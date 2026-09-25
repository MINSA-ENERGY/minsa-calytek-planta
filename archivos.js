// CALYTEK Planta — Archivos.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { diasPara, esLoteDeLaApp, fechaCorta, lista, plural, tipoDeArchivo } from './reglas.js';
import { $, avisar, el, enPlanta, esperaAutorizacion, estado, firmar, horaCorta, nombreDe, normaliza, opciones, porId, prealtaCambioTrasFirma, quien, renglon, selloSinFirma, sellosSinFirma, vivo } from './nucleo.js';
import { irA } from './navegacion.js';
import { subirEvidencia } from './gondolas.js';
import { sinMovimientoDe, verPrealta } from './prealtas.js';
import { irPadron, NOMBRE_PADRON, VIGENCIAS_PADRON, vigenciasDelCarrier } from './padron.js';
import { etiquetaCompuertaDe } from './hoy.js';

// ================================================================ ARCHIVOS (v0.33.0, sección aparte; artifact 1GvBJaYooYvjZT4rMRtL9Q)
// El árbol de la biblioteca Ambiental-CALYTEK leído por Graph (Sites.Selected; la misma ruta drive/root: con la que sube la
// evidencia): tres ramas —evidencia de báscula (carpeta por mes → lote por pesaje → foto, ticket, manifiesto), oficios ASEA y
// CSF de los carriers, y el buzón 99_Pendiente-Archivar, donde la app deja los lotes hasta que /archivar-calytek los acomode—.
// Cada carpeta se lee al abrirla (una llamada, paginada) y se recuerda hasta el siguiente Actualizar a mano. Los filtros
// (programa · tipo · buscador) actúan sobre lo ya leído: una carpeta cerrada no se juzga porque no se conoce. Solo lectura.
const RAMAS_ARCHIVOS = () => [
    { ruta: CONFIG.evidencia.destinoBase, titulo: 'Evidencia de báscula', nota: 'por mes · un lote por pesaje' },
    { ruta: CONFIG.archivos.padron, titulo: 'Oficios ASEA y CSF de los carriers', nota: '' },
    { ruta: CONFIG.buzon, titulo: 'Pendiente de archivar', nota: 'lo que la app acaba de subir', soloLotes: true }   // S-15: solo los lotes de la app, no el buzón entero
];
const ICONO_CARPETA = 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z';
const TIPOS_ARCHIVO = { ticket: 'ticket', foto: 'foto del indicador', manifiesto: 'manifiesto', oficio: 'oficio ASEA', csf: 'CSF', lote: 'registro que deja la app al pesar', otro: 'otro' };
function svgIcono(d, clase) {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('aria-hidden', 'true'); if (clase) s.setAttribute('class', clase);
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); s.appendChild(p);
    return s;
}
const tamano = n => n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : n >= 1024 ? `${Math.round(n / 1024)} KB` : `${n || 0} B`;
async function hijosDe(ruta) {
    const a = estado.archivos;
    if (!a) throw new Error('la biblioteca se está releyendo; abre la carpeta de nuevo');   // C-34 (v0.34.0)
    if (!a.ramas.has(ruta)) a.ramas.set(ruta, await estado.cliente.hijos(estado.siteId, ruta, m => avisar(m, 'ojo')));
    return a.ramas.get(ruta);
}
/** C-31 (v0.34.0): olvida lo leído de una rama y, si su carpeta está abierta en el árbol, la relee ya. La llama subirEvidencia. */
export function invalidarRama(ruta) {
    if (estado.archivos) estado.archivos.ramas.delete(ruta);
    for (const d of document.querySelectorAll('#arArbol details')) if (d.dataset.ruta === ruta) { delete d.dataset.leida; if (d.open && d.leer) d.leer(); }
}
export function pintarArchivos() {
    // el filtro de programa sale de las pre-altas cargadas (todas: una cerrada sigue teniendo sus lotes en la biblioteca)
    opciones($('arPrograma'), [...estado.prealtas].sort((x, y) => y.id - x.id), p => p.id, p => `${p.Title}${p.Estado === 'cerrada' ? ' · cerrada' : ''}`, 'Todos los programas');
    if (!estado.archivos) { estado.archivos = { biblioteca: null, ramas: new Map() }; construirArbol().catch(err => avisar('No se pudo armar el árbol: ' + err.message, 'error')); }   // C-34: sin catch, un throw dejaba «Abriendo…» para siempre
    else filtrarArbol();
}
async function construirArbol() {
    const a = estado.archivos;
    const caja = $('arArbol'); caja.textContent = ''; caja.appendChild(el('p', 'vacio', 'Abriendo la biblioteca…'));
    const lnk = $('lnkArchivosSharePoint'); lnk.href = `https://${CONFIG.sharepointHost}${CONFIG.sitio}`;
    try { a.biblioteca = await estado.cliente.biblioteca(estado.siteId); if (a.biblioteca.webUrl) lnk.href = a.biblioteca.webUrl; }
    catch (err) { avisar('No se pudo abrir la biblioteca: ' + err.message, 'error'); }
    if (a !== estado.archivos) return;   // hubo un Actualizar en medio: lo pinta la lectura nueva
    caja.textContent = '';
    for (const r of RAMAS_ARCHIVOS()) caja.appendChild(nodoCarpeta({ name: r.ruta }, r.ruta, r.titulo, r.nota, r.soloLotes));
    caja.firstChild.open = true;   // la evidencia abierta de entrada, como en el artifact
}
/** Una carpeta del árbol: <details> que lee sus hijos por Graph la primera vez que se abre. `titulo` solo en las tres raíces;
 *  `soloLotes` (S-15) deja ver únicamente las carpetas de lote de la app. `d.leer` la relee (C-31, invalidarRama). */
function nodoCarpeta(item, ruta, titulo, nota, soloLotes = false) {
    const d = el('details'); d.dataset.ruta = ruta; d.dataset.nombre = normaliza(titulo ? `${titulo} ${item.name}` : item.name);
    const s = el('summary'); s.appendChild(svgIcono('M9 6l6 6-6 6', 'flecha')); s.appendChild(svgIcono(ICONO_CARPETA));
    s.appendChild(el('span', '', titulo || item.name));
    if (nota) s.appendChild(el('span', 'sub', nota));
    if (titulo) s.title = ruta;
    const n = el('span', 'n', item.folder && item.folder.childCount !== undefined ? String(item.folder.childCount) : ''); s.appendChild(n);
    d.appendChild(s);
    const hijos = el('div', 'hijos'); d.appendChild(hijos);
    let leyendo = false;
    d.leer = async () => {
        if (leyendo || d.dataset.leida === '1') return;
        leyendo = true;
        hijos.textContent = ''; hijos.appendChild(el('p', 'vacio', 'Leyendo…'));
        let items;
        try { items = await hijosDe(ruta); }
        catch (err) { hijos.textContent = ''; hijos.appendChild(el('p', 'vacio', 'No se pudo leer: ' + err.message)); leyendo = false; return; }
        leyendo = false;
        hijos.textContent = ''; d.dataset.leida = '1';   // desde aquí el filtro sí la juzga
        if (items === null) { hijos.appendChild(el('p', 'vacio', 'Esta carpeta aún no existe en la biblioteca.')); n.textContent = '—'; return; }
        if (soloLotes) items = items.filter(it => it.folder && esLoteDeLaApp(it.name, CONFIG.evidencia.etiqueta));   // S-15
        if (!items.length) hijos.appendChild(el('p', 'vacio', soloLotes ? 'Sin lotes de la app.' : 'Vacía.'));
        // carpetas primero y lo más reciente arriba: los nombres de la casa empiezan por la fecha
        const orden = items.slice().sort((x, y) => ((y.folder ? 1 : 0) - (x.folder ? 1 : 0)) || y.name.localeCompare(x.name, 'es'));
        for (const it of orden) hijos.appendChild(it.folder ? nodoCarpeta(it, `${ruta}/${it.name}`) : nodoArchivo(it));
        const nc = orden.filter(x => x.folder).length, na = orden.length - nc;
        n.textContent = [nc ? plural(nc, soloLotes ? 'lote' : 'carpeta') : '', na ? plural(na, 'archivo') : ''].filter(Boolean).join(' · ') || '0';
        filtrarArbol();
    };
    d.addEventListener('toggle', () => { if (d.open) d.leer(); });
    return d;
}
/** Un archivo del árbol: enlace a SharePoint (abre en otra pestaña), extensión, nombre y «fecha · tamaño» en mono. */
function nodoArchivo(it) {
    const a = el('a', 'arch'); a.href = it.webUrl || '#'; a.target = '_blank'; a.rel = 'noopener';
    const tipo = tipoDeArchivo(it.name); a.dataset.tipo = tipo; a.dataset.nombre = normaliza(it.name);
    const ext = (String(it.name).match(/\.([a-z0-9]{1,4})$/i) || [])[1] || '';
    a.appendChild(el('span', 'ext', ext.toUpperCase()));
    a.appendChild(el('span', '', it.name));
    a.title = `${it.name} · ${TIPOS_ARCHIVO[tipo]}`;
    const dd = el('span', 'd', fechaCorta(it.lastModifiedDateTime)); dd.appendChild(el('span', 'tam', ` · ${tamano(Number(it.size) || 0)}`)); a.appendChild(dd);   // U-67: en celular la fecha baja de renglón y el tamaño se oculta
    return a;
}
/** Aplica los tres filtros sobre lo ya leído: un archivo pega por nombre y tipo; una carpeta leída se esconde si no le queda nada visible. */
export function filtrarArbol() {
    const q = normaliza($('arBusca').value.trim());
    const tipo = $('arTipo').value;
    const pre = porId(estado.prealtas, $('arPrograma').value);
    // programa = los folios de sus góndolas (E-26-00012 → e-26-00012, que es como viaja en el nombre del lote y de la foto)
    const folios = pre ? estado.embarques.filter(e => Number(e.PreAltaId) === pre.id && e.Title).map(e => normaliza(e.Title)) : [];
    const pegaNombre = nombre => (!q || nombre.includes(q)) && (!pre || folios.some(f => nombre.includes(f)));
    const hayFiltro = !!(q || tipo || pre);
    const caja = $('arArbol');
    for (const a of caja.querySelectorAll('.arch')) a.classList.toggle('oculto', hayFiltro && !(pegaNombre(a.dataset.nombre) && (!tipo || a.dataset.tipo === tipo)));
    for (const d of [...caja.querySelectorAll('details')].reverse()) {   // de adentro hacia afuera: el padre juzga hijos ya juzgados
        const alguno = [...d.querySelectorAll(':scope > .hijos > .arch, :scope > .hijos > details')].some(x => !x.classList.contains('oculto'));
        const propio = !tipo && pegaNombre(d.dataset.nombre);
        d.classList.toggle('oculto', hayFiltro && d.dataset.leida === '1' && !alguno && !propio);   // una carpeta sin leer no se juzga
    }
    $('arVacio').classList.toggle('oculto', !hayFiltro || [...caja.children].some(x => !x.classList.contains('oculto')));
    // U-135 (v0.74.0): con un filtro puesto, decir cuántas carpetas no se buscaron porque nadie las ha abierto, y ofrecer leerlas.
    const sinLeer = hayFiltro ? [...caja.querySelectorAll('details')].filter(d => d.dataset.leida !== '1') : [];
    $('arSinLeer').classList.toggle('oculto', !sinLeer.length);
    $('arSinLeerTexto').textContent = sinLeer.length ? `${sinLeer.length === 1 ? '1 carpeta sin abrir no se buscó' : `${sinLeer.length} carpetas sin abrir no se buscaron`}.` : '';
}
/** U-135: lee las carpetas que el filtro no pudo juzgar (un nivel; lo que traigan dentro vuelve a contarse). */
function leerCarpetasSinLeer() {
    for (const d of $('arArbol').querySelectorAll('details')) if (d.dataset.leida !== '1' && d.leer) d.leer();
}

// Pendiente revisar: pre-altas por firmar, excepciones y programas dormidos. Con algo, la tarjeta se pinta en ambar
// con el conteo en rojo (Carlos, 2026-09-08: es lo primero que hay que atender).
export function pintarPendientesHoy(borradores, pendientes, botonesExcepcion) {
    const pf = $('tbPendientes'); pf.textContent = '';
    const dormidas = estado.prealtas.filter(p => p.Estado === 'firmada').map(p => ({ p, sm: sinMovimientoDe(p) })).filter(x => x.sm);
    // S-01: sellos sin firma en PLANTA_Firmas (la compuerta no los acepta) y la lista misma si no esta provisionada.
    const ssf = sellosSinFirma().prealtas;
    const nPend = borradores.length + pendientes.length + dormidas.length + ssf.length + (estado.firmasError ? 1 : 0);
    $('tbPendientesTarjeta').classList.toggle('alerta', nPend > 0);
    $('tbPendientesN').classList.toggle('oculto', !nPend); $('tbPendientesN').textContent = String(nPend);
    if (!nPend) pf.appendChild(el('p', 'vacio', 'Nada pendiente.'));
    if (estado.firmasError) pf.appendChild(renglon('No se pudo leer el registro de firmas', `la app no firma ni autoriza y ningún sello vale sin su firma · Actualiza; si sigue, avisa a gerencia${estado.rol === 'gerencia' ? ` · PLANTA_Firmas: ${estado.firmasError} (permisos de la lista; setup-carlos.md, tarea 11)` : ''}`));   // U-31
    for (const p of ssf) pf.appendChild(renglon(`Firma · ${p.Title}`, prealtaCambioTrasFirma(p) ? `cambió después de que la firmó ${quien(p.FirmadaPor) || '?'} (carrier, unidades, choferes, corriente o generador) · la puerta no la ve · se vuelve a firmar desde su detalle` : `falta la firma del validador (trae el sello de ${quien(p.FirmadaPor) || '?'}, sin firma registrada) · la puerta no la ve · se firma desde su detalle`, 'Ver', () => verPrealta(vivo('prealtas', p))));   // U-28 / U-31
    for (const { p, sm } of dormidas) pf.appendChild(renglon(`Programa · ${p.Title}`, `${sm.motivo} · ¿se cierra? Sigue saliendo en la puerta`, 'Ver', () => verPrealta(vivo('prealtas', p))));
    for (const p of borradores) { const d = diasPara(p.FechaEstimada); pf.appendChild(renglon(`Pre-alta · ${p.Title}`, `firma del validador · 1er envío ${fechaCorta(p.FechaEstimada)}${d !== null ? ` (en ${d} días)` : ''} · capturó ${quien(p.CapturadaPor) || '?'}`, 'Ver', () => verPrealta(vivo('prealtas', p)))); }
    for (const e of pendientes) pf.appendChild(renglon(`Excepción · ${e.PlacaTractor}`, `${selloSinFirma(e)}autorización de gerencia · «${e.ExcepcionMotivo || 'sin motivo'}» · ${horaCorta(e.Arribo)}`, null, null, botonesExcepcion(e).map(b => ({ ...b, clase: b.accion === 'autorizar' ? '' : 'peligro' }))));
}

// C-29 (v0.34.0): UNA definición de «rechazo o excepción» y UN renglón para Hoy (los 10 últimos) y Reportes (todos). Antes el
// filtro y la lectura de CompuertaDetalle vivían copiados en los dos y la etiqueta difería (Hoy decía legal/comercial).
export const esRechazo = e => e.Etapa !== 'anulado' && (e.Etapa === 'rechazado' || e.Compuerta === 'excepcion-comercial');
export const rechazosYExcepciones = () => estado.embarques.filter(esRechazo).sort((a, b) => b.id - a.id);
/** C-54 (v0.51.0): la ÚNICA lectura de CompuertaDetalle — los hallazgos de las clases pedidas; detalle ilegible = ninguno. */
export function hallazgosDe(e, ...clases) {
    try { return JSON.parse(e.CompuertaDetalle || '[]').filter(h => clases.includes(h.clase)); } catch (_) { return []; }
}
export const reglasDe = (e, ...clases) => hallazgosDe(e, ...clases).map(h => h.regla).join(', ');
export function renglonRechazo(e) {
    const causa = reglasDe(e, 'legal', 'comercial');
    const r = renglon(`${e.Title || 'sin folio'} · ${e.PlacaTractor} · ${nombreDe(estado.carriers, e.CarrierId)}`, `${horaCorta(e.Arribo)} · ${causa}${e.ExcepcionAutorizo ? ' · autorizó ' + quien(e.ExcepcionAutorizo) : ''}`, 'Ver', () => abrirGondolaDe(vivo('embarques', e)));   // U-134 (v0.74.0)
    r.firstChild.firstChild.prepend(etiquetaCompuertaDe(e));   // U-143 (v0.73.0): la etiqueta va primero; al final caía sola en otro renglón
    return r;
}
export function pintarRechazosHoy() {
    const rj = $('tbRechazos'); rj.textContent = '';
    const rech = rechazosYExcepciones().slice(0, 10);
    if (!rech.length) rj.appendChild(el('p', 'vacio', 'Ningún rechazo ni excepción en lo cargado.'));   // mismo vacío que Reportes
    for (const e of rech) rj.appendChild(renglonRechazo(e));
}

// Vigencias como tiempo restante: barra llena = hoy vence; roja = ya vencio.
/**
 * U-29 (v0.27.0): las vigencias del PADRON (tarjeta, poliza, licencia, CSF de carriers/unidades/choferes activos) que
 * vencen dentro de la ventana, con la forma de un renglon del tablero. Hasta v0.26.0 Hoy solo miraba estado.vigencias
 * —al tablero solo se espeja la ASEA del carrier (C-05)— y decia «Nada vence» con una licencia venciendo la semana
 * siguiente. La ASEA del carrier que YA tiene renglon en el tablero no se repite (vive en las dos listas).
 */
function vigenciasPadronHoy() {
    const v = [];
    for (const [tipo, coleccion] of [['carriers', estado.carriers], ['unidades', estado.unidades], ['choferes', estado.choferes]]) {
        for (const it of coleccion.filter(x => x.Activo !== false)) {
            for (const [n, col] of VIGENCIAS_PADRON[tipo]) {
                if (tipo === 'carriers' && col === 'VigenciaASEA' && vigenciasDelCarrier(it).length) continue;
                const d = diasPara(it[col]);
                if (d !== null && d <= CONFIG.avisoVigenciaDias) v.push({ Title: `${NOMBRE_PADRON[tipo].replace(/^./, c => c.toUpperCase())} ${it.Title} · ${n}`, Vence: it[col], Fuente: 'padrón', AvisoDias: CONFIG.avisoVigenciaDias, ficha: { clave: tipo, id: it.id, carrier: tipo === 'carriers' ? it.id : Number(it.CarrierId) } });   // U-134: el renglón abre su ficha
            }
        }
    }
    return v;
}
/** U-134: de Hoy a la ficha del padrón; Volver lleva al expediente de su carrier. */
function abrirFichaDesdeHoy(f) {
    irA('padron');
    irPadron('ficha', { ficha: { clave: f.clave, id: f.id }, carrier: f.carrier, desde: 'carrier' });
}
/** U-134: de un rechazo a su góndola — En planta si sigue ahí (excepción en espera), si no el Historial buscando su folio o placa. */
function abrirGondolaDe(e) {
    estado.vistaGondolas = enPlanta(e) || esperaAutorizacion(e) ? 'planta' : 'historial';
    $('baBusca').value = estado.vistaGondolas === 'historial' ? (e.Title || e.PlacaTractor || '') : '';
    irA('bascula');
}
export function pintarVigenciasHoy() {
    const vg = $('tbVigencias'); vg.textContent = '';
    const prox = [...estado.vigencias.filter(v => v.Activo !== false), ...vigenciasPadronHoy()].map(v => ({ v, d: diasPara(v.Vence) })).filter(x => x.d !== null && x.d <= (Number(x.v.AvisoDias) || CONFIG.avisoVigenciaDias)).sort((a, b) => a.d - b.d);
    if (!prox.length) vg.appendChild(el('p', 'vacio', `Nada vence en ${CONFIG.avisoVigenciaDias} días.`));
    for (const { v, d } of prox) {
        const ventana = Number(v.AvisoDias) || CONFIG.avisoVigenciaDias;
        const clase = d < 0 ? 'mal' : d <= 7 ? 'ojo' : '';
        // U-134 (v0.74.0): la del padrón es un botón que abre la ficha del chofer, unidad o carrier; la del tablero no tiene ficha.
        const r = el(v.ficha ? 'button' : 'div', v.ficha ? 'vig abre' : 'vig');
        if (v.ficha) { r.type = 'button'; r.title = 'Abrir su ficha en el Padrón'; r.addEventListener('click', () => abrirFichaDesdeHoy(v.ficha)); }
        const t = el('span', '', v.Title); t.appendChild(el('small', '', `${fechaCorta(v.Vence)} · ${v.Fuente || ''}${v.Dueno ? ' · dueño ' + v.Dueno : ''}`)); r.appendChild(t);
        r.appendChild(el('span', 'd ' + clase, d < 0 ? `−${-d} d` : `${d} d`));
        const bar = el('span', 'bar'); const i = el('i', clase); i.style.width = Math.max(4, Math.min(100, Math.round((1 - d / ventana) * 100))) + '%'; bar.appendChild(i); r.appendChild(bar);
        vg.appendChild(r);
    }
}

$('btnArLeerTodas').addEventListener('click', leerCarpetasSinLeer);
// v0.66.0: cada conteo del padrón abre su grupo y lo trae a la vista.
