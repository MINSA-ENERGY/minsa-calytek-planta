// CALYTEK Planta — Pre-altas: lista, detalle y asistente.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { aIsoDia, basesRecientes, clienteDe, clientesPrealta, CORRIENTES, etiquetaCorriente, evaluarVigencia, fechaCorta, fechaDePestana, fechaMexico, huellaPrealta, limpiar, lista, mesesPrealtas, paraPatch, plural, prealtaSinMovimiento, PUEDE, siguienteFolio } from './reglas.js';
import { $, abrirForma, anclar, aplicar, avisar, cerrarForma, confirmar, el, embarquesDeProgramas, esColumnaFaltante, escribiendo, estado, etiqueta, firmar, hayCaptura, horaCorta, huellaForma, L, limpiarAvisos, motivoSinFirmas, nombreDe, opciones, pintarInsignias, porId, prealtaCambioTrasFirma, prealtaFirmada, quien, reanclar, refrescarCliente, textoDe, vivo } from './nucleo.js';
import { entrar, irA, repintar } from './navegacion.js';
import { asegurarFolioUnico, renglonOpcion } from './puerta.js';
import { pintarCertificadoEnDetalle } from './certificado.js';
import { activo, etiquetaVigencia, sinFecha, vigenciasPadron } from './padron.js';

// ================================================================ PRE-ALTAS

/**
 * Gondolas de una pre-alta: recibidas contra esperadas. Recibida = embarque que existe y llego;
 * el anulado (folio quemado) y el rechazado (no entro) no cuentan.
 */
export function gondolasDe(p, conteo = conteoRecibidas()) {
    return { rec: conteo.get(Number(p.id)) || 0, esp: Number(p.GondolasEsperadas) || 0 };
}
/** C-68 (v0.63.0): PreAltaId -> recibidas, en UNA pasada; pintarPrealtas la calcula una vez por repintado (antes, 3-5 por programa). */
function conteoRecibidas(emb = embarquesDeProgramas()) {
    const m = new Map();
    for (const e of emb) if (e.Etapa !== 'anulado' && e.Etapa !== 'rechazado') { const k = Number(e.PreAltaId); m.set(k, (m.get(k) || 0) + 1); }
    return m;
}
export function sinMovimientoDe(p, emb = embarquesDeProgramas()) { return prealtaSinMovimiento(p, emb, CONFIG.sinMovimientoDias); }
function barraAvance({ rec, esp }) {
    const d = el('div', 'avance');
    if (esp) {
        const b = el('span', 'barra' + (rec ? '' : ' esp'));
        const i = el('i'); i.style.width = Math.min(100, Math.round(rec / esp * 100)) + '%';
        b.appendChild(i); d.appendChild(b);
        d.appendChild(el('span', 'cifra', `${rec}/${esp}`));
    } else d.appendChild(el('span', 'cifra', `${plural(rec, 'recibida')} · sin estimado`));
    return d;
}
/** Las tres pestañas de Pre-altas (rediseño tanda 1, v0.53.0): estado de la pre-alta -> panel. */
const VISTAS_PREALTAS = { borrador: 'ppBorradores', firmada: 'ppFirmadas', cerrada: 'ppCerradas' };
/** U-101 (v0.57.0): el Estado con la palabra de su pestaña; «borrador» es el valor de la columna, no lo que lee la gente. */
const ESTADO_PREALTA = { borrador: 'por firmar', firmada: 'firmada', cerrada: 'cerrada' };
const estadoPrealta = e => ESTADO_PREALTA[e] || e || 'sin estado';
export function elegirVistaPrealtas(v) {
    estado.vistaPrealtas = VISTAS_PREALTAS[v] ? v : 'borrador';
    for (const b of document.querySelectorAll('#paTabs button, .pa-kpis button')) b.setAttribute('aria-pressed', String(b.dataset.pa === estado.vistaPrealtas));
    for (const [k, id] of Object.entries(VISTAS_PREALTAS)) $(id).hidden = k !== estado.vistaPrealtas;
}
/**
 * Pantalla principal, tanda 2 (v0.61.0; la K del lienzo EhJVtU3d4X6jUNuhJb6KQR): un programa es un renglón de columnas reales
 * —Programa · Folio · Corriente · Carrier · fecha de la pestaña · Góndolas · Ver— y los renglones van en cebra (`non`, lo
 * alterna pintarPrealtas porque las barras de mes cortan el nth-child). Corriente y carrier van juntos en `.cc` para que en
 * celular (la L) bajen a un renglón; en escritorio `.cc` es display: contents y cada uno es su columna. Las etiquetas de
 * excepción (estado que no es el de su pestaña, «sin firma», «¿se cierra?») van junto al nombre, como antes.
 */
const FECHA_PESTANA = { borrador: '1er envío', firmada: 'Firmada', cerrada: 'Cerrada' };
const fechaPestanaCorta = (p, grupo) => {
    const f = fechaDePestana(p, grupo);
    return !f ? '—' : grupo === 'borrador' ? fechaCorta(f) : fechaCorta(fechaMexico(new Date(f)));
};
/** C-69 (v0.63.0): las recibidas de programas SIN estimado van aparte (`sin`): sumadas a las otras inflaban «15/20». */
function sumaGondolas(ps, conteo) {
    return ps.reduce((a, p) => { const g = gondolasDe(p, conteo); if (g.esp) { a.rec += g.rec; a.esp += g.esp; } else a.sin += g.rec; return a; }, { rec: 0, esp: 0, sin: 0 });
}
/** El texto del subtotal de mes y del Total, uno solo (C-69): «3/12 góndolas», «3/12 góndolas + 2 sin estimado», o solo las recibidas. */
function textoGondolas({ rec, esp, sin }) {
    if (!esp) return `${plural(sin, 'góndola recibida', 'góndolas recibidas')} · sin estimado`;
    return `${rec}/${esp} góndolas` + (sin ? ` + ${sin} sin estimado` : '');
}
/** U-114 (v0.63.0): la rejilla es tabla para el lector de pantalla — role table/row/columnheader/cell, sin cambiar el CSS. */
export const celda = (tag, clase, texto, rol = 'cell') => { const c = el(tag, clase, texto); c.setAttribute('role', rol); return c; };
/**
 * C-74 (v0.63.0): el renglón de programa arma sus celdas explícitas —ya no parchea el DOM de renglon() con firstChild/lastChild—
 * y el CSS las ubica por NOMBRE de área de rejilla (--pa-areas), no por índice de columna. U-109: tocar el renglón abre la
 * ficha igual que Ver (el botón sigue siendo el control de teclado; no se anidan interactivos).
 */
function renglonPrograma(p, grupo, ctx, non = false) {
    const abrir = () => verPrealta(vivo('prealtas', p));
    const r = el('div', 'renglon prog' + (non ? ' non' : '')); r.setAttribute('role', 'row');
    const pr = celda('div', 'pr'), t = el('div', 't', p.Title);
    pr.appendChild(t);
    if (p.Estado !== grupo) t.appendChild(etiqueta(estadoPrealta(p.Estado), p.Estado));
    if (grupo === 'firmada' && !prealtaFirmada(p)) t.appendChild(etiqueta('sin firma', 'vencida'));   // S-01
    const sm = grupo === 'firmada' ? sinMovimientoDe(p, ctx.emb) : null;
    if (sm) { t.appendChild(etiqueta('¿se cierra?', 'aviso')); pr.appendChild(el('p', 'pista', `${sm.motivo}. Sigue saliendo en la puerta hasta que alguien cierre el programa.`)); }
    const cc = el('div', 'cc');
    const cor = etiquetaCorriente(p.Corriente) || 'sin corriente', car = nombreDe(estado.carriers, p.CarrierId);
    cc.appendChild(celda('span', 'cor', cor)).title = cor;   // U-112: en una línea con elipsis; el nombre completo al pasar el ratón
    cc.appendChild(celda('span', 'car', car)).title = car;
    const av = celda('div', 'gon'); av.appendChild(barraAvance(gondolasDe(p, ctx.conteo)));
    for (const c of [pr, celda('span', 'fol', p.Campana || '—'), cc, celda('span', 'fe', fechaPestanaCorta(p, grupo)), av]) r.appendChild(c);
    r.tabIndex = 0;   // v0.68.0: sin botón Ver (pedido de Carlos); el renglón es el control
    r.addEventListener('keydown', e => { if (e.target === r && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); abrir(); } });
    r.addEventListener('click', e => { if (!e.target.closest('button')) abrir(); });
    return r;
}
function encabezadoProgramas(grupo) {
    const h = el('div', 'pa-cab'); h.setAttribute('role', 'row');
    for (const s of ['Programa', 'Folio', 'Corriente', 'Carrier', FECHA_PESTANA[grupo], 'Góndolas']) h.appendChild(celda('span', '', s, 'columnheader'));
    return h;
}
/** U-111 (v0.63.0): la barra de mes va en la misma rejilla que los renglones y su subtotal cae en la columna Góndolas. */
function barraMes({ mes, ps }, ctx) {
    const b = el('div', 'pa-mes'); b.setAttribute('role', 'row');
    const m = celda('div', 'mt'); m.appendChild(el('b', '', mes)); m.appendChild(el('span', 'n', plural(ps.length, 'programa')));
    b.appendChild(m);
    b.appendChild(celda('span', 'g', textoGondolas(sumaGondolas(ps, ctx.conteo))));
    return b;
}
export function pintarPrealtas() {
    cerrarForma('paDetalle');
    const captura = PUEDE.capturarPrealta(estado.rol);
    $('paInicio').classList.toggle('oculto', !captura);
    $('btnNuevaPrealta').classList.toggle('oculto', !captura);   // además del contenedor: visible() de la E2E mira el botón
    // Un Estado que no sea uno de los tres cae en cerradas para que no desaparezca de la vista;
    // ahi su etiqueta lo delata (renglonPrograma la pinta cuando no coincide con la pestaña).
    const grupos = { borrador: [], firmada: [], cerrada: [] };
    for (const p of [...estado.prealtas].sort((a, b) => b.id - a.id)) (grupos[p.Estado] || grupos.cerrada).push(p);

    // Pantalla principal, tanda 1 (v0.60.0): los recientes viven en el pop-up #paRecientes; el botón de la banda solo sale si hay.
    const bases = captura ? basesRecientes(estado.prealtas) : [];
    $('btnPaBases').classList.toggle('oculto', !bases.length);
    if (!bases.length) cerrarForma('paRecientes');
    const cb = $('paBases'); cb.textContent = '';
    for (const p of bases) {
        const b = el('button', 'pa-base'); b.type = 'button';
        b.appendChild(el('b', '', p.Title));
        b.appendChild(el('small', '', `${etiquetaCorriente(p.Corriente) || 'sin corriente'} · ${nombreDe(estado.carriers, p.CarrierId)}`));
        b.appendChild(el('span', 'usar', 'Usar como base ›'));
        b.addEventListener('click', () => { cerrarForma('paRecientes'); usarComoBase(vivo('prealtas', p)); });
        cb.appendChild(b);
    }

    const emb = embarquesDeProgramas(), ctx = { emb, conteo: conteoRecibidas(emb) };   // C-68: una pasada por repintado
    const vacio = { borrador: 'Ninguna por firmar.', firmada: 'Ninguna firmada: la puerta no puede recibir.', cerrada: 'Ninguna cerrada.' };
    const TITULO_TABLA = { borrador: 'Programas por firmar', firmada: 'Programas firmados', cerrada: 'Programas cerrados' };
    for (const [grupo, cont] of [['borrador', 'paBorradores'], ['firmada', 'paFirmadas'], ['cerrada', 'paCerradas']]) {
        const c = $(cont); c.textContent = '';
        const ps = grupos[grupo];
        if (ps.length) {
            c.setAttribute('role', 'table'); c.setAttribute('aria-label', TITULO_TABLA[grupo]);   // U-114
            // Tanda 2 (v0.61.0): Por firmar sigue el orden de captura, sin meses; Firmadas y Cerradas van por mes de su fecha.
            c.appendChild(encabezadoProgramas(grupo));
            let k = 0;
            for (const b of grupo === 'borrador' ? [{ mes: null, ps }] : mesesPrealtas(ps, grupo)) {
                if (b.mes) c.appendChild(barraMes(b, ctx));
                for (const p of b.ps) c.appendChild(renglonPrograma(p, grupo, ctx, k++ % 2 === 1));
            }
            const tot = el('div', 'pa-total'); tot.setAttribute('role', 'row');
            tot.appendChild(celda('b', 'tt', 'Total'));
            tot.appendChild(celda('span', 'g', textoGondolas(sumaGondolas(ps, ctx.conteo))));
            c.appendChild(tot);
        } else { c.removeAttribute('role'); c.removeAttribute('aria-label'); }
        if (!ps.length) c.appendChild(el('p', 'vacio', estado.prealtas.length || grupo !== 'borrador' ? vacio[grupo] : 'No hay pre-altas. La primera góndola no puede entrar sin una firmada.'));
    }
    $('paNBorradores').textContent = String(grupos.borrador.length);
    $('paNFirmadas').textContent = String(grupos.firmada.length);
    $('paNCerradas').textContent = String(grupos.cerrada.length);
    for (const [id, g] of [['paKBorradores', 'borrador'], ['paKFirmadas', 'firmada'], ['paKCerradas', 'cerrada']]) $(id).textContent = String(grupos[g].length);   // la banda
    // Las pestañas con algo que atender se marcan: borradores por firmar y firmadas sin firma o que ya no se mueven.
    // U-110 (v0.63.0): los conteos de la banda siguen la MISMA regla (antes, «Por firmar» siempre ámbar y «Firmadas» siempre verde).
    const alerta = { borrador: grupos.borrador.length > 0, firmada: grupos.firmada.some(p => sinMovimientoDe(p, emb) || !prealtaFirmada(p)), cerrada: false };
    $('paNBorradores').classList.toggle('alerta', alerta.borrador);
    $('paNFirmadas').classList.toggle('alerta', alerta.firmada);
    for (const b of document.querySelectorAll('.pa-kpis button')) b.classList.toggle('alerta', alerta[b.dataset.pa]);

    const sb = sumaGondolas(grupos.borrador, ctx.conteo);
    $('paResBorradores').textContent = grupos.borrador.length
        ? `${plural(sb.esp, 'góndola comprometida', 'góndolas comprometidas')}; ninguna puede entrar hasta que se firme.`
        : '';
    // Tanda 2 (v0.61.0): el pie de Firmadas y Cerradas repetía la cuenta que ya da el renglón Total; queda vacío. Sin firmadas
    // el mensaje de la tabla sigue diciendo que la puerta no puede recibir (U-23, U-102).
    $('paResFirmadas').textContent = '';
    $('paResCerradas').textContent = '';
    elegirVistaPrealtas(estado.vistaPrealtas);
}

export function nuevaPrealta() {
    estado.prealtaEdit = null;
    $('paFormaTitulo').textContent = 'Nueva'; $('btnGuardarPrealta').textContent = 'Guardar para firma';
    estado.paAsis = { paso: 1, max: 1, revisar: false, genEdit: false, genDeId: null, otro: false, baseId: null };   // C-67: ids, no objetos
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false), c => c.id, c => c.Title);
    $('paCarrier').value = '';   // U-40: opciones() ya conserva el value; una pre-alta nueva empieza sin carrier
    pintarUnidadesChoferesPrealta();
    for (const id of ['paTitulo', 'paCliente', 'paPozoTitulo', 'paMes', 'paGenerador', 'paGeneradorDireccion', 'paGeneradorRegistro', 'paPozo', 'paFecha', 'paGondolas', 'paCorreoFecha', 'paCorreoRemitente', 'paNotas']) $(id).value = '';
    $('paCorriente').value = '';
    $('paMes').value = String(new Date().getFullYear());   // U-54: casi siempre es el año en curso
    armarTituloPrealta();
    abrirAsistentePrealta();
}
/**
 * P2 (rediseño tanda 1, v0.53.0): «Usar como base» abre la pre-alta nueva con cliente, generador, corriente y carrier del
 * programa elegido; pozo y envío quedan vacíos y el foco va al pozo. Si el carrier ya no está activo se deja sin elegir.
 * Tanda 2 (v0.54.0): el asistente abre en el paso 2 con el 1 y el 3 ya hechos.
 */
async function usarComoBase(p) {
    if (!p || !PUEDE.capturarPrealta(estado.rol) || !(await soltarCapturaPrealta())) return;
    nuevaPrealta();
    const f = textoDe;
    $('paCliente').value = clienteDe(p); $('paCorriente').value = f(p.Corriente);
    llenarGeneradorPrealta(p);
    if (estado.carriers.some(c => c.Activo !== false && Number(c.id) === Number(p.CarrierId))) $('paCarrier').value = f(p.CarrierId);
    pintarUnidadesChoferesPrealta(seleccionDe(p));   // U-94 (v0.56.0): la seleccion de la base, no todas
    armarTituloPrealta();
    Object.assign(estado.paAsis, { paso: 2, max: 3, baseId: p.id, genDeId: p.id });
    marcarOtroClientePrealta(); pintarAsistentePrealta();
    $('paAsis').dataset.huella = huellaForma($('paAsis'));   // lo copiado de la base no cuenta como captura: Cancelar no pregunta
    $('paPozoTitulo').focus();
}
/**
 * Editar un BORRADOR (v0.19.10, Carlos 2026-09-08: «ver el borrador no me da la opcion de editarlo»). Mismo
 * formulario que la alta, con `estado.prealtaEdit`; guardar es un PATCH que conserva Estado, Campana y quien capturo.
 * Una firmada no se edita: la puerta ya la usa; se cierra y se abre otra.
 */
export async function editarPrealta() {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'borrador' || !PUEDE.capturarPrealta(estado.rol)) return;
    if (!(await soltarCapturaPrealta())) return;   // U-95: el asistente conservado con una pre-alta a medias no se pisa callado
    estado.prealtaEdit = p;
    $('paFormaTitulo').textContent = `Editar ${p.Title}`; $('btnGuardarPrealta').textContent = 'Guardar cambios';
    estado.paAsis = { paso: 5, max: 5, revisar: true, genEdit: false, genDeId: null, otro: false, baseId: null };
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false || Number(c.id) === Number(p.CarrierId)), c => c.id, c => c.Title);
    const f = textoDe;   // C-26
    partirTituloPrealta(f(p.Title)); $('paCorriente').value = f(p.Corriente); llenarGeneradorPrealta(p);
    $('paPozo').value = f(p.Pozo); $('paCarrier').value = f(p.CarrierId);
    $('paFecha').value = p.FechaEstimada ? fechaCorta(p.FechaEstimada) : ''; $('paGondolas').value = f(p.GondolasEsperadas);
    $('paCorreoFecha').value = p.CorreoFecha ? fechaCorta(p.CorreoFecha) : ''; $('paCorreoRemitente').value = f(p.CorreoRemitente); $('paNotas').value = f(p.Notas);
    pintarUnidadesChoferesPrealta(seleccionDe(p));   // las que la pre-alta ya tenia; si no tenia ninguna guardada, todas
    armarTituloPrealta();
    cerrarForma('paDetalle');
    abrirAsistentePrealta();
}
/**
 * Nombre del programa = CLIENTE-POZO-AÑO, en mayusculas (el 3er campo era «mes» hasta la v0.19.18; id paMes se conserva) (Carlos, 2026-09-08: «GSM-IXACHI 15-2026»). Las tres partes se
 * capturan por separado y paTitulo (oculto) se arma solo; los guiones dentro de una parte se cambian por espacio para
 * que el nombre se pueda volver a partir al editar. El pozo del titulo rellena «Pozo / instalacion» si esta vacio.
 */
const PARTES_TITULO = ['paCliente', 'paPozoTitulo', 'paMes'];
const parteTitulo = id => $(id).value.trim().toUpperCase().replace(/\s*-\s*/g, ' ').replace(/\s+/g, ' ');
function armarTituloPrealta() {
    const partes = PARTES_TITULO.map(parteTitulo);
    const t = partes.every(Boolean) ? partes.join('-') : '';
    $('paTitulo').value = t; $('paTituloVista').textContent = t || '—';
    const pozo = $('paPozo'); if (partes[1] && (!pozo.value || pozo.value === pozo.dataset.auto)) { pozo.value = partes[1]; pozo.dataset.auto = partes[1]; }
}
function partirTituloPrealta(titulo) {
    const p = String(titulo || '').split('-').map(x => x.trim());
    const partes = p.length >= 3 ? [p[0], p.slice(1, -1).join(' '), p[p.length - 1]] : [titulo || '', '', ''];
    PARTES_TITULO.forEach((id, i) => { $(id).value = partes[i]; });
}
const seleccionDe = p => ({ unidades: lista(p.UnidadesIds), choferes: lista(p.ChoferesIds) });
/** U-92 (v0.56.0): si la corriente nueva no la ampara el carrier elegido, se suelta y SE DICE (antes se soltaba callado). */
function cambiarCorrientePrealta(k, texto) {
    const antes = $('paCarrier').value ? porId(estado.carriers, $('paCarrier').value) : null;
    $('paCorriente').value = k;
    if (antes && !carrierAmpara(antes, k)) {
        $('paCarrier').value = ''; pintarUnidadesChoferesPrealta();
        avisar(`${antes.Title} no ampara «${texto.toLowerCase()}»: se quitó del programa con sus unidades y choferes. Elige otro carrier en Transporte.`, 'ojo');
    }
    pintarAsistentePrealta();
}
/**
 * Las casillas de unidades y choferes del carrier elegido. `sel` ({ unidades, choferes }: ids) marca solo esas; sin `sel`, o con una
 * lista vacia, quedan todas (es lo que la puerta entiende). C-62/U-94 (v0.56.0): quien llama pasa la seleccion que ya existia.
 */
export function pintarUnidadesChoferesPrealta(sel) {
    const cid = Number($('paCarrier').value);
    const cajas = (cont, items, nombre, ids) => {
        cont.textContent = '';
        if (!items.length) { cont.appendChild(el('p', 'pista', `(ese carrier no tiene ${nombre} en el padrón)`)); return; }
        for (const it of items) {
            const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; c.value = String(it.id); c.checked = !ids || !ids.length || ids.includes(String(it.id));
            l.appendChild(c); l.appendChild(document.createTextNode(' ' + (it.PlacaPlana ? `${it.Title} / ${it.PlacaPlana}` : it.Title)));
            const v = etiquetaVigencia(nombre, it); if (v) l.appendChild(v);   // U-98 (v0.57.0): la vigencia a la vista, como en el padrón
            cont.appendChild(l);
        }
    };
    cajas($('paUnidades'), estado.unidades.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'unidades', sel && sel.unidades);
    cajas($('paChoferes'), estado.choferes.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'choferes', sel && sel.choferes);
}
function marcados(id) { return [...$(id).querySelectorAll('input:checked')].map(c => c.value).join(';'); }
/** C-65 (v0.57.0): el generador de un programa a sus tres campos; antes se copiaba a mano en tres sitios. */
function llenarGeneradorPrealta(p) {
    const f = textoDe;
    $('paGenerador').value = f(p.Generador); $('paGeneradorRegistro').value = f(p.GeneradorRegistro); $('paGeneradorDireccion').value = f(p.GeneradorDireccion);
}
/** C-65: lo que la forma guarda, armado una vez para la alta y la edición (la alta agrega Estado, CapturadaPor y Campana). */
function camposPrealta() {
    return {
        Title: valorPa('paTitulo'), Cliente: parteTitulo('paCliente'), Generador: valorPa('paGenerador'), GeneradorRegistro: valorPa('paGeneradorRegistro'), GeneradorDireccion: valorPa('paGeneradorDireccion'),   // Cliente: esquema v8, tanda 3
        Pozo: valorPa('paPozo'), Corriente: $('paCorriente').value, CarrierId: Number($('paCarrier').value),
        UnidadesIds: marcados('paUnidades'), ChoferesIds: marcados('paChoferes'),
        FechaEstimada: aIsoDia($('paFecha').value), GondolasEsperadas: valorPa('paGondolas') ? Number(valorPa('paGondolas')) : null,
        CorreoFecha: aIsoDia($('paCorreoFecha').value), CorreoRemitente: valorPa('paCorreoRemitente'), Notas: valorPa('paNotas')
    };
}
// ---------------------------------------------------------------- asistente de pre-alta (rediseño tanda 2, v0.54.0)
// Artifact 3PrGGb7ruvmrFhYNrrqdU1 v9: cinco pasos, una pantalla a la vez, en la pestaña (no en un pop-up). El estado del
// recorrido vive en estado.paAsis { paso, max (el más lejano visitado), revisar (vino de «Cambiar»: Siguiente regresa a Revisar),
// genEdit, genDe (el programa del que salió el generador), otro (cliente nuevo), base }. Lo capturado vive en los campos de
// siempre: guardarPrealta no cambió. P8: el botón dice qué falta y, tocado, lleva el foco ahí.
const PASOS_PREALTA = ['Cliente', 'Pozo y corriente', 'Transporte', 'Envío', 'Revisar'];
function abrirAsistentePrealta() {
    if (estado.pestana !== 'prealtas') irA('prealtas');   // Editar desde el detalle abierto en Hoy
    ocultarListoPrealta(); limpiarAvisos();
    $('p-prealtas').classList.add('asistiendo');
    $('paAsis').classList.remove('oculto');
    marcarOtroClientePrealta();   // C-64: al editar, un cliente que no es de los conocidos abre en «Otro»
    pintarAsistentePrealta();
    $('paAsis').dataset.huella = huellaForma($('paAsis'));
    window.scrollTo({ top: 0 });
    enfocarPasoPrealta();
}
function cerrarAsistentePrealta() {
    $('paAsis').classList.add('oculto'); $('p-prealtas').classList.remove('asistiendo');
    estado.paAsis = null; estado.prealtaEdit = null;
}
/** Cancelar o la miga: con algo cambiado desde que abrió, pregunta (U-09/U-34, lo que hacía Escape sobre el pop-up). */
export async function salirAsistentePrealta() {
    if (!(await soltarCapturaPrealta())) return;
    cerrarAsistentePrealta(); pintarPrealtas();
}
/** U-95 (v0.57.0): true si el asistente se puede soltar — cerrado, sin cambios desde que abrió, o quien captura dijo Descartar. */
async function soltarCapturaPrealta() {
    const d = $('paAsis');
    if (!estado.paAsis || !hayCaptura(d) || huellaForma(d) === d.dataset.huella) return true;
    const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: 'Esta pre-alta tiene datos sin guardar. Si sales, se pierden.' });
    return !!ok;
}
const valorPa = id => $(id).value.trim();
const fechaMala = id => { try { aIsoDia($(id).value); return false; } catch { return true; } };
/**
 * Lo que le falta a un paso para darse por hecho, con a dónde llevar el foco; null si está completo. 1–3 tienen obligatorios;
 * U-91/U-106 (v0.57.0): el rango del año (2) y las fechas y góndolas del 4 —opcionales, pero si se escriben, bien escritas— se
 * cotejan aquí y no al guardar, donde salían con el foco a un campo escondido.
 */
function faltaPrealta(n) {
    const primera = cont => $(cont).querySelector('.o:not(:disabled)') || $(cont).querySelector('input');
    if (n === 1) return valorPa('paCliente') ? null : { t: 'Elige el cliente', foco: () => primera('paClientes') };
    if (n === 2) {
        if (!valorPa('paPozoTitulo')) return { t: 'Escribe el pozo', foco: () => $('paPozoTitulo') };
        const anio = Number(valorPa('paMes')), hoy = new Date().getFullYear();   // U-54: entre el año pasado y el que viene
        if (!/^\d{4}$/.test(valorPa('paMes')) || anio < hoy - 1 || anio > hoy + 1) return { t: `Año: cuatro cifras, de ${hoy - 1} a ${hoy + 1}`, foco: () => $('paMes') };
        if (!$('paCorriente').value) return { t: 'Elige la corriente', foco: () => primera('paCorrientes') };
        return null;
    }
    if (n === 4) {
        if (fechaMala('paFecha')) return { t: 'Primer envío: fecha no válida', foco: () => $('paFecha') };
        if (valorPa('paGondolas') && !/^\d+$/.test(valorPa('paGondolas'))) return { t: 'Góndolas: un número entero', foco: () => $('paGondolas') };
        if (fechaMala('paCorreoFecha')) return { t: 'Correo: fecha no válida', foco: () => $('paCorreoFecha') };
        return null;
    }
    if (n === 3) {
        if (!$('paCarrier').value) return { t: 'Elige el carrier', foco: () => primera('paCarriers') };
        const ce = porId(estado.carriers, $('paCarrier').value);
        if (ce && !carrierAmpara(ce, $('paCorriente').value)) return { t: 'Elige un carrier que ampare la corriente', foco: () => primera('paCarriers') };
        for (const [cont, que] of [['paUnidades', 'una unidad'], ['paChoferes', 'un chofer']])
            if ($(cont).querySelector('input') && !$(cont).querySelector('input:checked')) return { t: `Marca al menos ${que}`, foco: () => $(cont).querySelector('input') };
    }
    return null;
}
function carrierAmpara(c, corriente) { const l = lista(c.Corrientes); return !corriente || !l.length || l.includes(corriente); }   // la misma regla que la compuerta
/**
 * P3: elegir un cliente conocido trae el generador de su último programa (con «Cambiar» en el paso 2). C-67 (v0.57.0): el
 * programa se resuelve por su id al tocar, no con el objeto que el renglón capturó al pintarse. v0.59.0: ya NO avanza al paso 2
 * (revierte U-105 a pedido de Carlos): el usuario pulsa «Siguiente».
 */
function elegirClientePrealta(clave, ultimoId) {
    const a = estado.paAsis;
    const cambio = parteTitulo('paCliente') !== clave;
    a.otro = false; $('paCliente').value = clave;
    const u = porId(estado.prealtas, ultimoId);
    if (u && (cambio || !valorPa('paGenerador'))) { llenarGeneradorPrealta(u); a.genDeId = u.Generador ? u.id : null; a.genEdit = false; }
    armarTituloPrealta(); pintarAsistentePrealta();
}
function resumenPasoPrealta(n) {
    if (n === 1) return parteTitulo('paCliente');
    if (n === 2) return [parteTitulo('paPozoTitulo'), etiquetaCorriente($('paCorriente').value)].filter(Boolean).join(' · ');
    if (n === 3) return $('paCarrier').value ? nombreDe(estado.carriers, $('paCarrier').value) : '';
    if (n === 4) return [valorPa('paFecha'), valorPa('paGondolas') && `${valorPa('paGondolas')} gón.`].filter(Boolean).join(' · ') || 'sin datos';
    return '';
}
/** Repinta los renglones de una lista de opciones sin perder el foco del teclado: el renglón enfocado se re-enfoca por su valor. */
function pintarOpcionesPa(cont, items) {
    const f = document.activeElement, v = f && f.parentElement === cont ? f.dataset.v : null;
    cont.textContent = '';
    for (const it of items) {
        const b = renglonOpcion(it);
        if (it.deshabilitada) b.disabled = true;
        b.addEventListener('click', it.alClic);
        cont.appendChild(b);
    }
    if (v !== null) { const n = [...cont.children].find(x => x.dataset.v === v); if (n) n.focus(); }
}
/**
 * C-64 (v0.58.0): el asistente se pinta por partes —la barra, el paso visible y el pie, más el resumen de escritorio— y los
 * pasos ocultos no se tocan: se pintan al entrar (irPasoPrealta). El pintor ya no escribe estado; a.otro lo fija quien carga
 * el cliente (marcarOtroClientePrealta). Antes eran 86 líneas que reconstruían los cinco pasos con cada tecla, Notas incluida.
 */
const PINTORES_PASO_PREALTA = { 1: pintarPaso1Prealta, 2: pintarPaso2Prealta, 3: pintarPaso3Prealta, 5: pintarPaso5Prealta };   // el 4 son campos fijos
function pintarAsistentePrealta() {
    const a = estado.paAsis; if (!a) return;
    pintarBarraPrealta(a);
    const pintor = PINTORES_PASO_PREALTA[a.paso]; if (pintor) pintor(a);
    pintarPiePrealta(a);
    pintarResumenPrealta(carrierPrealta());
}
const carrierPrealta = () => $('paCarrier').value ? porId(estado.carriers, $('paCarrier').value) : null;
/** «3 de 4» de las casillas de unidades o choferes; vacío si el carrier no tiene. */
function cuentaMarcadosPrealta(cont) {
    const t = $(cont).querySelectorAll('input').length;
    return t ? `${$(cont).querySelectorAll('input:checked').length} de ${t}` : '';
}
/** Un cliente escrito que no es de los conocidos es «Otro cliente». Lo llama quien carga el cliente en la forma, no el pintor. */
function marcarOtroClientePrealta() {
    const a = estado.paAsis, cli = parteTitulo('paCliente'); if (!a) return;
    a.otro = a.otro || (!!cli && !clientesPrealta(estado.prealtas).some(c => c.clave === cli));
}
function pintarBarraPrealta(a) {
    const n = a.paso;
    for (const b of $('paPasos').querySelectorAll('button')) {
        const i = Number(b.dataset.p), hecho = i !== n && i <= a.max && !faltaPrealta(i) && (i < n || a.revisar);
        b.className = i === n ? 'ahora' : hecho ? 'hecho' : i <= a.max ? 'visto' : '';
        b.disabled = i > a.max; if (i === n) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
        b.querySelector('small').textContent = hecho ? resumenPasoPrealta(i) : '';
    }
    for (let i = 1; i <= 5; i++) $('paPaso' + i).hidden = i !== n;
    $('paPasoK').textContent = `Paso ${n} de 5 · ${PASOS_PREALTA[n - 1]}${n === 4 ? ' · opcional' : ''}`;
}
/** Paso 1: clientes conocidos (por número de programas) y «Otro cliente». */
function pintarPaso1Prealta(a) {
    const cli = parteTitulo('paCliente');
    pintarOpcionesPa($('paClientes'), [...clientesPrealta(estado.prealtas).map(c => ({
        valor: c.clave, sel: !a.otro && cli === c.clave, titulo: c.clave, dato: plural(c.n, 'programa'),
        detalle: `${c.ultimo.Generador || 'sin generador'} · último ${c.ultimo.Campana || c.ultimo.Title}${c.ultimo.Pozo ? ', ' + c.ultimo.Pozo : ''}`,
        alClic: () => elegirClientePrealta(c.clave, c.ultimo.id)   // C-67: la clave y el id, no el objeto
    })), { valor: '__otro', sel: a.otro, titulo: 'Otro cliente', detalle: 'se captura una vez y queda para la próxima', clase: 'otro',
        alClic: () => { if (!a.otro) { a.otro = true; $('paCliente').value = ''; a.genDeId = null; armarTituloPrealta(); pintarAsistentePrealta(); } $('paCliente').focus(); } }]);
    $('paClienteOtro').classList.toggle('oculto', !a.otro);
}
/** Paso 2: la base, la corriente y el generador (aviso con «Cambiar», o sus campos si no hay o se pidió cambiarlo). */
function pintarPaso2Prealta(a) {
    const base = a.baseId ? porId(estado.prealtas, a.baseId) : null, genDe = a.genDeId ? porId(estado.prealtas, a.genDeId) : null;
    $('paBaseAviso').classList.toggle('oculto', !base);
    if (base) $('paBaseAviso').firstChild.textContent = $('paCarrier').value ? `Copiado de ${base.Title}: cliente, generador, corriente y transporte. Solo escribe el pozo.`
        : `Copiado de ${base.Title}: cliente, generador y corriente. Su carrier ya no está activo: escribe el pozo y elige otro en Transporte.`;   // U-94
    const corr = $('paCorriente').value;
    pintarOpcionesPa($('paCorrientes'), CORRIENTES.map(([k, t]) => ({ valor: k, sel: corr === k, titulo: t, detalle: k.startsWith('fluidos') ? 'fluido de perforación agotado' : 'recorte de perforación',
        alClic: () => { cambiarCorrientePrealta(k, t); } })));
    const gen = valorPa('paGenerador'), verCampos = a.genEdit || !gen;
    $('paGenAviso').classList.toggle('oculto', verCampos);
    $('paGenCampos').classList.toggle('oculto', !verCampos);
    if (!verCampos) {
        const reg = valorPa('paGeneradorRegistro');
        $('paGenAviso').firstChild.textContent = `Generador ${gen}${reg ? `, registro ${reg}` : ''}.${genDe ? ` Tomado de ${genDe.Campana || genDe.Title}.` : ''}`;
    }
}
/** Paso 3: el carrier se coteja ANTES de elegirlo (P5): vigencia a la vista; si su oficio no ampara la corriente, apagado. */
function pintarPaso3Prealta() {
    const cid = $('paCarrier').value, corr = $('paCorriente').value;
    const carriers = [...$('paCarrier').options].filter(o => o.value).map(o => porId(estado.carriers, o.value)).filter(Boolean);
    const elegido = carrierPrealta();
    // C-62 (v0.56.0): el pintor ya no suelta el carrier (lo hace cambiarCorrientePrealta, que avisa); un carrier que no ampara
    // —p. ej. al editar un borrador viejo— lo marca faltaPrealta(3) y Guardar no pasa.
    const cuenta = (col, c) => col.filter(x => Number(x.CarrierId) === Number(c.id) && x.Activo !== false).length;
    pintarOpcionesPa($('paCarriers'), carriers.map(c => {
        const ok = carrierAmpara(c, corr);
        return { valor: String(c.id), sel: String(c.id) === cid, titulo: c.Title, deshabilitada: !ok,
            dato: ok ? etiquetaVigencia('carriers', c) || undefined : etiqueta('no ampara', 'vencida'),
            detalle: ok ? `${plural(cuenta(estado.unidades, c), 'unidad', 'unidades')} · ${plural(cuenta(estado.choferes, c), 'chofer', 'choferes')}` : `Su oficio no ampara «${etiquetaCorriente(corr).toLowerCase()}»`,
            alClic: () => { if ($('paCarrier').value !== String(c.id)) { $('paCarrier').value = String(c.id); pintarUnidadesChoferesPrealta(); } pintarAsistentePrealta(); } };   // C-62: el mismo carrier no re-marca
    }));
    if (!carriers.length) $('paCarriers').appendChild(el('p', 'pista', 'No hay carriers activos en el padrón: da de alta uno.'));
    const h = elegido ? vigenciasPadron('carriers', elegido).filter(x => !sinFecha(x)) : [];
    $('paCarrierAviso').classList.toggle('oculto', !h.length);
    $('paCarrierAviso').classList.toggle('mal', h.some(x => !x.ok));
    // U-99 (v0.57.0): el remate de rechazo solo con algo ya vencido; «por vencer» (ok: true) dice su fecha límite, que ya trae el detalle.
    if (h.length) $('paCarrierAviso').firstChild.textContent = `${elegido.Title}: ${h.map(x => `${x.regla} ${x.detalle}`).join(' · ')}.${h.some(x => !x.ok) ? ' Si el envío llega con esto así, la puerta lo rechaza.' : ' Después de esa fecha la puerta lo rechaza.'}`;
    $('paTransporte').classList.toggle('oculto', !elegido);
    $('paNUnidades').textContent = cuentaMarcadosPrealta('paUnidades');
    $('paNChoferes').textContent = cuentaMarcadosPrealta('paChoferes');
}
/** Paso 5: Revisar, con «Cambiar» por sección (P9: cambias una y Siguiente te regresa aquí). */
function pintarPaso5Prealta(a) {
    const rv = $('paRevisar'); rv.textContent = '';
    const sec = (titulo, texto, paso, antes) => {
        const d = el('section'); d.appendChild(el('h3', '', titulo)); d.appendChild(el('p', '', texto));
        const b = el('button', 'enlace', 'Cambiar'); b.type = 'button'; b.addEventListener('click', () => { a.revisar = true; if (antes) antes(); irPasoPrealta(paso); }); d.appendChild(b);
        rv.appendChild(d);
    };
    const corr = $('paCorriente').value, elegido = carrierPrealta();
    const nom = valorPa('paTitulo'); const r0 = el('div', 'pa-nombre'); r0.appendChild(el('small', '', 'Programa')); r0.appendChild(el('b', 'mono', nom || '—')); rv.appendChild(r0);
    // U-90 (v0.57.0): cliente y generador en secciones aparte, porque viven en pasos distintos; «Cambiar» del generador abre
    // sus campos en el paso 2. U-104: lo que se guarda se repasa —dirección, pozo / instalación y notas incluidos—.
    sec('Cliente', parteTitulo('paCliente') || '—', 1);
    sec('Generador y origen', [valorPa('paGenerador') || 'sin generador', valorPa('paGeneradorRegistro'), valorPa('paGeneradorDireccion'), valorPa('paPozo') && `pozo / instalación ${valorPa('paPozo')}`].filter(Boolean).join(' · '), 2, () => { a.genEdit = true; });
    sec('Pozo y corriente', [parteTitulo('paPozoTitulo'), etiquetaCorriente(corr)].filter(Boolean).join(' · '), 2);
    sec('Transporte', elegido ? `${elegido.Title} · ${plural($('paUnidades').querySelectorAll('input:checked').length, 'unidad', 'unidades')} · ${plural($('paChoferes').querySelectorAll('input:checked').length, 'chofer', 'choferes')}` : '—', 3);
    sec('Envío y correo', [valorPa('paFecha') && `primer envío ${valorPa('paFecha')}`, valorPa('paGondolas') && plural(Number(valorPa('paGondolas')), 'góndola'), valorPa('paCorreoFecha') && `correo del ${valorPa('paCorreoFecha')}`, valorPa('paCorreoRemitente'), valorPa('paNotas') && `notas: ${valorPa('paNotas')}`].filter(Boolean).join(' · ') || 'sin datos (opcional)', 4);
}
/** P8: el botón dice qué falta; Atrás nombra a dónde vuelve. U-91 (v0.57.0): en Revisar, Guardar dice lo que le falte a 1–4. */
function pintarPiePrealta(a) {
    const n = a.paso, falta = n === 5 ? pendientePrealta() : faltaPrealta(n), sig = $('btnPaSiguiente'), gu = $('btnGuardarPrealta');
    $('btnPaAtras').hidden = n === 1; $('btnPaAtras').textContent = n > 1 ? `‹ ${PASOS_PREALTA[n - 2]}` : '';
    $('btnPaSaltar').hidden = n !== 4 || a.revisar;   // P7: el envío es opcional; desde «Cambiar» el botón ya regresa a Revisar
    sig.hidden = n === 5; gu.hidden = n !== 5;
    sig.classList.toggle('falta', n !== 5 && !!falta); gu.classList.toggle('falta', n === 5 && !!falta);
    sig.textContent = n === 5 ? '' : falta ? falta.t : a.revisar ? 'Volver a revisar ›' : n === 4 ? 'Revisar ›' : `Siguiente: ${PASOS_PREALTA[n]} ›`;
    if (n === 5) gu.textContent = falta ? falta.f.t : estado.prealtaEdit ? 'Guardar cambios' : 'Guardar para firma';
}
/** P4, escritorio: «Así va la pre-alta» a la derecha, en vivo. En celular no se ve: el nombre ya va arriba. */
function pintarResumenPrealta(carrier) {
    const r = $('paResumen'); r.textContent = '';
    r.appendChild(el('h2', '', 'Así va la pre-alta'));
    const dl = el('dl');
    const fila = (dt, dd) => { dl.appendChild(el('dt', '', dt)); const d = el('dd', dd ? '' : 'f'); if (dd instanceof Node) d.appendChild(dd); else d.textContent = dd || '—'; dl.appendChild(d); };
    const sec = t => dl.appendChild(el('span', 'sec', t));
    sec('Programa'); fila('Cliente', parteTitulo('paCliente')); fila('Pozo', parteTitulo('paPozoTitulo')); fila('Corriente', etiquetaCorriente($('paCorriente').value));
    sec('Generador'); fila('Razón social', valorPa('paGenerador')); fila('Registro', valorPa('paGeneradorRegistro'));
    sec('Transporte'); fila('Carrier', carrier ? carrier.Title : '');
    if (carrier) { fila('Unidades', cuentaMarcadosPrealta('paUnidades')); fila('Choferes', cuentaMarcadosPrealta('paChoferes')); }
    sec('Envío'); fila('Primer envío', valorPa('paFecha')); fila('Góndolas', valorPa('paGondolas'));
    r.appendChild(dl);
    r.appendChild(el('p', 'pista', 'Obligatorio: cliente, pozo, año, corriente, carrier y al menos una unidad y un chofer. Lo demás se puede completar después.'   /* U-131 (v0.73.0) */));
}
function enfocarPasoPrealta() { const a = estado.paAsis; if (a) $('paPaso' + a.paso).querySelector('.pregunta').focus({ preventScroll: true }); }
function irPasoPrealta(n) {
    const a = estado.paAsis; if (!a) return;
    a.paso = n; a.max = Math.max(a.max, n);
    limpiarAvisos(); pintarAsistentePrealta(); window.scrollTo({ top: 0 }); enfocarPasoPrealta();
}
function siguientePrealta() {
    const a = estado.paAsis; if (!a) return;
    const f = faltaPrealta(a.paso);
    if (f) { const x = f.foco(); if (x) x.focus(); return; }
    const p = pendientePrealta();
    irPasoPrealta(a.revisar ? (p ? p.paso : 5) : a.paso + 1);   // U-92: si «Cambiar» dejo otro paso incompleto, va ahi
}
/** El primer paso de 1–4 al que le falta algo, con lo que falta; null si todo está para guardarse (C-61, U-91). */
function pendientePrealta() {
    for (const paso of [1, 2, 3, 4]) { const f = faltaPrealta(paso); if (f) return { paso, f }; }
    return null;
}
/** U-100 (v0.57.0): Enter en un campo del paso visible es «Siguiente»; en Notas (textarea) sigue siendo un renglón nuevo. */
$('paAsis').addEventListener('keydown', ev => {
    const t = ev.target;
    if (ev.key !== 'Enter' || ev.isComposing || t.tagName !== 'INPUT' || t.type === 'checkbox' || !estado.paAsis) return;
    ev.preventDefault();
    if (estado.paAsis.paso === 5) guardarPrealta(); else siguientePrealta();
});
for (const id of PARTES_TITULO) $(id).addEventListener('input', armarTituloPrealta);
$('paCliente').addEventListener('input', marcarOtroClientePrealta);   // C-64: antes lo derivaba el pintor en cada repintado
$('paAsis').addEventListener('input', pintarAsistentePrealta);   // C-64: sin «change», que repintaba otra vez al salir de cada campo
$('btnPaSiguiente').addEventListener('click', siguientePrealta);
$('btnPaSaltar').addEventListener('click', siguientePrealta);   // el paso 4 no tiene obligatorios: saltar es seguir sin llenarlo
/** P7: − / + de góndolas esperadas; nunca baja de 0. Asignar .value no dispara «input», así que se repinta aquí. */
function sumarGondolasPrealta(d) {
    $('paGondolas').value = String(Math.max(0, (parseInt($('paGondolas').value, 10) || 0) + d));
    pintarAsistentePrealta();
}
$('btnPaGonMenos').addEventListener('click', () => sumarGondolasPrealta(-1));
$('btnPaGonMas').addEventListener('click', () => sumarGondolasPrealta(1));
$('btnPaAtras').addEventListener('click', () => irPasoPrealta(estado.paAsis.paso - 1));
for (const b of $('paPasos').querySelectorAll('button')) b.addEventListener('click', () => irPasoPrealta(Number(b.dataset.p)));
$('btnPaMiga').addEventListener('click', salirAsistentePrealta);
$('btnPaGenCambiar').addEventListener('click', () => { estado.paAsis.genEdit = true; pintarAsistentePrealta(); $('paGenerador').focus(); });
/** U-93 (v0.57.0): «Quitar la base» suelta lo que usarComoBase copió (cliente, generador, corriente, carrier) y vuelve al paso 1; lo tecleado —pozo, año— se queda. */
function quitarBasePrealta() {
    const a = estado.paAsis; if (!a) return;
    for (const id of ['paCliente', 'paGenerador', 'paGeneradorRegistro', 'paGeneradorDireccion']) $(id).value = '';
    $('paCorriente').value = ''; $('paCarrier').value = ''; pintarUnidadesChoferesPrealta();
    Object.assign(a, { baseId: null, genDeId: null, genEdit: false, otro: false });
    armarTituloPrealta(); irPasoPrealta(1);
}
$('btnPaQuitarBase').addEventListener('click', quitarBasePrealta);

export async function guardarPrealta() {
    if (!PUEDE.capturarPrealta(estado.rol)) { avisar('Tu rol no captura pre-altas.', 'error'); return; }
    if (!$('paTitulo').value.trim() || !$('paCorriente').value || !$('paCarrier').value) { avisar('Faltan cliente, pozo o año del programa, la corriente o el carrier.', 'error'); return; }
    // C-61 (v0.56.0): los pasos se revalidan al guardar. La barra de pasos deja llegar a Revisar sin pasar por Siguiente,
    // y unidades o choferes todos desmarcados se guardaban vacios, que la compuerta lee como «todas autorizadas».
    // U-91 (v0.57.0): el rango del año (U-54) y las fechas y góndolas del paso 4 viven ya en faltaPrealta: se salta al paso con el foco visible.
    const pend = pendientePrealta();
    if (pend) { irPasoPrealta(pend.paso); avisar(`${pend.f.t} antes de guardar.`, 'error'); const x = pend.f.foco(); if (x) x.focus(); return; }
    // C-59 (v0.56.0): se edita por el id que se abrio, aunque el refresco ya lo haya traido firmado: la relectura decide.
    // Antes, un borrador que otra sesion firmo caia al crearRenglon y nacia un programa duplicado.
    const edit = estado.prealtaEdit;
    await escribiendo('btnGuardarPrealta', async () => { try {   // C-24
        await refrescarCliente();
        if (edit) {
            const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, edit.id);
            if (fresco.Estado !== 'borrador') {
                aplicar('prealtas', edit, fresco);
                throw new Error(`ya no está por firmar: ${fresco.Estado === 'firmada' ? `la firmó ${quien(fresco.FirmadaPor) || 'otra sesión'}` : `la cerró ${quien(fresco.CerradaPor) || 'otra sesión'}`} mientras la editabas. Tus cambios no se guardaron; si hacen falta, cierra ese programa y captura otro.`);
            }
            const cambios = paraPatch(camposPrealta());   // C-65
            await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, edit.id, cambios);
            aplicar('prealtas', edit, cambios);   // C-23
            cerrarAsistentePrealta();
            if (estado.pestana === 'prealtas') pintarPrealtas(); else pintarInsignias();
            verPrealta(edit);
            avisar('Pre-alta actualizada. Sigue por firmar.', 'bien');   // U-101   // al final: abrir el detalle limpia los avisos
            return;
        }
        // C-04 (v0.22.0): la campana se escoge con la lista RELEIDA, no con la de memoria (que puede tener horas), y tras
        // escribir pasa por asegurarFolioUnico como R- y E-. Dos capturas en dos celulares recibian la misma L-AA-NNN.
        const todas = await estado.cliente.renglones(estado.siteId, L.prealtas);
        estado.prealtas = todas; reanclar();
        const campos = limpiar({
            ...camposPrealta(), Estado: 'borrador', CapturadaPor: estado.cuenta.username,   // C-65
            Campana: siguienteFolio('L', todas.map(p => p.Campana))   // L-AA-NNN (ticket 03)
        });
        const nuevo = await estado.cliente.crearRenglon(estado.siteId, L.prealtas, campos);
        await asegurarFolioUnico(nuevo, 'L');
        anclar('prealtas', nuevo);   // C-23
        cerrarAsistentePrealta();
        estado.vistaPrealtas = 'borrador';   // U-107 (v0.57.0): la lista abre donde quedó la nueva, no en la última pestaña elegida
        pintarPrealtas();
        mostrarListoPrealta(nuevo);
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}
/** P10: tras guardar una pre-alta nueva, la pestaña muestra qué sigue en vez de un aviso. «Capturar otra» abre el asistente vacío. */
function mostrarListoPrealta(p) {
    const carrier = porId(estado.carriers, p.CarrierId);
    $('paListoFolio').textContent = [p.Campana, p.Title].filter(Boolean).join(' · ');
    $('paListoPuerta').textContent = `Las góndolas de ${carrier ? carrier.Title : 'su carrier'} se cotejan solas contra este programa.`;
    $('p-prealtas').classList.add('en-listo'); $('paListo').classList.remove('oculto');
    window.scrollTo({ top: 0 });
    $('paListoTitulo').focus({ preventScroll: true });
}
export function ocultarListoPrealta() { $('paListo').classList.add('oculto'); $('p-prealtas').classList.remove('en-listo'); }
$('btnPaOtra').addEventListener('click', nuevaPrealta);
$('btnPaIrLista').addEventListener('click', () => { ocultarListoPrealta(); pintarPrealtas(); });

// Tras firmar / cerrar / eliminar: se cierra el pop-up y se repinta la pestana que esta abierta (Hoy o Pre-altas) sin
// borrar el aviso. Antes saltaba a Pre-altas aunque se hubiera abierto desde Hoy (Carlos, 2026-09-08).
function trasCambioPrealta() {
    cerrarForma('paDetalle'); repintar();
}
export function verPrealta(p) {
    estado.prealtaAbierta = p;
    estado.prealtaVista = { estado: p.Estado, huella: huellaPrealta(p) };   // C-60: lo que el detalle muestra, para cotejar al firmar
    abrirForma('paDetalle');   // antes de los avisos: con el dialog abierto, avisar() los pinta adentro
    $('paDetalleTitulo').textContent = `${p.Title} · ${estadoPrealta(p.Estado)}`;   // U-101
    const ul = $('paDetalleLista'); ul.textContent = '';
    const carrier = porId(estado.carriers, p.CarrierId);
    const filas = [
        ['Generador', `${p.Generador || '—'} · ${p.GeneradorRegistro || 'sin registro'}`], ['Pozo', p.Pozo || '—'], ['Corriente', etiquetaCorriente(p.Corriente) || '—'],
        ['Folio', p.Campana || '—'],   // U-129 (v0.73.0): el mismo L-AA-NNN que la lista llama Folio
        ['Góndolas', (({ rec, esp }) => esp ? `${rec} recibidas de ${esp} esperadas` : plural(rec, 'recibida'))(gondolasDe(p))],   // U-127 (v0.73.0): lo que se mira antes de «Cerrar el programa»
        ['Carrier', carrier ? `${carrier.Title} · ${carrier.AutorizacionASEA || 'sin autorización'}` : '—'],
        ['Unidades', lista(p.UnidadesIds).map(id => { const u = porId(estado.unidades, id); return u ? [u.Title, u.PlacaPlana].filter(Boolean).join('/') : `#${id}`;   /* U-128 */ }).join(', ') || '—'],
        ['Choferes', lista(p.ChoferesIds).map(id => nombreDe(estado.choferes, id)).join(', ') || '—'],
        ['Primer envío', fechaCorta(p.FechaEstimada)], ['Correo', [p.CorreoFecha ? fechaCorta(p.CorreoFecha) : '', p.CorreoRemitente].filter(Boolean).join(' · ') || '—'],   // U-128 (v0.73.0): sin separador colgando
        ['Capturó', quien(p.CapturadaPor) || '—'], ['Firmó', p.FirmadaPor ? `${quien(p.FirmadaPor)} · ${horaCorta(p.FirmadaEl)}${prealtaCambioTrasFirma(p) ? ' · cambió después de firmarse: la puerta no la ve hasta volver a firmar' : p.Estado === 'firmada' && !prealtaFirmada(p) ? ' · sello sin firma: la puerta no la ve' : ''}` : '—'], ['Cerró', p.CerradaPor ? `${quien(p.CerradaPor)} · ${horaCorta(p.CerradaEl)}` : '—'], ['Notas', p.Notas || '—']   // U-28 / U-31 (v0.26.0)
    ];
    for (const [k, v] of filas) { const li = el('li', '', k); li.appendChild(el('span', 'd', v)); if (k === 'Góndolas') li.lastChild.appendChild(barraAvance(gondolasDe(p))); ul.appendChild(li); }   // U-127: con la barra de la lista
    // Cotejo automatico de vigencias (lo que el validador firma que reviso).
    const vg = $('paDetalleVigencias'); vg.textContent = '';
    const hallazgos = [];
    if (carrier) {
        const c = evaluarVigencia('Autorización ASEA del carrier', carrier.VigenciaASEA, 'legal', CONFIG.avisoVigenciaDias); if (c) hallazgos.push(c);
        if (p.Corriente && lista(carrier.Corrientes).length && !lista(carrier.Corrientes).includes(p.Corriente)) hallazgos.push({ clase: 'legal', regla: 'Corriente', detalle: `el oficio del carrier no ampara ${etiquetaCorriente(p.Corriente)}` });
    } else hallazgos.push({ clase: 'legal', regla: 'Carrier', detalle: 'sin carrier' });
    for (const id of lista(p.UnidadesIds)) { const u = porId(estado.unidades, id); if (!u) continue;
        if (!u.FolioOficio && !(carrier && carrier.FolioOficio)) hallazgos.push({ clase: 'legal', regla: `Unidad ${u.Title}`, detalle: 'sin folio de oficio que la ampare (ni en la unidad ni en el carrier)' });
        for (const [n, f] of [['tarjeta', u.TarjetaVigencia], ['póliza', u.PolizaVigencia]]) { const h = evaluarVigencia(`Unidad ${u.Title} · ${n}`, f, 'comercial', CONFIG.avisoVigenciaDias); if (h) hallazgos.push(h); } }
    for (const id of lista(p.ChoferesIds)) { const ch = porId(estado.choferes, id); if (!ch) continue;
        const h = evaluarVigencia(`Chofer ${ch.Title} · licencia`, ch.LicenciaVigencia, 'comercial', CONFIG.avisoVigenciaDias); if (h) hallazgos.push(h); }
    if (!hallazgos.length) { const li = el('li', '', 'Vigencias '); li.appendChild(etiqueta('todo vigente', 'ok')); vg.appendChild(li); }
    for (const h of hallazgos) { const li = el('li', '', h.regla + ' '); li.appendChild(etiqueta(h.clase, h.clase)); li.appendChild(el('span', 'd', h.detalle)); vg.appendChild(li); }
    const hayLegal = hallazgos.some(h => h.clase === 'legal');
    const porFirmar = p.Estado === 'borrador' || (p.Estado === 'firmada' && !prealtaFirmada(p));   // S-01: el sello sin firma se firma aqui mismo
    $('btnFirmar').classList.toggle('oculto', !(porFirmar && PUEDE.firmarPrealta(estado.rol)));
    $('btnFirmar').title = motivoSinFirmas() || '';   // S-07
    $('btnEditarPrealta').classList.toggle('oculto', !(p.Estado === 'borrador' && PUEDE.capturarPrealta(estado.rol)));
    $('btnFirmar').disabled = hayLegal || !!estado.firmasError;   // S-07 + hallazgo legal (C-26: antes se asignaba dos veces)
    if (hayLegal && porFirmar) avisar('No se puede firmar con un hallazgo legal abierto: corrige el padrón (con el oficio a la vista) o cambia el carrier.', 'ojo');
    else if (porFirmar && p.Estado === 'firmada') avisar('Falta la firma del validador: trae el sello pero no la firma registrada (se escribió por fuera de la app o antes del corte). La puerta no la ve hasta que un validador o gerencia la firme.', 'ojo');   // U-31
    $('btnCerrarPrealta').classList.toggle('oculto', !(p.Estado === 'firmada' && PUEDE.capturarPrealta(estado.rol)));
    const sm = sinMovimientoDe(p);
    if (sm) avisar(`Este programa ${sm.motivo}. Sigue saliendo en la puerta hasta que se cierre; si ya no vienen más góndolas, ciérralo.`, 'ojo');
    // Un borrador equivocado se elimina; una firmada ya la vio la puerta y solo se CIERRA (2026-09-05).
    $('btnEliminarPrealta').classList.toggle('oculto', !(p.Estado === 'borrador' && PUEDE.corregir(estado.rol)));
    pintarCertificadoEnDetalle(p);   // v0.35.0
}

export async function eliminarPrealta() {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'borrador' || !PUEDE.corregir(estado.rol)) return;
    // La ventana de 90 dias no puede contestar esto: una pre-alta vieja tendria sus embarques
    // fuera de la carga y el borrador se borraria con historia colgando. Se pregunta EN VIVO.
    let citada;
    try { citada = (await estado.cliente.renglones(estado.siteId, L.embarques, `fields/PreAltaId eq ${p.id}`)).length; }
    catch (err) { avisar('No pude confirmar si tiene embarques (' + err.message + '). No se elimina: ciérrala.', 'error'); return; }
    if (citada) { avisar('Esta pre-alta ya tiene embarques: no se puede eliminar, ciérrala.', 'error'); return; }
    const { ok } = await confirmar({ titulo: 'Eliminar la pre-alta', peligro: true, ok: 'Eliminar',
        texto: `«${p.Title}» nunca se firmó, así que la puerta no la ha usado. Se borra el renglón; la campaña ${p.Campana || ''} queda libre.`, motivo: 'opcional', etiquetaMotivo: 'Por qué (opcional)' });
    if (!ok) return;
    try {
        await refrescarCliente();
        await estado.cliente.borrarRenglon(estado.siteId, L.prealtas, p.id);
        estado.prealtas = estado.prealtas.filter(x => x.id !== p.id);
        estado.prealtaAbierta = null;
        avisar('Pre-alta eliminada.', 'bien'); trasCambioPrealta();
    } catch (e) { avisar('No se pudo eliminar: ' + e.message, 'error'); }
}

export async function firmarPrealta() {
    const p = estado.prealtaAbierta; if (!p) return;
    await escribiendo('btnFirmar', async () => {   // C-24 / C-23
    const reFirma = p.Estado === 'firmada';   // S-01: trae el sello pero no su renglon en PLANTA_Firmas; solo falta la firma
    const { ok } = await confirmar({ titulo: 'Firmar la pre-alta', ok: 'Firmar',
        texto: reFirma && prealtaCambioTrasFirma(p) ? `«${p.Title}» cambió después de que ${quien(p.FirmadaPor)} la firmó, así que la puerta ya no la ve. Revisa carrier, unidades, choferes y generador arriba: con tu firma queda amparado lo que dice HOY.`
            : reFirma ? `«${p.Title}» trae el sello de ${quien(p.FirmadaPor)} pero no su firma registrada, así que la puerta no la ve. Con tu firma queda completa; queda registrado quién y cuándo.`
            : `«${p.Title}». Con tu firma la puerta empieza a aceptar sus góndolas; queda registrado quién y cuándo.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // C-60 (v0.56.0): se relee el renglón antes de firmar. Si otra sesión lo cerró, lo firmó o le cambió lo que se firma
        // mientras el detalle estaba abierto, no se firma lo que no se vio: se ancla lo fresco y se pide volver a mirarlo.
        const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, p.id);
        const vista = estado.prealtaVista || { estado: p.Estado, huella: huellaPrealta(p) };
        const cambio = fresco.Estado !== vista.estado || huellaPrealta(fresco) !== vista.huella;
        if (cambio) { aplicar('prealtas', p, fresco); trasCambioPrealta(); verPrealta(vivo('prealtas', p)); throw new Error(fresco.Estado === 'cerrada' ? 'el programa ya está cerrado (lo cerró otra sesión).'
            : fresco.Estado === 'firmada' && !reFirma ? `ya lo firmó ${quien(fresco.FirmadaPor) || 'otra sesión'}.` : 'cambió desde que lo abriste. Ábrelo otra vez y revísalo antes de firmar.'); }
        await firmar('prealta', p, huellaPrealta(fresco));   // S-01: primero la firma (403 si no eres validador/gerencia), luego el sello; S-27: con la huella de lo firmado
        const campos = { Estado: 'firmada', FirmadaPor: estado.cuenta.username, FirmadaEl: new Date().toISOString() };
        await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, campos);
        aplicar('prealtas', p, campos);   // C-23
        avisar('Pre-alta firmada.', 'bien'); trasCambioPrealta();
    } catch (e) { avisar('No se pudo firmar: ' + e.message, 'error'); }
    });
}
export async function cerrarPrealta() {
    const p = estado.prealtaAbierta; if (!p) return;
    await escribiendo('btnCerrarPrealta', async () => {   // C-24 / C-23
    const { ok } = await confirmar({ titulo: 'Cerrar el programa', ok: 'Cerrar el programa', peligro: true,   // U-130 (v0.73.0)
        texto: `«${p.Title}». La puerta dejará de aceptar góndolas contra él. No se borra: queda como historial.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // C-60 (v0.56.0): relectura. Cerrar lo ya cerrado no pisa el sello de quien cerró primero.
        const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, p.id);
        if (fresco.Estado === 'cerrada') { aplicar('prealtas', p, fresco); trasCambioPrealta(); avisar(`Ya estaba cerrado: lo cerró ${quien(fresco.CerradaPor) || 'otra sesión'}.`, 'ojo'); return; }
        const campos = { Estado: 'cerrada', CerradaPor: estado.cuenta.username, CerradaEl: new Date().toISOString() };
        try { await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, campos); }
        catch (err) {
            // Lista sin las columnas CerradaPor/CerradaEl (tarea 10 del setup): se cierra igual, sin sello, y se avisa.
            if (!esColumnaFaltante(err)) throw err;   // C-15
            await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, { Estado: 'cerrada' });
            delete campos.CerradaPor; delete campos.CerradaEl;
            avisar('Programa cerrado, pero sin registrar quién lo cerró' + (estado.rol === 'gerencia' ? ': la lista PLANTA_Prealtas no tiene todavía CerradaPor / CerradaEl (setup, tarea 10).' : '. Avisa a gerencia.'), 'ojo');   // U-31
        }
        aplicar('prealtas', p, campos); trasCambioPrealta();   // C-23
    } catch (e) { avisar('No se pudo cerrar: ' + e.message, 'error'); }
    });
}


// U-10 (v0.23.0): «+ Alta de carrier» desde la pre-alta abre la forma del padron ENCIMA (dialog anidado: la pre-alta
// capturada se queda atras, intacta) y al guardar el carrier nuevo queda elegido aqui, con sus unidades/choferes (vacios).
// Antes habia que Cancelar (se perdia todo), ir a Padron, dar de alta y reteclear los bloques 1 y 2.
export function elegirCarrierEnPrealta(c) {
    opciones($('paCarrier'), estado.carriers.filter(x => x.Activo !== false || x.id === c.id), x => x.id, x => x.Title);
    $('paCarrier').value = String(c.id);
    pintarUnidadesChoferesPrealta(); pintarAsistentePrealta();
}
