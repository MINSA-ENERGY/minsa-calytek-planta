// CALYTEK Planta — Archivos.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { esLoteDeLaApp, fechaCorta, plural, tipoDeArchivo } from './reglas.js';
import { $, avisar, el, estado, normaliza, opciones, porId } from './nucleo.js';

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

$('btnArLeerTodas').addEventListener('click', leerCarpetasSinLeer);
// v0.66.0: cada conteo del padrón abre su grupo y lo trae a la vista.
