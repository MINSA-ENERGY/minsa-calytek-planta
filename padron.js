// CALYTEK Planta — el padron por carrier.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { aIsoDia, etiquetaCorriente, evaluarVigencia, fechaCorta, fechaMexico, limpiar, lista, paraPatch, placaNormal, plural, PUEDE } from './reglas.js';
import { $, abrirForma, anclar, aplicar, asistentePadronAbierto, asistentePrealtaAbierto, avisar, cerrarForma, confirmar, el, esColumnaFaltante, escribiendo, estado, etiqueta, filtroTexto, hayCaptura, huellaForma, L, limpiarAvisos, nombreDe, normaliza, opciones, porId, quien, refrescarCliente, salir, textoDe, vivo } from './nucleo.js';
import { elegirCarrierEnPrealta } from './prealtas.js';

// ================================================================ PADRON

// Los grupos (pre-altas y padron) arrancan PLEGADOS en todo tamano (pedido de Carlos 2026-09-07):
// el usuario abre el que quiera y la app respeta lo que dejo abierto o cerrado al repintar.
// Ficha completa de un renglón del padrón: TODAS las columnas de la lista, no el resumen de una línea (Carlos, 2026-09-06).
// Se despliega al tocar el renglón; el resumen sigue siendo lo que se skimea.
// Vigencias que trae un renglón del padrón (vencida, por vencer o sin fecha), con la misma regla que la puerta.
export const VIGENCIAS_PADRON = {
    carriers: [['autorización ASEA', 'VigenciaASEA', 'legal'], ['CSF', 'CSFVigencia', 'comercial']],
    unidades: [['tarjeta de circulación', 'TarjetaVigencia', 'comercial'], ['póliza', 'PolizaVigencia', 'comercial']],
    choferes: [['licencia', 'LicenciaVigencia', 'comercial']],
};
export function vigenciasPadron(tipo, it) {
    return VIGENCIAS_PADRON[tipo].map(([n, col, cl]) => evaluarVigencia(n, it[col], cl, CONFIG.avisoVigenciaDias)).filter(Boolean);
}
export const sinFecha = h => h.detalle.startsWith('sin fecha');
// La etiqueta del renglón: lo peor que traiga (vencida > sin fecha > por vencer).
export function etiquetaVigencia(tipo, it) {
    const h = vigenciasPadron(tipo, it);
    const peor = h.find(x => !x.ok && !sinFecha(x)) || h.find(sinFecha) || h[0];
    // Vencida es roja siempre: aquí no se juzga la puerta (legal/comercial), se avisa que el papel ya no sirve.
    if (!peor) return null;
    if (!peor.ok && !sinFecha(peor)) return etiqueta('vencida', 'vencida');
    return etiqueta(sinFecha(peor) ? 'sin fecha' : 'por vencer', sinFecha(peor) ? peor.clase : 'aviso');
}
// El aviso del alta dice qué se guardó Y qué trae vencido o por vencer, para que quien captura no se entere hasta la puerta.
function avisarAlta(texto, tipo, it) {
    const h = vigenciasPadron(tipo, it).filter(x => !sinFecha(x));
    if (!h.length) { avisar(texto, 'bien'); return; }
    const vencidas = h.filter(x => !x.ok);
    avisar(`${texto} ⚠ ${h.map(x => `${x.regla} ${x.detalle}`).join(' · ')}`, vencidas.length ? 'error' : 'ojo');
}
const CAMPOS_PADRON = {
    carriers: [['Autorización ASEA', 'AutorizacionASEA'], ['Vence ASEA', 'VigenciaASEA', 'fecha'], ['Corrientes', 'Corrientes', 'lista'], ['Folio del oficio', 'FolioOficio'], ['Registro SCT', 'RegistroSCT'], ['CSF vence', 'CSFVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
    unidades: [['Placa plana', 'PlacaPlana'], ['Carrier', 'CarrierId', 'carrier'], ['Tipo', 'TipoUnidad'], ['Marca', 'Marca'], ['No. de serie', 'NumeroSerie'], ['Capacidad', 'CapacidadKg', 'kg'], ['Folio del oficio', 'FolioOficio'], ['Tarjeta de circulación', 'TarjetaCirc'], ['Tarjeta vence', 'TarjetaVigencia', 'fecha'], ['Póliza', 'Poliza'], ['Póliza vence', 'PolizaVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
    choferes: [['Carrier', 'CarrierId', 'carrier'], ['Licencia', 'Licencia'], ['Licencia vence', 'LicenciaVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
};
export const ESCRITORIO = window.matchMedia('(min-width: 900px)');
// U-11 (v0.22.0): en el celular el campo de peso NO levanta el teclado del sistema (inputmode=none): se captura con el
// teclado propio, que quedaba tapado por el del telefono. En escritorio el propio se oculta por CSS y el campo vuelve a numeric.
function ajustarTecladoPeso() { $('baKg').inputMode = ESCRITORIO.matches ? 'numeric' : 'none'; }
ESCRITORIO.addEventListener('change', ajustarTecladoPeso); ajustarTecladoPeso();
/**
 * v0.72.0 — EL PADRÓN POR CARRIER (artifact G77viANSEuwDFfDvkxVWFS, OK de Carlos 2026-09-24). Tres vistas en la pestaña, con la
 * banda de Pre-altas: 'lista' (carriers por estado, o la búsqueda de placa/chofer), 'carrier' (el expediente: unidades y choferes
 * en pestañas y el oficio a un lado) y 'ficha' (un renglón, con «Dar de baja» / «Eliminar» / «Reactivar» grande abajo). Los
 * estados son etiquetas (texto mono de color, sin píldora). estado.padronCarrier sigue siendo el carrier elegido que preselecciona
 * el alta de unidad/chofer (U-46): en el expediente y en la ficha es el de la vista; en la lista, ninguno.
 */
export const activo = x => x.Activo !== false;
const porTitulo = (a, b) => String(a.Title || '').localeCompare(String(b.Title || ''), 'es');
const nombrePd = (clave, x) => clave === 'unidades' ? `${x.Title}${x.PlacaPlana ? ' / ' + x.PlacaPlana : ''}` : x.Title;
const ARTICULO_PADRON = { carriers: 'el carrier', unidades: 'la unidad', choferes: 'el chofer' };
function tipoUnidadTexto(v) { const o = [...$('puuTipo').options].find(o => o.value === v); return o ? o.textContent : (v || '—'); }
/** Lo que un renglón trae vencido o por vencer (sin contar «sin fecha», igual que el resumen de antes). */
function avisosPadron(clave, x) { return vigenciasPadron(clave, x).filter(h => !sinFecha(h)); }
function delCarrier(clave, c) { return estado[clave].filter(x => Number(x.CarrierId) === c.id); }
/** Por atender = el oficio del carrier o la póliza/tarjeta/licencia de alguna de sus unidades o choferes ACTIVOS, vencida o por vencer. */
function pendientesCarrier(c) {
    return [['carriers', c], ...delCarrier('unidades', c).map(u => ['unidades', u]), ...delCarrier('choferes', c).map(h => ['choferes', h])]
        .filter(([clave, x]) => activo(x) && avisosPadron(clave, x).length);
}
function grupoCarrier(c) { return !activo(c) ? 'baja' : pendientesCarrier(c).length ? 'atender' : 'vigentes'; }
/** La etiqueta de un renglón: de baja > vencida > por vencer > sin fecha > vigente. */
function estadoPadron(clave, x) {
    if (!activo(x)) return etiqueta('de baja', 'baja');
    const h = vigenciasPadron(clave, x);
    const mala = h.find(v => !v.ok && !sinFecha(v)); if (mala) return etiqueta(`${mala.regla} vencida`, 'vencida');
    const aviso = h.find(v => v.ok); if (aviso) return etiqueta(`${aviso.regla} ${aviso.detalle.replace(/ \(.*\)$/, '')}`, 'aviso');
    const sf = h.find(sinFecha); if (sf) return etiqueta(`${sf.regla} sin fecha`, sf.clase);
    return etiqueta(clave === 'unidades' ? 'amparada' : 'vigente', 'ok');
}
function usoPadron(clave, x) {
    const num = Number(x.id), ids = { unidades: 'UnidadesIds', choferes: 'ChoferesIds' }[clave];
    const pre = clave === 'carriers' ? estado.prealtas.filter(p => Number(p.CarrierId) === num) : estado.prealtas.filter(p => lista(p[ids]).includes(String(x.id)));
    const emb = estado.embarques.filter(e => Number(e[COLUMNA_PADRON[clave]]) === num);
    return { prealtas: pre.length, gondolas: emb.length, ultima: emb.map(e => e.Arribo).filter(Boolean).sort().pop() || null };
}

// ---- piezas de pintura
const celdaPd = (clase, v) => { const s = el('span', clase); if (v instanceof Node) s.appendChild(v); else s.textContent = v === null || v === undefined || v === '' ? '—' : String(v); return s; };
function filaPd(celdas, alAbrir, baja) {
    const f = el('div', 'pd-fila' + (baja ? ' baja' : '')); f.setAttribute('role', 'button'); f.tabIndex = 0;
    for (const c of celdas) f.appendChild(c);
    f.addEventListener('click', alAbrir);
    f.addEventListener('keydown', ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); alAbrir(); } });
    return f;
}
function tablaPd(cont, cols, cabeza, filas, vacio) {
    cont.textContent = ''; cont.style.setProperty('--pd-cols', cols);
    const cab = el('div', 'pd-fila pd-cab'); cab.setAttribute('aria-hidden', 'true');
    for (const [t, c] of cabeza) cab.appendChild(el('span', c || '', t));
    cont.appendChild(cab);
    if (!filas.length) cont.appendChild(el('p', 'vacio', vacio));
    for (const f of filas) cont.appendChild(f);
}
function migasPd(partes) {
    const m = $('pdMigas'); m.textContent = ''; m.hidden = !partes.length;
    partes.forEach(([t, alClic], i) => {
        if (i) { const s = el('span', '', '/'); s.setAttribute('aria-hidden', 'true'); m.appendChild(s); }
        if (alClic) { const b = el('button', 'miga', t); b.type = 'button'; b.addEventListener('click', alClic); m.appendChild(b); }
        else { const b = el('b', '', t); b.setAttribute('aria-current', 'page'); m.appendChild(b); }
    });
}
function kpisPd(items) {
    const k = $('pdKpis'); k.textContent = ''; k.hidden = !items.length;
    for (const [n, t, puesto, alerta, alClic] of items) {
        const b = el('button'); b.type = 'button'; b.setAttribute('aria-pressed', String(puesto)); if (alerta) b.classList.add('alerta');
        b.appendChild(el('b', '', String(n))); b.appendChild(el('span', '', t)); b.addEventListener('click', alClic); k.appendChild(b);
    }
}
function cabeceraPd(migas, titulo, etq, sub) {
    migasPd(migas);
    const h = $('pdTitulo'); h.textContent = titulo; if (etq) h.appendChild(etq);
    $('pdSub').textContent = sub;
}
function dlPd() {
    const dl = el('dl');
    return { dl, sec: t => dl.appendChild(el('span', 'sec', t)), fila: (dt, dd) => {
        dl.appendChild(el('dt', '', dt));
        const d = el('dd', dd === null || dd === undefined || dd === '' ? 'f' : '');
        if (dd instanceof Node) d.appendChild(dd); else d.textContent = dd === null || dd === undefined || dd === '' ? '—' : String(dd);
        dl.appendChild(d);
    } };
}
function botonPd(a, clase, texto) { const b = el('button', clase, texto || a.texto); b.type = 'button'; b.dataset.accion = a.accion; b.addEventListener('click', a.alClic); return b; }

// ---- navegación
export function irPadron(v, extra = {}) { Object.assign(estado.padronVista, { v }, extra); pintarPadron(); window.scrollTo({ top: 0 }); }
function volverPadron() {
    const p = estado.padronVista;
    if (p.v === 'ficha' && p.desde !== 'buscar' && porId(estado.carriers, p.carrier)) irPadron('carrier');
    else irPadron('lista');
}

export function pintarPadron() {
    const puede = PUEDE.capturarPrealta(estado.rol);
    const p = estado.padronVista;
    if (p.v === 'asis' && !p.asis) p.v = 'lista';
    if (p.v !== 'asis') {   // v0.72.0: con el asistente abierto, repintar no suelta la edición ni cambia las opciones del carrier elegido
        for (const id of ['pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer']) cerrarForma(id);
        estado.padronEdit = null;
        for (const s of [$('puuCarrier'), $('pchCarrier')]) opciones(s, estado.carriers.filter(activo), c => c.id, c => c.Title);
    }
    // Lo que se veía pudo irse (se eliminó, o el refresco ya no lo trae): se cae a la vista de arriba.
    if (p.v === 'ficha' && !(p.ficha && porId(estado[p.ficha.clave], p.ficha.id))) p.v = porId(estado.carriers, p.carrier) && p.desde !== 'buscar' ? 'carrier' : 'lista';
    if (p.v === 'carrier' && !porId(estado.carriers, p.carrier)) p.v = 'lista';
    const c = p.v === 'lista' ? null : porId(estado.carriers, p.carrier);   // en el asistente: el carrier desde el que se abrió, si hay
    estado.padronCarrier = c ? c.id : null;
    $('pdLista').hidden = p.v !== 'lista'; $('pdExpediente').hidden = p.v !== 'carrier'; $('pdFicha').hidden = p.v !== 'ficha'; $('pdAsis').hidden = p.v !== 'asis';
    $('btnPdVolver').classList.toggle('oculto', p.v === 'lista' || p.v === 'asis');
    $('btnPdCancelar').classList.toggle('oculto', p.v !== 'asis');
    $('btnNuevoCarrier').classList.toggle('oculto', !puede || p.v !== 'lista');
    const altas = puede && p.v === 'carrier' && activo(c);
    $('btnNuevaUnidad').classList.toggle('oculto', !(altas && p.sub === 'unidades'));
    $('btnNuevoChofer').classList.toggle('oculto', !(altas && p.sub === 'choferes'));
    if (p.v === 'lista') pintarListaPd(); else if (p.v === 'carrier') pintarExpedientePd(c); else if (p.v === 'ficha') pintarFichaPd(c); else pintarAsistentePd(c);
}

function pintarListaPd() {
    const p = estado.padronVista, cs = [...estado.carriers].sort(porTitulo);
    const n = { vigentes: 0, atender: 0, baja: 0 }; for (const c of cs) n[grupoCarrier(c)]++;
    cabeceraPd([], 'Padrón', null, 'Un expediente por carrier: su oficio ASEA, sus unidades y sus choferes. Una placa que no está aquí no está amparada.');
    const tab = t => () => { p.tab = t; pintarPadron(); };
    kpisPd([[n.vigentes, 'Vigentes', p.tab === 'vigentes', false, tab('vigentes')], [n.atender, 'Por atender', p.tab === 'atender', n.atender > 0, tab('atender')], [n.baja, 'De baja', p.tab === 'baja', false, tab('baja')]]);
    $('pdnVigentes').textContent = n.vigentes || ''; $('pdnAtender').textContent = n.atender || ''; $('pdnBaja').textContent = n.baja || '';
    $('pdnAtender').classList.toggle('alerta', n.atender > 0);
    for (const b of $('pdTabs').querySelectorAll('button')) b.setAttribute('aria-pressed', String(b.dataset.pd === p.tab));
    $('pdpCarriers').hidden = p.tab === 'buscar'; $('pdpBuscar').hidden = p.tab !== 'buscar';
    if (p.tab === 'buscar') { pintarBusquedaPd(); return; }
    const filas = cs.filter(c => grupoCarrier(c) === p.tab).map(c => {
        const pend = pendientesCarrier(c);
        const est = !activo(c) ? etiqueta('de baja', 'baja')
            : pend.length ? etiqueta(`${pend.length} por atender`, pend.some(([k, x]) => avisosPadron(k, x).some(h => !h.ok)) ? 'vencida' : 'aviso')
            : etiqueta('vigente', 'ok');
        const a = el('span', 'a'); a.appendChild(el('b', 'n', c.Title)); a.appendChild(el('small', '', c.FolioOficio ? `Oficio ${c.FolioOficio}` : 'Sin folio de oficio'));
        return filaPd([a, celdaPd('x m', c.AutorizacionASEA), celdaPd('x m', delCarrier('unidades', c).filter(activo).length), celdaPd('x m', delCarrier('choferes', c).filter(activo).length), celdaPd('e', est)],
            () => irPadron('carrier', { carrier: c.id, sub: 'unidades' }), !activo(c));
    });
    const vacio = { vigentes: estado.carriers.length ? 'Ningún carrier sin pendientes.' : 'Sin carriers. La primera pre-alta necesita uno con su oficio ASEA transcrito: «+ Nuevo carrier».',
        atender: 'Nada por atender: todo el padrón está vigente.', baja: 'Ningún carrier de baja.' }[p.tab];
    tablaPd($('pdCarriers'), 'minmax(0,1.8fr) minmax(0,1fr) 4rem 4rem minmax(10rem,auto)', [['Carrier'], ['Autorización', 'x'], ['Unid.', 'x'], ['Chof.', 'x'], ['Estado', 'e']], filas, vacio);
    $('pdPieLista').textContent = p.tab === 'atender'
        ? `Un carrier cae aquí si su oficio, o la póliza, tarjeta o licencia de alguna de sus unidades o choferes, vence en ${CONFIG.avisoVigenciaDias} días o ya venció.`
        : filas.length ? 'El renglón abre el expediente del carrier.' : '';
}

function pintarBusquedaPd() {
    const texto = $('pdBusca').value.trim(), q = normaliza(texto), pega = filtroTexto(q);   // C-26
    const r = !q ? [] : [
        ...estado.unidades.filter(u => pega(u.Title, u.PlacaPlana, u.NumeroSerie)).map(u => ['unidades', u]),
        ...estado.choferes.filter(h => pega(h.Title, h.Licencia)).map(h => ['choferes', h]),
        ...estado.carriers.filter(c => pega(c.Title, c.AutorizacionASEA, c.FolioOficio)).map(c => ['carriers', c]),
    ];
    $('pdBuscaCuenta').textContent = q ? plural(r.length, 'resultado') : '';
    const filas = r.map(([clave, x]) => filaPd([celdaPd('x m', NOMBRE_PADRON[clave]), celdaPd('a t', nombrePd(clave, x)),
        celdaPd('c', clave === 'carriers' ? (x.FolioOficio ? `Oficio ${x.FolioOficio}` : '') : nombreDe(estado.carriers, x.CarrierId)), celdaPd('e', estadoPadron(clave, x))],
        () => irPadron('ficha', { ficha: { clave, id: x.id }, carrier: clave === 'carriers' ? x.id : Number(x.CarrierId), desde: 'buscar' }), !activo(x)));
    tablaPd($('pdResultados'), '5rem minmax(0,1.2fr) minmax(0,1.4fr) minmax(10rem,auto)', [['Tipo', 'x'], ['Unidad, chofer o carrier'], ['Carrier'], ['Estado', 'e']], filas,
        q ? `Nada con «${texto}». Si la placa no está, no está amparada: se transcribe del oficio del carrier.` : 'Escribe una placa, una serie, un nombre o una licencia.');
}

function pintarExpedientePd(c) {
    const p = estado.padronVista, us = delCarrier('unidades', c).sort(porTitulo), hs = delCarrier('choferes', c).sort(porTitulo), pend = pendientesCarrier(c);
    cabeceraPd([['Padrón', () => irPadron('lista')], [c.Title]], c.Title, activo(c) ? null : etiqueta('de baja', 'baja'),
        [c.FolioOficio && `Oficio ${c.FolioOficio}`, c.AutorizacionASEA && `Autorización ${c.AutorizacionASEA}`].filter(Boolean).join(' · ') || 'Sin datos del oficio capturados.');
    const sub = s => () => { p.sub = s; pintarPadron(); };
    const subPend = pend.some(([k]) => k === 'unidades') || !pend.some(([k]) => k === 'choferes') ? 'unidades' : 'choferes';
    // U-137 (v0.73.0): «Por atender» filtra las dos tablas a lo pendiente y queda marcado; Unidades o Choferes lo sueltan.
    const soloPend = !!p.soloPend && pend.length > 0, conPend = new Set(pend.map(([, x]) => x));
    const ver = xs => soloPend ? xs.filter(x => conPend.has(x)) : xs;
    kpisPd([[us.filter(activo).length, 'Unidades', !soloPend && p.sub === 'unidades', false, () => { p.soloPend = false; sub('unidades')(); }], [hs.filter(activo).length, 'Choferes', !soloPend && p.sub === 'choferes', false, () => { p.soloPend = false; sub('choferes')(); }],
        [pend.length, 'Por atender', soloPend, pend.length > 0, () => { p.soloPend = !soloPend; if (p.soloPend) p.sub = subPend; pintarPadron(); }]]);
    $('pdnUnidades').textContent = us.length || ''; $('pdnChoferes').textContent = hs.length || '';
    for (const b of $('pdSubTabs').querySelectorAll('button')) b.setAttribute('aria-pressed', String(b.dataset.sub === p.sub));
    $('pdpUnidades').hidden = p.sub !== 'unidades'; $('pdpChoferes').hidden = p.sub !== 'choferes';
    const ficha = (clave, x) => () => irPadron('ficha', { ficha: { clave, id: x.id }, desde: 'carrier' });
    tablaPd($('pdUnidades'), 'minmax(0,1.4fr) 7rem minmax(0,1fr) minmax(10rem,auto)', [['Placas'], ['Tipo', 'x'], ['Serie', 'x'], ['Estado', 'e']],
        ver(us).map(u => filaPd([celdaPd('a t', nombrePd('unidades', u)), celdaPd('x m', tipoUnidadTexto(u.TipoUnidad)), celdaPd('x m', u.NumeroSerie ? '…' + String(u.NumeroSerie).slice(-6) : ''), celdaPd('e', estadoPadron('unidades', u))], ficha('unidades', u), !activo(u))),
        'Sin unidades. Se transcriben del oficio del carrier con «+ Alta de unidad».');
    tablaPd($('pdChoferes'), 'minmax(0,1.6fr) minmax(0,1fr) minmax(10rem,auto)', [['Nombre'], ['Licencia', 'x'], ['Estado', 'e']],
        ver(hs).map(h => filaPd([celdaPd('a n', h.Title), celdaPd('x m', h.Licencia), celdaPd('e', estadoPadron('choferes', h))], ficha('choferes', h), !activo(h))),
        'Sin choferes. Se dan de alta con su licencia con «+ Alta de chofer».');
    const o = $('pdOficio'); o.textContent = ''; o.appendChild(el('h2', '', 'Oficio ASEA'));
    const { dl, sec, fila } = dlPd();
    sec('Autorización'); fila('Número', c.AutorizacionASEA); fila('Vence', c.VigenciaASEA ? fechaCorta(c.VigenciaASEA) : ''); fila('Estado', estadoPadron('carriers', c));
    sec('Oficio'); fila('Folio', c.FolioOficio); fila('Corrientes', lista(c.Corrientes).map(etiquetaCorriente).join(', '));
    sec('Registro'); fila('SCT', c.RegistroSCT); fila('CSF vence', c.CSFVigencia ? fechaCorta(c.CSFVigencia) : '');
    o.appendChild(dl);
    const b = el('button', 'secundario', 'Ver la ficha del carrier'); b.type = 'button'; b.id = 'btnPdFichaCarrier';
    b.addEventListener('click', () => irPadron('ficha', { ficha: { clave: 'carriers', id: c.id }, desde: 'carrier' }));
    o.appendChild(b);
}

function pintarFichaPd(c) {
    const p = estado.padronVista, clave = p.ficha.clave, x = porId(estado[clave], p.ficha.id);
    const migas = [['Padrón', () => irPadron('lista')]];
    if (c) migas.push([c.Title, () => irPadron('carrier')]);
    migas.push([clave === 'carriers' ? 'Ficha del carrier' : nombrePd(clave, x)]);
    const sub = clave === 'carriers' ? `Carrier${x.FolioOficio ? ' · oficio ' + x.FolioOficio : ''}`
        : clave === 'unidades' ? `${c ? c.Title : 'Sin carrier'} · ${tipoUnidadTexto(x.TipoUnidad)}`
        : `${c ? c.Title : 'Sin carrier'} · licencia ${x.Licencia || 'sin capturar'}`;
    cabeceraPd(migas, nombrePd(clave, x), estadoPadron(clave, x), sub);
    kpisPd([]);
    const d = $('pdFichaDatos'); d.textContent = '';
    const acciones = botonesPadron(clave, x);
    const cab = el('div', 'pd-cabeza'); cab.appendChild(el('h2', '', { carriers: 'Datos del oficio', unidades: 'Vehículo', choferes: 'Chofer' }[clave]));
    const editar = acciones.find(a => a.accion === 'editar'); if (editar) cab.appendChild(botonPd(editar, 'secundario'));
    d.appendChild(cab);
    const dl = el('dl', 'pd-pares');
    for (const [rotulo, col, fmt] of CAMPOS_PADRON[clave]) {
        if (fmt === 'fecha' || col === 'Activo' || col === 'Notas' || (fmt === 'carrier')) continue;
        if (clave === 'unidades' && col === 'PlacaPlana') continue;   // U-140 (v0.73.0): ya va en el título («tractor / plana»)
        const v = x[col];
        const txt = fmt === 'lista' ? lista(v).map(etiquetaCorriente).join(', ') : fmt === 'kg' ? (v ? `${Number(v).toLocaleString('es-MX')} kg` : '') : col === 'TipoUnidad' ? tipoUnidadTexto(v) : v;
        const par = el('div', txt ? '' : 'vacio'); par.appendChild(el('dt', '', rotulo)); par.appendChild(el('dd', txt ? 'mono' : 'f', txt ? String(txt) : '—')); dl.appendChild(par);   // U-140: en celular los vacíos no se pintan
    }
    d.appendChild(dl);
    const vacios = [...dl.querySelectorAll('div.vacio dt')].map(x => x.textContent);   // U-140 (v0.73.0): en celular, los vacíos juntos en un renglón
    if (vacios.length) d.appendChild(el('p', 'pista pd-falta', `Sin capturar: ${vacios.join(', ')}.`));
    d.appendChild(el('h3', '', 'Vigencias'));
    const ul = el('ul', 'pd-vigs');
    for (const [n, col, cl] of VIGENCIAS_PADRON[clave]) {
        const h = evaluarVigencia(n, x[col], cl, CONFIG.avisoVigenciaDias);
        const li = el('li'); li.appendChild(el('b', '', n[0].toUpperCase() + n.slice(1))); li.appendChild(el('span', 'f', x[col] ? fechaCorta(x[col]) : 'sin fecha'));
        li.appendChild(!h ? etiqueta('vigente', 'ok') : sinFecha(h) ? etiqueta('sin fecha', h.clase) : etiqueta(h.detalle.replace(/ \(.*\)$/, ''), h.ok ? 'aviso' : 'vencida'));
        ul.appendChild(li);
    }
    d.appendChild(ul);
    if (!activo(x)) d.appendChild(el('p', 'pista', `De baja: la puerta ya no ${clave === 'unidades' ? 'la' : 'lo'} ampara y el historial ${clave === 'unidades' ? 'la' : 'lo'} sigue viendo.`));
    // La acción que no se deshace va sola, grande, abajo (como «Cerrar el programa» en el detalle de la pre-alta, v0.68.0).
    const grande = acciones.find(a => a.accion !== 'editar');
    if (grande) d.appendChild(botonPd(grande, 'pd-grande' + (grande.accion === 'reactivar' ? ' bien' : ''), `${grande.texto} ${ARTICULO_PADRON[clave]}`));
    const u = usoPadron(clave, x), a = $('pdUso'); a.textContent = ''; a.appendChild(el('h2', '', 'Dónde se ha usado'));
    const { dl: dlu, sec, fila } = dlPd();
    sec('Historial'); fila(`Pre-altas que ${clave === 'unidades' ? 'la' : 'lo'} citan`, String(u.prealtas)); fila(`Góndolas (últimos ${CONFIG.ventanaDias} días)`, String(u.gondolas));
    if (clave !== 'carriers') fila('Última entrada', u.ultima ? fechaCorta(u.ultima) : '');
    if (x.Notas) { sec('Notas'); const nt = el('p', 'pd-notas', String(x.Notas).replace(/[\w.+-]+@[\w.-]+\.\w+/g, m => quien(m) || m)); dlu.appendChild(nt); }   // U-141 (v0.73.0): se guarda el correo, se lee el nombre
    a.appendChild(dlu);
    a.appendChild(el('p', 'pista', referenciasPadron(clave, x) ? 'Ya lo cita el historial: no se elimina, se da de baja.' : 'Nada lo cita todavía: se puede eliminar si se transcribió mal.'));
}
/**
 * Corregir el padron (2026-09-05). No hay edicion en sitio: una placa mal transcrita se corrige
 * dando de baja el renglon y transcribiendola de nuevo del oficio (misma regla del padron §4).
 *   - sin referencias (ninguna pre-alta, embarque, unidad o chofer lo cita): se ELIMINA.
 *   - con referencias: se da de BAJA (Activo = false): la puerta deja de ampararlo, el historial
 *     sigue apuntando a un renglon que existe. Se puede reactivar.
 */
const COLUMNA_PADRON = { carriers: 'CarrierId', unidades: 'UnidadId', choferes: 'ChoferId' };

/** Referencias FUERA de embarques (padron y pre-altas se cargan enteros: esto es exacto). */
function referenciasSinEmbarques(clave, x) {
    const id = String(x.id), num = Number(x.id);
    if (clave === 'carriers') return estado.unidades.filter(u => Number(u.CarrierId) === num).length + estado.choferes.filter(c => Number(c.CarrierId) === num).length
        + estado.prealtas.filter(p => Number(p.CarrierId) === num).length;
    if (clave === 'unidades') return estado.prealtas.filter(p => lista(p.UnidadesIds).includes(id)).length;
    return estado.prealtas.filter(p => lista(p.ChoferesIds).includes(id)).length;
}
/** Cuenta con lo que hay en memoria: sirve para ELEGIR el boton, nunca para autorizar el borrado. */
function referenciasPadron(clave, x) {
    const num = Number(x.id), col = COLUMNA_PADRON[clave];
    return referenciasSinEmbarques(clave, x) + estado.embarques.filter(e => Number(e[col]) === num).length;
}
/**
 * La cuenta que SI autoriza el borrado: pregunta a SharePoint por TODOS los embarques que citan
 * al renglon, no por los de la ventana de 90 dias. Un carrier citado hace un anio se veria en
 * cero desde la memoria y el boton ofreceria Eliminar sobre historia viva (cubeta 3).
 * Si la consulta falla, esta funcion no contesta cero: la excepcion sube y el borrado se niega.
 */
async function referenciasPadronVivas(clave, x) {
    const citas = await estado.cliente.renglones(estado.siteId, L.embarques, `fields/${COLUMNA_PADRON[clave]} eq ${Number(x.id)}`);
    return referenciasSinEmbarques(clave, x) + citas.length;
}
function botonesPadron(clave, x) {
    const editar = x.Activo !== false && PUEDE.capturarPrealta(estado.rol) ? [{ texto: 'Editar', accion: 'editar', clase: 'suave', alClic: () => abrirFormaPadron(clave, vivo(clave, x)) }] : [];
    if (!PUEDE.corregir(estado.rol)) return editar;
    if (x.Activo === false) return [{ texto: 'Reactivar', accion: 'reactivar', clase: 'suave', alClic: () => activarPadron(clave, vivo(clave, x), true) }];
    return editar.concat(referenciasPadron(clave, x) === 0
        ? [{ texto: 'Eliminar', accion: 'eliminar', clase: 'peligro', alClic: () => eliminarPadron(clave, vivo(clave, x)) }]
        : [{ texto: 'Dar de baja', accion: 'baja', clase: 'peligro', alClic: () => activarPadron(clave, vivo(clave, x), false) }]);
}
export const NOMBRE_PADRON = { carriers: 'carrier', unidades: 'unidad', choferes: 'chofer' };
export function vigenciasDelCarrier(c) { return estado.vigencias.filter(v => v.Rol === 'carrier' && String(v.Title).endsWith(`· ${c.Title}`)); }
// C-05 (v0.23.0): el SEGUNDO paso de una escritura en dos pasos (carrier + su vigencia del tablero) corre aparte: si falla,
// el primero SI quedo y el aviso lo dice en ambar. Antes salia «No se pudo guardar» sobre un carrier ya guardado, el
// operador lo reintentaba y lo duplicaba. Devuelve el mensaje del error, o null si el paso paso.
export async function segundoPaso(fn) { try { await fn(); return null; } catch (e) { return e && e.message ? e.message : String(e); } }
async function eliminarPadron(clave, x) {
    let citas;
    try { citas = await referenciasPadronVivas(clave, x); }
    catch (err) { avisar('No pude confirmar quién lo cita (' + err.message + '). No se elimina: dale de baja.', 'error'); return; }
    if (citas > 0) { avisar('Ya hay pre-altas o embarques que lo citan (incluso fuera de los últimos ' + CONFIG.ventanaDias + ' días): se da de baja, no se elimina.', 'error'); return; }
    const { ok } = await confirmar({ titulo: `Eliminar ${NOMBRE_PADRON[clave]}`, peligro: true, ok: 'Eliminar',
        texto: `«${x.Title}» no lo cita ninguna pre-alta ni embarque. Se borra del padrón${clave === 'carriers' ? ' junto con su vigencia ASEA' : ''}; si hacía falta corregirlo, se transcribe de nuevo del oficio.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // C-05 (v0.23.0): las vigencias ANTES que el carrier. Si falla a medias queda un carrier sin vigencia (se ve en el
        // padron y se reintenta), no una vigencia huerfana pintandose en «Hoy» sin carrier.
        if (clave === 'carriers') for (const v of vigenciasDelCarrier(x)) { await estado.cliente.borrarRenglon(estado.siteId, L.vigencias, v.id); estado.vigencias = estado.vigencias.filter(y => y.id !== v.id); }
        await estado.cliente.borrarRenglon(estado.siteId, L[clave], x.id);
        estado[clave] = estado[clave].filter(y => y.id !== x.id);
        avisar(`${NOMBRE_PADRON[clave]} eliminado.`, 'bien'); pintarPadron();
    } catch (e) { avisar('No se pudo eliminar: ' + e.message, 'error'); }
}
async function activarPadron(clave, x, activo) {
    const { ok, motivo } = await confirmar(activo
        ? { titulo: `Reactivar ${NOMBRE_PADRON[clave]}`, ok: 'Reactivar', texto: `«${x.Title}» vuelve a ampararse en la puerta y a salir en las listas.` }
        : { titulo: `Dar de baja ${NOMBRE_PADRON[clave]}`, peligro: true, ok: 'Dar de baja', motivo: 'opcional', etiquetaMotivo: 'Por qué (queda en Notas)',
            texto: `«${x.Title}» ya lo citan pre-altas o embarques, así que no se borra: deja de ampararse en la puerta y de salir en las listas, y el historial lo sigue viendo.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        const campos = { Activo: activo };
        // C-06 (v0.22.0): el motivo queda en Notas para las TRES claves, como promete el dialogo (unidades y choferes tienen
        // Notas desde la tarea 9). Si la lista aun no trae la columna, se da de baja igual, sin motivo, y se dice.
        if (!activo && motivo) campos.Notas = `${x.Notas ? x.Notas + '\n' : ''}BAJA ${fechaMexico()} por ${estado.cuenta.username}: ${motivo}`;
        if (activo) campos.Notas = `${x.Notas ? x.Notas + '\n' : ''}REACTIVADA ${fechaMexico()} por ${estado.cuenta.username}`;   // U-133 (v0.73.0): la reactivación también deja rastro
        let sinNotas = false;
        try { await estado.cliente.actualizarRenglon(estado.siteId, L[clave], x.id, campos); }
        catch (e) {
            if (!campos.Notas || !esColumnaFaltante(e)) throw e;   // C-15
            delete campos.Notas; sinNotas = true;
            await estado.cliente.actualizarRenglon(estado.siteId, L[clave], x.id, campos);
        }
        aplicar(clave, x, campos);   // C-23
        const pendiente = clave === 'carriers' ? await segundoPaso(async () => { for (const v of vigenciasDelCarrier(x)) { await estado.cliente.actualizarRenglon(estado.siteId, L.vigencias, v.id, { Activo: activo }); v.Activo = activo; } }) : null;
        if (pendiente) avisar(`${NOMBRE_PADRON[clave]} ${activo ? 'reactivado' : 'dado de baja'}, pero su vigencia ASEA del tablero no cambió (${pendiente}). Edítalo y guarda para sincronizarla.`, 'ojo');
        else if (activo) avisar(`${NOMBRE_PADRON[clave]} reactivado.`, 'bien');
        else if (sinNotas) avisar(`${NOMBRE_PADRON[clave]} dado de baja, pero el motivo NO quedó` + (estado.rol === 'gerencia' ? ': la lista no tiene todavía la columna Notas (setup, tarea 9).' : '. Avisa a gerencia.'), 'ojo');   // U-31
        else avisar(`${NOMBRE_PADRON[clave]} dado de baja${motivo ? ' · el motivo quedó en Notas' : ''}.`, 'bien');
        pintarPadron();
    } catch (e) { avisar('No se pudo cambiar: ' + e.message, 'error'); }
}

// ---- alta y edición del padrón. Un solo formulario por tipo: vacío da de alta, con `estado.padronEdit` edita ese renglón.
// Editar es un PATCH sobre el mismo renglón (SharePoint guarda la versión anterior): un dedazo en la póliza no obliga a
// dar de baja y transcribir de nuevo. Lo que se vacía en el formulario se borra en la lista (null), no se conserva.
export const FORMA_PADRON = {
    carriers: { forma: 'pdFormaCarrier',titulo: 'carrier', campos: ['pcTitle', 'pcAut', 'pcVig', 'pcFolio', 'pcSCT', 'pcCSF'] },
    unidades: { forma: 'pdFormaUnidad',titulo: 'unidad', campos: ['puuPlaca', 'puuPlana', 'puuFolio', 'puuCap', 'puuSerie', 'puuMarca', 'puuTarjeta', 'puuTarjetaVig', 'puuPoliza', 'puuPolizaVig'] },
    choferes: { forma: 'pdFormaChofer',titulo: 'chofer', campos: ['pchNombre', 'pchLic', 'pchLicVig'] },
};
function leerFormaPadron(clave) {
    if (clave === 'carriers') return {
        Title: $('pcTitle').value.trim(), AutorizacionASEA: $('pcAut').value.trim(), VigenciaASEA: aIsoDia($('pcVig').value),
        FolioOficio: $('pcFolio').value.trim(), Corrientes: [...document.querySelectorAll('input[name="pcCorr"]:checked')].map(c => c.value).join('; '),
        RegistroSCT: $('pcSCT').value.trim(), CSFVigencia: aIsoDia($('pcCSF').value), Activo: true };
    if (clave === 'unidades') return {
        Title: placaNormal($('puuPlaca').value), PlacaPlana: placaNormal($('puuPlana').value), CarrierId: Number($('puuCarrier').value),
        TipoUnidad: $('puuTipo').value, FolioOficio: folioUnidad($('puuFolio').value, $('puuCarrier').value), CapacidadKg: $('puuCap').value ? Number($('puuCap').value) : null,
        NumeroSerie: $('puuSerie').value.trim(), Marca: $('puuMarca').value.trim(), TarjetaCirc: $('puuTarjeta').value.trim(), TarjetaVigencia: aIsoDia($('puuTarjetaVig').value),
        Poliza: $('puuPoliza').value.trim(), PolizaVigencia: aIsoDia($('puuPolizaVig').value), Activo: true };
    return {
        Title: $('pchNombre').value.trim(), CarrierId: Number($('pchCarrier').value), Licencia: $('pchLic').value.trim(),
        LicenciaVigencia: aIsoDia($('pchLicVig').value), Activo: true };
}
// v0.19.9 (Carlos, 2026-09-08): el folio por unidad era redundante — cada placa se transcribe del oficio del carrier.
// Vacio hereda el folio del carrier; se teclea solo cuando la unidad la ampara un alcance distinto.
function folioUnidad(tecleado, carrierId) {
    const t = String(tecleado || '').trim(); if (t) return t;
    const c = porId(estado.carriers, carrierId); return c && c.FolioOficio ? String(c.FolioOficio).trim() : '';
}
function validarFormaPadron(clave) {
    if (clave === 'carriers' && !$('pcTitle').value.trim()) return 'Falta la razón social.';
    if (clave === 'unidades' && (!$('puuCarrier').value || !$('puuPlaca').value.trim())) return 'Carrier y placa son obligatorios.';
    if (clave === 'choferes' && (!$('pchCarrier').value || !$('pchNombre').value.trim())) return 'Carrier y nombre son obligatorios.';
    return null;
}
function llenarFormaPadron(clave, x) {
    const f = textoDe;   // C-26
    const fecha = v => v ? fechaCorta(v) : '';
    if (clave === 'carriers') {
        $('pcTitle').value = f(x.Title); $('pcAut').value = f(x.AutorizacionASEA); $('pcVig').value = fecha(x.VigenciaASEA); $('pcFolio').value = f(x.FolioOficio);
        $('pcSCT').value = f(x.RegistroSCT); $('pcCSF').value = fecha(x.CSFVigencia);
        const corr = lista(x.Corrientes); for (const c of document.querySelectorAll('input[name="pcCorr"]')) c.checked = corr.includes(c.value);
    } else if (clave === 'unidades') {
        $('puuCarrier').value = f(x.CarrierId); $('puuPlaca').value = f(x.Title); $('puuPlana').value = f(x.PlacaPlana); $('puuTipo').value = f(x.TipoUnidad) || 'gondola';
        $('puuFolio').value = f(x.FolioOficio); $('puuCap').value = f(x.CapacidadKg); $('puuSerie').value = f(x.NumeroSerie); $('puuMarca').value = f(x.Marca);
        $('puuTarjeta').value = f(x.TarjetaCirc); $('puuTarjetaVig').value = fecha(x.TarjetaVigencia); $('puuPoliza').value = f(x.Poliza); $('puuPolizaVig').value = fecha(x.PolizaVigencia);
    } else {
        $('pchCarrier').value = f(x.CarrierId); $('pchNombre').value = f(x.Title); $('pchLic').value = f(x.Licencia); $('pchLicVig').value = fecha(x.LicenciaVigencia);
    }
}
function vaciarFormaPadron(clave) {
    for (const id of FORMA_PADRON[clave].campos) $(id).value = '';
    if (clave === 'carriers') for (const c of document.querySelectorAll('input[name="pcCorr"]')) c.checked = false;
}
function tituloFormaPadron(clave) {
    const e = estado.padronEdit && estado.padronEdit.clave === clave ? estado.padronEdit.x : null;
    $(FORMA_PADRON[clave].forma).querySelector('h2').textContent = e ? `Editar ${FORMA_PADRON[clave].titulo}: ${e.Title}` : `Alta de ${FORMA_PADRON[clave].titulo}`;
}
const SELECT_CARRIER_PADRON = { unidades: 'puuCarrier', choferes: 'pchCarrier' };
export function abrirFormaPadron(clave, x = null) {
    estado.padronEdit = x ? { clave, x } : null;
    if (x) llenarFormaPadron(clave, x); else vaciarFormaPadron(clave);
    // U-46 (v0.29.0): el alta de unidad/chofer arranca con el carrier que el padrón ya tiene elegido (el filtro «Mostrando solo
    // lo de X») o, si no hay filtro, el último que se guardó en esta sesión. Transcribir 8 placas de un oficio eran 8 búsquedas
    // en el select. El value se conserva además entre altas consecutivas (opciones(), U-40).
    if (!x && SELECT_CARRIER_PADRON[clave]) {
        const sel = $(SELECT_CARRIER_PADRON[clave]);
        const quiero = estado.padronCarrier !== null && estado.padronCarrier !== undefined ? estado.padronCarrier : estado.ultimoCarrierPadron;
        if (quiero !== null && quiero !== undefined && [...sel.options].some(o => o.value === String(quiero))) sel.value = String(quiero);
    }
    tituloFormaPadron(clave);
    if (estado.pestana === 'padron') { abrirAsistentePadron(clave, x); return; }   // v0.72.0
    // Desde la pre-alta: si el Padrón tenía abierta el alta del mismo tipo, se suelta (el cuerpo es uno solo y se muda al dialog).
    const pv = estado.padronVista;
    if (asistentePadronAbierto() && pv.asis.clave === clave) { Object.assign(pv, pv.asis.volver); pv.asis = null; }
    cuerpoAlDialog(clave);
    abrirForma(FORMA_PADRON[clave].forma);
}
export function cerrarFormaPadron(clave) {
    estado.padronEdit = null; tituloFormaPadron(clave);
    cerrarForma(FORMA_PADRON[clave].forma);
    const pv = estado.padronVista;
    if (asistentePadronAbierto() && pv.asis.clave === clave) { Object.assign(pv, pv.asis.volver); pv.asis = null; cuerpoAlDialog(clave); }   // v0.72.0
}
// ---- v0.72.0 (tanda 2 del padrón por carrier): alta y edición EN LA PÁGINA, por pasos, con «Así va» a un lado. El cuerpo de la
// forma (#pdCuerpo<Clave>) es el mismo nodo que vive en el <dialog>: desde el Padrón se muda a #pdAsisCuerpo y se ve un paso a la
// vez; desde la pre-alta (U-10, «+ Alta de carrier» encima del asistente) vuelve al dialog y se ve entero. guardarPadron no cambió.
const CUERPO_PADRON = { carriers: 'pdCuerpoCarriers', unidades: 'pdCuerpoUnidades', choferes: 'pdCuerpoChoferes' };
const REVISAR_PADRON = ['Revisar', 'Revisa antes de guardar', 'Cotéjalo contra el oficio. Toca un paso de arriba para corregirlo.'];
const PASOS_PADRON = {
    carriers: [['Oficio', '¿Qué dice el oficio ASEA?', 'Razón social, autorización, vigencia, folio y corrientes, como vienen en el oficio.'],
        ['Registro', '¿Y su registro?', 'El registro SCT y la vigencia de la CSF.'], REVISAR_PADRON],
    unidades: [['Vehículo', '¿Qué placas trae el oficio?', 'Carrier, placas, tipo y capacidad.'],
        ['Documentos', '¿Y sus documentos?', 'Serie, tarjeta de circulación y póliza, con sus vigencias.'], REVISAR_PADRON],
    choferes: [['Chofer y licencia', '¿Quién maneja?', 'El nombre como viene en la licencia, su número y su vigencia.'], REVISAR_PADRON],
};
const OBLIGATORIOS_PADRON = {
    carriers: { 1: [['pcTitle', 'Falta la razón social.']] },
    unidades: { 1: [['puuCarrier', 'Falta el carrier.'], ['puuPlaca', 'Falta la placa del tractor.']] },
    choferes: { 1: [['pchCarrier', 'Falta el carrier.'], ['pchNombre', 'Falta el nombre.']] },
};
const ASIVA_PADRON = {
    carriers: [['Oficio', [['Razón social', 'pcTitle'], ['Autorización', 'pcAut'], ['Vence', 'pcVig'], ['Folio', 'pcFolio'], ['Corrientes', '@corr']]], ['Registro', [['SCT', 'pcSCT'], ['CSF vence', 'pcCSF']]]],
    unidades: [['Vehículo', [['Carrier', 'puuCarrier'], ['Placa tractor', 'puuPlaca'], ['Placa plana', 'puuPlana'], ['Tipo', 'puuTipo'], ['Capacidad', 'puuCap']]],
        ['Documentos', [['Serie', 'puuSerie'], ['Tarjeta vence', 'puuTarjetaVig'], ['Póliza', 'puuPoliza'], ['Póliza vence', 'puuPolizaVig']]]],
    choferes: [['Chofer', [['Carrier', 'pchCarrier'], ['Nombre', 'pchNombre'], ['Licencia', 'pchLic'], ['Vence', 'pchLicVig']]]],
};
function valorCampoPadron(id) {
    if (id === '@corr') return [...document.querySelectorAll('input[name="pcCorr"]:checked')].map(c => etiquetaCorriente(c.value)).join(', ');
    const c = $(id);
    if (c.tagName === 'SELECT') return c.value && c.selectedOptions[0] ? c.selectedOptions[0].textContent : '';
    if (id === 'puuCap' && c.value.trim() && isFinite(Number(c.value))) return `${Number(c.value).toLocaleString('es-MX')} kg`;   // U-139 (v0.73.0): como la ficha
    return c.value.trim();
}
function cuerpoAlDialog(clave) {
    const c = $(CUERPO_PADRON[clave]), d = $(FORMA_PADRON[clave].forma);
    if (c.parentElement !== d) d.insertBefore(c, d.querySelector('.dlg-botones'));
    for (const p of c.querySelectorAll('.pd-paso')) p.hidden = false;
}
function abrirAsistentePadron(clave, x) {
    const p = estado.padronVista;
    const volver = { v: p.v, carrier: p.carrier, sub: p.sub, ficha: p.ficha, desde: p.desde };
    $('pdAsisCuerpo').textContent = ''; $('pdAsisCuerpo').appendChild($(CUERPO_PADRON[clave]));
    limpiarAvisos();
    p.asis = { clave, paso: 1, max: x ? PASOS_PADRON[clave].length : 1, volver };
    p.v = 'asis';
    $('pdAsis').dataset.huella = huellaForma($('pdAsisCuerpo'));
    pintarPadron(); window.scrollTo({ top: 0 }); enfocarPasoPadron();
}
function enfocarPasoPadron() {
    const a = estado.padronVista.asis; if (!a) return;
    const campo = $(CUERPO_PADRON[a.clave]).querySelector(`.pd-paso[data-paso="${a.paso}"] :is(input, select)`);
    (campo || $('pdAsisPregunta')).focus({ preventScroll: true });
}
function irPasoPadron(n) {
    const a = estado.padronVista.asis; if (!a) return;
    a.paso = n; a.max = Math.max(a.max, n);
    limpiarAvisos(); pintarPadron(); window.scrollTo({ top: 0 }); enfocarPasoPadron();
}
/** Lo que le falta a un paso: los obligatorios y las fechas del paso (vacías o bien escritas). [id, texto] o null. */
function faltaPasoPadron(clave, paso) {
    for (const [id, texto] of (OBLIGATORIOS_PADRON[clave][paso] || [])) if (!$(id).value.trim()) return [id, texto];
    for (const f of $(CUERPO_PADRON[clave]).querySelectorAll(`.pd-paso[data-paso="${paso}"] input.fecha`)) {
        if (!f.value.trim()) continue;
        try { aIsoDia(f.value); } catch { return [f.id, `Fecha mal escrita en «${document.querySelector(`label[for="${f.id}"]`).textContent}»: dd/mm/aaaa.`]; }
    }
    return null;
}
/** La placa tecleada ya existe en el padrón (otra unidad): el aviso de «Así va» lo dice antes de guardar. */
function placaRepetidaPadron() {
    const a = estado.padronVista.asis; if (!a || a.clave !== 'unidades' || !$('puuPlaca').value.trim()) return null;
    const placa = placaNormal($('puuPlaca').value), edit = estado.padronEdit && estado.padronEdit.x;
    return estado.unidades.find(u => placaNormal(u.Title) === placa && (!edit || u.id !== edit.id)) || null;
}
function pintarAsiVaPadron() {
    const a = estado.padronVista.asis; if (!a) return;
    const r = $('pdAsiVa'); r.textContent = '';
    r.appendChild(el('h2', '', `Así va ${ARTICULO_PADRON[a.clave]}`));
    const { dl, sec, fila } = dlPd();
    for (const [s, campos] of ASIVA_PADRON[a.clave]) { sec(s); for (const [t, id] of campos) fila(t, valorCampoPadron(id)); }
    r.appendChild(dl);
    const rep = placaRepetidaPadron();
    if (rep) {
        const m = el('div', 'mensaje ' + (activo(rep) ? 'error' : 'ojo'));
        m.appendChild(el('span', '', activo(rep) ? `La placa ${rep.Title} ya está en el padrón con ${nombreDe(estado.carriers, rep.CarrierId)}.` : `La placa ${rep.Title} existe dada de baja. En vez de duplicarla, reactívala desde su ficha.`));
        if (!activo(rep)) { const b = el('button', 'secundario', 'Abrir su ficha'); b.type = 'button'; b.id = 'btnPdAbrirRepetida'; b.addEventListener('click', () => salirAsistentePadron({ v: 'ficha', ficha: { clave: 'unidades', id: rep.id }, carrier: Number(rep.CarrierId), desde: 'carrier' })); m.appendChild(b); }
        r.appendChild(m);
    }
    const ob = Object.values(OBLIGATORIOS_PADRON[a.clave]).flat().map(([id]) => document.querySelector(`label[for="${id}"]`).textContent.toLowerCase());
    r.appendChild(el('p', 'pista', `Obligatorio: ${ob.join(', ')}. Lo demás se puede completar después con Editar.`));
}
function pintarAsistentePd(c) {
    const a = estado.padronVista.asis, clave = a.clave, pasos = PASOS_PADRON[clave], edit = estado.padronEdit && estado.padronEdit.clave === clave ? estado.padronEdit.x : null;
    const nombre = NOMBRE_PADRON[clave], final = a.paso === pasos.length;
    const migas = [['Padrón', () => salirAsistentePadron({ v: 'lista' })]];
    if (c) migas.push([c.Title, () => salirAsistentePadron({ v: 'carrier', carrier: c.id })]);
    migas.push([edit ? `Editar ${nombre}` : `Alta de ${nombre}`]);
    cabeceraPd(migas, edit ? `Editar ${nombre}: ${edit.Title}` : `Alta de ${nombre}`, null,
        clave === 'carriers' ? 'Se transcribe del oficio ASEA del carrier.' : `${c ? 'Para ' + c.Title + '. ' : ''}Se transcribe del oficio, nunca de memoria.`);
    kpisPd([]);
    const nav = $('pdPasos'); nav.textContent = ''; nav.style.setProperty('--n', pasos.length);
    pasos.forEach(([t], i) => {
        const n = i + 1, b = el('button', n === a.paso ? 'ahora' : n <= a.max ? 'hecho' : ''); b.type = 'button'; b.dataset.p = String(n);
        b.disabled = n > a.max; b.setAttribute('aria-label', `Paso ${n} · ${t}`); if (n === a.paso) b.setAttribute('aria-current', 'step');
        b.appendChild(el('i')); b.appendChild(el('span', '', t)); b.addEventListener('click', () => irPasoPadron(n)); nav.appendChild(b);
    });
    const [, pregunta, sub] = pasos[a.paso - 1];
    $('pdAsisPregunta').textContent = pregunta; $('pdAsisSub').textContent = `Paso ${a.paso} de ${pasos.length} · ${sub}`;
    for (const p of $(CUERPO_PADRON[clave]).querySelectorAll('.pd-paso')) p.hidden = Number(p.dataset.paso) !== a.paso;
    $('pdAsisCuerpo').hidden = final;
    const rv = $('pdAsisRevisar'); rv.hidden = !final; rv.textContent = '';
    if (final) {
        const dl = el('dl', 'pd-pares');
        for (const [s, campos] of ASIVA_PADRON[clave]) for (const [t, id] of campos) {
            const v = valorCampoPadron(id), par = el('div'); par.appendChild(el('dt', '', t)); par.appendChild(el('dd', v ? 'mono' : 'f', v || '—')); dl.appendChild(par);
        }
        rv.appendChild(dl);
    }
    $('btnPdAtras').classList.toggle('oculto', a.paso === 1);
    $('btnPdSiguiente').textContent = final ? (edit ? 'Guardar cambios' : `Dar de alta ${ARTICULO_PADRON[clave]}`) : `Siguiente: ${pasos[a.paso][0].toLowerCase()}`;
    pintarAsiVaPadron();
}
/** Cancelar, las migas o «Abrir su ficha»: con algo cambiado desde que abrió, pregunta (el mismo «Descartar lo capturado»). */
async function salirAsistentePadron(destino = null) {
    const p = estado.padronVista; if (!asistentePadronAbierto()) return true;
    const cuerpo = $('pdAsisCuerpo');
    if (hayCaptura(cuerpo) && huellaForma(cuerpo) !== $('pdAsis').dataset.huella) {
        const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: `${estado.padronEdit ? 'Los cambios' : 'Lo que llevas de esta alta'} no se ha guardado. Si sales, se pierde.` });
        if (!ok) return false;
    }
    cerrarFormaPadron(p.asis.clave);
    if (destino) Object.assign(p, destino);
    pintarPadron(); window.scrollTo({ top: 0 });
    return true;
}
async function siguientePadron() {
    const a = estado.padronVista.asis; if (!a) return;
    const n = PASOS_PADRON[a.clave].length;
    if (a.paso < n) {
        const f = faltaPasoPadron(a.clave, a.paso);
        if (f) { avisar(f[1], 'error'); $(f[0]).focus(); return; }
        irPasoPadron(a.paso + 1); return;
    }
    for (let i = 1; i < n; i++) { const f = faltaPasoPadron(a.clave, i); if (f) { irPasoPadron(i); avisar(f[1], 'error'); $(f[0]).focus(); return; } }
    const b = $('btnPdSiguiente'); b.disabled = true;   // C-24: un doble toque no da dos altas
    try { await guardarPadron(a.clave); } finally { b.disabled = false; }
}
const AVISO_ALTA = { carriers: 'Carrier dado de alta. Ahora sus unidades, transcritas del oficio.', unidades: 'Unidad transcrita.', choferes: 'Chofer dado de alta.' };
const BOTON_GUARDAR_PADRON = { carriers: 'btnGuardarCarrier', unidades: 'btnGuardarUnidad', choferes: 'btnGuardarChofer' };
export async function guardarPadron(clave) {
    if (!PUEDE.capturarPrealta(estado.rol)) { avisar('Tu rol no edita el padrón.', 'error'); return; }
    const falta = validarFormaPadron(clave); if (falta) { avisar(falta, 'error'); return; }
    const edit = estado.padronEdit && estado.padronEdit.clave === clave ? estado.padronEdit.x : null;
    let campos;
    try { campos = leerFormaPadron(clave); } catch (e) { avisar(e.message, 'error'); return; }
    await escribiendo(BOTON_GUARDAR_PADRON[clave], async () => { try {   // C-24: un doble toque creaba dos carriers / unidades / choferes
        await refrescarCliente();
        if (edit) {
            delete campos.Activo;   // editar no reactiva ni da de baja
            // Cambiar una vigencia pide motivo (queda en Notas): editar tambien sirve para «arreglar» una poliza sin
            // tener el papel, y la app no distingue eso de una correccion legitima — el rastro lo deja quien captura.
            const dia = v => (v ? String(v).slice(0, 10) : '');
            const cambios = VIGENCIAS_PADRON[clave].filter(([, col]) => dia(campos[col]) !== dia(edit[col])).map(([n, col]) => `${n}: ${fechaCorta(edit[col])} → ${fechaCorta(campos[col])}`);
            if (cambios.length) {
                const { ok, motivo } = await confirmar({ titulo: 'Cambio de vigencia', ok: 'Guardar', motivo: true, etiquetaMotivo: 'Por qué cambia (queda en Notas)',
                    texto: `${cambios.join(' · ')}. La versión anterior queda en SharePoint; el motivo, en las Notas del ${NOMBRE_PADRON[clave]}.` });
                if (!ok) return;
                campos.Notas = `${edit.Notas ? edit.Notas + '\n' : ''}VIGENCIA ${fechaMexico()} por ${estado.cuenta.username}: ${cambios.join(' · ')} — ${motivo}`;
            }
            const vsAntes = clave === 'carriers' ? vigenciasDelCarrier(edit) : [];
            try { await estado.cliente.actualizarRenglon(estado.siteId, L[clave], edit.id, paraPatch(campos)); }
            catch (e) { if (campos.Notas && esColumnaFaltante(e)) throw new Error('la vigencia no se puede cambiar todavía' + (estado.rol === 'gerencia' ? ': la lista no tiene la columna Notas (tarea 9 de setup-carlos.md)' : '; avisa a gerencia')); throw e; }   // U-31
            aplicar(clave, edit, campos);   // C-23
            // La vigencia ASEA del tablero se llama por el carrier y guarda su fecha: se corrige junto con él. Es el segundo
            // paso (C-05): si falla, el carrier ya quedo y el aviso lo dice; guardar de nuevo la sincroniza (Activo incluido).
            const pendiente = clave === 'carriers' ? await segundoPaso(async () => {
                if (vsAntes.length) for (const v of vsAntes) { const c = { Title: `Autorización ASEA transporte · ${edit.Title}`, Vence: edit.VigenciaASEA || null, Folio: edit.FolioOficio || null, Activo: edit.Activo !== false }; await estado.cliente.actualizarRenglon(estado.siteId, L.vigencias, v.id, c); aplicar('vigencias', v, c); }
                else if (edit.VigenciaASEA) await altaVigencia(`Autorización ASEA transporte · ${edit.Title}`, 'tercero', 'carrier', 'legal', edit.VigenciaASEA, edit.FolioOficio);
            }) : null;
            cerrarFormaPadron(clave);
            if (pendiente) avisar(`${NOMBRE_PADRON[clave]} actualizado, pero su vigencia ASEA del tablero no se sincronizó (${pendiente}). Guárdalo de nuevo para reintentar.`, 'ojo');
            else avisarAlta(`${NOMBRE_PADRON[clave]} actualizado.`, clave, edit);
            pintarPadron();
        } else {
            const nuevo = await estado.cliente.crearRenglon(estado.siteId, L[clave], limpiar(campos));
            anclar(clave, nuevo);   // C-23
            estado.ultimoCarrierPadron = clave === 'carriers' ? nuevo.id : (nuevo.CarrierId ? Number(nuevo.CarrierId) : estado.ultimoCarrierPadron);   // U-46
            const pendiente = clave === 'carriers' && nuevo.VigenciaASEA ? await segundoPaso(() => altaVigencia(`Autorización ASEA transporte · ${nuevo.Title}`, 'tercero', 'carrier', 'legal', nuevo.VigenciaASEA, nuevo.FolioOficio)) : null;
            vaciarFormaPadron(clave); cerrarFormaPadron(clave);
            if (pendiente) avisar(`${NOMBRE_PADRON[clave]} dado de alta, pero su vigencia ASEA NO quedó en el tablero (${pendiente}). Edítalo y guarda para crearla.`, 'ojo');
            else avisarAlta(AVISO_ALTA[clave], clave, nuevo);
            // U-10 (v0.23.0): si el alta vino desde la pre-alta (que sigue abierta atras), el carrier nuevo queda elegido en ella.
            if (clave === 'carriers' && asistentePrealtaAbierto()) elegirCarrierEnPrealta(nuevo);
            if (estado.pestana === 'padron') Object.assign(estado.padronVista, clave === 'carriers' ? { v: 'carrier', carrier: nuevo.id, sub: 'unidades' }
                : { v: 'ficha', ficha: { clave, id: nuevo.id }, carrier: Number(nuevo.CarrierId), desde: 'carrier' });   // v0.72.0 / v0.72.0
            pintarPadron();
        }
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}
async function altaVigencia(titulo, titular, rol, fuente, vence, folio) {
    const v = await estado.cliente.crearRenglon(estado.siteId, L.vigencias, limpiar({ Title: titulo, Titular: titular, Rol: rol, Fuente: fuente, Vence: vence, Folio: folio, AvisoDias: CONFIG.avisoVigenciaDias, Activo: true }));
    estado.vigencias.push(v);
}


$('pdBusca').addEventListener('input', () => { if (estado.padronVista.tab === 'buscar') pintarBusquedaPd(); });
for (const b of $('pdTabs').querySelectorAll('button')) b.addEventListener('click', () => { estado.padronVista.tab = b.dataset.pd; pintarPadron(); if (b.dataset.pd === 'buscar') $('pdBusca').focus(); });
for (const b of $('pdSubTabs').querySelectorAll('button')) b.addEventListener('click', () => { estado.padronVista.sub = b.dataset.sub; pintarPadron(); });
$('btnPdVolver').addEventListener('click', volverPadron);
$('btnPdCancelar').addEventListener('click', () => salirAsistentePadron());
$('btnPdSiguiente').addEventListener('click', siguientePadron);
$('btnPdAtras').addEventListener('click', () => { const a = estado.padronVista.asis; if (a && a.paso > 1) irPasoPadron(a.paso - 1); });
for (const ev of ['input', 'change']) $('pdAsisCuerpo').addEventListener(ev, pintarAsiVaPadron);
