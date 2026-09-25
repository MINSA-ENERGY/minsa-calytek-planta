// CALYTEK Planta — Gondolas: lista, bascula y ticket.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { comprimir } from './imagen.js';
import { accionCorreccion, avisoNeto, compuerta, etiquetaCorriente, fechaCorta, fechaMexico, horaMexico, limpiar, lista, plural, PUEDE, siguienteFolio, siguientePaso, slug, yaCapturado } from './reglas.js';
import { $, anclar, aplicar, avisar, botonAccion, confirmar, el, embarquesDelAno, enListaPlanta, enPlanta, esColumnaFaltante, escribiendo, esperaAutorizacion, estado, etiqueta, excepcionAutorizada, filtroTexto, firmar, fundirEnVentana, horaCorta, L, nombreDe, normaliza, palabraCompuertaDe, pintarInsignias, porId, quien, refrescarCliente, selloSinFirma, VERSION, vivo } from './nucleo.js';
import { entrar, recargar, repintar } from './navegacion.js';
import { asegurarFolioUnico } from './puerta.js';
import { abrirCertificadoDeEmbarque, camposCancelacion, certificadosDe, certificadoVigente } from './certificado.js';
import { ESCRITORIO } from './padron.js';
import { cortesDia, pintarHoy } from './hoy.js';
import { hallazgosDe, invalidarRama, reglasDe } from './archivos.js';

// ================================================================ GONDOLAS (rediseño tanda 3, v0.48.0)

/**
 * Tanda 3 (decisiones 5, 6 y 10): una sección con cuatro pestañas —En planta · Cerradas hoy · Rechazos · Historial— que
 * funde la lista de báscula, la Fila del día (tanda 2) y Cerrados. Las cuatro se pintan en cada repintado y solo se ve la
 * elegida (estado.vistaGondolas), así los conteos de las pestañas nunca se quedan viejos.
 */
const VISTAS_GONDOLAS = { planta: 'gpPlanta', hoy: 'gpHoy', rechazos: 'gpRechazos', historial: 'gpHistorial' };
export function elegirVistaGondolas(v) {
    estado.vistaGondolas = VISTAS_GONDOLAS[v] ? v : 'planta';
    for (const b of document.querySelectorAll('#gTabs button, #gKpis button')) b.setAttribute('aria-pressed', String(b.dataset.g === estado.vistaGondolas));
    for (const [k, id] of Object.entries(VISTAS_GONDOLAS)) $(id).hidden = k !== estado.vistaGondolas;
}

/**
 * Una tabla de Góndolas. Un solo DOM para los dos anchos: tabla en escritorio y tarjeta en celular, donde cada celda cae en
 * su lugar por su letra `m` (a · b arriba, c · d en medio, e a lo ancho, x solo en escritorio; estilo.css .gtabla). La
 * última celda, sin encabezado, lleva los botones. `columnas`: [{ t, m, v: e => texto | Node, num, clase }].
 */
function tablaGondolas(caja, columnas, filas, botonesDe, vacio, claseDe = () => '') {
    caja.textContent = '';
    if (!filas.length) { caja.appendChild(el('p', 'vacio', vacio)); return; }
    const t = el('table', 'gtabla'), thead = el('thead'), cab = el('tr'), tb = el('tbody');
    for (const c of columnas) cab.appendChild(el('th', c.num ? 'num' : '', c.t));
    cab.appendChild(el('th'));
    thead.appendChild(cab); t.appendChild(thead);
    for (const e of filas) {
        const tr = el('tr', `gfila ${claseDe(e)}`.trim());
        for (const c of columnas) {
            const td = el('td', `m-${c.m}${c.num ? ' num' : ''}${c.clase ? ' ' + c.clase : ''}`);
            const v = c.v(e);
            if (v instanceof Node) td.appendChild(v); else td.textContent = v === null || v === undefined || v === '' ? '—' : v;
            tr.appendChild(td);
        }
        const ac = el('td', 'm-f acciones-g');
        for (const b of botonesDe(e)) ac.appendChild(botonAccion(b));
        tr.appendChild(ac);
        tb.appendChild(tr);
    }
    t.appendChild(tb); caja.appendChild(t);
}
const folioG = e => el('span', e.Title ? 'folio' : 'folio sin', e.Title || 'sin folio');
/** El programa y, debajo, el manifiesto: es lo que se busca con el papel en la mano. */
function programaG(e) {
    const s = el('span', 'prog', nombreDe(estado.prealtas, e.PreAltaId));
    if (e.Manifiesto) s.appendChild(el('small', 'mono', e.Manifiesto));
    return s;
}
export const kgG = n => `${Number(n).toLocaleString('es-MX')} kg`;
/** Lo último que se capturó de una góndola en planta: el peso bruto, o la hora a la que pasó (o llegó, si espera). */
function ultimoDatoG(e) {
    if (e.Etapa === 'bruto') return `bruto ${kgG(e.BrutoKg)}`;
    return `${esperaAutorizacion(e) ? 'arribo' : 'pasó'} ${horaMexico(e.Arribo, 'hora')}`;
}
/** Ticket (si hay folio), el certificado de la góndola cerrada y Anular/Eliminar. `lista` es la que recorre el ticket con ‹ ›. */
function botonesCerrada(e, lista) {
    const bs = [];
    if (e.Title) bs.push({ texto: 'Ticket', accion: 'ticket', alClic: () => abrirTicketPop(lista, lista.indexOf(e)) });
    // v0.39.0: el certificado es de ESTA góndola y se emite desde su renglón cerrado (decisión 9: solo gerencia emite).
    // U-70 (v0.40.0): una anulada con certificado lo sigue mostrando, para que gerencia llegue a cancelarlo.
    const certs = certificadosDe(e), vig = certs.find(x => x.Estado === 'vigente');
    if (e.Etapa === 'cerrado' && !vig && PUEDE.emitirCertificado(estado.rol)) bs.push({ texto: 'Emitir certificado', accion: 'certificado', clase: 'si', alClic: () => abrirCertificadoDeEmbarque(vivo('embarques', e)) });
    else if (certs.length) bs.push({ texto: 'Certificado', accion: 'certificado', alClic: () => abrirCertificadoDeEmbarque(vivo('embarques', e)) });
    return bs.concat(botonCorreccion(e));
}
function estadoCertificadoG(e) {
    const vig = certificadoVigente(e);
    return vig ? el('span', 'paso e-ok', `${vig.Title} vigente`) : el('span', 'paso e-warn', 'sin emitir');
}

/** La sección entera: las listas y, cerradas, las tarjetas de pesaje y de ticket (como al entrar). */
export function pintarGondolas() {
    cerrarAsistente();
    pintarListasGondolas();
}

// ---------------------------------------------------------------- tanda 5: el asistente de la báscula (M5–M8)
/**
 * Muestra el asistente sobre la góndola `e` en una de sus tres pantallas (baPesar · baPausa · baTicketCaja) y esconde la
 * lista. `paso` pinta la barra de cinco: 3 bruto, 4 tara, 5 ticket; 3.5 es la pausa de la descarga (tres hechos, ninguno
 * en curso: la góndola no está en la báscula).
 */
function mostrarAsistente(e, pantalla, paso) {
    estado.asisEmbarque = e;
    for (const id of ['baPesar', 'baPausa', 'baTicketCaja']) $(id).classList.toggle('oculto', id !== pantalla);
    $('baAsis').classList.remove('oculto');
    $('p-bascula').classList.add('asistiendo');
    [...$('baPasos').children].forEach((d, i) => {
        d.className = i + 1 < paso ? 'hecho' : i + 1 === paso ? 'ahora' : '';
        if (i + 1 === paso) d.dataset.tercio = '3'; else delete d.dataset.tercio;
    });
    $('baPasoK').textContent = pantalla === 'baPausa' ? 'Paso 3 de 5 · A descargar' : `Paso ${paso} de 5 · ${['', '', '', 'Bruto', 'Tara', 'Ticket'][paso]}`;
    $('baQuien').textContent = e.Title ? `${e.Title} · ${e.PlacaTractor || ''}` : [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' · ');
    $('baMiga').textContent = e.Title || e.PlacaTractor || 'Góndola';
    // U-88 (v0.51.0): cerrada, el dato que se le dice al chofer es el neto, no el bruto de la etapa anterior.
    const peso = e.NetoKg ? `neto ${kgG(e.NetoKg)}` : e.BrutoKg ? `bruto ${kgG(e.BrutoKg)}` : null;
    $('baQuienSub').textContent = [nombreDe(estado.prealtas, e.PreAltaId), nombreDe(estado.carriers, e.CarrierId), peso].filter(x => x && x !== '—').join(' · ');
    pintarCapturado(e);
    cerrarHojaCapturado();
    window.scrollTo({ top: 0 });
}
export function cerrarAsistente() {
    for (const id of ['baAsis', 'baPesar', 'baPausa', 'baTicketCaja']) $(id).classList.add('oculto');
    $('p-bascula').classList.remove('asistiendo');
    cerrarHojaCapturado();
    estado.asisEmbarque = null;
}
/** M5: tras el bruto, la pausa de la descarga — el folio ya nació y la góndola se retoma desde la lista. */
function mostrarPausa(e) {
    const dl = $('baPausaDatos'); dl.textContent = '';
    for (const [k, v] of [['Folio', e.Title], ['Bruto', `${kgG(e.BrutoKg)} · ${horaMexico(e.BrutoHora, 'hora')}`], ['Sigue', 'la tara, con la góndola vacía']]) {
        dl.appendChild(el('dt', '', k)); dl.appendChild(el('dd', k === 'Sigue' ? '' : 'mono', v));
    }
    mostrarAsistente(e, 'baPausa', 3.5);
}
/** M6 / M7: todo lo que la góndola ya trae, por paso. En escritorio es la columna derecha; en celular, la hoja. */
function pintarCapturado(e) {
    const dl = $('baCapturadoDl'); dl.textContent = '';
    const sec = (t, h) => { const d = el('div', 'sec', t); if (h) d.appendChild(el('em', '', h)); dl.appendChild(d); };
    // nombreDe() da «—» cuando no hay id: ese renglón no se pinta.
    const par = (k, v, clase) => { if (!v || v === '—') return; dl.appendChild(el('dt', '', k)); dl.appendChild(el('dd', clase || '', v)); };
    const pre = porId(estado.prealtas, e.PreAltaId);
    const corr = e.CorrienteDeclarada || (pre && pre.Corriente);
    sec('Programa y documentos', horaMexico(e.Arribo, 'hora'));
    par('Programa', nombreDe(estado.prealtas, e.PreAltaId)); par('Carrier', nombreDe(estado.carriers, e.CarrierId));
    par('Placas', [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' · '), 'mono');
    par('Chofer', e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId));
    par('Manifiesto', e.Manifiesto, 'mono'); par('Corriente', etiquetaCorriente(corr));
    par('Capturó', quien(e.CapturadoPor));
    const avisos = reglasDe(e, 'aviso');   // C-54
    sec('Veredicto');
    // C-55: un sello sin renglón en PLANTA_Firmas ya no se lee como autorización (lo mismo que dicen la lista y Hoy).
    const pc = palabraCompuertaDe(e), autorizo = excepcionAutorizada(e);
    par('Resultado', e.Compuerta !== 'excepcion-comercial' ? pc.palabra : autorizo ? `Espera · autorizó ${quien(e.ExcepcionAutorizo)}` : `Espera · ${selloSinFirma(e)}falta autorización`, 'e-' + pc.tono);
    par('Avisos', avisos, 'e-warn');
    if (e.BrutoKg) { sec('Peso bruto', horaMexico(e.BrutoHora, 'hora')); par('Folio', e.Title, 'mono'); par('Bruto', kgG(e.BrutoKg), 'mono'); }
    if (e.TaraKg) { sec('Tara', horaMexico(e.TaraHora, 'hora')); par('Tara', kgG(e.TaraKg), 'mono'); par('Neto', kgG(e.NetoKg), 'mono'); par('Ticket de báscula', e.TicketBascula, 'mono'); }
}
// U-81 (v0.51.0): el velo es un elemento (#baVelo) que recibe el toque y cierra la hoja; antes era una box-shadow y el toque
// caía en el teclado o en Guardar de atrás.
export function abrirHojaCapturado() { $('baCapturado').classList.add('abierta'); $('baVelo').hidden = false; $('btnLoCapturado').setAttribute('aria-expanded', 'true'); $('btnCerrarCapturado').focus(); }
export function cerrarHojaCapturado() { $('baCapturado').classList.remove('abierta'); $('baVelo').hidden = true; $('btnLoCapturado').setAttribute('aria-expanded', 'false'); }
/** Solo las cuatro listas y sus conteos: tras emitir un certificado no se toca el pesaje abierto. */
export function pintarListasGondolas() {
    const { hoy, dia, diaCierre } = cortesDia();
    const planta = enListaPlanta();
    const porCierre = (a, b) => String(b.TaraHora || '').localeCompare(String(a.TaraHora || ''));
    const cerradasHoy = estado.embarques.filter(e => e.Etapa === 'cerrado' && diaCierre(e) === hoy).sort(porCierre);
    const rechazosHoy = estado.embarques.filter(e => e.Etapa === 'rechazado' && dia(e) === hoy).sort((a, b) => b.id - a.id);
    const cuenta = (id, n, alerta) => { $(id).textContent = n ? String(n) : ''; $(id).classList.toggle('alerta', !!alerta && n > 0); };
    const esperan = planta.filter(esperaAutorizacion).length;
    cuenta('gnPlanta', planta.length, esperan > 0); cuenta('gnHoy', cerradasHoy.length); cuenta('gnRechazos', rechazosHoy.length, true);
    // v0.66.0: la banda cuenta lo mismo que las pestañas; en cero dice 0, no se vacía.
    for (const [id, n, alerta] of [['gkPlanta', planta.length, esperan > 0], ['gkHoy', cerradasHoy.length, false], ['gkRechazos', rechazosHoy.length, true]]) { $(id).textContent = String(n); $(id).parentElement.classList.toggle('alerta', alerta && n > 0); }
    elegirVistaGondolas(estado.vistaGondolas);
    $('btnNuevaGondola').classList.toggle('oculto', !PUEDE.puerta(estado.rol));   // tanda 2: la llegada se abre desde aquí; quien no captura no la ve

    // En planta: el siguiente paso de cada góndola en texto (reglas.siguientePaso). La que espera a gerencia sale en ámbar,
    // sin «Pesar» (decisión 10): se autoriza en Hoy › Pendiente revisar.
    const captura = PUEDE.puerta(estado.rol);
    tablaGondolas($('baLista'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Carrier', m: 'x', v: e => nombreDe(estado.carriers, e.CarrierId) },
        { t: 'Siguiente paso', m: 'e', v: e => { const s = siguientePaso(e, excepcionAutorizada(e)); return el('span', 'paso e-' + s.tono, s.texto); } },
        { t: 'Último dato', m: 'd', num: true, v: ultimoDatoG, clase: 'mono' }
    ], planta, e => {
        const bs = [];
        if (captura && enPlanta(e)) { const bruto = e.Etapa === 'compuerta'; bs.push({ texto: bruto ? 'Pesar bruto' : 'Pesar tara', accion: 'pesar', clase: 'si', alClic: () => abrirPesaje(vivo('embarques', e), bruto ? 'bruto' : 'tara') }); }
        return bs.concat(botonCorreccion(e));
    }, 'Nada en planta.', e => (esperaAutorizacion(e) ? 'espera' : ''));
    $('gPiePlanta').textContent = planta.length ? `${plural(planta.length, 'góndola en planta', 'góndolas en planta')}${esperan ? ` · ${esperan === 1 ? 'una espera' : `${esperan} esperan`} la autorización de gerencia` : ''}.` : '';

    // Cerradas hoy: el día de cierre es el de la tara (F1), y aquí vive el certificado (decisión 9).
    tablaGondolas($('gHoy'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Tara', m: 'x', v: e => horaMexico(e.TaraHora, 'hora'), clase: 'mono mudo' },
        { t: 'Neto', m: 'd', num: true, v: e => kgG(e.NetoKg), clase: 'mono' },
        { t: 'Ticket báscula', m: 'x', v: e => e.TicketBascula, clase: 'mono' },
        { t: 'Certificado', m: 'e', v: estadoCertificadoG }
    ], cerradasHoy, e => botonesCerrada(e, cerradasHoy), 'Ninguna góndola ha cerrado hoy.');
    const kgHoy = cerradasHoy.reduce((a, e) => a + (Number(e.NetoKg) || 0), 0);
    $('gPieHoy').textContent = cerradasHoy.length ? `${plural(cerradasHoy.length, 'cerrada', 'cerradas')} hoy · ${kgG(kgHoy)}${PUEDE.emitirCertificado(estado.rol) ? '' : ' · el certificado lo emite gerencia'}.` : '';

    // Rechazos de hoy: la causa sale de CompuertaDetalle (las reglas legales que fallaron, C-54) y, debajo, qué le pasó a
    // cada una (U-89): «Placa» sola no dice si no estaba en el padrón o la amparaba otro carrier.
    const causa = e => {
        const hs = hallazgosDe(e, 'legal'), s = el('span', 'causa');
        s.appendChild(el('span', 'paso e-bad', hs.map(h => h.regla).join(', ') || 'no entró'));
        const det = hs.map(h => h.detalle).filter(Boolean).join(' · ');
        if (det) s.appendChild(el('small', '', det));
        return s;
    };
    tablaGondolas($('gRechazos'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Carrier', m: 'c', v: e => nombreDe(estado.carriers, e.CarrierId) },
        { t: 'Hora', m: 'd', num: true, v: e => horaMexico(e.Arribo, 'hora'), clase: 'mono' },
        { t: 'Causa', m: 'e', v: causa }
    ], rechazosHoy, e => botonesCerrada(e, rechazosHoy.filter(x => x.Title)), 'Ningún rechazo hoy.');

    pintarHistorial();
    // F4: el CSV de todo lo cargado, para el reporte al cliente y la bitácora, sin copiar cifras de la pantalla.
    $('btnExportar').classList.toggle('oculto', !estado.embarques.length);
    // U-14: que la pantalla diga hasta dónde alcanza lo que muestra; «0 hace cuatro meses» es solo que no se cargó.
    $('tbAlcance').textContent = `Se cargan los últimos ${CONFIG.ventanaDias} días (desde el ${fechaCorta(estado.ventanaDesde)}) más todo lo que sigue abierto. El historial completo vive en SharePoint.`;
}

/**
 * Historial (antes Cerrados, 2026-09-08): todo folio que ya no está en planta —cerrado, anulado o rechazado— dentro de la
 * ventana cargada, del más reciente al más viejo, con su Ticket para reimprimir (la línea «Emitido» del ticket lleva la hora
 * de la reimpresión) y su certificado. El buscador filtra por folio, placa, manifiesto, carrier, ticket y folio CT (U-76).
 */
function pintarHistorial() {
    const q = normaliza($('baBusca').value.trim());
    const pega = filtroTexto(q);   // C-26
    const momento = e => e.TaraHora || e.AnuladoEl || e.Arribo || '';
    const todos = estado.embarques.filter(e => e.Title && !enPlanta(e)).sort((a, b) => momento(b).localeCompare(momento(a)));
    const visibles = todos.filter(e => pega(e.Title, e.PlacaTractor, e.PlacaPlana, e.Manifiesto, nombreDe(estado.carriers, e.CarrierId), e.TicketBascula, ...certificadosDe(e).map(c => c.Title)));
    const estadoH = e => {
        if (e.Etapa !== 'cerrado') return el('span', 'paso e-' + siguientePaso(e).tono, siguientePaso(e).texto);
        const vig = certificadoVigente(e);
        return vig ? el('span', 'paso e-ok', `cerrada · ${vig.Title}`) : el('span', 'paso e-ok', 'cerrada');
    };
    tablaGondolas($('baCerrados'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Estado', m: 'e', v: estadoH },
        { t: 'Neto', m: 'd', num: true, v: e => (e.NetoKg ? kgG(e.NetoKg) : null), clase: 'mono' },
        { t: 'Fecha', m: 'x', v: e => horaMexico(momento(e)), clase: 'mono mudo' }
    ], visibles, e => botonesCerrada(e, visibles), q ? `Nada coincide en los últimos ${CONFIG.ventanaDias} días.` : 'Ningún folio cerrado en la ventana cargada.', e => e.Etapa);
    $('baBuscaCuenta').textContent = q ? plural(visibles.length, 'resultado') : '';
    // U-76 (v0.42.0): cuántas cerradas siguen sin certificado, sobre todo lo cargado (no sobre lo filtrado).
    const cerradas = todos.filter(e => e.Etapa === 'cerrado'), sinCert = cerradas.filter(e => !certificadoVigente(e)).length, otras = todos.length - cerradas.length;
    $('gPieHistorial').textContent = todos.length ? `${cerradas.length} con neto${sinCert ? ` · ${sinCert} sin certificado` : ''}${otras ? ` · ${plural(otras, 'anulado o rechazado', 'anulados o rechazados')}` : ''}.` : '';
}

/** El boton de deshacer que le toca a un embarque segun su etapa (reglas.accionCorreccion), o ninguno. */
export function botonCorreccion(e) {
    if (!PUEDE.corregir(estado.rol)) return [];
    const a = accionCorreccion(e);
    if (a === 'eliminar') return [{ texto: 'Eliminar', accion: 'eliminar', clase: 'peligro', alClic: ev => eliminarEmbarque(vivo('embarques', e), ev.currentTarget) }];
    if (a === 'anular') return [{ texto: 'Anular', accion: 'anular', clase: 'peligro', alClic: ev => anularEmbarque(vivo('embarques', e), ev.currentTarget) }];   // C-24: el boton se deshabilita mientras dura
    return [];
}

/**
 * Deshacer una captura (pedido de Carlos, 2026-09-05). Dos caminos, decididos por la etapa:
 *   - sin folio (etapa compuerta): se BORRA el renglon; no hay nada que conservar.
 *   - con folio (R- o E-): se ANULA con motivo obligatorio; el folio no se reutiliza (ticket 03 §4)
 *     y las fotos ya subidas se quedan en el buzon con su _lote.json apuntando al folio anulado.
 * Despues se vuelve a correr la puerta y sale un folio nuevo.
 */
async function eliminarEmbarque(e, btn) {
    if (accionCorreccion(e) !== 'eliminar') { avisar('Este embarque ya tiene folio: se anula, no se elimina.', 'error'); return; }
    await escribiendo(btn, async () => {   // C-24
    const { ok, motivo } = await confirmar({
        titulo: 'Eliminar la captura', peligro: true, ok: 'Eliminar',
        texto: `${e.PlacaTractor}${e.Manifiesto ? ' · ' + e.Manifiesto : ''}, registrada ${horaCorta(e.Arribo)}. Todavía no tiene folio: el renglón se borra y no queda rastro en la app (SharePoint conserva la papelera).`,
        motivo: 'opcional', etiquetaMotivo: 'Por qué (queda solo en esta pantalla)'
    });
    if (!ok) return;
    try {
        await refrescarCliente();
        await estado.cliente.borrarRenglon(estado.siteId, L.embarques, e.id);
        estado.embarques = estado.embarques.filter(x => x.id !== e.id);
        avisar(`Captura eliminada${motivo ? ' (' + motivo + ')' : ''}. Si la góndola sigue en la puerta, vuelve a correr la compuerta.`, 'bien');
        repintar();   // C-14 (v0.25.0): como trasCambioPrealta. Antes iba a irA, que borraba el aviso «bien» y saltaba al tope
    } catch (err) { avisar('No se pudo eliminar: ' + (err && err.message ? err.message : err), 'error'); }
    });
}
async function anularEmbarque(e, btn) {
    if (accionCorreccion(e) !== 'anular') return;
    // U-70 (v0.40.0): una gondola con certificado vigente no se anula dejandolo vivo. Solo gerencia cancela certificados,
    // asi que solo gerencia puede anularla; y el certificado se cancela PRIMERO con el mismo motivo: si luego falla la
    // anulacion, queda una gondola cerrada sin certificado (se vuelve a emitir), nunca uno vigente sobre una anulada.
    const vigC = certificadoVigente(e);
    if (vigC && !PUEDE.emitirCertificado(estado.rol)) { avisar(`${e.Title} tiene el certificado ${vigC.Title} vigente: solo gerencia puede anular esta góndola (el certificado se cancela con ella).`, 'error'); return; }
    await escribiendo(btn, async () => {   // C-24 / C-23: boton deshabilitado y sin refresco mientras el confirm esta abierto
    const { ok, motivo } = await confirmar({
        titulo: `Anular ${e.Title}`, peligro: true, ok: 'Anular con este motivo', motivo: true,
        texto: `${e.PlacaTractor} · etapa ${e.Etapa}${e.NetoKg ? ` · neto ${kgG(e.NetoKg)}` : e.BrutoKg ? ` · bruto ${kgG(e.BrutoKg)}` : ''}. El folio ${e.Title} queda anulado y NO se vuelve a usar; el siguiente pesaje nace con folio nuevo. Las fotos ya subidas se conservan.${vigC ? ` Su certificado de tratamiento ${vigC.Title} se CANCELA con el mismo motivo: su QR dirá «cancelado».` : ''}`,
        etiquetaMotivo: 'Motivo de la anulación'
    });
    if (!ok) return;
    try {
        await refrescarCliente();
        if (vigC) {
            const cc = camposCancelacion(`Góndola ${e.Title} anulada: ${motivo}`);
            try { await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, vigC.id, cc); }
            catch (err) { throw new Error(`no se canceló el certificado ${vigC.Title} (${err.message}); la góndola NO se anuló`); }
            aplicar('certificados', vigC, cc);
        }
        const ahora = new Date().toISOString();
        const sello = `ANULADO por ${estado.cuenta.username} el ${horaCorta(ahora)} (estaba en ${e.Etapa}): ${motivo}`;
        const campos = { Etapa: 'anulado', AnuladoPor: estado.cuenta.username, AnuladoEl: ahora, AnuladoMotivo: motivo,
            Notas: e.Notas ? `${e.Notas}\n${sello}` : sello };
        try {
            await estado.cliente.actualizarRenglon(estado.siteId, L.embarques, e.id, campos);
        } catch (err) {
            // Lista sin actualizar: la opcion 'anulado' o las columnas Anulado* no existen todavia.
            if (esColumnaFaltante(err))   // C-15
                throw new Error('la app no puede anular todavía: avisa a gerencia' + (estado.rol === 'gerencia' ? ' (la lista PLANTA_Embarques no tiene la etapa «anulado» ni las columnas de anulación: herramientas-dev/provisionar.html, setup tarea 5). Detalle: ' + err.message : '.'));   // U-31
            throw err;
        }
        aplicar('embarques', e, campos);   // C-23
        avisar(`${e.Title} anulado. Para repesar la góndola, vuelve a correr la compuerta: saldrá un folio nuevo.`, 'bien');
        repintar();   // C-14 (v0.25.0): como trasCambioPrealta. Antes iba a irA, que borraba el aviso «bien» y saltaba al tope
    } catch (err) { avisar('No se pudo anular: ' + (err && err.message ? err.message : err), 'error'); }
    });
}

export async function autorizarExcepcion(e, btn) {
    if (!PUEDE.autorizarExcepcion(estado.rol)) return;
    await escribiendo(btn, async () => {   // C-24 / C-23
    const { ok } = await confirmar({ titulo: 'Autorizar la excepción', ok: 'Autorizar',
        texto: `${e.PlacaTractor} · ${nombreDe(estado.carriers, e.CarrierId)}. Motivo que dio la caseta: «${e.ExcepcionMotivo || 'sin motivo'}». Queda colgada de este embarque, no del carrier.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // S-01: primero la firma (403 si la cuenta no esta en Planta-Firmantes, que junta validador y gerencia), luego el sello.
        // El ROL lo exige firmaDe() al leer (S-11): una firma de excepcion de un validador no vale aunque el POST pase.
        await firmar('excepcion', e, e.ExcepcionMotivo);
        const campos = { ExcepcionAutorizo: estado.cuenta.username, ExcepcionEl: new Date().toISOString() };
        await estado.cliente.actualizarRenglon(estado.siteId, L.embarques, e.id, campos);
        aplicar('embarques', e, campos);   // C-23
        avisar('Excepción autorizada. Queda colgada de este embarque, no del carrier.', 'bien');
        pintarInsignias();
        pintarHoy();
    } catch (err) { avisar('No se pudo autorizar: ' + err.message, 'error'); }
    });
}

export function abrirPesaje(e, fase) {
    estado.pesando = { embarque: e, fase };
    estado.fotoBytes = null; soltarFotoPrevia();
    mostrarAsistente(e, 'baPesar', fase === 'bruto' ? 3 : 4);   // tanda 5: el pesaje es un paso del asistente, no una tarjeta bajo la lista
    // U-44 (v0.29.0): las dos fases titulan con la PLACA (lo que el basculista ve en el camión) y la tara además con el folio;
    // el manifiesto va al subtítulo. Antes la tara decía solo «Tara · E-26-00004» y el bruto no traía el manifiesto.
    $('baTitulo').textContent = fase === 'bruto' ? `Bruto · ${e.PlacaTractor}` : `Tara · ${e.PlacaTractor} · ${e.Title}`;
    const u = porId(estado.unidades, e.UnidadId);
    const manif = e.Manifiesto ? `Manifiesto ${e.Manifiesto}. ` : '';
    $('baSub').textContent = fase === 'bruto' ? `${manif}Primera pasada: la góndola cargada. Al guardar nace el folio E-.` :
        `${manif}Segunda pasada: la góndola vacía. Bruto ${kgG(e.BrutoKg)}${u && u.CapacidadKg ? ` · capacidad ${kgG(u.CapacidadKg)}` : ''}.`;   // U-87: con separador de miles, como la resta
    $('baKgLabel').textContent = fase === 'bruto' ? 'Bruto (kg)' : 'Tara (kg)';
    $('baKg').value = '';
    $('baFotoPrevia').classList.add('oculto');
    $('baFotoEstado').textContent = 'Sin foto todavía'; $('baFotoEstado').classList.remove('lista');
    $('btnFoto').textContent = 'Tomar foto';
    $('baNetoVivo').classList.toggle('oculto', fase !== 'tara');
    $('baEcuacion').classList.toggle('solo', fase !== 'tara');   // M7: el bruto es un solo término; la tara, la resta entera
    $('btnGuardarPeso').textContent = fase === 'bruto' ? 'Guardar bruto ›' : 'Guardar tara';
    if (fase === 'tara') {
        $('baNvBruto').textContent = Number(e.BrutoKg).toLocaleString('es-MX');
        $('baNvHora').textContent = e.BrutoHora ? ` · ${horaMexico(e.BrutoHora, 'hora')}` : '';
        $('baNvCap').textContent = u && u.CapacidadKg ? Number(u.CapacidadKg).toLocaleString('es-MX') : '';
        $('baNvCapBox').classList.toggle('oculto', !(u && u.CapacidadKg));
        revisarNeto();
    }
    // v0.39.0: el ticket de bascula (el que imprime el indicador) se captura al cerrar la TARA — cubre tara y destara,
    // y es lo que va impreso en el certificado de esa gondola (Carlos, 2026-09-22).
    $('baTicketCampo').classList.toggle('oculto', fase !== 'tara');
    $('baTicketBascula').value = '';
    $('baAvisoNeto').classList.add('oculto');
    $('baMotivoNetoCampo').classList.add('oculto');
    $('baMotivoNeto').value = '';
    if (ESCRITORIO.matches) $('baKg').focus();   // U-53: con ratón el foco va a los kilos; en celular el teclado es propio (inputmode none)
}

/** Revoca el blob URL de la vista previa (C-21, v0.25.0): cada foto comprimida quedaba viva hasta recargar la PWA. */
export function soltarFotoPrevia() {
    if (estado.fotoUrl) { URL.revokeObjectURL(estado.fotoUrl); estado.fotoUrl = null; }
}
export async function tomarFoto(archivo) {
    try {
        const { bytes } = await comprimir(archivo);
        estado.fotoBytes = bytes;
        const img = $('baFotoPrevia');
        soltarFotoPrevia();   // C-21: el JPEG anterior («Repetir») se libera; antes vivia en memoria todo el turno
        estado.fotoUrl = URL.createObjectURL(new Blob([bytes], { type: 'image/jpeg' }));
        img.src = estado.fotoUrl;
        img.classList.remove('oculto');
        $('baFotoEstado').textContent = `Lista · tomada ${horaCorta(new Date().toISOString())}`; $('baFotoEstado').classList.add('lista');
        $('btnFoto').textContent = 'Repetir';
    } catch (e) { avisar('No se pudo leer la foto: ' + e.message, 'error'); }
}

export function revisarNeto() {
    const p = estado.pesando; if (!p || p.fase !== 'tara') return null;
    const u = porId(estado.unidades, p.embarque.UnidadId);
    const aviso = avisoNeto(p.embarque.BrutoKg, $('baKg').value, u ? u.CapacidadKg : null, CONFIG.tolerancia);
    $('baAvisoNeto').classList.toggle('oculto', !aviso);
    $('baMotivoNetoCampo').classList.toggle('oculto', !aviso);
    if (aviso) $('baAvisoNeto').textContent = 'Revisa el indicador: ' + aviso;
    // I2: el neto se ve conforme se teclea; ambar si se sale de la banda; el boton dice lo que va a cerrar.
    const tara = Number($('baKg').value), bruto = Number(p.embarque.BrutoKg);
    const neto = Number.isFinite(tara) && tara > 0 ? bruto - tara : null;
    $('baNvNeto').textContent = neto === null ? '—' : neto.toLocaleString('es-MX');
    $('baNetoVivo').classList.toggle('neto-ok', neto !== null && !aviso);
    $('baNetoVivo').classList.toggle('neto-mal', neto !== null && !!aviso);
    $('btnGuardarPeso').textContent = neto === null || neto <= 0 ? 'Guardar tara' : `Guardar tara · neto ${neto.toLocaleString('es-MX')} kg ›`;
    return aviso;
}

export async function guardarPeso() {
    const p = estado.pesando; if (!p) return;
    const kg = Number($('baKg').value);
    // Kilos ENTEROS (el teclado y step=1 ya lo son; en computadora se podía teclear «12500.5» y el neto salía 11500.199999…)
    if (!Number.isInteger(kg) || kg <= 0) { avisar('Captura el peso en kilogramos, sin decimales.', 'error'); return; }
    // La tara igual o mayor que el bruto no es «fuera de banda»: daría neto 0 o negativo, y esa góndola ya no certifica.
    if (p.fase === 'tara' && kg >= Number(p.embarque.BrutoKg)) { avisar(`La tara (${kgG(kg)}) no puede ser igual o mayor que el bruto (${kgG(p.embarque.BrutoKg)}): revisa el indicador.`, 'error'); return; }
    if (!estado.fotoBytes) { avisar('Falta la foto del indicador: es lo que hace comprobable un peso tecleado.', 'error'); return; }
    // El ticket de bascula es obligatorio al cerrar: despues no hay donde capturarlo, y el certificado de la gondola lo lleva impreso.
    if (p.fase === 'tara' && !$('baTicketBascula').value.trim()) { avisar('Falta el número del ticket de báscula: es lo que va impreso en el certificado de esta góndola.', 'error'); $('baTicketBascula').focus(); return; }
    const aviso = revisarNeto();
    if (aviso && !$('baMotivoNeto').value.trim()) { avisar('El neto se sale de la banda: re-captura, o di por qué se cierra igual.', 'error'); return; }
    await escribiendo('btnGuardarPeso', async () => {   // C-24 / C-23: boton deshabilitado y sin refresco mientras sube
    $('avance').classList.remove('oculto');
    const paso = t => { $('textoAvance').textContent = t; };
    try {
        await refrescarCliente();
        const e = p.embarque;
        const ahora = new Date().toISOString();
        if (p.fase === 'bruto') await guardarBruto(e, kg, ahora, paso);
        else await guardarTara(e, kg, ahora, aviso, paso);
        soltarFotoPrevia(); estado.pesando = null; estado.fotoBytes = null;   // C-51: el JPEG ya subió; la pausa y el ticket no lo necesitan
        pintarTicket(e);   // tras el bruto también: «Ver ticket» de la pausa lo reimprime con el folio recién nacido
        pintarListasGondolas();   // la lista de atrás ya al día; el asistente sigue en pantalla
        // M5 / M8: el bruto lleva a la pausa de la descarga; la tara, al ticket con «Terminar».
        if (p.fase === 'bruto') mostrarPausa(e); else mostrarAsistente(e, 'baTicketCaja', 5);
    } catch (err) {
        pintarCapturado(p.embarque);   // decisión 11: si otra sesión avanzó la góndola, lo capturado ya dice lo suyo
        avisar('No se pudo guardar: ' + (err && err.message ? err.message : err), 'error');
    } finally {
        $('avance').classList.add('oculto');
    }
    });
}
// Si el PATCH falla despues de subir la foto, el lote se retira del buzon: antes quedaba una
// carpeta huerfana con _lote.json de un embarque que nunca avanzo (auditoria 2026-09-05).
async function guardarConLote(e, lote, campos, paso) {
    try { await estado.cliente.actualizarRenglon(estado.siteId, L.embarques, e.id, campos, paso); }
    catch (err) { try { await estado.cliente.borrarItemDrive(estado.siteId, lote.carpetaId); invalidarRama(CONFIG.buzon); } catch (_) { /* se reporta el error original */ } throw err; }
}
// C-26 (v0.28.0): las dos ramas de guardarPeso, cada una con su relectura; el re-anclaje en la ventana es anclar().
async function guardarBruto(e, kg, ahora, paso) {
    // Aqui nace el folio E-: el ticket se imprime con su numero desde la puerta. Se relee la
    // lista antes de escoger el numero (otro celular pudo tomar uno hace un segundo).
    paso('Asignando folio…');
    const delAno = await embarquesDelAno(paso);
    // `e` sigue siendo EL objeto del embarque: se le vuelcan los campos frescos y ocupa su
    // lugar en la lista de la ventana (si no, tendria una copia vieja en 'compuerta').
    const fresco = delAno.find(x => x.id === e.id);
    if (fresco) Object.assign(e, fresco);
    fundirEnVentana(delAno);
    anclar('embarques', e);
    // Decisión 11 (tanda 5): gana el primero, y se dice qué dejó y quién. Nada se sube: el lote todavía no existe.
    if (e.Etapa !== 'compuerta') throw new Error(`esta góndola ${yaCapturado(e, quien)} (lo movió otra sesión). Tu peso no se guardó; vuelve a la lista.`);
    // C-03 (v0.23.0): el folio se RESERVA (PATCH Title) y se confirma unico ANTES de subir la foto, para que la
    // carpeta, la foto y el _lote.json nazcan con el definitivo. Antes, si asegurarFolioUnico renumeraba, el lote ya
    // subido y BrutoFoto se quedaban con el folio viejo, que ahora era de otra gondola. Si la subida o el PATCH
    // final fallan, el renglon queda en compuerta CON folio y el reintento lo reusa (no nace otro numero).
    const folio = /^E-/.test(e.Title || '') ? e.Title : siguienteFolio('E', delAno.map(x => x.Title));
    await estado.cliente.actualizarRenglon(estado.siteId, L.embarques, e.id, { Title: folio }, paso);
    e.Title = folio;
    await asegurarFolioUnico(e, 'E', paso);
    paso('Subiendo la foto…');
    const lote = await subirEvidencia(e.Title, 'bruto', kg, paso);
    paso('Guardando…');
    const campos = { Etapa: 'bruto', BrutoKg: kg, BrutoHora: ahora, BrutoFoto: lote.ref };
    await guardarConLote(e, lote, campos, paso);
    aplicar('embarques', e, campos);
    avisar(`Bruto guardado. Folio ${e.Title}. La góndola puede descargar en la fosa.`, 'bien');
}
async function guardarTara(e, kg, ahora, aviso, paso) {
    // C-13 (v0.25.0): la tara relee SU renglon antes de subir la foto, como el bruto relee el anio. Gerencia pudo
    // anular la gondola desde Hoy mientras el basculista tecleaba: el PATCH de cierre pisaba «anulado» con «cerrado»
    // (renglon con AnuladoPor Y Etapa cerrado, contado en KPI y CSV). Y se re-ancla por id (C-12), como el bruto.
    paso('Revisando el embarque…');
    const vigente = await estado.cliente.renglon(estado.siteId, L.embarques, e.id, paso);
    Object.assign(e, vigente);
    anclar('embarques', e);
    if (e.Etapa !== 'bruto') throw new Error(`esta góndola ${yaCapturado(e, quien)} (lo movió otra sesión). Tu tara no se guardó; vuelve a la lista.`);   // decisión 11
    paso('Subiendo la foto…');
    const lote = await subirEvidencia(e.Title, 'tara', kg, paso);
    paso('Cerrando el embarque…');
    const neto = Number(e.BrutoKg) - kg;
    const campos = limpiar({ Etapa: 'cerrado', TaraKg: kg, TaraHora: ahora, TaraFoto: lote.ref, NetoKg: neto,
        TicketBascula: $('baTicketBascula').value.trim() || null,
        InicioAlmacen: e.BrutoHora || ahora,
        Notas: aviso ? `Neto fuera de banda (${aviso}). Motivo: ${$('baMotivoNeto').value.trim()}` : null });
    await guardarConLote(e, lote, campos, paso);
    aplicar('embarques', e, campos);
    avisar(`Góndola ${e.Title} cerrada: neto ${kgG(neto)}.`, 'bien');   // U-87: el vocabulario de la sección
}

/**
 * La foto del indicador va al buzon como un lote del contrato de nombres, con `_lote.json`
 * (ticket 06: la app apunta, no guarda). Devuelve el par (fecha, concepto) que el embarque conserva.
 */
/** `<destinoBase>/<AAAA>/<AAAA-MM>` a partir de la fecha del lote (AAAA-MM-DD). Carpeta por mes: ver config.js. */
export function destinoEvidencia(fecha) {
    return `${CONFIG.evidencia.destinoBase}/${fecha.slice(0, 4)}/${fecha.slice(0, 7)}`;
}

export async function subirEvidencia(folio, fase, kg, avisar) {
    const fecha = fechaMexico();
    const concepto = `${fase} ${folio} ${kg} kg`;
    const s = slug(concepto);
    const carpeta = `${fecha}_${CONFIG.evidencia.etiqueta}_${s}`;
    const { nombreReal, id: carpetaId } = await estado.cliente.crearCarpeta(estado.siteId, CONFIG.buzon, carpeta, avisar);
    const ruta = `${CONFIG.buzon}/${nombreReal}`;
    const archivo = `${fecha}_CALYTEK_Foto_${s}-01.jpg`;
    await estado.cliente.subirPieza(estado.siteId, ruta, archivo, estado.fotoBytes, 'image/jpeg', avisar);
    const manifiesto = {
        app: 'calytek-planta', contrato: 1, app_version: VERSION, unidad: 'CALYTEK',
        etiqueta: CONFIG.evidencia.etiqueta, destino: destinoEvidencia(fecha), tipo: 'foto',
        fecha, concepto, archivos: [archivo], paginas: 1, subido: new Date().toISOString(),
        embarque: folio
    };
    await estado.cliente.subirPieza(estado.siteId, ruta, '_lote.json',
        new TextEncoder().encode(JSON.stringify(manifiesto, null, 2) + '\n'), 'application/json', avisar);
    invalidarRama(CONFIG.buzon);   // C-31 (v0.34.0): Archivos ya no muestra el buzón de antes de este lote
    return { ref: `${fecha}|${concepto}`, carpetaId };
}

/** Ticket en ventana (v0.19.1): `lista` es lo que se ve en Cerrados; ‹ › y las flechas recorren. */
export function abrirTicketPop(lista, i) {
    const d = $('dlgTicket');
    // C-53 (v0.51.0): el renglón se resuelve por id en cada ‹ ›, como vivo(): la lista con que se pintó puede traer copias viejas.
    const actual = () => vivo('embarques', lista[i]);
    const pintar = () => {
        pintarTicket(actual(), $('ticketPop'));
        $('tkPos').textContent = `${i + 1} de ${lista.length} · ${actual().Title}`;
        $('tkAnt').disabled = i <= 0; $('tkSig').disabled = i >= lista.length - 1;
    };
    $('tkAnt').onclick = () => { if (i > 0) { i--; pintar(); } };
    $('tkSig').onclick = () => { if (i < lista.length - 1) { i++; pintar(); } };
    d.onkeydown = ev => { if (ev.key === 'ArrowLeft') $('tkAnt').onclick(); else if (ev.key === 'ArrowRight') $('tkSig').onclick(); };
    pintar();
    if (!d.open) d.showModal();
}
$('tkCerrar').addEventListener('click', () => $('dlgTicket').close());
$('tkImprimir').addEventListener('click', () => window.print());


function pintarTicket(e, t = $('ticket')) {
    t.textContent = '';
    t.appendChild(el('h3', '', 'CALYTEK · Planta de tratamiento de RME — Ticket de báscula'));
    const pre = porId(estado.prealtas, e.PreAltaId);
    const filas = [
        ['Folio', e.Title], ['Manifiesto', e.Manifiesto || '—'],
        ['Generador', pre ? `${pre.Generador || ''} ${pre.GeneradorRegistro ? '(' + pre.GeneradorRegistro + ')' : ''}` : '—'],
        ['Pozo / corriente', pre ? `${pre.Pozo || '—'} · ${etiquetaCorriente(e.CorrienteDeclarada || pre.Corriente)}` : (etiquetaCorriente(e.CorrienteDeclarada) || '—')],   // C-55
        ['Transportista', nombreDe(estado.carriers, e.CarrierId)],
        ['Placas', [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' / ') || '—'], ['Operador', e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId)],   // U-87: placas sin la diagonal colgando
        ['Arribo', horaCorta(e.Arribo)],
        ['Bruto', e.BrutoKg ? `${kgG(e.BrutoKg)} · ${horaCorta(e.BrutoHora)}` : '—'],
        ['Tara', e.TaraKg ? `${kgG(e.TaraKg)} · ${horaCorta(e.TaraHora)}` : 'pendiente'],
        ['NETO', e.NetoKg ? kgG(e.NetoKg) : 'pendiente'],
        ['Ticket de báscula', e.TicketBascula || '—'],
        ['Capturó', quien(e.CapturadoPor) || '—'],   // U-122 (v0.73.0): vacío lleva raya, como los demás renglones
        ['Emitido', `${horaCorta(new Date().toISOString())} · CALYTEK Planta ${VERSION}`]
    ];
    if (e.Etapa === 'anulado') filas.unshift(['ANULADO', `${horaCorta(e.AnuladoEl)} · ${quien(e.AnuladoPor)} · ${e.AnuladoMotivo || ''}`]);
    const tabla = el('table');
    for (const [k, v] of filas) {
        const tr = el('tr'); tr.appendChild(el('td', '', k)); tr.appendChild(el('td', k === 'NETO' ? 'grande' : 'mono', v)); tabla.appendChild(tr);
    }
    t.appendChild(tabla);
}


$('baBusca').addEventListener('input', pintarHistorial);

for (const b of document.querySelectorAll('#gTabs button, #gKpis button')) b.addEventListener('click', () => elegirVistaGondolas(b.dataset.g));
