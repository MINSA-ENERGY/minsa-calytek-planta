// CALYTEK Planta — Hoy (la consola).
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { fechaCorta, fechaMexico, horaMexico, lunesDe, plural, PUEDE } from './reglas.js';
import { $, avisar, botonAccion, el, enListaPlanta, estado, etiqueta, excepcionesPendientes, firmar, horaCorta, iso, motivoSinFirmas, nombreDe, palabraCompuertaDe, quien, selloSinFirma, vivo } from './nucleo.js';
import { autorizarExcepcion, botonCorreccion } from './gondolas.js';
import { celda } from './prealtas.js';
import { pintarPendientesHoy, pintarRechazosHoy, pintarVigenciasHoy } from './archivos.js';

// ================================================================ HOY (la consola)

// Tanda 3 (v0.48.0): la Fila del dia se fundio en Gondolas; de sus piezas queda la etiqueta de compuerta (Hoy y Reportes).
export function etiquetaCompuertaDe(e) {
    if (e.Etapa === 'anulado') return etiqueta('anulado', 'anulado');
    const pc = palabraCompuertaDe(e);   // C-55
    return etiqueta(pc.corta, pc.clase);
}
// Franja: la excepcion que espera a gerencia es la unica decision que «Hoy» le pide a alguien.
function pintarFranjaHoy(pendientes, botonesExcepcion) {
    const fr = $('hoyFranja'); fr.textContent = '';
    fr.classList.toggle('oculto', !pendientes.length);
    for (const e of pendientes) {
        const it = el('div', 'item');
        it.appendChild(el('b', 'w', 'Espera'));
        it.appendChild(el('span', 't', `${e.PlacaTractor} · ${nombreDe(estado.carriers, e.CarrierId)} · ${e.Manifiesto || 'sin manifiesto'}`));
        it.appendChild(el('span', 's', `${selloSinFirma(e)}«${e.ExcepcionMotivo || 'sin motivo'}» · ${quien(e.CapturadoPor) || '?'}, ${horaCorta(e.Arribo)}`));   // U-28: nombre, no correo
        const bs = botonesExcepcion(e);
        if (bs.length) { const d = el('div', 'botones'); for (const b of bs) d.appendChild(botonAccion(b)); it.appendChild(d); }
        fr.appendChild(it);
    }
}

// KPI: numero, tendencia y techo.
export function pintarKpisReportes({ cerradosHoy, cerradosAyer, cerradosSemana, activos, rechazosSemana, borradores }) {
    const kg = xs => xs.reduce((a, e) => a + (Number(e.NetoKg) || 0), 0);
    const k = $('tbKpis'); k.textContent = '';
    const kpi = (l, n, unidad, t, clase) => {
        const d = el('div', 'kpi ' + ({ mal: 'is-danger', ojo: 'is-warn', ok: 'is-ok', info: 'is-info' }[clase] || '')); d.appendChild(el('div', 'l', l));
        const num = el('div', 'n' + (clase === 'mal' || clase === 'ojo' ? ' ' + clase : ''), String(n)); if (unidad) num.appendChild(el('small', '', unidad)); d.appendChild(num);
        if (t) d.appendChild(t); k.appendChild(d); return d;
    };
    const tend = (h, a, texto) => {
        const t = el('div', 't'); const dif = h - a;
        if (dif !== 0) t.appendChild(el('span', dif > 0 ? 'sube' : 'baja', (dif > 0 ? '▲ ' : '▼ ') + Math.abs(dif) + ' '));
        t.appendChild(document.createTextNode(texto)); return t;
    };
    kpi('Góndolas cerradas hoy', cerradosHoy.length, null, tend(cerradosHoy.length, cerradosAyer.length, 'vs ayer'), 'ok');
    kpi('Toneladas netas hoy', (kg(cerradosHoy) / 1000).toFixed(1), 't', el('div', 't', cerradosHoy.length ? `${(kg(cerradosHoy) / 1000 / cerradosHoy.length).toFixed(1)} t por góndola · semana ${(kg(cerradosSemana) / 1000).toFixed(1)} t` : `semana ${(kg(cerradosSemana) / 1000).toFixed(1)} t`), 'ok');
    const enP = kpi('En planta ahora', activos.length, null, el('div', 't', `${activos.filter(e => e.Etapa === 'bruto').length} por tara · ${activos.filter(e => e.Etapa === 'compuerta').length} por bruto`), 'info');
    const med = el('div', 'medidor'); const mi = el('i'); mi.style.width = Math.min(100, Math.round(((cerradosHoy.length + activos.length) / CONFIG.techoGondolasDia) * 100)) + '%'; med.appendChild(mi); enP.appendChild(med);
    enP.appendChild(el('div', 't', `techo ${CONFIG.techoGondolasDia} al día`));
    kpi('Rechazos esta semana', rechazosSemana.length, null, el('div', 't', rechazosSemana.length ? 'legal · el residuo no entró' : 'ninguno'), rechazosSemana.length ? 'mal' : '');
    kpi('Pre-altas por firmar', borradores.length, null, el('div', 't', borradores.length ? 'esperan al validador' : 'todas firmadas'), borradores.length ? 'ojo' : '');
}


/** Los cortes de fecha que comparten Hoy y Reportes. Una gondola CERRADA cuenta el dia en que se cerro (hora de la tara),
 *  no el de arribo: la que llega 23:50 y cierra 00:10 es del dia siguiente, que es el que reporta la bascula (F1, 5-sep). */
export function cortesDia() {
    const hoy = fechaMexico();
    const ayer = fechaMexico(new Date(Date.now() - 86400000));
    const lunes = lunesDe(hoy);   // C-30 (v0.34.0): el mismo lunes que la barra de semanas de Reportes (antes: zona del dispositivo)
    const dia = e => e.Arribo ? fechaMexico(new Date(e.Arribo)) : '';
    const diaCierre = e => e.TaraHora ? fechaMexico(new Date(e.TaraHora)) : dia(e);
    const cerrados = f => estado.embarques.filter(e => e.Etapa === 'cerrado' && f(diaCierre(e)));
    return { hoy, ayer, lunes, dia, diaCierre, cerrados };
}
export function pintarHoy() {
    const activos = enListaPlanta();   // tanda 3: el mismo número que el rail y la pestaña En planta
    const pendientes = excepcionesPendientes();
    const borradores = estado.prealtas.filter(p => p.Estado === 'borrador');

    // v0.32.0: la fecha es el subtítulo; los KPI viven en Reportes. v0.66.0: el kicker es fijo (la banda de Pre-altas) y el rol ya no va aquí.
    $('hoyTitulo').textContent = `${new Date().toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City', weekday: 'long', day: 'numeric', month: 'long' })} · ${activos.length === 1 ? '1 góndola en planta' : `${activos.length} góndolas en planta`}${borradores.length ? ` · ${borradores.length === 1 ? '1 pre-alta por firmar' : `${borradores.length} pre-altas por firmar`}` : ''}`;

    // U-12 (v0.22.0): los mismos botones (Autorizar · Anular/Eliminar) salen en la franja y en su renglon de «Pendiente
    // revisar», que era la unica entrada de esa tarjeta sin accion; la excepcion se cuenta una sola vez (en la tarjeta).
    const botonesExcepcion = e => {
        const bs = [];
        if (PUEDE.autorizarExcepcion(estado.rol)) bs.push({ texto: 'Autorizar', accion: 'autorizar', clase: 'si', alClic: ev => autorizarExcepcion(vivo('embarques', e), ev.currentTarget), deshabilitado: motivoSinFirmas() });
        for (const b of botonCorreccion(e)) bs.push(b);   // U-52 (v0.29.0): Anular/Eliminar en rojo también en la franja; antes iban con la clase neutral «no»
        return bs;
    };
    pintarFranjaHoy(pendientes, botonesExcepcion);
    pintarPendientesHoy(borradores, pendientes, botonesExcepcion);
    pintarRechazosHoy();
    pintarVigenciasHoy();
}

/** CSV (UTF-8 con BOM, separado por coma, fechas en hora de Mexico) de todos los embarques cargados en la ventana. */
function exportarCsv() {
    const cab = ['Folio', 'Etapa', 'Compuerta', 'PlacaTractor', 'PlacaPlana', 'Carrier', 'Chofer', 'Manifiesto', 'Corriente', 'Programa', 'Arribo', 'BrutoKg', 'BrutoHora', 'TaraKg', 'TaraHora', 'NetoKg', 'CapturadoPor', 'ExcepcionAutorizo', 'AnuladoMotivo'];
    const hora = iso => iso ? horaMexico(iso, 'completa') : '';
    // Texto que empieza con = + - @ (o tabulador / retorno de carro, S-02) lleva apostrofo delante: Excel lo
    // ejecutaria como formula (auditoria 2026-09-08, hallazgo 4).
    const celda = v => { let t = v === null || v === undefined ? '' : String(v); if (/^[=+\-@\t\r]/.test(t)) t = "'" + t; return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };
    // Los nombres de columna son los de esquema.json (CorrienteDeclarada, PreAltaId): C-02, v0.21.0 — antes salian vacias.
    const filas = [...estado.embarques].sort((a, b) => a.id - b.id).map(e => [
        e.Title, e.Etapa, e.Compuerta, e.PlacaTractor, e.PlacaPlana, nombreDe(estado.carriers, e.CarrierId), e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId),
        e.Manifiesto, e.CorrienteDeclarada, nombreDe(estado.prealtas, e.PreAltaId), hora(e.Arribo), e.BrutoKg, hora(e.BrutoHora), e.TaraKg, hora(e.TaraHora), e.NetoKg,
        e.CapturadoPor, e.ExcepcionAutorizo, e.AnuladoMotivo]);
    const csv = '\ufeff' + [cab, ...filas].map(f => f.map(celda).join(',')).join('\r\n');
    const a = document.createElement('a');
    // C-78 (v0.73.0): el blob anterior se suelta al exportar otra vez, no por reloj (5 s no alcanzaban en un celular lento).
    if (exportarCsv.url) URL.revokeObjectURL(exportarCsv.url);
    a.href = exportarCsv.url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = `CALYTEK_Embarques_${fechaCorta(estado.ventanaDesde).replace(/\//g, '-')}_a_${fechaMexico()}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    avisar(`CSV con ${plural(filas.length, 'embarque')} descargado.`, 'bien');
}
$('btnExportar').addEventListener('click', exportarCsv);
