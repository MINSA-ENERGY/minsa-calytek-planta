// CALYTEK Planta — Hoy (la consola).
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { diasPara, esRechazo, fechaCorta, fechaMexico, horaMexico, lunesDe, plural, PUEDE, reglasDe } from './reglas.js';
import { $, avisar, botonAccion, el, enListaPlanta, enPlanta, esperaAutorizacion, estado, etiqueta, excepcionesPendientes, horaCorta, motivoSinFirmas, nombreDe, palabraCompuertaDe, prealtaCambioTrasFirma, quien, renglon, selloSinFirma, sellosSinFirma, vivo } from './nucleo.js';
import { autorizarExcepcion, botonCorreccion } from './gondolas.js';
import { irA } from './navegacion.js';
import { irPadron, NOMBRE_PADRON, VIGENCIAS_PADRON, vigenciasDelCarrier } from './padron.js';
import { sinMovimientoDe, verPrealta } from './prealtas.js';

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
    avisar(`CSV con ${plural(filas.length, 'góndola')} descargado.`, 'bien');
}
// C-86 (v0.79.0): lo que pinta Hoy vivía en archivos.js desde C-76; vuelve a su pantalla.
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
    for (const p of ssf) pf.appendChild(renglon(`Firma · ${p.Title}`, prealtaCambioTrasFirma(p) ? `cambió después de que la firmó ${quien(p.FirmadaPor) || '?'} (carrier, unidades, choferes, corriente o generador) · la puerta no la ve · se vuelve a firmar desde su detalle` : `falta la firma de la Responsable Ambiental (trae el sello de ${quien(p.FirmadaPor) || '?'}, sin firma registrada) · la puerta no la ve · se firma desde su detalle`, 'Ver', () => verPrealta(vivo('prealtas', p))));   // U-28 / U-31
    for (const { p, sm } of dormidas) pf.appendChild(renglon(`Programa · ${p.Title}`, `${sm.motivo} · ¿se cierra? Sigue saliendo en la puerta`, 'Ver', () => verPrealta(vivo('prealtas', p))));
    for (const p of borradores) { const d = diasPara(p.FechaEstimada); pf.appendChild(renglon(`Pre-alta · ${p.Title}`, `firma de la Responsable Ambiental · 1er envío ${fechaCorta(p.FechaEstimada)}${d !== null ? ` (en ${d} días)` : ''} · capturó ${quien(p.CapturadaPor) || '?'}`, 'Ver', () => verPrealta(vivo('prealtas', p)))); }
    for (const e of pendientes) pf.appendChild(renglon(`Excepción · ${e.PlacaTractor}`, `${selloSinFirma(e)}autorización de gerencia · «${e.ExcepcionMotivo || 'sin motivo'}» · ${horaCorta(e.Arribo)}`, null, null, botonesExcepcion(e).map(b => ({ ...b, clase: b.accion === 'autorizar' ? '' : 'peligro' }))));
}

// C-29 (v0.34.0): UNA definición de «rechazo o excepción» y UN renglón para Hoy (los 10 últimos) y Reportes (todos). Antes el
// filtro y la lectura de CompuertaDetalle vivían copiados en los dos y la etiqueta difería (Hoy decía legal/comercial).
export const rechazosYExcepciones = () => estado.embarques.filter(esRechazo).sort((a, b) => b.id - a.id);
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
function fichaDelTablero(v) {
    if (v.Rol !== 'carrier') return null;
    const c = estado.carriers.find(x => vigenciasDelCarrier(x).includes(v));
    return c ? { clave: 'carriers', id: c.id, carrier: c.id } : null;
}
export function pintarVigenciasHoy() {
    const vg = $('tbVigencias'); vg.textContent = '';
    const prox = [...estado.vigencias.filter(v => v.Activo !== false), ...vigenciasPadronHoy()].map(v => ({ v, d: diasPara(v.Vence) })).filter(x => x.d !== null && x.d <= (Number(x.v.AvisoDias) || CONFIG.avisoVigenciaDias)).sort((a, b) => a.d - b.d);
    if (!prox.length) vg.appendChild(el('p', 'vacio', `Nada vence en ${CONFIG.avisoVigenciaDias} días.`));
    for (const { v, d } of prox) {
        const ventana = Number(v.AvisoDias) || CONFIG.avisoVigenciaDias;
        const clase = d < 0 ? 'mal' : d <= 7 ? 'ojo' : '';
        // U-134 (v0.74.0): la del padrón es un botón que abre la ficha del chofer, unidad o carrier.
        // U-159 (v0.78.0): la ASEA del carrier que vive en el tablero también abre la ficha de su carrier; las demás del tablero no tienen ficha.
        const ficha = v.ficha || fichaDelTablero(v);
        const r = el(ficha ? 'button' : 'div', ficha ? 'vig abre' : 'vig');
        if (ficha) { r.type = 'button'; r.title = 'Abrir su ficha en el Padrón'; r.addEventListener('click', () => abrirFichaDesdeHoy(ficha)); }
        const t = el('span', '', v.Title); t.appendChild(el('small', '', `${fechaCorta(v.Vence)} · ${v.Fuente || ''}${v.Dueno ? ' · dueño ' + v.Dueno : ''}`)); r.appendChild(t);
        r.appendChild(el('span', 'd ' + clase, d < 0 ? `−${-d} d` : `${d} d`));
        const bar = el('span', 'bar'); const i = el('i', clase); i.style.width = Math.max(4, Math.min(100, Math.round((1 - d / ventana) * 100))) + '%'; bar.appendChild(i); r.appendChild(bar);
        vg.appendChild(r);
    }
}


$('btnExportar').addEventListener('click', exportarCsv);
