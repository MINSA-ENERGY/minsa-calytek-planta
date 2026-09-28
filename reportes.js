// CALYTEK Planta — Reportes.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { fechaCorta, sumarDias } from './reglas.js';
import { $, el, enPlanta, estado, nombreDe, renglon } from './nucleo.js';
import { gondolasDe } from './prealtas.js';
import { cortesDia, rechazosYExcepciones, renglonRechazo } from './hoy.js';

// ================================================================ REPORTES (v0.32.0, sección aparte; artifact 1GvBJaYooYvjZT4rMRtL9Q)
// Los cinco KPI que vivían en Hoy, más lo que se lee de lo cargado: avance por programa, por carrier, toneladas por semana,
// rechazos/excepciones y netos fuera de banda. Todo sale de estado.embarques (CONFIG.ventanaDias más lo abierto): no lee nada más.
const mesCorto = f => new Date(f + 'T12:00:00Z').toLocaleDateString('es-MX', { timeZone: 'UTC', month: 'short' }).replace('.', '');
export function pintarReportes() {
    const { hoy, ayer, lunes, dia, diaCierre, cerrados } = cortesDia();
    const activos = estado.embarques.filter(enPlanta);
    const borradores = estado.prealtas.filter(p => p.Estado === 'borrador');
    const kg = xs => xs.reduce((a, e) => a + (Number(e.NetoKg) || 0), 0);
    const t = x => (x / 1000).toFixed(1);
    $('repSub').textContent = `Lo que entró, lo que pesó y lo que no pasó, sobre lo cargado (${CONFIG.ventanaDias} días más lo abierto · ${estado.embarques.length} góndolas).`;
    pintarKpisReportes({ cerradosHoy: cerrados(d => d === hoy), cerradosAyer: cerrados(d => d === ayer), cerradosSemana: cerrados(d => d >= lunes),
        activos, borradores, rechazosSemana: estado.embarques.filter(e => e.Etapa === 'rechazado' && dia(e) >= lunes) });

    // barras: nombre · barra proporcional · cifra en mono
    const barras = (caja, filas, vacio) => {
        caja.textContent = '';
        if (!filas.length) { caja.appendChild(el('p', 'vacio', vacio)); return; }
        const tope = Math.max(1, ...filas.map(f => f.tope ?? f.valor));
        for (const f of filas) {
            caja.appendChild(el('span', 'nom', f.nombre));
            const b = el('div', 'b'); const i = el('i', f.clase || ''); i.style.width = Math.min(100, Math.round((f.valor / (f.tope || tope)) * 100)) + '%'; b.appendChild(i); caja.appendChild(b);
            caja.appendChild(el('span', 'n', f.cifra));
        }
    };
    const abiertas = estado.prealtas.filter(p => p.Estado !== 'cerrada' && p.Estado !== 'cerrado');
    barras($('repProgramas'), abiertas.map(p => {
        const g = gondolasDe(p); const neto = kg(estado.embarques.filter(e => Number(e.PreAltaId) === p.id && e.Etapa === 'cerrado'));
        return { nombre: p.Title, valor: g.rec, tope: Math.max(g.esp, g.rec, 1), clase: p.Estado === 'borrador' ? 'warn' : '', cifra: g.esp ? `${g.rec} / ${g.esp} · ${t(neto)} t` : `${g.rec} · ${t(neto)} t` };
    }), 'Sin programas abiertos.');
    const porCarrier = new Map();
    for (const e of estado.embarques.filter(e => e.Etapa !== 'anulado')) { const k = e.CarrierId ?? '?'; const c = porCarrier.get(k) || { n: 0, kg: 0 }; c.n++; if (e.Etapa === 'cerrado') c.kg += Number(e.NetoKg) || 0; porCarrier.set(k, c); }
    barras($('repCarriers'), [...porCarrier].sort((a, b) => b[1].n - a[1].n).map(([id, c]) => ({ nombre: id === '?' ? 'Sin carrier' : nombreDe(estado.carriers, id), valor: c.n, cifra: `${c.n} gónd. · ${t(c.kg)} t` })), 'Sin góndolas cargadas.');

    // toneladas netas por semana (lunes a domingo), las últimas 8; la semana en curso en el color de marca
    const sem = $('repSemanas'); sem.textContent = '';
    // C-30 (v0.34.0): el mismo «lunes» que cortesDia() (reglas.lunesDe sobre la fecha de México); antes se calculaba aquí con la zona del dispositivo
    const semanas = []; const l0 = lunes;
    for (let i = 7; i >= 0; i--) { const desde = sumarDias(l0, -7 * i); semanas.push({ desde, hasta: sumarDias(desde, 6), kg: 0 }); }
    for (const e of estado.embarques.filter(e => e.Etapa === 'cerrado')) { const f = diaCierre(e); const s = semanas.find(s => f >= s.desde && f <= s.hasta); if (s) s.kg += Number(e.NetoKg) || 0; }
    const topeSem = Math.max(1, ...semanas.map(s => s.kg));
    for (const s of semanas) {
        const d = el('div', s.desde === l0 ? 'hoy' : '');
        d.appendChild(el('span', 'cifra', s.kg ? t(s.kg) : ''));   // U-61 (v0.34.0): la cifra a la vista; el title era solo hover
        const i = el('i'); i.style.height = Math.max(1, Math.round((s.kg / topeSem) * 100)) + '%'; i.title = `${t(s.kg)} t`; d.appendChild(i);
        // U-60 (v0.34.1): día y mes en spans; en celular solo se ve el día y el mes va una vez en el pie
        const eti = el('span', 'eti', String(Number(s.desde.slice(8, 10)))); eti.appendChild(el('span', 'mes', ' ' + mesCorto(s.desde))); d.appendChild(eti); sem.appendChild(d);
    }
    const enCurso = semanas[semanas.length - 1];
    $('repSemanasSub').textContent = `Semanas del ${Number(semanas[0].desde.slice(8, 10))} ${mesCorto(semanas[0].desde)} al ${Number(enCurso.hasta.slice(8, 10))} ${mesCorto(enCurso.hasta)}. Semana en curso en azul: ${t(enCurso.kg)} t al ${new Date().toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City', weekday: 'long' })}. Tope de la escala: ${t(topeSem)} t. Solo cuentan las cerradas dentro de lo cargado.`;

    // rechazos y excepciones (misma definición que la tarjeta de Hoy, sin el tope de 10) y netos fuera de banda
    const rj = $('repRechazos'); rj.textContent = '';
    const rech = rechazosYExcepciones();
    if (!rech.length) rj.appendChild(el('p', 'vacio', 'Ningún rechazo ni excepción en lo cargado.'));
    for (const e of rech) rj.appendChild(renglonRechazo(e));
    const nt = $('repNetos'); nt.textContent = '';
    const fuera = estado.embarques.filter(e => e.Etapa === 'cerrado' && /Neto fuera de banda/.test(e.Notas || '')).sort((a, b) => b.id - a.id);
    const nCerr = cerrados(() => true).length;
    if (!fuera.length) nt.appendChild(el('p', 'vacio', !nCerr ? 'Ninguno: todavía no hay góndolas cerradas en lo cargado.' : nCerr === 1 ? 'Ninguno: la góndola cerrada quedó dentro de la banda.' : `Ninguno: las ${nCerr} cerradas quedaron dentro de la banda.`));
    for (const e of fuera) nt.appendChild(renglon(`${e.Title} · ${e.PlacaTractor}`, `${fechaCorta(e.TaraHora || e.Arribo)} · neto ${Number(e.NetoKg).toLocaleString('es-MX')} kg · ${(e.Notas || '').replace(/\n.*$/s, '')}`));
}

// C-86 (v0.79.0): desde hoy.js, donde vivía desde C-76; solo los pinta Reportes.
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
    kpi('Toneladas netas hoy', (kg(cerradosHoy) / 1000).toFixed(1), 't', el('div', 't', cerradosHoy.length ? `${(kg(cerradosHoy) / 1000 / cerradosHoy.length).toFixed(1)}\u00a0t por góndola · semana ${(kg(cerradosSemana) / 1000).toFixed(1)}\u00a0t` : `semana ${(kg(cerradosSemana) / 1000).toFixed(1)}\u00a0t`), 'ok');   // espacio duro: la «t» caía sola en otro renglón
    const enP = kpi('En planta ahora', activos.length, null, el('div', 't', `${activos.filter(e => e.Etapa === 'bruto').length} por tara · ${activos.filter(e => e.Etapa === 'compuerta').length} por bruto`), 'info');
    const med = el('div', 'medidor'); const mi = el('i'); mi.style.width = Math.min(100, Math.round(((cerradosHoy.length + activos.length) / CONFIG.techoGondolasDia) * 100)) + '%'; med.appendChild(mi); enP.appendChild(med);
    enP.appendChild(el('div', 't', `techo ${CONFIG.techoGondolasDia} al día`));
    kpi('Rechazos esta semana', rechazosSemana.length, null, el('div', 't', rechazosSemana.length ? 'legal · el residuo no entró' : 'ninguno'), rechazosSemana.length ? 'mal' : '');
    kpi('Pre-altas por firmar', borradores.length, null, el('div', 't', borradores.length ? 'esperan a la Responsable Ambiental' : 'todas firmadas'), borradores.length ? 'ojo' : '');
}
