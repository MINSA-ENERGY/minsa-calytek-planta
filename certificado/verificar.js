// Verificacion publica del certificado de tratamiento (v0.35.0). Sin dependencias, sin login, sin Graph.
// ?f=<folio>-<sufijo> -> ./datos/<folio>-<sufijo>.json (lo publica docs/exportar-planta.ps1 -PublicarCertificados).
// Todo va por textContent: nada de lo que llega del JSON se interpreta como HTML.
(function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const RE_F = /^CT-\d{2}-\d{4}-[a-z2-9]{8}$/;   // folio + sufijo, tal cual lo pinta la app (reglas.sufijoVerificacion)
    const ETIQUETAS = [
        ['folio', 'Folio'], ['estado', 'Estado'], ['generador', 'Generador'], ['registro', 'Registro de generador'], ['pozo', 'Pozo'],
        ['residuo', 'Residuo'], ['toneladas', 'Volumen tratado'], ['embarques', 'Embarques'], ['fechas', 'Fecha de recepción'],
        ['transportista', 'Transportista'], ['emitido', 'Emitido']
    ];
    const ESTADO = {
        vigente: ['ok', 'VIGENTE — este certificado fue emitido por MINSA ENERGY y sigue válido.'],
        sustituido: ['ojo', 'SUSTITUIDO — este certificado fue reemplazado por otro. El papel con este folio ya no vale.'],
        cancelado: ['mal', 'CANCELADO — este certificado fue cancelado por MINSA ENERGY. El papel con este folio no vale.']
    };
    function estado(clase, texto) { const e = $('estado'); e.className = 'estado ' + clase; e.textContent = texto; }
    function pintar(d) {
        const [clase, texto] = ESTADO[d.estado] || ['mal', `Estado desconocido: ${d.estado}`];
        estado(clase, texto);
        const dl = $('datos'); dl.textContent = '';
        const filas = {
            folio: d.folio, estado: d.estado, generador: d.generador, registro: d.registro, pozo: d.pozo, residuo: d.residuo,
            toneladas: d.kg != null ? `${(Number(d.kg) / 1000).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} t (${Number(d.kg).toLocaleString('es-MX')} kg)` : null,
            embarques: Array.isArray(d.embarques) ? `${d.embarques.length}: ${d.embarques.join(', ')}` : null,
            fechas: d.fechas, transportista: d.transportista, emitido: d.emitidoEl ? new Date(d.emitidoEl).toLocaleString('es-MX', { timeZone: 'America/Mexico_City', dateStyle: 'long', timeStyle: 'short' }) : null
        };
        for (const [k, eti] of ETIQUETAS) {
            if (filas[k] == null || filas[k] === '') continue;
            const dt = document.createElement('dt'); dt.textContent = eti;
            const dd = document.createElement('dd'); dd.textContent = String(filas[k]); if (k === 'folio' || k === 'embarques') dd.className = 'mono';
            dl.appendChild(dt); dl.appendChild(dd);
        }
        dl.classList.remove('oculto');
        const nota = $('nota');
        if (d.estado === 'sustituido' && d.sustituidoPor) { nota.textContent = `Sustituido por el folio ${d.sustituidoPor}.`; nota.classList.remove('oculto'); }
        else if (d.estado === 'cancelado' && d.motivo) { nota.textContent = `Motivo: ${d.motivo}`; nota.classList.remove('oculto'); }
        if (d.publicadoEl) { const p = document.createElement('p'); p.className = 'nota'; p.textContent = `Última publicación: ${new Date(d.publicadoEl).toLocaleString('es-MX', { timeZone: 'America/Mexico_City', dateStyle: 'long', timeStyle: 'short' })}`; nota.insertAdjacentElement('afterend', p); }
    }
    async function correr() {
        const f = new URLSearchParams(location.search).get('f') || '';
        if (!RE_F.test(f)) { estado('mal', 'El enlace no trae un folio válido. Escanee el QR del certificado; no teclee el folio a mano.'); return; }
        try {
            const r = await fetch(`./datos/${encodeURIComponent(f)}.json`, { cache: 'no-store' });
            if (r.status === 404) { estado('mal', `No hay ningún certificado publicado con el folio ${f.slice(0, 10)}. Puede ser un papel falso, un certificado recién emitido que aún no se publica, o un QR dañado.`); return; }
            if (!r.ok) throw new Error('HTTP ' + r.status);
            pintar(await r.json());
        } catch (e) { estado('mal', 'No se pudo consultar la verificación (' + e.message + '). Intente de nuevo en un momento.'); }
    }
    correr();
})();
