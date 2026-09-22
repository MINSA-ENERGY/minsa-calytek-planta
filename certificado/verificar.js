// Verificacion publica del certificado de tratamiento (v0.36.0). Sin dependencias, sin login, sin Graph.
//
// El repo de la app es PUBLICO, asi que lo publicado NO puede ser legible: cada certificado vive en
// ./datos/<sha256(folio-sufijo)>.json CIFRADO (AES-256-CBC + HMAC-SHA256, llaves derivadas del sufijo del QR con
// PBKDF2-SHA256, 100 000 iteraciones, sal «MINSA-CT:<folio>»). Quien escaneo el QR trae el sufijo y lee; quien lista el
// repo ve blobs y ni siquiera los folios. Lo escribe docs/exportar-planta.ps1 -PublicarCertificados con el mismo esquema.
// Todo va por textContent: nada de lo que llega del JSON se interpreta como HTML.
(function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const RE_F = /^(CT-\d{2}-\d{4})-([a-z2-9]{12})$/;   // folio + sufijo, tal cual lo pinta la app (reglas.sufijoVerificacion)
    const ITERACIONES = 100000;
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
    const enc = new TextEncoder();
    const hex = buf => [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
    const deB64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
    function estado(clase, texto) { const e = $('estado'); e.className = 'estado ' + clase; e.textContent = texto; }

    /** Llaves AES (32 B) y HMAC (32 B) a partir del sufijo; la sal lleva el folio para que dos certificados nunca compartan llave. */
    async function llaves(folio, sufijo) {
        const base = await crypto.subtle.importKey('raw', enc.encode(sufijo), 'PBKDF2', false, ['deriveBits']);
        const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: enc.encode('MINSA-CT:' + folio), iterations: ITERACIONES, hash: 'SHA-256' }, base, 512);
        const aes = await crypto.subtle.importKey('raw', bits.slice(0, 32), { name: 'AES-CBC' }, false, ['decrypt']);
        const mac = await crypto.subtle.importKey('raw', bits.slice(32, 64), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
        return { aes, mac };
    }
    async function descifrar(blob, folio, sufijo) {
        if (!blob || blob.v !== 1 || !blob.iv || !blob.ct || !blob.mac) throw new Error('formato desconocido');
        const iv = deB64(blob.iv), ct = deB64(blob.ct), mac = deB64(blob.mac);
        const k = await llaves(folio, sufijo);
        const juntos = new Uint8Array(iv.length + ct.length); juntos.set(iv); juntos.set(ct, iv.length);
        if (!(await crypto.subtle.verify('HMAC', k.mac, mac, juntos))) throw new Error('el sello no coincide: QR dañado o archivo alterado');
        const claro = await crypto.subtle.decrypt({ name: 'AES-CBC', iv }, k.aes, ct);
        return JSON.parse(new TextDecoder().decode(claro));
    }
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
        const m = RE_F.exec(f);
        if (!m) { estado('mal', 'El enlace no trae un folio válido. Escanee el QR del certificado; no teclee el folio a mano.'); return; }
        if (!(crypto && crypto.subtle)) { estado('mal', 'Este navegador no puede verificar (sin WebCrypto). Abra el enlace en un navegador actual.'); return; }
        const [, folio, sufijo] = m;
        try {
            const nombre = hex(await crypto.subtle.digest('SHA-256', enc.encode(f)));
            const r = await fetch(`./datos/${nombre}.json`, { cache: 'no-store' });
            if (r.status === 404) { estado('mal', `No hay ningún certificado publicado con el folio ${folio}. Puede ser un papel falso, un certificado recién emitido que aún no se publica, o un QR dañado.`); return; }
            if (!r.ok) throw new Error('HTTP ' + r.status);
            pintar(await descifrar(await r.json(), folio, sufijo));
        } catch (e) { estado('mal', 'No se pudo verificar (' + e.message + '). Intente de nuevo; si persiste, el QR o el archivo publicado no corresponden.'); }
    }
    correr();
})();
