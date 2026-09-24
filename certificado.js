// CALYTEK Planta — el certificado de tratamiento.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { aIsoDia, datosCertificado, fechaCorta, fechaMexico, lista, plural, PUEDE, registroCertificado, residuoDe, siguienteFolio, sufijoVerificacion, toneladas, urlVerificacion } from './reglas.js';
import { $, anclar, aplicar, avisar, confirmar, el, escribiendo, estado, etiqueta, firmaDe, firmar, iso, L, limpiarAvisos, motivoSinFirmas, porId, quien, refrescarCliente, renglon, textoDe, VERSION } from './nucleo.js';
import { asegurarFolioUnico } from './puerta.js';
import { pintarListasGondolas } from './gondolas.js';
import { segundoPaso } from './padron.js';

// ================================================================ CERTIFICADO DE TRATAMIENTO (v0.39.0)
//
// Decision de Carlos (2026-09-22, noche): UN certificado por GONDOLA — no por programa, como fue de la v0.35.0 a la
// v0.38.0. Un renglon de PLANTA_Embarques cerrado YA ES una gondola: un arribo de una unidad, con su manifiesto, su
// ticket de bascula y su neto. El disparador de emision es el cierre del EMBARQUE, no el del programa: en cuanto la
// gondola tiene neto y ticket, gerencia puede emitir y el generador no espera a que cierre el programa.
// Lo emite SOLO gerencia, se firma en PLANTA_Firmas (Tipo certificado, S-11: sin firma valida la app no lo imprime) y
// lleva QR a una pagina publica sin login (certificado/ de este repo) que muestra solo lo impreso. El papel se CONGELA
// al emitir (kg, manifiesto, ticket, generador…): si el embarque cambia despues, el certificado no se edita — se
// SUSTITUYE (nace otro; el viejo queda «sustituido») o se CANCELA. La pagina publica se regenera con
// docs/exportar-planta.ps1 -PublicarCertificados.  Memoria: calytek-certificado-tratamiento.

/** Dia (hora de Mexico) de un ISO datetime, como dd/mm/aaaa. aIsoDia NO sirve aqui: espera una fecha tecleada y lanza con un ISO. */
const diaCert = iso => (iso ? fechaCorta(fechaMexico(new Date(iso))) : '—');
/**
 * Los certificados de UNA gondola, del mas nuevo al mas viejo (una sustitucion deja el viejo aqui). No se muta lo que devuelve.
 * C-45 (v0.44.0): de un indice por EmbarqueId que se rehace solo cuando la lista cambia (carga nueva o push); antes cada
 * renglon de Cerrados filtraba y ordenaba la lista entera.
 */
let indiceCertificados = { lista: null, n: -1, porEmbarque: new Map() };
export function certificadosDe(e) {
    const lista = estado.certificados;
    if (indiceCertificados.lista !== lista || indiceCertificados.n !== lista.length) {
        const porEmbarque = new Map();
        for (const c of lista) { const k = Number(c.EmbarqueId); if (!porEmbarque.has(k)) porEmbarque.set(k, []); porEmbarque.get(k).push(c); }
        for (const v of porEmbarque.values()) v.sort((a, b) => b.id - a.id);
        indiceCertificados = { lista, n: lista.length, porEmbarque };
    }
    return indiceCertificados.porEmbarque.get(Number(e.id)) || [];
}
export function certificadoVigente(e) { return certificadosDe(e).find(c => c.Estado === 'vigente') || null; }
/** Los del PROGRAMA: solo para leerlos en el detalle de la pre-alta. No se emite nada desde ahi. C-45: los de sus gondolas cargadas. */
function certificadosDePrograma(p) { return estado.certificados.filter(c => Number(c.PreAltaId) === Number(p.id)).sort((a, b) => a.id - b.id); }
/** S-11 sobre el certificado: vale solo con un renglon de PLANTA_Firmas de Tipo certificado, firmado por gerencia y por la misma cuenta que EmitidoPor. */
function certificadoFirmado(c) { return !!firmaDe('certificado', c.id, c.EmitidoPor); }
const motivoSinCertificados = () => (estado.certificadosError ? `No se pudo leer la lista de certificados: no se emite hasta que se vea.${estado.rol === 'gerencia' ? ` (PLANTA_Certificados: ${estado.certificadosError} — si no existe, se crea con herramientas-dev/provisionar.html)` : ' Avisa a gerencia.'}` : null);

/**
 * v0.39.0: el detalle del programa ya no emite nada — el certificado es de la gondola. Aqui solo se LEE cuantos
 * lleva el programa y cuanto suman, para que gerencia no tenga que recorrer Cerrados para saberlo.
 */
export function pintarCertificadoEnDetalle(p) {
    const ul = $('paDetalleCertificado'); ul.textContent = '';
    if (p.Estado === 'borrador') return;
    const certs = certificadosDePrograma(p);
    // U-79 (v0.42.0): ámbar solo si hay góndolas cerradas del programa sin certificado vigente; sin ellas no falta nada.
    const pendientes = estado.embarques.filter(e => Number(e.PreAltaId) === Number(p.id) && e.Etapa === 'cerrado' && !certificadoVigente(e)).length;
    const faltan = pendientes ? ` · ${plural(pendientes, 'góndola cerrada', 'góndolas cerradas')} sin certificado` : '';
    if (!certs.length) {
        const li = el('li', '', 'Certificados de tratamiento ');
        li.appendChild(etiqueta('ninguno', pendientes ? 'aviso' : ''));
        li.appendChild(el('span', 'd', `Se emiten por góndola cerrada, desde Góndolas › Cerradas hoy o Historial${faltan}.`));
        ul.appendChild(li); return;
    }
    const vigentes = certs.filter(c => c.Estado === 'vigente');
    const kg = vigentes.reduce((s, c) => s + (Number(c.Kg) || 0), 0);
    const li = el('li', '', `Certificados de tratamiento: ${certs.length} `);
    li.appendChild(etiqueta(plural(vigentes.length, 'vigente'), pendientes ? 'aviso' : vigentes.length ? 'ok' : ''));
    li.appendChild(el('span', 'd', `${toneladas(kg)} t certificadas · ${certs.map(c => c.Title).join(', ')}${faltan} · uno por góndola: se emiten y se ven desde Góndolas › Cerradas hoy o Historial.`));
    ul.appendChild(li);
}

/**
 * U-72 (v0.45.0): lo que el papel CONGELA al emitir, en un solo sitio: la emision lo escribe y el borrador lo pinta con la
 * misma funcion, asi lo que gerencia ve antes de emitir es lo que sale. `gen` = generadorDe(p) o lo corregido en Sustituir.
 */
function camposPapel(p, e, d, gen) {
    const carrier = porId(estado.carriers, e.CarrierId || (p && p.CarrierId));
    return {
        Kg: d.kg, Manifiesto: d.manifiesto, TicketBascula: d.ticket, FechaRecepcion: d.fechaRecepcion, TipoBulto: d.tipoBulto,
        ...gen, Corriente: (p && p.Corriente) || null,
        Transportista: carrier ? `${carrier.Title} · autorización ${carrier.AutorizacionASEA || 'sin número'}` : null
    };
}
const generadorDe = p => ({ Generador: (p && p.Generador) || null, GeneradorRegistro: (p && p.GeneradorRegistro) || null, GeneradorDireccion: (p && p.GeneradorDireccion) || null, Pozo: (p && p.Pozo) || null });
/** U-72 (v0.45.0): los renglones del papel que saldrian «—». */
function vaciosDelPapel(c) {
    return [['ticket de báscula', c.TicketBascula], ['manifiesto', manifiestoCert(c)], ['generador', c.Generador], ['dirección', c.GeneradorDireccion],
        ['registro de generador', c.GeneradorRegistro], ['pozo', c.Pozo], ['transportista', c.Transportista]].filter(([, v]) => !v).map(([n]) => n);
}
/**
 * U-75 (v0.45.0): los datos del papel como lista de texto (los mismos renglones que la verificacion publica). En el celular
 * el papel a escala no se lee: la lista va arriba y el papel queda como vista previa que se amplia al tocarla. En
 * escritorio la lista se esconde por CSS: ahi el papel se lee entero.
 */
function pintarResumenCert(c) {
    const dl = $('ctResumen'); dl.textContent = '';
    dl.classList.toggle('oculto', !c);
    if (!c) return;
    const filas = [['Folio', c.Estado === 'borrador' ? 'se asigna al emitir' : c.Title], ['Estado', c.Estado === 'borrador' ? 'BORRADOR · sin emitir' : c.Estado],
        ['Generador', c.Generador], ['Dirección', c.GeneradorDireccion], ['Registro de generador', c.GeneradorRegistro], ['Residuo', residuoDe(c.Corriente)],
        ['Volumen', `${toneladas(c.Kg)} t`], ['Ticket de báscula', c.TicketBascula], ['Fecha de recepción', fechaRecepcionCert(c)],
        ['Manifiesto', manifiestoCert(c)], ['Pozo', c.Pozo], ['Transportista', c.Transportista]];
    for (const [eti, v] of filas) { dl.appendChild(el('dt', '', eti)); dl.appendChild(el('dd', v ? '' : 'vacio', v || '—')); }
}

/** Abre el certificado de UNA gondola (o el hueco donde iria). Es la unica puerta de emision desde la v0.39.0. */
export function abrirCertificadoDeEmbarque(e) {
    estado.certificadoEmbarque = e;
    limpiarAvisos(); mostrarCorreccion(false);
    pintarCertificadoDeEmbarque();
    const d = $('dlgCertificado'); if (!d.open) d.showModal();
}
/** Pinta el dialogo del certificado de `estado.certificadoEmbarque`: el papel (o el hueco) y que botones proceden. */
function pintarCertificadoDeEmbarque() {
    const e = estado.certificadoEmbarque; if (!e) return;
    const certs = certificadosDe(e);
    const vig = certificadoVigente(e);
    const cert = vig || certs[0] || null;
    estado.certificadoAbierto = cert;
    const t = $('certificado');
    const p = porId(estado.prealtas, e.PreAltaId);
    const d = datosCertificado(p, e);
    const puede = PUEDE.emitirCertificado(estado.rol);
    const bloqueo = motivoSinCertificados() || motivoSinFirmas() || (d.ok ? '' : `No se puede certificar: ${d.motivo}.`);
    // U-72 (v0.45.0): sin certificado, gerencia ve el papel como BORRADOR antes de emitir (y que renglones saldrian vacios).
    const borrador = !cert && puede && d.ok ? { Title: 'POR ASIGNAR', Estado: 'borrador', ...camposPapel(p, e, d, generadorDe(p)) } : null;
    if (cert || borrador) pintarCertificado(cert || borrador, t);
    else {
        t.textContent = ''; t.classList.remove('sustituido', 'cancelado', 'borrador');
        t.appendChild(el('p', 'pista', `La góndola ${e.Title || '(sin folio)'} todavía no tiene certificado de tratamiento. ${puede ? 'Solo gerencia lo emite; queda registrado quién y cuándo, y el papel lleva QR de verificación.' : 'Lo emite gerencia.'}`));
    }
    pintarResumenCert(cert || borrador);
    const vacios = borrador ? vaciosDelPapel(borrador) : [];
    $('btnEmitirCertificado').classList.toggle('oculto', !(puede && !vig));
    $('btnEmitirCertificado').disabled = !!bloqueo; $('btnEmitirCertificado').title = bloqueo;
    $('btnSustituirCertificado').classList.toggle('oculto', !(puede && vig));
    $('btnSustituirCertificado').disabled = !!bloqueo; $('btnSustituirCertificado').title = bloqueo;
    $('btnCancelarCertificado').classList.toggle('oculto', !(puede && vig));
    const imprimible = !!cert && cert.Estado === 'vigente' && certificadoFirmado(cert);
    $('ctImprimir').classList.toggle('oculto', !cert);
    $('ctImprimir').disabled = !imprimible;
    $('ctImprimir').title = !cert ? 'Todavía no hay certificado.' : imprimible ? '' : (cert.Estado !== 'vigente' ? `Este certificado está ${cert.Estado}: no se imprime como vigente.` : 'Sin firma registrada de gerencia: no se imprime.');
    // U-123 (v0.73.0): la identificación en un renglón y los avisos que piden actuar, cada uno en el suyo (antes, un solo párrafo
    // de hasta siete cláusulas). Todo sigue dentro de #ctEstado.
    const ident = (cert
        ? `${cert.Title} · ${imprimible ? 'vigente y firmado' : cert.Estado + (cert.Estado === 'vigente' ? ' · SIN FIRMA' : '')}`
        : `Góndola ${e.Title || '(sin folio)'} · sin certificado${borrador ? ' · abajo, el BORRADOR: así saldría el papel' : ''}`)
        + ` · góndola ${e.Title || '—'}${e.Manifiesto ? ' · manifiesto ' + e.Manifiesto : ''}`
        + (certs.length > 1 ? ` · ${certs.length} emitidos para esta góndola` : '');
    const avisosCert = [
        vig && e.Etapa === 'anulado' ? '⚠ La góndola está anulada: cancela este certificado.' : '',   // U-70 (v0.40.0)
        certs.filter(c => c.Estado === 'vigente').length > 1 ? '⚠ Hay más de uno vigente: vale el más nuevo; cancela los otros.' : '',   // C-39 (v0.40.0)
        vacios.length ? `⚠ Si se emite, saldrían vacíos: ${vacios.join(', ')}.` : '',   // U-72 (v0.45.0)
        puede && bloqueo ? bloqueo : ''   // U-73 (v0.42.0): en el celular no hay title; el motivo se lee aquí
    ].filter(Boolean);
    const ce = $('ctEstado'); ce.textContent = '';
    ce.appendChild(el('span', 'ct-id', ident));
    if (avisosCert.length) { const ul = el('ul', 'ct-avisos'); for (const a of avisosCert) ul.appendChild(el('li', '', a)); ce.appendChild(ul); }
}
/** U-69 (v0.40.0): el panel de correccion de Sustituir, precargado con lo que la gondola y su programa dicen HOY. */
export function mostrarCorreccion(si) {
    $('ctCorregir').classList.toggle('oculto', !si); $('ctBotones').classList.toggle('oculto', si);
    if (!si) return;
    const e = estado.certificadoEmbarque; const p = e && porId(estado.prealtas, e.PreAltaId);
    $('ctTicket').value = textoDe(e && e.TicketBascula); $('ctManifiesto').value = textoDe(e && e.Manifiesto);
    $('ctGenerador').value = textoDe(p && p.Generador); $('ctGeneradorDireccion').value = textoDe(p && p.GeneradorDireccion);
    $('ctGeneradorRegistro').value = textoDe(p && p.GeneradorRegistro); $('ctPozo').value = textoDe(p && p.Pozo); $('ctMotivo').value = '';
    $('ctTicket').focus();
}
/** Los campos de una cancelacion de certificado (C-50: uno solo para cancelar, compensar una emision y anular la gondola). */
export const camposCancelacion = motivo => ({ Estado: 'cancelado', CanceladoPor: estado.cuenta.username, CanceladoEl: new Date().toISOString(), Motivo: String(motivo).slice(0, 255) });

/**
 * Emite el certificado de la gondola abierta en el dialogo. Con `sustituye`, es una SUSTITUCION: el nuevo nace y el
 * viejo queda «sustituido» con SustituidoPor. El embarque se RELEE en vivo antes de congelar nada (gerencia pudo
 * anularlo, o el basculista corregir el neto, mientras el dialogo estaba abierto). Orden de escritura: renglon del
 * certificado -> folio unico -> FIRMA (403 si no eres gerencia) -> el viejo, si sustituye.
 * C-39 (v0.40.0): CUALQUIER falla despues de crear el renglon lo cancela —no solo la de la firma—, asi un fallo a medias
 * deja lo que habia antes (el viejo vigente, o nada) y el aviso nombra el paso. Si la compensacion tambien falla, el
 * aviso lo dice con el folio, y el dialogo marca «mas de uno vigente».
 * C-38 (v0.40.0): la lista de certificados se relee y MANDA: si otro toque u otra sesion ya emitio, no nace un segundo.
 * U-69 (v0.40.0): `corr` son los datos del panel de Sustituir. Ticket y manifiesto se corrigen en la GONDOLA (son su
 * dato) antes de congelar; generador, registro, direccion y pozo entran solo al certificado — el programa firmado no se toca.
 */
export async function emitirCertificado(sustituye = null, motivoSust = '', corr = null) {
    const e = estado.certificadoEmbarque; if (!e || !PUEDE.emitirCertificado(estado.rol)) return;
    const btn = sustituye ? 'ctCorregirOk' : 'btnEmitirCertificado';
    await escribiendo(btn, async () => {
    try { await refrescarCliente(); Object.assign(e, await estado.cliente.renglon(estado.siteId, L.embarques, e.id)); anclar('embarques', e); }
    catch (err) { avisar('No pude releer la góndola (' + err.message + '). No se emite.', 'error'); return; }
    if (corr) {
        const cambios = {};
        if (corr.TicketBascula !== textoDe(e.TicketBascula)) cambios.TicketBascula = corr.TicketBascula || null;
        if (corr.Manifiesto !== textoDe(e.Manifiesto)) cambios.Manifiesto = corr.Manifiesto || null;
        if (Object.keys(cambios).length) {
            try { await estado.cliente.actualizarRenglon(estado.siteId, L.embarques, e.id, cambios); aplicar('embarques', e, cambios); }
            catch (err) { avisar('No pude corregir la góndola (' + err.message + '). No se sustituye.', 'error'); return; }
        }
    }
    const p = porId(estado.prealtas, e.PreAltaId);
    const d = datosCertificado(p, e);
    if (!d.ok) { avisar('No se puede certificar: ' + d.motivo + '.', 'error'); pintarCertificadoDeEmbarque(); return; }
    const gen = generadorDe(p);
    if (corr) for (const k of Object.keys(gen)) gen[k] = corr[k] || null;
    const papel = camposPapel(p, e, d, gen);
    // Una gondola cerrada ANTES de la v0.39.0 no trae ticket de bascula: el papel sale sin ese dato y se avisa.
    if (!d.ticket) {
        const { ok } = await confirmar({ titulo: 'Sin ticket de báscula', ok: 'Emitir sin ticket',
            texto: `La góndola ${d.folio} no trae número de ticket de báscula. El certificado saldrá con ese renglón vacío.` });
        if (!ok) return;
    }
    if (!sustituye) {
        // U-72 (v0.45.0): el ticket ya tuvo su propio aviso arriba; el resto de los renglones vacios se nombra aqui.
        const vacios = vaciosDelPapel(papel).filter(n => n !== 'ticket de báscula');
        const { ok } = await confirmar({ titulo: 'Emitir el certificado de tratamiento', ok: 'Emitir',
            texto: `Góndola ${d.folio}${d.manifiesto ? `, manifiesto ${d.manifiesto}` : ''}: ${toneladas(d.kg)} t netas (${d.kg.toLocaleString('es-MX')} kg), recibida el ${diaCert(d.fechaRecepcion)}. Con tu firma queda emitido a nombre de ${gen.Generador || '(sin generador)'}; lo que diga el papel no cambia después: si algo está mal, se sustituye.${vacios.length ? ` OJO: saldrán vacíos («—») en el papel: ${vacios.join(', ')}.` : ''}` });
        if (!ok) return;
    }
    let nuevo = null, paso = '';
    try {
        // C-45 (v0.44.0): se relee solo lo de ESTA gondola (EmbarqueId indexada, una peticion), no la lista entera. El folio sale
        // de lo cargado mas eso; si otra sesion tomo el mismo numero, asegurarFolioUnico lo detecta y renumbera (C-16).
        const deLaGondola = await estado.cliente.renglones(estado.siteId, L.certificados, `fields/EmbarqueId eq ${Number(e.id)}`);
        const otro = deLaGondola.find(x => Number(x.EmbarqueId) === Number(e.id) && x.Estado === 'vigente' && (!sustituye || x.id !== sustituye.id));
        if (otro) throw new Error(`la góndola ya tiene el certificado ${otro.Title} vigente (${quien(otro.EmitidoPor)}). Actualiza antes de emitir otro`);
        if (sustituye && !deLaGondola.some(x => x.id === sustituye.id && x.Estado === 'vigente')) throw new Error(`${sustituye.Title} ya no está vigente. Actualiza`);
        const folio = siguienteFolio('C', [...estado.certificados, ...deLaGondola].map(x => x.Title));
        const ahora = new Date().toISOString();
        nuevo = await estado.cliente.crearRenglon(estado.siteId, L.certificados, registroCertificado({   // C-77 (v0.74.0)
            folio, prealtaId: p.id, embarqueId: e.id, papel, sufijo: sufijoVerificacion(), usuario: estado.cuenta.username, ahora, version: VERSION, sustituye, motivoSust }));
        try {
            paso = 'asegurar el folio'; await asegurarFolioUnico(nuevo, 'C');
            paso = 'firmar'; await firmar('certificado', nuevo);
            if (sustituye) {
                paso = `marcar ${sustituye.Title} como sustituido`;
                const campos = { Estado: 'sustituido', SustituidoPor: nuevo.Title, Motivo: String(motivoSust).slice(0, 255) };
                await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, sustituye.id, campos);
                aplicar('certificados', sustituye, campos);
            }
        } catch (err) {
            const cancel = camposCancelacion(`Emisión fallida al ${paso}`);   // S-22 (v0.41.0): el motivo sale en la verificación pública; el error de Graph solo en el aviso
            const sinCompensar = await segundoPaso(() => estado.cliente.actualizarRenglon(estado.siteId, L.certificados, nuevo.id, cancel));
            if (!sinCompensar) Object.assign(nuevo, cancel);
            estado.certificados.push(nuevo);
            throw new Error(`falló al ${paso} (${err.message}). ` + (sinCompensar
                ? `OJO: ${nuevo.Title} quedó creado y no se pudo cancelar (${sinCompensar}): cancélalo desde aquí.`
                : `${nuevo.Title} quedó cancelado${sustituye ? ` y ${sustituye.Title} sigue vigente` : ''}.`));
        }
        estado.certificados.push(nuevo);
        if (corr) mostrarCorreccion(false);
        avisar(`Certificado ${nuevo.Title} emitido${sustituye ? ` en sustitución de ${sustituye.Title}` : ''}. El QR contesta después de la próxima publicación (cada hora, al minuto 7).`, 'bien');   // U-77 (v0.42.0)
    } catch (err) { avisar('No se pudo emitir: ' + err.message, 'error'); }
    repintarTrasCertificado();
    });
}
/** C-50 (v0.42.0): lo que se repinta tras emitir, sustituir o cancelar: el diálogo, Cerrados y, si está abierto, el detalle del programa. */
function repintarTrasCertificado() {
    pintarCertificadoDeEmbarque();
    if (estado.pestana === 'bascula') pintarListasGondolas();   // tanda 3: las pestañas, sin tocar un pesaje abierto
    if (estado.prealtaAbierta) pintarCertificadoEnDetalle(estado.prealtaAbierta);
}
/** U-69 (v0.40.0): Sustituir ya no es un confirm con motivo: abre el panel de correccion precargado. */
export function sustituirCertificado() {
    const e = estado.certificadoEmbarque; const vig = e && certificadoVigente(e); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    limpiarAvisos(); mostrarCorreccion(true);
}
export async function confirmarSustitucion() {
    const e = estado.certificadoEmbarque; const vig = e && certificadoVigente(e); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    const motivo = $('ctMotivo').value.trim();
    if (!motivo) { avisar('Escribe qué estaba mal: queda registrado en los dos certificados.', 'error'); $('ctMotivo').focus(); return; }
    const corr = { TicketBascula: $('ctTicket').value.trim(), Manifiesto: $('ctManifiesto').value.trim(), Generador: $('ctGenerador').value.trim(),
        GeneradorDireccion: $('ctGeneradorDireccion').value.trim(), GeneradorRegistro: $('ctGeneradorRegistro').value.trim(), Pozo: $('ctPozo').value.trim() };
    const tocaGondola = corr.TicketBascula !== textoDe(e.TicketBascula) || corr.Manifiesto !== textoDe(e.Manifiesto);
    // U-72 (v0.45.1, revisor): Sustituir tambien emite — su confirm nombra lo que saldria vacio con los datos del panel,
    // armado con la misma camposPapel que congela. El ticket se queda fuera: emitirCertificado le da su propio aviso.
    const p = porId(estado.prealtas, e.PreAltaId);
    const gen = generadorDe(p); for (const k of Object.keys(gen)) gen[k] = corr[k] || null;
    const vacios = vaciosDelPapel(camposPapel(p, e, datosCertificado(p, { ...e, TicketBascula: corr.TicketBascula || null, Manifiesto: corr.Manifiesto || null }), gen)).filter(n => n !== 'ticket de báscula');
    const { ok } = await confirmar({ titulo: 'Sustituir el certificado', ok: 'Sustituir',
        texto: `${vig.Title} queda como «sustituido» (su QR lo dirá) y nace uno nuevo con los datos del panel${tocaGondola ? '; el ticket y el manifiesto se corrigen también en la góndola' : ''}. «${motivo}» queda registrado en los dos.${vacios.length ? ` OJO: saldrán vacíos («—») en el papel: ${vacios.join(', ')}.` : ''}` });
    if (!ok) return;
    await emitirCertificado(vig, motivo, corr);
}
export async function cancelarCertificado() {
    const e = estado.certificadoEmbarque; const vig = e && certificadoVigente(e); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    await escribiendo('btnCancelarCertificado', async () => {
    const { ok, motivo } = await confirmar({ titulo: 'Cancelar el certificado', ok: 'Cancelar el certificado', peligro: true, motivo: true, etiquetaMotivo: 'Por qué se cancela (sale en la verificación pública del QR)',
        texto: `${vig.Title} deja de valer: su QR dirá «cancelado». No se borra: queda como historial. Si hace falta uno bueno, después se emite otro.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        const campos = camposCancelacion(motivo);
        await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, vig.id, campos);
        aplicar('certificados', vig, campos);
        avisar(`Certificado ${vig.Title} cancelado.`, 'bien');
    } catch (err) { avisar('No se pudo cancelar: ' + err.message, 'error'); }
    repintarTrasCertificado();
    });
}
/** QR como SVG en el DOM (sin innerHTML): modulos del generador vendorizado (qrcode-generator, MIT). Sin la libreria, un recuadro que lo dice. */
function qrSvg(texto, lado) {
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('class', 'qr'); svg.setAttribute('width', lado); svg.setAttribute('height', lado); svg.setAttribute('role', 'img');
    if (typeof window.qrcode !== 'function') { svg.setAttribute('viewBox', '0 0 10 10'); const r = document.createElementNS(NS, 'rect'); r.setAttribute('width', 10); r.setAttribute('height', 10); r.setAttribute('fill', 'none'); r.setAttribute('stroke', '#000'); svg.appendChild(r); svg.dataset.sinLibreria = '1'; return svg; }
    const q = window.qrcode(0, 'M'); q.addData(texto); q.make();
    const n = q.getModuleCount();
    svg.setAttribute('viewBox', `0 0 ${n} ${n}`); svg.setAttribute('shape-rendering', 'crispEdges');
    const fondo = document.createElementNS(NS, 'rect'); fondo.setAttribute('width', n); fondo.setAttribute('height', n); fondo.setAttribute('fill', '#fff'); svg.appendChild(fondo);
    let d = '';
    for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (q.isDark(r, c)) d += `M${c} ${r}h1v1h-1z`;
    const path = document.createElementNS(NS, 'path'); path.setAttribute('d', d); path.setAttribute('fill', '#131c2c'); svg.appendChild(path);
    const t = document.createElementNS(NS, 'title'); t.textContent = texto; svg.appendChild(t);
    return svg;
}
/** El manifiesto del papel. Un certificado de la v0.35-v0.38 (uno por programa) traia varios en `Manifiestos`. */
function manifiestoCert(c) { return c.Manifiesto || lista(c.Manifiestos).join(' · ') || null; }
/** La fecha de recepcion del papel: un dia. Los certificados viejos traian un rango (PrimerCierre/UltimoCierre). */
function fechaRecepcionCert(c) {
    if (c.FechaRecepcion) return diaCert(c.FechaRecepcion);
    if (!c.PrimerCierre && !c.UltimoCierre) return null;
    return c.PrimerCierre && c.UltimoCierre && fechaMexico(new Date(c.PrimerCierre)) !== fechaMexico(new Date(c.UltimoCierre))
        ? `del ${diaCert(c.PrimerCierre)} al ${diaCert(c.UltimoCierre)}` : diaCert(c.UltimoCierre || c.PrimerCierre);
}
/**
 * La tabla de incisos, un renglon por bulto. Hoy NO se pinta (ver pintarCertificado): espera al contenedor de marina.
 * C-49 (v0.42.0): lee Embarques/Manifiestos, que ninguna emision escribe hoy; al encenderla hay que escribirlos tambien.
 * Se deja tal como estaba en la v0.38.0 a proposito — apagar no es rediseniar: el rotulo de la segunda columna lo
 * decide Carlos el dia que se encienda, con un caso marino real delante (hoy la celda trae el folio del embarque).
 */
function tablaIncisosCert(c) {
    const folios = lista(c.Embarques), manif = lista(c.Manifiestos);
    const tabla = el('table', 'ct-incisos'); const th = el('tr'); for (const h of ['#', 'FOLIO DE EMBARQUE', 'MANIFIESTO', 'NETO (KG)']) th.appendChild(el('th', '', h)); tabla.appendChild(th);
    folios.forEach((f, i) => { const tr = el('tr'); tr.appendChild(el('td', 'n', String(i + 1))); tr.appendChild(el('td', 'mono', f)); tr.appendChild(el('td', 'mono', manif[i] || '—')); tr.appendChild(el('td', 'mono kg', '')); tabla.appendChild(tr); });
    // El neto por bulto no se congela en el renglon (solo el total, Kg): se toma del embarque cargado si esta en la ventana; si no, en blanco.
    [...tabla.querySelectorAll('td.kg')].forEach((td, i) => { const e = estado.embarques.find(x => x.Title === folios[i]); td.textContent = e && e.NetoKg ? Number(e.NetoKg).toLocaleString('es-MX') : ''; });
    return tabla;
}

/** Pinta el formato FO-CT-01 (carta horizontal) dentro de `t`. Solo lee el renglon del certificado: es lo que se congelo al emitir. */
function pintarCertificado(c, t) {
    const CFG = CONFIG.certificado;
    t.textContent = '';
    t.classList.toggle('sustituido', c.Estado === 'sustituido'); t.classList.toggle('cancelado', c.Estado === 'cancelado'); t.classList.toggle('borrador', c.Estado === 'borrador');
    const marco = el('div', 'ct-marco'); const hoja = el('div', 'ct-hoja'); marco.appendChild(hoja); t.appendChild(marco);
    // v0.37.0: la marca de agua es el SIMBOLO solo, no el lockup (Carlos, 23-sep); gris horneado y 4.3 in contra el lado corto, como las plantillas de la casa (marca_agua.py)
    const agua = el('img', 'ct-agua'); agua.src = './marca/simbolo-agua.svg'; agua.alt = ''; hoja.appendChild(agua);
    if (c.Estado !== 'vigente') hoja.appendChild(el('div', 'ct-sello', c.Estado === 'cancelado' ? 'CANCELADO' : c.Estado === 'borrador' ? 'BORRADOR · SIN EMITIR' : `SUSTITUIDO POR ${c.SustituidoPor || '—'}`));
    const folioCaja = el('div', 'ct-folio'); folioCaja.appendChild(el('div', 'ct-eti', 'FOLIO')); folioCaja.appendChild(el('div', 'ct-folio-num', c.Title)); hoja.appendChild(folioCaja);
    const logo = el('img', 'ct-logo'); logo.src = './marca/lockup.svg'; logo.alt = 'MINSA ENERGY'; hoja.appendChild(logo);
    hoja.appendChild(el('h1', 'ct-titulo', 'CERTIFICADO DE TRATAMIENTO'));
    const sub = el('div', 'ct-sub'); sub.appendChild(el('i')); sub.appendChild(el('span', '', 'RECORTES DE PERFORACIÓN · RESIDUO DE MANEJO ESPECIAL')); sub.appendChild(el('i')); hoja.appendChild(sub);
    const inciso = (eti, valor, clase = '') => { const d = el('div', 'ct-inciso'); d.appendChild(el('span', 'ct-eti', eti)); d.appendChild(el('span', 'ct-val ' + clase, valor || '—')); return d; };
    const fila = (...incisos) => { const f = el('div', 'ct-fila'); for (const i of incisos) f.appendChild(i); return f; };
    const filaCentrada = (...incisos) => { const f = fila(...incisos); f.classList.add('centrada'); return f; };
    // v0.39.0 (Carlos, 22-sep noche): arriba queda SOLO el generador — POZO baja al bloque del manifiesto y REGISTRO
    // DE GENERADOR pasa a renglon completo. Sin rotulos de bloque visibles: la separacion es por acomodo.
    hoja.appendChild(fila(inciso('GENERADOR:', c.Generador, 'fuerte')));
    hoja.appendChild(fila(inciso('DIRECCIÓN:', c.GeneradorDireccion)));   // v0.38.0 (Carlos): renglon nuevo bajo GENERADOR; los certificados anteriores no lo traen y sale «—»
    hoja.appendChild(fila(inciso('REGISTRO DE GENERADOR:', c.GeneradorRegistro, 'mono')));
    const prosa = el('p', 'ct-prosa'); prosa.appendChild(el('b', '', 'MATERIAS INDUSTRIALIZADAS CCMV DEL NORTE, S.A. DE C.V.')); prosa.appendChild(document.createTextNode(' — MINSA ENERGY — certifica que ha recibido para su tratamiento en su planta '));
    prosa.appendChild(el('b', '', 'CALYTEK')); prosa.appendChild(document.createTextNode(', de una forma ambientalmente segura y conforme a los términos de su autorización, el residuo de:')); hoja.appendChild(prosa);
    hoja.appendChild(el('div', 'ct-residuo', residuoDe(c.Corriente)));
    // v0.38.0 (Carlos, 23-sep): fuera la regla azul bajo el residuo y mas aire debajo; VOLUMEN/…/FECHA centrados; firma al centro; leyenda del QR corta; pie sin la linea «Emitido...».
    // v0.39.0 (Carlos, 22-sep noche): donde iba EMBARQUES ahora va el TICKET DE BASCULA de la gondola, y el MANIFIESTO
    // —uno solo— baja con el POZO al renglon de abajo.
    hoja.appendChild(filaCentrada(inciso('VOLUMEN:', `${toneladas(c.Kg)} TON.`, 'volumen'), inciso('TICKET DE BÁSCULA:', c.TicketBascula, 'mono'), inciso('FECHA DE RECEPCIÓN:', fechaRecepcionCert(c), 'fuerte')));
    hoja.appendChild(fila(inciso('MANIFIESTO:', manifiestoCert(c), 'mono'), inciso('POZO:', c.Pozo, 'fuerte')));
    hoja.appendChild(fila(inciso('TRANSPORTISTA:', c.Transportista, 'fuerte')));   // renglon entero: un numero de autorizacion no se parte (revisor v0.35.1)
    hoja.appendChild(fila(inciso('AUTORIZACIÓN DE LA PLANTA:', CFG.autorizacionPlanta || 'pendiente (ASEA-03-011-A)', 'mono')));
    // Tabla de incisos: APAGADA desde la v0.39.0 — un certificado ampara UNA gondola y un solo manifiesto. No se borra:
    // el dia que llegue residuo de plataforma marina, varios contenedores viajan en un embarque y la tabla es justo lo
    // que hace falta (decision de Carlos). C-49 (v0.42.0): hoy NUNCA se enciende — datosCertificado fija TipoBulto 'gondola'
    // y la emision no escribe Embarques/Manifiestos. Encenderla pide las dos cosas, no solo un TipoBulto = 'contenedor'.
    if (c.TipoBulto === 'contenedor') hoja.appendChild(tablaIncisosCert(c));
    const pieFirma = el('div', 'ct-firmas');
    const firma = el('div', 'ct-firma'); firma.appendChild(el('div', 'ct-raya')); firma.appendChild(el('div', 'ct-nombre', CFG.responsableTecnico || '(nombre pendiente)')); firma.appendChild(el('div', 'ct-cargo', 'Responsable técnico de planta · CALYTEK')); pieFirma.appendChild(firma);
    const url = urlVerificacion(CFG.urlVerificacion, c);
    const qrCaja = el('div', 'ct-qr');
    if (url) { qrCaja.appendChild(qrSvg(url, 84)); qrCaja.appendChild(el('div', 'ct-qr-leyenda', 'Verifique este certificado con el código QR')); }
    else qrCaja.appendChild(el('div', 'ct-qr-leyenda', c.Estado === 'borrador' ? 'El QR se genera al emitir' : 'Sin QR: certificado sin sufijo de verificación'));
    pieFirma.appendChild(qrCaja); hoja.appendChild(pieFirma);
    const pie = el('div', 'ct-pie');
    pie.appendChild(el('div', '', `Recibido en las instalaciones de MINSA ENERGY · Planta CALYTEK, ${CFG.domicilioPlanta} · ${CFG.resolutivo} · RFC ${CFG.rfc}`));
    pie.appendChild(el('div', '', 'Este certificado ampara únicamente el residuo y la cantidad aquí descritos.'));
    hoja.appendChild(pie);
}
export function imprimirCertificado() {
    if ($('ctImprimir').disabled) return;
    // La hoja de impresion (carta horizontal, solo el formato) se enciende SOLO mientras dura el print: @page no admite clase
    // y el ticket termico / Reportes no deben salir en carta horizontal.  La CSP no deja un <style> inline; un <link> propio si.
    // v0.37.0: se enciende por `media` (print / not all), NO por `disabled`: un <link> apagado y vuelto a encender ya no se aplicaba
    // al imprimir (medido con Edge headless: la SEGUNDA impresion de la sesion salia en blanco y vertical).
    document.body.classList.add('imprimiendo-certificado'); $('cssImpresionCert').media = 'print';
    // v0.36.1: «Guardar como PDF» toma el nombre del archivo de document.title — con el folio, cada certificado sale con el suyo (pedido de Carlos).
    const tituloApp = document.title; const c = estado.certificadoAbierto; if (c && c.Title) document.title = `Certificado de tratamiento ${c.Title}`;
    // quitar corre UNA vez: por afterprint o por el respaldo de 60 s (algun WebView no manda afterprint). El respaldo se cancela al
    // salir por afterprint — si sobreviviera, pisaria document.title un minuto despues (lo cazo la E2E en tiempo virtual).
    // C-46 (v0.42.0): el respaldo se arma ANTES de print() — un afterprint dentro de print() ya lo encuentra — y `hecho` lo vuelve idempotente.
    let hecho = false;
    const quitar = () => { if (hecho) return; hecho = true; clearTimeout(respaldo); document.body.classList.remove('imprimiendo-certificado'); $('cssImpresionCert').media = 'not all'; document.title = tituloApp; window.removeEventListener('afterprint', quitar); };
    const respaldo = setTimeout(quitar, 60000);
    window.addEventListener('afterprint', quitar);
    window.print();
}
