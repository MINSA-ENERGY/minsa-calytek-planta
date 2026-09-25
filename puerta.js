// CALYTEK Planta — la puerta (llegada de la gondola).
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { compuerta, CORRIENTES, etiquetaCorriente, lista, palabraCompuerta, placaNormal, plural, PUEDE, registroPuerta, siguienteFolio } from './reglas.js';
import { $, anclar, avisar, el, embarquesDelAno, escribiendo, estado, firmar, fundirEnVentana, L, limpiarAvisos, nombreDe, opciones, porId, prealtaFirmada, refrescarCliente, renglon, vivo } from './nucleo.js';
import { entrar, irA } from './navegacion.js';
import { abrirPesaje, elegirVistaGondolas } from './gondolas.js';
import { gondolasDe } from './prealtas.js';
import { etiquetaVigencia } from './padron.js';
import { esRechazo } from './archivos.js';

// ================================================================ PUERTA

export function pintarPuerta() {
    const firmadas = estado.prealtas.filter(prealtaFirmada);
    // Aviso informativo: pre-altas en borrador (sin firma) — la puerta no las ve en el selector hasta que se firmen.
    // S-01: la que trae sello pero no renglon en PLANTA_Firmas cuenta igual: la compuerta manda.
    const borradores = estado.prealtas.filter(p => p.Estado === 'borrador' || (p.Estado === 'firmada' && !prealtaFirmada(p)));
    const pp = $('puPendientes'); pp.classList.toggle('oculto', !borradores.length);
    if (borradores.length) pp.textContent = `${borradores.length === 1 ? 'Hay 1 pre-alta por firmar' : `Hay ${borradores.length} pre-altas por firmar`}: ${borradores.map(p => p.Title).join(' · ')}. Sus góndolas no pueden entrar hasta que el validador firme.`;
    opciones($('puPrealta'), firmadas, p => p.id, p => `${p.Title} · ${etiquetaCorriente(p.Corriente) || '?'} · ${nombreDe(estado.carriers, p.CarrierId)}`);
    if (!$('puPrealta').value && firmadas.length === 1) $('puPrealta').value = String(firmadas[0].id);   // U-40: una sola firmada no se hace elegir
    pintarProgramasPuerta(firmadas);
    pintarChoferesPuerta();
    pintarUnidadesPuerta();
    $('btnCompuerta').disabled = !PUEDE.puerta(estado.rol);
    $('puSoloLectura').classList.toggle('oculto', PUEDE.puerta(estado.rol));   // U-38 (v0.26.0): texto fijo, no avisar(): el refresco silencioso lo repetía cada 2 min y pisaba el aviso que se leía
    $('puSoloLectura').textContent = estado.rol === 'validador' ? 'Tu rol es validador: aquí solo ves; capturas y firmas en Pre-altas.' : 'Tu rol es de lectura: puedes ver, no capturar.';   // U-50: al validador no se le dice «lectura»
    irSubpaso(estado.subpasoPuerta);   // el repintado (refresco de 2 min) no mueve de pantalla: solo la repinta
    pintarPrevioPuerta();
}

// ---------------------------------------------------------------- tanda 4: el asistente de la llegada (M1–M3)

/** Las tres pantallas del paso 1. La barra de cinco pasos avanza por tercios dentro del primero. */
const SUBPASOS = { 1: 'Programa', 2: 'Vehículo y chofer', 3: 'Carga' };
const SIGUIENTE_SUBPASO = { 1: 'Siguiente · el vehículo ›', 2: 'Siguiente · la carga ›' };
export function irSubpaso(s, enfocar = false) {
    s = SUBPASOS[s] ? Number(s) : 1;
    const cambio = s !== estado.subpasoPuerta;
    estado.subpasoPuerta = s;
    for (const n of [1, 2, 3]) $('puBloque' + n).hidden = n !== s;
    for (const b of $('puSubpasos').querySelectorAll('button')) { if (Number(b.dataset.s) === s) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current'); }
    $('puPaso1').dataset.tercio = String(s);
    $('puPasoK').textContent = `Paso 1 de 5 · ${SUBPASOS[s]}`;
    $('btnPuAtras').hidden = s === 1;
    if (s > 1) $('btnPuAtras').textContent = `‹ ${SUBPASOS[s - 1]}`;
    $('btnPuSiguiente').hidden = s === 3 || (s === 1 && !estado.prealtas.some(prealtaFirmada));   // U-118
    if (s < 3) $('btnPuSiguiente').textContent = SIGUIENTE_SUBPASO[s];
    $('btnCompuerta').hidden = s !== 3;
    if (cambio && estado.pestana === 'puerta') window.scrollTo({ top: 0 });
    if (enfocar) { const q = $('puBloque' + s).querySelector('.pregunta'); if (q) { q.tabIndex = -1; q.focus({ preventScroll: true }); } }
}
/** Una empieza nueva: con el programa ya elegido (las góndolas del mismo programa llegan en fila) arranca en el vehículo. */
export const subpasoInicial = () => ($('puPrealta').value ? 2 : 1);

/**
 * M2: un renglón tocable que actúa como radio. `valor` va en data-v; el que está elegido lleva .sel y aria-pressed, como los
 * chips de unidad de siempre (U-33). `dato` es la columna derecha: texto o un nodo (la etiqueta de vigencia).
 */
export function renglonOpcion({ valor, sel, titulo, mono = false, detalle, dato, clase = '' }) {
    const b = el('button', `o ${clase}${sel ? ' sel' : ''}`.trim()); b.type = 'button';
    b.dataset.v = String(valor);
    b.setAttribute('aria-pressed', String(!!sel));
    b.appendChild(el('span', 'r'));
    const t = el('span', 't'); t.appendChild(el('b', mono ? 'mono' : '', titulo));
    if (detalle) t.appendChild(el('small', '', detalle));
    b.appendChild(t);
    if (dato instanceof Node) { dato.classList.add('d'); b.appendChild(dato); } else if (dato) b.appendChild(el('span', 'd', dato));
    return b;
}
export function marcarOpcion(cont, valor) {
    for (const x of cont.querySelectorAll('.o')) { const on = valor !== '' && x.dataset.v === String(valor); x.classList.toggle('sel', on); x.setAttribute('aria-pressed', String(on)); }
}
/** Elegir en un renglón es elegir en el select oculto: el listener de `change` de siempre repinta lo que depende de él. */
function elegirEnSelect(id, v) { $(id).value = String(v); $(id).dispatchEvent(new Event('change', { bubbles: true })); }

function pintarProgramasPuerta(firmadas) {
    const cont = $('puProgramas'); cont.textContent = '';
    // U-118 (v0.73.0): sin firmadas no hay a dónde avanzar: la salida a Pre-altas es un botón y «Siguiente» se esconde (irSubpaso).
    $('puPistaPrograma').hidden = !firmadas.length;
    if (!firmadas.length) {
        cont.appendChild(el('p', 'pista', 'No hay programas firmados todavía. Se captura y se firma primero en Pre-altas.'));
        const ir = el('button', 'secundario', 'Ir a Pre-altas ›'); ir.type = 'button'; ir.id = 'btnPuIrPrealtas';
        ir.addEventListener('click', () => irA('prealtas'));
        cont.appendChild(ir);
        return;
    }
    for (const p of firmadas) {
        const { rec, esp } = gondolasDe(p);
        const b = renglonOpcion({ valor: p.id, sel: String(p.id) === $('puPrealta').value, titulo: p.Title,
            detalle: [nombreDe(estado.carriers, p.CarrierId), etiquetaCorriente(p.Corriente)].filter(Boolean).join(' · '),
            dato: esp ? `${rec} de ${esp}` : plural(rec, 'recibida') });
        b.addEventListener('click', () => elegirEnSelect('puPrealta', p.id));
        cont.appendChild(b);
    }
}
export function pintarCorrientesPuerta() {
    const cont = $('puCorrientes'); cont.textContent = '';
    for (const [valor, titulo] of CORRIENTES) {   // C-55: del catálogo, no del select oculto
        const b = renglonOpcion({ valor, sel: valor === $('puCorriente').value, titulo });
        b.addEventListener('click', () => elegirEnSelect('puCorriente', valor));
        cont.appendChild(b);
    }
}
/** Escritorio: «Así va la góndola» a la derecha, en vivo; la misma estructura que «Así va la pre-alta». En celular no se ve. */
function pintarResumenPuerta() {
    const r = $('puResumen'); r.textContent = '';
    r.appendChild(el('h2', '', 'Así va la góndola'));
    const dl = el('dl');
    const fila = (dt, dd) => { dl.appendChild(el('dt', '', dt)); dl.appendChild(el('dd', dd ? '' : 'f', dd || '—')); };
    const sec = t => dl.appendChild(el('span', 'sec', t));
    const pre = porId(estado.prealtas, $('puPrealta').value);
    sec('Programa'); fila('Pre-alta', pre ? pre.Title : ''); fila('Carrier', pre && pre.CarrierId ? nombreDe(estado.carriers, pre.CarrierId) : '');
    if (pre) { const { rec, esp } = gondolasDe(pre); fila('Recibidas', esp ? `${rec} de ${esp}` : String(rec)); }
    const chofer = $('puChoferNombre').value.trim() || (porId(estado.choferes, $('puChofer').value) || {}).Title;
    sec('Vehículo y chofer'); fila('Placa tractor', placaNormal($('puPlaca').value)); fila('Placa plana', placaNormal($('puPlacaPlana').value)); fila('Chofer', chofer);
    sec('Carga'); fila('Manifiesto', $('puManifiesto').value.trim()); fila('Corriente', $('puCorriente').value ? etiquetaCorriente($('puCorriente').value) : ''); fila('Art. 79', $('pu79').checked ? 'identificado y etiquetado' : '');
    r.appendChild(dl);
    // U-120 (v0.73.0): la pista sale de CAMPOS_PUERTA, la misma lista que cuenta el botón; el chofer no detiene, manda a Espera.
    r.appendChild(el('p', 'pista', `Obligatorio: ${CAMPOS_PUERTA.map(([, t]) => t.replace(/^(el|la) /, '')).join(', ').replace(/, ([^,]*)$/, ' y $1')}. Sin chofer la góndola sale en Espera. Después se revisan los documentos.`));
}
/** Dónde vive cada campo que la compuerta necesita, y qué recibe el foco cuando falta (el select oculto no puede). */
const PANTALLA_DE = { puPrealta: 1, puPlaca: 2, puPlacaPlana: 2, puChofer: 2, puChoferNombre: 2, puManifiesto: 3, puCorriente: 3 };
export function enfocarCampoPuerta(id) {
    irSubpaso(PANTALLA_DE[id] || 1);
    if (id === 'puPlaca' || id === 'puPlacaPlana') $('puTeclear').open = true;
    const destino = id === 'puPrealta' ? $('puProgramas').querySelector('.o') : id === 'puCorriente' ? $('puCorrientes').querySelector('.o') : $(id);
    if (destino) destino.focus();
}
// C-56 (v0.52.0): subpasoDeRegla vive en reglas.js, con tabla explícita y su prueba contra cada nombre que emite compuerta().
// I6 (7-sep): las unidades que la pre-alta autorizo (o, si no marco ninguna, todas las activas del carrier).
// Tocar una llena placa tractor y plana; la seleccion se marca comparando con lo que hay en el campo, asi que
// teclear otra placa la desmarca sola. La compuerta sigue evaluando la placa del campo, no la seleccion.
export function pintarUnidadesPuerta() {
    const cont = $('puUnidades'); cont.textContent = '';
    const pre = porId(estado.prealtas, $('puPrealta').value);
    $('puUnidadesTitulo').textContent = 'Tocar la unidad llena las dos placas.';
    if (!pre) { cont.appendChild(el('p', 'pista', 'Elige el programa para ver sus unidades.')); return; }
    const autorizadas = lista(pre.UnidadesIds).map(Number);
    let us = estado.unidades.filter(u => u.Activo !== false && Number(u.CarrierId) === Number(pre.CarrierId));
    if (autorizadas.length) us = us.filter(u => autorizadas.includes(u.id));
    $('puUnidadesTitulo').textContent = us.length ? `${plural(us.length, 'unidad', 'unidades')} en este programa; tocar una llena las dos placas.` : 'Tocar la unidad llena las dos placas.';
    // Tanda 4: sin unidades en el padrón no hay renglón que tocar, así que las placas se teclean a la vista.
    if (!us.length) { $('puTeclear').open = true; cont.appendChild(el('p', 'pista', 'Este programa no tiene unidades en el padrón: teclea la placa y saldrá como no amparada.')); return; }
    const actual = placaNormal($('puPlaca').value);
    for (const u of us) {
        // M2 (tanda 4): el chip de siempre como renglón tocable. Conserva .u y data-placa (U-17 / C-28) y el aria-pressed de U-33.
        const v = etiquetaVigencia('unidades', u);
        const b = renglonOpcion({ valor: placaNormal(u.Title), sel: placaNormal(u.Title) === actual, clase: 'u', mono: true,
            titulo: u.PlacaPlana ? `${u.Title} · ${u.PlacaPlana}` : u.Title,
            detalle: [u.TipoUnidad || 'unidad', u.CapacidadKg ? `${Number(u.CapacidadKg).toLocaleString('es-MX')} kg` : null, v ? null : 'vigencias al día'].filter(Boolean).join(' · '),
            dato: v });
        b.dataset.placa = placaNormal(u.Title);   // U-17: la seleccion al teclear compara contra esto, no contra el texto del chip
        b.addEventListener('click', () => {
            const x = vivo('unidades', u);   // C-58 (v0.51.0): la unidad VIVA al clic; la Puerta no se repinta con captura a medias
            $('puPlaca').value = x.Title; $('puPlacaPlana').value = x.PlacaPlana || '';
            marcarChip(cont, b.dataset.placa);   // C-28
            pintarPrevioPuerta();
        });
        cont.appendChild(b);
    }
}
/** C-28 (v0.28.0): el chip de unidad elegido, por clase Y aria-pressed, igual al clic que al teclear la placa (U-17 / U-33). */
export function marcarChip(cont, placa) {
    for (const x of cont.querySelectorAll('.u')) { const on = !!placa && x.dataset.placa === placa; x.classList.toggle('sel', on); x.setAttribute('aria-pressed', String(on)); }
}
export function pintarChoferesPuerta() {
    const pre = porId(estado.prealtas, $('puPrealta').value);
    const carrierId = pre ? Number(pre.CarrierId) : null;
    const ch = estado.choferes.filter(x => x.Activo !== false && (!carrierId || Number(x.CarrierId) === carrierId));
    // U-21 (v0.26.0): como los chips de unidad — primero los que la pre-alta autorizó; los demás del carrier siguen
    // eligibles pero marcados «fuera del programa», porque la compuerta los saca ámbar (reglas.js «Chofer vs pre-alta»).
    const autorizados = pre ? lista(pre.ChoferesIds).map(String) : [];
    const fuera = x => autorizados.length > 0 && !autorizados.includes(String(x.id));
    const orden = ch.filter(x => !fuera(x)).concat(ch.filter(fuera));
    opciones($('puChofer'), orden, x => x.id, x => fuera(x) ? `${x.Title} · fuera del programa` : x.Title, '— no está en el padrón —');
    // M2 (tanda 4): los mismos choferes, en el mismo orden, como renglones; el select oculto sigue siendo lo que se evalúa.
    const cont = $('puChoferes'); cont.textContent = '';
    if (!pre) { cont.appendChild(el('p', 'pista', 'Elige el programa para ver sus choferes.')); return; }
    if (!orden.length) { $('puChoferOtro').open = true; cont.appendChild(el('p', 'pista', 'El carrier no tiene choferes en el padrón: teclea el nombre como viene en la licencia; saldrá en Espera.')); return; }
    for (const x of orden) {
        const b = renglonOpcion({ valor: x.id, sel: String(x.id) === $('puChofer').value, titulo: x.Title,
            detalle: fuera(x) ? 'fuera del programa' : 'licencia federal', dato: etiquetaVigencia('choferes', x) || el('span', 'e-ok', 'vigente'),
            clase: fuera(x) ? 'fuera' : '' });
        b.addEventListener('click', () => elegirEnSelect('puChofer', x.id));
        cont.appendChild(b);
    }
}

/**
 * Lee la forma de la puerta y corre la compuerta con lo que haya capturado. PURA respecto del DOM
 * de salida: no pinta nada. La usan el boton (veredicto) y la vista previa en vivo, para que las
 * dos digan siempre lo mismo — si divergieran, la que miente es la que el operador lee primero.
 */
function evaluarPuerta() {
    const pre = porId(estado.prealtas, $('puPrealta').value);
    const carrier = pre ? porId(estado.carriers, pre.CarrierId) : null;
    const placa = placaNormal($('puPlaca').value);
    const unidad = estado.unidades.find(u => placaNormal(u.Title) === placa && u.Activo !== false) || null;
    const chofer = porId(estado.choferes, $('puChofer').value);
    const campos = {
        manifiesto: $('puManifiesto').value.trim(), placaTractor: $('puPlaca').value.trim(),
        placaPlana: $('puPlacaPlana').value.trim(), corriente: $('puCorriente').value,
        choferNombre: $('puChoferNombre').value.trim() || (chofer ? chofer.Title : ''), art79: $('pu79').checked
    };
    // Si la unidad esta en el padron pero con OTRO carrier, es como si no estuviera: la ampara otro oficio.
    const unidadDelCarrier = unidad && carrier && Number(unidad.CarrierId) === carrier.id ? unidad : null;
    const r = compuerta({ prealta: pre, carrier, unidad: unidadDelCarrier, chofer, ...campos, avisoDias: CONFIG.avisoVigenciaDias });
    if (unidad && !unidadDelCarrier) r.hallazgos.unshift({ clase: 'legal', regla: 'Placa', ok: false, detalle: `la placa ${placa} está en el padrón pero amparada por otro carrier (${nombreDe(estado.carriers, unidad.CarrierId)})` });
    if (unidad && !unidadDelCarrier && r.resultado !== 'rechazo-legal') r.resultado = 'rechazo-legal';
    if (!campos.art79) r.hallazgos.push({ clase: 'aviso', regla: 'Art. 79', ok: true, detalle: 'no se marcó la verificación de identificación/etiquetado/envasado' });
    return { ...r, campos, pre, carrier, unidad: unidadDelCarrier, chofer, placa };
}

export function correrCompuerta() {
    limpiarAvisos();
    // U-01 (v0.21.0): mientras falte un campo de CAMPOS_PUERTA el boton dice «Faltan N datos» y NO abre el
    // veredicto: enfoca el primer faltante. Antes con manifiesto o corriente vacios abria «No entra» y un toque
    // mas dejaba un folio R- por un dato que no se habia tecleado (reglas.js los cuenta como hallazgo legal).
    // Tanda 4: el faltante puede estar en otra pantalla del asistente; se va a ella antes de enfocarlo.
    const falta = CAMPOS_PUERTA.find(([id]) => !String($(id).value).trim());
    if (falta) { avisar(`Falta ${falta[1]}.`, 'error'); enfocarCampoPuerta(falta[0]); return; }
    const e = evaluarPuerta();
    if (!e.placa) { avisar('Falta la placa del tractor.', 'error'); enfocarCampoPuerta('puPlaca'); return; }
    estado.ultimaCompuerta = e;
    pintarResultadoCompuerta();
}

/** Los campos que la compuerta necesita para que su respuesta sea un pronostico y no un «faltan datos». */
const CAMPOS_PUERTA = [['puPrealta', 'el programa'], ['puPlaca', 'la placa del tractor'], ['puManifiesto', 'el manifiesto'], ['puCorriente', 'la corriente']];

/**
 * G3 + G4: el estado de cada bloque y la lista viva de reglas. Corre la MISMA evaluacion que el
 * boton; mientras falte un campo la lista se atenua y la linea dice que falta, para que no se lea
 * como veredicto (el veredicto sigue siendo la pantalla completa, sin cambios).
 */
export function pintarPrevioPuerta() {
    const e = evaluarPuerta();
    const faltan = CAMPOS_PUERTA.filter(([id]) => !String($(id).value).trim()).map(([, n]) => n);

    // U-48 (v0.29.0): cada dato del bloque es una lista de alternativas (el chofer vale por el select O por el nombre en
    // licencia). El bloque 2 contaba solo la placa y decía «Listo» mientras la lista viva anunciaba Espera por el chofer;
    // el chofer sigue sin ser obligatorio para correr la compuerta (CAMPOS_PUERTA no cambia).
    const hay = alts => alts.some(id => String($(id).value).trim());
    // Tanda 4: el estado de cada pantalla vive en su botón de la fila de subpasos (antes, en la cabecera de cada bloque).
    const bloques = [['puSub1', 'puEst1', [['puPrealta', 'programa']], 'sin'],
                     ['puSub2', 'puEst2', [['puPlaca', 'la placa'], ['puChofer', 'puChoferNombre', 'el chofer']], 'falta'],
                     ['puSub3', 'puEst3', [['puManifiesto', 'el manifiesto'], ['puCorriente', 'la corriente']], 'falta']];
    for (const [bloque, est, datos, pendiente] of bloques) {
        const faltantes = datos.filter(d => !hay(d.slice(0, -1))).map(d => d[d.length - 1]);
        const ok = !faltantes.length, n = faltantes.length;
        $(bloque).classList.toggle('listo', ok);
        // D2 (2026-09-08): el estado cuenta lo que falta; el detalle («falta la placa») queda en el title y en el texto.
        // U-85 (v0.51.0): el «· programa» va en su propio span para que el celular muestre solo el conteo (estilo.css .est .det).
        $(est).textContent = ok ? 'Listo' : n === 1 ? 'Falta 1' : `Faltan ${n}`;
        if (!ok && n === 1) $(est).appendChild(el('span', 'det', ` · ${faltantes[0]}`));
        $(est).title = ok ? '' : `${pendiente} ${faltantes.join(' y ')}`;
    }

    // D2 (2026-09-08): el boton dice «Faltan 2 datos · el manifiesto, la corriente» hasta que todo esta capturado.
    // Sigue siendo el mismo boton y sigue corriendo la compuerta: solo cambia lo que dice (tanda 4: «Revisar documentos»).
    const btn = $('btnCompuerta');
    btn.classList.toggle('incompleto', faltan.length > 0);
    btn.textContent = '';
    if (faltan.length) {
        btn.appendChild(document.createTextNode(faltan.length === 1 ? 'Falta 1 dato' : `Faltan ${faltan.length} datos`));
        btn.appendChild(el('small', '', faltan.join(', ')));
    } else btn.textContent = 'Revisar documentos ›';

    pintarVivoPuerta(e, faltan);
    pintarResumenPuerta();
}
/**
 * M3 (tanda 4): la lista larga «Lo que va a revisar» se vuelve un renglón al pie: «Hasta ahora · 6 en verde · 1 aviso».
 * Mientras falte un dato no se cuentan las reglas que solo se quejan de ESE dato (el manifiesto vacío sale «legal» en
 * reglas.js): contarlas diría «no entra» por algo que todavía no se teclea. Con todo capturado, dice qué va a salir.
 */
const REGLAS_DEL_CAMPO = { 'el programa': ['Pre-alta', 'Carrier'], 'la placa del tractor': ['Placa'], 'el manifiesto': ['Manifiesto'], 'la corriente': ['Corriente'] };
function pintarVivoPuerta(e, faltan) {
    const vivo = $('puVivo'); vivo.textContent = '';
    if (faltan.includes('el programa')) { vivo.textContent = 'Elige el programa: la revisión en vivo empieza con él.'; return; }
    const callar = new Set(faltan.flatMap(f => REGLAS_DEL_CAMPO[f] || []));
    if (!String($('puChofer').value).trim() && !$('puChoferNombre').value.trim()) callar.add('Chofer');
    if (faltan.length) callar.add('Art. 79');   // la casilla vive en la última pantalla: avisar antes de llegar a ella es ruido
    const hs = e.hallazgos.filter(h => !callar.has(h.regla));
    const n = c => hs.filter(h => h.clase === c).length;
    const cifra = (texto, tono) => { vivo.appendChild(document.createTextNode(' · ')); vivo.appendChild(el('b', 'e-' + tono, texto)); };
    if (faltan.length) vivo.appendChild(el('span', '', 'Hasta ahora'));
    else { const def = VEREDICTOS[e.resultado]; vivo.appendChild(document.createTextNode('Va a salir ')); vivo.appendChild(el('b', def.clase.replace('v-', 'e-'), def.palabra)); }
    cifra(`${n('ok')} en verde`, 'ok');
    if (n('aviso')) cifra(plural(n('aviso'), 'aviso'), 'warn');
    if (n('comercial')) cifra(`${n('comercial')} para gerencia`, 'warn');
    if (n('legal')) cifra(`${n('legal')} no ${n('legal') === 1 ? 'pasa' : 'pasan'}`, 'bad');
    if (faltan.length) vivo.appendChild(el('span', 'falta', ` · falta ${faltan.join(', ')}`));
}

/**
 * EL SEMAFORO. La compuerta tiene tres salidas y la pantalla se ve distinta en cada una: el color
 * cubre la cabecera, la palabra es enorme y la regla que decidio va primero y en negrita. Un chofer
 * a tres metros sabe si entra sin leer nada (mockup aprobado 2026-09-05).
 */
// Tanda 4 (M4): cada veredicto dice qué hacer después — el guion para el chofer y un botón que dice a dónde lleva.
// La frase de «Pasa» se arma con la cuenta de reglas (pintarResultadoCompuerta).
const VEREDICTOS = {
    pasa: { clase: 'v-ok', palabra: palabraCompuerta('pasa').palabra, frase: '', boton: 'Seguir a peso bruto ›', guion: '«Pásate a la báscula.»' },
    'rechazo-legal': { clase: 'v-bad', palabra: palabraCompuerta('rechazo-legal').palabra, frase: 'Falta un requisito legal · sin dispensa', boton: 'Registrar el rechazo' },
    'excepcion-comercial': { clase: 'v-warn', palabra: palabraCompuerta('excepcion-comercial').palabra, frase: 'Falta un documento comercial · lo autoriza gerencia', boton: 'Mandar a gerencia y volver a la lista', guion: '«Espérate en el patio; gerencia está autorizando.»' }
};
/**
 * Pinta una lista de hallazgos. Compartida por el veredicto y por la vista previa de la puerta:
 * la que decide (`decide`) va primero y en negrita; sin decide, el orden es el que trajo la regla.
 */
function pintarHallazgos(ul, hallazgos, decide) {
    const orden = [...hallazgos].sort((a, b) => (a.clase === decide ? 0 : 1) - (b.clase === decide ? 0 : 1));
    ul.textContent = '';
    for (const h of orden) {
        const li = el('li', h.clase === 'ok' ? 'ok' : h.clase);
        if (decide && h.clase === decide) li.classList.add('culpable');
        li.appendChild(el('i'));
        const texto = el('span', '', h.regla);
        if (h.detalle) texto.appendChild(el('span', 'd', h.detalle));
        li.appendChild(texto);
        li.appendChild(el('small', '', h.clase === 'ok' ? 'ok' : h.clase));
        ul.appendChild(li);
    }
}

export function cerrarVeredicto() {
    const v = $('veredicto');
    const estabaAbierto = !v.classList.contains('oculto');
    v.classList.add('oculto');
    v.classList.remove('v-ok', 'v-bad', 'v-warn');
    // El foco vuelve a donde estaba antes de abrirlo (teclado y lector de pantalla; F3 de la auditoria del 7-sep).
    if (estabaAbierto && estado.focoAntesVeredicto && document.contains(estado.focoAntesVeredicto)) estado.focoAntesVeredicto.focus();
    estado.focoAntesVeredicto = null;
}
function pintarResultadoCompuerta() {
    const r = estado.ultimaCompuerta;
    const def = VEREDICTOS[r.resultado];
    const v = $('veredicto');
    v.classList.remove('oculto', 'v-ok', 'v-bad', 'v-warn');
    v.classList.add(def.clase);
    const verdes = r.hallazgos.filter(h => h.clase === 'ok').length, avisos = r.hallazgos.filter(h => h.clase === 'aviso').length;
    $('vkK').textContent = `Paso 2 de 5 · Veredicto · ${placaNormal(r.campos.placaTractor) || '—'}`;
    $('vkW').textContent = def.palabra;
    $('vkM').textContent = def.frase || `${verdes} en verde${avisos ? ` · ${plural(avisos, 'aviso')}` : ''}`;
    // El folio R- se conoce al registrar; el E- nace en el bruto: se anuncia con puntos, no con un numero que cambie.
    const aa = String(new Date().getFullYear()).slice(-2);
    $('vkFolio').textContent = r.resultado === 'pasa' ? `→ E-${aa}-·····` : r.resultado === 'rechazo-legal' ? `R-${aa}-····` : '';

    const ctx = $('vkCtx'); ctx.textContent = '';
    const tarjeta = (b, s) => { const d = el('div'); d.appendChild(el('b', '', b)); d.appendChild(el('span', '', s)); ctx.appendChild(d); };
    tarjeta(placaNormal(r.campos.placaTractor) || '—', r.unidad ? `tractor · oficio ${r.unidad.FolioOficio || '?'}` : 'tractor · no está en el padrón');
    if (r.campos.placaPlana) tarjeta(placaNormal(r.campos.placaPlana), 'plana');
    tarjeta(r.campos.manifiesto || '—', 'manifiesto');
    tarjeta(r.campos.choferNombre || (r.chofer ? r.chofer.Title : '—'), r.chofer ? 'chofer del padrón' : 'chofer fuera del padrón');

    // Las reglas: primero las que decidieron (legal si es rechazo, comercial si es excepcion), luego el resto.
    const decide = r.resultado === 'rechazo-legal' ? 'legal' : r.resultado === 'excepcion-comercial' ? 'comercial' : null;
    pintarHallazgos($('puHallazgos'), r.hallazgos, decide);
    // M4 (tanda 4): lo único que hay que leer va arriba —la regla que decidió, o los avisos si pasa— y las verdes se pliegan
    // detrás de «Ver las N reglas». Sin avisos ni culpable, la lista entera queda plegada: todo está en verde.
    const plegables = r.hallazgos.filter(h => h.clase === 'ok').length;
    $('puHallazgos').classList.toggle('plegada', plegables > 0);
    $('btnVerReglas').classList.toggle('oculto', plegables === 0);
    $('btnVerReglas').textContent = `Ver las ${r.hallazgos.length} reglas`;
    $('vkLbl').textContent = decide ? 'Qué lo decidió' : avisos ? plural(avisos, 'aviso · no detiene', 'avisos · no detienen') : 'Todo en verde';

    $('puExcepcion').classList.toggle('oculto', r.resultado !== 'excepcion-comercial');
    const guion = $('vkGuion'), nota = $('vkNota');
    nota.classList.toggle('oculto', r.resultado !== 'rechazo-legal');
    if (r.resultado === 'rechazo-legal') {
        const culpables = r.hallazgos.filter(h => h.clase === 'legal').map(h => h.regla.toLowerCase());
        guion.lastElementChild.textContent = `«No puede descargar: falta lo legal (${culpables.join(', ')}).»`;
        nota.textContent = '¿Fue un error al teclear? «Corregir lo capturado» antes de registrar. Si no, se corrige en el oficio o en la pre-alta, no aquí, y el chofer se lleva su constancia con el folio R-.';
    } else guion.lastElementChild.textContent = def.guion;
    $('btnRegistrarPuerta').textContent = def.boton;
    v.scrollTop = 0;
    // Dialogo modal: el foco entra al veredicto y sale con Esc (F3). El boton principal recibe el foco.
    estado.focoAntesVeredicto = document.activeElement;
    // U-119 (v0.73.0): en un rechazo legal el foco va a «Corregir lo capturado»: un Enter de más ya no emite el folio R-.
    (r.resultado === 'rechazo-legal' ? $('btnVolverVeredicto') : $('btnRegistrarPuerta')).focus();
}
document.addEventListener('keydown', ev => {
    const v = $('veredicto');
    if (v.classList.contains('oculto')) return;
    if (ev.key === 'Escape') { ev.preventDefault(); cerrarVeredicto(); return; }
    // U-16 (v0.22.0): el veredicto es role=dialog aria-modal pero no es un <dialog>: Tab se salia a la puerta que esta
    // detras. Se cicla entre lo enfocable del veredicto (motivo de la excepcion, si esta, y los dos botones).
    if (ev.key !== 'Tab') return;
    const focables = [...v.querySelectorAll('button, textarea, input, select')].filter(x => !x.disabled && x.offsetParent !== null);
    if (!focables.length) return;
    const i = focables.indexOf(document.activeElement);
    const siguiente = ev.shiftKey ? (i <= 0 ? focables[focables.length - 1] : focables[i - 1]) : (i < 0 || i === focables.length - 1 ? focables[0] : focables[i + 1]);
    ev.preventDefault(); siguiente.focus();
});

export async function registrarPuerta() {
    const r = estado.ultimaCompuerta; if (!r) return;
    // C-52 (v0.51.0): el veredicto se calculó al abrirlo y puede llevar minutos abierto; si el refresco trajo otro padrón o pre-alta,
    // se vuelve a evaluar y, si cambió, se pinta el nuevo y NO se registra hasta un segundo toque. Antes se escribía el viejo.
    const fresca = evaluarPuerta();
    if (fresca.resultado !== r.resultado || JSON.stringify(fresca.hallazgos) !== JSON.stringify(r.hallazgos)) {
        const foco = estado.focoAntesVeredicto;
        estado.ultimaCompuerta = fresca; pintarResultadoCompuerta(); estado.focoAntesVeredicto = foco;
        avisar('El padrón o la pre-alta cambiaron mientras el veredicto estaba abierto: revisa el veredicto nuevo y vuelve a tocar el botón.', 'ojo');
        return;
    }
    if (r.resultado === 'excepcion-comercial' && !$('puMotivo').value.trim()) {
        avisar('La excepción lleva motivo escrito, no una casilla.', 'error');
        $('puMotivo').setAttribute('aria-invalid', 'true'); $('puMotivo').focus();   // U-39: el campo que falta se marca y recibe el foco
        return;
    }
    const textoBoton = $('btnRegistrarPuerta').textContent;
    await escribiendo('btnRegistrarPuerta', async () => { try {   // C-24
        await refrescarCliente();
        const ahora = new Date().toISOString();
        const esRechazo = r.resultado === 'rechazo-legal';
        // El folio R- se asigna aqui; el E- NO: nace en la primera pasada de bascula (ticket 03 regla 1).
        // Se relee la lista antes de escoger el numero: el estado local puede tener horas.
        // Se relee el ANIO, no la ventana: el consecutivo R- corre de enero a diciembre.
        const av = texto => { $('btnRegistrarPuerta').textContent = texto; };   // C-17: el boton dice que reintenta
        const delAno = esRechazo ? await embarquesDelAno(av) : [];
        if (esRechazo) fundirEnVentana(delAno);
        const folio = esRechazo ? siguienteFolio('R', delAno.map(e => e.Title)) : '';
        const campos = registroPuerta(r, { folio, ahora, motivo: $('puMotivo').value, usuario: estado.cuenta.username });   // C-77 (v0.74.0): el armado vive en reglas.js
        const nuevo = await estado.cliente.crearRenglon(estado.siteId, L.embarques, campos, av);
        if (esRechazo) await asegurarFolioUnico(nuevo, 'R', av);
        anclar('embarques', nuevo);   // C-23: sin duplicar el id si el refresco ya lo trajo
        for (const id of ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre', 'puMotivo']) $(id).value = '';
        $('pu79').checked = false;
        // U-02 (v0.21.0): tambien el chofer (la siguiente gondola heredaba su ChoferId) y se repintan chips y vista
        // previa: antes los bloques seguian en «Listo» con los campos vacios. El programa y la corriente se quedan:
        // las gondolas del mismo programa llegan en fila (tanda 4: por eso la siguiente arranca en el vehículo).
        $('puChofer').value = ''; delete $('puChoferNombre').dataset.auto;
        $('puTeclear').open = false; $('puChoferOtro').open = false;
        cerrarVeredicto();
        estado.ultimaCompuerta = null;
        pintarChoferesPuerta(); pintarUnidadesPuerta(); pintarPrevioPuerta();
        irSubpaso(subpasoInicial());
        // M4 (tanda 4): cada veredicto lleva a lo que sigue. Pasa → el peso bruto de ESA góndola; Espera → la lista, donde queda
        // en ámbar; No entra → Rechazos, con su folio R-. El aviso va después de irA(), que limpia los avisos.
        irA('bascula');
        elegirVistaGondolas(esRechazo ? 'rechazos' : 'planta');
        if (r.resultado === 'pasa') abrirPesaje(vivo('embarques', nuevo), 'bruto');
        avisar(esRechazo ? `Rechazo registrado con folio ${nuevo.Title}. La góndola no entra.` :
            r.resultado === 'pasa' ? 'Registrado: pasa. Sigue el peso bruto, con la góndola llena en la báscula.' :
            'Registrado. Gerencia lo ve en Hoy › Pendiente revisar; en Góndolas queda en ámbar hasta que autorice.', 'bien');
    } catch (e) {
        avisar('No se pudo registrar: ' + (e && e.message ? e.message : e), 'error');
    } finally { $('btnRegistrarPuerta').textContent = textoBoton; } });
}

/**
 * Dos celulares pueden pedir el mismo folio en el mismo segundo. Despues de escribir se relee la
 * lista: si el folio quedo repetido, el renglon MAS NUEVO (id mayor) se renumera al siguiente libre.
 * Un hueco se investiga; un duplicado no puede existir.
 * Parametrizada por lista y columna (C-04, v0.22.0): la campana L- de las pre-altas pasa por aqui igual que R- y E-.
 */
const FOLIOS = {
    E: { lista: () => L.embarques, campo: 'Title', releer: embarquesDelAno },
    R: { lista: () => L.embarques, campo: 'Title', releer: embarquesDelAno },
    L: { lista: () => L.prealtas, campo: 'Campana', releer: av => estado.cliente.renglones(estado.siteId, L.prealtas, null, av) },
    C: { lista: () => L.certificados, campo: 'Title', releer: av => estado.cliente.renglones(estado.siteId, L.certificados, null, av) }   // v0.35.0: CT-AA-NNNN
};
export async function asegurarFolioUnico(renglon, tipo, avisar) {
    const { lista: lst, campo, releer } = FOLIOS[tipo];
    // C-16 (v0.25.0): la unicidad solo necesita saber si ALGUIEN mas tiene ese folio: un filtro de igualdad (una peticion)
    // en vez de releer el anio entero (hasta 4 paginas de 500) por segunda vez. El anio solo se relee si hubo choque.
    const iguales = await estado.cliente.renglones(estado.siteId, lst(), `fields/${campo} eq '${String(renglon[campo]).replace(/'/g, "''")}'`, avisar);
    if (iguales.length <= 1) return renglon[campo];
    const masNuevo = iguales.reduce((a, b) => (a.id > b.id ? a : b));
    if (masNuevo.id !== renglon.id) return renglon[campo];
    const todos = await releer(avisar);   // el folio lleva el anio: el resto no puede chocar
    const nuevoFolio = siguienteFolio(tipo, todos.map(x => x[campo]));
    await estado.cliente.actualizarRenglon(estado.siteId, lst(), renglon.id, { [campo]: nuevoFolio });
    renglon[campo] = nuevoFolio;
    return nuevoFolio;
}


/** U-40: el programa elegido también es captura. */
export function camposCapturaPuerta() { return ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre', 'puMotivo', 'puPrealta']; }   // función y no const: capturaAMedias() la llama desde arriba
export function puertaConCaptura() { return camposCapturaPuerta().some(id => $(id).value.trim()); }
