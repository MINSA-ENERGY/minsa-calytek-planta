// CALYTEK Planta — control de recepcion. Sesion 1: la puerta (pre-alta, compuerta, bascula).
//
// Estructura heredada de minsa-captura-app: MSAL por REDIRECCION (el popup se rompe en
// celulares), token en sessionStorage, nada de innerHTML (todo textContent), y la app apunta a
// la biblioteca (ticket 06): las fotos del indicador van al buzon con el contrato de nombres.
//
// Lo que la app NO protege: quien decide que puede escribir cada persona es SharePoint. Los
// roles de PLANTA_Roles son para que cada quien vea su pantalla, no una barrera; la traza real
// de quien firmo o autorizo la deja SharePoint en Creado por / Modificado por de cada renglon.

import { CONFIG } from './config.js';
import { CORRIENTES, placaNormal, subpasoDeRegla } from './reglas.js';
import { $, abrirForma, avisar, cerrarForma, confirmar, el, estado, hayCaptura, huellaForma, limpiarAvisos, porId, salir, VERSION, vivo } from './nucleo.js';
import { arrancar, botonesRail, entrar, irDesdePestana, recargar, repintar, soltarPesaje } from './navegacion.js';
import { camposCapturaPuerta, cerrarVeredicto, correrCompuerta, enfocarCampoPuerta, irSubpaso, marcarChip, marcarOpcion, pintarChoferesPuerta, pintarCorrientesPuerta, pintarPrevioPuerta, pintarUnidadesPuerta, puertaConCaptura, registrarPuerta, subpasoInicial } from './puerta.js';
import { abrirHojaCapturado, abrirTicketPop, cerrarAsistente, cerrarHojaCapturado, guardarPeso, kgG, revisarNeto, tomarFoto } from './gondolas.js';
import { cerrarPrealta, editarPrealta, elegirVistaPrealtas, eliminarPrealta, firmarPrealta, guardarPrealta, nuevaPrealta, pintarUnidadesChoferesPrealta, salirAsistentePrealta } from './prealtas.js';
import { cancelarCertificado, confirmarSustitucion, emitirCertificado, imprimirCertificado, mostrarCorreccion, sustituirCertificado } from './certificado.js';
import { abrirFormaPadron, cerrarFormaPadron, FORMA_PADRON, guardarPadron } from './padron.js';
import './hoy.js';
import './reportes.js';
import { filtrarArbol } from './archivos.js';

// ================================================================ arranque

$('btnEntrar').addEventListener('click', entrar);
// Tema: sin valor guardado sigue al sistema; el pill fija claro u oscuro (localStorage, por dispositivo) y no vuelve al sistema.
const TEMA_LLAVE = 'calytek-planta-tema';
function aplicarTema(tema) {
    if (tema === 'claro' || tema === 'oscuro') document.documentElement.dataset.tema = tema;
    else delete document.documentElement.dataset.tema;
    for (const b of document.querySelectorAll('.tema button')) b.setAttribute('aria-pressed', String(b.dataset.tema === tema));
    pintarBarraEstado();
}
// La barra de estado del teléfono toma su color del <meta theme-color>. Con dos metas por prefers-color-scheme el
// botón sol/luna de la app cambiaba el fondo sin avisarle, y la barra quedaba del otro tono (foto de Carlos, 2026-09-07).
// Se le da SIEMPRE el --bg que la página está usando de verdad.
function pintarBarraEstado() {
    const bg = getComputedStyle(document.documentElement).getPropertyValue('--bg').trim();
    const m = $('metaTema'); if (m && bg) m.setAttribute('content', bg);
}
window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', pintarBarraEstado);
try { aplicarTema(localStorage.getItem(TEMA_LLAVE)); } catch (e) { aplicarTema(null); }
for (const b of document.querySelectorAll('.tema button')) b.addEventListener('click', () => {
    // Cada boton FIJA su tema; tocarlo dos veces no cambia nada. Antes el segundo toque volvia al sistema,
    // y en un celular en oscuro eso se veia como oscuro sin anillo, y claro-claro alternaba (Carlos, 2026-09-06).
    const nuevo = b.dataset.tema;
    aplicarTema(nuevo);
    try { localStorage.setItem(TEMA_LLAVE, nuevo); } catch (e) { /* sin almacenamiento */ }
});
$('btnSalir').addEventListener('click', () => { $('menuRail').open = false; salir(); });
$('btnSalirMovil').addEventListener('click', () => { $('menuMovil').open = false; salir(); });
// v0.32.0: el rail se PLIEGA a 64 px con la marca y se despliega con el chevron (Proyectos v0.28.1); se recuerda por dispositivo.
const RAIL_LLAVE = 'calytek-planta-rail';
function plegarRail(p) {
    $('app').classList.toggle('rail-plegado', p);
    $('btnMarca').setAttribute('aria-expanded', String(!p)); $('btnPlegar').setAttribute('aria-expanded', String(!p));
    try { localStorage.setItem(RAIL_LLAVE, p ? 'plegado' : 'abierto'); } catch (e) { /* sin almacenamiento */ }
}
try { plegarRail(localStorage.getItem(RAIL_LLAVE) === 'plegado'); } catch (e) { /* sin almacenamiento */ }
$('btnMarca').addEventListener('click', () => plegarRail(true));
$('btnPlegar').addEventListener('click', () => plegarRail(false));
// Imprimir los reportes: la hoja de impresión solo deja ver el ticket; con esta clase deja ver la sección (estilo.css @media print).
$('btnImprimirReportes').addEventListener('click', () => { document.body.classList.add('imprimiendo-reportes'); window.print(); });
// Archivos (v0.33.0): los filtros no releen nada, solo esconden y muestran lo ya leído.
$('arBusca').addEventListener('input', filtrarArbol);
for (const id of ['arTipo', 'arPrograma']) $(id).addEventListener('change', filtrarArbol);
window.addEventListener('afterprint', () => document.body.classList.remove('imprimiendo-reportes'));
// El menu «···» se cierra al elegir algo o al tocar fuera.
document.addEventListener('click', ev => { for (const id of ['menuRail', 'menuMovil']) { const m = $(id); if (m.open && !m.contains(ev.target)) m.open = false; } });   // C-35: los dos menús en un listener
$('btnActualizar').addEventListener('click', () => { $('menuRail').open = false; recargar(); });
$('btnActualizarMovil').addEventListener('click', () => { $('menuMovil').open = false; recargar(); });
// Al volver a la app (el celular estuvo en el bolsillo, la pestana en segundo plano) se relee si
// el estado tiene mas de un minuto. Sin aviso: solo cambia lo que se ve.
document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && estado.siteId && Date.now() - estado.cargadoEl > 60000) recargar(true);
});
// Y cada CONFIG.refrescoMs mientras la app esta a la vista: la excepcion que gerencia autoriza desde su sesion le
// llega a la caseta sin tocar Actualizar (H8, auditoria del 7-sep). recargar() ya no pisa un pesaje ni un veredicto abiertos.
if (CONFIG.refrescoMs > 0) setInterval(() => {
    if (document.visibilityState === 'visible' && estado.siteId && Date.now() - estado.cargadoEl > CONFIG.refrescoMs - 5000) recargar(true);
}, CONFIG.refrescoMs);
for (const b of botonesRail()) b.addEventListener('click', () => irDesdePestana(b.dataset.p));   // C-27 / U-41
// Tanda 2 (v0.47.0): la llegada se abre con «+ Nueva góndola» y regresa por la miga «Góndolas».
// Tanda 4: una góndola nueva empieza en la primera pantalla que falta — el vehículo, si el programa ya está elegido. Si hay
// una captura a medias (se salió a mirar la lista y se volvió), se retoma donde estaba.
$('btnNuevaGondola').addEventListener('click', () => {
    if (!['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre'].some(id => $(id).value.trim())) estado.subpasoPuerta = subpasoInicial();
    irDesdePestana('puerta');
});
$('migaGondolasPuerta').addEventListener('click', () => irDesdePestana('bascula'));
/** v0.72.0 (a pedido de Carlos): Cancelar con algo capturado pregunta, como en Pre-altas; Descartar deja la puerta en cero. */
$('btnCancelarPuerta').addEventListener('click', async () => {
    if (puertaConCaptura()) {
        const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: 'Esta góndola tiene datos sin guardar. Si sales, se pierden.' });
        if (!ok) return;
        for (const id of [...camposCapturaPuerta(), 'puChofer']) $(id).value = '';
        $('pu79').checked = false; delete $('puChoferNombre').dataset.auto;
        $('puTeclear').open = false; $('puChoferOtro').open = false;
        cerrarVeredicto(); estado.ultimaCompuerta = null;
        pintarChoferesPuerta(); pintarUnidadesPuerta(); pintarPrevioPuerta();
        irSubpaso(subpasoInicial());
    }
    irDesdePestana('bascula');
});
// v0.72.0 (bug que reportó Carlos): regresar y cambiar de programa dejaba las placas y el chofer del programa anterior
// —de otro carrier— capturados y en «Así va la góndola». Al cambiar de programa se limpian los datos que dependen de él.
let prealtaPuertaPrevia = '';
$('puPrealta').addEventListener('change', () => {
    const v = $('puPrealta').value;
    if (prealtaPuertaPrevia && v !== prealtaPuertaPrevia) for (const id of ['puPlaca', 'puPlacaPlana', 'puChofer', 'puChoferNombre']) $(id).value = '';
    prealtaPuertaPrevia = v;
    marcarOpcion($('puProgramas'), v); pintarChoferesPuerta(); pintarUnidadesPuerta(); pintarPrevioPuerta(); });
$('puChofer').addEventListener('change', () => marcarOpcion($('puChoferes'), $('puChofer').value));
// C-55 (v0.52.0): los tres controles de corriente salen del mismo catálogo (CORRIENTES, reglas.js).
for (const id of ['puCorriente', 'paCorriente']) for (const [v, txt] of CORRIENTES) { const o = el('option', '', txt); o.value = v; $(id).appendChild(o); }
for (const [v, txt] of CORRIENTES) { const l = el('label', 'chk'), c = el('input'); c.type = 'checkbox'; c.name = 'pcCorr'; c.value = v; l.appendChild(c); l.appendChild(document.createTextNode(' ' + txt)); $('pcCorrientes').appendChild(l); }
$('puCorriente').addEventListener('change', () => marcarOpcion($('puCorrientes'), $('puCorriente').value));
for (const b of $('puSubpasos').querySelectorAll('button')) b.addEventListener('click', () => irSubpaso(Number(b.dataset.s), true));
$('btnPuAtras').addEventListener('click', () => irSubpaso(estado.subpasoPuerta - 1, true));
$('btnPuSiguiente').addEventListener('click', () => {
    // U-83 (v0.51.0): sin programa no hay unidades ni choferes que tocar; se queda y enfoca la lista, como «Faltan N datos».
    if (estado.subpasoPuerta === 1 && !$('puPrealta').value) { avisar('Falta el programa: elígelo para ver sus unidades y choferes.', 'error'); enfocarCampoPuerta('puPrealta'); return; }
    irSubpaso(estado.subpasoPuerta + 1, true);
});
$('btnVerReglas').addEventListener('click', () => { $('puHallazgos').classList.remove('plegada'); $('btnVerReglas').classList.add('oculto'); });
pintarCorrientesPuerta();
// U-17 (v0.22.0): se compara con data-placa; el textContent del chip trae pegado el <small> («55XY9Kgóndola · 20,000 kg»)
// y una unidad sin placa plana nunca se marcaba al teclear.
$('puPlaca').addEventListener('input', () => marcarChip($('puUnidades'), placaNormal($('puPlaca').value)));   // C-28: clase y aria-pressed
$('puChofer').addEventListener('change', () => {
    // El nombre se rellena con el del padrón y se REEMPLAZA al cambiar de chofer, salvo que alguien lo haya editado a mano
    // (se distingue porque el valor ya no es el que puso la app). Antes solo se llenaba si estaba vacío y se quedaba el anterior.
    const ch = porId(estado.choferes, $('puChofer').value), n = $('puChoferNombre');
    if (ch && (!n.value || n.value === n.dataset.auto)) { n.value = ch.Title; n.dataset.auto = ch.Title; }
});
// La vista previa se repinta con cada tecla: `compuerta()` es pura y no toca la red.
for (const id of ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre']) $(id).addEventListener('input', pintarPrevioPuerta);
for (const id of ['puChofer', 'puCorriente', 'pu79']) $(id).addEventListener('change', pintarPrevioPuerta);
$('btnCompuerta').addEventListener('click', correrCompuerta);
$('puMotivo').addEventListener('input', () => $('puMotivo').removeAttribute('aria-invalid'));   // U-39: al escribir el motivo se quita la marca
$('btnRegistrarPuerta').addEventListener('click', registrarPuerta);
// Tanda 4: «Corregir lo capturado» regresa a la pantalla de la regla que decidió (o a la carga, si pasó).
$('btnVolverVeredicto').addEventListener('click', () => {
    const r = estado.ultimaCompuerta, decide = r && (r.resultado === 'rechazo-legal' ? 'legal' : r.resultado === 'excepcion-comercial' ? 'comercial' : null);
    const culpable = decide && r.hallazgos.find(h => h.clase === decide);
    cerrarVeredicto(); estado.ultimaCompuerta = null;
    irSubpaso(culpable ? subpasoDeRegla(culpable.regla) : 3);
});
// Teclado numerico de la bascula (celular): digitos enteros, sin coma. En computadora se oculta por CSS.
$('teclado').addEventListener('click', ev => {
    const b = ev.target.closest('button[data-k]'); if (!b) return;
    const kg = $('baKg'); const k = b.dataset.k;
    kg.value = k === 'borrar' ? kg.value.slice(0, -1) : (kg.value + k).replace(/^0+(?=\d)/, '').slice(0, 6);
    revisarNeto();
});
$('btnFoto').addEventListener('click', () => $('baFotoEntrada').click());
$('baFotoEntrada').addEventListener('change', e => { if (e.target.files[0]) tomarFoto(e.target.files[0]); e.target.value = ''; });
$('baKg').addEventListener('input', revisarNeto);
$('btnGuardarPeso').addEventListener('click', guardarPeso);
$('btnCancelarPeso').addEventListener('click', async () => {
    // U-24 (v0.26.0): con kg o foto ya capturados pregunta, como las formas (U-09); vacío cierra directo. En el celular
    // Cancelar queda justo bajo el Guardar pegado y un toque con guante tiraba la foto del indicador sin avisar.
    // C-27 (v0.28.0): la compuerta es soltarPesaje(), la misma que al cambiar de pestaña.
    if (await soltarPesaje()) repintar();   // C-12: la lista se repinta (un refresco con el pesaje abierto la dejaba vieja)
});
$('btnImprimir').addEventListener('click', () => window.print());
// Tanda 5: los botones del asistente de la báscula. «Góndolas» de las migas suelta el pesaje como Cancelar (pregunta si hay algo).
$('btnBaGondolas').addEventListener('click', async () => {
    if (!$('baPesar').classList.contains('oculto')) { if (await soltarPesaje()) repintar(); return; }
    cerrarAsistente(); repintar();
});
$('btnPausaLista').addEventListener('click', () => { cerrarAsistente(); repintar(); });
$('btnPausaTicket').addEventListener('click', () => { const e = estado.asisEmbarque; if (e) abrirTicketPop([vivo('embarques', e)], 0); });
$('btnTerminar').addEventListener('click', () => {   // M8: de vuelta a la lista, con el aviso de lo que se cerró
    const e = estado.asisEmbarque;
    cerrarAsistente(); repintar();
    if (e) avisar(`${e.Title} cerrada · neto ${kgG(e.NetoKg)}. Queda en Cerradas hoy.`, 'bien');
});
$('btnLoCapturado').addEventListener('click', abrirHojaCapturado);
$('btnCerrarCapturado').addEventListener('click', () => { cerrarHojaCapturado(); $('btnLoCapturado').focus(); });
$('baVelo').addEventListener('click', () => { cerrarHojaCapturado(); $('btnLoCapturado').focus(); });
$('baCapturado').addEventListener('keydown', ev => {
    if (!$('baCapturado').classList.contains('abierta') || $('baVelo').hidden) return;
    if (ev.key === 'Escape') { cerrarHojaCapturado(); $('btnLoCapturado').focus(); return; }
    // U-81: con la hoja abierta, Tab no sale al contenido atenuado de atrás (como el veredicto, U-16).
    if (ev.key !== 'Tab') return;
    const focables = [...$('baCapturado').querySelectorAll('button, a[href], input, select, textarea')].filter(x => !x.disabled && x.offsetParent !== null);
    if (!focables.length) return;
    const i = focables.indexOf(document.activeElement);
    ev.preventDefault();
    focables[ev.shiftKey ? (i <= 0 ? focables.length - 1 : i - 1) : (i < 0 || i === focables.length - 1 ? 0 : i + 1)].focus();
});
$('btnNuevaPrealta').addEventListener('click', nuevaPrealta);
$('btnPaBases').addEventListener('click', () => abrirForma('paRecientes'));
$('btnPaBasesCerrar').addEventListener('click', () => cerrarForma('paRecientes'));
for (const b of $('paTabs').querySelectorAll('button')) b.addEventListener('click', () => elegirVistaPrealtas(b.dataset.pa));
for (const b of document.querySelectorAll('.pa-kpis button')) b.addEventListener('click', () => elegirVistaPrealtas(b.dataset.pa));   // U-110: el conteo lleva a su pestaña
$('paCarrier').addEventListener('change', pintarUnidadesChoferesPrealta);
$('btnAltaCarrierPrealta').addEventListener('click', () => abrirFormaPadron('carriers'));
$('btnGuardarPrealta').addEventListener('click', guardarPrealta);
$('btnCancelarPrealta').addEventListener('click', salirAsistentePrealta);
// Escape cierra el <dialog> sin pasar por Cancelar: la edicion pendiente del padron se suelta igual.
for (const clave of Object.keys(FORMA_PADRON)) $(FORMA_PADRON[clave].forma).addEventListener('close', () => { if (estado.padronEdit && estado.padronEdit.clave === clave) estado.padronEdit = null; });
// U-09 (v0.22.0): Escape sobre un formulario CON algo capturado pregunta antes de tirarlo — el <dialog> nativo cerraba sin
// pasar por Cancelar y la pre-alta de 14 campos se perdia. Vacio (o solo con los valores por omision) cierra directo; el
// boton Cancelar sigue cerrando sin preguntar. Un select cuenta solo si no esta en su primera opcion.
for (const id of ['pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer']) $(id).addEventListener('cancel', async ev => {
    if (!hayCaptura($(id)) || huellaForma($(id)) === $(id).dataset.huella) return;   // U-34: sin cambios desde que abrió → cierra directo
    ev.preventDefault();
    const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: 'Este formulario tiene datos sin guardar. Si lo cierras, se pierden.' });
    if (ok) cerrarForma(id);
});
$('btnFirmar').addEventListener('click', firmarPrealta);
$('btnCerrarPrealta').addEventListener('click', cerrarPrealta);
$('btnEmitirCertificado').addEventListener('click', () => emitirCertificado());   // v0.39.0: sobre la gondola abierta
$('btnSustituirCertificado').addEventListener('click', sustituirCertificado);
$('btnCancelarCertificado').addEventListener('click', cancelarCertificado);
$('ctCorregirOk').addEventListener('click', confirmarSustitucion);   // U-69 (v0.40.0)
$('ctCorregirVolver').addEventListener('click', () => { limpiarAvisos(); mostrarCorreccion(false); });
$('ctCerrar').addEventListener('click', () => $('dlgCertificado').close());
// U-75 (v0.45.0): en el celular el papel es vista previa; tocarlo (o Enter/espacio) lo amplia y lo regresa. En escritorio no cambia nada.
{ const v = document.querySelector('#dlgCertificado .ct-vista');
  const alternar = () => { const si = v.classList.toggle('ampliada'); v.setAttribute('aria-pressed', String(si)); };
  v.addEventListener('click', alternar);
  v.addEventListener('keydown', ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); alternar(); } });
  $('dlgCertificado').addEventListener('close', () => { v.classList.remove('ampliada'); v.setAttribute('aria-pressed', 'false'); }); }
// C-38 (v0.40.0): el estado se suelta en el evento close, no en el boton: Escape cerraba sin soltarlo. Solo si el dialogo
// SIGUE cerrado: el close llega encolado y, si se reabrio en ese hueco, vaciaba la gondola recien abierta (lo cazo la E2E;
// es la misma carrera que obs. 603 midio en confirmar()).
$('dlgCertificado').addEventListener('close', () => { if ($('dlgCertificado').open) return; estado.certificadoAbierto = null; estado.certificadoEmbarque = null; mostrarCorreccion(false); });
$('ctImprimir').addEventListener('click', imprimirCertificado);
$('btnEliminarPrealta').addEventListener('click', eliminarPrealta);
$('btnEditarPrealta').addEventListener('click', editarPrealta);
$('btnVolverPrealtas').addEventListener('click', () => cerrarForma('paDetalle'));
$('btnNuevoCarrier').addEventListener('click', ev => { ev.preventDefault(); ev.stopPropagation(); abrirFormaPadron('carriers'); });
$('btnCancelarCarrier').addEventListener('click', () => cerrarFormaPadron('carriers'));
$('btnGuardarCarrier').addEventListener('click', () => guardarPadron('carriers'));
$('btnNuevaUnidad').addEventListener('click', ev => { ev.preventDefault(); ev.stopPropagation(); abrirFormaPadron('unidades'); });
$('btnCancelarUnidad').addEventListener('click', () => cerrarFormaPadron('unidades'));
$('btnGuardarUnidad').addEventListener('click', () => guardarPadron('unidades'));
$('btnNuevoChofer').addEventListener('click', ev => { ev.preventDefault(); ev.stopPropagation(); abrirFormaPadron('choferes'); });
$('btnCancelarChofer').addEventListener('click', () => cerrarFormaPadron('choferes'));
$('btnGuardarChofer').addEventListener('click', () => guardarPadron('choferes'));

$('pie').textContent = `CALYTEK Planta ${VERSION}`;
arrancar();

// S-06 (v0.21.0): el service worker nuevo ya NO se activa solo sobre una pestana abierta. Antes skipWaiting +
// controllerchange recargaban la pagina en cuanto se publicaba una version: a media tara se perdian foto y peso.
// Ahora el SW nuevo queda ESPERANDO, la app ofrece «Actualizar» y solo al tocarlo (o al cerrar la app) se activa.
if ('serviceWorker' in navigator) {
    const habiaControlador = !!navigator.serviceWorker.controller;
    let recargandoSw = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
        if (!habiaControlador || recargandoSw) return;
        recargandoSw = true; window.location.reload();
    });
    const ofrecer = sw => {
        const caja = $('nuevaVersion'); if (!caja || !sw) return;
        caja.classList.remove('oculto');
        $('btnNuevaVersion').onclick = () => { caja.classList.add('oculto'); sw.postMessage('activar'); };
    };
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').then(r => {
            if (typeof r.addEventListener !== 'function') return r.update && r.update();
            if (r.waiting) ofrecer(r.waiting);
            r.addEventListener('updatefound', () => {
                const nuevo = r.installing; if (!nuevo) return;
                nuevo.addEventListener('statechange', () => { if (nuevo.state === 'installed' && navigator.serviceWorker.controller) ofrecer(nuevo); });
            });
            return r.update();
        }).catch(() => {});
    });
}
