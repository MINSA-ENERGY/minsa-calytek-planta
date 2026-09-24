// CALYTEK Planta — la sesion (entrar, arrancar), irA, el rail, recargar y la compuerta de captura a medias.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { lista, rolDe } from './reglas.js';
import { $, asistentePadronAbierto, asistentePrealtaAbierto, avisar, cargarTodo, confirmar, el, escrituras, estado, limpiarAvisos, pasoEntrada, pca, pintarInsignias, pintarSync, ponerQuien, prepararMsal, refrescarCliente } from './nucleo.js';
import { cerrarVeredicto, pintarPuerta, puertaConCaptura } from './puerta.js';
import { cerrarAsistente, pintarGondolas, soltarFotoPrevia } from './gondolas.js';
import { ocultarListoPrealta, pintarPrealtas } from './prealtas.js';
import { pintarPadron } from './padron.js';
import { pintarHoy } from './hoy.js';
import { pintarReportes } from './reportes.js';
import { pintarArchivos } from './archivos.js';

export async function entrar() {
    pasoEntrada('Entrando…');
    try {
        await prepararMsal();
        if (pca.getAllAccounts().length === 0) { pasoEntrada('Abriendo el inicio de sesión de Microsoft…'); await pca.loginRedirect({ scopes: CONFIG.scopes }); return; }
        await sesionIniciada();
    } catch (e) {
        pasoEntrada(null);
        avisar('No se pudo entrar: ' + (e && e.message ? e.message : e), 'error');
        $('textoEntrar').textContent = 'Vuelve a intentarlo.';
    }
}
export async function arrancar() {
    if (window.self !== window.top) return;
    try {
        const respuesta = await prepararMsal();
        if (respuesta || pca.getAllAccounts().length > 0) {
            pasoEntrada('Entrando…');
            await sesionIniciada();
        }
    } catch (e) {
        pasoEntrada(null);
        avisar('No se pudo terminar el inicio de sesión: ' + (e && e.message ? e.message : e), 'error');
        $('textoEntrar').textContent = 'Vuelve a intentarlo.';
    }
}

async function sesionIniciada() {
    estado.cuenta = pca.getAllAccounts()[0];
    await refrescarCliente();
    ponerQuien(estado.cuenta.username);
    $('btnSalir').classList.remove('oculto');
    const lnk = $('lnkSharePoint'); lnk.href = `https://${CONFIG.sharepointHost}${CONFIG.sitio}`; lnk.classList.remove('oculto');   // v0.32.0
    pasoEntrada('Abriendo el sitio de CALYTEK…');
    estado.siteId = await estado.cliente.sitio(CONFIG.sharepointHost, CONFIG.sitio);
    pasoEntrada('Leyendo las listas…');
    await cargarTodo();
    estado.rol = rolDe(estado.cuenta.username, estado.roles);
    ponerQuien(estado.cuenta.username, estado.rol);
    $('pantallaEntrar').classList.add('oculto');
    $('rail').classList.remove('oculto');
    $('barraMovil').classList.remove('oculto');
    $('syncMovil').classList.remove('oculto'); pintarSync();
    // Pantalla inicial por rol: quien captura abre en Góndolas (decisión 12 del rediseño, v0.47.0; antes en la puerta),
    // quien firma en las pre-altas, gerencia y lectura en «Hoy» (la consola). Igual en celular y en computadora.
    irA(estado.rol === 'trazabilidad' ? 'bascula' : estado.rol === 'validador' ? 'prealtas' : 'hoy');
}

/**
 * Vuelve a leer las siete listas y repinta la pestana abierta. Antes el estado se leia UNA vez al
 * entrar: gerencia con «Hoy» abierto toda la manana no veia la excepcion nueva, y el basculista no
 * veia la gondola que la caseta acababa de registrar en otro celular (auditoria 2026-09-05).
 * Corre por el boton Actualizar y solo al volver a la app tras un rato (visibilitychange).
 *//**
 * Hay una captura a medias en pantalla: repintar la pestana la borraria. Antes solo protegia el pesaje y el
 * veredicto; el alta de un carrier en el celular se perdia a los 2 minutos por el refresco automatico y al
 * volver a la app (Carlos, 2026-09-08). Cubre todo formulario abierto y la puerta con algo tecleado.
 */
function capturaAMedias() {
    const abierto = id => !$(id).classList.contains('oculto');
    // Tanda 5 (decisión 11): el asistente de la báscula abierto —pesaje, pausa «A descargar» o ticket— cuenta entero.
    if (abierto('baAsis') || abierto('veredicto')) return true;
    if (estado.pestana === 'prealtas' && asistentePrealtaAbierto()) return true;
    if (estado.pestana === 'padron' && asistentePadronAbierto()) return true;   // v0.72.0: el alta del padrón en página   // v0.54.0: el asistente de pre-alta a la vista
    if (['paDetalle', 'paRecientes', 'pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer'].some(id => $(id).open)) return true;
    if (estado.pestana === 'puerta' && puertaConCaptura()) return true;
    return false;
}

let recargando = false;
export async function recargar(silencioso = false) {
    if (recargando || !estado.siteId) return;
    // C-23 (v0.28.0): con una escritura en vuelo o un confirm abierto NO se sustituyen las listas: el Object.assign que sigue
    // al await caeria sobre un objeto huerfano (la gondola anulada seguia «En planta», gerencia autorizaba dos veces). Antes
    // solo se frenaba el REPINTADO (capturaAMedias); cargarTodo() corria igual. El timer lo vuelve a intentar al minuto.
    // C-38 (v0.40.0): tambien con el certificado abierto — gerencia esta leyendo un papel que el refresco repintaria debajo.
    // C-53 (v0.51.0): y con el ticket abierto — es el papel que se lleva el chofer; ‹ › no deben recorrer otra lista a media lectura.
    if (escrituras > 0 || $('dlg').open || $('dlgCertificado').open || $('dlgTicket').open) { if (!silencioso && escrituras > 0) avisar('Espera a que termine de guardar y vuelve a actualizar.', 'ojo'); return; }
    recargando = true;
    for (const id of ['btnActualizar', 'btnActualizarMovil']) $(id).disabled = true;
    pintarSync(true);
    try {
        await refrescarCliente();
        await cargarTodo();
        // v0.33.0: el Actualizar a mano relee el arbol; el refresco de 2 min no lo tira (se pierden las carpetas abiertas).
        // C-34 (v0.34.0): se tira aqui, ya leido todo, y no antes del await: en medio un toggle del arbol caia sobre null.
        if (!silencioso) estado.archivos = null;
        estado.rol = rolDe(estado.cuenta.username, estado.roles);
        ponerQuien(estado.cuenta.username, estado.rol);
        // No pisar una captura a medias. El refresco SILENCIOSO repinta en su lugar (U-04, v0.21.0): irA()
        // limpiaba el aviso que se estaba leyendo y mandaba la pagina al tope cada 2 minutos.
        if (capturaAMedias()) pintarInsignias();
        else if (silencioso) repintar();
        else irA(estado.pestana);
        if (!silencioso) avisar('Datos actualizados.', 'bien');
    } catch (e) {
        avisar('No se pudo actualizar: ' + (e && e.message ? e.message : e), 'error');
    } finally {
        recargando = false;
        for (const id of ['btnActualizar', 'btnActualizarMovil']) $(id).disabled = false;
        pintarSync();
    }
}

// ---------------------------------------------------------------- navegacion

const PINTORES = { hoy: () => pintarHoy(), puerta: () => pintarPuerta(), bascula: () => pintarGondolas(), prealtas: () => pintarPrealtas(), padron: () => pintarPadron(), reportes: () => pintarReportes(), archivos: () => pintarArchivos() };
const SECCIONES = Object.keys(PINTORES);
/** Los botones del rail: las pestañas de siempre más la sección aparte (Reportes, v0.32.0). */
export const botonesRail = () => [...$('pestanas').querySelectorAll('button'), ...$('pestanasExtra').querySelectorAll('button')];
/** Repinta la pestana abierta SIN tocar avisos, veredicto ni scroll (refresco silencioso y cambios de pre-alta). */
export function repintar() { pintarInsignias(); PINTORES[estado.pestana](); }
/**
 * C-27 / U-41 (v0.28.0): salir de la bascula con un pesaje abierto pasa por la misma compuerta que «Cancelar» (U-24): con kg o
 * foto pregunta, y al salir suelta estado.pesando, la foto y su blob URL. Antes irA() ocultaba la tarjeta sin preguntar y el
 * basculista volvia a fotografiar el indicador; el JPEG comprimido quedaba vivo hasta el siguiente pesaje. Devuelve si se salio.
 */
const pesajeConAlgo = () => !$('baPesar').classList.contains('oculto') && !!($('baKg').value.trim() || estado.fotoBytes);
function descartarPesaje() { estado.pesando = null; estado.fotoBytes = null; soltarFotoPrevia(); cerrarAsistente(); }
export async function soltarPesaje() {
    if (pesajeConAlgo()) {
        const { ok } = await confirmar({ titulo: 'Cancelar el pesaje', peligro: true, ok: 'Descartar', texto: 'Se pierden el peso tecleado y la foto del indicador; habría que volver a tomarla.' });
        if (!ok) return false;
    }
    descartarPesaje();
    return true;
}
/** El clic en una pestana: sincrono salvo que haya que preguntar (la E2E y el usuario esperan la pestana pintada al soltar). */
/** Tanda 2 (v0.47.0): el boton «Góndolas» del rail cubre la lista (p-bascula) y la llegada (p-puerta, hasta el asistente de la tanda 4); entra por la lista. */
const RAIL_DE = { puerta: 'gondolas', bascula: 'gondolas' };
const PESTANA_DE = { gondolas: 'bascula' };
export function irDesdePestana(p) {
    p = PESTANA_DE[p] || p;
    if (p !== 'bascula' && pesajeConAlgo()) { soltarPesaje().then(ok => { if (ok) irA(p); }); return; }
    irA(p);   // vacio, en la pausa o en el ticket: irA lo suelta sin preguntar (C-51)
}
/**
 * C-51 (v0.51.0): salir de Góndolas suelta el asistente en TODOS sus estados. Antes solo se soltaba con #baPesar visible: desde la
 * pausa «A descargar» o el ticket, #baAsis se quedaba abierto en una sección oculta y capturaAMedias() congelaba el refresco de
 * todas las pestañas. Una sola salida, llamada desde irA(): el rail, la miga y registrarPuerta pasan por ahí.
 */
function salirDelAsistente() { if (!$('baAsis').classList.contains('oculto') || estado.pesando) descartarPesaje(); }
/* R6 (v0.65.0): el cuadro de rótulo dice la sección, su número de hoja entre las del rail y la fecha de hoy. */
function pintarRotulo(p) {
    const bs = [...botonesRail()], i = bs.findIndex(b => b.dataset.p === p);
    if (i < 0) return;
    $('rotuloT').textContent = 'MINSA ENERGY - PLANTA CALYTEK - ' + bs[i].querySelector('span').textContent.toUpperCase();
    $('rotuloH').textContent = (i + 1) + ' / ' + bs.length;
    $('rotuloF').textContent = new Date().toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-').replace('.', '');
}
export function irA(p) {
    if (p !== 'bascula') salirDelAsistente();
    ocultarListoPrealta();   // P10: la confirmación no guarda nada; salir o volver por el rail la suelta
    if (p === 'padron' && estado.padronVista.v !== 'asis') estado.padronVista.v = 'lista';   // v0.72.0: el rail lleva a la lista; v0.72.0: un alta a medias se conserva
    estado.pestana = p;
    for (const b of botonesRail()) { if (b.dataset.p === (RAIL_DE[p] || p)) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); }   // U-57 (v0.29.0): <nav> con aria-current, como .mn-rail de la piel; el role=tablist prometía flechas y tabpanel que no había
    for (const s of SECCIONES) $('p-' + s).classList.toggle('oculto', s !== p);
    pintarRotulo(RAIL_DE[p] || p);
    limpiarAvisos();
    cerrarVeredicto();
    repintar();
    window.scrollTo({ top: 0 });
}


setInterval(() => { if (estado.siteId) pintarSync(recargando); }, 15000);
