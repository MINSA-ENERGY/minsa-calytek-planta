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
import { crearCliente } from './graph.js';
import { comprimir } from './imagen.js';
import { compuerta, siguienteFolio, avisoNeto, placaNormal, fechaMexico, horaMexico, slug, rolDe, PUEDE, lista, diasPara, evaluarVigencia, accionCorreccion, prealtaSinMovimiento, fechaCorta, aIsoDia, autoformatoFecha, plural, limpiar, paraPatch, tipoDeArchivo, lunesDe, sumarDias, esLoteDeLaApp, residuoDe, sufijoVerificacion, resumenCertificado, urlVerificacion, toneladas } from './reglas.js';

const VERSION = '0.35.1';   // la misma cadena va en package.json y en sw.js (CACHE); test/version.test.js lo exige
const $ = id => document.getElementById(id);
const L = CONFIG.listas;

const estado = {
    cuenta: null, token: null, cliente: null, siteId: null, rol: 'lectura',
    carriers: [], unidades: [], choferes: [], prealtas: [], embarques: [], vigencias: [], roles: [],
    firmas: [], firmasError: null,   // S-01 / S-07 (v0.25.0): renglones de PLANTA_Firmas; si la lista no se pudo leer, el motivo (y la app no firma ni autoriza)
    certificados: [], certificadosError: null,   // v0.35.0: renglones de PLANTA_Certificados; si la lista no existe todavia, el motivo (y no se emite)
    certificadoAbierto: null,   // el certificado pintado en dlgCertificado
    ultimaCompuerta: null,   // {resultado, hallazgos, campos}
    pesando: null,           // {embarque, fase: 'bruto'|'tara'}
    fotoBytes: null,
    fotoUrl: null,           // blob URL de la vista previa; se revoca al reemplazar la foto o cancelar (C-21)
    prealtaAbierta: null,
    prealtaEdit: null,       // borrador de pre-alta en edicion (paForma)
    padronEdit: null,        // {clave, x} del renglon del padron en edicion, o null
    focoAntesVeredicto: null, // elemento con el foco antes de abrir el veredicto (vuelve ahi al cerrarlo)
    padronFicha: null,       // 'tipo:id' de la ficha del padron desplegada
    padronCarrier: null,     // id del carrier que filtra unidades y choferes en el padron
    ultimoCarrierPadron: null,   // U-46: el ultimo carrier dado de alta o usado en un alta de unidad/chofer (solo esta sesion)
    pestana: 'hoy',
    archivos: null,          // v0.33.0: { biblioteca, ramas: Map ruta -> hijos } de la seccion Archivos; null = se relee al pintar
    cargadoEl: 0,
    ventanaDesde: ''      // ISO: inicio de la ventana de carga (cubeta 3)
};

// NO llamar `msal` a esta variable: taparia el global del bundle UMD.
const pca = new msal.PublicClientApplication({
    auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: new URL('./', window.location.href).href
    },
    // S-04: nada de cookies (storeAuthStateInCookie ya no existe en msal-browser 5.x; S-16 quito la opcion muerta).
    // sessionStorage sigue siendo decision cerrada.
    cache: { cacheLocation: 'sessionStorage' }
});

// ---------------------------------------------------------------- avisos y utilerias de DOM

// Un aviso REEMPLAZA al anterior: la pantalla dice el estado de la ultima accion, no la historia.
// Antes se apilaban (tres «Fecha no valida» del mismo intento, y el error seguia arriba del «dado de alta»; foto de Carlos 2026-09-06).
function avisar(texto, clase = '') {
    limpiarAvisos();
    const d = document.createElement('div');
    d.className = 'mensaje' + (clase ? ' ' + clase : '');
    d.textContent = texto;
    // U-08 (v0.22.0): con la pantalla de entrada visible el aviso va DENTRO de ella, sobre el boton. #avisos vive
    // arriba de una seccion de 100dvh en un body sin scroll: «No se pudo entrar» empujaba «Entrar» fuera de la vista.
    if (!$('pantallaEntrar').classList.contains('oculto')) { const z = $('entradaAviso'); z.textContent = texto; z.className = d.className; return; }
    $('avisos').setAttribute('aria-live', clase === 'error' ? 'assertive' : 'polite');   // U-32 (v0.26.0): el error interrumpe; lo demás espera su turno
    $('avisos').appendChild(d);
    // U-10 (v0.23.0): con dos formas abiertas (la del carrier encima de la pre-alta) el aviso va a la de ENCIMA, que es la
    // ultima abierta y, en el DOM, la ultima de las abiertas (las del padron van despues de #paForma).
    const abiertas = document.querySelectorAll('dialog.dlg-forma[open]');
    const dlg = abiertas[abiertas.length - 1];
    if (dlg) { const z = dlg.querySelector('.dlg-avisos'); z.textContent = ''; z.appendChild(d.cloneNode(true)); dlg.scrollTo({ top: 0, behavior: 'smooth' }); return; }
    // U-39 (v0.29.0): el veredicto es una <section> fija a pantalla completa (z-index 30) y tapa #avisos: «lleva motivo escrito»
    // y «No se pudo registrar» se pintan dentro de su cuerpo, junto a las acciones, y se hace scroll hasta ahí.
    if (!$('veredicto').classList.contains('oculto')) { const z = $('vkAvisos'); z.textContent = ''; z.appendChild(d.cloneNode(true)); z.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); return; }
    // Sin scrollTo (U-03, v0.21.0): #avisos es sticky y se ve donde este el usuario; el salto al tope alejaba
    // al basculista del formulario de pesaje con cada «Captura el peso» / «Falta la foto».
}
function limpiarAvisos() { $('avisos').textContent = ''; $('entradaAviso').textContent = ''; $('entradaAviso').className = 'mensaje oculto'; for (const z of document.querySelectorAll('.dlg-avisos')) z.textContent = ''; }   // U-39: incluye la del veredicto
// Los formularios de alta viven en <dialog> (v0.19.8): abrir es showModal, cerrar es close. Idempotentes.
function abrirForma(id) { const d = $(id); limpiarAvisos(); if (!d.open) d.showModal(); d.scrollTo({ top: 0 }); d.dataset.huella = huellaForma(d); }
/** Valores de la forma en una cadena (U-34, v0.26.0): se toma al abrir y Escape pregunta solo si cambió — un Editar sin tocar nada cierra directo. */
const huellaForma = d => JSON.stringify([...d.querySelectorAll('input:not([type=hidden]), textarea, select')].map(c => c.type === 'checkbox' || c.type === 'radio' ? c.checked : String(c.value)));
function cerrarForma(id) { const d = $(id); if (d.open) d.close(); }

function el(tag, clase, texto) {
    const e = document.createElement(tag);
    if (clase) e.className = clase;
    if (texto !== undefined) e.textContent = texto;
    return e;
}
/**
 * Renglon de lista con un boton principal y, opcionalmente, acciones secundarias
 * ({texto, alClic, accion, clase}) apiladas a la derecha; `accion` sale como data-accion.
 */
function renglon(titulo, sub, boton, alClic, extras = []) {
    const r = el('div', 'renglon');
    const izq = el('div');
    izq.appendChild(el('div', 't', titulo));
    if (sub) izq.appendChild(el('div', 's', sub));
    r.appendChild(izq);
    const botones = [];
    if (boton) botones.push({ texto: boton, alClic });
    botones.push(...extras);
    if (botones.length === 1) r.appendChild(botonAccion(botones[0]));
    else if (botones.length) { const d = el('div', 'acciones'); for (const b of botones) d.appendChild(botonAccion(b)); r.appendChild(d); }
    return r;
}
function botonAccion({ texto, alClic, accion, clase, deshabilitado }) {
    const b = el('button', clase || '', texto);
    b.type = 'button';
    if (accion) b.dataset.accion = accion;
    if (deshabilitado) { b.disabled = true; b.title = deshabilitado; }   // S-07: el motivo se lee al pasar el dedo / el raton
    b.addEventListener('click', alClic);
    return b;
}

/**
 * Confirmacion propia (sustituye al confirm() nativo). Devuelve {ok, motivo}. Con `motivo: true`
 * el motivo es obligatorio y el boton OK no procede sin el; con 'opcional' se muestra y no se exige.
 */
function confirmar({ titulo, texto, ok = 'Confirmar', motivo = false, etiquetaMotivo = 'Motivo', peligro = false }) {
    return new Promise(resolver => {
        const d = $('dlg');
        $('dlgTitulo').textContent = titulo;
        $('dlgTexto').textContent = texto || '';
        $('dlgOk').textContent = ok;
        $('dlgMotivoEtiqueta').textContent = etiquetaMotivo + (motivo === true ? ' · obligatorio' : '');
        $('dlgMotivoCampo').classList.toggle('oculto', !motivo);
        $('dlgMotivo').value = '';
        $('dlgError').classList.add('oculto');
        d.classList.toggle('peligro', peligro);
        const cerrar = ok2 => {
            $('dlgOk').onclick = $('dlgCancelar').onclick = d.onclose = null;
            if (d.open) d.close();
            resolver({ ok: ok2, motivo: $('dlgMotivo').value.trim() });
        };
        $('dlgOk').onclick = () => {
            // El aviso va en texto dentro del dialogo, no en la burbuja de validacion del navegador
            // (en el celular se pierde detras del teclado y en headless cerraba el dialogo).
            if (motivo === true && !$('dlgMotivo').value.trim()) { $('dlgError').classList.remove('oculto'); $('dlgMotivo').focus(); return; }
            cerrar(true);
        };
        $('dlgCancelar').onclick = () => cerrar(false);
        // Escape. Solo si el dialogo YA esta cerrado: el evento close del confirm ANTERIOR llega ENCOLADO (un tick despues
        // de su close()) y, si este confirm abrio en ese hueco, caia en este handler y lo cerraba sin respuesta — 2 de 8
        // corridas E2E de gerencia en rojo en pasos distintos, medido 2026-09-19 (v0.24.0) instrumentando close/showModal.
        d.onclose = () => { if (!d.open) cerrar(false); };
        $('dlgMotivo').oninput = () => $('dlgError').classList.add('oculto');
        d.showModal();
        if (motivo) $('dlgMotivo').focus();
    });
}
function etiqueta(texto, clase) { return el('span', 'etiqueta ' + (clase || ''), texto); }
/**
 * Un <div> que se abre al tocarlo (tarjeta de la fila del dia, renglon del padron) se vuelve alcanzable con
 * teclado y lector de pantalla (U-15, v0.22.0): role=button, tabindex=0, aria-expanded, y Enter/Espacio hacen
 * lo mismo que el clic. La tecla solo cuenta sobre el propio nodo: un Enter en un boton interior no lo abre.
 */
function desplegable(nodo, abierto, alClic) {
    nodo.setAttribute('role', 'button'); nodo.tabIndex = 0; nodo.setAttribute('aria-expanded', String(!!abierto));
    nodo.addEventListener('click', alClic);
    nodo.addEventListener('keydown', ev => { if (ev.target === nodo && (ev.key === 'Enter' || ev.key === ' ')) { ev.preventDefault(); alClic(ev); } });
}
/**
 * U-40 (v0.29.0): conserva el value anterior si sigue entre las opciones nuevas. Antes cada repintado (cambio de pestaña,
 * refresco de 2 minutos) rehacía el select en «— elige —» y la puerta perdía el programa elegido (los campos de texto sí
 * sobrevivían). Quien quiera empezar en blanco pone sel.value = '' después.
 */
function opciones(sel, items, valor, textoDe, primera = '— elige —') {
    const antes = sel.value;
    sel.textContent = '';
    const o = el('option', '', primera); o.value = ''; sel.appendChild(o);
    for (const it of items) {
        const op = el('option', '', textoDe(it)); op.value = String(valor(it)); sel.appendChild(op);
    }
    if (antes && [...sel.options].some(op => op.value === antes)) sel.value = antes;
}
// Fechas: dd/mm/aaaa en pantalla, ISO en SharePoint. C-25 (v0.28.0): fechaCorta, aIsoDia y el autoformato viven en reglas.js
// (puras, con casos en reglas.test.js); aqui solo queda el enganche al DOM.
// C-10 (v0.22.0): las cuatro horas de la app salen de reglas.horaMexico (hourCycle h23); antes tres usaban hour12:false
// (que en Chromium puede dar «24:05») y la fila del dia salia en 12 h con AM/PM: Hoy mezclaba los dos.
const horaCorta = iso => horaMexico(iso, 'fecha');
for (const inp of document.querySelectorAll('input.fecha')) inp.addEventListener('input', () => { inp.value = autoformatoFecha(inp.value); });
function porId(coleccion, id) { return coleccion.find(x => x.id === Number(id)) || null; }
/**
 * C-12 (v0.25.0): el objeto VIVO de estado[clave] con el id de `x`, resuelto AL CLIC. Los handlers de renglon capturan el
 * objeto de la lista con que se pinto; tras un refresco con captura a medias (que no repinta) ese objeto ya no esta en
 * estado.* y la tara / la anulacion / la firma se escribian sobre una copia rancia. Si ya no existe, se conserva el viejo
 * y la operacion lo reporta (el bruto y la tara releen el renglon antes de escribir).
 */
function vivo(clave, x) { return (x && porId(estado[clave], x.id)) || x; }
/**
 * C-23 (v0.28.0): un Object.assign DESPUES de un await cae sobre el objeto vivo por id, ademas del que el handler tenia en la
 * mano (C-12 resolvia al clic; el motivo de la anulacion se teclea 10-30 s y el refresco podia sustituir la lista debajo).
 * Devuelve el vivo.
 */
function aplicar(clave, obj, campos) { const v = vivo(clave, obj); Object.assign(obj, campos); if (v !== obj) Object.assign(v, campos); return v; }
/** C-23 / C-26: `x` ocupa su lugar en estado[clave] por id (sin duplicar el renglon si un refresco ya lo trajo); si no estaba, entra. */
function anclar(clave, x) { const i = estado[clave].findIndex(y => y.id === x.id); if (i >= 0) estado[clave][i] = x; else estado[clave].push(x); return x; }
/**
 * C-24 / C-23 (v0.28.0): toda escritura al tenant pasa por aqui. Deshabilita el boton que la disparo mientras dura (un segundo
 * toque con guante y senal lenta creaba dos carriers, dos firmas) y cuenta las escrituras en vuelo, que recargar() respeta.
 * `btn` es un id, un elemento o null (escrituras que no nacen de un boton fijo).
 */
let escrituras = 0;
async function escribiendo(btn, fn) {
    const b = typeof btn === 'string' ? $(btn) : btn;
    if (b && b.disabled) return undefined;
    if (b) b.disabled = true;
    escrituras++;
    try { return await fn(); }
    finally { escrituras--; if (b) b.disabled = false; }
}
/** C-26: un solo formateador para llenar formas (null/undefined -> ''). */
const textoDe = v => v === null || v === undefined ? '' : String(v);
/** C-26: el filtro del buscador (padron y cerrados): sin texto pega todo. */
const filtroTexto = q => (...campos) => !q || campos.some(v => normaliza(v).includes(q));
const haySel = sel => sel !== null && sel !== undefined;
function nombreDe(coleccion, id) { if (id === null || id === undefined || id === '') return '—'; const x = porId(coleccion, id); return x ? x.Title : `#${id}`; }

// ---------------------------------------------------------------- sesion

async function token() {
    const cuentas = pca.getAllAccounts();
    // S-13 (v0.30.0, msal-browser 5.x): sin caer al iframe oculto. En 5.x el iframe/popup exige una pagina «redirect bridge»
    // que esta app no tiene (su redirectUri es index.html, que procesa la respuesta con handleRedirectPromise como en 4.x);
    // sin bridge, la renovacion por iframe muere en `redirect_bridge_timeout` (BrowserAuthError, no InteractionRequired) y el
    // catch de refrescarCliente() no la mandaria al login. Con AccessTokenAndRefreshToken MSAL usa cache y refresh token y,
    // si no alcanzan, lanza InteractionRequiredAuthError (no_tokens_found / refresh_token_expired) -> acquireTokenRedirect.
    const r = await pca.acquireTokenSilent({ scopes: CONFIG.scopes, account: cuentas[0], cacheLookupPolicy: msal.CacheLookupPolicy.AccessTokenAndRefreshToken });
    return r.accessToken;
}
let msalListo = false;
async function prepararMsal() {
    if (msalListo) return null;
    await pca.initialize();
    const respuesta = await pca.handleRedirectPromise();
    msalListo = true;
    return respuesta;
}
// U-20 (v0.22.0): el progreso del inicio de sesion se escribe en el PROPIO boton (con anillo girando), no solo en la
// linea mono de 9.5 px al pie: con la senal lenta de la caseta parecia que no habia pasado nada y se volvia a tocar.
function pasoEntrada(texto) {
    const b = $('btnEntrar');
    b.disabled = !!texto; b.classList.toggle('ocupado', !!texto);
    b.textContent = texto || 'Entrar con cuenta MINSA';
    $('textoEntrar').textContent = texto || 'CALYTEK · Planta';
}
async function entrar() {
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
async function arrancar() {
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
async function salir() {
    // Salir esta a un toque en la barra del celular: se confirma para no cerrar la sesion con la gondola esperando (F5).
    const { ok } = await confirmar({ titulo: 'Salir de la app', ok: 'Salir', texto: 'Se cierra la sesión de MINSA en este dispositivo. Lo capturado ya está guardado; lo que esté a medias en pantalla se pierde.' });
    if (!ok) return;
    try { await pca.logoutRedirect({ account: estado.cuenta }); }
    catch (_) { sessionStorage.clear(); window.location.reload(); }
}
async function refrescarCliente() {
    try {
        estado.token = await token();
    } catch (e) {
        // La sesion silenciosa caduca (horas en la caseta): MSAL pide interaccion. Se manda al
        // login por redireccion en vez de dejar un "No se pudo guardar" que nadie sabe resolver.
        const pideInteraccion = (typeof msal !== 'undefined' && msal.InteractionRequiredAuthError && e instanceof msal.InteractionRequiredAuthError)
            || (e && e.errorCode === 'interaction_required');
        if (pideInteraccion) { avisar('La sesión caducó: volviendo a entrar…', 'ojo'); await pca.acquireTokenRedirect({ scopes: CONFIG.scopes, account: pca.getAllAccounts()[0] }); }
        throw e;
    }
    estado.cliente = crearCliente(CONFIG.graph, estado.token);
}
/**
 * U-51 / U-47 (v0.29.0): el rail y el menú «···» dicen QUIÉN entró por su nombre (quien(): PLANTA_Roles, luego la cuenta MSAL,
 * luego el correo) y el correo queda en title; antes el rail partía el correo a media palabra y el celular no lo decía en ningún lado.
 */
function ponerQuien(correo, rol = '') {
    const nombre = quien(correo);
    for (const q of document.querySelectorAll('.quien')) { q.textContent = ''; q.appendChild(el('b', '', nombre)); if (rol) q.appendChild(el('span', 'rol', rol)); q.title = correo; }   // v0.32.0: dos renglones, como el rail de Proyectos
    $('rolMovil').textContent = rol;
    $('quienMovil').textContent = nombre; $('quienMovil').title = correo;
    $('correoMovil').textContent = nombre === correo ? '' : correo;
}
// I5 (7-sep): cuando se leyeron las listas por ultima vez, en la barra movil y en el rail. Ambar pasados 5 minutos.
let avisoReintento = false;
function pintarSync(leyendo = false, texto = 'Leyendo las listas…') {
    const t = Date.now() - estado.cargadoEl;
    const hace = !estado.cargadoEl ? '' : t < 60000 ? `hace ${Math.max(1, Math.round(t / 1000))} s` : t < 3600000 ? `hace ${Math.round(t / 60000)} min` : `hace ${Math.round(t / 3600000)} h`;
    const dice = leyendo ? texto : estado.cargadoEl ? `Al día · leído ${hace}` : '';
    for (const x of document.querySelectorAll('.sync')) {
        // U-63 (v0.34.0): el punto del renglón de sesión del rail no lleva texto (la hora sigue dentro del «···», decisión cerrada)
        if (x.classList.contains('sync-punto')) { x.title = dice; x.setAttribute('aria-label', dice); } else x.textContent = dice;
        x.classList.toggle('viejo', !leyendo && t > 300000);
        x.classList.toggle('leyendo', leyendo);
    }
    // U-63: en escritorio el reintento de C-17 quedaba detrás del «···»; se dice también en #avisos y se limpia al terminar de leer
    if (leyendo && texto !== 'Leyendo las listas…') { avisar(texto, 'ojo'); avisoReintento = true; }
    else if (!leyendo && avisoReintento) { avisoReintento = false; limpiarAvisos(); }
}
setInterval(() => { if (estado.siteId) pintarSync(recargando); }, 15000);

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
    // Pantalla inicial por rol: quien captura abre en la puerta, quien firma en las pre-altas,
    // gerencia y lectura en «Hoy» (la consola). Igual en celular y en computadora.
    irA(estado.rol === 'trazabilidad' ? 'puerta' : estado.rol === 'validador' ? 'prealtas' : 'hoy');
}

/** Un embarque esta "en planta" si paso la compuerta (o le autorizaron la excepcion) y no ha cerrado.
 *  Es LA definicion: la usan Bascula, Hoy y las insignias, para que los tres numeros coincidan. */
function enPlanta(e) {
    return (e.Etapa === 'compuerta' && (e.Compuerta === 'pasa' || excepcionAutorizada(e)))
        || e.Etapa === 'bruto';
}
function excepcionesPendientes() {
    return estado.embarques.filter(e => e.Etapa === 'compuerta' && e.Compuerta === 'excepcion-comercial' && !excepcionAutorizada(e));
}

/**
 * S-01, mitad del tenant (v0.24.0): la firma de la pre-alta y la autorizacion de la excepcion viven en PLANTA_Firmas,
 * una lista SIN herencia donde solo validador + gerencia escriben (setup-carlos.md, tarea 11). La COMPUERTA MANDA
 * (decision de Carlos, 2026-09-19): sin renglon ahi, la pre-alta no sale en la puerta y la excepcion no pasa a bascula,
 * aunque el sello (FirmadaPor / ExcepcionAutorizo) este escrito. S-07 (v0.25.0): la compuerta falla CERRADA. Con la lista
 * ya provisionada (tarea 11), si PLANTA_Firmas no se puede leer —no aparece en /lists, 404, 403— la app NO firma ni
 * autoriza y no acepta ningun sello sin firma; «Hoy» lo dice en Pendiente revisar. Antes volvia al «sello solo» y una
 * cuenta a la que le quitaran Leer en la lista corria como antes de S-01. Otro error (red, 5xx) sube y la recarga falla entera.
 */
async function cargarFirmas(c, s, avisar) {
    try { const f = await c.renglones(s, L.firmas, null, avisar); estado.firmasError = null; return f; }
    catch (e) {
        if (!(e && (e.status === 404 || e.status === 403))) throw e;   // C-15: por status (idDeLista tipa su 404), no por texto
        estado.firmasError = String(e.message); return [];
    }
}
/**
 * C-15 (v0.27.0): «la lista no tiene todavia esa columna / esa opcion» se decide por el 400 de Graph (invalidRequest:
 * «Field 'X' is not recognized», «not a valid option»), no por regex sobre el texto. Un 401, 403, 404 o 5xx NO es una
 * columna faltante y sube tal cual: antes «invalid» casaba con InvalidAuthenticationToken y «no existe» con el 404 del sitio.
 */
function esColumnaFaltante(e) { return !!e && e.status === 400; }
/** v0.35.0: PLANTA_Certificados puede no existir todavia (provisionar.html la crea): la app sigue, sin emitir, y lo dice a gerencia. */
async function cargarCertificados(c, s, avisar) {
    try { const x = await c.renglones(s, L.certificados, null, avisar); estado.certificadosError = null; return x; }
    catch (e) {
        if (!(e && (e.status === 404 || e.status === 403))) throw e;
        estado.certificadosError = String(e.message); return [];
    }
}
/**
 * S-11 (v0.28.0): un renglon de PLANTA_Firmas solo vale si su Firmante tiene HOY en PLANTA_Roles el rol que el Tipo exige
 * (excepcion -> gerencia; prealta -> validador o gerencia) Y es la misma cuenta que el sello (ExcepcionAutorizo / FirmadaPor).
 * Planta-Firmantes junta validador y gerencia con Colaborar: antes un validador podia crear por Graph una firma de excepcion
 * y la app la aceptaba. PLANTA_Roles es de solo lectura para Miembros, asi que nadie se da el rol solo.
 */
const ROL_FIRMA = { excepcion: ['gerencia'], prealta: ['validador', 'gerencia'], certificado: ['gerencia'] };   // certificado: v0.35.0
const mismaCuenta = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
function firmaDe(tipo, id, sello) {
    return estado.firmas.find(f => f.Tipo === tipo && Number(f.ObjetoId) === Number(id)
        && (ROL_FIRMA[tipo] || []).includes(rolDe(f.Firmante, estado.roles))
        && (sello === undefined || mismaCuenta(f.Firmante, sello))) || null;
}
function prealtaFirmada(p) { return p.Estado === 'firmada' && !!firmaDe('prealta', p.id, p.FirmadaPor); }
function excepcionAutorizada(e) { return !!e.ExcepcionAutorizo && !!firmaDe('excepcion', e.id, e.ExcepcionAutorizo); }
/** Sellos sin firma que valga (de antes del corte, por fuera de la app, o firmada por quien no tiene el rol): la compuerta no los acepta. */
function sellosSinFirma() {
    return { prealtas: estado.prealtas.filter(p => p.Estado === 'firmada' && !firmaDe('prealta', p.id, p.FirmadaPor)),
             embarques: estado.embarques.filter(e => e.Etapa === 'compuerta' && !!e.ExcepcionAutorizo && !firmaDe('excepcion', e.id, e.ExcepcionAutorizo)) };
}
const selloSinFirma = e => (e.ExcepcionAutorizo ? `sello de ${quien(e.ExcepcionAutorizo)} sin firma registrada · ` : '');   // U-28 / U-31 (v0.26.0)
/** S-07: por que la app no puede firmar ni autorizar ahora mismo, o null. Va como `title` del boton deshabilitado. */
const motivoSinFirmas = () => (estado.firmasError ? `No se pudo leer el registro de firmas: no se firma ni se autoriza hasta que se vea. Actualiza; si sigue, avisa a gerencia.${estado.rol === 'gerencia' ? ` (PLANTA_Firmas: ${estado.firmasError})` : ''}` : null);   // U-31: lenguaje de planta primero; el detalle técnico solo a gerencia
/** El renglon de firma se escribe PRIMERO: para quien no esta en el grupo de firmantes es un 403, y ahi termina. */
async function firmar(tipo, objeto, motivo) {
    if (estado.firmasError) throw new Error(motivoSinFirmas());
    const f = await estado.cliente.crearRenglon(estado.siteId, L.firmas, limpiar({
        Title: `${tipo} · ${objeto.Title || objeto.PlacaTractor || objeto.id}`, Tipo: tipo, ObjetoId: objeto.id,
        Firmante: estado.cuenta.username, FirmadoEl: new Date().toISOString(), Motivo: motivo || null }));
    estado.firmas.push(f);
    return f;
}
function pintarInsignias() {
    const b = estado.embarques.filter(enPlanta).length;
    const p = estado.prealtas.filter(x => x.Estado === 'borrador').length;
    $('nBascula').textContent = String(b); $('nBascula').hidden = b === 0;
    $('nPrealtas').textContent = String(p); $('nPrealtas').hidden = p === 0 || !PUEDE.firmarPrealta(estado.rol);
    // D2 (2026-09-08): carril Puerta -> Bascula -> Ticket en «Hoy». Puerta = gondolas registradas hoy (sin las anuladas),
    // Bascula = las que estan en planta (misma definicion que la insignia), Ticket = las cerradas hoy.
    const flujo = $('flujoHoy');
    if (flujo) {
        const hoy = fechaMexico(new Date());
        const deHoy = estado.embarques.filter(e => e.Etapa !== 'anulado' && e.Arribo && fechaMexico(new Date(e.Arribo)) === hoy);
        const pon = (k, v) => { const b = flujo.querySelector(`[data-e="${k}"] b`); if (b) b.textContent = String(v); };
        pon('puerta', deHoy.length);
        pon('bascula', b);
        pon('ticket', deHoy.filter(e => e.Etapa === 'cerrado').length);
    }
}

// ---------------------------------------------------------------- carga acotada (cubeta 3)

/** Los embarques que la app considera "abiertos": no han cerrado, sin importar su edad. */
// 'descargando' se retiro el 2026-09-07 (F6, decision de Carlos): nada la escribia. La opcion sigue en la lista de SharePoint
// (provisionar solo agrega); un renglon viejo con ese valor ya no cuenta como abierto.
const FILTRO_ABIERTOS = "fields/Etapa eq 'compuerta' or fields/Etapa eq 'bruto'";
const iso = d => new Date(d).toISOString();

/**
 * Embarques de los ultimos CONFIG.ventanaDias MAS todos los abiertos (union, sin repetir).
 * La segunda consulta no es un lujo: una gondola registrada hace cuatro meses y nunca cerrada
 * saldria de la ventana y desapareceria de Bascula › En planta con la fosa todavia ocupada.
 * Las dos columnas del filtro (Arribo, Etapa) estan indexadas (esquema.json, cubeta 1).
 */
async function cargarEmbarques(c, s, avisar) {
    const desde = iso(Date.now() - CONFIG.ventanaDias * 86400000);
    estado.ventanaDesde = desde;
    const [recientes, abiertos] = await Promise.all([
        c.renglones(s, L.embarques, `fields/Arribo ge '${desde}'`, avisar),
        c.renglones(s, L.embarques, FILTRO_ABIERTOS, avisar)
    ]);
    const porId = new Map();
    for (const e of recientes) porId.set(e.id, e);
    for (const e of abiertos) if (!porId.has(e.id)) porId.set(e.id, e);
    return [...porId.values()];
}

/**
 * Embarques del ANIO en curso: la unica base valida para elegir folio. La ventana de 90 dias no
 * sirve aqui — en abril dejaria fuera enero y `siguienteFolio` reiniciaria el consecutivo sobre
 * numeros ya emitidos. El folio SI se reinicia cada anio, y eso lo hace `siguienteFolio` solo
 * (su regex lleva el AA), asi que el anio en curso es exactamente lo que hay que mirar.
 */
async function embarquesDelAno(avisar) {
    return await estado.cliente.renglones(estado.siteId, L.embarques,
        `fields/Arribo ge '${iso(new Date(new Date().getFullYear(), 0, 1))}'`, avisar);
}

/** Mete en la lista de la ventana los renglones frescos que le correspondan (por id). */
function fundirEnVentana(frescos) {
    const porId = new Map(estado.embarques.map(e => [e.id, e]));
    for (const f of frescos) {
        const v = porId.get(f.id);
        if (v) Object.assign(v, f);
        else if (f.Arribo && f.Arribo >= estado.ventanaDesde) estado.embarques.push(f);
    }
}

async function cargarTodo() {
    const c = estado.cliente, s = estado.siteId;
    // C-17 (v0.25.0): graph.js reintenta 429/503/red caida hasta 5.6 s; antes nadie recibia el aviso y la pantalla se
    // quedaba en «Leyendo las listas…». La franja de sync dice que esta reintentando.
    const av = texto => pintarSync(true, texto);
    [estado.carriers, estado.unidades, estado.choferes, estado.prealtas, estado.embarques, estado.vigencias, estado.roles, estado.firmas, estado.certificados] =
        await Promise.all([
            c.renglones(s, L.carriers, null, av), c.renglones(s, L.unidades, null, av), c.renglones(s, L.choferes, null, av),
            c.renglones(s, L.prealtas, null, av), cargarEmbarques(c, s, av), c.renglones(s, L.vigencias, null, av),
            c.renglones(s, L.roles, null, av), cargarFirmas(c, s, av), cargarCertificados(c, s, av)
        ]);
    estado.cargadoEl = Date.now();
    reanclar();
}

/**
 * C-01 (v0.21.0): lo que la pantalla tiene «en la mano» —el embarque que se esta pesando, la pre-alta
 * abierta o en edicion, el renglon del padron en edicion— apunta a un OBJETO de la lista anterior.
 * cargarTodo() sustituye las listas enteras, asi que tras un refresco silencioso (cada 2 min, al volver
 * del bolsillo) ese objeto queda huerfano: la tara se guardaba sobre el, pintarBascula leia la lista
 * nueva y la gondola seguia «en planta» (un segundo toque subia otro lote). Aqui se vuelve a apuntar
 * por id al objeto fresco; si ya no esta (lo borro otra sesion), se conserva el viejo y el guardado
 * lo reporta como hoy.
 */
function reanclar() {
    const fresco = (col, x) => (x && porId(col, x.id)) || x;
    if (estado.pesando) estado.pesando.embarque = fresco(estado.embarques, estado.pesando.embarque);
    estado.prealtaAbierta = fresco(estado.prealtas, estado.prealtaAbierta);
    estado.prealtaEdit = fresco(estado.prealtas, estado.prealtaEdit);
    if (estado.padronEdit) estado.padronEdit.x = fresco(estado[estado.padronEdit.clave], estado.padronEdit.x);
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
    if (abierto('baPesar') || abierto('veredicto')) return true;
    if (['paForma', 'paDetalle', 'pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer'].some(id => $(id).open)) return true;
    if (estado.pestana === 'puerta' && ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre', 'puMotivo', 'puPrealta'].some(id => $(id).value.trim())) return true;   // U-40: el programa elegido también es captura
    return false;
}

let recargando = false;
async function recargar(silencioso = false) {
    if (recargando || !estado.siteId) return;
    // C-23 (v0.28.0): con una escritura en vuelo o un confirm abierto NO se sustituyen las listas: el Object.assign que sigue
    // al await caeria sobre un objeto huerfano (la gondola anulada seguia «En planta», gerencia autorizaba dos veces). Antes
    // solo se frenaba el REPINTADO (capturaAMedias); cargarTodo() corria igual. El timer lo vuelve a intentar al minuto.
    if (escrituras > 0 || $('dlg').open) { if (!silencioso && escrituras > 0) avisar('Espera a que termine de guardar y vuelve a actualizar.', 'ojo'); return; }
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

const PINTORES = { hoy: () => pintarHoy(), puerta: () => pintarPuerta(), bascula: () => pintarBascula(), prealtas: () => pintarPrealtas(), padron: () => pintarPadron(), reportes: () => pintarReportes(), archivos: () => pintarArchivos() };
const SECCIONES = Object.keys(PINTORES);
/** Los botones del rail: las pestañas de siempre más la sección aparte (Reportes, v0.32.0). */
const botonesRail = () => [...$('pestanas').querySelectorAll('button'), ...$('pestanasExtra').querySelectorAll('button')];
/** Repinta la pestana abierta SIN tocar avisos, veredicto ni scroll (refresco silencioso y cambios de pre-alta). */
function repintar() { pintarInsignias(); PINTORES[estado.pestana](); }
/**
 * C-27 / U-41 (v0.28.0): salir de la bascula con un pesaje abierto pasa por la misma compuerta que «Cancelar» (U-24): con kg o
 * foto pregunta, y al salir suelta estado.pesando, la foto y su blob URL. Antes irA() ocultaba la tarjeta sin preguntar y el
 * basculista volvia a fotografiar el indicador; el JPEG comprimido quedaba vivo hasta el siguiente pesaje. Devuelve si se salio.
 */
const pesajeConAlgo = () => !$('baPesar').classList.contains('oculto') && !!($('baKg').value.trim() || estado.fotoBytes);
function descartarPesaje() { estado.pesando = null; estado.fotoBytes = null; soltarFotoPrevia(); $('baPesar').classList.add('oculto'); }
async function soltarPesaje() {
    if (pesajeConAlgo()) {
        const { ok } = await confirmar({ titulo: 'Cancelar el pesaje', peligro: true, ok: 'Descartar', texto: 'Se pierden el peso tecleado y la foto del indicador; habría que volver a tomarla.' });
        if (!ok) return false;
    }
    descartarPesaje();
    return true;
}
/** El clic en una pestana: sincrono salvo que haya que preguntar (la E2E y el usuario esperan la pestana pintada al soltar). */
function irDesdePestana(p) {
    if (p !== 'bascula' && pesajeConAlgo()) { soltarPesaje().then(ok => { if (ok) irA(p); }); return; }
    if (p !== 'bascula' && !$('baPesar').classList.contains('oculto')) descartarPesaje();   // vacio: se suelta sin preguntar, como Cancelar
    irA(p);
}
function irA(p) {
    estado.pestana = p;
    for (const b of botonesRail()) { if (b.dataset.p === p) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); }   // U-57 (v0.29.0): <nav> con aria-current, como .mn-rail de la piel; el role=tablist prometía flechas y tabpanel que no había
    for (const s of SECCIONES) $('p-' + s).classList.toggle('oculto', s !== p);
    limpiarAvisos();
    cerrarVeredicto();
    repintar();
    window.scrollTo({ top: 0 });
}

// ================================================================ PUERTA

function pintarPuerta() {
    const firmadas = estado.prealtas.filter(prealtaFirmada);
    // Aviso informativo: pre-altas en borrador (sin firma) — la puerta no las ve en el selector hasta que se firmen.
    // S-01: la que trae sello pero no renglon en PLANTA_Firmas cuenta igual: la compuerta manda.
    const borradores = estado.prealtas.filter(p => p.Estado === 'borrador' || (p.Estado === 'firmada' && !prealtaFirmada(p)));
    const pp = $('puPendientes'); pp.classList.toggle('oculto', !borradores.length);
    if (borradores.length) pp.textContent = `${borradores.length === 1 ? 'Hay 1 pre-alta por firmar' : `Hay ${borradores.length} pre-altas por firmar`}: ${borradores.map(p => p.Title).join(' · ')}. Sus góndolas no pueden entrar hasta que el validador firme.`;
    opciones($('puPrealta'), firmadas, p => p.id, p => `${p.Title} · ${p.Corriente || '?'} · ${nombreDe(estado.carriers, p.CarrierId)}`);
    if (!$('puPrealta').value && firmadas.length === 1) $('puPrealta').value = String(firmadas[0].id);   // U-40: una sola firmada no se hace elegir
    pintarChoferesPuerta();
    pintarUnidadesPuerta();
    $('btnCompuerta').disabled = !PUEDE.puerta(estado.rol);
    $('puSoloLectura').classList.toggle('oculto', PUEDE.puerta(estado.rol));   // U-38 (v0.26.0): texto fijo, no avisar(): el refresco silencioso lo repetía cada 2 min y pisaba el aviso que se leía
    $('puSoloLectura').textContent = estado.rol === 'validador' ? 'Tu rol es validador: aquí solo ves; capturas y firmas en Pre-altas.' : 'Tu rol es de lectura: puedes ver, no capturar.';   // U-50: al validador no se le dice «lectura»
    pintarPrevioPuerta();
}
// I6 (7-sep): las unidades que la pre-alta autorizo (o, si no marco ninguna, todas las activas del carrier).
// Tocar una llena placa tractor y plana; la seleccion se marca comparando con lo que hay en el campo, asi que
// teclear otra placa la desmarca sola. La compuerta sigue evaluando la placa del campo, no la seleccion.
function pintarUnidadesPuerta() {
    const cont = $('puUnidades'); cont.textContent = '';
    const pre = porId(estado.prealtas, $('puPrealta').value);
    if (!pre) { cont.appendChild(el('p', 'pista', 'Elige el programa para ver sus unidades.')); return; }
    const autorizadas = lista(pre.UnidadesIds).map(Number);
    let us = estado.unidades.filter(u => u.Activo !== false && Number(u.CarrierId) === Number(pre.CarrierId));
    if (autorizadas.length) us = us.filter(u => autorizadas.includes(u.id));
    $('puUnidadesTitulo').textContent = us.length ? `Unidad del carrier · ${us.length} en este programa` : 'Unidad del carrier';
    if (!us.length) { cont.appendChild(el('p', 'pista', 'Este programa no tiene unidades en el padrón: teclea la placa y saldrá como no amparada.')); return; }
    const actual = placaNormal($('puPlaca').value);
    for (const u of us) {
        const b = el('button', 'u' + (placaNormal(u.Title) === actual ? ' sel' : '')); b.type = 'button';
        b.setAttribute('aria-pressed', String(placaNormal(u.Title) === actual));   // U-33 (v0.26.0): el estado seleccionado existe para teclado y lector, no solo como clase
        b.dataset.placa = placaNormal(u.Title);   // U-17: la seleccion al teclear compara contra esto, no contra el texto del chip
        const t = el('span', 't', u.PlacaPlana ? `${u.Title} / ${u.PlacaPlana}` : u.Title);
        const v = etiquetaVigencia('unidades', u);
        t.appendChild(el('small', '', [u.TipoUnidad || 'unidad', u.CapacidadKg ? `${Number(u.CapacidadKg).toLocaleString('es-MX')} kg` : null, v ? null : 'vigencias al día'].filter(Boolean).join(' · ')));
        b.appendChild(t);
        if (v) b.appendChild(v);
        b.addEventListener('click', () => {
            $('puPlaca').value = u.Title; $('puPlacaPlana').value = u.PlacaPlana || '';
            marcarChip(cont, b.dataset.placa);   // C-28
            pintarPrevioPuerta();
        });
        cont.appendChild(b);
    }
}
/** C-28 (v0.28.0): el chip de unidad elegido, por clase Y aria-pressed, igual al clic que al teclear la placa (U-17 / U-33). */
function marcarChip(cont, placa) {
    for (const x of cont.querySelectorAll('.u')) { const on = !!placa && x.dataset.placa === placa; x.classList.toggle('sel', on); x.setAttribute('aria-pressed', String(on)); }
}
function pintarChoferesPuerta() {
    const pre = porId(estado.prealtas, $('puPrealta').value);
    const carrierId = pre ? Number(pre.CarrierId) : null;
    const ch = estado.choferes.filter(x => x.Activo !== false && (!carrierId || Number(x.CarrierId) === carrierId));
    // U-21 (v0.26.0): como los chips de unidad — primero los que la pre-alta autorizó; los demás del carrier siguen
    // eligibles pero marcados «fuera del programa», porque la compuerta los saca ámbar (reglas.js «Chofer vs pre-alta»).
    const autorizados = pre ? lista(pre.ChoferesIds).map(String) : [];
    const fuera = x => autorizados.length > 0 && !autorizados.includes(String(x.id));
    const orden = ch.filter(x => !fuera(x)).concat(ch.filter(fuera));
    opciones($('puChofer'), orden, x => x.id, x => fuera(x) ? `${x.Title} · fuera del programa` : x.Title, '— no está en el padrón —');
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

function correrCompuerta() {
    limpiarAvisos();
    // U-01 (v0.21.0): mientras falte un campo de CAMPOS_PUERTA el boton dice «Faltan N datos» y NO abre el
    // veredicto: enfoca el primer faltante. Antes con manifiesto o corriente vacios abria «No entra» y un toque
    // mas dejaba un folio R- por un dato que no se habia tecleado (reglas.js los cuenta como hallazgo legal).
    const falta = CAMPOS_PUERTA.find(([id]) => !String($(id).value).trim());
    if (falta) { avisar(`Falta ${falta[1]}.`, 'error'); $(falta[0]).focus(); return; }
    const e = evaluarPuerta();
    if (!e.placa) { avisar('Falta la placa del tractor.', 'error'); $('puPlaca').focus(); return; }
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
function pintarPrevioPuerta() {
    const e = evaluarPuerta();
    const faltan = CAMPOS_PUERTA.filter(([id]) => !String($(id).value).trim()).map(([, n]) => n);

    // U-48 (v0.29.0): cada dato del bloque es una lista de alternativas (el chofer vale por el select O por el nombre en
    // licencia). El bloque 2 contaba solo la placa y decía «Listo» mientras la lista viva anunciaba Espera por el chofer;
    // el chofer sigue sin ser obligatorio para correr la compuerta (CAMPOS_PUERTA no cambia).
    const hay = alts => alts.some(id => String($(id).value).trim());
    const bloques = [['puBloque1', 'puEst1', [['puPrealta', 'programa']], 'sin'],
                     ['puBloque2', 'puEst2', [['puPlaca', 'la placa'], ['puChofer', 'puChoferNombre', 'el chofer']], 'falta'],
                     ['puBloque3', 'puEst3', [['puManifiesto', 'el manifiesto'], ['puCorriente', 'la corriente']], 'falta']];
    for (const [bloque, est, datos, pendiente] of bloques) {
        const faltantes = datos.filter(d => !hay(d.slice(0, -1))).map(d => d[d.length - 1]);
        const ok = !faltantes.length, n = faltantes.length;
        $(bloque).classList.toggle('listo', ok);
        // D2 (2026-09-08): el estado cuenta lo que falta; el detalle («falta la placa») queda en el title y en el texto.
        $(est).textContent = ok ? 'Listo' : n === 1 ? `Falta 1 · ${faltantes[0]}` : `Faltan ${n}`;
        $(est).title = ok ? '' : `${pendiente} ${faltantes.join(' y ')}`;
    }

    const decide = e.resultado === 'rechazo-legal' ? 'legal' : e.resultado === 'excepcion-comercial' ? 'comercial' : null;
    pintarHallazgos($('puPrevioReglas'), e.hallazgos, faltan.length ? null : decide);

    // D2 (2026-09-08): el boton dice «Faltan 2 datos · el manifiesto, la corriente» hasta que todo esta capturado.
    // Sigue siendo el mismo boton y sigue corriendo la compuerta: solo cambia lo que dice.
    const btn = $('btnCompuerta');
    btn.classList.toggle('incompleto', faltan.length > 0);
    btn.textContent = '';
    if (faltan.length) {
        btn.appendChild(document.createTextNode(faltan.length === 1 ? 'Falta 1 dato' : `Faltan ${faltan.length} datos`));
        btn.appendChild(el('small', '', faltan.join(', ')));
    } else btn.textContent = 'Correr la compuerta';

    const cuenta = $('puPrevioCuenta');
    cuenta.classList.remove('v-ok', 'v-warn', 'v-bad');
    $('puPrevio').classList.toggle('incompleto', faltan.length > 0);
    if (faltan.length) { cuenta.textContent = `Falta ${faltan.join(', ')}. Hasta entonces esto no es un pronóstico.`; return; }
    const def = VEREDICTOS[e.resultado];
    cuenta.classList.add(def.clase);
    const porque = decide ? ` — ${e.hallazgos.filter(h => h.clase === decide).map(h => h.regla.toLowerCase()).join(', ')}` : '';
    cuenta.textContent = '';
    cuenta.appendChild(document.createTextNode('Con lo capturado va a salir '));
    cuenta.appendChild(el('b', '', def.palabra));
    cuenta.appendChild(document.createTextNode(porque + '.'));
}

/**
 * EL SEMAFORO. La compuerta tiene tres salidas y la pantalla se ve distinta en cada una: el color
 * cubre la cabecera, la palabra es enorme y la regla que decidio va primero y en negrita. Un chofer
 * a tres metros sabe si entra sin leer nada (mockup aprobado 2026-09-05).
 */
const VEREDICTOS = {
    pasa: { clase: 'v-ok', palabra: 'Pasa', frase: 'Todo lo legal está amparado. Que suba a la báscula.', boton: 'Registrar: la góndola pasa a báscula' },
    'rechazo-legal': { clase: 'v-bad', palabra: 'No entra', frase: 'Rechazo legal. El residuo no se recibe y no hay dispensa.', boton: 'Registrar el rechazo' },
    'excepcion-comercial': { clase: 'v-warn', palabra: 'Espera', frase: 'Excepción comercial. Puede entrar si gerencia lo autoriza con motivo.', boton: 'Registrar y pedir autorización' }
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

function cerrarVeredicto() {
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
    $('vkW').textContent = def.palabra;
    $('vkM').textContent = def.frase;
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
    $('vkLbl').textContent = decide ? 'Qué lo decidió' : `${r.hallazgos.length} reglas · ${plural(r.hallazgos.filter(h => h.clase === 'aviso').length, 'aviso')}`;

    $('puExcepcion').classList.toggle('oculto', r.resultado !== 'excepcion-comercial');
    const guion = $('vkGuion');
    guion.classList.toggle('oculto', r.resultado !== 'rechazo-legal');
    if (r.resultado === 'rechazo-legal') {
        const culpables = r.hallazgos.filter(h => h.clase === 'legal').map(h => h.regla.toLowerCase());
        guion.lastElementChild.textContent = `Falta lo legal (${culpables.join(', ')}). Se corrige en el oficio o en la pre-alta, no aquí. Se lleva su constancia con el folio R-.`;
    }
    $('btnRegistrarPuerta').textContent = def.boton;
    v.scrollTop = 0;
    // Dialogo modal: el foco entra al veredicto y sale con Esc (F3). El boton principal recibe el foco.
    estado.focoAntesVeredicto = document.activeElement;
    $('btnRegistrarPuerta').focus();
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

async function registrarPuerta() {
    const r = estado.ultimaCompuerta; if (!r) return;
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
        const campos = limpiar({
            Title: folio, Etapa: esRechazo ? 'rechazado' : 'compuerta',
            PreAltaId: r.pre ? r.pre.id : null, Manifiesto: r.campos.manifiesto, Arribo: ahora,
            CarrierId: r.carrier ? r.carrier.id : null, UnidadId: r.unidad ? r.unidad.id : null,
            PlacaTractor: placaNormal(r.campos.placaTractor), PlacaPlana: placaNormal(r.campos.placaPlana),
            ChoferId: r.chofer ? r.chofer.id : null, ChoferNombre: r.campos.choferNombre,
            CorrienteDeclarada: r.campos.corriente, Compuerta: r.resultado,
            CompuertaDetalle: JSON.stringify(r.hallazgos), Verificacion79: !!r.campos.art79,
            ExcepcionMotivo: r.resultado === 'excepcion-comercial' ? $('puMotivo').value.trim() : null,
            CapturadoPor: estado.cuenta.username
        });
        const nuevo = await estado.cliente.crearRenglon(estado.siteId, L.embarques, campos, av);
        if (esRechazo) await asegurarFolioUnico(nuevo, 'R', av);
        anclar('embarques', nuevo);   // C-23: sin duplicar el id si el refresco ya lo trajo
        avisar(esRechazo ? `Rechazo registrado con folio ${nuevo.Title}. La góndola no entra.` :
            r.resultado === 'pasa' ? 'Registrado. Ya aparece en Báscula › En planta.' :
            'Registrado. Gerencia lo ve en Báscula › Excepciones por autorizar.', 'bien');
        for (const id of ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre', 'puMotivo']) $(id).value = '';
        $('pu79').checked = false;
        // U-02 (v0.21.0): tambien el chofer (la siguiente gondola heredaba su ChoferId) y se repintan chips y vista
        // previa: antes los bloques seguian en «Listo» con los campos vacios. El programa y la corriente se quedan:
        // las gondolas del mismo programa llegan en fila.
        $('puChofer').value = ''; delete $('puChoferNombre').dataset.auto;
        cerrarVeredicto();
        pintarInsignias();
        estado.ultimaCompuerta = null;
        pintarUnidadesPuerta(); pintarPrevioPuerta();
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
async function asegurarFolioUnico(renglon, tipo, avisar) {
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

// ================================================================ BASCULA

function pintarBascula() {
    // Las excepciones por autorizar ya no viven aqui: estan en «Hoy», que es donde gerencia mira.
    const lista2 = $('baLista'); lista2.textContent = '';
    const activos = estado.embarques.filter(enPlanta);
    if (!activos.length) lista2.appendChild(el('p', 'pista', 'Nada en planta.'));
    for (const e of activos) {
        const faseBruto = e.Etapa === 'compuerta';
        lista2.appendChild(renglon(`${e.Title || '(sin folio)'} · ${e.PlacaTractor}${e.Manifiesto ? ' · ' + e.Manifiesto : ''}`,
            faseBruto ? `Pasó la compuerta ${horaCorta(e.Arribo)} · falta el BRUTO` : `Bruto ${e.BrutoKg} kg ${horaCorta(e.BrutoHora)} · falta la TARA`,
            PUEDE.puerta(estado.rol) ? (faseBruto ? 'Pesar bruto' : 'Pesar tara') : null,
            () => abrirPesaje(vivo('embarques', e), faseBruto ? 'bruto' : 'tara'),
            botonCorreccion(e)));
    }
    $('baPesar').classList.add('oculto');
    $('baTicketCaja').classList.add('oculto');
    pintarCerrados();
}

/**
 * Cerrados (2026-09-08): todo embarque con folio que ya no esta en planta — cerrado, anulado o
 * rechazado — dentro de la ventana cargada, del mas reciente al mas viejo, con su Ticket para
 * volver a verlo o reimprimirlo (el chofer que vuelve por el ticket perdido, gerencia cotejando un
 * neto de la semana pasada). No es archivo: el completo vive en SharePoint, y la linea «Emitido»
 * del ticket lleva la hora de la reimpresion. Fuera de la ventana no se busca (decision 8-sep:
 * se agrega cuando aparezca el primer caso real).
 */
function pintarCerrados() {
    const caja = $('baCerrados'); caja.textContent = '';
    const q = normaliza($('baBusca').value.trim());
    const pega = filtroTexto(q);   // C-26
    const momento = e => e.TaraHora || e.AnuladoEl || e.Arribo || '';
    const todos = estado.embarques.filter(e => e.Title && !enPlanta(e)).sort((a, b) => momento(b).localeCompare(momento(a)));
    const cerrados = todos.filter(e => e.Etapa === 'cerrado');
    const res = $('baResCerrados'); res.textContent = '';
    res.appendChild(document.createTextNode(`${cerrados.length} con neto${todos.length - cerrados.length ? ` · ${plural(todos.length - cerrados.length, 'anulado o rechazado', 'anulados o rechazados')}` : ''} · últimos ${CONFIG.ventanaDias} días`));
    let n = 0;
    const visibles = [];
    for (const e of todos) {
        if (!pega(e.Title, e.PlacaTractor, e.PlacaPlana, e.Manifiesto, nombreDe(estado.carriers, e.CarrierId))) continue;
        n++;
        // Solo el titular por hilera (Carlos, 8-sep): el detalle vive en el ticket. Ticket abre la
        // ventana sobre la lista tal como quedó filtrada, para recorrerla con ‹ ›.
        visibles.push(e);
        const i = visibles.length - 1;
        const r = renglon(`${e.Title} · ${e.PlacaTractor || ''}${e.Manifiesto ? ' · ' + e.Manifiesto : ''}`, null, 'Ticket',
            () => abrirTicketPop(visibles, i));
        if (e.Etapa !== 'cerrado') { r.classList.add(e.Etapa); r.firstChild.firstChild.appendChild(etiqueta(e.Etapa, e.Etapa)); }
        caja.appendChild(r);
    }
    if (!n) caja.appendChild(el('p', 'pista', q ? 'Nada coincide en los últimos ' + CONFIG.ventanaDias + ' días.' : 'Ningún folio cerrado en la ventana cargada.'));
    $('baBuscaCuenta').textContent = q ? plural(n, 'resultado') : '';
}

/** El boton de deshacer que le toca a un embarque segun su etapa (reglas.accionCorreccion), o ninguno. */
function botonCorreccion(e) {
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
    await escribiendo(btn, async () => {   // C-24 / C-23: boton deshabilitado y sin refresco mientras el confirm esta abierto
    const { ok, motivo } = await confirmar({
        titulo: `Anular ${e.Title}`, peligro: true, ok: 'Anular con este motivo', motivo: true,
        texto: `${e.PlacaTractor} · etapa ${e.Etapa}${e.NetoKg ? ` · neto ${e.NetoKg} kg` : e.BrutoKg ? ` · bruto ${e.BrutoKg} kg` : ''}. El folio ${e.Title} queda anulado y NO se vuelve a usar; el siguiente pesaje nace con folio nuevo. Las fotos ya subidas se conservan.`,
        etiquetaMotivo: 'Motivo de la anulación'
    });
    if (!ok) return;
    try {
        await refrescarCliente();
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

async function autorizarExcepcion(e, btn) {
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

function abrirPesaje(e, fase) {
    estado.pesando = { embarque: e, fase };
    estado.fotoBytes = null; soltarFotoPrevia();
    $('baPesar').classList.remove('oculto');
    $('baTicketCaja').classList.add('oculto');
    // U-44 (v0.29.0): las dos fases titulan con la PLACA (lo que el basculista ve en el camión) y la tara además con el folio;
    // el manifiesto va al subtítulo. Antes la tara decía solo «Tara · E-26-00004» y el bruto no traía el manifiesto.
    $('baTitulo').textContent = fase === 'bruto' ? `Bruto · ${e.PlacaTractor}` : `Tara · ${e.PlacaTractor} · ${e.Title}`;
    const u = porId(estado.unidades, e.UnidadId);
    const manif = e.Manifiesto ? `Manifiesto ${e.Manifiesto}. ` : '';
    $('baSub').textContent = fase === 'bruto' ? `${manif}Primera pasada: la góndola cargada. Al guardar nace el folio E-.` :
        `${manif}Segunda pasada: la góndola vacía. Bruto ${e.BrutoKg} kg${u && u.CapacidadKg ? ` · capacidad ${u.CapacidadKg} kg` : ''}.`;
    $('baKgLabel').textContent = fase === 'bruto' ? 'Bruto (kg)' : 'Tara (kg)';
    $('baKg').value = '';
    $('baFotoPrevia').classList.add('oculto');
    $('baFotoEstado').textContent = 'Sin foto todavía'; $('baFotoEstado').classList.remove('lista');
    $('btnFoto').textContent = 'Tomar foto';
    $('baNetoVivo').classList.toggle('oculto', fase !== 'tara');
    $('btnGuardarPeso').textContent = fase === 'bruto' ? 'Guardar bruto' : 'Guardar tara';
    if (fase === 'tara') {
        $('baNvBruto').textContent = Number(e.BrutoKg).toLocaleString('es-MX');
        $('baNvCap').textContent = u && u.CapacidadKg ? Number(u.CapacidadKg).toLocaleString('es-MX') : '—';
        revisarNeto();
    }
    $('baAvisoNeto').classList.add('oculto');
    $('baMotivoNetoCampo').classList.add('oculto');
    $('baMotivoNeto').value = '';
    $('baPesar').scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (ESCRITORIO.matches) $('baKg').focus();   // U-53: con ratón el foco va a los kilos; en celular el teclado es propio (inputmode none)
}

/** Revoca el blob URL de la vista previa (C-21, v0.25.0): cada foto comprimida quedaba viva hasta recargar la PWA. */
function soltarFotoPrevia() {
    if (estado.fotoUrl) { URL.revokeObjectURL(estado.fotoUrl); estado.fotoUrl = null; }
}
async function tomarFoto(archivo) {
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

function revisarNeto() {
    const p = estado.pesando; if (!p || p.fase !== 'tara') return null;
    const u = porId(estado.unidades, p.embarque.UnidadId);
    const aviso = avisoNeto(p.embarque.BrutoKg, $('baKg').value, u ? u.CapacidadKg : null, CONFIG.tolerancia);
    $('baAvisoNeto').classList.toggle('oculto', !aviso);
    $('baMotivoNetoCampo').classList.toggle('oculto', !aviso);
    if (aviso) $('baAvisoNeto').textContent = 'Revisa el indicador: ' + aviso;
    // I2: el neto se ve conforme se teclea; ambar si se sale de la banda; el boton dice lo que va a cerrar.
    const tara = Number($('baKg').value), bruto = Number(p.embarque.BrutoKg);
    const neto = Number.isFinite(tara) && tara > 0 ? bruto - tara : null;
    const nv = $('baNvNeto'); nv.textContent = neto === null ? '—' : neto.toLocaleString('es-MX');
    nv.parentElement.className = neto === null ? '' : aviso ? 'neto-mal' : 'neto-ok';
    $('btnGuardarPeso').textContent = neto === null || neto <= 0 ? 'Guardar tara' : `Guardar tara · neto ${neto.toLocaleString('es-MX')} kg`;
    return aviso;
}

async function guardarPeso() {
    const p = estado.pesando; if (!p) return;
    const kg = Number($('baKg').value);
    if (!Number.isFinite(kg) || kg <= 0) { avisar('Captura el peso en kilogramos.', 'error'); return; }
    if (!estado.fotoBytes) { avisar('Falta la foto del indicador: es lo que hace comprobable un peso tecleado.', 'error'); return; }
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
        $('baPesar').classList.add('oculto'); soltarFotoPrevia();
        pintarTicket(e);
        pintarBascula();
        $('baTicketCaja').classList.remove('oculto');
        $('baTicketCaja').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
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
    if (e.Etapa !== 'compuerta') throw new Error(`este embarque ya está en ${e.Etapa} (lo movió otra sesión). Actualiza la lista.`);
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
    if (e.Etapa !== 'bruto') throw new Error(`este embarque ya está en ${e.Etapa} (lo movió otra sesión). Actualiza la lista.`);
    paso('Subiendo la foto…');
    const lote = await subirEvidencia(e.Title, 'tara', kg, paso);
    paso('Cerrando el embarque…');
    const neto = Number(e.BrutoKg) - kg;
    const campos = limpiar({ Etapa: 'cerrado', TaraKg: kg, TaraHora: ahora, TaraFoto: lote.ref, NetoKg: neto,
        InicioAlmacen: e.BrutoHora || ahora,
        Notas: aviso ? `Neto fuera de banda (${aviso}). Motivo: ${$('baMotivoNeto').value.trim()}` : null });
    await guardarConLote(e, lote, campos, paso);
    aplicar('embarques', e, campos);
    avisar(`Embarque ${e.Title} cerrado: neto ${neto} kg.`, 'bien');
}

/**
 * La foto del indicador va al buzon como un lote del contrato de nombres, con `_lote.json`
 * (ticket 06: la app apunta, no guarda). Devuelve el par (fecha, concepto) que el embarque conserva.
 */
/** `<destinoBase>/<AAAA>/<AAAA-MM>` a partir de la fecha del lote (AAAA-MM-DD). Carpeta por mes: ver config.js. */
export function destinoEvidencia(fecha) {
    return `${CONFIG.evidencia.destinoBase}/${fecha.slice(0, 4)}/${fecha.slice(0, 7)}`;
}

async function subirEvidencia(folio, fase, kg, avisar) {
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
function abrirTicketPop(lista, i) {
    const d = $('dlgTicket');
    const pintar = () => {
        pintarTicket(lista[i], $('ticketPop'));
        $('tkPos').textContent = `${i + 1} de ${lista.length} · ${lista[i].Title}`;
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

/**
 * Nombre y apellido de una cuenta (Carlos, 8-sep): el renglon guarda el CORREO (auditable, es lo que
 * SharePoint conoce); en pantalla se muestra el `Nombre` de PLANTA_Roles, y si esa fila no lo trae,
 * el nombre que da Entra para la cuenta activa, y al final el correo tal cual.
 */
function quien(correo) {
    if (!correo) return '';
    const r = estado.roles.find(x => String(x.Title || '').toLowerCase() === String(correo).toLowerCase());
    if (r && r.Nombre) return r.Nombre;
    if (estado.cuenta && estado.cuenta.username && estado.cuenta.username.toLowerCase() === String(correo).toLowerCase() && estado.cuenta.name) return estado.cuenta.name;
    return correo;
}

function pintarTicket(e, t = $('ticket')) {
    t.textContent = '';
    t.appendChild(el('h3', '', 'CALYTEK · Planta de tratamiento de RME — Ticket de báscula'));
    const pre = porId(estado.prealtas, e.PreAltaId);
    const filas = [
        ['Folio', e.Title], ['Manifiesto', e.Manifiesto || '—'],
        ['Generador', pre ? `${pre.Generador || ''} ${pre.GeneradorRegistro ? '(' + pre.GeneradorRegistro + ')' : ''}` : '—'],
        ['Pozo / corriente', pre ? `${pre.Pozo || '—'} · ${e.CorrienteDeclarada || pre.Corriente || ''}` : (e.CorrienteDeclarada || '—')],
        ['Transportista', nombreDe(estado.carriers, e.CarrierId)],
        ['Placas', `${e.PlacaTractor || ''} / ${e.PlacaPlana || ''}`], ['Operador', e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId)],
        ['Arribo', horaCorta(e.Arribo)],
        ['Bruto', e.BrutoKg ? `${e.BrutoKg} kg · ${horaCorta(e.BrutoHora)}` : '—'],
        ['Tara', e.TaraKg ? `${e.TaraKg} kg · ${horaCorta(e.TaraHora)}` : 'pendiente'],
        ['NETO', e.NetoKg ? `${e.NetoKg} kg` : 'pendiente'],
        ['Capturó', quien(e.CapturadoPor)], ['Emitido', `${horaCorta(new Date().toISOString())} · CALYTEK Planta ${VERSION}`]
    ];
    if (e.Etapa === 'anulado') filas.unshift(['ANULADO', `${horaCorta(e.AnuladoEl)} · ${quien(e.AnuladoPor)} · ${e.AnuladoMotivo || ''}`]);
    const tabla = el('table');
    for (const [k, v] of filas) {
        const tr = el('tr'); tr.appendChild(el('td', '', k)); tr.appendChild(el('td', k === 'NETO' ? 'grande' : 'mono', v)); tabla.appendChild(tr);
    }
    t.appendChild(tabla);
}

// ================================================================ PRE-ALTAS

/**
 * Gondolas de una pre-alta: recibidas contra esperadas. Recibida = embarque que existe y llego;
 * el anulado (folio quemado) y el rechazado (no entro) no cuentan.
 */
function gondolasDe(p) {
    const rec = estado.embarques.filter(e => Number(e.PreAltaId) === p.id && e.Etapa !== 'anulado' && e.Etapa !== 'rechazado').length;
    return { rec, esp: Number(p.GondolasEsperadas) || 0 };
}
function sinMovimientoDe(p) { return prealtaSinMovimiento(p, estado.embarques, CONFIG.sinMovimientoDias); }
function barraAvance({ rec, esp }) {
    const d = el('div', 'avance');
    if (esp) {
        const b = el('span', 'barra' + (rec ? '' : ' esp'));
        const i = el('i'); i.style.width = Math.min(100, Math.round(rec / esp * 100)) + '%';
        b.appendChild(i); d.appendChild(b);
        d.appendChild(el('span', 'cifra', `${rec}/${esp}`));
    } else d.appendChild(el('span', 'cifra', `${plural(rec, 'recibida')} · sin estimado`));
    return d;
}
function pintarPrealtas() {
    cerrarForma('paForma'); cerrarForma('paDetalle');
    $('btnNuevaPrealta').classList.toggle('oculto', !PUEDE.capturarPrealta(estado.rol));
    // Un Estado que no sea uno de los tres cae en cerradas para que no desaparezca de la vista;
    // ahi su etiqueta lo delata (abajo se pinta cuando no coincide con el grupo).
    const grupos = { borrador: [], firmada: [], cerrada: [] };
    for (const p of [...estado.prealtas].sort((a, b) => b.id - a.id)) (grupos[p.Estado] || grupos.cerrada).push(p);

    const pintar = (cont, ps, grupo) => {
        cont.textContent = '';
        for (const p of ps) {
            const g = gondolasDe(p);
            const r = renglon(p.Title, `${p.Generador || '?'} · ${p.Pozo || '?'} · ${p.Corriente || '?'} · ${nombreDe(estado.carriers, p.CarrierId)}${grupo === 'borrador' ? ` · 1er envío ${fechaCorta(p.FechaEstimada)}` : ''}`, 'Ver', () => verPrealta(vivo('prealtas', p)));
            r.classList.add('conavance');
            if (p.Campana) r.firstChild.firstChild.appendChild(el('span', 'folio', p.Campana));
            if (p.Estado !== grupo) r.firstChild.firstChild.appendChild(etiqueta(p.Estado || 'sin estado', p.Estado));
            if (grupo === 'firmada' && !prealtaFirmada(p)) r.firstChild.firstChild.appendChild(etiqueta('sin firma', 'vencida'));   // S-01
            const sm = grupo === 'firmada' ? sinMovimientoDe(p) : null;
            if (sm) { r.firstChild.firstChild.appendChild(etiqueta('¿se cierra?', 'aviso')); r.firstChild.appendChild(el('p', 'pista', `${sm.motivo}. Sigue saliendo en la puerta hasta que alguien cierre el programa.`)); }
            r.firstChild.appendChild(barraAvance(g));
            cont.appendChild(r);
        }
    };
    pintar($('paBorradores'), grupos.borrador, 'borrador');
    pintar($('paFirmadas'), grupos.firmada, 'firmada');
    pintar($('paCerradas'), grupos.cerrada, 'cerrada');

    if (!estado.prealtas.length) $('paBorradores').appendChild(el('p', 'pista', 'No hay pre-altas. La primera góndola no puede entrar sin una firmada.'));

    const suma = ps => ps.reduce((a, p) => { const g = gondolasDe(p); a.rec += g.rec; a.esp += g.esp; return a; }, { rec: 0, esp: 0 });
    const sb = suma(grupos.borrador), sf = suma(grupos.firmada), sc = suma(grupos.cerrada);
    $('paResBorradores').textContent = grupos.borrador.length
        ? `${grupos.borrador.length} · ${plural(sb.esp, 'góndola comprometida', 'góndolas comprometidas')}, ninguna puede entrar`
        : 'ninguno pendiente de firma';
    $('paResFirmadas').textContent = grupos.firmada.length
        ? `${grupos.firmada.length} · ${sf.rec} de ${sf.esp} góndolas recibidas`
        : 'ninguna · la puerta no puede recibir';
    $('paResCerradas').textContent = grupos.cerrada.length
        ? `${grupos.cerrada.length} · ${plural(sc.rec, 'góndola')} en total`
        : 'ninguna';
    // U-23 (v0.26.0): Firmadas se queda visible aunque esté vacío — su resumen «la puerta no puede recibir» es el aviso
    // que importa justo cuando no hay ninguna. Solo Cerradas se oculta vacío.
    $('paGrupoFirmadas').classList.remove('oculto');
    $('paGrupoCerradas').classList.toggle('oculto', !grupos.cerrada.length);
}

function nuevaPrealta() {
    estado.prealtaEdit = null;
    $('paFormaTitulo').textContent = 'Nueva pre-alta'; $('btnGuardarPrealta').textContent = 'Guardar como borrador';
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false), c => c.id, c => c.Title);
    $('paCarrier').value = '';   // U-40: opciones() ya conserva el value; una pre-alta nueva empieza sin carrier
    pintarUnidadesChoferesPrealta();
    for (const id of ['paTitulo', 'paCliente', 'paPozoTitulo', 'paMes', 'paGenerador', 'paGeneradorRegistro', 'paPozo', 'paFecha', 'paGondolas', 'paCorreoFecha', 'paCorreoRemitente', 'paNotas']) $(id).value = '';
    $('paCorriente').value = '';
    $('paMes').value = String(new Date().getFullYear());   // U-54: casi siempre es el año en curso
    armarTituloPrealta();
    pintarEstadoPrealta();
    abrirForma('paForma');
}
/**
 * Editar un BORRADOR (v0.19.10, Carlos 2026-09-08: «ver el borrador no me da la opcion de editarlo»). Mismo
 * formulario que la alta, con `estado.prealtaEdit`; guardar es un PATCH que conserva Estado, Campana y quien capturo.
 * Una firmada no se edita: la puerta ya la usa; se cierra y se abre otra.
 */
function editarPrealta() {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'borrador' || !PUEDE.capturarPrealta(estado.rol)) return;
    estado.prealtaEdit = p;
    $('paFormaTitulo').textContent = `Editar pre-alta: ${p.Title}`; $('btnGuardarPrealta').textContent = 'Guardar cambios';
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false || Number(c.id) === Number(p.CarrierId)), c => c.id, c => c.Title);
    const f = textoDe;   // C-26
    partirTituloPrealta(f(p.Title)); $('paCorriente').value = f(p.Corriente); $('paGenerador').value = f(p.Generador);
    $('paGeneradorRegistro').value = f(p.GeneradorRegistro); $('paPozo').value = f(p.Pozo); $('paCarrier').value = f(p.CarrierId);
    $('paFecha').value = p.FechaEstimada ? fechaCorta(p.FechaEstimada) : ''; $('paGondolas').value = f(p.GondolasEsperadas);
    $('paCorreoFecha').value = p.CorreoFecha ? fechaCorta(p.CorreoFecha) : ''; $('paCorreoRemitente').value = f(p.CorreoRemitente); $('paNotas').value = f(p.Notas);
    pintarUnidadesChoferesPrealta();
    // Se marcan las que la pre-alta ya tenia; si no tenia ninguna guardada, quedan todas (es lo que la puerta entiende).
    for (const [cont, ids] of [['paUnidades', lista(p.UnidadesIds)], ['paChoferes', lista(p.ChoferesIds)]])
        if (ids.length) for (const c of $(cont).querySelectorAll('input')) c.checked = ids.includes(c.value);
    armarTituloPrealta();
    pintarEstadoPrealta();
    cerrarForma('paDetalle');
    abrirForma('paForma');
}
/**
 * Nombre del programa = CLIENTE-POZO-AÑO, en mayusculas (el 3er campo era «mes» hasta la v0.19.18; id paMes se conserva) (Carlos, 2026-09-08: «GSM-IXACHI 15-2026»). Las tres partes se
 * capturan por separado y paTitulo (oculto) se arma solo; los guiones dentro de una parte se cambian por espacio para
 * que el nombre se pueda volver a partir al editar. El pozo del titulo rellena «Pozo / instalacion» si esta vacio.
 */
const PARTES_TITULO = ['paCliente', 'paPozoTitulo', 'paMes'];
const parteTitulo = id => $(id).value.trim().toUpperCase().replace(/\s*-\s*/g, ' ').replace(/\s+/g, ' ');
function armarTituloPrealta() {
    const partes = PARTES_TITULO.map(parteTitulo);
    const t = partes.every(Boolean) ? partes.join('-') : '';
    $('paTitulo').value = t; $('paTituloVista').textContent = t || '—';
    const pozo = $('paPozo'); if (partes[1] && (!pozo.value || pozo.value === pozo.dataset.auto)) { pozo.value = partes[1]; pozo.dataset.auto = partes[1]; }
}
function partirTituloPrealta(titulo) {
    const p = String(titulo || '').split('-').map(x => x.trim());
    const partes = p.length >= 3 ? [p[0], p.slice(1, -1).join(' '), p[p.length - 1]] : [titulo || '', '', ''];
    PARTES_TITULO.forEach((id, i) => { $(id).value = partes[i]; });
}
function pintarUnidadesChoferesPrealta() {
    const cid = Number($('paCarrier').value);
    const cajas = (cont, items, nombre) => {
        cont.textContent = '';
        if (!items.length) { cont.appendChild(el('p', 'pista', `(ese carrier no tiene ${nombre} en el padrón)`)); return; }
        for (const it of items) {
            const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; c.value = String(it.id); c.checked = true;
            l.appendChild(c); l.appendChild(document.createTextNode(' ' + (it.PlacaPlana ? `${it.Title} / ${it.PlacaPlana}` : it.Title))); cont.appendChild(l);
        }
    };
    cajas($('paUnidades'), estado.unidades.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'unidades');
    cajas($('paChoferes'), estado.choferes.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'choferes');
}
function marcados(id) { return [...$(id).querySelectorAll('input:checked')].map(c => c.value).join(';'); }
// I4 (7-sep): cada bloque de la pre-alta dice si esta completo o que le falta, igual que la puerta.
// Solo el 1 y el 3 tienen obligatorios (los mismos que valida guardarPrealta); el 2 y el 4 son opcionales y lo dicen.
const BLOQUES_PREALTA = [
    ['paBloque1', 'paEst1', ['paCliente', 'paPozoTitulo', 'paMes', 'paCorriente'], [], 'faltan cliente, pozo, año o corriente'],
    ['paBloque2', 'paEst2', [], ['paGenerador', 'paGeneradorRegistro', 'paPozo'], 'opcional'],
    ['paBloque3', 'paEst3', ['paCarrier'], [], 'falta el carrier'],
    ['paBloque4', 'paEst4', [], ['paFecha', 'paGondolas', 'paCorreoFecha', 'paCorreoRemitente'], 'opcional'],
];
function pintarEstadoPrealta() {
    const lleno = id => String($(id).value).trim() !== '';
    for (const [bloque, est, oblig, opc, pendiente] of BLOQUES_PREALTA) {
        const ok = oblig.every(lleno);
        const n = opc.filter(lleno).length;
        $(bloque).classList.toggle('listo', ok && (!opc.length || n > 0));
        $(est).textContent = !ok ? pendiente : opc.length ? (n ? `${n} de ${opc.length}` : pendiente) : 'completo';
    }
}
for (const id of PARTES_TITULO) $(id).addEventListener('input', armarTituloPrealta);
$('paForma').addEventListener('input', pintarEstadoPrealta);
$('paForma').addEventListener('change', pintarEstadoPrealta);

async function guardarPrealta() {
    if (!PUEDE.capturarPrealta(estado.rol)) { avisar('Tu rol no captura pre-altas.', 'error'); return; }
    if (!$('paTitulo').value.trim() || !$('paCorriente').value || !$('paCarrier').value) { avisar('Faltan cliente, pozo o año del programa, la corriente o el carrier.', 'error'); return; }
    // U-54 (v0.29.0): el año del nombre es un número de cuatro cifras entre el año pasado y el que viene; antes «202» o «x»
    // se quedaban en el título que la puerta ve en el select y en el ticket.
    const anio = Number($('paMes').value.trim()), hoyAnio = new Date().getFullYear();
    if (!/^\d{4}$/.test($('paMes').value.trim()) || anio < hoyAnio - 1 || anio > hoyAnio + 1) { avisar(`El año del programa va con cuatro cifras, entre ${hoyAnio - 1} y ${hoyAnio + 1}.`, 'error'); $('paMes').focus(); return; }
    const edit = estado.prealtaEdit && estado.prealtaEdit.Estado === 'borrador' ? estado.prealtaEdit : null;
    await escribiendo('btnGuardarPrealta', async () => { try {   // C-24
        await refrescarCliente();
        if (edit) {
            const cambios = paraPatch({
                Title: $('paTitulo').value.trim(), Generador: $('paGenerador').value.trim(), GeneradorRegistro: $('paGeneradorRegistro').value.trim(),
                Pozo: $('paPozo').value.trim(), Corriente: $('paCorriente').value, CarrierId: Number($('paCarrier').value),
                UnidadesIds: marcados('paUnidades'), ChoferesIds: marcados('paChoferes'),
                FechaEstimada: aIsoDia($('paFecha').value), GondolasEsperadas: $('paGondolas').value ? Number($('paGondolas').value) : null,
                CorreoFecha: aIsoDia($('paCorreoFecha').value), CorreoRemitente: $('paCorreoRemitente').value.trim(), Notas: $('paNotas').value.trim()
            });
            await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, edit.id, cambios);
            aplicar('prealtas', edit, cambios);   // C-23
            estado.prealtaEdit = null;
            cerrarForma('paForma');
            if (estado.pestana === 'prealtas') pintarPrealtas(); else pintarInsignias();
            verPrealta(edit);
            avisar('Borrador actualizado. Sigue pendiente de firma.', 'bien');   // al final: abrir el detalle limpia los avisos
            return;
        }
        // C-04 (v0.22.0): la campana se escoge con la lista RELEIDA, no con la de memoria (que puede tener horas), y tras
        // escribir pasa por asegurarFolioUnico como R- y E-. Dos capturas en dos celulares recibian la misma L-AA-NNN.
        const todas = await estado.cliente.renglones(estado.siteId, L.prealtas);
        estado.prealtas = todas; reanclar();
        const campos = limpiar({
            Title: $('paTitulo').value.trim(), Estado: 'borrador', Generador: $('paGenerador').value.trim(),
            GeneradorRegistro: $('paGeneradorRegistro').value.trim(), Pozo: $('paPozo').value.trim(),
            Corriente: $('paCorriente').value, CarrierId: Number($('paCarrier').value),
            UnidadesIds: marcados('paUnidades'), ChoferesIds: marcados('paChoferes'),
            FechaEstimada: aIsoDia($('paFecha').value), GondolasEsperadas: $('paGondolas').value ? Number($('paGondolas').value) : null,
            CorreoFecha: aIsoDia($('paCorreoFecha').value), CorreoRemitente: $('paCorreoRemitente').value.trim(),
            CapturadaPor: estado.cuenta.username, Notas: $('paNotas').value.trim(),
            Campana: siguienteFolio('L', todas.map(p => p.Campana))   // L-AA-NNN (ticket 03)
        });
        const nuevo = await estado.cliente.crearRenglon(estado.siteId, L.prealtas, campos);
        await asegurarFolioUnico(nuevo, 'L');
        anclar('prealtas', nuevo);   // C-23
        cerrarForma('paForma');
        avisar('Pre-alta guardada como borrador. Falta la firma del validador.', 'bien');
        pintarPrealtas();
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}

// Tras firmar / cerrar / eliminar: se cierra el pop-up y se repinta la pestana que esta abierta (Hoy o Pre-altas) sin
// borrar el aviso. Antes saltaba a Pre-altas aunque se hubiera abierto desde Hoy (Carlos, 2026-09-08).
function trasCambioPrealta() {
    cerrarForma('paDetalle'); repintar();
}
function verPrealta(p) {
    estado.prealtaAbierta = p;
    abrirForma('paDetalle');   // antes de los avisos: con el dialog abierto, avisar() los pinta adentro
    $('paDetalleTitulo').textContent = `${p.Title} · ${p.Estado}`;
    const ul = $('paDetalleLista'); ul.textContent = '';
    const carrier = porId(estado.carriers, p.CarrierId);
    const filas = [
        ['Generador', `${p.Generador || '—'} · ${p.GeneradorRegistro || 'sin registro'}`], ['Pozo', p.Pozo || '—'], ['Corriente', p.Corriente || '—'],
        ['Campaña', p.Campana || '—'], ['Carrier', carrier ? `${carrier.Title} · ${carrier.AutorizacionASEA || 'sin autorización'}` : '—'],
        ['Unidades', lista(p.UnidadesIds).map(id => { const u = porId(estado.unidades, id); return u ? `${u.Title}/${u.PlacaPlana || ''}` : `#${id}`; }).join(', ') || '—'],
        ['Choferes', lista(p.ChoferesIds).map(id => nombreDe(estado.choferes, id)).join(', ') || '—'],
        ['Primer envío', fechaCorta(p.FechaEstimada)], ['Correo', `${fechaCorta(p.CorreoFecha)} · ${p.CorreoRemitente || ''}`],
        ['Capturó', quien(p.CapturadaPor) || '—'], ['Firmó', p.FirmadaPor ? `${quien(p.FirmadaPor)} · ${horaCorta(p.FirmadaEl)}${p.Estado === 'firmada' && !prealtaFirmada(p) ? ' · sello sin firma: la puerta no la ve' : ''}` : '—'], ['Cerró', p.CerradaPor ? `${quien(p.CerradaPor)} · ${horaCorta(p.CerradaEl)}` : '—'], ['Notas', p.Notas || '—']   // U-28 / U-31 (v0.26.0)
    ];
    for (const [k, v] of filas) { const li = el('li', '', k); li.appendChild(el('span', 'd', v)); ul.appendChild(li); }
    // Cotejo automatico de vigencias (lo que el validador firma que reviso).
    const vg = $('paDetalleVigencias'); vg.textContent = '';
    const hallazgos = [];
    if (carrier) {
        const c = evaluarVigencia('Autorización ASEA del carrier', carrier.VigenciaASEA, 'legal', CONFIG.avisoVigenciaDias); if (c) hallazgos.push(c);
        if (p.Corriente && lista(carrier.Corrientes).length && !lista(carrier.Corrientes).includes(p.Corriente)) hallazgos.push({ clase: 'legal', regla: 'Corriente', detalle: `el oficio del carrier no ampara ${p.Corriente}` });
    } else hallazgos.push({ clase: 'legal', regla: 'Carrier', detalle: 'sin carrier' });
    for (const id of lista(p.UnidadesIds)) { const u = porId(estado.unidades, id); if (!u) continue;
        if (!u.FolioOficio && !(carrier && carrier.FolioOficio)) hallazgos.push({ clase: 'legal', regla: `Unidad ${u.Title}`, detalle: 'sin folio de oficio que la ampare (ni en la unidad ni en el carrier)' });
        for (const [n, f] of [['tarjeta', u.TarjetaVigencia], ['póliza', u.PolizaVigencia]]) { const h = evaluarVigencia(`Unidad ${u.Title} · ${n}`, f, 'comercial', CONFIG.avisoVigenciaDias); if (h) hallazgos.push(h); } }
    for (const id of lista(p.ChoferesIds)) { const ch = porId(estado.choferes, id); if (!ch) continue;
        const h = evaluarVigencia(`Chofer ${ch.Title} · licencia`, ch.LicenciaVigencia, 'comercial', CONFIG.avisoVigenciaDias); if (h) hallazgos.push(h); }
    if (!hallazgos.length) { const li = el('li', '', 'Vigencias '); li.appendChild(etiqueta('todo vigente', 'ok')); vg.appendChild(li); }
    for (const h of hallazgos) { const li = el('li', '', h.regla + ' '); li.appendChild(etiqueta(h.clase, h.clase)); li.appendChild(el('span', 'd', h.detalle)); vg.appendChild(li); }
    const hayLegal = hallazgos.some(h => h.clase === 'legal');
    const porFirmar = p.Estado === 'borrador' || (p.Estado === 'firmada' && !prealtaFirmada(p));   // S-01: el sello sin firma se firma aqui mismo
    $('btnFirmar').classList.toggle('oculto', !(porFirmar && PUEDE.firmarPrealta(estado.rol)));
    $('btnFirmar').title = motivoSinFirmas() || '';   // S-07
    $('btnEditarPrealta').classList.toggle('oculto', !(p.Estado === 'borrador' && PUEDE.capturarPrealta(estado.rol)));
    $('btnFirmar').disabled = hayLegal || !!estado.firmasError;   // S-07 + hallazgo legal (C-26: antes se asignaba dos veces)
    if (hayLegal && porFirmar) avisar('No se puede firmar con un hallazgo legal abierto: corrige el padrón (con el oficio a la vista) o cambia el carrier.', 'ojo');
    else if (porFirmar && p.Estado === 'firmada') avisar('Falta la firma del validador: trae el sello pero no la firma registrada (se escribió por fuera de la app o antes del corte). La puerta no la ve hasta que un validador o gerencia la firme.', 'ojo');   // U-31
    $('btnCerrarPrealta').classList.toggle('oculto', !(p.Estado === 'firmada' && PUEDE.capturarPrealta(estado.rol)));
    const sm = sinMovimientoDe(p);
    if (sm) avisar(`Este programa ${sm.motivo}. Sigue saliendo en la puerta hasta que se cierre; si ya no vienen más góndolas, ciérralo.`, 'ojo');
    // Un borrador equivocado se elimina; una firmada ya la vio la puerta y solo se CIERRA (2026-09-05).
    $('btnEliminarPrealta').classList.toggle('oculto', !(p.Estado === 'borrador' && PUEDE.corregir(estado.rol)));
    pintarCertificadoEnDetalle(p);   // v0.35.0
}

async function eliminarPrealta() {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'borrador' || !PUEDE.corregir(estado.rol)) return;
    // La ventana de 90 dias no puede contestar esto: una pre-alta vieja tendria sus embarques
    // fuera de la carga y el borrador se borraria con historia colgando. Se pregunta EN VIVO.
    let citada;
    try { citada = (await estado.cliente.renglones(estado.siteId, L.embarques, `fields/PreAltaId eq ${p.id}`)).length; }
    catch (err) { avisar('No pude confirmar si tiene embarques (' + err.message + '). No se elimina: ciérrala.', 'error'); return; }
    if (citada) { avisar('Esta pre-alta ya tiene embarques: no se puede eliminar, ciérrala.', 'error'); return; }
    const { ok } = await confirmar({ titulo: 'Eliminar el borrador', peligro: true, ok: 'Eliminar',
        texto: `«${p.Title}» nunca se firmó, así que la puerta no la ha usado. Se borra el renglón; la campaña ${p.Campana || ''} queda libre.`, motivo: 'opcional', etiquetaMotivo: 'Por qué (opcional)' });
    if (!ok) return;
    try {
        await refrescarCliente();
        await estado.cliente.borrarRenglon(estado.siteId, L.prealtas, p.id);
        estado.prealtas = estado.prealtas.filter(x => x.id !== p.id);
        estado.prealtaAbierta = null;
        avisar('Borrador eliminado.', 'bien'); trasCambioPrealta();
    } catch (e) { avisar('No se pudo eliminar: ' + e.message, 'error'); }
}

async function firmarPrealta() {
    const p = estado.prealtaAbierta; if (!p) return;
    await escribiendo('btnFirmar', async () => {   // C-24 / C-23
    const reFirma = p.Estado === 'firmada';   // S-01: trae el sello pero no su renglon en PLANTA_Firmas; solo falta la firma
    const { ok } = await confirmar({ titulo: 'Firmar la pre-alta', ok: 'Firmar',
        texto: reFirma ? `«${p.Title}» trae el sello de ${quien(p.FirmadaPor)} pero no su firma registrada, así que la puerta no la ve. Con tu firma queda completa; queda registrado quién y cuándo.`
            : `«${p.Title}». Con tu firma la puerta empieza a aceptar sus góndolas; queda registrado quién y cuándo.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        await firmar('prealta', p);   // S-01: primero la firma (403 si no eres validador/gerencia), luego el sello
        const campos = { Estado: 'firmada', FirmadaPor: estado.cuenta.username, FirmadaEl: new Date().toISOString() };
        await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, campos);
        aplicar('prealtas', p, campos);   // C-23
        avisar('Pre-alta firmada.', 'bien'); trasCambioPrealta();
    } catch (e) { avisar('No se pudo firmar: ' + e.message, 'error'); }
    });
}
async function cerrarPrealta() {
    const p = estado.prealtaAbierta; if (!p) return;
    await escribiendo('btnCerrarPrealta', async () => {   // C-24 / C-23
    const { ok } = await confirmar({ titulo: 'Cerrar el programa', ok: 'Cerrar',
        texto: `«${p.Title}». La puerta dejará de aceptar góndolas contra él. No se borra: queda como historial.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        const campos = { Estado: 'cerrada', CerradaPor: estado.cuenta.username, CerradaEl: new Date().toISOString() };
        try { await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, campos); }
        catch (err) {
            // Lista sin las columnas CerradaPor/CerradaEl (tarea 10 del setup): se cierra igual, sin sello, y se avisa.
            if (!esColumnaFaltante(err)) throw err;   // C-15
            await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, p.id, { Estado: 'cerrada' });
            delete campos.CerradaPor; delete campos.CerradaEl;
            avisar('Programa cerrado, pero sin registrar quién lo cerró' + (estado.rol === 'gerencia' ? ': la lista PLANTA_Prealtas no tiene todavía CerradaPor / CerradaEl (setup, tarea 10).' : '. Avisa a gerencia.'), 'ojo');   // U-31
        }
        aplicar('prealtas', p, campos); trasCambioPrealta();   // C-23
    } catch (e) { avisar('No se pudo cerrar: ' + e.message, 'error'); }
    });
}

// ================================================================ CERTIFICADO DE TRATAMIENTO (v0.35.0)
//
// Decision de Carlos (2026-09-22): UN certificado por PROGRAMA (pre-alta cerrada), lo emite SOLO gerencia, se firma en
// PLANTA_Firmas (Tipo certificado, S-11: sin firma valida la app no lo imprime) y lleva QR a una pagina publica sin login
// (certificado/ de este repo) que muestra solo lo impreso. El papel se CONGELA al emitir (kg, folios, generador…): si el
// programa cambia despues, el certificado no se edita — se SUSTITUYE (nace otro; el viejo queda «sustituido») o se CANCELA.
// Formato FO-CT-01 (rumbo D «Diploma» + incisos de Global Water, mockup M del artifact Design; memoria
// calytek-certificado-tratamiento). La pagina publica se regenera con docs/exportar-planta.ps1 -PublicarCertificados.

/** Dia (hora de Mexico) de un ISO datetime, como dd/mm/aaaa. aIsoDia NO sirve aqui: espera una fecha tecleada y lanza con un ISO. */
const diaCert = iso => (iso ? fechaCorta(fechaMexico(new Date(iso))) : '—');
function certificadosDe(p) { return estado.certificados.filter(c => Number(c.PreAltaId) === Number(p.id)).sort((a, b) => b.id - a.id); }
function certificadoVigente(p) { return certificadosDe(p).find(c => c.Estado === 'vigente') || null; }
/** S-11 sobre el certificado: vale solo con un renglon de PLANTA_Firmas de Tipo certificado, firmado por gerencia y por la misma cuenta que EmitidoPor. */
function certificadoFirmado(c) { return !!firmaDe('certificado', c.id, c.EmitidoPor); }
const motivoSinCertificados = () => (estado.certificadosError ? `No se pudo leer la lista de certificados: no se emite hasta que se vea.${estado.rol === 'gerencia' ? ` (PLANTA_Certificados: ${estado.certificadosError} — si no existe, se crea con herramientas-dev/provisionar.html)` : ' Avisa a gerencia.'}` : null);

function pintarCertificadoEnDetalle(p) {
    const ul = $('paDetalleCertificado'); ul.textContent = '';
    const cerrada = p.Estado === 'cerrada';
    const certs = cerrada ? certificadosDe(p) : [];
    const vig = cerrada ? certificadoVigente(p) : null;
    if (cerrada) {
        if (!certs.length) { const li = el('li', '', 'Certificado de tratamiento '); li.appendChild(etiqueta('sin emitir', 'aviso')); li.appendChild(el('span', 'd', PUEDE.emitirCertificado(estado.rol) ? 'Solo gerencia lo emite; queda registrado quién y cuándo, y el papel lleva QR de verificación.' : 'Lo emite gerencia.')); ul.appendChild(li); }
        for (const c of certs) {
            const li = el('li', '', `Certificado ${c.Title} `);
            li.appendChild(etiqueta(c.Estado, c.Estado === 'vigente' ? 'ok' : c.Estado === 'cancelado' ? 'vencida' : ''));
            if (c.Estado === 'vigente' && !certificadoFirmado(c)) li.appendChild(etiqueta('sin firma', 'vencida'));
            li.appendChild(el('span', 'd', `${toneladas(c.Kg)} t · ${lista(c.Embarques).length} embarques · emitido ${horaCorta(c.EmitidoEl)} por ${quien(c.EmitidoPor) || '—'}${c.SustituidoPor ? ` · sustituido por ${c.SustituidoPor}` : ''}${c.Estado === 'cancelado' ? ` · cancelado ${horaCorta(c.CanceladoEl)} por ${quien(c.CanceladoPor) || '—'}` : ''}${c.Motivo ? ` · ${c.Motivo}` : ''}`));
            ul.appendChild(li);
        }
    }
    const puede = cerrada && PUEDE.emitirCertificado(estado.rol);
    const bloqueo = motivoSinCertificados() || motivoSinFirmas() || '';
    $('btnEmitirCertificado').classList.toggle('oculto', !(puede && !vig));
    $('btnEmitirCertificado').disabled = !!bloqueo; $('btnEmitirCertificado').title = bloqueo;
    $('btnSustituirCertificado').classList.toggle('oculto', !(puede && vig));
    $('btnSustituirCertificado').disabled = !!bloqueo; $('btnSustituirCertificado').title = bloqueo;
    $('btnCancelarCertificado').classList.toggle('oculto', !(puede && vig));
    $('btnVerCertificado').classList.toggle('oculto', !certs.length);
}

/**
 * Emite el certificado del programa abierto. Con `sustituye`, es una SUSTITUCION: el nuevo nace y el viejo queda
 * «sustituido» con SustituidoPor. Los embarques se leen EN VIVO por PreAltaId (la ventana de 90 dias dejaria fuera
 * los de un programa largo). Orden de escritura: renglon del certificado -> folio unico -> FIRMA (403 si no eres
 * gerencia: entonces el renglon recien creado se cancela con ese motivo, para que no quede un certificado sin firma
 * que parezca vigente) -> el viejo, si sustituye.
 */
async function emitirCertificado(sustituye = null, motivoSust = '') {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'cerrada' || !PUEDE.emitirCertificado(estado.rol)) return;
    const btn = sustituye ? 'btnSustituirCertificado' : 'btnEmitirCertificado';
    await escribiendo(btn, async () => {
    let live;
    try { await refrescarCliente(); live = await estado.cliente.renglones(estado.siteId, L.embarques, `fields/PreAltaId eq ${p.id}`); }
    catch (err) { avisar('No pude leer los embarques del programa (' + err.message + '). No se emite.', 'error'); return; }
    const r = resumenCertificado(p, live);
    if (!r.embarques.length) { avisar('Este programa no tiene embarques cerrados con neto: no hay nada que certificar.', 'error'); return; }
    if (!sustituye) {
        const { ok } = await confirmar({ titulo: 'Emitir el certificado de tratamiento', ok: 'Emitir',
            texto: `«${p.Title}»: ${plural(r.embarques.length, 'embarque cerrado', 'embarques cerrados')}, ${toneladas(r.kg)} t netas (${r.kg.toLocaleString('es-MX')} kg), del ${diaCert(r.primerCierre)} al ${diaCert(r.ultimoCierre)}. Con tu firma queda emitido a nombre de ${p.Generador || '(sin generador)'}; lo que diga el papel no cambia después: si algo está mal, se sustituye.` });
        if (!ok) return;
    }
    const carrier = porId(estado.carriers, p.CarrierId);
    let nuevo = null;
    try {
        const certsDelAno = await estado.cliente.renglones(estado.siteId, L.certificados, null);
        const folio = siguienteFolio('C', certsDelAno.map(x => x.Title));
        const ahora = new Date().toISOString();
        nuevo = await estado.cliente.crearRenglon(estado.siteId, L.certificados, limpiar({
            Title: folio, PreAltaId: p.id, Estado: 'vigente', Kg: r.kg,
            Embarques: r.folios.join('; '), Manifiestos: r.manifiestos.join('; '),
            PrimerCierre: r.primerCierre, UltimoCierre: r.ultimoCierre,
            Generador: p.Generador || null, GeneradorRegistro: p.GeneradorRegistro || null, Pozo: p.Pozo || null, Corriente: p.Corriente || null,
            Transportista: carrier ? `${carrier.Title} · autorización ${carrier.AutorizacionASEA || 'sin número'}` : null,
            Sufijo: sufijoVerificacion(), EmitidoPor: estado.cuenta.username, EmitidoEl: ahora, Version: VERSION,
            Motivo: sustituye ? `Sustituye a ${sustituye.Title}: ${motivoSust}`.slice(0, 255) : null
        }));
        await asegurarFolioUnico(nuevo, 'C');
        try { await firmar('certificado', nuevo); }
        catch (err) {
            const cancel = { Estado: 'cancelado', CanceladoPor: estado.cuenta.username, CanceladoEl: new Date().toISOString(), Motivo: ('Sin firma: ' + err.message).slice(0, 255) };
            try { await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, nuevo.id, cancel); } catch (_) { /* queda sin firma: la app no lo imprime (certificadoFirmado) */ }
            Object.assign(nuevo, cancel); estado.certificados.push(nuevo);
            throw err;
        }
        estado.certificados.push(nuevo);
        if (sustituye) {
            const campos = { Estado: 'sustituido', SustituidoPor: nuevo.Title, Motivo: String(motivoSust).slice(0, 255) };
            await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, sustituye.id, campos);
            aplicar('certificados', sustituye, campos);
        }
        avisar(`Certificado ${nuevo.Title} emitido${sustituye ? ` en sustitución de ${sustituye.Title}` : ''}. Publica la verificación con el exporte (-PublicarCertificados) para que el QR conteste.`, 'bien');
        pintarCertificadoEnDetalle(p);
        verCertificado(nuevo);
    } catch (e) { avisar('No se pudo emitir: ' + e.message, 'error'); pintarCertificadoEnDetalle(p); }
    });
}
async function sustituirCertificado() {
    const p = estado.prealtaAbierta; const vig = p && certificadoVigente(p); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    const { ok, motivo } = await confirmar({ titulo: 'Sustituir el certificado', ok: 'Sustituir', motivo: true, etiquetaMotivo: 'Qué estaba mal',
        texto: `${vig.Title} queda como «sustituido» (su QR lo dirá) y nace uno nuevo con los embarques cerrados de HOY del programa. El motivo queda registrado en los dos.` });
    if (!ok) return;
    await emitirCertificado(vig, motivo);
}
async function cancelarCertificado() {
    const p = estado.prealtaAbierta; const vig = p && certificadoVigente(p); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    await escribiendo('btnCancelarCertificado', async () => {
    const { ok, motivo } = await confirmar({ titulo: 'Cancelar el certificado', ok: 'Cancelar el certificado', peligro: true, motivo: true, etiquetaMotivo: 'Por qué se cancela',
        texto: `${vig.Title} deja de valer: su QR dirá «cancelado». No se borra: queda como historial. Si hace falta uno bueno, después se emite otro.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        const campos = { Estado: 'cancelado', CanceladoPor: estado.cuenta.username, CanceladoEl: new Date().toISOString(), Motivo: String(motivo).slice(0, 255) };
        await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, vig.id, campos);
        aplicar('certificados', vig, campos);
        avisar(`Certificado ${vig.Title} cancelado.`, 'bien'); pintarCertificadoEnDetalle(p);
    } catch (e) { avisar('No se pudo cancelar: ' + e.message, 'error'); }
    });
}
function verCertificado(c) {
    const p = estado.prealtaAbierta;
    const cert = c || (p && (certificadoVigente(p) || certificadosDe(p)[0]));
    if (!cert) return;
    estado.certificadoAbierto = cert;
    pintarCertificado(cert, $('certificado'));
    const d = $('dlgCertificado'); if (!d.open) d.showModal();
    // Sin firma valida (S-11) o sin estar vigente, se VE pero no se imprime como bueno: el papel llevaria un estado que no es.
    const imprimible = cert.Estado === 'vigente' && certificadoFirmado(cert);
    $('ctImprimir').disabled = !imprimible;
    $('ctImprimir').title = imprimible ? '' : (cert.Estado !== 'vigente' ? `Este certificado está ${cert.Estado}: no se imprime como vigente.` : 'Sin firma registrada de gerencia: no se imprime.');
    $('ctEstado').textContent = imprimible ? `${cert.Title} · vigente y firmado` : `${cert.Title} · ${cert.Estado}${cert.Estado === 'vigente' ? ' · SIN FIRMA' : ''}`;
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
/** Pinta el formato FO-CT-01 (carta horizontal) dentro de `t`. Solo lee el renglon del certificado: es lo que se congelo al emitir. */
function pintarCertificado(c, t) {
    const CFG = CONFIG.certificado;
    t.textContent = '';
    t.classList.toggle('sustituido', c.Estado === 'sustituido'); t.classList.toggle('cancelado', c.Estado === 'cancelado');
    const marco = el('div', 'ct-marco'); const hoja = el('div', 'ct-hoja'); marco.appendChild(hoja); t.appendChild(marco);
    const agua = el('img', 'ct-agua'); agua.src = './marca/lockup.svg'; agua.alt = ''; hoja.appendChild(agua);
    if (c.Estado !== 'vigente') hoja.appendChild(el('div', 'ct-sello', c.Estado === 'cancelado' ? 'CANCELADO' : `SUSTITUIDO POR ${c.SustituidoPor || '—'}`));
    hoja.appendChild(el('div', 'ct-formato', CFG.formato));
    const folioCaja = el('div', 'ct-folio'); folioCaja.appendChild(el('div', 'ct-eti', 'FOLIO')); folioCaja.appendChild(el('div', 'ct-folio-num', c.Title)); hoja.appendChild(folioCaja);
    const logo = el('img', 'ct-logo'); logo.src = './marca/lockup.svg'; logo.alt = 'MINSA ENERGY'; hoja.appendChild(logo);
    hoja.appendChild(el('h1', 'ct-titulo', 'CERTIFICADO DE TRATAMIENTO'));
    const sub = el('div', 'ct-sub'); sub.appendChild(el('i')); sub.appendChild(el('span', '', 'RECORTES DE PERFORACIÓN · RESIDUO DE MANEJO ESPECIAL')); sub.appendChild(el('i')); hoja.appendChild(sub);
    const inciso = (eti, valor, clase = '') => { const d = el('div', 'ct-inciso'); d.appendChild(el('span', 'ct-eti', eti)); d.appendChild(el('span', 'ct-val ' + clase, valor || '—')); return d; };
    const fila = (...incisos) => { const f = el('div', 'ct-fila'); for (const i of incisos) f.appendChild(i); return f; };
    hoja.appendChild(fila(inciso('GENERADOR:', c.Generador, 'fuerte')));
    hoja.appendChild(fila(inciso('REGISTRO DE GENERADOR:', c.GeneradorRegistro, 'mono'), inciso('POZO:', c.Pozo, 'fuerte')));
    const prosa = el('p', 'ct-prosa'); prosa.appendChild(el('b', '', 'MATERIAS INDUSTRIALIZADAS CCMV DEL NORTE, S.A. DE C.V.')); prosa.appendChild(document.createTextNode(' — MINSA ENERGY — certifica que ha recibido para su tratamiento en su planta '));
    prosa.appendChild(el('b', '', 'CALYTEK')); prosa.appendChild(document.createTextNode(', de una forma ambientalmente segura y conforme a los términos de su autorización, el residuo de:')); hoja.appendChild(prosa);
    hoja.appendChild(el('div', 'ct-residuo', residuoDe(c.Corriente)));
    hoja.appendChild(el('div', 'ct-regla'));
    const folios = lista(c.Embarques), manif = lista(c.Manifiestos);
    hoja.appendChild(fila(inciso('VOLUMEN:', `${toneladas(c.Kg)} TON.`, 'volumen'), inciso('EMBARQUES:', String(folios.length), 'mono'), inciso('FECHA DE RECEPCIÓN:', c.PrimerCierre && c.UltimoCierre && fechaMexico(new Date(c.PrimerCierre)) !== fechaMexico(new Date(c.UltimoCierre)) ? `del ${diaCert(c.PrimerCierre)} al ${diaCert(c.UltimoCierre)}` : diaCert(c.UltimoCierre || c.PrimerCierre), 'fuerte')));
    hoja.appendChild(fila(inciso('TRANSPORTISTA:', c.Transportista, 'fuerte')));   // renglon entero: un numero de autorizacion no se parte (revisor v0.35.1)
    hoja.appendChild(fila(inciso('AUTORIZACIÓN DE LA PLANTA:', CFG.autorizacionPlanta || 'pendiente (ASEA-03-011-A)', 'mono')));
    // Incisos: un renglon por embarque cerrado — folio · manifiesto · neto. Es lo que sustituye al «ticket de bascula» de un certificado por embarque.
    const tabla = el('table', 'ct-incisos'); const th = el('tr'); for (const h of ['#', 'FOLIO DE EMBARQUE', 'MANIFIESTO', 'NETO (KG)']) th.appendChild(el('th', '', h)); tabla.appendChild(th);
    folios.forEach((f, i) => { const tr = el('tr'); tr.appendChild(el('td', 'n', String(i + 1))); tr.appendChild(el('td', 'mono', f)); tr.appendChild(el('td', 'mono', manif[i] || '—')); tr.appendChild(el('td', 'mono kg', '')); tabla.appendChild(tr); });
    hoja.appendChild(tabla);
    // El neto por embarque no se congela en el renglon (solo la suma, Kg): se toma del embarque cargado si esta en la ventana; si no, en blanco.
    [...tabla.querySelectorAll('td.kg')].forEach((td, i) => { const e = estado.embarques.find(x => x.Title === folios[i]); td.textContent = e && e.NetoKg ? Number(e.NetoKg).toLocaleString('es-MX') : ''; });
    const pieFirma = el('div', 'ct-firmas');
    const firma = el('div', 'ct-firma'); firma.appendChild(el('div', 'ct-raya')); firma.appendChild(el('div', 'ct-nombre', CFG.responsableTecnico || '(nombre pendiente)')); firma.appendChild(el('div', 'ct-cargo', 'Responsable técnico de planta · CALYTEK')); pieFirma.appendChild(firma);
    const url = urlVerificacion(CFG.urlVerificacion, c);
    const qrCaja = el('div', 'ct-qr');
    if (url) { qrCaja.appendChild(qrSvg(url, 84)); qrCaja.appendChild(el('div', 'ct-qr-leyenda', `Verifique este certificado en\n${CFG.urlVerificacion.replace(/^https?:\/\//, '')}`)); }
    else qrCaja.appendChild(el('div', 'ct-qr-leyenda', 'Sin QR: certificado sin sufijo de verificación'));
    pieFirma.appendChild(qrCaja); hoja.appendChild(pieFirma);
    const pie = el('div', 'ct-pie');
    pie.appendChild(el('div', '', `Recibido en las instalaciones de MINSA ENERGY · Planta CALYTEK, ${CFG.domicilioPlanta} · ${CFG.resolutivo} · RFC ${CFG.rfc}`));
    pie.appendChild(el('div', '', `Este certificado ampara únicamente el residuo y la cantidad aquí descritos. Emitido ${horaCorta(c.EmitidoEl)} por ${quien(c.EmitidoPor) || '—'} · CALYTEK Planta ${c.Version || VERSION}`));
    hoja.appendChild(pie);
}
function imprimirCertificado() {
    if ($('ctImprimir').disabled) return;
    // La hoja de impresion (carta horizontal, solo el formato) se enciende SOLO mientras dura el print: @page no admite clase
    // y el ticket termico / Reportes no deben salir en carta horizontal.  La CSP no deja un <style> inline; un <link> propio si.
    document.body.classList.add('imprimiendo-certificado'); $('cssImpresionCert').disabled = false;
    const quitar = () => { document.body.classList.remove('imprimiendo-certificado'); $('cssImpresionCert').disabled = true; window.removeEventListener('afterprint', quitar); };
    window.addEventListener('afterprint', quitar);
    window.print();
    setTimeout(quitar, 60000);   // por si afterprint no llega (algun WebView)
}

// ================================================================ PADRON

// Los grupos (pre-altas y padron) arrancan PLEGADOS en todo tamano (pedido de Carlos 2026-09-07):
// el usuario abre el que quiera y la app respeta lo que dejo abierto o cerrado al repintar.
// Ficha completa de un renglón del padrón: TODAS las columnas de la lista, no el resumen de una línea (Carlos, 2026-09-06).
// Se despliega al tocar el renglón; el resumen sigue siendo lo que se skimea.
// Vigencias que trae un renglón del padrón (vencida, por vencer o sin fecha), con la misma regla que la puerta.
const VIGENCIAS_PADRON = {
    carriers: [['autorización ASEA', 'VigenciaASEA', 'legal'], ['CSF', 'CSFVigencia', 'comercial']],
    unidades: [['tarjeta de circulación', 'TarjetaVigencia', 'comercial'], ['póliza', 'PolizaVigencia', 'comercial']],
    choferes: [['licencia', 'LicenciaVigencia', 'comercial']],
};
function vigenciasPadron(tipo, it) {
    return VIGENCIAS_PADRON[tipo].map(([n, col, cl]) => evaluarVigencia(n, it[col], cl, CONFIG.avisoVigenciaDias)).filter(Boolean);
}
const sinFecha = h => h.detalle.startsWith('sin fecha');
// La etiqueta del renglón: lo peor que traiga (vencida > sin fecha > por vencer).
function etiquetaVigencia(tipo, it) {
    const h = vigenciasPadron(tipo, it);
    const peor = h.find(x => !x.ok && !sinFecha(x)) || h.find(sinFecha) || h[0];
    // Vencida es roja siempre: aquí no se juzga la puerta (legal/comercial), se avisa que el papel ya no sirve.
    if (!peor) return null;
    if (!peor.ok && !sinFecha(peor)) return etiqueta('vencida', 'vencida');
    return etiqueta(sinFecha(peor) ? 'sin fecha' : 'por vencer', sinFecha(peor) ? peor.clase : 'aviso');
}
// El aviso del alta dice qué se guardó Y qué trae vencido o por vencer, para que quien captura no se entere hasta la puerta.
function avisarAlta(texto, tipo, it) {
    const h = vigenciasPadron(tipo, it).filter(x => !sinFecha(x));
    if (!h.length) { avisar(texto, 'bien'); return; }
    const vencidas = h.filter(x => !x.ok);
    avisar(`${texto} ⚠ ${h.map(x => `${x.regla} ${x.detalle}`).join(' · ')}`, vencidas.length ? 'error' : 'ojo');
}
const CAMPOS_PADRON = {
    carriers: [['Autorización ASEA', 'AutorizacionASEA'], ['Vence ASEA', 'VigenciaASEA', 'fecha'], ['Corrientes', 'Corrientes', 'lista'], ['Folio del oficio', 'FolioOficio'], ['Registro SCT', 'RegistroSCT'], ['CSF vence', 'CSFVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
    unidades: [['Placa plana', 'PlacaPlana'], ['Carrier', 'CarrierId', 'carrier'], ['Tipo', 'TipoUnidad'], ['Marca', 'Marca'], ['No. de serie', 'NumeroSerie'], ['Capacidad', 'CapacidadKg', 'kg'], ['Folio del oficio', 'FolioOficio'], ['Tarjeta de circulación', 'TarjetaCirc'], ['Tarjeta vence', 'TarjetaVigencia', 'fecha'], ['Póliza', 'Poliza'], ['Póliza vence', 'PolizaVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
    choferes: [['Carrier', 'CarrierId', 'carrier'], ['Licencia', 'Licencia'], ['Licencia vence', 'LicenciaVigencia', 'fecha'], ['Activo', 'Activo', 'si'], ['Notas', 'Notas']],
};
function fichaPadron(tipo, x) {
    const dl = el('dl', 'ficha oculto');
    for (const [rotulo, col, fmt] of CAMPOS_PADRON[tipo]) {
        const v = x[col];
        let txt;
        if (fmt === 'fecha') txt = fechaCorta(v);
        else if (fmt === 'lista') txt = lista(v).join(', ') || '—';
        else if (fmt === 'carrier') txt = nombreDe(estado.carriers, v);
        else if (fmt === 'kg') txt = v ? `${v} kg` : '—';
        else if (fmt === 'si') txt = v === false ? 'No (dado de baja)' : 'Sí';
        else txt = (v === null || v === undefined || v === '') ? '—' : String(v);
        dl.appendChild(el('dt', null, rotulo)); dl.appendChild(el('dd', null, txt));
    }
    return dl;
}
// Tocar el renglón despliega/oculta su ficha; para el carrier además filtra unidades y choferes.
function conFicha(r, tipo, x, extra) {
    const ficha = fichaPadron(tipo, x);
    r.firstChild.appendChild(ficha);
    // Los botones (Editar, Dar de baja…) solo se ven con la ficha abierta: la lista se lee limpia y se actúa sobre un inciso a la vez.
    r.classList.toggle('plegado', estado.padronFicha !== `${tipo}:${x.id}`);
    if (estado.padronFicha === `${tipo}:${x.id}`) ficha.classList.remove('oculto');
    desplegable(r.firstChild, estado.padronFicha === `${tipo}:${x.id}`, () => {
        const llave = `${tipo}:${x.id}`;
        estado.padronFicha = estado.padronFicha === llave ? null : llave;
        if (extra) extra(); else { ficha.classList.toggle('oculto', estado.padronFicha !== llave); r.classList.toggle('plegado', estado.padronFicha !== llave); r.firstChild.setAttribute('aria-expanded', String(estado.padronFicha === llave)); }
    });
    return r;
}
// Plegar un grupo del padrón suelta lo que tenía seleccionado (pedido de Carlos 2026-09-07): un filtro o una
// ficha escondidos dentro de un grupo cerrado seguían actuando sin que nadie los viera (la foto del chofer único).
for (const [id, tipo] of [['pdGrupoCarriers', 'carriers'], ['pdGrupoUnidades', 'unidades'], ['pdGrupoChoferes', 'choferes']]) {
    // U-56 (v0.29.0): en escritorio la cabecera no tiene chevron ni cursor de mano y los tres grupos van abiertos (C-08); un clic
    // en «Unidades» para enfocar la columna la plegaba sin señal de cómo volver. Los botones del summary («+ Alta») ya cortan el evento.
    const sinPlegar = ev => { if (ESCRITORIO.matches && !ev.target.closest('button')) ev.preventDefault(); };
    $(id).querySelector('summary').addEventListener('click', sinPlegar);
    $(id).querySelector('summary').addEventListener('keydown', ev => { if (ev.key === 'Enter' || ev.key === ' ') sinPlegar(ev); });
    $(id).addEventListener('toggle', () => {
        if ($(id).open) return;
        let cambio = false;
        if (tipo === 'carriers' && estado.padronCarrier !== null && estado.padronCarrier !== undefined) { estado.padronCarrier = null; cambio = true; }
        if (estado.padronFicha && estado.padronFicha.startsWith(tipo + ':')) { estado.padronFicha = null; cambio = true; }
        if (cambio) pintarPadron();
    });
}
function elegirCarrierPadron(c) {
    estado.padronCarrier = estado.padronCarrier === c.id ? null : c.id;   // segundo toque = quitar el filtro
    estado.padronFicha = estado.padronCarrier === null ? null : `carriers:${c.id}`;
    pintarPadron();
}
// I1 (7-sep): el resumen de cada grupo dice cuantos hay y cuantos estan vencidos o por vencer, para que el grupo
// plegado siga informando. En escritorio (>= 900 px) las tres columnas se abren solas porque ahi caben; en celular
// se quedan plegadas (pedido de Carlos, v0.16.0). Reversible: quitar abrirGruposPadronEscritorio().
const ESCRITORIO = window.matchMedia('(min-width: 900px)');
function abrirGruposPadronEscritorio() { if (ESCRITORIO.matches) for (const id of ['pdGrupoCarriers', 'pdGrupoUnidades', 'pdGrupoChoferes']) $(id).open = true; }
ESCRITORIO.addEventListener('change', abrirGruposPadronEscritorio);
// U-11 (v0.22.0): en el celular el campo de peso NO levanta el teclado del sistema (inputmode=none): se captura con el
// teclado propio, que quedaba tapado por el del telefono. En escritorio el propio se oculta por CSS y el campo vuelve a numeric.
function ajustarTecladoPeso() { $('baKg').inputMode = ESCRITORIO.matches ? 'numeric' : 'none'; }
ESCRITORIO.addEventListener('change', ajustarTecladoPeso); ajustarTecladoPeso();
function resumenPadron(tipo, items, unidad) {
    const vivos = items.filter(x => x.Activo !== false);
    const vencidos = vivos.filter(x => vigenciasPadron(tipo, x).some(h => !h.ok && !sinFecha(h))).length;
    const porVencer = vivos.filter(x => { const h = vigenciasPadron(tipo, x); return !h.some(v => !v.ok && !sinFecha(v)) && h.some(v => v.ok); }).length;
    const [uno, varios] = Array.isArray(unidad) ? unidad : [unidad, unidad + 's'];   // U-36: 'activo' → activos; ['con licencia', 'con licencia'] no cambia
    const s = el('span', '', plural(vivos.length, uno, varios));
    if (vencidos) { s.appendChild(document.createTextNode(' · ')); s.appendChild(el('span', 'mal', `${vencidos} vencid${vencidos === 1 ? 'o' : 'os'}`)); }
    if (porVencer) { s.appendChild(document.createTextNode(' · ')); s.appendChild(el('span', 'ojo', `${porVencer} por vencer`)); }
    return s;
}
const normaliza = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
/** Lo que distingue a cada grupo del padron (C-18): contenedor, campos que busca el filtro de texto, si el carrier elegido lo acota, titulo y extra de la ficha. */
const GRUPOS_PADRON = {
    carriers: { contenedor: 'pdCarriers', porCarrier: false, busca: c => [c.Title, c.AutorizacionASEA, c.FolioOficio], titulo: c => c.Title, extra: c => () => elegirCarrierPadron(c) },
    unidades: { contenedor: 'pdUnidades', porCarrier: true, busca: u => [u.Title, u.PlacaPlana, u.NumeroSerie, nombreDe(estado.carriers, u.CarrierId)], titulo: u => `${u.Title}${u.PlacaPlana ? ' / ' + u.PlacaPlana : ''}` },
    choferes: { contenedor: 'pdChoferes', porCarrier: true, busca: ch => [ch.Title, ch.Licencia, nombreDe(estado.carriers, ch.CarrierId)], titulo: ch => ch.Title }
};
/** Etiqueta de estado del renglon: «baja» si esta inactivo, si no la de vigencia (vencida / por vencer), si la hay. */
function marcaEstado(r, clave, x) {
    if (x.Activo === false) { r.classList.add('baja'); r.firstChild.firstChild.appendChild(etiqueta('baja', 'baja')); return; }
    const v = etiquetaVigencia(clave, x); if (v) r.firstChild.firstChild.appendChild(v);
}
function pintarPadron() {
    const puede = PUEDE.capturarPrealta(estado.rol);
    for (const id of ['btnNuevoCarrier', 'btnNuevaUnidad', 'btnNuevoChofer']) $(id).classList.toggle('oculto', !puede);
    for (const id of ['pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer']) cerrarForma(id);
    estado.padronEdit = null;
    abrirGruposPadronEscritorio();
    const sel = estado.padronCarrier;
    const q = normaliza($('pdBusca').value.trim());
    const pega = filtroTexto(q);   // C-26
    $('pdResCarriers').textContent = ''; $('pdResCarriers').appendChild(resumenPadron('carriers', estado.carriers, 'activo'));
    $('pdResUnidades').textContent = ''; $('pdResUnidades').appendChild(resumenPadron('unidades', estado.unidades, 'amparada'));
    $('pdResChoferes').textContent = ''; $('pdResChoferes').appendChild(resumenPadron('choferes', estado.choferes, ['con licencia', 'con licencia']));
    let encontrados = 0;
    // C-18 (v0.25.0): un solo bucle sobre GRUPOS_PADRON en vez de tres gemelos; lo que cambia por grupo vive en la tabla.
    const filtraCarrier = x => !haySel(sel) || Number(x.CarrierId) === Number(sel);   // C-26
    const contenedores = {};
    for (const [clave, g] of Object.entries(GRUPOS_PADRON)) {
        const cont = $(g.contenedor); cont.textContent = ''; contenedores[clave] = cont;
        for (const x of estado[clave]) {
            if (g.porCarrier && !filtraCarrier(x)) continue;
            if (!pega(...g.busca(x))) continue;
            encontrados++;
            const r = renglon(g.titulo(x), null, null, null, botonesPadron(clave, x));
            marcaEstado(r, clave, x);
            if (clave === 'carriers') { r.classList.toggle('sel', sel === x.id); r.firstChild.setAttribute('aria-pressed', String(sel === x.id)); }   // U-33: el filtro por carrier como estado accesible
            cont.appendChild(conFicha(r, clave, x, g.extra ? g.extra(x) : undefined));
        }
    }
    const c1 = contenedores.carriers, c2 = contenedores.unidades, c3 = contenedores.choferes;
    if (haySel(sel)) c1.appendChild(el('p', 'filtro', `Mostrando solo lo de ${nombreDe(estado.carriers, sel)} · toca el carrier otra vez para ver todo`));
    if (!estado.carriers.length) c1.appendChild(el('p', 'pista', 'Sin carriers. La primera pre-alta necesita uno con su oficio ASEA transcrito.'));
    // U-18 (v0.22.0): un grupo sin renglones lo dice, y distinto si el vacio es por el filtro de carrier. Antes solo
    // Carriers tenia estado vacio y «0 unidades» se leia igual que «este carrier no tiene».
    const filtrado = haySel(sel);
    if (!q && !c2.querySelector('.renglon')) c2.appendChild(el('p', 'pista', filtrado ? `${nombreDe(estado.carriers, sel)} no tiene unidades en el padrón.` : 'Sin unidades. Se transcriben del oficio del carrier con «+ Alta».'));
    if (!q && !c3.querySelector('.renglon')) c3.appendChild(el('p', 'pista', filtrado ? `${nombreDe(estado.carriers, sel)} no tiene choferes en el padrón.` : 'Sin choferes. Se dan de alta con su licencia con «+ Alta».'));
    // Con texto en el buscador: los grupos con resultado se abren, los vacios lo dicen, y la cuenta va junto al campo.
    if (q) {
        for (const [g, cont] of [['pdGrupoCarriers', c1], ['pdGrupoUnidades', c2], ['pdGrupoChoferes', c3]]) {
            const hay = cont.querySelector('.renglon');
            if (!hay) cont.appendChild(el('p', 'sin-resultado', 'Nada con ese texto.'));
            $(g).open = !!hay;
        }
    }
    $('pdBuscaCuenta').textContent = q ? plural(encontrados, 'resultado') : '';
    const sels = [$('puuCarrier'), $('pchCarrier')];
    for (const s of sels) opciones(s, estado.carriers.filter(c => c.Activo !== false), c => c.id, c => c.Title);
}

/**
 * Corregir el padron (2026-09-05). No hay edicion en sitio: una placa mal transcrita se corrige
 * dando de baja el renglon y transcribiendola de nuevo del oficio (misma regla del padron §4).
 *   - sin referencias (ninguna pre-alta, embarque, unidad o chofer lo cita): se ELIMINA.
 *   - con referencias: se da de BAJA (Activo = false): la puerta deja de ampararlo, el historial
 *     sigue apuntando a un renglon que existe. Se puede reactivar.
 */
const COLUMNA_PADRON = { carriers: 'CarrierId', unidades: 'UnidadId', choferes: 'ChoferId' };

/** Referencias FUERA de embarques (padron y pre-altas se cargan enteros: esto es exacto). */
function referenciasSinEmbarques(clave, x) {
    const id = String(x.id), num = Number(x.id);
    if (clave === 'carriers') return estado.unidades.filter(u => Number(u.CarrierId) === num).length + estado.choferes.filter(c => Number(c.CarrierId) === num).length
        + estado.prealtas.filter(p => Number(p.CarrierId) === num).length;
    if (clave === 'unidades') return estado.prealtas.filter(p => lista(p.UnidadesIds).includes(id)).length;
    return estado.prealtas.filter(p => lista(p.ChoferesIds).includes(id)).length;
}
/** Cuenta con lo que hay en memoria: sirve para ELEGIR el boton, nunca para autorizar el borrado. */
function referenciasPadron(clave, x) {
    const num = Number(x.id), col = COLUMNA_PADRON[clave];
    return referenciasSinEmbarques(clave, x) + estado.embarques.filter(e => Number(e[col]) === num).length;
}
/**
 * La cuenta que SI autoriza el borrado: pregunta a SharePoint por TODOS los embarques que citan
 * al renglon, no por los de la ventana de 90 dias. Un carrier citado hace un anio se veria en
 * cero desde la memoria y el boton ofreceria Eliminar sobre historia viva (cubeta 3).
 * Si la consulta falla, esta funcion no contesta cero: la excepcion sube y el borrado se niega.
 */
async function referenciasPadronVivas(clave, x) {
    const citas = await estado.cliente.renglones(estado.siteId, L.embarques, `fields/${COLUMNA_PADRON[clave]} eq ${Number(x.id)}`);
    return referenciasSinEmbarques(clave, x) + citas.length;
}
function botonesPadron(clave, x) {
    const editar = x.Activo !== false && PUEDE.capturarPrealta(estado.rol) ? [{ texto: 'Editar', accion: 'editar', clase: 'suave', alClic: () => abrirFormaPadron(clave, vivo(clave, x)) }] : [];
    if (!PUEDE.corregir(estado.rol)) return editar;
    if (x.Activo === false) return [{ texto: 'Reactivar', accion: 'reactivar', clase: 'suave', alClic: () => activarPadron(clave, vivo(clave, x), true) }];
    return editar.concat(referenciasPadron(clave, x) === 0
        ? [{ texto: 'Eliminar', accion: 'eliminar', clase: 'peligro', alClic: () => eliminarPadron(clave, vivo(clave, x)) }]
        : [{ texto: 'Dar de baja', accion: 'baja', clase: 'peligro', alClic: () => activarPadron(clave, vivo(clave, x), false) }]);
}
const NOMBRE_PADRON = { carriers: 'carrier', unidades: 'unidad', choferes: 'chofer' };
function vigenciasDelCarrier(c) { return estado.vigencias.filter(v => v.Rol === 'carrier' && String(v.Title).endsWith(`· ${c.Title}`)); }
// C-05 (v0.23.0): el SEGUNDO paso de una escritura en dos pasos (carrier + su vigencia del tablero) corre aparte: si falla,
// el primero SI quedo y el aviso lo dice en ambar. Antes salia «No se pudo guardar» sobre un carrier ya guardado, el
// operador lo reintentaba y lo duplicaba. Devuelve el mensaje del error, o null si el paso paso.
async function segundoPaso(fn) { try { await fn(); return null; } catch (e) { return e && e.message ? e.message : String(e); } }
async function eliminarPadron(clave, x) {
    let citas;
    try { citas = await referenciasPadronVivas(clave, x); }
    catch (err) { avisar('No pude confirmar quién lo cita (' + err.message + '). No se elimina: dale de baja.', 'error'); return; }
    if (citas > 0) { avisar('Ya hay pre-altas o embarques que lo citan (incluso fuera de los últimos ' + CONFIG.ventanaDias + ' días): se da de baja, no se elimina.', 'error'); return; }
    const { ok } = await confirmar({ titulo: `Eliminar ${NOMBRE_PADRON[clave]}`, peligro: true, ok: 'Eliminar',
        texto: `«${x.Title}» no lo cita ninguna pre-alta ni embarque. Se borra del padrón${clave === 'carriers' ? ' junto con su vigencia ASEA' : ''}; si hacía falta corregirlo, se transcribe de nuevo del oficio.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // C-05 (v0.23.0): las vigencias ANTES que el carrier. Si falla a medias queda un carrier sin vigencia (se ve en el
        // padron y se reintenta), no una vigencia huerfana pintandose en «Hoy» sin carrier.
        if (clave === 'carriers') for (const v of vigenciasDelCarrier(x)) { await estado.cliente.borrarRenglon(estado.siteId, L.vigencias, v.id); estado.vigencias = estado.vigencias.filter(y => y.id !== v.id); }
        await estado.cliente.borrarRenglon(estado.siteId, L[clave], x.id);
        estado[clave] = estado[clave].filter(y => y.id !== x.id);
        avisar(`${NOMBRE_PADRON[clave]} eliminado.`, 'bien'); pintarPadron();
    } catch (e) { avisar('No se pudo eliminar: ' + e.message, 'error'); }
}
async function activarPadron(clave, x, activo) {
    const { ok, motivo } = await confirmar(activo
        ? { titulo: `Reactivar ${NOMBRE_PADRON[clave]}`, ok: 'Reactivar', texto: `«${x.Title}» vuelve a ampararse en la puerta y a salir en las listas.` }
        : { titulo: `Dar de baja ${NOMBRE_PADRON[clave]}`, peligro: true, ok: 'Dar de baja', motivo: 'opcional', etiquetaMotivo: 'Por qué (queda en Notas)',
            texto: `«${x.Title}» ya lo citan pre-altas o embarques, así que no se borra: deja de ampararse en la puerta y de salir en las listas, y el historial lo sigue viendo.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        const campos = { Activo: activo };
        // C-06 (v0.22.0): el motivo queda en Notas para las TRES claves, como promete el dialogo (unidades y choferes tienen
        // Notas desde la tarea 9). Si la lista aun no trae la columna, se da de baja igual, sin motivo, y se dice.
        if (!activo && motivo) campos.Notas = `${x.Notas ? x.Notas + '\n' : ''}BAJA ${fechaMexico()} por ${estado.cuenta.username}: ${motivo}`;
        let sinNotas = false;
        try { await estado.cliente.actualizarRenglon(estado.siteId, L[clave], x.id, campos); }
        catch (e) {
            if (!campos.Notas || !esColumnaFaltante(e)) throw e;   // C-15
            delete campos.Notas; sinNotas = true;
            await estado.cliente.actualizarRenglon(estado.siteId, L[clave], x.id, campos);
        }
        aplicar(clave, x, campos);   // C-23
        const pendiente = clave === 'carriers' ? await segundoPaso(async () => { for (const v of vigenciasDelCarrier(x)) { await estado.cliente.actualizarRenglon(estado.siteId, L.vigencias, v.id, { Activo: activo }); v.Activo = activo; } }) : null;
        if (pendiente) avisar(`${NOMBRE_PADRON[clave]} ${activo ? 'reactivado' : 'dado de baja'}, pero su vigencia ASEA del tablero no cambió (${pendiente}). Edítalo y guarda para sincronizarla.`, 'ojo');
        else if (activo) avisar(`${NOMBRE_PADRON[clave]} reactivado.`, 'bien');
        else if (sinNotas) avisar(`${NOMBRE_PADRON[clave]} dado de baja, pero el motivo NO quedó` + (estado.rol === 'gerencia' ? ': la lista no tiene todavía la columna Notas (setup, tarea 9).' : '. Avisa a gerencia.'), 'ojo');   // U-31
        else avisar(`${NOMBRE_PADRON[clave]} dado de baja${motivo ? ' · el motivo quedó en Notas' : ''}.`, 'bien');
        pintarPadron();
    } catch (e) { avisar('No se pudo cambiar: ' + e.message, 'error'); }
}

// ---- alta y edición del padrón. Un solo formulario por tipo: vacío da de alta, con `estado.padronEdit` edita ese renglón.
// Editar es un PATCH sobre el mismo renglón (SharePoint guarda la versión anterior): un dedazo en la póliza no obliga a
// dar de baja y transcribir de nuevo. Lo que se vacía en el formulario se borra en la lista (null), no se conserva.
const FORMA_PADRON = {
    carriers: { forma: 'pdFormaCarrier', grupo: 'pdGrupoCarriers', titulo: 'carrier', campos: ['pcTitle', 'pcAut', 'pcVig', 'pcFolio', 'pcSCT', 'pcCSF'] },
    unidades: { forma: 'pdFormaUnidad', grupo: 'pdGrupoUnidades', titulo: 'unidad', campos: ['puuPlaca', 'puuPlana', 'puuFolio', 'puuCap', 'puuSerie', 'puuMarca', 'puuTarjeta', 'puuTarjetaVig', 'puuPoliza', 'puuPolizaVig'] },
    choferes: { forma: 'pdFormaChofer', grupo: 'pdGrupoChoferes', titulo: 'chofer', campos: ['pchNombre', 'pchLic', 'pchLicVig'] },
};
function leerFormaPadron(clave) {
    if (clave === 'carriers') return {
        Title: $('pcTitle').value.trim(), AutorizacionASEA: $('pcAut').value.trim(), VigenciaASEA: aIsoDia($('pcVig').value),
        FolioOficio: $('pcFolio').value.trim(), Corrientes: [...document.querySelectorAll('input[name="pcCorr"]:checked')].map(c => c.value).join('; '),
        RegistroSCT: $('pcSCT').value.trim(), CSFVigencia: aIsoDia($('pcCSF').value), Activo: true };
    if (clave === 'unidades') return {
        Title: placaNormal($('puuPlaca').value), PlacaPlana: placaNormal($('puuPlana').value), CarrierId: Number($('puuCarrier').value),
        TipoUnidad: $('puuTipo').value, FolioOficio: folioUnidad($('puuFolio').value, $('puuCarrier').value), CapacidadKg: $('puuCap').value ? Number($('puuCap').value) : null,
        NumeroSerie: $('puuSerie').value.trim(), Marca: $('puuMarca').value.trim(), TarjetaCirc: $('puuTarjeta').value.trim(), TarjetaVigencia: aIsoDia($('puuTarjetaVig').value),
        Poliza: $('puuPoliza').value.trim(), PolizaVigencia: aIsoDia($('puuPolizaVig').value), Activo: true };
    return {
        Title: $('pchNombre').value.trim(), CarrierId: Number($('pchCarrier').value), Licencia: $('pchLic').value.trim(),
        LicenciaVigencia: aIsoDia($('pchLicVig').value), Activo: true };
}
// v0.19.9 (Carlos, 2026-09-08): el folio por unidad era redundante — cada placa se transcribe del oficio del carrier.
// Vacio hereda el folio del carrier; se teclea solo cuando la unidad la ampara un alcance distinto.
function folioUnidad(tecleado, carrierId) {
    const t = String(tecleado || '').trim(); if (t) return t;
    const c = porId(estado.carriers, carrierId); return c && c.FolioOficio ? String(c.FolioOficio).trim() : '';
}
function validarFormaPadron(clave) {
    if (clave === 'carriers' && !$('pcTitle').value.trim()) return 'Falta la razón social.';
    if (clave === 'unidades' && (!$('puuCarrier').value || !$('puuPlaca').value.trim())) return 'Carrier y placa son obligatorios.';
    if (clave === 'choferes' && (!$('pchCarrier').value || !$('pchNombre').value.trim())) return 'Carrier y nombre son obligatorios.';
    return null;
}
function llenarFormaPadron(clave, x) {
    const f = textoDe;   // C-26
    const fecha = v => v ? fechaCorta(v) : '';
    if (clave === 'carriers') {
        $('pcTitle').value = f(x.Title); $('pcAut').value = f(x.AutorizacionASEA); $('pcVig').value = fecha(x.VigenciaASEA); $('pcFolio').value = f(x.FolioOficio);
        $('pcSCT').value = f(x.RegistroSCT); $('pcCSF').value = fecha(x.CSFVigencia);
        const corr = lista(x.Corrientes); for (const c of document.querySelectorAll('input[name="pcCorr"]')) c.checked = corr.includes(c.value);
    } else if (clave === 'unidades') {
        $('puuCarrier').value = f(x.CarrierId); $('puuPlaca').value = f(x.Title); $('puuPlana').value = f(x.PlacaPlana); $('puuTipo').value = f(x.TipoUnidad) || 'gondola';
        $('puuFolio').value = f(x.FolioOficio); $('puuCap').value = f(x.CapacidadKg); $('puuSerie').value = f(x.NumeroSerie); $('puuMarca').value = f(x.Marca);
        $('puuTarjeta').value = f(x.TarjetaCirc); $('puuTarjetaVig').value = fecha(x.TarjetaVigencia); $('puuPoliza').value = f(x.Poliza); $('puuPolizaVig').value = fecha(x.PolizaVigencia);
    } else {
        $('pchCarrier').value = f(x.CarrierId); $('pchNombre').value = f(x.Title); $('pchLic').value = f(x.Licencia); $('pchLicVig').value = fecha(x.LicenciaVigencia);
    }
}
function vaciarFormaPadron(clave) {
    for (const id of FORMA_PADRON[clave].campos) $(id).value = '';
    if (clave === 'carriers') for (const c of document.querySelectorAll('input[name="pcCorr"]')) c.checked = false;
}
function tituloFormaPadron(clave) {
    const e = estado.padronEdit && estado.padronEdit.clave === clave ? estado.padronEdit.x : null;
    $(FORMA_PADRON[clave].forma).querySelector('h2').textContent = e ? `Editar ${FORMA_PADRON[clave].titulo}: ${e.Title}` : `Alta de ${FORMA_PADRON[clave].titulo}`;
}
const SELECT_CARRIER_PADRON = { unidades: 'puuCarrier', choferes: 'pchCarrier' };
function abrirFormaPadron(clave, x = null) {
    estado.padronEdit = x ? { clave, x } : null;
    if (x) llenarFormaPadron(clave, x); else vaciarFormaPadron(clave);
    // U-46 (v0.29.0): el alta de unidad/chofer arranca con el carrier que el padrón ya tiene elegido (el filtro «Mostrando solo
    // lo de X») o, si no hay filtro, el último que se guardó en esta sesión. Transcribir 8 placas de un oficio eran 8 búsquedas
    // en el select. El value se conserva además entre altas consecutivas (opciones(), U-40).
    if (!x && SELECT_CARRIER_PADRON[clave]) {
        const sel = $(SELECT_CARRIER_PADRON[clave]);
        const quiero = estado.padronCarrier !== null && estado.padronCarrier !== undefined ? estado.padronCarrier : estado.ultimoCarrierPadron;
        if (quiero !== null && quiero !== undefined && [...sel.options].some(o => o.value === String(quiero))) sel.value = String(quiero);
    }
    tituloFormaPadron(clave);
    $(FORMA_PADRON[clave].grupo).open = true;
    abrirForma(FORMA_PADRON[clave].forma);
}
function cerrarFormaPadron(clave) {
    estado.padronEdit = null; tituloFormaPadron(clave);
    cerrarForma(FORMA_PADRON[clave].forma);
}
const AVISO_ALTA = { carriers: 'Carrier dado de alta. Ahora sus unidades, transcritas del oficio.', unidades: 'Unidad transcrita.', choferes: 'Chofer dado de alta.' };
const BOTON_GUARDAR_PADRON = { carriers: 'btnGuardarCarrier', unidades: 'btnGuardarUnidad', choferes: 'btnGuardarChofer' };
async function guardarPadron(clave) {
    if (!PUEDE.capturarPrealta(estado.rol)) { avisar('Tu rol no edita el padrón.', 'error'); return; }
    const falta = validarFormaPadron(clave); if (falta) { avisar(falta, 'error'); return; }
    const edit = estado.padronEdit && estado.padronEdit.clave === clave ? estado.padronEdit.x : null;
    let campos;
    try { campos = leerFormaPadron(clave); } catch (e) { avisar(e.message, 'error'); return; }
    await escribiendo(BOTON_GUARDAR_PADRON[clave], async () => { try {   // C-24: un doble toque creaba dos carriers / unidades / choferes
        await refrescarCliente();
        if (edit) {
            delete campos.Activo;   // editar no reactiva ni da de baja
            // Cambiar una vigencia pide motivo (queda en Notas): editar tambien sirve para «arreglar» una poliza sin
            // tener el papel, y la app no distingue eso de una correccion legitima — el rastro lo deja quien captura.
            const dia = v => (v ? String(v).slice(0, 10) : '');
            const cambios = VIGENCIAS_PADRON[clave].filter(([, col]) => dia(campos[col]) !== dia(edit[col])).map(([n, col]) => `${n}: ${fechaCorta(edit[col])} → ${fechaCorta(campos[col])}`);
            if (cambios.length) {
                const { ok, motivo } = await confirmar({ titulo: 'Cambio de vigencia', ok: 'Guardar', motivo: true, etiquetaMotivo: 'Por qué cambia (queda en Notas)',
                    texto: `${cambios.join(' · ')}. La versión anterior queda en SharePoint; el motivo, en las Notas del ${NOMBRE_PADRON[clave]}.` });
                if (!ok) return;
                campos.Notas = `${edit.Notas ? edit.Notas + '\n' : ''}VIGENCIA ${fechaMexico()} por ${estado.cuenta.username}: ${cambios.join(' · ')} — ${motivo}`;
            }
            const vsAntes = clave === 'carriers' ? vigenciasDelCarrier(edit) : [];
            try { await estado.cliente.actualizarRenglon(estado.siteId, L[clave], edit.id, paraPatch(campos)); }
            catch (e) { if (campos.Notas && esColumnaFaltante(e)) throw new Error('la vigencia no se puede cambiar todavía' + (estado.rol === 'gerencia' ? ': la lista no tiene la columna Notas (tarea 9 de setup-carlos.md)' : '; avisa a gerencia')); throw e; }   // U-31
            aplicar(clave, edit, campos);   // C-23
            // La vigencia ASEA del tablero se llama por el carrier y guarda su fecha: se corrige junto con él. Es el segundo
            // paso (C-05): si falla, el carrier ya quedo y el aviso lo dice; guardar de nuevo la sincroniza (Activo incluido).
            const pendiente = clave === 'carriers' ? await segundoPaso(async () => {
                if (vsAntes.length) for (const v of vsAntes) { const c = { Title: `Autorización ASEA transporte · ${edit.Title}`, Vence: edit.VigenciaASEA || null, Folio: edit.FolioOficio || null, Activo: edit.Activo !== false }; await estado.cliente.actualizarRenglon(estado.siteId, L.vigencias, v.id, c); aplicar('vigencias', v, c); }
                else if (edit.VigenciaASEA) await altaVigencia(`Autorización ASEA transporte · ${edit.Title}`, 'tercero', 'carrier', 'legal', edit.VigenciaASEA, edit.FolioOficio);
            }) : null;
            cerrarFormaPadron(clave);
            if (pendiente) avisar(`${NOMBRE_PADRON[clave]} actualizado, pero su vigencia ASEA del tablero no se sincronizó (${pendiente}). Guárdalo de nuevo para reintentar.`, 'ojo');
            else avisarAlta(`${NOMBRE_PADRON[clave]} actualizado.`, clave, edit);
            pintarPadron();
        } else {
            const nuevo = await estado.cliente.crearRenglon(estado.siteId, L[clave], limpiar(campos));
            anclar(clave, nuevo);   // C-23
            estado.ultimoCarrierPadron = clave === 'carriers' ? nuevo.id : (nuevo.CarrierId ? Number(nuevo.CarrierId) : estado.ultimoCarrierPadron);   // U-46
            const pendiente = clave === 'carriers' && nuevo.VigenciaASEA ? await segundoPaso(() => altaVigencia(`Autorización ASEA transporte · ${nuevo.Title}`, 'tercero', 'carrier', 'legal', nuevo.VigenciaASEA, nuevo.FolioOficio)) : null;
            vaciarFormaPadron(clave); cerrarFormaPadron(clave);
            if (pendiente) avisar(`${NOMBRE_PADRON[clave]} dado de alta, pero su vigencia ASEA NO quedó en el tablero (${pendiente}). Edítalo y guarda para crearla.`, 'ojo');
            else avisarAlta(AVISO_ALTA[clave], clave, nuevo);
            // U-10 (v0.23.0): si el alta vino desde la pre-alta (que sigue abierta atras), el carrier nuevo queda elegido en ella.
            if (clave === 'carriers' && $('paForma').open) elegirCarrierEnPrealta(nuevo);
            pintarPadron();
        }
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}
async function altaVigencia(titulo, titular, rol, fuente, vence, folio) {
    const v = await estado.cliente.crearRenglon(estado.siteId, L.vigencias, limpiar({ Title: titulo, Titular: titular, Rol: rol, Fuente: fuente, Vence: vence, Folio: folio, AvisoDias: CONFIG.avisoVigenciaDias, Activo: true }));
    estado.vigencias.push(v);
}

// ================================================================ HOY (la consola)

// C-07 (v0.23.0): pintarHoy se parte por tarjeta (franja, KPI, fila del dia, pendientes, rechazos, vigencias) y la fila
// del dia se pinta dos veces —tarjetas en celular, tabla en escritorio— con ESTAS mismas piezas. Antes cada vista traia
// su copia de ETAPAS, los segmentos y la etiqueta de compuerta, y un cambio de etapa habia que hacerlo en dos sitios.
const ETAPAS_FILA = { compuerta: 1, bruto: 2, cerrado: 3, rechazado: 0, anulado: 0 };
function segmentosEtapa(e) {
    const seg = el('span', 'etapa');
    for (let i = 0; i < 3; i++) { const s = el('i'); if (e.Etapa === 'rechazado' && i === 0) s.className = 'x'; else if (i < (ETAPAS_FILA[e.Etapa] ?? 0)) s.className = 'f'; seg.appendChild(s); }
    return seg;
}
function etiquetaCompuertaDe(e) {
    if (e.Etapa === 'anulado') return etiqueta('anulado', 'anulado');
    return etiqueta(e.Compuerta === 'pasa' ? 'pasa' : e.Compuerta === 'rechazo-legal' ? 'no entró' : excepcionAutorizada(e) ? 'excepción ok' : 'espera',
                    e.Compuerta === 'pasa' ? 'pasa' : e.Compuerta === 'rechazo-legal' ? 'rechazo-legal' : 'excepcion-comercial');
}
/** «neto 21,220» / «bruto 44,600» / null; `rotuloNeto` vacio deja el neto solo (la tabla ya tiene la columna). */
function pesoFila(e, rotuloNeto = 'neto ') {
    return e.NetoKg ? `${rotuloNeto}${Number(e.NetoKg).toLocaleString('es-MX')}` : e.BrutoKg ? `bruto ${Number(e.BrutoKg).toLocaleString('es-MX')}` : null;
}
/** Ticket (si hay folio) + Anular/Eliminar. `detener` frena la propagacion: en la tarjeta el clic tambien la despliega. */
function botonesFila(e, abrirTicket, detener) {
    const bf = el('div', 'botones-fila');
    if (e.Title) { const b = el('button', '', 'Ticket'); b.type = 'button'; b.addEventListener('click', ev => { if (detener) ev.stopPropagation(); abrirTicket(e); }); bf.appendChild(b); }
    for (const b of botonCorreccion(e)) { const x = botonAccion(b); if (detener) x.addEventListener('click', ev => ev.stopPropagation()); bf.appendChild(x); }
    return bf;
}
const horaFila = e => horaMexico(e.Arribo, 'hora');

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
function pintarKpisReportes({ cerradosHoy, cerradosAyer, cerradosSemana, activos, rechazosSemana, borradores }) {
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

// ================================================================ REPORTES (v0.32.0, sección aparte; artifact 1GvBJaYooYvjZT4rMRtL9Q)
// Los cinco KPI que vivían en Hoy, más lo que se lee de lo cargado: avance por programa, por carrier, toneladas por semana,
// rechazos/excepciones y netos fuera de banda. Todo sale de estado.embarques (CONFIG.ventanaDias más lo abierto): no lee nada más.
const mesCorto = f => new Date(f + 'T12:00:00Z').toLocaleDateString('es-MX', { timeZone: 'UTC', month: 'short' }).replace('.', '');
function pintarReportes() {
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
    if (!rech.length) rj.appendChild(el('p', 'vacio', 'Ninguno en lo cargado.'));
    for (const e of rech) rj.appendChild(renglonRechazo(e));
    const nt = $('repNetos'); nt.textContent = '';
    const fuera = estado.embarques.filter(e => e.Etapa === 'cerrado' && /Neto fuera de banda/.test(e.Notas || '')).sort((a, b) => b.id - a.id);
    if (!fuera.length) nt.appendChild(el('p', 'vacio', `Ninguno: las ${cerrados(() => true).length} cerradas quedaron dentro de la banda.`));
    for (const e of fuera) nt.appendChild(renglon(`${e.Title} · ${e.PlacaTractor}`, `${fechaCorta(e.TaraHora || e.Arribo)} · neto ${Number(e.NetoKg).toLocaleString('es-MX')} kg · ${(e.Notas || '').replace(/\n.*$/s, '')}`));
}

// ================================================================ ARCHIVOS (v0.33.0, sección aparte; artifact 1GvBJaYooYvjZT4rMRtL9Q)
// El árbol de la biblioteca Ambiental-CALYTEK leído por Graph (Sites.Selected; la misma ruta drive/root: con la que sube la
// evidencia): tres ramas —evidencia de báscula (carpeta por mes → lote por pesaje → foto, ticket, manifiesto), oficios ASEA y
// CSF de los carriers, y el buzón 99_Pendiente-Archivar, donde la app deja los lotes hasta que /archivar-calytek los acomode—.
// Cada carpeta se lee al abrirla (una llamada, paginada) y se recuerda hasta el siguiente Actualizar a mano. Los filtros
// (programa · tipo · buscador) actúan sobre lo ya leído: una carpeta cerrada no se juzga porque no se conoce. Solo lectura.
const RAMAS_ARCHIVOS = () => [
    { ruta: CONFIG.evidencia.destinoBase, titulo: 'Evidencia de báscula', nota: 'por mes · un lote por pesaje' },
    { ruta: CONFIG.archivos.padron, titulo: 'Oficios ASEA y CSF de los carriers', nota: '' },
    { ruta: CONFIG.buzon, titulo: 'Pendiente de archivar', nota: 'lo que la app acaba de subir', soloLotes: true }   // S-15: solo los lotes de la app, no el buzón entero
];
const ICONO_CARPETA = 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z';
const TIPOS_ARCHIVO = { ticket: 'ticket', foto: 'foto del indicador', manifiesto: 'manifiesto', oficio: 'oficio ASEA', csf: 'CSF', lote: 'registro que deja la app al pesar', otro: 'otro' };
function svgIcono(d, clase) {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('aria-hidden', 'true'); if (clase) s.setAttribute('class', clase);
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); s.appendChild(p);
    return s;
}
const tamano = n => n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : n >= 1024 ? `${Math.round(n / 1024)} KB` : `${n || 0} B`;
async function hijosDe(ruta) {
    const a = estado.archivos;
    if (!a) throw new Error('la biblioteca se está releyendo; abre la carpeta de nuevo');   // C-34 (v0.34.0)
    if (!a.ramas.has(ruta)) a.ramas.set(ruta, await estado.cliente.hijos(estado.siteId, ruta, m => avisar(m, 'ojo')));
    return a.ramas.get(ruta);
}
/** C-31 (v0.34.0): olvida lo leído de una rama y, si su carpeta está abierta en el árbol, la relee ya. La llama subirEvidencia. */
function invalidarRama(ruta) {
    if (estado.archivos) estado.archivos.ramas.delete(ruta);
    for (const d of document.querySelectorAll('#arArbol details')) if (d.dataset.ruta === ruta) { delete d.dataset.leida; if (d.open && d.leer) d.leer(); }
}
function pintarArchivos() {
    // el filtro de programa sale de las pre-altas cargadas (todas: una cerrada sigue teniendo sus lotes en la biblioteca)
    opciones($('arPrograma'), [...estado.prealtas].sort((x, y) => y.id - x.id), p => p.id, p => `${p.Title}${p.Estado === 'cerrada' ? ' · cerrada' : ''}`, 'Todos los programas');
    if (!estado.archivos) { estado.archivos = { biblioteca: null, ramas: new Map() }; construirArbol().catch(err => avisar('No se pudo armar el árbol: ' + err.message, 'error')); }   // C-34: sin catch, un throw dejaba «Abriendo…» para siempre
    else filtrarArbol();
}
async function construirArbol() {
    const a = estado.archivos;
    const caja = $('arArbol'); caja.textContent = ''; caja.appendChild(el('p', 'vacio', 'Abriendo la biblioteca…'));
    const lnk = $('lnkArchivosSharePoint'); lnk.href = `https://${CONFIG.sharepointHost}${CONFIG.sitio}`;
    try { a.biblioteca = await estado.cliente.biblioteca(estado.siteId); if (a.biblioteca.webUrl) lnk.href = a.biblioteca.webUrl; }
    catch (err) { avisar('No se pudo abrir la biblioteca: ' + err.message, 'error'); }
    if (a !== estado.archivos) return;   // hubo un Actualizar en medio: lo pinta la lectura nueva
    caja.textContent = '';
    for (const r of RAMAS_ARCHIVOS()) caja.appendChild(nodoCarpeta({ name: r.ruta }, r.ruta, r.titulo, r.nota, r.soloLotes));
    caja.firstChild.open = true;   // la evidencia abierta de entrada, como en el artifact
}
/** Una carpeta del árbol: <details> que lee sus hijos por Graph la primera vez que se abre. `titulo` solo en las tres raíces;
 *  `soloLotes` (S-15) deja ver únicamente las carpetas de lote de la app. `d.leer` la relee (C-31, invalidarRama). */
function nodoCarpeta(item, ruta, titulo, nota, soloLotes = false) {
    const d = el('details'); d.dataset.ruta = ruta; d.dataset.nombre = normaliza(titulo ? `${titulo} ${item.name}` : item.name);
    const s = el('summary'); s.appendChild(svgIcono('M9 6l6 6-6 6', 'flecha')); s.appendChild(svgIcono(ICONO_CARPETA));
    s.appendChild(el('span', '', titulo || item.name));
    if (nota) s.appendChild(el('span', 'sub', nota));
    if (titulo) s.title = ruta;
    const n = el('span', 'n', item.folder && item.folder.childCount !== undefined ? String(item.folder.childCount) : ''); s.appendChild(n);
    d.appendChild(s);
    const hijos = el('div', 'hijos'); d.appendChild(hijos);
    let leyendo = false;
    d.leer = async () => {
        if (leyendo || d.dataset.leida === '1') return;
        leyendo = true;
        hijos.textContent = ''; hijos.appendChild(el('p', 'vacio', 'Leyendo…'));
        let items;
        try { items = await hijosDe(ruta); }
        catch (err) { hijos.textContent = ''; hijos.appendChild(el('p', 'vacio', 'No se pudo leer: ' + err.message)); leyendo = false; return; }
        leyendo = false;
        hijos.textContent = ''; d.dataset.leida = '1';   // desde aquí el filtro sí la juzga
        if (items === null) { hijos.appendChild(el('p', 'vacio', 'Esta carpeta aún no existe en la biblioteca.')); n.textContent = '—'; return; }
        if (soloLotes) items = items.filter(it => it.folder && esLoteDeLaApp(it.name, CONFIG.evidencia.etiqueta));   // S-15
        if (!items.length) hijos.appendChild(el('p', 'vacio', soloLotes ? 'Sin lotes de la app.' : 'Vacía.'));
        // carpetas primero y lo más reciente arriba: los nombres de la casa empiezan por la fecha
        const orden = items.slice().sort((x, y) => ((y.folder ? 1 : 0) - (x.folder ? 1 : 0)) || y.name.localeCompare(x.name, 'es'));
        for (const it of orden) hijos.appendChild(it.folder ? nodoCarpeta(it, `${ruta}/${it.name}`) : nodoArchivo(it));
        const nc = orden.filter(x => x.folder).length, na = orden.length - nc;
        n.textContent = [nc ? plural(nc, soloLotes ? 'lote' : 'carpeta') : '', na ? plural(na, 'archivo') : ''].filter(Boolean).join(' · ') || '0';
        filtrarArbol();
    };
    d.addEventListener('toggle', () => { if (d.open) d.leer(); });
    return d;
}
/** Un archivo del árbol: enlace a SharePoint (abre en otra pestaña), extensión, nombre y «fecha · tamaño» en mono. */
function nodoArchivo(it) {
    const a = el('a', 'arch'); a.href = it.webUrl || '#'; a.target = '_blank'; a.rel = 'noopener';
    const tipo = tipoDeArchivo(it.name); a.dataset.tipo = tipo; a.dataset.nombre = normaliza(it.name);
    const ext = (String(it.name).match(/\.([a-z0-9]{1,4})$/i) || [])[1] || '';
    a.appendChild(el('span', 'ext', ext.toUpperCase()));
    a.appendChild(el('span', '', it.name));
    a.title = `${it.name} · ${TIPOS_ARCHIVO[tipo]}`;
    const dd = el('span', 'd', fechaCorta(it.lastModifiedDateTime)); dd.appendChild(el('span', 'tam', ` · ${tamano(Number(it.size) || 0)}`)); a.appendChild(dd);   // U-67: en celular la fecha baja de renglón y el tamaño se oculta
    return a;
}
/** Aplica los tres filtros sobre lo ya leído: un archivo pega por nombre y tipo; una carpeta leída se esconde si no le queda nada visible. */
function filtrarArbol() {
    const q = normaliza($('arBusca').value.trim());
    const tipo = $('arTipo').value;
    const pre = porId(estado.prealtas, $('arPrograma').value);
    // programa = los folios de sus góndolas (E-26-00012 → e-26-00012, que es como viaja en el nombre del lote y de la foto)
    const folios = pre ? estado.embarques.filter(e => Number(e.PreAltaId) === pre.id && e.Title).map(e => normaliza(e.Title)) : [];
    const pegaNombre = nombre => (!q || nombre.includes(q)) && (!pre || folios.some(f => nombre.includes(f)));
    const hayFiltro = !!(q || tipo || pre);
    const caja = $('arArbol');
    for (const a of caja.querySelectorAll('.arch')) a.classList.toggle('oculto', hayFiltro && !(pegaNombre(a.dataset.nombre) && (!tipo || a.dataset.tipo === tipo)));
    for (const d of [...caja.querySelectorAll('details')].reverse()) {   // de adentro hacia afuera: el padre juzga hijos ya juzgados
        const alguno = [...d.querySelectorAll(':scope > .hijos > .arch, :scope > .hijos > details')].some(x => !x.classList.contains('oculto'));
        const propio = !tipo && pegaNombre(d.dataset.nombre);
        d.classList.toggle('oculto', hayFiltro && d.dataset.leida === '1' && !alguno && !propio);   // una carpeta sin leer no se juzga
    }
    $('arVacio').classList.toggle('oculto', !hayFiltro || [...caja.children].some(x => !x.classList.contains('oculto')));
}

// Fila del dia: los embarques de hoy; si no hay, los ultimos 5. Tarjetas (celular) y tabla (escritorio) de la MISMA lista
// y con las mismas piezas de arriba. I3 (7-sep): en celular la tabla de 8 columnas escondia seis; tocar la tarjeta muestra sus botones.
function pintarFilaDia(hoy, dia) {
    const deHoy = estado.embarques.filter(e => dia(e) === hoy).sort((a, b) => a.id - b.id);
    const fila = deHoy.length ? deHoy : [...estado.embarques].sort((a, b) => b.id - a.id).slice(0, 5);
    // v0.19.2 (Carlos, 8-sep): Ticket abre la misma ventana que Cerrados, sin salir de Hoy; ‹ › recorren los de la fila con folio.
    const conFolio = fila.filter(x => x.Title);
    const abrirTicketDeHoy = e => abrirTicketPop(conFolio, conFolio.indexOf(e));
    const tw = $('tbFila'); tw.textContent = '';
    const tt = $('tbFilaTarjetas'); tt.textContent = '';
    if (!fila.length) { tw.appendChild(el('p', 'vacio', 'Ninguna góndola registrada todavía.')); tt.appendChild(el('p', 'vacio', 'Ninguna góndola registrada todavía.')); return; }
    if (!deHoy.length) { tw.appendChild(el('p', 'vacio', 'Hoy no ha llegado ninguna; estos son los últimos.')); tt.appendChild(el('p', 'vacio', 'Hoy no ha llegado ninguna; estos son los últimos.')); }
    for (const e of fila) {
        const t = el('div', 'ftar' + (e.Etapa === 'anulado' ? ' anulado' : ''));
        // U-58 (v0.29.0): la tarjeta era role=button con botones adentro (interactivo anidado, deuda de U-15). Ahora el que
        // despliega es el FOLIO, un <button> con aria-expanded y aria-controls a la zona de botones; la tarjeta entera sigue
        // abriendo al toque (sin role ni tabindex) porque con guante el dedo no apunta al folio.
        const bf = botonesFila(e, abrirTicketDeHoy, true);
        const f = el(bf.childElementCount ? 'button' : 'span', 'f' + (e.Title ? '' : ' mudo'), e.Title || 'sin folio');
        t.appendChild(f);
        t.appendChild(segmentosEtapa(e));
        const c = el('span', 'c');
        c.textContent = [e.PlacaTractor || '—', nombreDe(estado.carriers, e.CarrierId), pesoFila(e), horaFila(e)].filter(Boolean).join(' · ');
        c.appendChild(etiquetaCompuertaDe(e));
        t.appendChild(c);
        if (bf.childElementCount) {
            bf.id = `ftarBotones${e.id}`; t.appendChild(bf);
            f.type = 'button'; f.setAttribute('aria-expanded', 'false'); f.setAttribute('aria-controls', bf.id);
            const alternar = () => { t.classList.toggle('abierta'); f.setAttribute('aria-expanded', String(t.classList.contains('abierta'))); };
            f.addEventListener('click', ev => { ev.stopPropagation(); alternar(); });
            t.addEventListener('click', alternar);
        }
        tt.appendChild(t);
    }
    const t = el('table', 'fila'); const th = el('tr');
    for (const c of ['Folio', 'Unidad', 'Transportista', 'Etapa', 'Compuerta', 'Neto kg', 'Arribo', '']) th.appendChild(el('th', '', c));
    t.appendChild(th);
    for (const e of fila) {
        const tr = el('tr', e.Etapa === 'anulado' ? 'anulado' : '');
        tr.appendChild(el('td', 'mono' + (e.Title ? '' : ' mudo'), e.Title || '—'));
        tr.appendChild(el('td', 'mono', e.PlacaTractor || '—'));
        tr.appendChild(el('td', '', nombreDe(estado.carriers, e.CarrierId)));
        const et = el('td'); et.appendChild(segmentosEtapa(e)); tr.appendChild(et);
        const co = el('td'); co.appendChild(etiquetaCompuertaDe(e)); tr.appendChild(co);
        tr.appendChild(el('td', 'num mono' + (e.NetoKg ? '' : ' mudo'), pesoFila(e, '') || '—'));
        tr.appendChild(el('td', 'mono', horaFila(e)));
        const ac = el('td'); ac.appendChild(botonesFila(e, abrirTicketDeHoy, false)); tr.appendChild(ac);
        t.appendChild(tr);
    }
    tw.appendChild(t);
}

// Pendiente revisar: pre-altas por firmar, excepciones y programas dormidos. Con algo, la tarjeta se pinta en ambar
// con el conteo en rojo (Carlos, 2026-09-08: es lo primero que hay que atender).
function pintarPendientesHoy(borradores, pendientes, botonesExcepcion) {
    const pf = $('tbPendientes'); pf.textContent = '';
    const dormidas = estado.prealtas.filter(p => p.Estado === 'firmada').map(p => ({ p, sm: sinMovimientoDe(p) })).filter(x => x.sm);
    // S-01: sellos sin firma en PLANTA_Firmas (la compuerta no los acepta) y la lista misma si no esta provisionada.
    const ssf = sellosSinFirma().prealtas;
    const nPend = borradores.length + pendientes.length + dormidas.length + ssf.length + (estado.firmasError ? 1 : 0);
    $('tbPendientesTarjeta').classList.toggle('alerta', nPend > 0);
    $('tbPendientesN').classList.toggle('oculto', !nPend); $('tbPendientesN').textContent = String(nPend);
    if (!nPend) pf.appendChild(el('p', 'vacio', 'Nada pendiente.'));
    if (estado.firmasError) pf.appendChild(renglon('No se pudo leer el registro de firmas', `la app no firma ni autoriza y ningún sello vale sin su firma · Actualiza; si sigue, avisa a gerencia${estado.rol === 'gerencia' ? ` · PLANTA_Firmas: ${estado.firmasError} (permisos de la lista; setup-carlos.md, tarea 11)` : ''}`));   // U-31
    for (const p of ssf) pf.appendChild(renglon(`Firma · ${p.Title}`, `falta la firma del validador (trae el sello de ${quien(p.FirmadaPor) || '?'}, sin firma registrada) · la puerta no la ve · se firma desde su detalle`, 'Ver', () => verPrealta(vivo('prealtas', p))));   // U-28 / U-31
    for (const { p, sm } of dormidas) pf.appendChild(renglon(`Programa · ${p.Title}`, `${sm.motivo} · ¿se cierra? Sigue saliendo en la puerta`, 'Ver', () => verPrealta(vivo('prealtas', p))));
    for (const p of borradores) { const d = diasPara(p.FechaEstimada); pf.appendChild(renglon(`Pre-alta · ${p.Title}`, `firma del validador · 1er envío ${fechaCorta(p.FechaEstimada)}${d !== null ? ` (en ${d} días)` : ''} · capturó ${quien(p.CapturadaPor) || '?'}`, 'Ver', () => verPrealta(vivo('prealtas', p)))); }
    for (const e of pendientes) pf.appendChild(renglon(`Excepción · ${e.PlacaTractor}`, `${selloSinFirma(e)}autorización de gerencia · «${e.ExcepcionMotivo || 'sin motivo'}» · ${horaCorta(e.Arribo)}`, null, null, botonesExcepcion(e).map(b => ({ ...b, clase: b.accion === 'autorizar' ? '' : 'peligro' }))));
}

// C-29 (v0.34.0): UNA definición de «rechazo o excepción» y UN renglón para Hoy (los 10 últimos) y Reportes (todos). Antes el
// filtro y la lectura de CompuertaDetalle vivían copiados en los dos y la etiqueta difería (Hoy decía legal/comercial).
const esRechazo = e => e.Etapa !== 'anulado' && (e.Etapa === 'rechazado' || e.Compuerta === 'excepcion-comercial');
const rechazosYExcepciones = () => estado.embarques.filter(esRechazo).sort((a, b) => b.id - a.id);
function renglonRechazo(e) {
    let causa = '';
    try { causa = JSON.parse(e.CompuertaDetalle || '[]').filter(h => h.clase === 'legal' || h.clase === 'comercial').map(h => h.regla).join(', '); } catch (_) { /* detalle ilegible */ }
    const r = renglon(`${e.Title || '(excepción)'} · ${e.PlacaTractor} · ${nombreDe(estado.carriers, e.CarrierId)}`, `${horaCorta(e.Arribo)} · ${causa}${e.ExcepcionAutorizo ? ' · autorizó ' + quien(e.ExcepcionAutorizo) : ''}`);
    r.firstChild.firstChild.appendChild(etiquetaCompuertaDe(e));
    return r;
}
function pintarRechazosHoy() {
    const rj = $('tbRechazos'); rj.textContent = '';
    const rech = rechazosYExcepciones().slice(0, 10);
    if (!rech.length) rj.appendChild(el('p', 'pista', 'Ninguno.'));
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
                if (d !== null && d <= CONFIG.avisoVigenciaDias) v.push({ Title: `${NOMBRE_PADRON[tipo].replace(/^./, c => c.toUpperCase())} ${it.Title} · ${n}`, Vence: it[col], Fuente: 'padrón', AvisoDias: CONFIG.avisoVigenciaDias });
            }
        }
    }
    return v;
}
function pintarVigenciasHoy() {
    const vg = $('tbVigencias'); vg.textContent = '';
    const prox = [...estado.vigencias.filter(v => v.Activo !== false), ...vigenciasPadronHoy()].map(v => ({ v, d: diasPara(v.Vence) })).filter(x => x.d !== null && x.d <= (Number(x.v.AvisoDias) || CONFIG.avisoVigenciaDias)).sort((a, b) => a.d - b.d);
    if (!prox.length) vg.appendChild(el('p', 'vacio', `Nada vence en ${CONFIG.avisoVigenciaDias} días.`));
    for (const { v, d } of prox) {
        const ventana = Number(v.AvisoDias) || CONFIG.avisoVigenciaDias;
        const clase = d < 0 ? 'mal' : d <= 7 ? 'ojo' : '';
        const r = el('div', 'vig');
        const t = el('span', '', v.Title); t.appendChild(el('small', '', `${fechaCorta(v.Vence)} · ${v.Fuente || ''}${v.Dueno ? ' · dueño ' + v.Dueno : ''}`)); r.appendChild(t);
        r.appendChild(el('span', 'd ' + clase, d < 0 ? `−${-d} d` : `${d} d`));
        const bar = el('span', 'bar'); const i = el('i', clase); i.style.width = Math.max(4, Math.min(100, Math.round((1 - d / ventana) * 100))) + '%'; bar.appendChild(i); r.appendChild(bar);
        vg.appendChild(r);
    }
}

/** Los cortes de fecha que comparten Hoy y Reportes. Una gondola CERRADA cuenta el dia en que se cerro (hora de la tara),
 *  no el de arribo: la que llega 23:50 y cierra 00:10 es del dia siguiente, que es el que reporta la bascula (F1, 5-sep). */
function cortesDia() {
    const hoy = fechaMexico();
    const ayer = fechaMexico(new Date(Date.now() - 86400000));
    const lunes = lunesDe(hoy);   // C-30 (v0.34.0): el mismo lunes que la barra de semanas de Reportes (antes: zona del dispositivo)
    const dia = e => e.Arribo ? fechaMexico(new Date(e.Arribo)) : '';
    const diaCierre = e => e.TaraHora ? fechaMexico(new Date(e.TaraHora)) : dia(e);
    const cerrados = f => estado.embarques.filter(e => e.Etapa === 'cerrado' && f(diaCierre(e)));
    return { hoy, ayer, lunes, dia, diaCierre, cerrados };
}
function pintarHoy() {
    const { hoy, dia } = cortesDia();
    const activos = estado.embarques.filter(enPlanta);
    const pendientes = excepcionesPendientes();
    const borradores = estado.prealtas.filter(p => p.Estado === 'borrador');

    // v0.32.0: la fecha es el subtítulo y el rol el kicker (la cabecera de Proyectos); los KPI viven en Reportes.
    $('hoyTitulo').textContent = `${new Date().toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City', weekday: 'long', day: 'numeric', month: 'long' })} · ${activos.length === 1 ? '1 góndola en planta' : `${activos.length} góndolas en planta`}${borradores.length ? ` · ${borradores.length === 1 ? '1 pre-alta por firmar' : `${borradores.length} pre-altas por firmar`}` : ''}`;
    $('hoyKicker').textContent = estado.rol;

    // U-12 (v0.22.0): los mismos botones (Autorizar · Anular/Eliminar) salen en la franja y en su renglon de «Pendiente
    // revisar», que era la unica entrada de esa tarjeta sin accion; la excepcion se cuenta una sola vez (en la tarjeta).
    const botonesExcepcion = e => {
        const bs = [];
        if (PUEDE.autorizarExcepcion(estado.rol)) bs.push({ texto: 'Autorizar', accion: 'autorizar', clase: 'si', alClic: ev => autorizarExcepcion(vivo('embarques', e), ev.currentTarget), deshabilitado: motivoSinFirmas() });
        for (const b of botonCorreccion(e)) bs.push(b);   // U-52 (v0.29.0): Anular/Eliminar en rojo también en la franja; antes iban con la clase neutral «no»
        return bs;
    };
    pintarFranjaHoy(pendientes, botonesExcepcion);
    pintarFilaDia(hoy, dia);
    // Exportar lo cargado a CSV (F4): para el reporte al cliente y la bitacora, sin copiar cifras de la pantalla.
    $('btnExportar').classList.toggle('oculto', !estado.embarques.length);
    // Que la consola diga hasta donde alcanza lo que muestra: sin esta linea, «4 rechazos esta
    // semana» y «0 hace cuatro meses» se leen igual y el segundo es solo que no se cargo.
    // U-14 (v0.22.0): va en #tbAlcance, fuera de la tabla, que en el celular esta oculta: ahi nunca se veia.
    $('tbAlcance').textContent = `Se cargan los últimos ${CONFIG.ventanaDias} días (desde el ${fechaCorta(estado.ventanaDesde)}) más todo lo que sigue abierto. El historial completo vive en SharePoint.`;
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
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = `CALYTEK_Embarques_${fechaCorta(estado.ventanaDesde).replace(/\//g, '-')}_a_${fechaMexico()}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    avisar(`CSV con ${plural(filas.length, 'embarque')} descargado.`, 'bien');
}
$('btnExportar').addEventListener('click', exportarCsv);
$('pdBusca').addEventListener('input', () => { estado.padronFicha = null; pintarPadron(); });
$('baBusca').addEventListener('input', pintarCerrados);

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
$('puPrealta').addEventListener('change', () => { pintarChoferesPuerta(); pintarUnidadesPuerta(); pintarPrevioPuerta(); });
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
$('btnVolverVeredicto').addEventListener('click', () => { cerrarVeredicto(); estado.ultimaCompuerta = null; });
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
// El boton vive dentro del <summary>: sin preventDefault el clic pliega el grupo (igual que en el padron).
$('btnNuevaPrealta').addEventListener('click', ev => { ev.preventDefault(); ev.stopPropagation(); $('paGrupoBorradores').open = true; nuevaPrealta(); });
$('paCarrier').addEventListener('change', pintarUnidadesChoferesPrealta);
// U-10 (v0.23.0): «+ Alta de carrier» desde la pre-alta abre la forma del padron ENCIMA (dialog anidado: la pre-alta
// capturada se queda atras, intacta) y al guardar el carrier nuevo queda elegido aqui, con sus unidades/choferes (vacios).
// Antes habia que Cancelar (se perdia todo), ir a Padron, dar de alta y reteclear los bloques 1 y 2.
function elegirCarrierEnPrealta(c) {
    opciones($('paCarrier'), estado.carriers.filter(x => x.Activo !== false || x.id === c.id), x => x.id, x => x.Title);
    $('paCarrier').value = String(c.id);
    pintarUnidadesChoferesPrealta(); pintarEstadoPrealta();
}
$('btnAltaCarrierPrealta').addEventListener('click', () => abrirFormaPadron('carriers'));
$('btnGuardarPrealta').addEventListener('click', guardarPrealta);
$('btnCancelarPrealta').addEventListener('click', () => cerrarForma('paForma'));
// Escape cierra el <dialog> sin pasar por Cancelar: la edicion pendiente del padron se suelta igual.
for (const clave of Object.keys(FORMA_PADRON)) $(FORMA_PADRON[clave].forma).addEventListener('close', () => { if (estado.padronEdit && estado.padronEdit.clave === clave) estado.padronEdit = null; });
// U-09 (v0.22.0): Escape sobre un formulario CON algo capturado pregunta antes de tirarlo — el <dialog> nativo cerraba sin
// pasar por Cancelar y la pre-alta de 14 campos se perdia. Vacio (o solo con los valores por omision) cierra directo; el
// boton Cancelar sigue cerrando sin preguntar. Un select cuenta solo si no esta en su primera opcion.
const hayCaptura = d => [...d.querySelectorAll('input:not([type=hidden]):not([type=checkbox]), textarea, select')].some(c => c.tagName === 'SELECT' ? c.selectedIndex > 0 : String(c.value).trim() !== '');
for (const id of ['paForma', 'pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer']) $(id).addEventListener('cancel', async ev => {
    if (!hayCaptura($(id)) || huellaForma($(id)) === $(id).dataset.huella) return;   // U-34: sin cambios desde que abrió → cierra directo
    ev.preventDefault();
    const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: 'Este formulario tiene datos sin guardar. Si lo cierras, se pierden.' });
    if (ok) cerrarForma(id);
});
$('btnFirmar').addEventListener('click', firmarPrealta);
$('btnCerrarPrealta').addEventListener('click', cerrarPrealta);
$('btnEmitirCertificado').addEventListener('click', () => emitirCertificado());   // v0.35.0
$('btnSustituirCertificado').addEventListener('click', sustituirCertificado);
$('btnCancelarCertificado').addEventListener('click', cancelarCertificado);
$('btnVerCertificado').addEventListener('click', () => verCertificado());
$('ctCerrar').addEventListener('click', () => { $('dlgCertificado').close(); estado.certificadoAbierto = null; });
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
