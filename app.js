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
import { compuerta, siguienteFolio, avisoNeto, placaNormal, fechaMexico, horaMexico, slug, rolDe, PUEDE, lista, diasPara, evaluarVigencia, accionCorreccion, prealtaSinMovimiento, fechaCorta, aIsoDia, autoformatoFecha, plural, limpiar, paraPatch, tipoDeArchivo, lunesDe, sumarDias, esLoteDeLaApp, residuoDe, sufijoVerificacion, datosCertificado, urlVerificacion, toneladas, siguientePaso, yaCapturado, CORRIENTES, etiquetaCorriente, palabraCompuerta, subpasoDeRegla, clienteDe, huellaPrealta, firmaAmparaPrealta, basesRecientes, clientesPrealta, fechaDePestana, mesesPrealtas } from './reglas.js';

const VERSION = '0.66.1';   // la misma cadena va en package.json y en sw.js (CACHE); test/version.test.js lo exige
const $ = id => document.getElementById(id);
const L = CONFIG.listas;

const estado = {
    cuenta: null, token: null, cliente: null, siteId: null, rol: 'lectura',
    carriers: [], unidades: [], choferes: [], prealtas: [], embarques: [], vigencias: [], roles: [],
    firmas: [], firmasError: null,   // S-01 / S-07 (v0.25.0): renglones de PLANTA_Firmas; si la lista no se pudo leer, el motivo (y la app no firma ni autoriza)
    certificados: [], certificadosError: null,   // v0.35.0: renglones de PLANTA_Certificados; si la lista no existe todavia, el motivo (y no se emite)
    certificadoAbierto: null,   // el certificado pintado en dlgCertificado
    certificadoEmbarque: null,  // v0.39.0: la gondola (embarque cerrado) cuyo certificado se esta viendo o emitiendo
    ultimaCompuerta: null,   // {resultado, hallazgos, campos}
    pesando: null,           // {embarque, fase: 'bruto'|'tara'}
    fotoBytes: null,
    fotoUrl: null,           // blob URL de la vista previa; se revoca al reemplazar la foto o cancelar (C-21)
    prealtaAbierta: null,
    prealtaVista: null,      // C-60 (v0.56.0): { estado, huella } de la pre-alta tal como la pinto el detalle
    prealtaEdit: null,       // borrador de pre-alta en edicion (asistente #paAsis)
    paAsis: null,            // v0.54.0: el recorrido del asistente de pre-alta (paso, max, revisar...); null = cerrado
    padronEdit: null,        // {clave, x} del renglon del padron en edicion, o null
    focoAntesVeredicto: null, // elemento con el foco antes de abrir el veredicto (vuelve ahi al cerrarlo)
    padronFicha: null,       // 'tipo:id' de la ficha del padron desplegada
    padronCarrier: null,     // id del carrier que filtra unidades y choferes en el padron
    ultimoCarrierPadron: null,   // U-46: el ultimo carrier dado de alta o usado en un alta de unidad/chofer (solo esta sesion)
    pestana: 'hoy',
    vistaPrealtas: 'borrador', // rediseño de Pre-altas tanda 1 (v0.53.0): la pestaña elegida (borrador · firmada · cerrada)
    vistaGondolas: 'planta', // tanda 3 (v0.48.0): la pestaña elegida dentro de Góndolas (planta · hoy · rechazos · historial)
    subpasoPuerta: 1,        // tanda 4 (v0.49.0): la pantalla del paso 1 del asistente (1 programa · 2 vehículo y chofer · 3 carga)
    archivos: null,          // v0.33.0: { biblioteca, ramas: Map ruta -> hijos } de la seccion Archivos; null = se relee al pintar
    cargadoEl: 0,
    ventanaDesde: '',     // ISO: inicio de la ventana de carga (cubeta 3)
    fueraDeVentana: new Map()   // C-68 (v0.63.0): PreAltaId -> embarques de los programas que empezaron antes de la ventana
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
    // ultima abierta y, en el DOM, la ultima de las abiertas (las del padron van al final del body).
    // U-71 (v0.40.0): .con-avisos suma los modales que no son forma pero tienen su zona de avisos (el del certificado).
    const abiertas = document.querySelectorAll('dialog.dlg-forma[open], dialog.con-avisos[open]');
    const dlg = abiertas[abiertas.length - 1];
    if (dlg) { const z = dlg.querySelector('.dlg-avisos'); z.textContent = ''; z.appendChild(d.cloneNode(true)); dlg.scrollTo({ top: 0, behavior: 'smooth' }); return; }
    // Tanda 2 de Pre-altas (v0.54.0): el asistente dejó de ser <dialog>; su zona de avisos hace el papel de la del pop-up.
    if (estado.pestana === 'prealtas' && asistentePrealtaAbierto()) { const z = $('paAvisos'); z.textContent = ''; z.appendChild(d.cloneNode(true)); z.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); return; }
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
    b.textContent = texto || 'Entrar con mi cuenta de MINSA';
    $('textoEntrar').textContent = texto || 'CALYTEK · Planta';
    $('textoEntrar').classList.toggle('estado', !!texto);   // v0.45.3: como proyectos (U-05), un ESTADO se lee; la marca en reposo, chica
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
    // Pantalla inicial por rol: quien captura abre en Góndolas (decisión 12 del rediseño, v0.47.0; antes en la puerta),
    // quien firma en las pre-altas, gerencia y lectura en «Hoy» (la consola). Igual en celular y en computadora.
    irA(estado.rol === 'trazabilidad' ? 'bascula' : estado.rol === 'validador' ? 'prealtas' : 'hoy');
}

/** Un embarque esta "en planta" si paso la compuerta (o le autorizaron la excepcion) y no ha cerrado: lo que se PESA.
 *  Lo usan los botones de pesar, el Historial (lo que ya no esta en planta) y el KPI de Reportes. */
function enPlanta(e) {
    return (e.Etapa === 'compuerta' && (e.Compuerta === 'pasa' || excepcionAutorizada(e)))
        || e.Etapa === 'bruto';
}
const esperaAutorizacion = e => e.Etapa === 'compuerta' && e.Compuerta === 'excepcion-comercial' && !excepcionAutorizada(e);
function excepcionesPendientes() { return estado.embarques.filter(esperaAutorizacion); }
/** Tanda 3 (decision 10): la lista «En planta» de Góndolas = lo que se pesa + lo que espera la autorizacion de gerencia
 *  (en ambar), por hora de llegada. Es LA definicion del numero del rail, de la pestaña y del subtitulo de Hoy. */
const enListaPlanta = () => estado.embarques.filter(e => enPlanta(e) || esperaAutorizacion(e)).sort((a, b) => String(a.Arribo || '').localeCompare(String(b.Arribo || '')));

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
/**
 * v0.35.0: PLANTA_Certificados puede no existir todavia (provisionar.html la crea): la app sigue, sin emitir, y lo dice a gerencia.
 * C-45 (v0.44.0): solo los certificados de las gondolas CARGADAS. Nace uno por gondola y nunca se borra, asi que la lista
 * entera crece sin tope; los ids de PLANTA_Embarques son crecientes, asi que «EmbarqueId ge el menor id cargado» (columna
 * indexada, filtro numerico) trae todos los de la ventana y los abiertos. Sin gondolas no hay nada que certificar.
 */
async function cargarCertificados(c, s, avisar, embarques) {
    if (!embarques.length) { estado.certificadosError = null; return []; }
    const desde = Math.min(...embarques.map(e => Number(e.id)));
    try { const x = await c.renglones(s, L.certificados, `fields/EmbarqueId ge ${desde}`, avisar); estado.certificadosError = null; return x; }
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
/**
 * S-27 (v0.56.0): la firma de pre-alta vale solo si su huella (Motivo) sigue cuadrando con el renglón: tras firmada, un
 * cambio de carrier, corriente, unidades, choferes o generador —por la app o por Graph— la deja sin firma y la puerta no la ve.
 */
const firmasPrealta = p => estado.firmas.filter(f => f.Tipo === 'prealta' && Number(f.ObjetoId) === Number(p.id)
    && ROL_FIRMA.prealta.includes(rolDe(f.Firmante, estado.roles)) && mismaCuenta(f.Firmante, p.FirmadaPor));
function prealtaFirmada(p) { return p.Estado === 'firmada' && firmasPrealta(p).some(f => firmaAmparaPrealta(f.Motivo, p, f.Created)); }   // S-29: Created lo pone SharePoint
/** Firmada con sello y firma, pero el renglón cambió después: se dice así, no «sin firma». */
const prealtaCambioTrasFirma = p => p.Estado === 'firmada' && !prealtaFirmada(p) && firmasPrealta(p).length > 0;
function excepcionAutorizada(e) { return !!e.ExcepcionAutorizo && !!firmaDe('excepcion', e.id, e.ExcepcionAutorizo); }
/** C-55 (v0.52.0): la compuerta de un embarque en palabras, con la firma y no el sello decidiendo la excepción. */
const palabraCompuertaDe = e => palabraCompuerta(e.Compuerta, excepcionAutorizada(e));
/** Sellos sin firma que valga (de antes del corte, por fuera de la app, o firmada por quien no tiene el rol): la compuerta no los acepta. */
function sellosSinFirma() {
    return { prealtas: estado.prealtas.filter(p => p.Estado === 'firmada' && !prealtaFirmada(p)),
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
    const b = enListaPlanta().length;
    const p = estado.prealtas.filter(x => x.Estado === 'borrador').length;
    $('nGondolas').textContent = String(b); $('nGondolas').hidden = b === 0;
    $('nPrealtas').textContent = String(p); $('nPrealtas').hidden = p === 0 || !PUEDE.firmarPrealta(estado.rol);
    // El carril Puerta -> Bascula -> Ticket de «Hoy» (D2, 8-sep) salio en la tanda 2 del rediseño (decision 6).
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

/**
 * C-68 (v0.63.0): las góndolas de un programa que empezó ANTES de la ventana no están todas en estado.embarques (90 días):
 * una cerrada vieja pintaba 0/20 y una firmada larga «sin un solo arribo». Sus embarques se leen aparte por PreAltaId
 * (columna indexada), de 15 en 15, y NO entran a estado.embarques: Báscula, Historial y Reportes siguen mirando la ventana.
 * «Empezó» = Created de la pre-alta (lo pone SharePoint; ningún embarque es anterior), o su firma si no viene. Las cerradas
 * ya no reciben góndolas y se leen una vez por sesión; las firmadas, en cada carga.
 */
async function cargarProgramasViejos(c, s, avisar) {
    const desde = Date.parse(estado.ventanaDesde);
    const viejo = p => (p.Estado === 'firmada' || p.Estado === 'cerrada') && !(Date.parse(p.Created || p.FirmadaEl || '') >= desde);
    const antes = estado.fueraDeVentana, ahora = new Map(), faltan = [];
    for (const p of estado.prealtas.filter(viejo)) {
        if (p.Estado === 'cerrada' && antes.has(p.id)) ahora.set(p.id, antes.get(p.id));
        else { ahora.set(p.id, []); faltan.push(p.id); }
    }
    for (let i = 0; i < faltan.length; i += 15) {
        const ids = faltan.slice(i, i + 15);
        for (const e of await c.renglones(s, L.embarques, ids.map(id => `fields/PreAltaId eq ${id}`).join(' or '), avisar)) ahora.get(Number(e.PreAltaId))?.push(e);
    }
    estado.fueraDeVentana = ahora;
}
/** Los embarques que cuentan para los programas: la ventana más los de programas viejos (C-68), sin repetir. */
function embarquesDeProgramas() {
    if (!estado.fueraDeVentana.size) return estado.embarques;
    const porId = new Map(estado.embarques.map(e => [e.id, e]));
    for (const filas of estado.fueraDeVentana.values()) for (const e of filas) if (!porId.has(e.id)) porId.set(e.id, e);
    return [...porId.values()];
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
    // C-45 (v0.44.0): los certificados se acotan con los ids de los embarques, asi que van encadenados a ellos (y el resto en paralelo).
    const embarquesYCertificados = cargarEmbarques(c, s, av).then(async emb => [emb, await cargarCertificados(c, s, av, emb)]);
    [estado.carriers, estado.unidades, estado.choferes, estado.prealtas, [estado.embarques, estado.certificados], estado.vigencias, estado.roles, estado.firmas] =
        await Promise.all([
            c.renglones(s, L.carriers, null, av), c.renglones(s, L.unidades, null, av), c.renglones(s, L.choferes, null, av),
            c.renglones(s, L.prealtas, null, av), embarquesYCertificados, c.renglones(s, L.vigencias, null, av),
            c.renglones(s, L.roles, null, av), cargarFirmas(c, s, av)
        ]);
    await cargarProgramasViejos(c, s, av);
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
    estado.certificadoEmbarque = fresco(estado.embarques, estado.certificadoEmbarque);   // C-38 (v0.40.0)
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
    if (estado.pestana === 'prealtas' && asistentePrealtaAbierto()) return true;   // v0.54.0: el asistente de pre-alta a la vista
    if (['paDetalle', 'paRecientes', 'pdFormaCarrier', 'pdFormaUnidad', 'pdFormaChofer'].some(id => $(id).open)) return true;
    if (estado.pestana === 'puerta' && ['puManifiesto', 'puPlaca', 'puPlacaPlana', 'puChoferNombre', 'puMotivo', 'puPrealta'].some(id => $(id).value.trim())) return true;   // U-40: el programa elegido también es captura
    return false;
}

let recargando = false;
async function recargar(silencioso = false) {
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
const botonesRail = () => [...$('pestanas').querySelectorAll('button'), ...$('pestanasExtra').querySelectorAll('button')];
/** Repinta la pestana abierta SIN tocar avisos, veredicto ni scroll (refresco silencioso y cambios de pre-alta). */
function repintar() { pintarInsignias(); PINTORES[estado.pestana](); }
/**
 * C-27 / U-41 (v0.28.0): salir de la bascula con un pesaje abierto pasa por la misma compuerta que «Cancelar» (U-24): con kg o
 * foto pregunta, y al salir suelta estado.pesando, la foto y su blob URL. Antes irA() ocultaba la tarjeta sin preguntar y el
 * basculista volvia a fotografiar el indicador; el JPEG comprimido quedaba vivo hasta el siguiente pesaje. Devuelve si se salio.
 */
const pesajeConAlgo = () => !$('baPesar').classList.contains('oculto') && !!($('baKg').value.trim() || estado.fotoBytes);
function descartarPesaje() { estado.pesando = null; estado.fotoBytes = null; soltarFotoPrevia(); cerrarAsistente(); }
async function soltarPesaje() {
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
function irDesdePestana(p) {
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
function irA(p) {
    if (p !== 'bascula') salirDelAsistente();
    ocultarListoPrealta();   // P10: la confirmación no guarda nada; salir o volver por el rail la suelta
    estado.pestana = p;
    for (const b of botonesRail()) { if (b.dataset.p === (RAIL_DE[p] || p)) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); }   // U-57 (v0.29.0): <nav> con aria-current, como .mn-rail de la piel; el role=tablist prometía flechas y tabpanel que no había
    for (const s of SECCIONES) $('p-' + s).classList.toggle('oculto', s !== p);
    pintarRotulo(RAIL_DE[p] || p);
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
function irSubpaso(s, enfocar = false) {
    s = SUBPASOS[s] ? Number(s) : 1;
    const cambio = s !== estado.subpasoPuerta;
    estado.subpasoPuerta = s;
    for (const n of [1, 2, 3]) $('puBloque' + n).hidden = n !== s;
    for (const b of $('puSubpasos').querySelectorAll('button')) { if (Number(b.dataset.s) === s) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current'); }
    $('puPaso1').dataset.tercio = String(s);
    $('puPasoK').textContent = `Paso 1 de 5 · ${SUBPASOS[s]}`;
    $('btnPuAtras').hidden = s === 1;
    if (s > 1) $('btnPuAtras').textContent = `‹ ${SUBPASOS[s - 1]}`;
    $('btnPuSiguiente').hidden = s === 3;
    if (s < 3) $('btnPuSiguiente').textContent = SIGUIENTE_SUBPASO[s];
    $('btnCompuerta').hidden = s !== 3;
    if (cambio && estado.pestana === 'puerta') window.scrollTo({ top: 0 });
    if (enfocar) { const q = $('puBloque' + s).querySelector('.pregunta'); if (q) { q.tabIndex = -1; q.focus({ preventScroll: true }); } }
}
/** Una empieza nueva: con el programa ya elegido (las góndolas del mismo programa llegan en fila) arranca en el vehículo. */
const subpasoInicial = () => ($('puPrealta').value ? 2 : 1);

/**
 * M2: un renglón tocable que actúa como radio. `valor` va en data-v; el que está elegido lleva .sel y aria-pressed, como los
 * chips de unidad de siempre (U-33). `dato` es la columna derecha: texto o un nodo (la etiqueta de vigencia).
 */
function renglonOpcion({ valor, sel, titulo, mono = false, detalle, dato, clase = '' }) {
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
function marcarOpcion(cont, valor) {
    for (const x of cont.querySelectorAll('.o')) { const on = valor !== '' && x.dataset.v === String(valor); x.classList.toggle('sel', on); x.setAttribute('aria-pressed', String(on)); }
}
/** Elegir en un renglón es elegir en el select oculto: el listener de `change` de siempre repinta lo que depende de él. */
function elegirEnSelect(id, v) { $(id).value = String(v); $(id).dispatchEvent(new Event('change', { bubbles: true })); }

function pintarProgramasPuerta(firmadas) {
    const cont = $('puProgramas'); cont.textContent = '';
    if (!firmadas.length) { cont.appendChild(el('p', 'pista', 'No hay programas firmados todavía.')); return; }
    for (const p of firmadas) {
        const { rec, esp } = gondolasDe(p);
        const b = renglonOpcion({ valor: p.id, sel: String(p.id) === $('puPrealta').value, titulo: p.Title,
            detalle: [nombreDe(estado.carriers, p.CarrierId), etiquetaCorriente(p.Corriente)].filter(Boolean).join(' · '),
            dato: esp ? `${rec} de ${esp}` : plural(rec, 'recibida') });
        b.addEventListener('click', () => elegirEnSelect('puPrealta', p.id));
        cont.appendChild(b);
    }
}
function pintarCorrientesPuerta() {
    const cont = $('puCorrientes'); cont.textContent = '';
    for (const [valor, titulo] of CORRIENTES) {   // C-55: del catálogo, no del select oculto
        const b = renglonOpcion({ valor, sel: valor === $('puCorriente').value, titulo });
        b.addEventListener('click', () => elegirEnSelect('puCorriente', valor));
        cont.appendChild(b);
    }
}
/** Arriba del asistente, la góndola que se está capturando: placas o «Góndola nueva», y el programa y el chofer debajo. */
function pintarQuienPuerta() {
    const placa = placaNormal($('puPlaca').value), plana = placaNormal($('puPlacaPlana').value);
    $('puQuien').textContent = placa ? (plana ? `${placa} · ${plana}` : placa) : 'Góndola nueva';
    $('puQuien').classList.toggle('mono', !!placa);
    const pre = porId(estado.prealtas, $('puPrealta').value);
    const chofer = $('puChoferNombre').value.trim() || (porId(estado.choferes, $('puChofer').value) || {}).Title;
    $('puQuienSub').textContent = pre ? [pre.Title, pre.CarrierId ? nombreDe(estado.carriers, pre.CarrierId) : null, chofer].filter(Boolean).join(' · ') : 'sin programa todavía';
}
/** Dónde vive cada campo que la compuerta necesita, y qué recibe el foco cuando falta (el select oculto no puede). */
const PANTALLA_DE = { puPrealta: 1, puPlaca: 2, puPlacaPlana: 2, puChofer: 2, puChoferNombre: 2, puManifiesto: 3, puCorriente: 3 };
function enfocarCampoPuerta(id) {
    irSubpaso(PANTALLA_DE[id] || 1);
    if (id === 'puPlaca' || id === 'puPlacaPlana') $('puTeclear').open = true;
    const destino = id === 'puPrealta' ? $('puProgramas').querySelector('.o') : id === 'puCorriente' ? $('puCorrientes').querySelector('.o') : $(id);
    if (destino) destino.focus();
}
// C-56 (v0.52.0): subpasoDeRegla vive en reglas.js, con tabla explícita y su prueba contra cada nombre que emite compuerta().
// I6 (7-sep): las unidades que la pre-alta autorizo (o, si no marco ninguna, todas las activas del carrier).
// Tocar una llena placa tractor y plana; la seleccion se marca comparando con lo que hay en el campo, asi que
// teclear otra placa la desmarca sola. La compuerta sigue evaluando la placa del campo, no la seleccion.
function pintarUnidadesPuerta() {
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

function correrCompuerta() {
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
function pintarPrevioPuerta() {
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
    pintarQuienPuerta();
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

// ================================================================ GONDOLAS (rediseño tanda 3, v0.48.0)

/**
 * Tanda 3 (decisiones 5, 6 y 10): una sección con cuatro pestañas —En planta · Cerradas hoy · Rechazos · Historial— que
 * funde la lista de báscula, la Fila del día (tanda 2) y Cerrados. Las cuatro se pintan en cada repintado y solo se ve la
 * elegida (estado.vistaGondolas), así los conteos de las pestañas nunca se quedan viejos.
 */
const VISTAS_GONDOLAS = { planta: 'gpPlanta', hoy: 'gpHoy', rechazos: 'gpRechazos', historial: 'gpHistorial' };
function elegirVistaGondolas(v) {
    estado.vistaGondolas = VISTAS_GONDOLAS[v] ? v : 'planta';
    for (const b of document.querySelectorAll('#gTabs button, #gKpis button')) b.setAttribute('aria-pressed', String(b.dataset.g === estado.vistaGondolas));
    for (const [k, id] of Object.entries(VISTAS_GONDOLAS)) $(id).hidden = k !== estado.vistaGondolas;
}

/**
 * Una tabla de Góndolas. Un solo DOM para los dos anchos: tabla en escritorio y tarjeta en celular, donde cada celda cae en
 * su lugar por su letra `m` (a · b arriba, c · d en medio, e a lo ancho, x solo en escritorio; estilo.css .gtabla). La
 * última celda, sin encabezado, lleva los botones. `columnas`: [{ t, m, v: e => texto | Node, num, clase }].
 */
function tablaGondolas(caja, columnas, filas, botonesDe, vacio, claseDe = () => '') {
    caja.textContent = '';
    if (!filas.length) { caja.appendChild(el('p', 'vacio', vacio)); return; }
    const t = el('table', 'gtabla'), thead = el('thead'), cab = el('tr'), tb = el('tbody');
    for (const c of columnas) cab.appendChild(el('th', c.num ? 'num' : '', c.t));
    cab.appendChild(el('th'));
    thead.appendChild(cab); t.appendChild(thead);
    for (const e of filas) {
        const tr = el('tr', `gfila ${claseDe(e)}`.trim());
        for (const c of columnas) {
            const td = el('td', `m-${c.m}${c.num ? ' num' : ''}${c.clase ? ' ' + c.clase : ''}`);
            const v = c.v(e);
            if (v instanceof Node) td.appendChild(v); else td.textContent = v === null || v === undefined || v === '' ? '—' : v;
            tr.appendChild(td);
        }
        const ac = el('td', 'm-f acciones-g');
        for (const b of botonesDe(e)) ac.appendChild(botonAccion(b));
        tr.appendChild(ac);
        tb.appendChild(tr);
    }
    t.appendChild(tb); caja.appendChild(t);
}
const folioG = e => el('span', e.Title ? 'folio' : 'folio sin', e.Title || 'sin folio');
/** El programa y, debajo, el manifiesto: es lo que se busca con el papel en la mano. */
function programaG(e) {
    const s = el('span', 'prog', nombreDe(estado.prealtas, e.PreAltaId));
    if (e.Manifiesto) s.appendChild(el('small', 'mono', e.Manifiesto));
    return s;
}
const kgG = n => `${Number(n).toLocaleString('es-MX')} kg`;
/** Lo último que se capturó de una góndola en planta: el peso bruto, o la hora a la que pasó (o llegó, si espera). */
function ultimoDatoG(e) {
    if (e.Etapa === 'bruto') return `bruto ${kgG(e.BrutoKg)}`;
    return `${esperaAutorizacion(e) ? 'arribo' : 'pasó'} ${horaMexico(e.Arribo, 'hora')}`;
}
/** Ticket (si hay folio), el certificado de la góndola cerrada y Anular/Eliminar. `lista` es la que recorre el ticket con ‹ ›. */
function botonesCerrada(e, lista) {
    const bs = [];
    if (e.Title) bs.push({ texto: 'Ticket', accion: 'ticket', alClic: () => abrirTicketPop(lista, lista.indexOf(e)) });
    // v0.39.0: el certificado es de ESTA góndola y se emite desde su renglón cerrado (decisión 9: solo gerencia emite).
    // U-70 (v0.40.0): una anulada con certificado lo sigue mostrando, para que gerencia llegue a cancelarlo.
    const certs = certificadosDe(e), vig = certs.find(x => x.Estado === 'vigente');
    if (e.Etapa === 'cerrado' && !vig && PUEDE.emitirCertificado(estado.rol)) bs.push({ texto: 'Emitir certificado', accion: 'certificado', clase: 'si', alClic: () => abrirCertificadoDeEmbarque(vivo('embarques', e)) });
    else if (certs.length) bs.push({ texto: 'Certificado', accion: 'certificado', alClic: () => abrirCertificadoDeEmbarque(vivo('embarques', e)) });
    return bs.concat(botonCorreccion(e));
}
function estadoCertificadoG(e) {
    const vig = certificadoVigente(e);
    return vig ? el('span', 'paso e-ok', `${vig.Title} vigente`) : el('span', 'paso e-warn', 'sin emitir');
}

/** La sección entera: las listas y, cerradas, las tarjetas de pesaje y de ticket (como al entrar). */
function pintarGondolas() {
    cerrarAsistente();
    pintarListasGondolas();
}

// ---------------------------------------------------------------- tanda 5: el asistente de la báscula (M5–M8)
/**
 * Muestra el asistente sobre la góndola `e` en una de sus tres pantallas (baPesar · baPausa · baTicketCaja) y esconde la
 * lista. `paso` pinta la barra de cinco: 3 bruto, 4 tara, 5 ticket; 3.5 es la pausa de la descarga (tres hechos, ninguno
 * en curso: la góndola no está en la báscula).
 */
function mostrarAsistente(e, pantalla, paso) {
    estado.asisEmbarque = e;
    for (const id of ['baPesar', 'baPausa', 'baTicketCaja']) $(id).classList.toggle('oculto', id !== pantalla);
    $('baAsis').classList.remove('oculto');
    $('p-bascula').classList.add('asistiendo');
    [...$('baPasos').children].forEach((d, i) => {
        d.className = i + 1 < paso ? 'hecho' : i + 1 === paso ? 'ahora' : '';
        if (i + 1 === paso) d.dataset.tercio = '3'; else delete d.dataset.tercio;
    });
    $('baPasoK').textContent = pantalla === 'baPausa' ? 'Paso 3 de 5 · A descargar' : `Paso ${paso} de 5 · ${['', '', '', 'Bruto', 'Tara', 'Ticket'][paso]}`;
    $('baQuien').textContent = e.Title ? `${e.Title} · ${e.PlacaTractor || ''}` : [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' · ');
    $('baMiga').textContent = e.Title || e.PlacaTractor || 'Góndola';
    // U-88 (v0.51.0): cerrada, el dato que se le dice al chofer es el neto, no el bruto de la etapa anterior.
    const peso = e.NetoKg ? `neto ${kgG(e.NetoKg)}` : e.BrutoKg ? `bruto ${kgG(e.BrutoKg)}` : null;
    $('baQuienSub').textContent = [nombreDe(estado.prealtas, e.PreAltaId), nombreDe(estado.carriers, e.CarrierId), peso].filter(x => x && x !== '—').join(' · ');
    pintarCapturado(e);
    cerrarHojaCapturado();
    window.scrollTo({ top: 0 });
}
function cerrarAsistente() {
    for (const id of ['baAsis', 'baPesar', 'baPausa', 'baTicketCaja']) $(id).classList.add('oculto');
    $('p-bascula').classList.remove('asistiendo');
    cerrarHojaCapturado();
    estado.asisEmbarque = null;
}
/** M5: tras el bruto, la pausa de la descarga — el folio ya nació y la góndola se retoma desde la lista. */
function mostrarPausa(e) {
    const dl = $('baPausaDatos'); dl.textContent = '';
    for (const [k, v] of [['Folio', e.Title], ['Bruto', `${kgG(e.BrutoKg)} · ${horaMexico(e.BrutoHora, 'hora')}`], ['Sigue', 'la tara, con la góndola vacía']]) {
        dl.appendChild(el('dt', '', k)); dl.appendChild(el('dd', k === 'Sigue' ? '' : 'mono', v));
    }
    mostrarAsistente(e, 'baPausa', 3.5);
}
/** M6 / M7: todo lo que la góndola ya trae, por paso. En escritorio es la columna derecha; en celular, la hoja. */
function pintarCapturado(e) {
    const dl = $('baCapturadoDl'); dl.textContent = '';
    const sec = (t, h) => { const d = el('div', 'sec', t); if (h) d.appendChild(el('em', '', h)); dl.appendChild(d); };
    // nombreDe() da «—» cuando no hay id: ese renglón no se pinta.
    const par = (k, v, clase) => { if (!v || v === '—') return; dl.appendChild(el('dt', '', k)); dl.appendChild(el('dd', clase || '', v)); };
    const pre = porId(estado.prealtas, e.PreAltaId);
    const corr = e.CorrienteDeclarada || (pre && pre.Corriente);
    sec('Programa y documentos', horaMexico(e.Arribo, 'hora'));
    par('Programa', nombreDe(estado.prealtas, e.PreAltaId)); par('Carrier', nombreDe(estado.carriers, e.CarrierId));
    par('Placas', [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' · '), 'mono');
    par('Chofer', e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId));
    par('Manifiesto', e.Manifiesto, 'mono'); par('Corriente', etiquetaCorriente(corr));
    par('Capturó', quien(e.CapturadoPor));
    const avisos = reglasDe(e, 'aviso');   // C-54
    sec('Veredicto');
    // C-55: un sello sin renglón en PLANTA_Firmas ya no se lee como autorización (lo mismo que dicen la lista y Hoy).
    const pc = palabraCompuertaDe(e), autorizo = excepcionAutorizada(e);
    par('Resultado', e.Compuerta !== 'excepcion-comercial' ? pc.palabra : autorizo ? `Espera · autorizó ${quien(e.ExcepcionAutorizo)}` : `Espera · ${selloSinFirma(e)}falta autorización`, 'e-' + pc.tono);
    par('Avisos', avisos, 'e-warn');
    if (e.BrutoKg) { sec('Peso bruto', horaMexico(e.BrutoHora, 'hora')); par('Folio', e.Title, 'mono'); par('Bruto', kgG(e.BrutoKg), 'mono'); }
    if (e.TaraKg) { sec('Tara', horaMexico(e.TaraHora, 'hora')); par('Tara', kgG(e.TaraKg), 'mono'); par('Neto', kgG(e.NetoKg), 'mono'); par('Ticket de báscula', e.TicketBascula, 'mono'); }
}
// U-81 (v0.51.0): el velo es un elemento (#baVelo) que recibe el toque y cierra la hoja; antes era una box-shadow y el toque
// caía en el teclado o en Guardar de atrás.
function abrirHojaCapturado() { $('baCapturado').classList.add('abierta'); $('baVelo').hidden = false; $('btnLoCapturado').setAttribute('aria-expanded', 'true'); $('btnCerrarCapturado').focus(); }
function cerrarHojaCapturado() { $('baCapturado').classList.remove('abierta'); $('baVelo').hidden = true; $('btnLoCapturado').setAttribute('aria-expanded', 'false'); }
/** Solo las cuatro listas y sus conteos: tras emitir un certificado no se toca el pesaje abierto. */
function pintarListasGondolas() {
    const { hoy, dia, diaCierre } = cortesDia();
    const planta = enListaPlanta();
    const porCierre = (a, b) => String(b.TaraHora || '').localeCompare(String(a.TaraHora || ''));
    const cerradasHoy = estado.embarques.filter(e => e.Etapa === 'cerrado' && diaCierre(e) === hoy).sort(porCierre);
    const rechazosHoy = estado.embarques.filter(e => e.Etapa === 'rechazado' && dia(e) === hoy).sort((a, b) => b.id - a.id);
    const cuenta = (id, n, alerta) => { $(id).textContent = n ? String(n) : ''; $(id).classList.toggle('alerta', !!alerta && n > 0); };
    const esperan = planta.filter(esperaAutorizacion).length;
    cuenta('gnPlanta', planta.length, esperan > 0); cuenta('gnHoy', cerradasHoy.length); cuenta('gnRechazos', rechazosHoy.length, true);
    // v0.66.0: la banda cuenta lo mismo que las pestañas; en cero dice 0, no se vacía.
    for (const [id, n, alerta] of [['gkPlanta', planta.length, esperan > 0], ['gkHoy', cerradasHoy.length, false], ['gkRechazos', rechazosHoy.length, true]]) { $(id).textContent = String(n); $(id).parentElement.classList.toggle('alerta', alerta && n > 0); }
    elegirVistaGondolas(estado.vistaGondolas);
    $('btnNuevaGondola').classList.toggle('oculto', !PUEDE.puerta(estado.rol));   // tanda 2: la llegada se abre desde aquí; quien no captura no la ve

    // En planta: el siguiente paso de cada góndola en texto (reglas.siguientePaso). La que espera a gerencia sale en ámbar,
    // sin «Pesar» (decisión 10): se autoriza en Hoy › Pendiente revisar.
    const captura = PUEDE.puerta(estado.rol);
    tablaGondolas($('baLista'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Carrier', m: 'x', v: e => nombreDe(estado.carriers, e.CarrierId) },
        { t: 'Siguiente paso', m: 'e', v: e => { const s = siguientePaso(e, excepcionAutorizada(e)); return el('span', 'paso e-' + s.tono, s.texto); } },
        { t: 'Último dato', m: 'd', num: true, v: ultimoDatoG, clase: 'mono' }
    ], planta, e => {
        const bs = [];
        if (captura && enPlanta(e)) { const bruto = e.Etapa === 'compuerta'; bs.push({ texto: bruto ? 'Pesar bruto' : 'Pesar tara', accion: 'pesar', clase: 'si', alClic: () => abrirPesaje(vivo('embarques', e), bruto ? 'bruto' : 'tara') }); }
        return bs.concat(botonCorreccion(e));
    }, 'Nada en planta.', e => (esperaAutorizacion(e) ? 'espera' : ''));
    $('gPiePlanta').textContent = planta.length ? `${plural(planta.length, 'góndola en planta', 'góndolas en planta')}${esperan ? ` · ${esperan === 1 ? 'una espera' : `${esperan} esperan`} la autorización de gerencia` : ''}.` : '';

    // Cerradas hoy: el día de cierre es el de la tara (F1), y aquí vive el certificado (decisión 9).
    tablaGondolas($('gHoy'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Tara', m: 'x', v: e => horaMexico(e.TaraHora, 'hora'), clase: 'mono mudo' },
        { t: 'Neto', m: 'd', num: true, v: e => kgG(e.NetoKg), clase: 'mono' },
        { t: 'Ticket báscula', m: 'x', v: e => e.TicketBascula, clase: 'mono' },
        { t: 'Certificado', m: 'e', v: estadoCertificadoG }
    ], cerradasHoy, e => botonesCerrada(e, cerradasHoy), 'Ninguna góndola ha cerrado hoy.');
    const kgHoy = cerradasHoy.reduce((a, e) => a + (Number(e.NetoKg) || 0), 0);
    $('gPieHoy').textContent = cerradasHoy.length ? `${plural(cerradasHoy.length, 'cerrada', 'cerradas')} hoy · ${kgG(kgHoy)}${PUEDE.emitirCertificado(estado.rol) ? '' : ' · el certificado lo emite gerencia'}.` : '';

    // Rechazos de hoy: la causa sale de CompuertaDetalle (las reglas legales que fallaron, C-54) y, debajo, qué le pasó a
    // cada una (U-89): «Placa» sola no dice si no estaba en el padrón o la amparaba otro carrier.
    const causa = e => {
        const hs = hallazgosDe(e, 'legal'), s = el('span', 'causa');
        s.appendChild(el('span', 'paso e-bad', hs.map(h => h.regla).join(', ') || 'no entró'));
        const det = hs.map(h => h.detalle).filter(Boolean).join(' · ');
        if (det) s.appendChild(el('small', '', det));
        return s;
    };
    tablaGondolas($('gRechazos'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Carrier', m: 'c', v: e => nombreDe(estado.carriers, e.CarrierId) },
        { t: 'Hora', m: 'd', num: true, v: e => horaMexico(e.Arribo, 'hora'), clase: 'mono' },
        { t: 'Causa', m: 'e', v: causa }
    ], rechazosHoy, e => botonesCerrada(e, rechazosHoy.filter(x => x.Title)), 'Ningún rechazo hoy.');

    pintarHistorial();
    // F4: el CSV de todo lo cargado, para el reporte al cliente y la bitácora, sin copiar cifras de la pantalla.
    $('btnExportar').classList.toggle('oculto', !estado.embarques.length);
    // U-14: que la pantalla diga hasta dónde alcanza lo que muestra; «0 hace cuatro meses» es solo que no se cargó.
    $('tbAlcance').textContent = `Se cargan los últimos ${CONFIG.ventanaDias} días (desde el ${fechaCorta(estado.ventanaDesde)}) más todo lo que sigue abierto. El historial completo vive en SharePoint.`;
}

/**
 * Historial (antes Cerrados, 2026-09-08): todo folio que ya no está en planta —cerrado, anulado o rechazado— dentro de la
 * ventana cargada, del más reciente al más viejo, con su Ticket para reimprimir (la línea «Emitido» del ticket lleva la hora
 * de la reimpresión) y su certificado. El buscador filtra por folio, placa, manifiesto, carrier, ticket y folio CT (U-76).
 */
function pintarHistorial() {
    const q = normaliza($('baBusca').value.trim());
    const pega = filtroTexto(q);   // C-26
    const momento = e => e.TaraHora || e.AnuladoEl || e.Arribo || '';
    const todos = estado.embarques.filter(e => e.Title && !enPlanta(e)).sort((a, b) => momento(b).localeCompare(momento(a)));
    const visibles = todos.filter(e => pega(e.Title, e.PlacaTractor, e.PlacaPlana, e.Manifiesto, nombreDe(estado.carriers, e.CarrierId), e.TicketBascula, ...certificadosDe(e).map(c => c.Title)));
    const estadoH = e => {
        if (e.Etapa !== 'cerrado') return el('span', 'paso e-' + siguientePaso(e).tono, siguientePaso(e).texto);
        const vig = certificadoVigente(e);
        return vig ? el('span', 'paso e-ok', `cerrada · ${vig.Title}`) : el('span', 'paso e-ok', 'cerrada');
    };
    tablaGondolas($('baCerrados'), [
        { t: 'Folio', m: 'a', v: folioG },
        { t: 'Placa', m: 'b', v: e => e.PlacaTractor, clase: 'mono' },
        { t: 'Programa', m: 'c', v: programaG },
        { t: 'Estado', m: 'e', v: estadoH },
        { t: 'Neto', m: 'd', num: true, v: e => (e.NetoKg ? kgG(e.NetoKg) : null), clase: 'mono' },
        { t: 'Fecha', m: 'x', v: e => horaMexico(momento(e)), clase: 'mono mudo' }
    ], visibles, e => botonesCerrada(e, visibles), q ? `Nada coincide en los últimos ${CONFIG.ventanaDias} días.` : 'Ningún folio cerrado en la ventana cargada.', e => e.Etapa);
    $('baBuscaCuenta').textContent = q ? plural(visibles.length, 'resultado') : '';
    // U-76 (v0.42.0): cuántas cerradas siguen sin certificado, sobre todo lo cargado (no sobre lo filtrado).
    const cerradas = todos.filter(e => e.Etapa === 'cerrado'), sinCert = cerradas.filter(e => !certificadoVigente(e)).length, otras = todos.length - cerradas.length;
    $('gPieHistorial').textContent = todos.length ? `${cerradas.length} con neto${sinCert ? ` · ${sinCert} sin certificado` : ''}${otras ? ` · ${plural(otras, 'anulado o rechazado', 'anulados o rechazados')}` : ''}.` : '';
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
    // U-70 (v0.40.0): una gondola con certificado vigente no se anula dejandolo vivo. Solo gerencia cancela certificados,
    // asi que solo gerencia puede anularla; y el certificado se cancela PRIMERO con el mismo motivo: si luego falla la
    // anulacion, queda una gondola cerrada sin certificado (se vuelve a emitir), nunca uno vigente sobre una anulada.
    const vigC = certificadoVigente(e);
    if (vigC && !PUEDE.emitirCertificado(estado.rol)) { avisar(`${e.Title} tiene el certificado ${vigC.Title} vigente: solo gerencia puede anular esta góndola (el certificado se cancela con ella).`, 'error'); return; }
    await escribiendo(btn, async () => {   // C-24 / C-23: boton deshabilitado y sin refresco mientras el confirm esta abierto
    const { ok, motivo } = await confirmar({
        titulo: `Anular ${e.Title}`, peligro: true, ok: 'Anular con este motivo', motivo: true,
        texto: `${e.PlacaTractor} · etapa ${e.Etapa}${e.NetoKg ? ` · neto ${kgG(e.NetoKg)}` : e.BrutoKg ? ` · bruto ${kgG(e.BrutoKg)}` : ''}. El folio ${e.Title} queda anulado y NO se vuelve a usar; el siguiente pesaje nace con folio nuevo. Las fotos ya subidas se conservan.${vigC ? ` Su certificado de tratamiento ${vigC.Title} se CANCELA con el mismo motivo: su QR dirá «cancelado».` : ''}`,
        etiquetaMotivo: 'Motivo de la anulación'
    });
    if (!ok) return;
    try {
        await refrescarCliente();
        if (vigC) {
            const cc = camposCancelacion(`Góndola ${e.Title} anulada: ${motivo}`);
            try { await estado.cliente.actualizarRenglon(estado.siteId, L.certificados, vigC.id, cc); }
            catch (err) { throw new Error(`no se canceló el certificado ${vigC.Title} (${err.message}); la góndola NO se anuló`); }
            aplicar('certificados', vigC, cc);
        }
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
    mostrarAsistente(e, 'baPesar', fase === 'bruto' ? 3 : 4);   // tanda 5: el pesaje es un paso del asistente, no una tarjeta bajo la lista
    // U-44 (v0.29.0): las dos fases titulan con la PLACA (lo que el basculista ve en el camión) y la tara además con el folio;
    // el manifiesto va al subtítulo. Antes la tara decía solo «Tara · E-26-00004» y el bruto no traía el manifiesto.
    $('baTitulo').textContent = fase === 'bruto' ? `Bruto · ${e.PlacaTractor}` : `Tara · ${e.PlacaTractor} · ${e.Title}`;
    const u = porId(estado.unidades, e.UnidadId);
    const manif = e.Manifiesto ? `Manifiesto ${e.Manifiesto}. ` : '';
    $('baSub').textContent = fase === 'bruto' ? `${manif}Primera pasada: la góndola cargada. Al guardar nace el folio E-.` :
        `${manif}Segunda pasada: la góndola vacía. Bruto ${kgG(e.BrutoKg)}${u && u.CapacidadKg ? ` · capacidad ${kgG(u.CapacidadKg)}` : ''}.`;   // U-87: con separador de miles, como la resta
    $('baKgLabel').textContent = fase === 'bruto' ? 'Bruto (kg)' : 'Tara (kg)';
    $('baKg').value = '';
    $('baFotoPrevia').classList.add('oculto');
    $('baFotoEstado').textContent = 'Sin foto todavía'; $('baFotoEstado').classList.remove('lista');
    $('btnFoto').textContent = 'Tomar foto';
    $('baNetoVivo').classList.toggle('oculto', fase !== 'tara');
    $('baEcuacion').classList.toggle('solo', fase !== 'tara');   // M7: el bruto es un solo término; la tara, la resta entera
    $('btnGuardarPeso').textContent = fase === 'bruto' ? 'Guardar bruto ›' : 'Guardar tara';
    if (fase === 'tara') {
        $('baNvBruto').textContent = Number(e.BrutoKg).toLocaleString('es-MX');
        $('baNvHora').textContent = e.BrutoHora ? ` · ${horaMexico(e.BrutoHora, 'hora')}` : '';
        $('baNvCap').textContent = u && u.CapacidadKg ? Number(u.CapacidadKg).toLocaleString('es-MX') : '';
        $('baNvCapBox').classList.toggle('oculto', !(u && u.CapacidadKg));
        revisarNeto();
    }
    // v0.39.0: el ticket de bascula (el que imprime el indicador) se captura al cerrar la TARA — cubre tara y destara,
    // y es lo que va impreso en el certificado de esa gondola (Carlos, 2026-09-22).
    $('baTicketCampo').classList.toggle('oculto', fase !== 'tara');
    $('baTicketBascula').value = '';
    $('baAvisoNeto').classList.add('oculto');
    $('baMotivoNetoCampo').classList.add('oculto');
    $('baMotivoNeto').value = '';
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
    $('baNvNeto').textContent = neto === null ? '—' : neto.toLocaleString('es-MX');
    $('baNetoVivo').classList.toggle('neto-ok', neto !== null && !aviso);
    $('baNetoVivo').classList.toggle('neto-mal', neto !== null && !!aviso);
    $('btnGuardarPeso').textContent = neto === null || neto <= 0 ? 'Guardar tara' : `Guardar tara · neto ${neto.toLocaleString('es-MX')} kg ›`;
    return aviso;
}

async function guardarPeso() {
    const p = estado.pesando; if (!p) return;
    const kg = Number($('baKg').value);
    if (!Number.isFinite(kg) || kg <= 0) { avisar('Captura el peso en kilogramos.', 'error'); return; }
    if (!estado.fotoBytes) { avisar('Falta la foto del indicador: es lo que hace comprobable un peso tecleado.', 'error'); return; }
    // El ticket de bascula es obligatorio al cerrar: despues no hay donde capturarlo, y el certificado de la gondola lo lleva impreso.
    if (p.fase === 'tara' && !$('baTicketBascula').value.trim()) { avisar('Falta el número del ticket de báscula: es lo que va impreso en el certificado de esta góndola.', 'error'); $('baTicketBascula').focus(); return; }
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
        soltarFotoPrevia(); estado.pesando = null; estado.fotoBytes = null;   // C-51: el JPEG ya subió; la pausa y el ticket no lo necesitan
        pintarTicket(e);   // tras el bruto también: «Ver ticket» de la pausa lo reimprime con el folio recién nacido
        pintarListasGondolas();   // la lista de atrás ya al día; el asistente sigue en pantalla
        // M5 / M8: el bruto lleva a la pausa de la descarga; la tara, al ticket con «Terminar».
        if (p.fase === 'bruto') mostrarPausa(e); else mostrarAsistente(e, 'baTicketCaja', 5);
    } catch (err) {
        pintarCapturado(p.embarque);   // decisión 11: si otra sesión avanzó la góndola, lo capturado ya dice lo suyo
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
    // Decisión 11 (tanda 5): gana el primero, y se dice qué dejó y quién. Nada se sube: el lote todavía no existe.
    if (e.Etapa !== 'compuerta') throw new Error(`esta góndola ${yaCapturado(e, quien)} (lo movió otra sesión). Tu peso no se guardó; vuelve a la lista.`);
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
    if (e.Etapa !== 'bruto') throw new Error(`esta góndola ${yaCapturado(e, quien)} (lo movió otra sesión). Tu tara no se guardó; vuelve a la lista.`);   // decisión 11
    paso('Subiendo la foto…');
    const lote = await subirEvidencia(e.Title, 'tara', kg, paso);
    paso('Cerrando el embarque…');
    const neto = Number(e.BrutoKg) - kg;
    const campos = limpiar({ Etapa: 'cerrado', TaraKg: kg, TaraHora: ahora, TaraFoto: lote.ref, NetoKg: neto,
        TicketBascula: $('baTicketBascula').value.trim() || null,
        InicioAlmacen: e.BrutoHora || ahora,
        Notas: aviso ? `Neto fuera de banda (${aviso}). Motivo: ${$('baMotivoNeto').value.trim()}` : null });
    await guardarConLote(e, lote, campos, paso);
    aplicar('embarques', e, campos);
    avisar(`Góndola ${e.Title} cerrada: neto ${kgG(neto)}.`, 'bien');   // U-87: el vocabulario de la sección
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
    // C-53 (v0.51.0): el renglón se resuelve por id en cada ‹ ›, como vivo(): la lista con que se pintó puede traer copias viejas.
    const actual = () => vivo('embarques', lista[i]);
    const pintar = () => {
        pintarTicket(actual(), $('ticketPop'));
        $('tkPos').textContent = `${i + 1} de ${lista.length} · ${actual().Title}`;
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
        ['Pozo / corriente', pre ? `${pre.Pozo || '—'} · ${etiquetaCorriente(e.CorrienteDeclarada || pre.Corriente)}` : (etiquetaCorriente(e.CorrienteDeclarada) || '—')],   // C-55
        ['Transportista', nombreDe(estado.carriers, e.CarrierId)],
        ['Placas', [e.PlacaTractor, e.PlacaPlana].filter(Boolean).join(' / ') || '—'], ['Operador', e.ChoferNombre || nombreDe(estado.choferes, e.ChoferId)],   // U-87: placas sin la diagonal colgando
        ['Arribo', horaCorta(e.Arribo)],
        ['Bruto', e.BrutoKg ? `${kgG(e.BrutoKg)} · ${horaCorta(e.BrutoHora)}` : '—'],
        ['Tara', e.TaraKg ? `${kgG(e.TaraKg)} · ${horaCorta(e.TaraHora)}` : 'pendiente'],
        ['NETO', e.NetoKg ? kgG(e.NetoKg) : 'pendiente'],
        ['Ticket de báscula', e.TicketBascula || '—'],
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
function gondolasDe(p, conteo = conteoRecibidas()) {
    return { rec: conteo.get(Number(p.id)) || 0, esp: Number(p.GondolasEsperadas) || 0 };
}
/** C-68 (v0.63.0): PreAltaId -> recibidas, en UNA pasada; pintarPrealtas la calcula una vez por repintado (antes, 3-5 por programa). */
function conteoRecibidas(emb = embarquesDeProgramas()) {
    const m = new Map();
    for (const e of emb) if (e.Etapa !== 'anulado' && e.Etapa !== 'rechazado') { const k = Number(e.PreAltaId); m.set(k, (m.get(k) || 0) + 1); }
    return m;
}
function sinMovimientoDe(p, emb = embarquesDeProgramas()) { return prealtaSinMovimiento(p, emb, CONFIG.sinMovimientoDias); }
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
/** Las tres pestañas de Pre-altas (rediseño tanda 1, v0.53.0): estado de la pre-alta -> panel. */
const VISTAS_PREALTAS = { borrador: 'ppBorradores', firmada: 'ppFirmadas', cerrada: 'ppCerradas' };
/** U-101 (v0.57.0): el Estado con la palabra de su pestaña; «borrador» es el valor de la columna, no lo que lee la gente. */
const ESTADO_PREALTA = { borrador: 'por firmar', firmada: 'firmada', cerrada: 'cerrada' };
const estadoPrealta = e => ESTADO_PREALTA[e] || e || 'sin estado';
function elegirVistaPrealtas(v) {
    estado.vistaPrealtas = VISTAS_PREALTAS[v] ? v : 'borrador';
    for (const b of document.querySelectorAll('#paTabs button, .pa-kpis button')) b.setAttribute('aria-pressed', String(b.dataset.pa === estado.vistaPrealtas));
    for (const [k, id] of Object.entries(VISTAS_PREALTAS)) $(id).hidden = k !== estado.vistaPrealtas;
}
/**
 * Pantalla principal, tanda 2 (v0.61.0; la K del lienzo EhJVtU3d4X6jUNuhJb6KQR): un programa es un renglón de columnas reales
 * —Programa · Folio · Corriente · Carrier · fecha de la pestaña · Góndolas · Ver— y los renglones van en cebra (`non`, lo
 * alterna pintarPrealtas porque las barras de mes cortan el nth-child). Corriente y carrier van juntos en `.cc` para que en
 * celular (la L) bajen a un renglón; en escritorio `.cc` es display: contents y cada uno es su columna. Las etiquetas de
 * excepción (estado que no es el de su pestaña, «sin firma», «¿se cierra?») van junto al nombre, como antes.
 */
const FECHA_PESTANA = { borrador: '1er envío', firmada: 'Firmada', cerrada: 'Cerrada' };
const fechaPestanaCorta = (p, grupo) => {
    const f = fechaDePestana(p, grupo);
    return !f ? '—' : grupo === 'borrador' ? fechaCorta(f) : fechaCorta(fechaMexico(new Date(f)));
};
/** C-69 (v0.63.0): las recibidas de programas SIN estimado van aparte (`sin`): sumadas a las otras inflaban «15/20». */
function sumaGondolas(ps, conteo) {
    return ps.reduce((a, p) => { const g = gondolasDe(p, conteo); if (g.esp) { a.rec += g.rec; a.esp += g.esp; } else a.sin += g.rec; return a; }, { rec: 0, esp: 0, sin: 0 });
}
/** El texto del subtotal de mes y del Total, uno solo (C-69): «3/12 góndolas», «3/12 góndolas + 2 sin estimado», o solo las recibidas. */
function textoGondolas({ rec, esp, sin }) {
    if (!esp) return `${plural(sin, 'góndola recibida', 'góndolas recibidas')} · sin estimado`;
    return `${rec}/${esp} góndolas` + (sin ? ` + ${sin} sin estimado` : '');
}
/** U-114 (v0.63.0): la rejilla es tabla para el lector de pantalla — role table/row/columnheader/cell, sin cambiar el CSS. */
const celda = (tag, clase, texto, rol = 'cell') => { const c = el(tag, clase, texto); c.setAttribute('role', rol); return c; };
/**
 * C-74 (v0.63.0): el renglón de programa arma sus celdas explícitas —ya no parchea el DOM de renglon() con firstChild/lastChild—
 * y el CSS las ubica por NOMBRE de área de rejilla (--pa-areas), no por índice de columna. U-109: tocar el renglón abre la
 * ficha igual que Ver (el botón sigue siendo el control de teclado; no se anidan interactivos).
 */
function renglonPrograma(p, grupo, ctx, non = false) {
    const abrir = () => verPrealta(vivo('prealtas', p));
    const r = el('div', 'renglon prog' + (non ? ' non' : '')); r.setAttribute('role', 'row');
    const pr = celda('div', 'pr'), t = el('div', 't', p.Title);
    pr.appendChild(t);
    if (p.Estado !== grupo) t.appendChild(etiqueta(estadoPrealta(p.Estado), p.Estado));
    if (grupo === 'firmada' && !prealtaFirmada(p)) t.appendChild(etiqueta('sin firma', 'vencida'));   // S-01
    const sm = grupo === 'firmada' ? sinMovimientoDe(p, ctx.emb) : null;
    if (sm) { t.appendChild(etiqueta('¿se cierra?', 'aviso')); pr.appendChild(el('p', 'pista', `${sm.motivo}. Sigue saliendo en la puerta hasta que alguien cierre el programa.`)); }
    const cc = el('div', 'cc');
    const cor = etiquetaCorriente(p.Corriente) || 'sin corriente', car = nombreDe(estado.carriers, p.CarrierId);
    cc.appendChild(celda('span', 'cor', cor)).title = cor;   // U-112: en una línea con elipsis; el nombre completo al pasar el ratón
    cc.appendChild(celda('span', 'car', car)).title = car;
    const av = celda('div', 'gon'); av.appendChild(barraAvance(gondolasDe(p, ctx.conteo)));
    const ver = celda('div', 'ver'); ver.appendChild(botonAccion({ texto: 'Ver', alClic: abrir }));
    for (const c of [pr, celda('span', 'fol', p.Campana || '—'), cc, celda('span', 'fe', fechaPestanaCorta(p, grupo)), av, ver]) r.appendChild(c);
    r.addEventListener('click', e => { if (!e.target.closest('button')) abrir(); });
    return r;
}
function encabezadoProgramas(grupo) {
    const h = el('div', 'pa-cab'); h.setAttribute('role', 'row');
    for (const s of ['Programa', 'Folio', 'Corriente', 'Carrier', FECHA_PESTANA[grupo], 'Góndolas']) h.appendChild(celda('span', '', s, 'columnheader'));
    h.appendChild(celda('span', '', '', 'columnheader')).setAttribute('aria-label', 'Abrir');
    return h;
}
/** U-111 (v0.63.0): la barra de mes va en la misma rejilla que los renglones y su subtotal cae en la columna Góndolas. */
function barraMes({ mes, ps }, ctx) {
    const b = el('div', 'pa-mes'); b.setAttribute('role', 'row');
    const m = celda('div', 'mt'); m.appendChild(el('b', '', mes)); m.appendChild(el('span', 'n', plural(ps.length, 'programa')));
    b.appendChild(m);
    b.appendChild(celda('span', 'g', textoGondolas(sumaGondolas(ps, ctx.conteo))));
    return b;
}
function pintarPrealtas() {
    cerrarForma('paDetalle');
    const captura = PUEDE.capturarPrealta(estado.rol);
    $('paInicio').classList.toggle('oculto', !captura);
    $('btnNuevaPrealta').classList.toggle('oculto', !captura);   // además del contenedor: visible() de la E2E mira el botón
    // Un Estado que no sea uno de los tres cae en cerradas para que no desaparezca de la vista;
    // ahi su etiqueta lo delata (renglonPrograma la pinta cuando no coincide con la pestaña).
    const grupos = { borrador: [], firmada: [], cerrada: [] };
    for (const p of [...estado.prealtas].sort((a, b) => b.id - a.id)) (grupos[p.Estado] || grupos.cerrada).push(p);

    // Pantalla principal, tanda 1 (v0.60.0): los recientes viven en el pop-up #paRecientes; el botón de la banda solo sale si hay.
    const bases = captura ? basesRecientes(estado.prealtas) : [];
    $('btnPaBases').classList.toggle('oculto', !bases.length);
    if (!bases.length) cerrarForma('paRecientes');
    const cb = $('paBases'); cb.textContent = '';
    for (const p of bases) {
        const b = el('button', 'pa-base'); b.type = 'button';
        b.appendChild(el('b', '', p.Title));
        b.appendChild(el('small', '', `${etiquetaCorriente(p.Corriente) || 'sin corriente'} · ${nombreDe(estado.carriers, p.CarrierId)}`));
        b.appendChild(el('span', 'usar', 'Usar como base ›'));
        b.addEventListener('click', () => { cerrarForma('paRecientes'); usarComoBase(vivo('prealtas', p)); });
        cb.appendChild(b);
    }

    const emb = embarquesDeProgramas(), ctx = { emb, conteo: conteoRecibidas(emb) };   // C-68: una pasada por repintado
    const vacio = { borrador: 'Ninguna por firmar.', firmada: 'Ninguna firmada: la puerta no puede recibir.', cerrada: 'Ninguna cerrada.' };
    const TITULO_TABLA = { borrador: 'Programas por firmar', firmada: 'Programas firmados', cerrada: 'Programas cerrados' };
    for (const [grupo, cont] of [['borrador', 'paBorradores'], ['firmada', 'paFirmadas'], ['cerrada', 'paCerradas']]) {
        const c = $(cont); c.textContent = '';
        const ps = grupos[grupo];
        if (ps.length) {
            c.setAttribute('role', 'table'); c.setAttribute('aria-label', TITULO_TABLA[grupo]);   // U-114
            // Tanda 2 (v0.61.0): Por firmar sigue el orden de captura, sin meses; Firmadas y Cerradas van por mes de su fecha.
            c.appendChild(encabezadoProgramas(grupo));
            let k = 0;
            for (const b of grupo === 'borrador' ? [{ mes: null, ps }] : mesesPrealtas(ps, grupo)) {
                if (b.mes) c.appendChild(barraMes(b, ctx));
                for (const p of b.ps) c.appendChild(renglonPrograma(p, grupo, ctx, k++ % 2 === 1));
            }
            const tot = el('div', 'pa-total'); tot.setAttribute('role', 'row');
            tot.appendChild(celda('b', 'tt', 'Total'));
            tot.appendChild(celda('span', 'g', textoGondolas(sumaGondolas(ps, ctx.conteo))));
            c.appendChild(tot);
        } else { c.removeAttribute('role'); c.removeAttribute('aria-label'); }
        if (!ps.length) c.appendChild(el('p', 'vacio', estado.prealtas.length || grupo !== 'borrador' ? vacio[grupo] : 'No hay pre-altas. La primera góndola no puede entrar sin una firmada.'));
    }
    $('paNBorradores').textContent = String(grupos.borrador.length);
    $('paNFirmadas').textContent = String(grupos.firmada.length);
    $('paNCerradas').textContent = String(grupos.cerrada.length);
    for (const [id, g] of [['paKBorradores', 'borrador'], ['paKFirmadas', 'firmada'], ['paKCerradas', 'cerrada']]) $(id).textContent = String(grupos[g].length);   // la banda
    // Las pestañas con algo que atender se marcan: borradores por firmar y firmadas sin firma o que ya no se mueven.
    // U-110 (v0.63.0): los conteos de la banda siguen la MISMA regla (antes, «Por firmar» siempre ámbar y «Firmadas» siempre verde).
    const alerta = { borrador: grupos.borrador.length > 0, firmada: grupos.firmada.some(p => sinMovimientoDe(p, emb) || !prealtaFirmada(p)), cerrada: false };
    $('paNBorradores').classList.toggle('alerta', alerta.borrador);
    $('paNFirmadas').classList.toggle('alerta', alerta.firmada);
    for (const b of document.querySelectorAll('.pa-kpis button')) b.classList.toggle('alerta', alerta[b.dataset.pa]);

    const sb = sumaGondolas(grupos.borrador, ctx.conteo);
    $('paResBorradores').textContent = grupos.borrador.length
        ? `${plural(sb.esp, 'góndola comprometida', 'góndolas comprometidas')}; ninguna puede entrar hasta que se firme.`
        : '';
    // Tanda 2 (v0.61.0): el pie de Firmadas y Cerradas repetía la cuenta que ya da el renglón Total; queda vacío. Sin firmadas
    // el mensaje de la tabla sigue diciendo que la puerta no puede recibir (U-23, U-102).
    $('paResFirmadas').textContent = '';
    $('paResCerradas').textContent = '';
    elegirVistaPrealtas(estado.vistaPrealtas);
}

function nuevaPrealta() {
    estado.prealtaEdit = null;
    $('paFormaTitulo').textContent = 'Nueva'; $('btnGuardarPrealta').textContent = 'Guardar para firma';
    estado.paAsis = { paso: 1, max: 1, revisar: false, genEdit: false, genDeId: null, otro: false, baseId: null };   // C-67: ids, no objetos
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false), c => c.id, c => c.Title);
    $('paCarrier').value = '';   // U-40: opciones() ya conserva el value; una pre-alta nueva empieza sin carrier
    pintarUnidadesChoferesPrealta();
    for (const id of ['paTitulo', 'paCliente', 'paPozoTitulo', 'paMes', 'paGenerador', 'paGeneradorDireccion', 'paGeneradorRegistro', 'paPozo', 'paFecha', 'paGondolas', 'paCorreoFecha', 'paCorreoRemitente', 'paNotas']) $(id).value = '';
    $('paCorriente').value = '';
    $('paMes').value = String(new Date().getFullYear());   // U-54: casi siempre es el año en curso
    armarTituloPrealta();
    abrirAsistentePrealta();
}
/**
 * P2 (rediseño tanda 1, v0.53.0): «Usar como base» abre la pre-alta nueva con cliente, generador, corriente y carrier del
 * programa elegido; pozo y envío quedan vacíos y el foco va al pozo. Si el carrier ya no está activo se deja sin elegir.
 * Tanda 2 (v0.54.0): el asistente abre en el paso 2 con el 1 y el 3 ya hechos.
 */
async function usarComoBase(p) {
    if (!p || !PUEDE.capturarPrealta(estado.rol) || !(await soltarCapturaPrealta())) return;
    nuevaPrealta();
    const f = textoDe;
    $('paCliente').value = clienteDe(p); $('paCorriente').value = f(p.Corriente);
    llenarGeneradorPrealta(p);
    if (estado.carriers.some(c => c.Activo !== false && Number(c.id) === Number(p.CarrierId))) $('paCarrier').value = f(p.CarrierId);
    pintarUnidadesChoferesPrealta(seleccionDe(p));   // U-94 (v0.56.0): la seleccion de la base, no todas
    armarTituloPrealta();
    Object.assign(estado.paAsis, { paso: 2, max: 3, baseId: p.id, genDeId: p.id });
    marcarOtroClientePrealta(); pintarAsistentePrealta();
    $('paAsis').dataset.huella = huellaForma($('paAsis'));   // lo copiado de la base no cuenta como captura: Cancelar no pregunta
    $('paPozoTitulo').focus();
}
/**
 * Editar un BORRADOR (v0.19.10, Carlos 2026-09-08: «ver el borrador no me da la opcion de editarlo»). Mismo
 * formulario que la alta, con `estado.prealtaEdit`; guardar es un PATCH que conserva Estado, Campana y quien capturo.
 * Una firmada no se edita: la puerta ya la usa; se cierra y se abre otra.
 */
async function editarPrealta() {
    const p = estado.prealtaAbierta; if (!p || p.Estado !== 'borrador' || !PUEDE.capturarPrealta(estado.rol)) return;
    if (!(await soltarCapturaPrealta())) return;   // U-95: el asistente conservado con una pre-alta a medias no se pisa callado
    estado.prealtaEdit = p;
    $('paFormaTitulo').textContent = `Editar ${p.Title}`; $('btnGuardarPrealta').textContent = 'Guardar cambios';
    estado.paAsis = { paso: 5, max: 5, revisar: true, genEdit: false, genDeId: null, otro: false, baseId: null };
    opciones($('paCarrier'), estado.carriers.filter(c => c.Activo !== false || Number(c.id) === Number(p.CarrierId)), c => c.id, c => c.Title);
    const f = textoDe;   // C-26
    partirTituloPrealta(f(p.Title)); $('paCorriente').value = f(p.Corriente); llenarGeneradorPrealta(p);
    $('paPozo').value = f(p.Pozo); $('paCarrier').value = f(p.CarrierId);
    $('paFecha').value = p.FechaEstimada ? fechaCorta(p.FechaEstimada) : ''; $('paGondolas').value = f(p.GondolasEsperadas);
    $('paCorreoFecha').value = p.CorreoFecha ? fechaCorta(p.CorreoFecha) : ''; $('paCorreoRemitente').value = f(p.CorreoRemitente); $('paNotas').value = f(p.Notas);
    pintarUnidadesChoferesPrealta(seleccionDe(p));   // las que la pre-alta ya tenia; si no tenia ninguna guardada, todas
    armarTituloPrealta();
    cerrarForma('paDetalle');
    abrirAsistentePrealta();
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
const seleccionDe = p => ({ unidades: lista(p.UnidadesIds), choferes: lista(p.ChoferesIds) });
/** U-92 (v0.56.0): si la corriente nueva no la ampara el carrier elegido, se suelta y SE DICE (antes se soltaba callado). */
function cambiarCorrientePrealta(k, texto) {
    const antes = $('paCarrier').value ? porId(estado.carriers, $('paCarrier').value) : null;
    $('paCorriente').value = k;
    if (antes && !carrierAmpara(antes, k)) {
        $('paCarrier').value = ''; pintarUnidadesChoferesPrealta();
        avisar(`${antes.Title} no ampara «${texto.toLowerCase()}»: se quitó del programa con sus unidades y choferes. Elige otro carrier en Transporte.`, 'ojo');
    }
    pintarAsistentePrealta();
}
/**
 * Las casillas de unidades y choferes del carrier elegido. `sel` ({ unidades, choferes }: ids) marca solo esas; sin `sel`, o con una
 * lista vacia, quedan todas (es lo que la puerta entiende). C-62/U-94 (v0.56.0): quien llama pasa la seleccion que ya existia.
 */
function pintarUnidadesChoferesPrealta(sel) {
    const cid = Number($('paCarrier').value);
    const cajas = (cont, items, nombre, ids) => {
        cont.textContent = '';
        if (!items.length) { cont.appendChild(el('p', 'pista', `(ese carrier no tiene ${nombre} en el padrón)`)); return; }
        for (const it of items) {
            const l = el('label', 'chk'); const c = el('input'); c.type = 'checkbox'; c.value = String(it.id); c.checked = !ids || !ids.length || ids.includes(String(it.id));
            l.appendChild(c); l.appendChild(document.createTextNode(' ' + (it.PlacaPlana ? `${it.Title} / ${it.PlacaPlana}` : it.Title)));
            const v = etiquetaVigencia(nombre, it); if (v) l.appendChild(v);   // U-98 (v0.57.0): la vigencia a la vista, como en el padrón
            cont.appendChild(l);
        }
    };
    cajas($('paUnidades'), estado.unidades.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'unidades', sel && sel.unidades);
    cajas($('paChoferes'), estado.choferes.filter(u => Number(u.CarrierId) === cid && u.Activo !== false), 'choferes', sel && sel.choferes);
}
function marcados(id) { return [...$(id).querySelectorAll('input:checked')].map(c => c.value).join(';'); }
/** C-65 (v0.57.0): el generador de un programa a sus tres campos; antes se copiaba a mano en tres sitios. */
function llenarGeneradorPrealta(p) {
    const f = textoDe;
    $('paGenerador').value = f(p.Generador); $('paGeneradorRegistro').value = f(p.GeneradorRegistro); $('paGeneradorDireccion').value = f(p.GeneradorDireccion);
}
/** C-65: lo que la forma guarda, armado una vez para la alta y la edición (la alta agrega Estado, CapturadaPor y Campana). */
function camposPrealta() {
    return {
        Title: valorPa('paTitulo'), Cliente: parteTitulo('paCliente'), Generador: valorPa('paGenerador'), GeneradorRegistro: valorPa('paGeneradorRegistro'), GeneradorDireccion: valorPa('paGeneradorDireccion'),   // Cliente: esquema v8, tanda 3
        Pozo: valorPa('paPozo'), Corriente: $('paCorriente').value, CarrierId: Number($('paCarrier').value),
        UnidadesIds: marcados('paUnidades'), ChoferesIds: marcados('paChoferes'),
        FechaEstimada: aIsoDia($('paFecha').value), GondolasEsperadas: valorPa('paGondolas') ? Number(valorPa('paGondolas')) : null,
        CorreoFecha: aIsoDia($('paCorreoFecha').value), CorreoRemitente: valorPa('paCorreoRemitente'), Notas: valorPa('paNotas')
    };
}
// ---------------------------------------------------------------- asistente de pre-alta (rediseño tanda 2, v0.54.0)
// Artifact 3PrGGb7ruvmrFhYNrrqdU1 v9: cinco pasos, una pantalla a la vez, en la pestaña (no en un pop-up). El estado del
// recorrido vive en estado.paAsis { paso, max (el más lejano visitado), revisar (vino de «Cambiar»: Siguiente regresa a Revisar),
// genEdit, genDe (el programa del que salió el generador), otro (cliente nuevo), base }. Lo capturado vive en los campos de
// siempre: guardarPrealta no cambió. P8: el botón dice qué falta y, tocado, lleva el foco ahí.
const PASOS_PREALTA = ['Cliente', 'Pozo y corriente', 'Transporte', 'Envío', 'Revisar'];
const asistentePrealtaAbierto = () => !$('paAsis').classList.contains('oculto');
function abrirAsistentePrealta() {
    if (estado.pestana !== 'prealtas') irA('prealtas');   // Editar desde el detalle abierto en Hoy
    ocultarListoPrealta(); limpiarAvisos();
    $('p-prealtas').classList.add('asistiendo');
    $('paAsis').classList.remove('oculto');
    marcarOtroClientePrealta();   // C-64: al editar, un cliente que no es de los conocidos abre en «Otro»
    pintarAsistentePrealta();
    $('paAsis').dataset.huella = huellaForma($('paAsis'));
    window.scrollTo({ top: 0 });
    enfocarPasoPrealta();
}
function cerrarAsistentePrealta() {
    $('paAsis').classList.add('oculto'); $('p-prealtas').classList.remove('asistiendo');
    estado.paAsis = null; estado.prealtaEdit = null;
}
/** Cancelar o la miga: con algo cambiado desde que abrió, pregunta (U-09/U-34, lo que hacía Escape sobre el pop-up). */
async function salirAsistentePrealta() {
    if (!(await soltarCapturaPrealta())) return;
    cerrarAsistentePrealta(); pintarPrealtas();
}
/** U-95 (v0.57.0): true si el asistente se puede soltar — cerrado, sin cambios desde que abrió, o quien captura dijo Descartar. */
async function soltarCapturaPrealta() {
    const d = $('paAsis');
    if (!estado.paAsis || !hayCaptura(d) || huellaForma(d) === d.dataset.huella) return true;
    const { ok } = await confirmar({ titulo: 'Descartar lo capturado', peligro: true, ok: 'Descartar', texto: 'Esta pre-alta tiene datos sin guardar. Si sales, se pierden.' });
    return !!ok;
}
const valorPa = id => $(id).value.trim();
const fechaMala = id => { try { aIsoDia($(id).value); return false; } catch { return true; } };
/**
 * Lo que le falta a un paso para darse por hecho, con a dónde llevar el foco; null si está completo. 1–3 tienen obligatorios;
 * U-91/U-106 (v0.57.0): el rango del año (2) y las fechas y góndolas del 4 —opcionales, pero si se escriben, bien escritas— se
 * cotejan aquí y no al guardar, donde salían con el foco a un campo escondido.
 */
function faltaPrealta(n) {
    const primera = cont => $(cont).querySelector('.o:not(:disabled)') || $(cont).querySelector('input');
    if (n === 1) return valorPa('paCliente') ? null : { t: 'Elige el cliente', foco: () => primera('paClientes') };
    if (n === 2) {
        if (!valorPa('paPozoTitulo')) return { t: 'Escribe el pozo', foco: () => $('paPozoTitulo') };
        const anio = Number(valorPa('paMes')), hoy = new Date().getFullYear();   // U-54: entre el año pasado y el que viene
        if (!/^\d{4}$/.test(valorPa('paMes')) || anio < hoy - 1 || anio > hoy + 1) return { t: `Año: cuatro cifras, de ${hoy - 1} a ${hoy + 1}`, foco: () => $('paMes') };
        if (!$('paCorriente').value) return { t: 'Elige la corriente', foco: () => primera('paCorrientes') };
        return null;
    }
    if (n === 4) {
        if (fechaMala('paFecha')) return { t: 'Primer envío: fecha no válida', foco: () => $('paFecha') };
        if (valorPa('paGondolas') && !/^\d+$/.test(valorPa('paGondolas'))) return { t: 'Góndolas: un número entero', foco: () => $('paGondolas') };
        if (fechaMala('paCorreoFecha')) return { t: 'Correo: fecha no válida', foco: () => $('paCorreoFecha') };
        return null;
    }
    if (n === 3) {
        if (!$('paCarrier').value) return { t: 'Elige el carrier', foco: () => primera('paCarriers') };
        const ce = porId(estado.carriers, $('paCarrier').value);
        if (ce && !carrierAmpara(ce, $('paCorriente').value)) return { t: 'Elige un carrier que ampare la corriente', foco: () => primera('paCarriers') };
        for (const [cont, que] of [['paUnidades', 'una unidad'], ['paChoferes', 'un chofer']])
            if ($(cont).querySelector('input') && !$(cont).querySelector('input:checked')) return { t: `Marca al menos ${que}`, foco: () => $(cont).querySelector('input') };
    }
    return null;
}
function carrierAmpara(c, corriente) { const l = lista(c.Corrientes); return !corriente || !l.length || l.includes(corriente); }   // la misma regla que la compuerta
/**
 * P3: elegir un cliente conocido trae el generador de su último programa (con «Cambiar» en el paso 2). C-67 (v0.57.0): el
 * programa se resuelve por su id al tocar, no con el objeto que el renglón capturó al pintarse. v0.59.0: ya NO avanza al paso 2
 * (revierte U-105 a pedido de Carlos): el usuario pulsa «Siguiente».
 */
function elegirClientePrealta(clave, ultimoId) {
    const a = estado.paAsis;
    const cambio = parteTitulo('paCliente') !== clave;
    a.otro = false; $('paCliente').value = clave;
    const u = porId(estado.prealtas, ultimoId);
    if (u && (cambio || !valorPa('paGenerador'))) { llenarGeneradorPrealta(u); a.genDeId = u.Generador ? u.id : null; a.genEdit = false; }
    armarTituloPrealta(); pintarAsistentePrealta();
}
function resumenPasoPrealta(n) {
    if (n === 1) return parteTitulo('paCliente');
    if (n === 2) return [parteTitulo('paPozoTitulo'), etiquetaCorriente($('paCorriente').value)].filter(Boolean).join(' · ');
    if (n === 3) return $('paCarrier').value ? nombreDe(estado.carriers, $('paCarrier').value) : '';
    if (n === 4) return [valorPa('paFecha'), valorPa('paGondolas') && `${valorPa('paGondolas')} gón.`].filter(Boolean).join(' · ') || 'sin datos';
    return '';
}
/** Repinta los renglones de una lista de opciones sin perder el foco del teclado: el renglón enfocado se re-enfoca por su valor. */
function pintarOpcionesPa(cont, items) {
    const f = document.activeElement, v = f && f.parentElement === cont ? f.dataset.v : null;
    cont.textContent = '';
    for (const it of items) {
        const b = renglonOpcion(it);
        if (it.deshabilitada) b.disabled = true;
        b.addEventListener('click', it.alClic);
        cont.appendChild(b);
    }
    if (v !== null) { const n = [...cont.children].find(x => x.dataset.v === v); if (n) n.focus(); }
}
/**
 * C-64 (v0.58.0): el asistente se pinta por partes —la barra, el paso visible y el pie, más el resumen de escritorio— y los
 * pasos ocultos no se tocan: se pintan al entrar (irPasoPrealta). El pintor ya no escribe estado; a.otro lo fija quien carga
 * el cliente (marcarOtroClientePrealta). Antes eran 86 líneas que reconstruían los cinco pasos con cada tecla, Notas incluida.
 */
const PINTORES_PASO_PREALTA = { 1: pintarPaso1Prealta, 2: pintarPaso2Prealta, 3: pintarPaso3Prealta, 5: pintarPaso5Prealta };   // el 4 son campos fijos
function pintarAsistentePrealta() {
    const a = estado.paAsis; if (!a) return;
    pintarBarraPrealta(a);
    const pintor = PINTORES_PASO_PREALTA[a.paso]; if (pintor) pintor(a);
    pintarPiePrealta(a);
    pintarResumenPrealta(carrierPrealta());
}
const carrierPrealta = () => $('paCarrier').value ? porId(estado.carriers, $('paCarrier').value) : null;
/** «3 de 4» de las casillas de unidades o choferes; vacío si el carrier no tiene. */
function cuentaMarcadosPrealta(cont) {
    const t = $(cont).querySelectorAll('input').length;
    return t ? `${$(cont).querySelectorAll('input:checked').length} de ${t}` : '';
}
/** Un cliente escrito que no es de los conocidos es «Otro cliente». Lo llama quien carga el cliente en la forma, no el pintor. */
function marcarOtroClientePrealta() {
    const a = estado.paAsis, cli = parteTitulo('paCliente'); if (!a) return;
    a.otro = a.otro || (!!cli && !clientesPrealta(estado.prealtas).some(c => c.clave === cli));
}
function pintarBarraPrealta(a) {
    const n = a.paso;
    for (const b of $('paPasos').querySelectorAll('button')) {
        const i = Number(b.dataset.p), hecho = i !== n && i <= a.max && !faltaPrealta(i) && (i < n || a.revisar);
        b.className = i === n ? 'ahora' : hecho ? 'hecho' : i <= a.max ? 'visto' : '';
        b.disabled = i > a.max; if (i === n) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
        b.querySelector('small').textContent = hecho ? resumenPasoPrealta(i) : '';
    }
    for (let i = 1; i <= 5; i++) $('paPaso' + i).hidden = i !== n;
    $('paPasoK').textContent = `Paso ${n} de 5 · ${PASOS_PREALTA[n - 1]}${n === 4 ? ' · opcional' : ''}`;
}
/** Paso 1: clientes conocidos (por número de programas) y «Otro cliente». */
function pintarPaso1Prealta(a) {
    const cli = parteTitulo('paCliente');
    pintarOpcionesPa($('paClientes'), [...clientesPrealta(estado.prealtas).map(c => ({
        valor: c.clave, sel: !a.otro && cli === c.clave, titulo: c.clave, dato: plural(c.n, 'programa'),
        detalle: `${c.ultimo.Generador || 'sin generador'} · último ${c.ultimo.Campana || c.ultimo.Title}${c.ultimo.Pozo ? ', ' + c.ultimo.Pozo : ''}`,
        alClic: () => elegirClientePrealta(c.clave, c.ultimo.id)   // C-67: la clave y el id, no el objeto
    })), { valor: '__otro', sel: a.otro, titulo: 'Otro cliente', detalle: 'se captura una vez y queda para la próxima', clase: 'otro',
        alClic: () => { if (!a.otro) { a.otro = true; $('paCliente').value = ''; a.genDeId = null; armarTituloPrealta(); pintarAsistentePrealta(); } $('paCliente').focus(); } }]);
    $('paClienteOtro').classList.toggle('oculto', !a.otro);
}
/** Paso 2: la base, la corriente y el generador (aviso con «Cambiar», o sus campos si no hay o se pidió cambiarlo). */
function pintarPaso2Prealta(a) {
    const base = a.baseId ? porId(estado.prealtas, a.baseId) : null, genDe = a.genDeId ? porId(estado.prealtas, a.genDeId) : null;
    $('paBaseAviso').classList.toggle('oculto', !base);
    if (base) $('paBaseAviso').firstChild.textContent = $('paCarrier').value ? `Copiado de ${base.Title}: cliente, generador, corriente y transporte. Solo escribe el pozo.`
        : `Copiado de ${base.Title}: cliente, generador y corriente. Su carrier ya no está activo: escribe el pozo y elige otro en Transporte.`;   // U-94
    const corr = $('paCorriente').value;
    pintarOpcionesPa($('paCorrientes'), CORRIENTES.map(([k, t]) => ({ valor: k, sel: corr === k, titulo: t, detalle: k.startsWith('fluidos') ? 'fluido de perforación agotado' : 'recorte de perforación',
        alClic: () => { cambiarCorrientePrealta(k, t); } })));
    const gen = valorPa('paGenerador'), verCampos = a.genEdit || !gen;
    $('paGenAviso').classList.toggle('oculto', verCampos);
    $('paGenCampos').classList.toggle('oculto', !verCampos);
    if (!verCampos) {
        const reg = valorPa('paGeneradorRegistro');
        $('paGenAviso').firstChild.textContent = `Generador ${gen}${reg ? `, registro ${reg}` : ''}.${genDe ? ` Tomado de ${genDe.Campana || genDe.Title}.` : ''}`;
    }
}
/** Paso 3: el carrier se coteja ANTES de elegirlo (P5): vigencia a la vista; si su oficio no ampara la corriente, apagado. */
function pintarPaso3Prealta() {
    const cid = $('paCarrier').value, corr = $('paCorriente').value;
    const carriers = [...$('paCarrier').options].filter(o => o.value).map(o => porId(estado.carriers, o.value)).filter(Boolean);
    const elegido = carrierPrealta();
    // C-62 (v0.56.0): el pintor ya no suelta el carrier (lo hace cambiarCorrientePrealta, que avisa); un carrier que no ampara
    // —p. ej. al editar un borrador viejo— lo marca faltaPrealta(3) y Guardar no pasa.
    const cuenta = (col, c) => col.filter(x => Number(x.CarrierId) === Number(c.id) && x.Activo !== false).length;
    pintarOpcionesPa($('paCarriers'), carriers.map(c => {
        const ok = carrierAmpara(c, corr);
        return { valor: String(c.id), sel: String(c.id) === cid, titulo: c.Title, deshabilitada: !ok,
            dato: ok ? etiquetaVigencia('carriers', c) || undefined : etiqueta('no ampara', 'vencida'),
            detalle: ok ? `${plural(cuenta(estado.unidades, c), 'unidad', 'unidades')} · ${plural(cuenta(estado.choferes, c), 'chofer', 'choferes')}` : `Su oficio no ampara «${etiquetaCorriente(corr).toLowerCase()}»`,
            alClic: () => { if ($('paCarrier').value !== String(c.id)) { $('paCarrier').value = String(c.id); pintarUnidadesChoferesPrealta(); } pintarAsistentePrealta(); } };   // C-62: el mismo carrier no re-marca
    }));
    if (!carriers.length) $('paCarriers').appendChild(el('p', 'pista', 'No hay carriers activos en el padrón: da de alta uno.'));
    const h = elegido ? vigenciasPadron('carriers', elegido).filter(x => !sinFecha(x)) : [];
    $('paCarrierAviso').classList.toggle('oculto', !h.length);
    $('paCarrierAviso').classList.toggle('mal', h.some(x => !x.ok));
    // U-99 (v0.57.0): el remate de rechazo solo con algo ya vencido; «por vencer» (ok: true) dice su fecha límite, que ya trae el detalle.
    if (h.length) $('paCarrierAviso').firstChild.textContent = `${elegido.Title}: ${h.map(x => `${x.regla} ${x.detalle}`).join(' · ')}.${h.some(x => !x.ok) ? ' Si el envío llega con esto así, la puerta lo rechaza.' : ' Después de esa fecha la puerta lo rechaza.'}`;
    $('paTransporte').classList.toggle('oculto', !elegido);
    $('paNUnidades').textContent = cuentaMarcadosPrealta('paUnidades');
    $('paNChoferes').textContent = cuentaMarcadosPrealta('paChoferes');
}
/** Paso 5: Revisar, con «Cambiar» por sección (P9: cambias una y Siguiente te regresa aquí). */
function pintarPaso5Prealta(a) {
    const rv = $('paRevisar'); rv.textContent = '';
    const sec = (titulo, texto, paso, antes) => {
        const d = el('section'); d.appendChild(el('h3', '', titulo)); d.appendChild(el('p', '', texto));
        const b = el('button', 'enlace', 'Cambiar'); b.type = 'button'; b.addEventListener('click', () => { a.revisar = true; if (antes) antes(); irPasoPrealta(paso); }); d.appendChild(b);
        rv.appendChild(d);
    };
    const corr = $('paCorriente').value, elegido = carrierPrealta();
    const nom = valorPa('paTitulo'); const r0 = el('div', 'pa-nombre'); r0.appendChild(el('small', '', 'Programa')); r0.appendChild(el('b', 'mono', nom || '—')); rv.appendChild(r0);
    // U-90 (v0.57.0): cliente y generador en secciones aparte, porque viven en pasos distintos; «Cambiar» del generador abre
    // sus campos en el paso 2. U-104: lo que se guarda se repasa —dirección, pozo / instalación y notas incluidos—.
    sec('Cliente', parteTitulo('paCliente') || '—', 1);
    sec('Generador y origen', [valorPa('paGenerador') || 'sin generador', valorPa('paGeneradorRegistro'), valorPa('paGeneradorDireccion'), valorPa('paPozo') && `pozo / instalación ${valorPa('paPozo')}`].filter(Boolean).join(' · '), 2, () => { a.genEdit = true; });
    sec('Pozo y corriente', [parteTitulo('paPozoTitulo'), etiquetaCorriente(corr)].filter(Boolean).join(' · '), 2);
    sec('Transporte', elegido ? `${elegido.Title} · ${plural($('paUnidades').querySelectorAll('input:checked').length, 'unidad', 'unidades')} · ${plural($('paChoferes').querySelectorAll('input:checked').length, 'chofer', 'choferes')}` : '—', 3);
    sec('Envío y correo', [valorPa('paFecha') && `primer envío ${valorPa('paFecha')}`, valorPa('paGondolas') && plural(Number(valorPa('paGondolas')), 'góndola'), valorPa('paCorreoFecha') && `correo del ${valorPa('paCorreoFecha')}`, valorPa('paCorreoRemitente'), valorPa('paNotas') && `notas: ${valorPa('paNotas')}`].filter(Boolean).join(' · ') || 'sin datos (opcional)', 4);
}
/** P8: el botón dice qué falta; Atrás nombra a dónde vuelve. U-91 (v0.57.0): en Revisar, Guardar dice lo que le falte a 1–4. */
function pintarPiePrealta(a) {
    const n = a.paso, falta = n === 5 ? pendientePrealta() : faltaPrealta(n), sig = $('btnPaSiguiente'), gu = $('btnGuardarPrealta');
    $('btnPaAtras').hidden = n === 1; $('btnPaAtras').textContent = n > 1 ? `‹ ${PASOS_PREALTA[n - 2]}` : '';
    $('btnPaSaltar').hidden = n !== 4 || a.revisar;   // P7: el envío es opcional; desde «Cambiar» el botón ya regresa a Revisar
    sig.hidden = n === 5; gu.hidden = n !== 5;
    sig.classList.toggle('falta', n !== 5 && !!falta); gu.classList.toggle('falta', n === 5 && !!falta);
    sig.textContent = n === 5 ? '' : falta ? falta.t : a.revisar ? 'Volver a revisar ›' : n === 4 ? 'Revisar ›' : `Siguiente: ${PASOS_PREALTA[n]} ›`;
    if (n === 5) gu.textContent = falta ? falta.f.t : estado.prealtaEdit ? 'Guardar cambios' : 'Guardar para firma';
}
/** P4, escritorio: «Así va la pre-alta» a la derecha, en vivo. En celular no se ve: el nombre ya va arriba. */
function pintarResumenPrealta(carrier) {
    const r = $('paResumen'); r.textContent = '';
    r.appendChild(el('h2', '', 'Así va la pre-alta'));
    const dl = el('dl');
    const fila = (dt, dd) => { dl.appendChild(el('dt', '', dt)); const d = el('dd', dd ? '' : 'f'); if (dd instanceof Node) d.appendChild(dd); else d.textContent = dd || '—'; dl.appendChild(d); };
    const sec = t => dl.appendChild(el('span', 'sec', t));
    sec('Programa'); fila('Cliente', parteTitulo('paCliente')); fila('Pozo', parteTitulo('paPozoTitulo')); fila('Corriente', etiquetaCorriente($('paCorriente').value));
    sec('Generador'); fila('Razón social', valorPa('paGenerador')); fila('Registro', valorPa('paGeneradorRegistro'));
    sec('Transporte'); fila('Carrier', carrier ? carrier.Title : '');
    if (carrier) { fila('Unidades', cuentaMarcadosPrealta('paUnidades')); fila('Choferes', cuentaMarcadosPrealta('paChoferes')); }
    sec('Envío'); fila('Primer envío', valorPa('paFecha')); fila('Góndolas', valorPa('paGondolas'));
    r.appendChild(dl);
    r.appendChild(el('p', 'pista', 'Obligatorio: cliente, pozo, año, corriente y carrier. Lo demás se puede completar después.'));
}
function enfocarPasoPrealta() { const a = estado.paAsis; if (a) $('paPaso' + a.paso).querySelector('.pregunta').focus({ preventScroll: true }); }
function irPasoPrealta(n) {
    const a = estado.paAsis; if (!a) return;
    a.paso = n; a.max = Math.max(a.max, n);
    limpiarAvisos(); pintarAsistentePrealta(); window.scrollTo({ top: 0 }); enfocarPasoPrealta();
}
function siguientePrealta() {
    const a = estado.paAsis; if (!a) return;
    const f = faltaPrealta(a.paso);
    if (f) { const x = f.foco(); if (x) x.focus(); return; }
    const p = pendientePrealta();
    irPasoPrealta(a.revisar ? (p ? p.paso : 5) : a.paso + 1);   // U-92: si «Cambiar» dejo otro paso incompleto, va ahi
}
/** El primer paso de 1–4 al que le falta algo, con lo que falta; null si todo está para guardarse (C-61, U-91). */
function pendientePrealta() {
    for (const paso of [1, 2, 3, 4]) { const f = faltaPrealta(paso); if (f) return { paso, f }; }
    return null;
}
/** U-100 (v0.57.0): Enter en un campo del paso visible es «Siguiente»; en Notas (textarea) sigue siendo un renglón nuevo. */
$('paAsis').addEventListener('keydown', ev => {
    const t = ev.target;
    if (ev.key !== 'Enter' || ev.isComposing || t.tagName !== 'INPUT' || t.type === 'checkbox' || !estado.paAsis) return;
    ev.preventDefault();
    if (estado.paAsis.paso === 5) guardarPrealta(); else siguientePrealta();
});
for (const id of PARTES_TITULO) $(id).addEventListener('input', armarTituloPrealta);
$('paCliente').addEventListener('input', marcarOtroClientePrealta);   // C-64: antes lo derivaba el pintor en cada repintado
$('paAsis').addEventListener('input', pintarAsistentePrealta);   // C-64: sin «change», que repintaba otra vez al salir de cada campo
$('btnPaSiguiente').addEventListener('click', siguientePrealta);
$('btnPaSaltar').addEventListener('click', siguientePrealta);   // el paso 4 no tiene obligatorios: saltar es seguir sin llenarlo
/** P7: − / + de góndolas esperadas; nunca baja de 0. Asignar .value no dispara «input», así que se repinta aquí. */
function sumarGondolasPrealta(d) {
    $('paGondolas').value = String(Math.max(0, (parseInt($('paGondolas').value, 10) || 0) + d));
    pintarAsistentePrealta();
}
$('btnPaGonMenos').addEventListener('click', () => sumarGondolasPrealta(-1));
$('btnPaGonMas').addEventListener('click', () => sumarGondolasPrealta(1));
$('btnPaAtras').addEventListener('click', () => irPasoPrealta(estado.paAsis.paso - 1));
for (const b of $('paPasos').querySelectorAll('button')) b.addEventListener('click', () => irPasoPrealta(Number(b.dataset.p)));
$('btnPaMiga').addEventListener('click', salirAsistentePrealta);
$('btnPaGenCambiar').addEventListener('click', () => { estado.paAsis.genEdit = true; pintarAsistentePrealta(); $('paGenerador').focus(); });
/** U-93 (v0.57.0): «Quitar la base» suelta lo que usarComoBase copió (cliente, generador, corriente, carrier) y vuelve al paso 1; lo tecleado —pozo, año— se queda. */
function quitarBasePrealta() {
    const a = estado.paAsis; if (!a) return;
    for (const id of ['paCliente', 'paGenerador', 'paGeneradorRegistro', 'paGeneradorDireccion']) $(id).value = '';
    $('paCorriente').value = ''; $('paCarrier').value = ''; pintarUnidadesChoferesPrealta();
    Object.assign(a, { baseId: null, genDeId: null, genEdit: false, otro: false });
    armarTituloPrealta(); irPasoPrealta(1);
}
$('btnPaQuitarBase').addEventListener('click', quitarBasePrealta);

async function guardarPrealta() {
    if (!PUEDE.capturarPrealta(estado.rol)) { avisar('Tu rol no captura pre-altas.', 'error'); return; }
    if (!$('paTitulo').value.trim() || !$('paCorriente').value || !$('paCarrier').value) { avisar('Faltan cliente, pozo o año del programa, la corriente o el carrier.', 'error'); return; }
    // C-61 (v0.56.0): los pasos se revalidan al guardar. La barra de pasos deja llegar a Revisar sin pasar por Siguiente,
    // y unidades o choferes todos desmarcados se guardaban vacios, que la compuerta lee como «todas autorizadas».
    // U-91 (v0.57.0): el rango del año (U-54) y las fechas y góndolas del paso 4 viven ya en faltaPrealta: se salta al paso con el foco visible.
    const pend = pendientePrealta();
    if (pend) { irPasoPrealta(pend.paso); avisar(`${pend.f.t} antes de guardar.`, 'error'); const x = pend.f.foco(); if (x) x.focus(); return; }
    // C-59 (v0.56.0): se edita por el id que se abrio, aunque el refresco ya lo haya traido firmado: la relectura decide.
    // Antes, un borrador que otra sesion firmo caia al crearRenglon y nacia un programa duplicado.
    const edit = estado.prealtaEdit;
    await escribiendo('btnGuardarPrealta', async () => { try {   // C-24
        await refrescarCliente();
        if (edit) {
            const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, edit.id);
            if (fresco.Estado !== 'borrador') {
                aplicar('prealtas', edit, fresco);
                throw new Error(`ya no está por firmar: ${fresco.Estado === 'firmada' ? `la firmó ${quien(fresco.FirmadaPor) || 'otra sesión'}` : `la cerró ${quien(fresco.CerradaPor) || 'otra sesión'}`} mientras la editabas. Tus cambios no se guardaron; si hacen falta, cierra ese programa y captura otro.`);
            }
            const cambios = paraPatch(camposPrealta());   // C-65
            await estado.cliente.actualizarRenglon(estado.siteId, L.prealtas, edit.id, cambios);
            aplicar('prealtas', edit, cambios);   // C-23
            cerrarAsistentePrealta();
            if (estado.pestana === 'prealtas') pintarPrealtas(); else pintarInsignias();
            verPrealta(edit);
            avisar('Pre-alta actualizada. Sigue por firmar.', 'bien');   // U-101   // al final: abrir el detalle limpia los avisos
            return;
        }
        // C-04 (v0.22.0): la campana se escoge con la lista RELEIDA, no con la de memoria (que puede tener horas), y tras
        // escribir pasa por asegurarFolioUnico como R- y E-. Dos capturas en dos celulares recibian la misma L-AA-NNN.
        const todas = await estado.cliente.renglones(estado.siteId, L.prealtas);
        estado.prealtas = todas; reanclar();
        const campos = limpiar({
            ...camposPrealta(), Estado: 'borrador', CapturadaPor: estado.cuenta.username,   // C-65
            Campana: siguienteFolio('L', todas.map(p => p.Campana))   // L-AA-NNN (ticket 03)
        });
        const nuevo = await estado.cliente.crearRenglon(estado.siteId, L.prealtas, campos);
        await asegurarFolioUnico(nuevo, 'L');
        anclar('prealtas', nuevo);   // C-23
        cerrarAsistentePrealta();
        estado.vistaPrealtas = 'borrador';   // U-107 (v0.57.0): la lista abre donde quedó la nueva, no en la última pestaña elegida
        pintarPrealtas();
        mostrarListoPrealta(nuevo);
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}
/** P10: tras guardar una pre-alta nueva, la pestaña muestra qué sigue en vez de un aviso. «Capturar otra» abre el asistente vacío. */
function mostrarListoPrealta(p) {
    const carrier = porId(estado.carriers, p.CarrierId);
    $('paListoFolio').textContent = [p.Campana, p.Title].filter(Boolean).join(' · ');
    $('paListoPuerta').textContent = `Las góndolas de ${carrier ? carrier.Title : 'su carrier'} se cotejan solas contra este programa.`;
    $('p-prealtas').classList.add('en-listo'); $('paListo').classList.remove('oculto');
    window.scrollTo({ top: 0 });
    $('paListoTitulo').focus({ preventScroll: true });
}
function ocultarListoPrealta() { $('paListo').classList.add('oculto'); $('p-prealtas').classList.remove('en-listo'); }
$('btnPaOtra').addEventListener('click', nuevaPrealta);
$('btnPaIrLista').addEventListener('click', () => { ocultarListoPrealta(); pintarPrealtas(); });

// Tras firmar / cerrar / eliminar: se cierra el pop-up y se repinta la pestana que esta abierta (Hoy o Pre-altas) sin
// borrar el aviso. Antes saltaba a Pre-altas aunque se hubiera abierto desde Hoy (Carlos, 2026-09-08).
function trasCambioPrealta() {
    cerrarForma('paDetalle'); repintar();
}
function verPrealta(p) {
    estado.prealtaAbierta = p;
    estado.prealtaVista = { estado: p.Estado, huella: huellaPrealta(p) };   // C-60: lo que el detalle muestra, para cotejar al firmar
    abrirForma('paDetalle');   // antes de los avisos: con el dialog abierto, avisar() los pinta adentro
    $('paDetalleTitulo').textContent = `${p.Title} · ${estadoPrealta(p.Estado)}`;   // U-101
    const ul = $('paDetalleLista'); ul.textContent = '';
    const carrier = porId(estado.carriers, p.CarrierId);
    const filas = [
        ['Generador', `${p.Generador || '—'} · ${p.GeneradorRegistro || 'sin registro'}`], ['Pozo', p.Pozo || '—'], ['Corriente', etiquetaCorriente(p.Corriente) || '—'],
        ['Campaña', p.Campana || '—'], ['Carrier', carrier ? `${carrier.Title} · ${carrier.AutorizacionASEA || 'sin autorización'}` : '—'],
        ['Unidades', lista(p.UnidadesIds).map(id => { const u = porId(estado.unidades, id); return u ? `${u.Title}/${u.PlacaPlana || ''}` : `#${id}`; }).join(', ') || '—'],
        ['Choferes', lista(p.ChoferesIds).map(id => nombreDe(estado.choferes, id)).join(', ') || '—'],
        ['Primer envío', fechaCorta(p.FechaEstimada)], ['Correo', `${fechaCorta(p.CorreoFecha)} · ${p.CorreoRemitente || ''}`],
        ['Capturó', quien(p.CapturadaPor) || '—'], ['Firmó', p.FirmadaPor ? `${quien(p.FirmadaPor)} · ${horaCorta(p.FirmadaEl)}${prealtaCambioTrasFirma(p) ? ' · cambió después de firmarse: la puerta no la ve hasta volver a firmar' : p.Estado === 'firmada' && !prealtaFirmada(p) ? ' · sello sin firma: la puerta no la ve' : ''}` : '—'], ['Cerró', p.CerradaPor ? `${quien(p.CerradaPor)} · ${horaCorta(p.CerradaEl)}` : '—'], ['Notas', p.Notas || '—']   // U-28 / U-31 (v0.26.0)
    ];
    for (const [k, v] of filas) { const li = el('li', '', k); li.appendChild(el('span', 'd', v)); ul.appendChild(li); }
    // Cotejo automatico de vigencias (lo que el validador firma que reviso).
    const vg = $('paDetalleVigencias'); vg.textContent = '';
    const hallazgos = [];
    if (carrier) {
        const c = evaluarVigencia('Autorización ASEA del carrier', carrier.VigenciaASEA, 'legal', CONFIG.avisoVigenciaDias); if (c) hallazgos.push(c);
        if (p.Corriente && lista(carrier.Corrientes).length && !lista(carrier.Corrientes).includes(p.Corriente)) hallazgos.push({ clase: 'legal', regla: 'Corriente', detalle: `el oficio del carrier no ampara ${etiquetaCorriente(p.Corriente)}` });
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
    const { ok } = await confirmar({ titulo: 'Eliminar la pre-alta', peligro: true, ok: 'Eliminar',
        texto: `«${p.Title}» nunca se firmó, así que la puerta no la ha usado. Se borra el renglón; la campaña ${p.Campana || ''} queda libre.`, motivo: 'opcional', etiquetaMotivo: 'Por qué (opcional)' });
    if (!ok) return;
    try {
        await refrescarCliente();
        await estado.cliente.borrarRenglon(estado.siteId, L.prealtas, p.id);
        estado.prealtas = estado.prealtas.filter(x => x.id !== p.id);
        estado.prealtaAbierta = null;
        avisar('Pre-alta eliminada.', 'bien'); trasCambioPrealta();
    } catch (e) { avisar('No se pudo eliminar: ' + e.message, 'error'); }
}

async function firmarPrealta() {
    const p = estado.prealtaAbierta; if (!p) return;
    await escribiendo('btnFirmar', async () => {   // C-24 / C-23
    const reFirma = p.Estado === 'firmada';   // S-01: trae el sello pero no su renglon en PLANTA_Firmas; solo falta la firma
    const { ok } = await confirmar({ titulo: 'Firmar la pre-alta', ok: 'Firmar',
        texto: reFirma && prealtaCambioTrasFirma(p) ? `«${p.Title}» cambió después de que ${quien(p.FirmadaPor)} la firmó, así que la puerta ya no la ve. Revisa carrier, unidades, choferes y generador arriba: con tu firma queda amparado lo que dice HOY.`
            : reFirma ? `«${p.Title}» trae el sello de ${quien(p.FirmadaPor)} pero no su firma registrada, así que la puerta no la ve. Con tu firma queda completa; queda registrado quién y cuándo.`
            : `«${p.Title}». Con tu firma la puerta empieza a aceptar sus góndolas; queda registrado quién y cuándo.` });
    if (!ok) return;
    try {
        await refrescarCliente();
        // C-60 (v0.56.0): se relee el renglón antes de firmar. Si otra sesión lo cerró, lo firmó o le cambió lo que se firma
        // mientras el detalle estaba abierto, no se firma lo que no se vio: se ancla lo fresco y se pide volver a mirarlo.
        const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, p.id);
        const vista = estado.prealtaVista || { estado: p.Estado, huella: huellaPrealta(p) };
        const cambio = fresco.Estado !== vista.estado || huellaPrealta(fresco) !== vista.huella;
        if (cambio) { aplicar('prealtas', p, fresco); trasCambioPrealta(); verPrealta(vivo('prealtas', p)); throw new Error(fresco.Estado === 'cerrada' ? 'el programa ya está cerrado (lo cerró otra sesión).'
            : fresco.Estado === 'firmada' && !reFirma ? `ya lo firmó ${quien(fresco.FirmadaPor) || 'otra sesión'}.` : 'cambió desde que lo abriste. Ábrelo otra vez y revísalo antes de firmar.'); }
        await firmar('prealta', p, huellaPrealta(fresco));   // S-01: primero la firma (403 si no eres validador/gerencia), luego el sello; S-27: con la huella de lo firmado
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
        // C-60 (v0.56.0): relectura. Cerrar lo ya cerrado no pisa el sello de quien cerró primero.
        const fresco = await estado.cliente.renglon(estado.siteId, L.prealtas, p.id);
        if (fresco.Estado === 'cerrada') { aplicar('prealtas', p, fresco); trasCambioPrealta(); avisar(`Ya estaba cerrado: lo cerró ${quien(fresco.CerradaPor) || 'otra sesión'}.`, 'ojo'); return; }
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
function certificadosDe(e) {
    const lista = estado.certificados;
    if (indiceCertificados.lista !== lista || indiceCertificados.n !== lista.length) {
        const porEmbarque = new Map();
        for (const c of lista) { const k = Number(c.EmbarqueId); if (!porEmbarque.has(k)) porEmbarque.set(k, []); porEmbarque.get(k).push(c); }
        for (const v of porEmbarque.values()) v.sort((a, b) => b.id - a.id);
        indiceCertificados = { lista, n: lista.length, porEmbarque };
    }
    return indiceCertificados.porEmbarque.get(Number(e.id)) || [];
}
function certificadoVigente(e) { return certificadosDe(e).find(c => c.Estado === 'vigente') || null; }
/** Los del PROGRAMA: solo para leerlos en el detalle de la pre-alta. No se emite nada desde ahi. C-45: los de sus gondolas cargadas. */
function certificadosDePrograma(p) { return estado.certificados.filter(c => Number(c.PreAltaId) === Number(p.id)).sort((a, b) => a.id - b.id); }
/** S-11 sobre el certificado: vale solo con un renglon de PLANTA_Firmas de Tipo certificado, firmado por gerencia y por la misma cuenta que EmitidoPor. */
function certificadoFirmado(c) { return !!firmaDe('certificado', c.id, c.EmitidoPor); }
const motivoSinCertificados = () => (estado.certificadosError ? `No se pudo leer la lista de certificados: no se emite hasta que se vea.${estado.rol === 'gerencia' ? ` (PLANTA_Certificados: ${estado.certificadosError} — si no existe, se crea con herramientas-dev/provisionar.html)` : ' Avisa a gerencia.'}` : null);

/**
 * v0.39.0: el detalle del programa ya no emite nada — el certificado es de la gondola. Aqui solo se LEE cuantos
 * lleva el programa y cuanto suman, para que gerencia no tenga que recorrer Cerrados para saberlo.
 */
function pintarCertificadoEnDetalle(p) {
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
function abrirCertificadoDeEmbarque(e) {
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
    $('ctEstado').textContent = (cert
        ? `${cert.Title} · ${imprimible ? 'vigente y firmado' : cert.Estado + (cert.Estado === 'vigente' ? ' · SIN FIRMA' : '')}`
        : `Góndola ${e.Title || '(sin folio)'} · sin certificado${borrador ? ' · abajo, el BORRADOR: así saldría el papel' : ''}`)
        + ` · góndola ${e.Title || '—'}${e.Manifiesto ? ' · manifiesto ' + e.Manifiesto : ''}`
        + (certs.length > 1 ? ` · ${certs.length} emitidos para esta góndola` : '')
        + (vig && e.Etapa === 'anulado' ? ' · ⚠ LA GÓNDOLA ESTÁ ANULADA: cancela este certificado' : '')   // U-70 (v0.40.0)
        + (certs.filter(c => c.Estado === 'vigente').length > 1 ? ' · ⚠ HAY MÁS DE UNO VIGENTE: vale el más nuevo; cancela los otros' : '')   // C-39 (v0.40.0)
        + (vacios.length ? ` · ⚠ saldrían vacíos: ${vacios.join(', ')}` : '')   // U-72 (v0.45.0)
        + (puede && bloqueo ? ` · ${bloqueo}` : '');   // U-73 (v0.42.0): en el celular no hay title; el motivo se lee aquí
}
/** U-69 (v0.40.0): el panel de correccion de Sustituir, precargado con lo que la gondola y su programa dicen HOY. */
function mostrarCorreccion(si) {
    $('ctCorregir').classList.toggle('oculto', !si); $('ctBotones').classList.toggle('oculto', si);
    if (!si) return;
    const e = estado.certificadoEmbarque; const p = e && porId(estado.prealtas, e.PreAltaId);
    $('ctTicket').value = textoDe(e && e.TicketBascula); $('ctManifiesto').value = textoDe(e && e.Manifiesto);
    $('ctGenerador').value = textoDe(p && p.Generador); $('ctGeneradorDireccion').value = textoDe(p && p.GeneradorDireccion);
    $('ctGeneradorRegistro').value = textoDe(p && p.GeneradorRegistro); $('ctPozo').value = textoDe(p && p.Pozo); $('ctMotivo').value = '';
    $('ctTicket').focus();
}
/** Los campos de una cancelacion de certificado (C-50: uno solo para cancelar, compensar una emision y anular la gondola). */
const camposCancelacion = motivo => ({ Estado: 'cancelado', CanceladoPor: estado.cuenta.username, CanceladoEl: new Date().toISOString(), Motivo: String(motivo).slice(0, 255) });

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
async function emitirCertificado(sustituye = null, motivoSust = '', corr = null) {
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
        nuevo = await estado.cliente.crearRenglon(estado.siteId, L.certificados, limpiar({
            Title: folio, PreAltaId: p.id, EmbarqueId: e.id, Estado: 'vigente', ...papel,
            Sufijo: sufijoVerificacion(), EmitidoPor: estado.cuenta.username, EmitidoEl: ahora, Version: VERSION,
            Motivo: sustituye ? `Sustituye a ${sustituye.Title}: ${motivoSust}`.slice(0, 255) : null
        }));
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
function sustituirCertificado() {
    const e = estado.certificadoEmbarque; const vig = e && certificadoVigente(e); if (!vig || !PUEDE.emitirCertificado(estado.rol)) return;
    limpiarAvisos(); mostrarCorreccion(true);
}
async function confirmarSustitucion() {
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
async function cancelarCertificado() {
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
function imprimirCertificado() {
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
    $('pdkCarriers').textContent = String(estado.carriers.length); $('pdkUnidades').textContent = String(estado.unidades.length); $('pdkChoferes').textContent = String(estado.choferes.length);   // v0.66.0: la banda
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
            if (clave === 'carriers' && asistentePrealtaAbierto()) elegirCarrierEnPrealta(nuevo);
            pintarPadron();
        }
    } catch (e) { avisar('No se pudo guardar: ' + e.message, 'error'); } });
}
async function altaVigencia(titulo, titular, rol, fuente, vence, folio) {
    const v = await estado.cliente.crearRenglon(estado.siteId, L.vigencias, limpiar({ Title: titulo, Titular: titular, Rol: rol, Fuente: fuente, Vence: vence, Folio: folio, AvisoDias: CONFIG.avisoVigenciaDias, Activo: true }));
    estado.vigencias.push(v);
}

// ================================================================ HOY (la consola)

// Tanda 3 (v0.48.0): la Fila del dia se fundio en Gondolas; de sus piezas queda la etiqueta de compuerta (Hoy y Reportes).
function etiquetaCompuertaDe(e) {
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
    for (const p of ssf) pf.appendChild(renglon(`Firma · ${p.Title}`, prealtaCambioTrasFirma(p) ? `cambió después de que la firmó ${quien(p.FirmadaPor) || '?'} (carrier, unidades, choferes, corriente o generador) · la puerta no la ve · se vuelve a firmar desde su detalle` : `falta la firma del validador (trae el sello de ${quien(p.FirmadaPor) || '?'}, sin firma registrada) · la puerta no la ve · se firma desde su detalle`, 'Ver', () => verPrealta(vivo('prealtas', p))));   // U-28 / U-31
    for (const { p, sm } of dormidas) pf.appendChild(renglon(`Programa · ${p.Title}`, `${sm.motivo} · ¿se cierra? Sigue saliendo en la puerta`, 'Ver', () => verPrealta(vivo('prealtas', p))));
    for (const p of borradores) { const d = diasPara(p.FechaEstimada); pf.appendChild(renglon(`Pre-alta · ${p.Title}`, `firma del validador · 1er envío ${fechaCorta(p.FechaEstimada)}${d !== null ? ` (en ${d} días)` : ''} · capturó ${quien(p.CapturadaPor) || '?'}`, 'Ver', () => verPrealta(vivo('prealtas', p)))); }
    for (const e of pendientes) pf.appendChild(renglon(`Excepción · ${e.PlacaTractor}`, `${selloSinFirma(e)}autorización de gerencia · «${e.ExcepcionMotivo || 'sin motivo'}» · ${horaCorta(e.Arribo)}`, null, null, botonesExcepcion(e).map(b => ({ ...b, clase: b.accion === 'autorizar' ? '' : 'peligro' }))));
}

// C-29 (v0.34.0): UNA definición de «rechazo o excepción» y UN renglón para Hoy (los 10 últimos) y Reportes (todos). Antes el
// filtro y la lectura de CompuertaDetalle vivían copiados en los dos y la etiqueta difería (Hoy decía legal/comercial).
const esRechazo = e => e.Etapa !== 'anulado' && (e.Etapa === 'rechazado' || e.Compuerta === 'excepcion-comercial');
const rechazosYExcepciones = () => estado.embarques.filter(esRechazo).sort((a, b) => b.id - a.id);
/** C-54 (v0.51.0): la ÚNICA lectura de CompuertaDetalle — los hallazgos de las clases pedidas; detalle ilegible = ninguno. */
function hallazgosDe(e, ...clases) {
    try { return JSON.parse(e.CompuertaDetalle || '[]').filter(h => clases.includes(h.clase)); } catch (_) { return []; }
}
const reglasDe = (e, ...clases) => hallazgosDe(e, ...clases).map(h => h.regla).join(', ');
function renglonRechazo(e) {
    const causa = reglasDe(e, 'legal', 'comercial');
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
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = `CALYTEK_Embarques_${fechaCorta(estado.ventanaDesde).replace(/\//g, '-')}_a_${fechaMexico()}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    avisar(`CSV con ${plural(filas.length, 'embarque')} descargado.`, 'bien');
}
$('btnExportar').addEventListener('click', exportarCsv);
$('pdBusca').addEventListener('input', () => { estado.padronFicha = null; pintarPadron(); });
$('baBusca').addEventListener('input', pintarHistorial);
for (const b of document.querySelectorAll('#gTabs button, #gKpis button')) b.addEventListener('click', () => elegirVistaGondolas(b.dataset.g));
// v0.66.0: cada conteo del padrón abre su grupo y lo trae a la vista.
for (const b of $('pdKpis').querySelectorAll('button')) b.addEventListener('click', () => { const g = $(b.dataset.grupo); g.open = true; g.scrollIntoView({ block: 'start', behavior: 'smooth' }); });

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
$('puPrealta').addEventListener('change', () => { marcarOpcion($('puProgramas'), $('puPrealta').value); pintarChoferesPuerta(); pintarUnidadesPuerta(); pintarPrevioPuerta(); });
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
// U-10 (v0.23.0): «+ Alta de carrier» desde la pre-alta abre la forma del padron ENCIMA (dialog anidado: la pre-alta
// capturada se queda atras, intacta) y al guardar el carrier nuevo queda elegido aqui, con sus unidades/choferes (vacios).
// Antes habia que Cancelar (se perdia todo), ir a Padron, dar de alta y reteclear los bloques 1 y 2.
function elegirCarrierEnPrealta(c) {
    opciones($('paCarrier'), estado.carriers.filter(x => x.Activo !== false || x.id === c.id), x => x.id, x => x.Title);
    $('paCarrier').value = String(c.id);
    pintarUnidadesChoferesPrealta(); pintarAsistentePrealta();
}
$('btnAltaCarrierPrealta').addEventListener('click', () => abrirFormaPadron('carriers'));
$('btnGuardarPrealta').addEventListener('click', guardarPrealta);
$('btnCancelarPrealta').addEventListener('click', salirAsistentePrealta);
// Escape cierra el <dialog> sin pasar por Cancelar: la edicion pendiente del padron se suelta igual.
for (const clave of Object.keys(FORMA_PADRON)) $(FORMA_PADRON[clave].forma).addEventListener('close', () => { if (estado.padronEdit && estado.padronEdit.clave === clave) estado.padronEdit = null; });
// U-09 (v0.22.0): Escape sobre un formulario CON algo capturado pregunta antes de tirarlo — el <dialog> nativo cerraba sin
// pasar por Cancelar y la pre-alta de 14 campos se perdia. Vacio (o solo con los valores por omision) cierra directo; el
// boton Cancelar sigue cerrando sin preguntar. Un select cuenta solo si no esta en su primera opcion.
const hayCaptura = d => [...d.querySelectorAll('input:not([type=hidden]):not([type=checkbox]), textarea, select')].some(c => c.tagName === 'SELECT' ? c.selectedIndex > 0 : String(c.value).trim() !== '');
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
