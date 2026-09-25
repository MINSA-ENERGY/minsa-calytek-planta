// CALYTEK Planta — estado, utilerias de DOM, avisos, sesion y lecturas compartidas. No importa ninguna pantalla: se evalua primero.
// Salió de app.js en C-76 (v0.75.0): código movido tal cual; solo se agregaron import/export.

import { CONFIG } from './config.js';
import { crearCliente } from './graph.js';
import { autoformatoFecha, compuerta, firmaAmparaPrealta, horaMexico, limpiar, lista, palabraCompuerta, PUEDE, rolDe } from './reglas.js';

export const VERSION = '0.75.2';   // la misma cadena va en package.json y en sw.js (CACHE); test/version.test.js lo exige
export const $ = id => document.getElementById(id);
export const L = CONFIG.listas;

export const estado = {
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
    padronVista: { v: 'lista', tab: 'vigentes', sub: 'unidades', carrier: null, ficha: null, desde: null },   // v0.72.0: vista del padron por carrier
    padronCarrier: null,     // id del carrier del expediente o la ficha del padron (preselecciona el alta, U-46)
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
export const pca = new msal.PublicClientApplication({
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
export function avisar(texto, clase = '') {
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
export function limpiarAvisos() { $('avisos').textContent = ''; $('entradaAviso').textContent = ''; $('entradaAviso').className = 'mensaje oculto'; for (const z of document.querySelectorAll('.dlg-avisos')) z.textContent = ''; }   // U-39: incluye la del veredicto
// Los formularios de alta viven en <dialog> (v0.19.8): abrir es showModal, cerrar es close. Idempotentes.
export function abrirForma(id) { const d = $(id); limpiarAvisos(); if (!d.open) d.showModal(); d.scrollTo({ top: 0 }); d.dataset.huella = huellaForma(d); }
/** Valores de la forma en una cadena (U-34, v0.26.0): se toma al abrir y Escape pregunta solo si cambió — un Editar sin tocar nada cierra directo. */
export const huellaForma = d => JSON.stringify([...d.querySelectorAll('input:not([type=hidden]), textarea, select')].map(c => c.type === 'checkbox' || c.type === 'radio' ? c.checked : String(c.value)));
export function cerrarForma(id) { const d = $(id); if (d.open) d.close(); }

export function el(tag, clase, texto) {
    const e = document.createElement(tag);
    if (clase) e.className = clase;
    if (texto !== undefined) e.textContent = texto;
    return e;
}
/**
 * Renglon de lista con un boton principal y, opcionalmente, acciones secundarias
 * ({texto, alClic, accion, clase}) apiladas a la derecha; `accion` sale como data-accion.
 */
export function renglon(titulo, sub, boton, alClic, extras = []) {
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
export function botonAccion({ texto, alClic, accion, clase, deshabilitado }) {
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
export function confirmar({ titulo, texto, ok = 'Confirmar', motivo = false, etiquetaMotivo = 'Motivo', peligro = false }) {
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
export function etiqueta(texto, clase) { return el('span', 'etiqueta ' + (clase || ''), texto); }
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
export function opciones(sel, items, valor, textoDe, primera = '— elige —') {
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
export const horaCorta = iso => horaMexico(iso, 'fecha');
for (const inp of document.querySelectorAll('input.fecha')) inp.addEventListener('input', () => { inp.value = autoformatoFecha(inp.value); });
export function porId(coleccion, id) { return coleccion.find(x => x.id === Number(id)) || null; }
/**
 * C-12 (v0.25.0): el objeto VIVO de estado[clave] con el id de `x`, resuelto AL CLIC. Los handlers de renglon capturan el
 * objeto de la lista con que se pinto; tras un refresco con captura a medias (que no repinta) ese objeto ya no esta en
 * estado.* y la tara / la anulacion / la firma se escribian sobre una copia rancia. Si ya no existe, se conserva el viejo
 * y la operacion lo reporta (el bruto y la tara releen el renglon antes de escribir).
 */
export function vivo(clave, x) { return (x && porId(estado[clave], x.id)) || x; }
/**
 * C-23 (v0.28.0): un Object.assign DESPUES de un await cae sobre el objeto vivo por id, ademas del que el handler tenia en la
 * mano (C-12 resolvia al clic; el motivo de la anulacion se teclea 10-30 s y el refresco podia sustituir la lista debajo).
 * Devuelve el vivo.
 */
export function aplicar(clave, obj, campos) { const v = vivo(clave, obj); Object.assign(obj, campos); if (v !== obj) Object.assign(v, campos); return v; }
/** C-23 / C-26: `x` ocupa su lugar en estado[clave] por id (sin duplicar el renglon si un refresco ya lo trajo); si no estaba, entra. */
export function anclar(clave, x) { const i = estado[clave].findIndex(y => y.id === x.id); if (i >= 0) estado[clave][i] = x; else estado[clave].push(x); return x; }
/**
 * C-24 / C-23 (v0.28.0): toda escritura al tenant pasa por aqui. Deshabilita el boton que la disparo mientras dura (un segundo
 * toque con guante y senal lenta creaba dos carriers, dos firmas) y cuenta las escrituras en vuelo, que recargar() respeta.
 * `btn` es un id, un elemento o null (escrituras que no nacen de un boton fijo).
 */
export let escrituras = 0;
export async function escribiendo(btn, fn) {
    const b = typeof btn === 'string' ? $(btn) : btn;
    if (b && b.disabled) return undefined;
    if (b) b.disabled = true;
    escrituras++;
    try { return await fn(); }
    finally { escrituras--; if (b) b.disabled = false; }
}
/** C-26: un solo formateador para llenar formas (null/undefined -> ''). */
export const textoDe = v => v === null || v === undefined ? '' : String(v);
/** C-26: el filtro del buscador (padron y cerrados): sin texto pega todo. */
export const filtroTexto = q => (...campos) => !q || campos.some(v => normaliza(v).includes(q));
const haySel = sel => sel !== null && sel !== undefined;
export function nombreDe(coleccion, id) { if (id === null || id === undefined || id === '') return '—'; const x = porId(coleccion, id); return x ? x.Title : `#${id}`; }

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
export async function prepararMsal() {
    if (msalListo) return null;
    await pca.initialize();
    const respuesta = await pca.handleRedirectPromise();
    msalListo = true;
    return respuesta;
}
// U-20 (v0.22.0): el progreso del inicio de sesion se escribe en el PROPIO boton (con anillo girando), no solo en la
// linea mono de 9.5 px al pie: con la senal lenta de la caseta parecia que no habia pasado nada y se volvia a tocar.
export function pasoEntrada(texto) {
    const b = $('btnEntrar');
    b.disabled = !!texto; b.classList.toggle('ocupado', !!texto);
    b.textContent = texto || 'Entrar con mi cuenta de MINSA';
    $('textoEntrar').textContent = texto || 'CALYTEK · Planta';
    $('textoEntrar').classList.toggle('estado', !!texto);   // v0.45.3: como proyectos (U-05), un ESTADO se lee; la marca en reposo, chica
}
export async function salir() {
    // Salir esta a un toque en la barra del celular: se confirma para no cerrar la sesion con la gondola esperando (F5).
    const { ok } = await confirmar({ titulo: 'Salir de la app', ok: 'Salir', texto: 'Se cierra la sesión de MINSA en este dispositivo. Lo capturado ya está guardado; lo que esté a medias en pantalla se pierde.' });
    if (!ok) return;
    try { await pca.logoutRedirect({ account: estado.cuenta }); }
    catch (_) { sessionStorage.clear(); window.location.reload(); }
}
export async function refrescarCliente() {
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
export function ponerQuien(correo, rol = '') {
    const nombre = quien(correo);
    for (const q of document.querySelectorAll('.quien')) { q.textContent = ''; q.appendChild(el('b', '', nombre)); if (rol) q.appendChild(el('span', 'rol', rol)); q.title = correo; }   // v0.32.0: dos renglones, como el rail de Proyectos
    $('rolMovil').textContent = rol;
    $('quienMovil').textContent = nombre; $('quienMovil').title = correo;
    $('correoMovil').textContent = nombre === correo ? '' : correo;
}
// I5 (7-sep): cuando se leyeron las listas por ultima vez, en la barra movil y en el rail. Ambar pasados 5 minutos.
let avisoReintento = false;
export function pintarSync(leyendo = false, texto = 'Leyendo las listas…') {
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


/** Un embarque esta "en planta" si paso la compuerta (o le autorizaron la excepcion) y no ha cerrado: lo que se PESA.
 *  Lo usan los botones de pesar, el Historial (lo que ya no esta en planta) y el KPI de Reportes. */
export function enPlanta(e) {
    return (e.Etapa === 'compuerta' && (e.Compuerta === 'pasa' || excepcionAutorizada(e)))
        || e.Etapa === 'bruto';
}
export const esperaAutorizacion = e => e.Etapa === 'compuerta' && e.Compuerta === 'excepcion-comercial' && !excepcionAutorizada(e);
export function excepcionesPendientes() { return estado.embarques.filter(esperaAutorizacion); }
/** Tanda 3 (decision 10): la lista «En planta» de Góndolas = lo que se pesa + lo que espera la autorizacion de gerencia
 *  (en ambar), por hora de llegada. Es LA definicion del numero del rail, de la pestaña y del subtitulo de Hoy. */
export const enListaPlanta = () => estado.embarques.filter(e => enPlanta(e) || esperaAutorizacion(e)).sort((a, b) => String(a.Arribo || '').localeCompare(String(b.Arribo || '')));

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
export function esColumnaFaltante(e) { return !!e && e.status === 400; }
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
export function firmaDe(tipo, id, sello) {
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
export function prealtaFirmada(p) { return p.Estado === 'firmada' && firmasPrealta(p).some(f => firmaAmparaPrealta(f.Motivo, p, f.Created)); }   // S-29: Created lo pone SharePoint
/** Firmada con sello y firma, pero el renglón cambió después: se dice así, no «sin firma». */
export const prealtaCambioTrasFirma = p => p.Estado === 'firmada' && !prealtaFirmada(p) && firmasPrealta(p).length > 0;
export function excepcionAutorizada(e) { return !!e.ExcepcionAutorizo && !!firmaDe('excepcion', e.id, e.ExcepcionAutorizo); }
/** C-55 (v0.52.0): la compuerta de un embarque en palabras, con la firma y no el sello decidiendo la excepción. */
export const palabraCompuertaDe = e => palabraCompuerta(e.Compuerta, excepcionAutorizada(e));
/** Sellos sin firma que valga (de antes del corte, por fuera de la app, o firmada por quien no tiene el rol): la compuerta no los acepta. */
export function sellosSinFirma() {
    return { prealtas: estado.prealtas.filter(p => p.Estado === 'firmada' && !prealtaFirmada(p)),
             embarques: estado.embarques.filter(e => e.Etapa === 'compuerta' && !!e.ExcepcionAutorizo && !firmaDe('excepcion', e.id, e.ExcepcionAutorizo)) };
}
export const selloSinFirma = e => (e.ExcepcionAutorizo ? `sello de ${quien(e.ExcepcionAutorizo)} sin firma registrada · ` : '');   // U-28 / U-31 (v0.26.0)
/** S-07: por que la app no puede firmar ni autorizar ahora mismo, o null. Va como `title` del boton deshabilitado. */
export const motivoSinFirmas = () => (estado.firmasError ? `No se pudo leer el registro de firmas: no se firma ni se autoriza hasta que se vea. Actualiza; si sigue, avisa a gerencia.${estado.rol === 'gerencia' ? ` (PLANTA_Firmas: ${estado.firmasError})` : ''}` : null);   // U-31: lenguaje de planta primero; el detalle técnico solo a gerencia
/** El renglon de firma se escribe PRIMERO: para quien no esta en el grupo de firmantes es un 403, y ahi termina. */
export async function firmar(tipo, objeto, motivo) {
    if (estado.firmasError) throw new Error(motivoSinFirmas());
    const f = await estado.cliente.crearRenglon(estado.siteId, L.firmas, limpiar({
        Title: `${tipo} · ${objeto.Title || objeto.PlacaTractor || objeto.id}`, Tipo: tipo, ObjetoId: objeto.id,
        Firmante: estado.cuenta.username, FirmadoEl: new Date().toISOString(), Motivo: motivo || null }));
    estado.firmas.push(f);
    return f;
}
export function pintarInsignias() {
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
export const iso = d => new Date(d).toISOString();

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
export async function embarquesDelAno(avisar) {
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
export function embarquesDeProgramas() {
    if (!estado.fueraDeVentana.size) return estado.embarques;
    const porId = new Map(estado.embarques.map(e => [e.id, e]));
    for (const filas of estado.fueraDeVentana.values()) for (const e of filas) if (!porId.has(e.id)) porId.set(e.id, e);
    return [...porId.values()];
}

/** Mete en la lista de la ventana los renglones frescos que le correspondan (por id). */
export function fundirEnVentana(frescos) {
    const porId = new Map(estado.embarques.map(e => [e.id, e]));
    for (const f of frescos) {
        const v = porId.get(f.id);
        if (v) Object.assign(v, f);
        else if (f.Arribo && f.Arribo >= estado.ventanaDesde) estado.embarques.push(f);
    }
}

export async function cargarTodo() {
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
export function reanclar() {
    const fresco = (col, x) => (x && porId(col, x.id)) || x;
    if (estado.pesando) estado.pesando.embarque = fresco(estado.embarques, estado.pesando.embarque);
    estado.prealtaAbierta = fresco(estado.prealtas, estado.prealtaAbierta);
    estado.prealtaEdit = fresco(estado.prealtas, estado.prealtaEdit);
    if (estado.padronEdit) estado.padronEdit.x = fresco(estado[estado.padronEdit.clave], estado.padronEdit.x);
    estado.certificadoEmbarque = fresco(estado.embarques, estado.certificadoEmbarque);   // C-38 (v0.40.0)
}


export const hayCaptura = d => [...d.querySelectorAll('input:not([type=hidden]):not([type=checkbox]), textarea, select')].some(c => c.tagName === 'SELECT' ? c.selectedIndex > 0 : String(c.value).trim() !== '');

/**
 * Nombre y apellido de una cuenta (Carlos, 8-sep): el renglon guarda el CORREO (auditable, es lo que
 * SharePoint conoce); en pantalla se muestra el `Nombre` de PLANTA_Roles, y si esa fila no lo trae,
 * el nombre que da Entra para la cuenta activa, y al final el correo tal cual.
 */
export function quien(correo) {
    if (!correo) return '';
    const r = estado.roles.find(x => String(x.Title || '').toLowerCase() === String(correo).toLowerCase());
    if (r && r.Nombre) return r.Nombre;
    if (estado.cuenta && estado.cuenta.username && estado.cuenta.username.toLowerCase() === String(correo).toLowerCase() && estado.cuenta.name) return estado.cuenta.name;
    return correo;
}

export const asistentePrealtaAbierto = () => !$('paAsis').classList.contains('oculto');

export const normaliza = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

export const asistentePadronAbierto = () => estado.padronVista.v === 'asis' && !!estado.padronVista.asis;
