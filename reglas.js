// Reglas de la puerta — funciones PURAS, sin red ni DOM, para poder probarlas con node.
//
// Aqui vive lo que el mapa control-planta-calytek decidio y la entrevista del 2026-09-04
// confirmo. Cada regla dice de donde sale. Cambiar un umbral es cambiar una constante; cambiar
// una CLASE (legal <-> comercial) es cambiar una decision, y eso se anota en docs/plan.md.
//
//   legal     -> rechazo duro, sin dispensa (padron seccion 1: la autorizacion ASEA del carrier,
//                la placa amparada, la corriente amparada; decision 4: la pre-alta sin firmar)
//   comercial -> excepcion por embarque, con motivo y autoridad gerencial (poliza, tarjeta,
//                licencia, CSF)
//   aviso     -> no bloquea: un vencimiento cercano, un dato que conviene mirar

export const CLASE = { LEGAL: 'legal', COMERCIAL: 'comercial', AVISO: 'aviso' };

/** Normaliza una placa para comparar: sin espacios, guiones ni minusculas. */
export function placaNormal(p) {
    return String(p || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/** 'a; b ;c' -> ['a','b','c'] */
export function lista(texto) {
    return String(texto || '').split(';').map(s => s.trim()).filter(Boolean);
}

/** Dias entre hoy y una fecha ISO (negativo si ya paso). null si no hay fecha. */
export function diasPara(fechaIso, hoy = new Date()) {
    if (!fechaIso) return null;
    const f = new Date(fechaIso);
    if (Number.isNaN(f.getTime())) return null;
    const dia = 24 * 60 * 60 * 1000;
    return Math.floor((f.getTime() - hoy.getTime()) / dia);
}

/**
 * Evalua una vigencia y devuelve un hallazgo o null.
 * @param {string} que        nombre legible del documento
 * @param {string} fechaIso   vencimiento
 * @param {'legal'|'comercial'} clase
 * @param {number} avisoDias  con cuantos dias de anticipacion se avisa
 */
export function evaluarVigencia(que, fechaIso, clase, avisoDias, hoy = new Date()) {
    const d = diasPara(fechaIso, hoy);
    if (d === null) return { clase, regla: que, ok: false, detalle: 'sin fecha de vigencia capturada' };
    const ddmm = `${fechaIso.slice(8, 10)}/${fechaIso.slice(5, 7)}/${fechaIso.slice(0, 4)}`;   // dd/mm/aaaa, el estándar de la casa
    if (d < 0) return { clase, regla: que, ok: false, detalle: `vencida hace ${plural(-d, 'día')} (${ddmm})` };
    if (d <= avisoDias) return { clase: CLASE.AVISO, regla: que, ok: true, detalle: `vence en ${plural(d, 'día')} (${ddmm})` };
    return null;
}

/**
 * LA COMPUERTA DOCUMENTAL (ticket 04 seccion 3). Corre ANTES de la bascula, sin tocar el residuo.
 *
 * @param {object} p
 * @param {object} p.prealta      renglon de PLANTA_PreAltas (fields) — puede ser null
 * @param {object} p.carrier      renglon de PLANTA_Carriers — puede ser null
 * @param {object} p.unidad       renglon de PLANTA_Unidades que caso con la placa — puede ser null
 * @param {object} p.chofer       renglon de PLANTA_Choferes que caso — puede ser null
 * @param {string} p.placaTractor lo que se leyo en la puerta
 * @param {string} p.placaPlana   lo que se leyo en la puerta
 * @param {string} p.corriente    la declarada en el manifiesto
 * @param {string} p.manifiesto   numero de manifiesto leido
 * @param {number} p.avisoDias
 * @param {Date}   [p.hoy]
 * @returns {{resultado:'pasa'|'rechazo-legal'|'excepcion-comercial', hallazgos:Array}}
 */
export function compuerta(p) {
    const hoy = p.hoy || new Date();
    const h = [];
    const legal = (regla, detalle) => h.push({ clase: CLASE.LEGAL, regla, ok: false, detalle });
    const comercial = (regla, detalle) => h.push({ clase: CLASE.COMERCIAL, regla, ok: false, detalle });
    const ok = (regla, detalle) => h.push({ clase: 'ok', regla, ok: true, detalle });
    const vig = (que, fecha, clase) => {
        const r = evaluarVigencia(que, fecha, clase, p.avisoDias, hoy);
        if (r) h.push(r); else ok(que, 'vigente');
    };

    // 0. Manifiesto: sin numero no hay embarque que registrar (c6).
    if (!String(p.manifiesto || '').trim()) legal('Manifiesto', 'sin número de manifiesto');

    // 1. Pre-alta firmada (decision 4). Sin firma no hay contra que cotejar: rechazo legal.
    if (!p.prealta) legal('Pre-alta', 'no hay pre-alta para este embarque');
    else if (p.prealta.Estado !== 'firmada') legal('Pre-alta', `la pre-alta está en "${p.prealta.Estado}", no firmada`);
    else ok('Pre-alta', `firmada por ${p.prealta.FirmadaPor || '?'}`);

    // 2. Carrier: autorizacion ASEA vigente (padron seccion 1, LEGAL).
    if (!p.carrier) legal('Carrier', 'el carrier no está en el padrón');
    else {
        if (!String(p.carrier.AutorizacionASEA || '').trim()) legal('Autorización ASEA del carrier', 'sin número de autorización');
        vig('Autorización ASEA del carrier', p.carrier.VigenciaASEA, CLASE.LEGAL);
        // 3. Corriente amparada (padron seccion 2, eje 3).
        const corrientes = lista(p.carrier.Corrientes);
        if (!p.corriente) legal('Corriente', 'no se declaró la corriente del manifiesto');
        else if (corrientes.length && !corrientes.includes(p.corriente)) {
            legal('Corriente', `el oficio del carrier no ampara "${p.corriente}" (ampara: ${corrientes.join(', ')})`);
        } else ok('Corriente', `${p.corriente} amparada`);
        if (p.prealta && p.prealta.CarrierId && p.carrier.id && Number(p.prealta.CarrierId) !== Number(p.carrier.id)) {
            legal('Carrier vs pre-alta', 'el carrier no es el de la pre-alta firmada');
        }
        vig('CSF del carrier', p.carrier.CSFVigencia, CLASE.COMERCIAL);
    }

    // 4. Unidad: placa amparada por el oficio (padron seccion 4, LEGAL). Dos placas, dos veces.
    if (!p.unidad) legal('Placa', `la placa ${placaNormal(p.placaTractor)} no está en el padrón: abrir el oficio, nunca darla de alta a mano`);
    else {
        // El folio de la unidad hereda el del carrier si va vacio (v0.19.9): la placa se transcribio de ese mismo oficio.
        const folio = String(p.unidad.FolioOficio || (p.carrier && p.carrier.FolioOficio) || '').trim();
        if (!folio) legal('Placa amparada', 'ni la unidad ni el carrier tienen folio de oficio que la ampare');
        else ok('Placa amparada', `oficio ${folio}`);
        if (p.placaPlana && placaNormal(p.unidad.PlacaPlana) !== placaNormal(p.placaPlana)) {
            legal('Placa de la plana', `se leyó ${placaNormal(p.placaPlana)} y el padrón tiene ${placaNormal(p.unidad.PlacaPlana) || '(vacía)'}`);
        }
        if (p.prealta && lista(p.prealta.UnidadesIds).length && !lista(p.prealta.UnidadesIds).includes(String(p.unidad.id))) {
            comercial('Unidad vs pre-alta', 'la unidad no venía en la pre-alta firmada');
        }
        vig('Tarjeta de circulación', p.unidad.TarjetaVigencia, CLASE.COMERCIAL);
        vig('Póliza de la unidad', p.unidad.PolizaVigencia, CLASE.COMERCIAL);
    }

    // 5. Chofer: COMERCIAL (no aparece en el oficio).
    if (!p.chofer) comercial('Chofer', 'el chofer no está en el padrón');
    else {
        vig('Licencia del chofer', p.chofer.LicenciaVigencia, CLASE.COMERCIAL);
        if (p.prealta && lista(p.prealta.ChoferesIds).length && !lista(p.prealta.ChoferesIds).includes(String(p.chofer.id))) {
            comercial('Chofer vs pre-alta', 'el chofer no venía en la pre-alta firmada');
        }
    }

    const hayLegal = h.some(x => x.clase === CLASE.LEGAL);
    const hayComercial = h.some(x => x.clase === CLASE.COMERCIAL);
    const resultado = hayLegal ? 'rechazo-legal' : hayComercial ? 'excepcion-comercial' : 'pasa';
    return { resultado, hallazgos: h };
}

/**
 * Folios (ticket 03 seccion 4): ciegos, secuenciales por anio, nunca reutilizados.
 *   E-26-00001  embarque que CALYTEK recibio y peso (nace en la primera pasada)
 *   R-26-0001   rechazo documental (no consume E-)
 *   L-26-001    campana (pre-alta), ticket 03: tres digitos
 * @param {'E'|'R'|'L'} tipo
 * @param {string[]} existentes  todos los folios ya emitidos (cualquier anio)
 * @param {Date} [hoy]
 */
export function siguienteFolio(tipo, existentes, hoy = new Date()) {
    const aa = String(hoy.getFullYear()).slice(-2);
    const ancho = { E: 5, R: 4, L: 3 }[tipo] || 4;
    const re = new RegExp(`^${tipo}-${aa}-(\\d+)$`);
    let max = 0;
    for (const f of existentes || []) {
        const m = re.exec(String(f || '').trim());
        if (m) max = Math.max(max, Number(m[1]));
    }
    return `${tipo}-${aa}-${String(max + 1).padStart(ancho, '0')}`;
}

/**
 * Tolerancia del neto (decision 6). No es barrera: devuelve el aviso y la app pide confirmar.
 * @returns {null|string}  motivo del aviso
 */
export function avisoNeto(brutoKg, taraKg, capacidadKg, tol) {
    const b = Number(brutoKg), t = Number(taraKg);
    if (!Number.isFinite(b) || !Number.isFinite(t)) return 'bruto o tara no son numeros';
    if (t <= 0 || b <= 0) return 'bruto y tara deben ser mayores que cero';
    if (b <= t) return `el bruto (${b} kg) no es mayor que la tara (${t} kg)`;
    const neto = b - t;
    const cap = Number(capacidadKg);
    if (Number.isFinite(cap) && cap > 0) {
        if (neto < cap * tol.minFraccion) return `neto ${neto} kg es menos del ${Math.round(tol.minFraccion * 100)} % de la capacidad de la unidad (${cap} kg)`;
        if (neto > cap * tol.maxFraccion) return `neto ${neto} kg supera el ${Math.round(tol.maxFraccion * 100)} % de la capacidad de la unidad (${cap} kg)`;
    }
    return null;
}

/**
 * Una pre-alta FIRMADA que ya no se mueve (Carlos, 2026-09-08: «¿siguen saliendo en la puerta si no llegan
 * los camiones?» — si, hasta cerrarla a mano; esto la senala). Devuelve null si no hay nada que decir, o
 * { motivo, desde } con el texto y la fecha del ultimo movimiento. Dos causas, la primera que aplique:
 *  - ya recibio todas las gondolas esperadas (esp > 0 y rec >= esp);
 *  - sin arribos en `dias` contados desde su ultimo movimiento (ultimo arribo, o si no hubo, la fecha estimada
 *    del primer envio o la firma, la mas reciente). Sin ninguna fecha no se puede juzgar: null.
 */
export function prealtaSinMovimiento(p, embarques, dias, hoy = new Date()) {
    if (!p || p.Estado !== 'firmada') return null;
    const suyos = (embarques || []).filter(e => Number(e.PreAltaId) === Number(p.id) && e.Etapa !== 'anulado' && e.Etapa !== 'rechazado');
    const esp = Number(p.GondolasEsperadas) || 0;
    if (esp && suyos.length >= esp) return { motivo: `ya recibió sus ${esp} góndola(s)`, desde: null };
    const fechas = suyos.map(e => e.Arribo).concat(suyos.length ? [] : [p.FechaEstimada, p.FirmadaEl]).filter(Boolean).map(f => new Date(f)).filter(f => !Number.isNaN(f.getTime()));
    if (!fechas.length) return null;
    const ultimo = new Date(Math.max(...fechas.map(f => f.getTime())));
    const d = Math.floor((hoy.getTime() - ultimo.getTime()) / 86400000);
    if (d < dias) return null;
    return { motivo: suyos.length ? `sin arribos desde hace ${d} días` : `firmada y sin un solo arribo en ${d} días`, desde: ultimo.toISOString() };
}

/** Fecha YYYY-MM-DD en hora de Mexico (contrato de nombres: nunca UTC). */
export function fechaMexico(ahora = new Date()) {
    const partes = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'America/Mexico_City', year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(ahora);
    const v = t => partes.find(x => x.type === t).value;
    return `${v('year')}-${v('month')}-${v('day')}`;
}

/**
 * Hora de Mexico como texto: 'fecha' → dd/mm/aaaa HH:MM (pantalla y ticket), 'hora' → HH:MM (fila del dia),
 * 'completa' → con segundos (CSV). Un solo helper para los cuatro sitios (C-10, v0.22.0) y con hourCycle h23,
 * nunca hour12:false: en Chromium hour12:false cae en h24 para varias locales e imprime «24:05» a medianoche —
 * justo la gondola que cierra 00:10 (F1). Vacio → '—'; lo que no es fecha se devuelve tal cual.
 */
export function horaMexico(iso, modo = 'fecha') {
    if (!iso) return '—';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return String(iso);
    const base = { timeZone: 'America/Mexico_City', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' };
    if (modo === 'hora') return d.toLocaleTimeString('es-MX', base);
    const fecha = { day: '2-digit', month: '2-digit', year: 'numeric' };
    return d.toLocaleString('es-MX', modo === 'completa' ? { ...base, ...fecha, second: '2-digit' } : { ...base, ...fecha });
}

/** Slug del contrato de nombres (pasos c01-c08 de contrato-nombres-captura.md). */
export function slug(texto) {
    let s = String(texto || '').trim().toLowerCase();
    s = s.normalize('NFD').replace(/[̀-ͯ]/g, '');
    // c03-c04 y todo lo demas que no sea letra, digito, espacio o guion (el punto medio, las
    // comillas curvas): lo que no cabe en un nombre de carpeta no cabe en el slug.
    s = s.replace(/[^a-z0-9\s-]/g, '');
    s = s.replace(/\s+/g, '-').replace(/-+/g, '-');
    s = s.slice(0, 50).replace(/-$/, '');
    return s || 'sin-concepto';
}

/** Rol efectivo de un correo segun PLANTA_Roles. Sin renglon: 'lectura'. */
export function rolDe(correo, roles) {
    const c = String(correo || '').trim().toLowerCase();
    const r = (roles || []).find(x => String(x.Title || '').trim().toLowerCase() === c && x.Activo !== false);
    return r ? r.Rol : 'lectura';
}

export const PUEDE = {
    capturarPrealta: rol => ['trazabilidad', 'gerencia'].includes(rol),
    firmarPrealta: rol => ['validador', 'gerencia'].includes(rol),
    puerta: rol => ['trazabilidad', 'gerencia'].includes(rol),
    autorizarExcepcion: rol => rol === 'gerencia',
    // Deshacer una captura equivocada (pedido de Carlos, 2026-09-05). Quien captura puede corregir
    // lo suyo en el momento: la gondola esta esperando. La traza queda en el renglon y en el
    // historial de versiones de SharePoint.
    corregir: rol => ['trazabilidad', 'gerencia'].includes(rol),
    ver: () => true
};

/**
 * Que se puede hacer con un embarque equivocado (2026-09-05):
 *   'eliminar' — sin folio todavia (etapa compuerta: paso o espera autorizacion). No hay nada que
 *                conservar: ni folio ni foto. Se borra el renglon.
 *   'anular'   — ya tiene folio (R- del rechazo, o E- del bruto en adelante). El folio NUNCA se
 *                reutiliza (ticket 03 §4) y la foto ya esta en el buzon: el renglon se queda con
 *                etapa 'anulado', motivo obligatorio y quien lo anulo.
 *   null       — ya anulado.
 */
export function accionCorreccion(embarque) {
    if (!embarque || embarque.Etapa === 'anulado') return null;
    return embarque.Etapa === 'compuerta' && !embarque.Title ? 'eliminar' : 'anular';
}

// ---------------------------------------------------------------- C-25 (v0.28.0): la frontera de fechas y utilerias puras
// Vivian en app.js, que no se importa desde node: la E2E solo pegaba ISO y nadie probaba «16/03/26», «31/04/2026» ni el
// mensaje de error. Fechas: el estandar de la casa es dd/mm/aaaa (Carlos, 2026-09-05), en pantalla, en el ticket y al
// capturar; lo guardado en SharePoint sigue siendo ISO.
export function fechaCorta(iso) {
    if (!iso) return '—';
    const s = String(iso);
    return /^\d{4}-\d{2}-\d{2}/.test(s) ? `${s.slice(8, 10)}/${s.slice(5, 7)}/${s.slice(0, 4)}` : s;
}
/** Acepta dd/mm/aaaa (lo que teclea la gente) y aaaa-mm-dd (pegados y pruebas). Vacio = null; cualquier otra cosa lanza. */
export function aIsoDia(texto) {
    const s = String(texto || '').trim();
    if (!s) return null;
    let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/);
    let y, mo, d;
    // Ano de dos cifras = 20aa: en la puerta se teclea «16/03/26» (fotos de Carlos, 2026-09-06) y ninguna vigencia es del siglo pasado.
    if (m) { d = +m[1]; mo = +m[2]; y = m[3].length === 2 ? 2000 + +m[3] : +m[3]; }
    else if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else throw new Error(`Fecha «${s}» no válida: escríbela como dd/mm/aaaa`);
    const f = new Date(y, mo - 1, d, 12);
    if (f.getFullYear() !== y || f.getMonth() !== mo - 1 || f.getDate() !== d) throw new Error(`Fecha «${s}» no existe: escríbela como dd/mm/aaaa`);
    return f.toISOString();
}
/** Al teclear una fecha: solo digitos y las barras se ponen solas (05092026 -> 05/09/2026); un ISO pegado se muestra dd/mm/aaaa. */
export function autoformatoFecha(valor) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(valor)) return fechaCorta(valor);
    const dig = valor.replace(/\D/g, '').slice(0, 8);
    return dig.length > 4 ? `${dig.slice(0, 2)}/${dig.slice(2, 4)}/${dig.slice(4)}` : dig.length > 2 ? `${dig.slice(0, 2)}/${dig.slice(2)}` : dig;
}
/** «1 gondola» / «3 gondolas» (U-36): sin «(s)». El plural se pasa solo cuando no es singular + «s». */
export const plural = (n, uno, varios = uno + 's') => `${n} ${n === 1 ? uno : varios}`;
/** Para el POST: lo vacio no viaja. */
export function limpiar(obj) { const o = {}; for (const k in obj) if (obj[k] !== null && obj[k] !== undefined && obj[k] !== '') o[k] = obj[k]; return o; }
/** Para el PATCH: lo vacio va como null para que SharePoint lo borre; `limpiar()` lo omitiria y el dato viejo sobreviviria. */
export function paraPatch(campos) { const o = {}; for (const k in campos) o[k] = campos[k] === '' || campos[k] === undefined ? null : campos[k]; return o; }
