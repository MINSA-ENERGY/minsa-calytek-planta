// Cliente de Microsoft Graph para listas y evidencia — adaptado de minsa-captura-app/app/subir.js.
//
// Todo pasa por conReintento: en un celular en la caseta, un fallo de red o un 429 es el caso
// normal. Lo que NO se reintenta es un 403 o un 404: esos no mejoran repitiendo.

const REINTENTOS = 4;
const dormir = ms => new Promise(res => setTimeout(res, ms));

function valeReintentar(estado) {
    return estado === 429 || estado === 503 || estado === 504 || estado === 0;
}

export async function conReintento(hacer, alAvisar) {
    let espera = 800;
    for (let intento = 1; intento <= REINTENTOS; intento++) {
        let r;
        try {
            r = await hacer();
        } catch (e) {
            if (intento === REINTENTOS) throw e;
            if (alAvisar) alAvisar(`sin conexión, reintentando (${intento}/${REINTENTOS - 1})`);
            await dormir(espera); espera *= 2;
            continue;
        }
        if (r.ok) return r;
        if (!valeReintentar(r.status) || intento === REINTENTOS) return r;
        const dice = Number(r.headers.get('Retry-After'));
        const cuanto = Number.isFinite(dice) && dice > 0 ? dice * 1000 : espera;
        if (alAvisar) alAvisar(`el servidor pidió esperar, reintentando (${intento}/${REINTENTOS - 1})`);
        await dormir(cuanto);
        espera *= 2;
    }
    throw new Error('se agotaron los reintentos');
}

/** Cuerpo de error de Graph, leido UNA vez: codigo (invalidRequest, itemNotFound, accessDenied…) y mensaje. */
async function cuerpoDeError(r) {
    try { const j = await r.json(); return { codigo: j?.error?.code || '', mensaje: j?.error?.message || '' }; }
    catch (_) { return { codigo: '', mensaje: '' }; }   // no era JSON
}
function frase(status, d) {
    const detalle = d.mensaje || d.codigo;
    if (status === 403) return `sin permiso (403). ${detalle}`;
    if (status === 404) return `no existe (404). ${detalle}`;
    if (status === 401) return 'la sesión caducó (401). Vuelve a entrar.';
    return `HTTP ${status}. ${detalle}`;
}

/** Mensaje util a partir de una respuesta fallida. No nombra la operacion: la pone quien llama. */
export async function motivo(r) { return frase(r.status, await cuerpoDeError(r)); }

/**
 * C-15 (v0.27.0): el Error que sale del cliente trae `status` (HTTP) y `codigo` (error.code de Graph) ademas del texto.
 * app.js decide por ellos —esColumnaFaltante, cargarFirmas—, no por regex sobre el mensaje: hasta v0.26.0 «invalid»
 * casaba con InvalidAuthenticationToken y una sesion caducada se leia como «la lista no tiene la columna».
 */
export async function fallo(r, prefijo) {
    const d = await cuerpoDeError(r);
    return Object.assign(new Error(prefijo + frase(r.status, d)), { status: r.status, codigo: d.codigo });
}

/** Codifica una ruta para Graph SIN destruir las diagonales. */
export function rutaUrl(ruta) {
    return String(ruta).split('/').filter(s => s !== '').map(encodeURIComponent).join('/');
}

/**
 * Convierte un item de Graph (con .fields) en un objeto plano con `id` numerico.
 * Las columnas se leen por su nombre INTERNO (esquema.json).
 */
export function aplanar(item) {
    const f = item && item.fields ? item.fields : (item || {});
    // Tanda 5 (v0.50.0, decisión 11): quién tocó el renglón por última vez, sin columna nueva — SharePoint ya lo sabe.
    // Lo lee el aviso de «gana el primero» cuando otra sesión avanzó la góndola; el guion bajo lo aparta de las columnas.
    const por = item && item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName;
    return { ...f, id: Number(item.id ?? f.id), ...(por ? { _por: por } : {}) };
}

export function crearCliente(graph, token) {
    const cab = { Authorization: 'Bearer ' + token };
    const json = { 'Content-Type': 'application/json' };
    const listasPorNombre = new Map();

    async function pedir(url, opciones = {}, avisar) {
        return conReintento(() => fetch(url, {
            ...opciones,
            headers: { ...cab, ...(opciones.headers || {}) }
        }), avisar);
    }

    return {
        async sitio(host, ruta) {
            const r = await pedir(`${graph}/sites/${host}:${ruta}`);
            if (!r.ok) throw await fallo(r, `no se pudo abrir el sitio ${ruta}: `);
            return (await r.json()).id;
        },

        /** Lista los nombres de listas del sitio (para provisionar y para resolver ids). */
        async listas(siteId) {
            // S-07 (v0.25.0): sigue @odata.nextLink. Antes se quedaba con la primera pagina de 200: en un sitio con mas
            // listas, PLANTA_Firmas podia caer en la segunda y la app la daba por inexistente.
            let url = `${graph}/sites/${siteId}/lists?$select=id,name,displayName&$top=200`;
            const v = [];
            while (url) {
                const r = await pedir(url);
                if (!r.ok) throw await fallo(r, 'no se pudieron ver las listas del sitio: ');
                const j = await r.json();
                v.push(...j.value);
                url = j['@odata.nextLink'] || null;
            }
            for (const l of v) { listasPorNombre.set(l.displayName, l.id); listasPorNombre.set(l.name, l.id); }
            return v;
        },

        async idDeLista(siteId, nombre) {
            if (!listasPorNombre.has(nombre)) await this.listas(siteId);
            const id = listasPorNombre.get(nombre);
            // C-15: excepcion tipada como el 404 de Graph, para que quien llama no dependa del texto.
            if (!id) throw Object.assign(new Error(`no existe la lista ${nombre} en el sitio: hay que provisionarla (herramientas-dev/provisionar.html)`), { status: 404, codigo: 'listaNoExiste' });
            return id;
        },

        /** Columnas reales de una lista (nombre interno + tipo), para cotejar contra esquema.json. */
        async columnas(siteId, listaId) {
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/columns?$top=200`);
            if (!r.ok) throw await fallo(r, 'no se pudieron leer las columnas: ');
            return (await r.json()).value;
        },

        async crearLista(siteId, cuerpo) {
            const r = await pedir(`${graph}/sites/${siteId}/lists`, {
                method: 'POST', headers: json, body: JSON.stringify(cuerpo)
            });
            if (!r.ok) throw await fallo(r, `no se pudo crear la lista ${cuerpo.displayName}: `);
            return await r.json();
        },

        async agregarColumna(siteId, listaId, columna) {
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/columns`, {
                method: 'POST', headers: json, body: JSON.stringify(columna)
            });
            if (!r.ok) throw await fallo(r, `no se pudo crear la columna ${columna.name}: `);
            return await r.json();
        },

        /**
         * Todos los renglones de una lista, aplanados. $top=500 y sigue @odata.nextLink: a
         * 8 gondolas/dia son ~2,000 embarques/ano, asi que hay que paginar de verdad.
         * @param {string[]} [campos]  v0.76.0: solo esas columnas y el id (sin createdBy/lastModifiedBy/eTag...): para la
         *                            lista que crece sin tope y se lee entera (PLANTA_Firmas), la mitad del peso por pagina.
         * @param {string} [filtro]  OData, p. ej. "fields/Estado eq 'firmada'" (necesita columna indexada o la
         *                            cabecera Prefer: HonorNonIndexedQueriesWarningMayFailRandomly)
         */
        async renglones(siteId, nombreLista, filtro, avisar, campos) {
            const listaId = await this.idDeLista(siteId, nombreLista);
            let url = `${graph}/sites/${siteId}/lists/${listaId}/items?`
                + (campos ? `$select=id&expand=fields($select=${campos.join(',')})` : 'expand=fields') + '&$top=500'
                + (filtro ? `&$filter=${encodeURIComponent(filtro)}` : '');
            const todos = [];
            while (url) {
                const r = await pedir(url, { headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } }, avisar);
                if (!r.ok) throw await fallo(r, `no se pudieron leer los renglones de ${nombreLista}: `);
                const j = await r.json();
                for (const it of j.value) todos.push(aplanar(it));
                url = j['@odata.nextLink'] || null;
            }
            return todos;
        },

        /** UN renglon por id, aplanado (C-13, v0.25.0): para releer el estado vigente antes de un PATCH que cierra. */
        async renglon(siteId, nombreLista, id, avisar) {
            const listaId = await this.idDeLista(siteId, nombreLista);
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`, {}, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo leer el renglón ${id} de ${nombreLista}: `);
            return aplanar(await r.json());
        },

        async crearRenglon(siteId, nombreLista, campos, avisar) {
            const listaId = await this.idDeLista(siteId, nombreLista);
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/items`, {
                method: 'POST', headers: json, body: JSON.stringify({ fields: campos })
            }, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo escribir en ${nombreLista}: `);
            return aplanar(await r.json());
        },

        async actualizarRenglon(siteId, nombreLista, id, campos, avisar) {
            const listaId = await this.idDeLista(siteId, nombreLista);
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/items/${id}/fields`, {
                method: 'PATCH', headers: json, body: JSON.stringify(campos)
            }, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo actualizar el renglón ${id} de ${nombreLista}: `);
            return await r.json();
        },

        /** Borra un renglon. Solo para lo que NUNCA tuvo folio ni firma (app.js decide); lo demas se ANULA, no se borra. */
        async borrarRenglon(siteId, nombreLista, id, avisar) {
            const listaId = await this.idDeLista(siteId, nombreLista);
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/items/${id}`, { method: 'DELETE' }, avisar);
            if (!r.ok && r.status !== 404) throw await fallo(r, `no se pudo borrar el renglón ${id} de ${nombreLista}: `);
        },

        /** Cambia las opciones (u otro atributo) de una columna existente: PATCH sobre la definicion. */
        async actualizarColumna(siteId, listaId, columnaId, cuerpo, avisar) {
            const r = await pedir(`${graph}/sites/${siteId}/lists/${listaId}/columns/${columnaId}`, {
                method: 'PATCH', headers: json, body: JSON.stringify(cuerpo)
            }, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo actualizar la columna ${columnaId}: `);
            return await r.json();
        },

        /** v0.33.0 (Archivos): la biblioteca del sitio; su webUrl es el destino de «Ver en SharePoint» (la URL no se deduce). */
        async biblioteca(siteId) {
            const r = await pedir(`${graph}/sites/${siteId}/drive?$select=id,webUrl`);
            if (!r.ok) throw await fallo(r, 'no se pudo abrir la biblioteca: ');
            return await r.json();
        },

        /**
         * v0.33.0 (Archivos): los hijos de una carpeta de la biblioteca por su ruta, paginando ($top=200 + nextLink). Devuelve
         * null si la carpeta NO existe (404): el arbol lo dice en vez de fallar — 02_Planta/Bascula nace con el primer lote
         * que /archivar-calytek acomode. Misma ruta drive/root: con la que la app sube la evidencia (Sites.Selected alcanza).
         */
        async hijos(siteId, ruta, avisar) {
            let url = `${graph}/sites/${siteId}/drive/root:/${rutaUrl(ruta)}:/children?$select=id,name,size,folder,file,lastModifiedDateTime,webUrl&$top=200`;
            const todos = [];
            while (url) {
                const r = await pedir(url, {}, avisar);
                if (r.status === 404) return null;
                if (!r.ok) throw await fallo(r, `no se pudo leer la carpeta ${ruta}: `);
                const j = await r.json();
                todos.push(...j.value);
                url = j['@odata.nextLink'] || null;
            }
            return todos;
        },

        /** Borra una carpeta o archivo de la biblioteca por id. Solo se usa para deshacer un lote que quedo a medias. */
        async borrarItemDrive(siteId, itemId, avisar) {
            const r = await pedir(`${graph}/sites/${siteId}/drive/items/${itemId}`, { method: 'DELETE' }, avisar);
            if (!r.ok && r.status !== 404) throw await fallo(r, `no se pudo borrar el elemento ${itemId}: `);
        },

        /** Carpeta en la biblioteca (para la evidencia). conflictBehavior rename: dos lotes iguales no se mezclan. */
        async crearCarpeta(siteId, rutaPadre, nombre, avisar) {
            const r = await pedir(`${graph}/sites/${siteId}/drive/root:/${rutaUrl(rutaPadre)}:/children`, {
                method: 'POST', headers: json,
                body: JSON.stringify({ name: nombre, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' })
            }, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo crear la carpeta ${nombre}: `);
            const j = await r.json();
            return { nombreReal: j.name, id: j.id };
        },

        async subirPieza(siteId, rutaCarpeta, nombreArchivo, bytes, tipoMime, avisar) {
            const url = `${graph}/sites/${siteId}/drive/root:/${rutaUrl(rutaCarpeta)}/${rutaUrl(nombreArchivo)}:/content`;
            const r = await pedir(url, { method: 'PUT', headers: { 'Content-Type': tipoMime }, body: bytes }, avisar);
            if (!r.ok) throw await fallo(r, `no se pudo subir ${nombreArchivo}: `);
            return await r.json();
        }
    };
}

/**
 * Traduce una columna de esquema.json al cuerpo que Graph espera en POST /columns.
 * Tipos: text | note | number | dateTime | boolean | choice.
 */
export function columnaGraph(c) {
    const base = { name: c.name || c.nombre, displayName: c.titulo || c.displayName, required: !!c.obligatorio, indexed: !!c.indexada };
    switch (c.tipo) {
        case 'text': return { ...base, text: { allowMultipleLines: false, maxLength: 255 } };
        case 'note': return { ...base, text: { allowMultipleLines: true, textType: 'plain' } };
        case 'number': return { ...base, number: { decimalPlaces: 'automatic' } };
        case 'dateTime': return { ...base, dateTime: { format: 'dateTime' } };
        case 'boolean': return { ...base, boolean: {} };
        case 'choice': return { ...base, choice: { allowTextEntry: false, choices: c.opciones || [], displayAs: 'dropDownMenu' } };
        default: throw new Error(`tipo de columna desconocido: ${c.tipo}`);
    }
}
