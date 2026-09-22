// Configuracion de CALYTEK Planta (control de recepcion).
//
// TODO ESTO ES PUBLICO POR DISENO (misma regla que minsa-captura-app): en una app de pagina
// unica el client id y el tenant id no son secretos; lo que impide que alguien monte una
// pagina falsa con este client id es la lista de URL de redireccion registradas en Entra.
// NUNCA agregar aqui un client secret: esta app no lleva ninguno.

export const CONFIG = {
    // App registration "CALYTEK Planta" — la crea Carlos (docs/setup-carlos.md, tarea 1) y
    // pega aqui el Application (client) ID (hecho el 2026-09-05).
    clientId: '748d77cb-6e30-49b0-a5d3-05e47f25eb20',   // registrada por Carlos el 2026-09-05
    tenantId: 'c28754af-c62e-44db-a72a-3eeab634074b',

    // Sites.Selected: el token NO alcanza nada por si mismo. El acceso lo da la autorizacion
    // de la app sobre el sitio de CALYTEK (docs/otorgar-permiso-sitio.ps1), y encima aplican
    // los permisos de la persona en ese sitio.
    scopes: ['https://graph.microsoft.com/Sites.Selected'],

    graph: 'https://graph.microsoft.com/v1.0',
    sharepointHost: 'minsaenergy.sharepoint.com',

    // Un solo sitio: la app es de la planta. Ruta VERIFICADA contra el tenant el 2026-08-16
    // (minsa-captura-app/docs/plan.md). Nunca deducirla del nombre de la biblioteca.
    sitio: '/sites/Ambiental-CALYTEK',

    // Donde deja la evidencia (fotos del indicador): el buzon de la biblioteca, en una
    // carpeta por embarque con el contrato de nombres de minsa-energy/contrato-nombres-captura.md.
    // La skill /archivar-calytek la acomoda despues. Destino declarado en el _lote.json:
    buzon: '99_Pendiente-Archivar',
    evidencia: {
        etiqueta: 'Embarque',
        // Destino POR ANIO Y MES (2026-09-06): a ~16 lotes/dia una sola carpeta cruza los 5,000
        // elementos en ~15 meses y SharePoint deja de mostrarla. El _lote.json declara
        // `<destinoBase>/<AAAA>/<AAAA-MM>` con la fecha del lote; /archivar-calytek crea el mes
        // que falte (diagnostico `fechado` de inventario-buzon.py). Antes era fijo: 02_Planta/Otros.
        // DEUDA (2026-09-04): la BASE deberia salir de CAT_Evidencia_MINSA como en captura.
        destinoBase: '02_Planta/Bascula'
    },

    // ARCHIVOS (v0.33.0): las tres ramas del arbol que la seccion lee por Graph. La evidencia sale de evidencia.destinoBase
    // y el buzon de `buzon`; aqui solo la que no existia: donde viven los oficios ASEA y las CSF de los carriers. El artifact
    // del rediseno decia «05_Padron», pero en la biblioteca 05_ es Comercial-y-Cotizaciones y los oficios de transporte
    // estan en 04_SGI (permisos-arranque.md, 2026-08-22). Una carpeta que aun no exista sale como «aun no existe», no como error.
    archivos: { padron: '04_SGI/02_Ambiental/Permisos-y-Autorizaciones' },

    // Prefijo comun de las listas (esquema.json). Cambiarlo obliga a re-provisionar.
    listas: {
        carriers: 'PLANTA_Carriers',
        unidades: 'PLANTA_Unidades',
        choferes: 'PLANTA_Choferes',
        prealtas: 'PLANTA_PreAltas',
        embarques: 'PLANTA_Embarques',
        vigencias: 'PLANTA_Vigencias',
        roles: 'PLANTA_Roles',
        firmas: 'PLANTA_Firmas',   // S-01 (v0.24.0): sin herencia, solo validador + gerencia escriben; la compuerta manda
        certificados: 'PLANTA_Certificados'   // v0.39.0: un certificado de tratamiento por GONDOLA (embarque cerrado); lo emite gerencia y lo firma en PLANTA_Firmas
    },

    // Certificado de tratamiento (FO-CT-01, v0.35.0; decision de Carlos 2026-09-22). Lo que va impreso y no vive en
    // ninguna lista. Los dos placeholders siguen SIN DATO (KB: calytek/permisos-arranque.md): el papel los pinta como
    // «pendiente» hasta que alguien los llene aqui.
    certificado: {
        formato: 'FO-CT-01 · rev. 0',   // v0.39.0: YA NO se imprime (Carlos, 22-sep noche); queda como el nombre del formato para la KB
        autorizacionPlanta: '',            // No. de autorizacion ASEA-03-011-A de la planta (tramite no obtenido al 2026-09-22)
        responsableTecnico: '',            // nombre del responsable tecnico de planta (unica firma desde el 2026-09-22)
        resolutivo: 'Resolutivo ASEA/UGI/DGGEERC/1536/2025',
        domicilioPlanta: 'km 2 carretera Piedras Negras – Mata Espino, Tlalixcoyan, Veracruz',
        rfc: 'MIC200817HF0',
        // Pagina publica de verificacion (sin login). El QR lleva <urlVerificacion>?f=<folio>-<sufijo>. La pagina vive en
        // certificado/ de este mismo repo y lee certificado/datos/<folio>-<sufijo>.json, que REGENERA docs/exportar-planta.ps1
        // -PublicarCertificados (la app no puede escribir en el repo: es GitHub Pages).
        urlVerificacion: 'https://planta.minsaenergy.com/certificado/'
    },

    // Tolerancia del neto contra la capacidad de la unidad (decision 6): fuera de la banda se
    // pide re-captura antes de cerrar el folio. Es cinturon, no barrera: con motivo se cierra.
    tolerancia: { minFraccion: 0.3, maxFraccion: 1.25 },

    // Dias de anticipacion con que la puerta avisa (no bloquea) de un vencimiento cercano.
    avisoVigenciaDias: 30,
    // Una pre-alta FIRMADA sin gondolas en estos dias (contados desde su ultimo movimiento: firma, fecha estimada o
    // ultimo arribo) se senala en Hoy y en Pre-altas como «¿se cierra?». Solo informa: cerrar sigue siendo a mano.
    sinMovimientoDias: 15,
    // Cada cuanto se releen las listas mientras la app esta a la vista (ms). 0 = solo con Actualizar y al volver del segundo plano.
    refrescoMs: 120000,

    // CARGA ACOTADA (cubeta 3, 2026-09-06): la app leia PLANTA_Embarques COMPLETA en cada carga.
    // A ~2,000 embarques al ano eso crece lineal y sin techo. Desde ahora lee solo los ultimos
    // `ventanaDias` (columna Arribo, indexada) MAS los que siguen abiertos sin importar su edad
    // (Etapa compuerta/bruto, indexada): una gondola atorada nunca desaparece.
    // Los catalogos (carriers, unidades, choferes, vigencias, roles) y las pre-altas se siguen
    // leyendo enteros: son acotados por naturaleza (una pre-alta por PROGRAMA, no por gondola).
    // Lo que NO sale de la ventana: elegir folio (mira el ANIO entero) y decidir un borrado
    // (consulta en vivo). Ver app.js `embarquesDelAno` y `referenciasPadronVivas`.
    ventanaDias: 90,

    // Techo operativo de la planta (hallazgo de la entrevista 2026-09-04: ~8 gondolas/dia). Solo
    // para el medidor de «Hoy»; no bloquea nada.
    techoGondolasDia: 8
};
