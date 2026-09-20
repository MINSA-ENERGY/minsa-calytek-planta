// Service worker — solo para que la app abra rapido y sobreviva a una senal mala en la caseta.
//
// REGLA QUE NO SE TOCA (heredada de minsa-captura-app, hallazgo A7): aqui se cachea UNICAMENTE
// el armazon estatico. Nada de Graph, nada del login de Microsoft. Cachear una
// respuesta de Graph dejaria embarques y placas en el disco del telefono.
//
// Cada peticion del armazon lleva `cache: 'reload'`: GitHub Pages sirve con max-age=600 y sin
// eso el service worker nuevo se llena con los archivos VIEJOS (medido en captura, 2026-08-17).

// La cache lleva la MISMA cadena que VERSION (app.js) y package.json: test/version.test.js falla si difieren (C-09, v0.21.0).
const CACHE = 'calytek-planta-v0.31.0';

function traerDeLaRed(recurso) {
    return fetch(new Request(recurso, { cache: 'reload', credentials: 'same-origin' }));
}

const ARMAZON = [
    './',
    './index.html',
    './estilo.css',
    './app.js',
    './config.js',
    './graph.js',
    './reglas.js',
    './imagen.js',
    './esquema.json',
    './manifest.json',
    './vendor/msal-browser.min.js',
    './vendor/fuentes/oswald.woff2',
    './vendor/fuentes/manrope.woff2',
    './vendor/fuentes/jetbrains-mono.woff2',
    './iconos/icono-192.png',
    './iconos/icono-512.png',
    './iconos/icono-512-recortable.png',
    './marca/lockup.svg',
    './marca/lockup-oscuro.svg'
];

self.addEventListener('install', evento => {
    evento.waitUntil(
        caches.open(CACHE)
            .then(c => Promise.all(ARMAZON.map(recurso =>
                traerDeLaRed(recurso).then(respuesta => {
                    if (!respuesta || !respuesta.ok) throw new Error(`no se pudo precargar ${recurso}`);
                    return c.put(recurso, respuesta);
                })
            )))
        // Sin skipWaiting (S-06, v0.21.0): el SW nuevo espera a que la app lo pida («Actualizar») o a que se
        // cierre la ultima pestana. Activarlo solo recargaba la pagina a media captura.
    );
});
self.addEventListener('message', evento => { if (evento.data === 'activar') self.skipWaiting(); });

self.addEventListener('activate', evento => {
    evento.waitUntil(
        caches.keys()
            .then(llaves => Promise.all(llaves.filter(k => k !== CACHE).map(k => caches.delete(k))))
            .then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', evento => {
    const url = new URL(evento.request.url);
    if (url.origin !== self.location.origin) return;
    if (evento.request.method !== 'GET') return;
    evento.respondWith(
        traerDeLaRed(evento.request.url)
            .then(respuesta => {
                if (respuesta && respuesta.ok) {
                    const copia = respuesta.clone();
                    caches.open(CACHE).then(c => c.put(evento.request, copia));
                }
                return respuesta;
            })
            .catch(() => caches.match(evento.request).then(r => r || caches.match('./index.html')))
    );
});
