// Capturas de la app a anchos REALES, por CDP, con el modo SOLO MIRAR de test/pruebas.html (U-115).
// Mismo metodo que minsa-proyectos-app/herramientas-dev/capturas.mjs. Desde app/:
//
//   node test/capturas.mjs                                   -> todas las vistas, 390 y 1366, a %TEMP%\capturas-planta
//   node test/capturas.mjs --vistas hoy,archivos --anchos 390 --salida ..\..\capturas\v0.74.0
//
// POR QUE CDP Y NO --window-size: Edge headless no baja de 504 px de viewport; la unica via a 390 px
// reales es Emulation.setDeviceMetricsOverride. Cada vista espera a que la corrida E2E termine (el
// <pre id=resultados> se oculta) y deja una linea en _mediciones.txt: overflowX (un negativo es la
// barra de scroll vertical, no un desborde) · docH · chicos(N) = objetivos tactiles de menos de 36 px.
// veredicto-* y certificado-* salen vacios: al final de la E2E no queda programa firmado ni
// certificado vigente (esas dos pantallas se revisan por lectura).

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const VISTAS_TODAS = {
    'hoy': 'vista=hoy',
    'puerta-1': 'vista=puerta&s=1', 'puerta-2': 'vista=puerta&s=2', 'puerta-3': 'vista=puerta&s=3',
    'veredicto-pasa': 'vista=puerta&vk=1', 'veredicto-rechazo': 'vista=puerta&vk=rechazo',
    'gondolas': 'vista=gondolas', 'gondolas-historial': 'vista=gondolas&g=historial', 'gondolas-rechazos': 'vista=gondolas&g=rechazos',
    'asis-bruto': 'vista=gondolas&asis=bruto', 'asis-tara': 'vista=gondolas&asis=tara', 'asis-pausa': 'vista=gondolas&asis=pausa',
    'asis-ticket': 'vista=gondolas&asis=ticket', 'asis-hoja': 'vista=gondolas&asis=tara&hoja=1',
    'prealtas': 'vista=prealtas',
    'padron': 'vista=padron', 'padron-carrier': 'vista=padron&pd=carrier', 'padron-ficha': 'vista=padron&pd=ficha', 'padron-alta': 'vista=padron&pd=alta',
    'reportes': 'vista=reportes', 'archivos': 'vista=archivos', 'certificado': 'vista=certificado',
    // con un filtro puesto: el aviso de carpetas sin abrir (U-135). El texto se teclea tras la corrida, como lo haría el usuario.
    // el menú «···» del rail abierto en oscuro, con el texto de sincronía más largo (v0.75.2: se salía del rail y quedaba recortado).
    'menu-rail': ['vista=hoy&tema=oscuro', "(() => { document.getElementById('syncRail').textContent = 'Al día · leído hace 44 s'; document.getElementById('menuRail').open = true; })()"],
    // el detalle de una pre-alta (v0.75.2: Volver arriba a la izquierda; a 390 el título se leía letra por letra).
    'prealta-detalle': ['vista=prealtas', "document.querySelector('#p-prealtas .pa-tabla [tabindex=\"0\"]')?.click()"],
    'archivos-filtro': ['vista=archivos', "(() => { const b = document.getElementById('arBusca'); b.value = 'E-25-00042'; b.dispatchEvent(new Event('input', { bubbles: true })); })()"],
};

const arg = (nombre, def) => { const i = process.argv.indexOf('--' + nombre); return i > 0 ? process.argv[i + 1] : def; };
const salida = path.resolve(arg('salida', path.join(os.tmpdir(), 'capturas-planta')));
const anchos = arg('anchos', '390,1366').split(',').map(Number);
const pedidas = arg('vistas', '');
const vistas = Object.entries(VISTAS_TODAS).filter(([n]) => !pedidas || pedidas.split(',').includes(n));
if (pedidas) for (const n of pedidas.split(',')) if (!VISTAS_TODAS[n]) { console.error(`vista desconocida: ${n} (hay: ${Object.keys(VISTAS_TODAS).join(', ')})`); process.exit(2); }
fs.mkdirSync(salida, { recursive: true });

const dormir = ms => new Promise(r => setTimeout(r, ms));
// obs. 618 (28-sep): servidor y Edge en puertos EFIMEROS. Con 8080 y 9334 fijos, si otra sesion (o el arnes de proyectos)
// ya los tenia, el hijo moria en silencio y se capturaba la app de OTRA sesion con los nombres de vistas de esta.
const srv = spawn('node', ['servidor-local.js', 'test/pruebas.html', '--puerto', '0'], { cwd: APP, stdio: ['ignore', 'pipe', 'ignore'] });
const PUERTO_WEB = await new Promise(res => {
    let buf = ''; const t = setTimeout(() => res(null), 5000);
    srv.stdout.on('data', d => { buf += d; const m = /PUERTO (\d+)/.exec(buf); if (m) { clearTimeout(t); res(Number(m[1])); } });
    srv.on('exit', () => { clearTimeout(t); res(null); });
});
if (!PUERTO_WEB) { console.error('SERVIDOR SIN PUERTO: murio al arrancar o no anuncio «PUERTO n» en 5 s. Nada se midio.'); srv.kill(); process.exit(2); }
// Con puerto 0, Edge escribe el que le dieron en <perfil>/DevToolsActivePort; el perfil es de esta corrida, asi que es SU Edge.
const perfil = path.join(os.tmpdir(), 'cdp-planta-' + Date.now());
const edge = spawn(EDGE, ['--headless=new', '--disable-gpu', '--remote-debugging-port=0', '--user-data-dir=' + perfil, 'about:blank'], { stdio: 'ignore' });
let PUERTO_CDP = null;
for (let i = 0; i < 75 && !PUERTO_CDP; i++) {
    try { PUERTO_CDP = Number(fs.readFileSync(path.join(perfil, 'DevToolsActivePort'), 'utf8').split('\n')[0]) || null; } catch { await dormir(200); }
}
if (!PUERTO_CDP) { console.error('EDGE SIN PUERTO CDP: no escribio DevToolsActivePort en 15 s. Nada se midio.'); edge.kill(); srv.kill(); process.exit(2); }
let ws; let n = 0; const pend = new Map();
const cdp = (method, params = {}, sessionId) => new Promise(res => { const id = ++n; pend.set(id, res); ws.send(JSON.stringify({ id, method, params, sessionId })); });
try {
    await dormir(2500);
    let ver; for (let i = 0; i < 50 && !ver; i++) { try { ver = await (await fetch(`http://127.0.0.1:${PUERTO_CDP}/json/version`)).json(); } catch { await dormir(200); } }
    ws = new WebSocket(ver.webSocketDebuggerUrl);
    await new Promise(r => ws.onopen = r);
    ws.onmessage = e => { const m = JSON.parse(e.data); if (m.id && pend.has(m.id)) { pend.get(m.id)(m.result || m); pend.delete(m.id); } };
    const med = [];
    for (const [nombre, def] of vistas) for (const w of anchos) {
        const [q, accion] = Array.isArray(def) ? def : [def, null];
        const { targetId } = await cdp('Target.createTarget', { url: 'about:blank' });
        const { sessionId } = await cdp('Target.attachToTarget', { targetId, flatten: true });
        await cdp('Emulation.setDeviceMetricsOverride', { width: w, height: w < 900 ? 844 : 900, deviceScaleFactor: 1, mobile: false }, sessionId);
        await cdp('Page.enable', {}, sessionId);
        await cdp('Page.navigate', { url: `http://localhost:${PUERTO_WEB}/?rol=gerencia&${q}` }, sessionId);
        let listo = false;
        for (let i = 0; i < 240 && !listo; i++) { await dormir(500); const r = await cdp('Runtime.evaluate', { expression: "document.getElementById('resultados')?.hidden === true", returnByValue: true }, sessionId); listo = r.result?.value === true; }
        // El resumen de la E2E de ESTA vista va a su linea: comparar-capturas.py rechaza una base con fallas (obs. 615).
        const e2e = listo ? String((await cdp('Runtime.evaluate', { expression: 'document.title', returnByValue: true }, sessionId)).result?.value || '').replace(/^PRUEBAS TERMINADAS:\s*/, '') : 'NO TERMINO';
        if (accion) await cdp('Runtime.evaluate', { expression: accion }, sessionId);
        await dormir(1500);
        const m = await cdp('Runtime.evaluate', { returnByValue: true, expression: `(() => { const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden'; };
          const chicos = [...document.querySelectorAll('button, a[href], input, select, summary, [role=button]')].filter(vis).filter(e => { const r = e.getBoundingClientRect(); return r.height < 36 && e.type !== 'checkbox' && e.type !== 'radio'; }).map(e => (e.id || e.className || e.tagName) + ':' + Math.round(e.getBoundingClientRect().height)).slice(0, 12);
          return 'overflowX=' + (document.documentElement.scrollWidth - innerWidth) + ' docH=' + document.documentElement.scrollHeight + ' chicos(' + chicos.length + ') ' + chicos.join(' '); })()` }, sessionId);
        const { data } = await cdp('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false }, sessionId);
        fs.writeFileSync(path.join(salida, `${nombre}-${w}.png`), Buffer.from(data, 'base64'));
        med.push(`${nombre}-${w}${listo ? '' : ' [SIN VISTA: la corrida no termino]'} · E2E ${e2e} · ${m.result?.value}`);
        console.log(med.at(-1));
        await cdp('Target.closeTarget', { targetId });
    }
    // Se funde con lo que ya había: una corrida por vista no pisa las medidas de las otras (revisor-entregable, v0.74.0).
    const archivoMed = path.join(salida, '_mediciones.txt');
    const clave = l => l.split(' ')[0];
    const previas = fs.existsSync(archivoMed) ? fs.readFileSync(archivoMed, 'utf8').split('\n').filter(l => l && !med.some(x => clave(x) === clave(l))) : [];
    fs.writeFileSync(archivoMed, [...previas, ...med].sort().join('\n'));
    console.log(`capturas en ${salida}`);
} finally {
    try { ws?.close(); } catch {} edge.kill(); srv.kill();
    await dormir(500); try { fs.rmSync(perfil, { recursive: true, force: true }); } catch { /* el perfil temporal no siempre se suelta */ }
}
