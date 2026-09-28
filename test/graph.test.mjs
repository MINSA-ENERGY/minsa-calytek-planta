// node test/graph.test.mjs — C-95 (v0.81.0): conReintento y el cache de ids de lista (C-92) no tenian prueba.
import assert from 'node:assert/strict';
import { conReintento, crearCliente } from '../graph.js';

const resp = (status, cuerpo = {}) => ({ ok: status >= 200 && status < 300, status, headers: { get: () => null }, json: async () => cuerpo });

// 503 y luego 200: se reintenta y se avisa una vez.
{
    let n = 0; const avisos = [];
    const r = await conReintento(async () => (++n === 1 ? resp(503) : resp(200)), t => avisos.push(t));
    assert.equal(r.status, 200); assert.equal(n, 2); assert.equal(avisos.length, 1);
}
// 403 no mejora repitiendo: una sola llamada.
{
    let n = 0;
    const r = await conReintento(async () => { n++; return resp(403); });
    assert.equal(r.status, 403); assert.equal(n, 1);
}
// Un fallo de red (TypeError de fetch) se reintenta y la segunda vuelta entra.
{
    let n = 0;
    const r = await conReintento(async () => { if (++n === 1) throw new TypeError('Failed to fetch'); return resp(200); });
    assert.equal(r.status, 200); assert.equal(n, 2);
}
// C-92: dos clientes con el mismo Map comparten los ids, y varias listas pedidas a la vez leen /lists UNA vez.
{
    let lecturas = 0;
    globalThis.fetch = async url => {
        if (String(url).includes('/lists?')) { lecturas++; return resp(200, { value: [{ id: 'id-a', name: 'A', displayName: 'A' }, { id: 'id-b', name: 'B', displayName: 'B' }] }); }
        return resp(404);
    };
    const cache = new Map();
    const c1 = crearCliente('https://graph.example.invalid/v1.0', 't1', cache);
    const ids = await Promise.all([c1.idDeLista('s', 'A'), c1.idDeLista('s', 'B'), c1.idDeLista('s', 'A')]);
    assert.deepEqual(ids, ['id-a', 'id-b', 'id-a']); assert.equal(lecturas, 1);
    const c2 = crearCliente('https://graph.example.invalid/v1.0', 't2', cache);
    assert.equal(await c2.idDeLista('s', 'B'), 'id-b'); assert.equal(lecturas, 1);
}
// C-88 (v0.82.0): un POST NO se repite ante red caida, timeout o 504 — pudo haberse escrito; sale resultadoIncierto.
for (const hacerFalla of [() => { throw new TypeError('Failed to fetch'); }, () => { throw new DOMException('timeout', 'TimeoutError'); }, () => resp(504)]) {
    let n = 0;
    await assert.rejects(conReintento(async () => { n++; return hacerFalla(); }, null, false), e => e.codigo === 'resultadoIncierto');
    assert.equal(n, 1);
}
// ...y el resultado incierto pide la relectura (planta:releer) aunque quien llama se trague el error, como hacen los 12 llamadores.
{
    let releidas = 0;
    globalThis.window = new EventTarget();
    window.addEventListener('planta:releer', () => releidas++);
    try { await conReintento(async () => resp(504), null, false); } catch (_) { /* el llamador avisa y no relanza */ }
    await new Promise(res => setTimeout(res, 5));
    assert.equal(releidas, 1);
    delete globalThis.window;
}
// ...pero 429/503 (no se proceso) si se reintentan, y un 400 sale como respuesta sin reintento.
{
    let n = 0;
    const r = await conReintento(async () => (++n === 1 ? resp(503) : resp(201)), null, false);
    assert.equal(r.status, 201); assert.equal(n, 2);
    n = 0;
    const r2 = await conReintento(async () => { n++; return resp(400); }, null, false);
    assert.equal(r2.status, 400); assert.equal(n, 1);
}
// C-88 + C-93: el cliente manda POST sin reintento y TODA peticion con un AbortSignal (uno nuevo por intento).
{
    const vistos = [];
    globalThis.fetch = async (url, op) => {
        vistos.push({ metodo: op.method || 'GET', signal: op.signal });
        if (String(url).includes('/lists?')) return resp(200, { value: [{ id: 'id-a', name: 'A', displayName: 'A' }] });
        throw new TypeError('Failed to fetch');
    };
    const c = crearCliente('https://graph.example.invalid/v1.0', 't', new Map());
    await assert.rejects(c.crearRenglon('s', 'A', { Title: 'x' }), e => e.codigo === 'resultadoIncierto');
    assert.equal(vistos.filter(v => v.metodo === 'POST').length, 1);
    assert.ok(vistos.every(v => v.signal instanceof AbortSignal));
}
console.log('graph: ok (reintento 503 y red, 403 sin reintento, cache de listas compartido y una sola lectura de /lists, POST sin reintento incierto, timeout en cada peticion)');
