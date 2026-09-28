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
console.log('graph: ok (reintento 503 y red, 403 sin reintento, cache de listas compartido y una sola lectura de /lists)');
