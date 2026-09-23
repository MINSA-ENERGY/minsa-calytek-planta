// node test/cifrado.test.mjs — C-41 (v0.43.0): el esquema cifrado de la verificacion publica vive en TRES lenguajes
// (certificado/verificar.js con WebCrypto, .github/scripts/cifrado-certificado.mjs con node:crypto y
// docs/exportar-planta.ps1 con .NET) y hasta hoy ninguna prueba los cotejaba. Aqui cada uno descifra lo que cifro otro
// y los tres sacan el MISMO nombre de archivo. S-19: ese nombre sale de los bytes 64..96 del PBKDF2, no de un sha256.
//   · el .ps1 se corre en vivo si hay powershell (Windows) y SIEMPRE se coteja contra test/fixtures/cifrado-ps1.json,
//     que produjo el .ps1 y esta versionado (regenerarlo: node test/cifrado.test.mjs --fijar).
//   · verificar.js se corre entero en un vm con document/fetch falsos: lo que se prueba es la pagina, no una copia.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash, pbkdf2Sync } from 'node:crypto';
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import { material, cifrar, descifrar } from '../.github/scripts/cifrado-certificado.mjs';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURE = join(raiz, 'test', 'fixtures', 'cifrado-ps1.json');
const EXPORTADOR = join(raiz, '..', 'docs', 'exportar-planta.ps1');
const folio = 'CT-26-0042', sufijo = 'k7m2p9q4r8s3';
const doc = {
    folio, estado: 'vigente', generador: 'Perforadora Ñandú & <b>Hijos</b>', registro: 'GEN-123', direccion: 'Km 12 carr. Poza Rica–Tuxpan', pozo: 'Rabasa-101',
    residuo: 'RECORTES DE PERFORACION', kg: 12345, manifiesto: 'M-001', ticket: 'T-77', embarques: ['E-1', 'E-2'], fechas: '22/09/2026',
    transportista: null, emitidoEl: '2026-09-22T18:00:00Z', sustituidoPor: null, motivo: null, publicadoEl: '2026-09-22T19:00:00Z'
};
const alterar = blob => { const b = JSON.parse(blob); const ct = Buffer.from(b.ct, 'base64'); ct[0] ^= 1; b.ct = ct.toString('base64'); return JSON.stringify(b); };

// --- 1. El esquema fijado: nombre = hex(PBKDF2 bytes 64..96), llaves = bytes 0..64, y ya NO es sha256(folio-sufijo).
const k = material(folio, sufijo);
const crudo = pbkdf2Sync(sufijo, 'MINSA-CT:' + folio, 100000, 96, 'sha256');
assert.equal(k.nombre, crudo.subarray(64, 96).toString('hex') + '.json', 'S-19: el nombre son los bytes 64..96 del PBKDF2');
assert.deepEqual([Buffer.from(k.aes), Buffer.from(k.mac)], [crudo.subarray(0, 32), crudo.subarray(32, 64)], 'las llaves no cambian al pedir 96 B');
const legado = createHash('sha256').update(folio + '-' + sufijo).digest('hex') + '.json';
assert.notEqual(k.nombre, legado, 'S-19: el nombre ya no es el sha256 rapido');
const blobNode = cifrar(doc, k);
assert.deepEqual(descifrar(blobNode, k), doc, 'Node descifra lo que cifro');
assert.throws(() => descifrar(alterar(blobNode), k), /mac/, 'Node rechaza un blob alterado');

// --- 2. El .ps1: en vivo (si hay powershell) y contra el fixture versionado.
const fijar = process.argv.includes('--fijar');
let vivo = null;
if (process.platform === 'win32') {
    const dir = mkdtempSync(join(tmpdir(), 'cifrado-ps1-'));
    try {
        const ent = join(dir, 'entrada.json'), sal = join(dir, 'salida.json');
        writeFileSync(ent, JSON.stringify({ folio, sufijo, doc, blobNode, blobAlterado: alterar(blobNode) }), 'utf8');
        const r = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(raiz, 'test', 'cifrado-ps1.ps1'), ent, sal, EXPORTADOR], { encoding: 'utf8' });
        assert.equal(r.status, 0, 'el arnes del .ps1 termina bien: ' + r.stderr + r.stdout);
        vivo = JSON.parse(readFileSync(sal, 'utf8'));
    } finally { rmSync(dir, { recursive: true, force: true }); }
    assert.equal(vivo.nombre, k.nombre, 'el .ps1 saca el mismo nombre de archivo que Node');
    assert.deepEqual(descifrar(vivo.blob, k), doc, 'Node descifra lo que cifro el .ps1');
    assert.deepEqual(vivo.descifrado, doc, 'el .ps1 descifra lo que cifro Node');
    assert.equal(vivo.alterado, 'mac', 'el .ps1 rechaza un blob alterado');
    if (fijar) writeFileSync(FIXTURE, JSON.stringify({ generadoPor: 'docs/exportar-planta.ps1 via test/cifrado-ps1.ps1', folio, sufijo, doc, nombre: vivo.nombre, blob: vivo.blob }, null, 2) + '\n', 'utf8');
} else {
    assert.ok(!fijar, '--fijar necesita powershell');
    console.log('cifrado: OMITIDO el .ps1 en vivo (sin powershell en esta plataforma); se coteja solo el fixture');
}
const fx = JSON.parse(readFileSync(FIXTURE, 'utf8'));
const kfx = material(fx.folio, fx.sufijo);
assert.equal(fx.nombre, kfx.nombre, 'el nombre que dio el .ps1 al fijar el fixture sigue siendo el de Node');
assert.deepEqual(descifrar(fx.blob, kfx), fx.doc, 'Node descifra el blob versionado del .ps1');

// --- 3. verificar.js entero, en un vm, contra blobs de los dos publicadores.
const fuente = readFileSync(join(raiz, 'certificado', 'verificar.js'), 'utf8');
async function verificar(f, archivos) {
    const nodo = () => ({ className: '', textContent: '', hijos: [], classList: { remove() {}, add() {} }, appendChild(c) { this.hijos.push(c); }, insertAdjacentElement() {} });
    const els = {};
    const pedidos = [];
    const ctx = {
        document: { getElementById: id => (els[id] ??= nodo()), createElement: () => nodo() },
        location: { search: '?f=' + encodeURIComponent(f) }, URLSearchParams, TextEncoder, TextDecoder, atob, crypto: globalThis.crypto,
        fetch: async url => { pedidos.push(url); const a = archivos[url.replace('./datos/', '')]; return a ? { status: 200, ok: true, json: async () => JSON.parse(a) } : { status: 404, ok: false }; }
    };
    vm.runInNewContext(fuente, ctx);
    for (let i = 0; i < 500 && !(els.estado && els.estado.className); i++) await new Promise(r => setTimeout(r, 10));
    return { clase: els.estado?.className, texto: els.estado?.textContent, datos: (els.datos?.hijos || []).map(h => h.textContent), pedidos };
}
const f = folio + '-' + sufijo;
const deNode = await verificar(f, { [k.nombre]: blobNode });
assert.equal(deNode.clase, 'estado ok', 'verificar.js descifra lo que cifro Node: ' + deNode.texto);
assert.ok(deNode.datos.includes(doc.generador) && deNode.datos.includes(doc.direccion), 'verificar.js pinta generador y direccion');
assert.equal(deNode.pedidos[0], './datos/' + k.nombre, 'S-19: verificar.js pide primero el nombre derivado del PBKDF2');
const dePs1 = await verificar(fx.folio + '-' + fx.sufijo, { [fx.nombre]: fx.blob });
assert.equal(dePs1.clase, 'estado ok', 'verificar.js descifra lo que cifro el .ps1: ' + dePs1.texto);
if (vivo) assert.equal((await verificar(f, { [vivo.nombre]: vivo.blob })).clase, 'estado ok', 'verificar.js descifra el blob en vivo del .ps1');
// v0.45.2: el respaldo de transicion salio — un archivo con el nombre viejo (sha256) ya NO se lee, y se pide un solo nombre.
const transicion = await verificar(f, { [legado]: blobNode });
assert.ok(transicion.clase === 'estado ojo' && transicion.pedidos.length === 1, 'S-19: el nombre viejo ya no se consulta (sin oraculo): ' + transicion.clase + ' · ' + transicion.pedidos.join(', '));
const alterado = await verificar(f, { [k.nombre]: alterar(blobNode) });
assert.ok(alterado.clase === 'estado mal' && /sello/.test(alterado.texto), 'verificar.js rechaza un blob alterado: ' + alterado.texto);
assert.equal((await verificar(f, {})).clase, 'estado ojo', 'sin archivo: aviso ambar (U-78)');

console.log(`cifrado: ok (Node · verificar.js · .ps1 ${vivo ? 'en vivo + ' : ''}fixture: mismo nombre derivado del PBKDF2, cada uno descifra al otro, blob alterado rechazado, el nombre viejo ya no se consulta)`);
