// node test/publicador.test.mjs — el publicador de la verificacion (.github/scripts/publicar-certificados.mjs) contra un
// Graph falso (test/fixtures/graph-falso-publicador.mjs). v0.40.0: S-18 (vigente solo con firma de gerencia), C-39 (un
// vigente por gondola: el firmado mas nuevo) y C-40 (la fecha en hora de Mexico, aunque el runner corra en UTC).
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { material, descifrar } from '../.github/scripts/cifrado-certificado.mjs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const falso = pathToFileURL(join(raiz, 'test', 'fixtures', 'graph-falso-publicador.mjs')).href;
const publicador = join(raiz, '.github', 'scripts', 'publicar-certificados.mjs');
const correr = (cwd, extra = {}) => spawnSync(process.execPath, ['--import', falso, publicador], { cwd, encoding: 'utf8', env: { ...process.env, TENANT_ID: 't', CLIENT_ID: 'c', CLIENT_SECRET: 's', TZ: 'UTC', ...extra } });

const dir = mkdtempSync(join(tmpdir(), 'publicador-'));
try {
    // S-19 (v0.43.0): lo publicado antes con el nombre sha256(folio-sufijo) se borra en la primera corrida (republicar todo una vez).
    const legado = createHash('sha256').update('CT-26-0001-abcdefghjkmn').digest('hex') + '.json';
    mkdirSync(join(dir, 'certificado', 'datos'), { recursive: true }); writeFileSync(join(dir, 'certificado', 'datos', legado), '{}');
    const r = correr(dir);
    assert.equal(r.status, 0, 'el publicador termina bien: ' + r.stderr);
    const publicados = readdirSync(join(dir, 'certificado', 'datos'));
    assert.ok(!publicados.includes(legado) && publicados.includes(material('CT-26-0001', 'abcdefghjkmn').nombre), 'S-19: el nombre viejo se borra y el nuevo se escribe');
    // v0.43.0: nombre y llaves por el modulo del esquema (lo coteja test/cifrado.test.mjs contra verificar.js y el .ps1)
    const leer = (folio, suf) => { const k = material(folio, suf); return descifrar(readFileSync(join(dir, 'certificado', 'datos', k.nombre), 'utf8'), k); };
    const c1 = leer('CT-26-0001', 'abcdefghjkmn'), c2 = leer('CT-26-0002', 'bcdefghjkmnp'), c3 = leer('CT-26-0003', 'cdefghjkmnpq'), c4 = leer('CT-26-0004', 'defghjkmnpqr');
    assert.equal(c2.estado, 'vigente', 'S-18: firmado por gerencia (correo sin distinguir mayusculas) -> vigente');
    assert.deepEqual([c1.estado, c1.sustituidoPor], ['sustituido', 'CT-26-0002'], 'C-39: dos vigentes firmados en la misma gondola -> el mas nuevo manda');
    assert.equal(c3.estado, 'sin-firma', 'S-18: firma de un validador no vale para un certificado');
    assert.equal(c4.estado, 'sin-firma', 'S-18: sin renglon en PLANTA_Firmas no es vigente');
    assert.equal(c1.fechas, '22/09/2026', 'C-40: 02:30Z del 23 es el 22 en Mexico (TZ=UTC como el runner)');
    assert.equal(leer('CT-26-0005', 'efghjkmnpqrs').motivo, 'Emisión fallida', 'S-22: la cancelacion automatica no publica el error de Graph');
    assert.equal(leer('CT-26-0006', 'fghjkmnpqrst').motivo, 'El generador pidió otra razón social', 'S-22: el motivo tecleado por gerencia si se publica');
    // v0.76.0: la segunda corrida no deriva ninguna llave (el indice manda), no reescribe nada y el indice no lleva sufijos.
    const antes = Object.fromEntries(readdirSync(join(dir, 'certificado', 'datos')).map(a => [a, readFileSync(join(dir, 'certificado', 'datos', a), 'utf8')]));
    const rIdx = correr(dir);
    assert.equal(rIdx.status, 0, 'segunda corrida: ' + rIdx.stderr);
    assert.match(rIdx.stdout, / 0 escritos .* 0 llaves derivadas/, 'v0.76.0: sin cambios no se deriva ni se escribe: ' + rIdx.stdout);
    for (const [a, t] of Object.entries(antes)) assert.equal(readFileSync(join(dir, 'certificado', 'datos', a), 'utf8'), t, 'v0.76.0: el archivo no cambia: ' + a);
    const rutaIdx = join(dir, '.github', 'indice-certificados.json'), idx = readFileSync(rutaIdx, 'utf8');
    assert.ok(!/abcdefghjkmn|bcdefghjkmnp/.test(idx), 'v0.76.0: el indice (repo publico) no lleva el sufijo del QR');
    // Con otra llave del publicador (secreto rotado) las huellas no casan: se recalcula todo y lo publicado queda igual.
    const rRot = correr(dir, { CLIENT_SECRET: 'otro' });
    assert.match(rRot.stdout, / 0 escritos .* [1-9]\d* llaves derivadas/, 'v0.76.0: secreto rotado -> se derivan de nuevo, sin reescribir: ' + rRot.stdout);
    // Si falta un archivo publicado, el indice no lo tapa: se vuelve a escribir.
    const uno = material('CT-26-0002', 'bcdefghjkmnp').nombre; rmSync(join(dir, 'certificado', 'datos', uno));
    const rFalta = correr(dir, { CLIENT_SECRET: 'otro' });
    assert.match(rFalta.stdout, / 1 escritos .* 1 llaves derivadas/, 'v0.76.0: archivo borrado -> se rederiva y reescribe solo ese: ' + rFalta.stdout);
    assert.equal(leer('CT-26-0002', 'bcdefghjkmnp').estado, 'vigente', 'v0.76.0: el reescrito sigue vigente');
    // Si el publicador no ve PLANTA_Firmas (llega vacia con vigentes) aborta: no pone NO VALIDO a los buenos.
    const dir2 = mkdtempSync(join(tmpdir(), 'publicador-'));
    try { const r2 = correr(dir2, { FALSO_SIN_FIRMAS: '1' }); assert.equal(r2.status, 4, 'sin firmas legibles el publicador aborta con 4'); }
    finally { rmSync(dir2, { recursive: true, force: true }); }
} finally { rmSync(dir, { recursive: true, force: true }); }
console.log('publicador: ok (vigente / sustituido / sin-firma x2, fecha en hora de Mexico, aborta sin firmas, motivo publico fijo en cancelacion automatica)');
