// node test/publicador.test.mjs — el publicador de la verificacion (.github/scripts/publicar-certificados.mjs) contra un
// Graph falso (test/fixtures/graph-falso-publicador.mjs). v0.40.0: S-18 (vigente solo con firma de gerencia), C-39 (un
// vigente por gondola: el firmado mas nuevo) y C-40 (la fecha en hora de Mexico, aunque el runner corra en UTC).
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash, pbkdf2Sync, createDecipheriv } from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const falso = pathToFileURL(join(raiz, 'test', 'fixtures', 'graph-falso-publicador.mjs')).href;
const publicador = join(raiz, '.github', 'scripts', 'publicar-certificados.mjs');
const correr = (cwd, extra = {}) => spawnSync(process.execPath, ['--import', falso, publicador], { cwd, encoding: 'utf8', env: { ...process.env, TENANT_ID: 't', CLIENT_ID: 'c', CLIENT_SECRET: 's', TZ: 'UTC', ...extra } });

const dir = mkdtempSync(join(tmpdir(), 'publicador-'));
try {
    const r = correr(dir);
    assert.equal(r.status, 0, 'el publicador termina bien: ' + r.stderr);
    const leer = (folio, suf) => {
        const b = JSON.parse(readFileSync(join(dir, 'certificado', 'datos', createHash('sha256').update(folio + '-' + suf).digest('hex') + '.json'), 'utf8'));
        const m = pbkdf2Sync(suf, 'MINSA-CT:' + folio, 100000, 64, 'sha256');
        const d = createDecipheriv('aes-256-cbc', m.subarray(0, 32), Buffer.from(b.iv, 'base64'));
        return JSON.parse(Buffer.concat([d.update(Buffer.from(b.ct, 'base64')), d.final()]).toString());
    };
    const c1 = leer('CT-26-0001', 'abcdefghjkmn'), c2 = leer('CT-26-0002', 'bcdefghjkmnp'), c3 = leer('CT-26-0003', 'cdefghjkmnpq'), c4 = leer('CT-26-0004', 'defghjkmnpqr');
    assert.equal(c2.estado, 'vigente', 'S-18: firmado por gerencia (correo sin distinguir mayusculas) -> vigente');
    assert.deepEqual([c1.estado, c1.sustituidoPor], ['sustituido', 'CT-26-0002'], 'C-39: dos vigentes firmados en la misma gondola -> el mas nuevo manda');
    assert.equal(c3.estado, 'sin-firma', 'S-18: firma de un validador no vale para un certificado');
    assert.equal(c4.estado, 'sin-firma', 'S-18: sin renglon en PLANTA_Firmas no es vigente');
    assert.equal(c1.fechas, '22/09/2026', 'C-40: 02:30Z del 23 es el 22 en Mexico (TZ=UTC como el runner)');
    // Si el publicador no ve PLANTA_Firmas (llega vacia con vigentes) aborta: no pone NO VALIDO a los buenos.
    const dir2 = mkdtempSync(join(tmpdir(), 'publicador-'));
    try { const r2 = correr(dir2, { FALSO_SIN_FIRMAS: '1' }); assert.equal(r2.status, 4, 'sin firmas legibles el publicador aborta con 4'); }
    finally { rmSync(dir2, { recursive: true, force: true }); }
} finally { rmSync(dir, { recursive: true, force: true }); }
console.log('publicador: ok (vigente / sustituido / sin-firma x2, fecha en hora de Mexico, aborta sin firmas)');
