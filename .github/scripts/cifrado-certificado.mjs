// El esquema cifrado de la verificacion publica del certificado (C-41, v0.43.0): UN solo sitio en Node. Lo importan
// publicar-certificados.mjs y test/cifrado.test.mjs; las otras dos implementaciones (certificado/verificar.js con
// WebCrypto y docs/exportar-planta.ps1 con .NET) se cotejan contra esta en test/cifrado.test.mjs.
//   material = PBKDF2-SHA256(sufijo, sal 'MINSA-CT:<folio>', 100000 it., 96 B)
//   llaves   = material[0..32] AES-256-CBC · material[32..64] HMAC-SHA256
//   archivo  = certificado/datos/<hex(material[64..96])>.json
//   blob     = { v:1, iv, ct, mac } con mac = HMAC(iv || ct), todo en base64
// S-19 (v0.43.0): el nombre ya no es sha256(folio-sufijo) — ese era un oraculo de UN hash por intento que se saltaba el
// PBKDF2. Ahora probar un sufijo contra el nombre cuesta las mismas 100 000 iteraciones que probarlo contra el blob. Los
// primeros 64 B no cambian al pedir 96 (PBKDF2 los calcula por bloque), asi que los blobs ya publicados descifran igual.
import { pbkdf2Sync, randomBytes, createCipheriv, createDecipheriv, createHmac, timingSafeEqual } from 'node:crypto';

export const ITERACIONES = 100000;

export function material(folio, sufijo) {
    const m = pbkdf2Sync(sufijo, 'MINSA-CT:' + folio, ITERACIONES, 96, 'sha256');
    return { aes: m.subarray(0, 32), mac: m.subarray(32, 64), nombre: m.subarray(64, 96).toString('hex') + '.json' };
}
export function cifrar(doc, k) {
    const iv = randomBytes(16);
    const c = createCipheriv('aes-256-cbc', k.aes, iv);
    const ct = Buffer.concat([c.update(Buffer.from(JSON.stringify(doc), 'utf8')), c.final()]);
    const mac = createHmac('sha256', k.mac).update(Buffer.concat([iv, ct])).digest();
    return JSON.stringify({ v: 1, iv: iv.toString('base64'), ct: ct.toString('base64'), mac: mac.toString('base64') });
}
export function descifrar(json, k) {
    const b = JSON.parse(json);
    if (!b || b.v !== 1 || !b.iv || !b.ct || !b.mac) throw new Error('formato desconocido');
    const iv = Buffer.from(b.iv, 'base64'), ct = Buffer.from(b.ct, 'base64'), mac = Buffer.from(b.mac, 'base64');
    const esperado = createHmac('sha256', k.mac).update(Buffer.concat([iv, ct])).digest();
    if (mac.length !== esperado.length || !timingSafeEqual(mac, esperado)) throw new Error('mac');
    const d = createDecipheriv('aes-256-cbc', k.aes, iv);
    return JSON.parse(Buffer.concat([d.update(ct), d.final()]).toString('utf8'));
}
