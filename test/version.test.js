// node test/version.test.js — una sola versión en los tres sitios (C-09, v0.21.0).
// package.json llevaba 17 versiones atrasado y la caché del SW se bumpeaba aparte: el ticket y el _lote.json
// publican VERSION, y un SW con caché vieja no refresca. Aquí falla el cierre de versión si alguno se olvidó.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const leer = f => readFileSync(join(raiz, f), 'utf8');
const pkg = JSON.parse(leer('package.json')).version;
const app = (leer('nucleo.js').match(/const VERSION = '([^']+)'/) || [])[1];
const sw = (leer('sw.js').match(/const CACHE = 'calytek-planta-v([^']+)'/) || [])[1];
assert.ok(/^\d+\.\d+\.\d+$/.test(pkg), `package.json version «${pkg}» no es semver`);
assert.equal(app, pkg, `nucleo.js VERSION (${app}) ≠ package.json (${pkg})`);
assert.equal(sw, pkg, `sw.js CACHE (${sw}) ≠ package.json (${pkg})`);
console.log(`version: ok (${pkg} en package.json, nucleo.js y sw.js)`);
