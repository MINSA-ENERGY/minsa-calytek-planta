// node test/vendor.test.js — lo vendorizado es lo que dice INTEGRIDAD.md y lo que exige index.html (S-03 / S-05, v0.21.0).
// Sin package-lock no hay npm audit: la única librería con superficie (msal-browser maneja el token) se
// inventaría a mano en vendor/INTEGRIDAD.md. Esta prueba cierra el hueco por el otro lado: el archivo del
// repo tiene el sha256 registrado, y el integrity= de index.html es el sha384 de ESE archivo — un PR que
// cambie uno sin el otro falla aquí antes de que el navegador se niegue a cargarlo.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const bytes = readFileSync(join(raiz, 'vendor', 'msal-browser.min.js'));
const sha256 = createHash('sha256').update(bytes).digest('hex');
const sha384 = 'sha384-' + createHash('sha384').update(bytes).digest('base64');

const integridad = readFileSync(join(raiz, 'vendor', 'INTEGRIDAD.md'), 'utf8');
const filas = [...integridad.matchAll(/^\| `msal-browser\.min\.js` \|[^\n]*`([0-9a-f]{64})`/gm)].map(m => m[1]);
assert.ok(filas.length > 0, 'INTEGRIDAD.md no trae ninguna fila de msal-browser.min.js con sha256');
assert.equal(filas[filas.length - 1], sha256, 'el sha256 del archivo no es el de la ÚLTIMA fila de INTEGRIDAD.md (¿subió el vendor sin anotarlo?)');

const html = readFileSync(join(raiz, 'index.html'), 'utf8');
const m = html.match(/<script src="\.\/vendor\/msal-browser\.min\.js" integrity="([^"]+)"/);
assert.ok(m, 'index.html carga msal-browser.min.js sin integrity=');
assert.equal(m[1], sha384, `integrity= de index.html (${m[1]}) ≠ sha384 del archivo (${sha384})`);
console.log(`vendor: ok (msal-browser sha256 ${sha256.slice(0, 12)}…, integrity cotejado)`);
