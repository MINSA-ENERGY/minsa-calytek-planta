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
// S-13 (v0.30.0): la E2E sustituye window.msal por un falso, asi que un bundle que cambie la API sale verde alli. Aqui se
// evalua el UMD REAL con un `exports` propio (el bundle no toca window/document al cargarse) y se exige lo que app.js usa.
const exportsMsal = {};
new Function('exports', 'module', 'define', bytes.toString('utf8'))(exportsMsal, { exports: exportsMsal }, undefined);
const PCA = exportsMsal.PublicClientApplication;
assert.equal(typeof PCA, 'function', 'el bundle no exporta PublicClientApplication');
const metodos = new Set();
for (let p = PCA.prototype; p && p !== Object.prototype; p = Object.getPrototypeOf(p)) Object.getOwnPropertyNames(p).forEach(n => metodos.add(n));
const usados = ['initialize', 'getAllAccounts', 'acquireTokenSilent', 'acquireTokenRedirect', 'handleRedirectPromise', 'loginRedirect', 'logoutRedirect'];
const faltan = usados.filter(n => !metodos.has(n));
assert.deepEqual(faltan, [], `msal.PublicClientApplication ya no expone: ${faltan.join(', ')} (app.js los llama)`);
assert.equal(typeof exportsMsal.InteractionRequiredAuthError, 'function', 'el bundle no exporta InteractionRequiredAuthError (refrescarCliente lo usa)');
assert.equal(exportsMsal.CacheLookupPolicy && exportsMsal.CacheLookupPolicy.AccessTokenAndRefreshToken, 2, 'CacheLookupPolicy.AccessTokenAndRefreshToken ≠ 2 (token() lo usa para no caer al iframe)');
const version = (bytes.toString('utf8', 0, 80).match(/@azure\/msal-browser v(\d+\.\d+\.\d+)/) || [])[1];
assert.ok(version && integridad.includes(`@azure/msal-browser ${version} |`), `la cabecera del bundle dice v${version} y INTEGRIDAD.md no tiene fila con esa version`);
console.log(`vendor: ok (msal-browser ${version}, sha256 ${sha256.slice(0, 12)}…, integrity cotejado, API de app.js presente)`);
