// node test/sw.test.js — que la precarga del service worker no se quede corta (heredado de captura).
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const sw = readFileSync(join(raiz, 'sw.js'), 'utf8');
const FUERA = new Set(['sw.js', 'servidor-local.js']);
const modulos = readdirSync(raiz).filter(f => f.endsWith('.js') && !FUERA.has(f));
assert.ok(modulos.length > 0);
for (const m of modulos) assert.ok(sw.includes(`'./${m}'`), `sw.js precarga ${m}`);
assert.ok(sw.includes(`'./esquema.json'`), 'sw.js precarga esquema.json');
assert.ok(/const CACHE = 'calytek-planta-v\d+\.\d+\.\d+'/.test(sw), 'sw.js versiona su caché con la versión de la app (C-09)');
// S-06: el SW nuevo no se activa solo sobre una pestaña abierta; espera el mensaje «activar» de la app.
assert.ok(!/then\(\(\) => self\.skipWaiting\(\)\)/.test(sw), 'install no llama skipWaiting');
assert.ok(/addEventListener\('message'/.test(sw) && /skipWaiting\(\)/.test(sw), 'skipWaiting solo por mensaje de la app');
assert.ok(!/\.addAll\s*\(/.test(sw), 'sin addAll: pasa por la caché HTTP');
assert.ok(/cache:\s*'reload'/.test(sw), 'cada petición del armazón lleva cache: reload');
assert.ok(!/graph\.microsoft\.com|login\.microsoftonline\.com/.test(sw), 'el sw no menciona Graph ni login');
// C-44 (v0.40.0): el fetch solo intercepta el armazón; la verificación pública (certificado/) nunca pasa por la caché.
assert.ok(/if \(!RUTAS_ARMAZON\.has\(url\.pathname\)\) return;/.test(sw), 'el fetch sale temprano fuera del armazón');
assert.ok(!/'\.\/certificado\//.test(sw), 'certificado/ no está en el armazón');
// S-34 (v0.78.0), portado de proyectos: /guardar solo con Origin del propio servidor; .git/ y _salida-dev.json responden 403; Host validado.
{
  const srv = readFileSync(join(raiz, 'servidor-local.js'), 'utf8');
  assert.ok(srv.includes('ORIGENES.has(req.headers.origin)') && srv.includes('|| rutaVedada(destino)'), 'servidor-local.js debe filtrar Origin en /guardar y las rutas vedadas sobre el destino resuelto');
  assert.ok(srv.includes('if (!hostValido(req.headers.host))'), 'servidor-local.js debe rechazar un Host ajeno antes de servir');
  const { rutaVedada, hostValido } = await import('../servidor-local.js');   // importarlo NO levanta el servidor (solo escucha como programa principal)
  const { resolve } = await import('node:path');
  for (const [rel, esperado] of [['/.git/HEAD', true], ['/.gitignore', true], ['/_salida-dev.json', true], ['/_SALIDA-DEV.JSON', true], ['/x/../.git/HEAD', true], ['/test/../_salida-dev.json', true], ['/\\.git/HEAD', process.platform === 'win32'], ['/_salida-dev.json::$DATA', true], ['/_salida-dev.json.', true], ['/index.html', false], ['/test/pruebas.html', false], ['/vendor/x.js', false]])
    assert.equal(rutaVedada(resolve(raiz, '.' + rel), raiz), esperado, `rutaVedada(${rel})`);
  for (const [h, esperado] of [['localhost:8080', true], ['127.0.0.1:8080', true], ['LOCALHOST:8080', true], ['evil.example:8080', false], ['localhost:9999', false], ['', false], [undefined, false]])
    assert.equal(hostValido(h), esperado, `hostValido(${h})`);
}
console.log('sw: ok');
