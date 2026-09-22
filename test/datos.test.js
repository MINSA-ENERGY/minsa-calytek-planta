// node test/datos.test.js — guardia contra datos personales en el repo PUBLICO (auditoria 2026-09-08).
//
// Dos capas. La fija, que viaja con el repo: ningun archivo rastreado trae un correo del tenant
// (@minsaenergy) ni una sesion de MSAL pegada (eyJ...). La privada: si junto al repo existe
// ../../herramientas-dev/datos-prohibidos.txt (vive FUERA del repo, en la maquina de desarrollo),
// se buscan tambien esas cadenas. Listarlas aqui seria publicarlas; por eso el archivo es externo.
import assert from 'node:assert/strict';
import { execSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const ESTE = 'test/datos.test.js';
const rastreados = execSync('git ls-files', { cwd: raiz, encoding: 'utf8' })
    .split(/\r?\n/).filter(Boolean)
    .filter(f => /\.(js|mjs|json|html|css|md|yml|svg|txt)$/.test(f) && f !== ESTE && !f.startsWith('vendor/'));
assert.ok(rastreados.length > 10, 'git ls-files no devolvio archivos');

const fijas = [
    [/[a-z0-9._-]+@minsaenergy\.[a-z]+/i, 'correo del tenant'],
    [/eyJ[A-Za-z0-9_-]{30,}\.[A-Za-z0-9_-]{30,}/, 'token JWT'],
    // v0.36.1: el VALOR pegado (literal de 20+ tras : o =), no el nombre — el workflow de los QR lo toma de ${{ secrets.* }} y el
    // script lo manda como campo OAuth `client_secret`, que es como se llama.
    [/(?:client_secret|clientSecret)\s*[:=]\s*['"]?[A-Za-z0-9~._-]{20,}/i, 'client secret']
];
const escapar = s => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
const privado = resolve(raiz, '..', 'herramientas-dev', 'datos-prohibidos.txt');
const privadas = existsSync(privado)
    ? readFileSync(privado, 'utf8').split(/\r?\n/).map(s => s.trim()).filter(s => s && !s.startsWith('#'))
        .map(s => [new RegExp(escapar(s), 'i'), 'dato de la lista privada'])
    : [];

const fallas = [];
for (const f of rastreados) {
    const texto = readFileSync(join(raiz, f), 'utf8');
    for (const [re, que] of [...fijas, ...privadas]) {
        const m = texto.match(re);
        if (m) fallas.push(`${f}: ${que} (${m[0].slice(0, 40)})`);
    }
}
assert.deepEqual(fallas, [], 'datos que no deben estar en el repo publico:\n  ' + fallas.join('\n  '));
console.log(`datos: ok (${rastreados.length} archivos, ${fijas.length} reglas fijas, ${privadas.length} privadas${privadas.length ? '' : ' — lista privada no encontrada'})`);
