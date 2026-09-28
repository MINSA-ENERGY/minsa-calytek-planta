// node test/imports.test.mjs — C-83 (v0.79.0): ningun import de un modulo propio queda sin uso ni tapado por un nombre local.
// C-76 partio app.js en 11 modulos y dejo 19 asi: cada uno fabricaba una dependencia (a veces circular) que nadie usaba, y un
// import tapado por un `const` del mismo nombre hace creer que se llama al compartido. El texto de comentarios, cadenas y
// plantillas NO cuenta como uso («por firmar» no usa firmar); lo que va dentro de ${...} si.
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..');
const FUERA = new Set(['sw.js', 'servidor-local.js', 'config.js']);

/** El codigo sin comentarios ni texto literal: cadenas y el texto de las plantillas pasan a espacios; ${...} se conserva. */
export function soloCodigo(src) {
    let out = '', i = 0, previo = '';
    const pila = [];   // profundidad de llaves de cada ${ abierto
    const puedeRegex = () => /[(,=:[!&|?{};+\-*%<>~^]$|^$|\breturn$|\btypeof$/.test(previo);
    while (i < src.length) {
        const c = src[i], d = src[i + 1];
        if (c === '/' && d === '/') { while (i < src.length && src[i] !== '\n') i++; continue; }
        if (c === '/' && d === '*') { const f = src.indexOf('*/', i + 2); i = f < 0 ? src.length : f + 2; out += ' '; continue; }
        if (c === "'" || c === '"') { i++; while (i < src.length && src[i] !== c) { if (src[i] === '\\') i++; i++; } i++; out += "''"; previo = "'"; continue; }
        if (c === '/' && puedeRegex()) {   // literal de regex: se salta entero (puede traer comillas o acentos graves)
            i++; let clase = false;
            while (i < src.length && (clase || src[i] !== '/')) { if (src[i] === '\\') i++; else if (src[i] === '[') clase = true; else if (src[i] === ']') clase = false; i++; }
            i++; while (/[a-z]/.test(src[i] || '')) i++;
            out += '/r/'; previo = '/'; continue;
        }
        if (c === '`' || (c === '}' && pila.length && pila[pila.length - 1] === 0)) {   // texto de plantilla (al abrirla o al cerrar un ${})
            if (c === '}') pila.pop();
            i++;
            while (i < src.length && src[i] !== '`' && !(src[i] === '$' && src[i + 1] === '{')) { if (src[i] === '\\') i++; i++; }
            if (src[i] === '$') { i += 2; pila.push(0); out += ' ( '; previo = '('; }
            else { i++; out += '``'; previo = '`'; }
            continue;
        }
        if (pila.length) { if (c === '{') pila[pila.length - 1]++; else if (c === '}') pila[pila.length - 1]--; }
        out += c;
        if (!/\s/.test(c)) previo = /[\w$]/.test(c) ? (previo + c).slice(-8) : c;
        i++;
    }
    return out;
}

const IMPORT = /import\s*\{([^}]*)\}\s*from\s*'\.\/([\w-]+)\.js';?/g;
const esc = n => n.replace(/\$/g, '\\$');
const hallazgos = [];
for (const f of readdirSync(raiz).filter(x => x.endsWith('.js') && !FUERA.has(x))) {
    const src = readFileSync(join(raiz, f), 'utf8');
    const cuerpo = soloCodigo(src.replace(IMPORT, ''));
    for (const [, lista, mod] of src.matchAll(IMPORT)) {
        for (const n of lista.split(',').map(s => s.trim()).filter(Boolean)) {
            if (!new RegExp(`(^|[^\\w$.])${esc(n)}(?![\\w$])|\\.\\.\\.${esc(n)}(?![\\w$])`).test(cuerpo)) hallazgos.push(`${f}: import ${n} de ${mod}.js sin uso`);
            if (new RegExp(`\\b(const|let|var|function)\\s+${esc(n)}(?![\\w$])|(^|[^\\w$.])${esc(n)}\\s*=>`).test(cuerpo)) hallazgos.push(`${f}: import ${n} de ${mod}.js tapado por un nombre local`);
        }
    }
}
// El detector mismo: lo que antes pasaba (texto, plantilla, regex con comillas) no cuenta como uso; ${} si.
assert.equal(/\bfirmar\b/.test(soloCodigo("x = `pre-altas por firmar`; // firmar\n")), false);
assert.equal(/\bfirmar\b/.test(soloCodigo('x = `a ${firmar(1)} b`;')), true);
assert.equal(/\bdentro\b/.test(soloCodigo("const r = /[\"'`]/g; y = 'dentro';")), false);
assert.equal(/\bdentro\b/.test(soloCodigo('const r = /[`]/g; dentro();')), true);
assert.deepEqual(hallazgos, [], hallazgos.join('\n'));
console.log('imports: ok');
