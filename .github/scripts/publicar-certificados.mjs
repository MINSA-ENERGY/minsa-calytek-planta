// Publica la verificacion de los certificados (corre en GitHub Actions; ver ../workflows/publicar-certificados.yml).
// MISMO esquema que docs/exportar-planta.ps1 v1.10.0 (paso 3c) y que certificado/verificar.js:
//   archivo  = certificado/datos/<sha256(folio-sufijo)>.json
//   llaves   = PBKDF2-SHA256(sufijo, sal 'MINSA-CT:<folio>', 100000 it.) -> 32 B AES-256-CBC + 32 B HMAC-SHA256
//   blob     = { v:1, iv, ct, mac }  con mac = HMAC(iv || ct), todo en base64
//   contenido = solo lo impreso + estado + publicadoEl; publicadoEl no fuerza reescritura (se descifra el existente y se compara)
// Sin dependencias: fetch y crypto de Node 22. Solo lee del tenant; escribe solo en certificado/datos/.
// S-18 (app v0.40.0): «vigente» se publica SOLO con la firma que la app exige (certificadoFirmado): un renglon de
// PLANTA_Firmas Tipo=certificado, ObjetoId = el certificado, Firmante = EmitidoPor, y ese Firmante con rol gerencia
// ACTIVO en PLANTA_Roles. Sin eso sale «sin-firma» (la pagina dice NO VALIDO). Y una gondola tiene UN vigente: si una
// sustitucion quedo a medias (C-39) el mas nuevo firmado manda y los otros salen «sustituido» por el.
// C-40 (v0.40.0): las fechas en hora de Mexico, como el papel — el runner de Actions corre en UTC.
import { createHash, pbkdf2Sync, randomBytes, createCipheriv, createDecipheriv, createHmac, timingSafeEqual } from 'node:crypto';
import { readFileSync, writeFileSync, readdirSync, unlinkSync, mkdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) { console.error('faltan PUBLICADOR_TENANT_ID / PUBLICADOR_CLIENT_ID / PUBLICADOR_CLIENT_SECRET (secrets del repo)'); process.exit(2); }
const HOST = 'minsaenergy.sharepoint.com', RUTA_SITIO = '/sites/Ambiental-CALYTEK', LISTA = 'PLANTA_Certificados';
const CARPETA = join(process.cwd(), 'certificado', 'datos');
const RE_NOMBRE = /^CT-\d{2}-\d{4}-[a-z2-9]{12}$/;

async function token() {
    const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' })
    });
    if (!r.ok) throw new Error('token: ' + r.status + ' ' + (await r.text()).slice(0, 300));
    return (await r.json()).access_token;
}
async function graph(url, tk) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + tk } });
    if (!r.ok) throw new Error(`graph ${r.status} ${url.slice(0, 120)}: ${(await r.text()).slice(0, 300)}`);
    return r.json();
}
function llaves(folio, sufijo) {
    const m = pbkdf2Sync(sufijo, 'MINSA-CT:' + folio, 100000, 64, 'sha256');
    return { aes: m.subarray(0, 32), mac: m.subarray(32, 64) };
}
function cifrar(doc, folio, sufijo) {
    const k = llaves(folio, sufijo), iv = randomBytes(16);
    const c = createCipheriv('aes-256-cbc', k.aes, iv);
    const ct = Buffer.concat([c.update(Buffer.from(JSON.stringify(doc), 'utf8')), c.final()]);
    const mac = createHmac('sha256', k.mac).update(Buffer.concat([iv, ct])).digest();
    return JSON.stringify({ v: 1, iv: iv.toString('base64'), ct: ct.toString('base64'), mac: mac.toString('base64') });
}
function descifrar(json, folio, sufijo) {
    const b = JSON.parse(json); const k = llaves(folio, sufijo);
    const iv = Buffer.from(b.iv, 'base64'), ct = Buffer.from(b.ct, 'base64'), mac = Buffer.from(b.mac, 'base64');
    const esperado = createHmac('sha256', k.mac).update(Buffer.concat([iv, ct])).digest();
    if (mac.length !== esperado.length || !timingSafeEqual(mac, esperado)) throw new Error('mac');
    const d = createDecipheriv('aes-256-cbc', k.aes, iv);
    return JSON.parse(Buffer.concat([d.update(ct), d.final()]).toString('utf8'));
}
const FMT_DIA = new Intl.DateTimeFormat('en-GB', { timeZone: 'America/Mexico_City', day: '2-digit', month: '2-digit', year: 'numeric' });
const dia = iso => FMT_DIA.format(new Date(iso));   // dd/mm/aaaa en hora de Mexico (C-40)

const tk = await token();
const site = await graph(`https://graph.microsoft.com/v1.0/sites/${HOST}:${RUTA_SITIO}`, tk);
const listas = await graph(`https://graph.microsoft.com/v1.0/sites/${site.id}/lists?$select=id,name,displayName&$top=200`, tk);
async function leerLista(nombre) {
    const l = listas.value.find(x => x.name === nombre || x.displayName === nombre);
    if (!l) { console.error(`la lista ${nombre} no existe en el sitio`); process.exit(3); }
    let url = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${l.id}/items?expand=fields&$top=500`;
    const out = [];
    while (url) { const p = await graph(url, tk); out.push(...p.value.map(it => ({ ...it.fields, _id: Number(it.id) }))); url = p['@odata.nextLink']; }
    return out;
}
const renglones = await leerLista(LISTA);
const firmas = await leerLista('PLANTA_Firmas'), roles = await leerLista('PLANTA_Roles');
const minus = s => String(s || '').trim().toLowerCase();
const esGerencia = correo => roles.some(r => minus(r.Title) === minus(correo) && r.Activo !== false && r.Rol === 'gerencia');
const firmado = f => firmas.some(x => x.Tipo === 'certificado' && Number(x.ObjetoId) === f._id && minus(x.Firmante) === minus(f.EmitidoPor) && esGerencia(x.Firmante));
// El vigente que manda por gondola: el firmado de id mas alto. Los certificados de programa (v0.35-v0.38) no traen EmbarqueId.
const mandaPorGondola = new Map();
for (const f of renglones) if (f.Estado === 'vigente' && f.EmbarqueId != null && firmado(f)) { const k = Number(f.EmbarqueId), a = mandaPorGondola.get(k); if (!a || f._id > a._id) mandaPorGondola.set(k, f); }
function estadoPublico(f) {
    if (f.Estado !== 'vigente') return { estado: f.Estado, sustituidoPor: f.SustituidoPor ?? null };
    if (!firmado(f)) return { estado: 'sin-firma', sustituidoPor: null };
    const m = f.EmbarqueId != null ? mandaPorGondola.get(Number(f.EmbarqueId)) : null;
    if (m && m._id !== f._id) return { estado: 'sustituido', sustituidoPor: m.Title };
    return { estado: 'vigente', sustituidoPor: null };
}
// Si PLANTA_Firmas llega VACIA habiendo vigentes, el permiso del publicador no alcanza la lista (tiene la herencia rota):
// publicar asi pondria NO VALIDO a todos los certificados buenos. Se aborta y el Action queda en rojo; lo publicado no cambia.
if (!firmas.length && renglones.some(f => f.Estado === 'vigente')) { console.error('PLANTA_Firmas se leyo vacia con certificados vigentes: el publicador no ve las firmas. No se publica.'); process.exit(4); }
let sinFirma = 0;

mkdirSync(CARPETA, { recursive: true });
const ahora = new Date().toISOString();
const esperados = new Set();
let escritos = 0, sinSufijo = 0, iguales = 0;
for (const f of renglones) {
    const nombre = `${f.Title || ''}-${f.Sufijo || ''}`;
    if (!RE_NOMBRE.test(nombre)) { sinSufijo++; continue; }
    const residuo = 'RECORTES DE PERFORACION';   // v0.38.0: sin la base, como el papel (Carlos, 23-sep)
    // v0.39.0: un certificado por GONDOLA trae UN manifiesto y UN ticket de bascula. Los emitidos de la v0.35.0 a la
    // v0.38.0 (uno por programa) traen la lista de embarques y el rango de fechas: se publican con lo que tengan.
    const emb = String(f.Embarques || '').split(';').map(s => s.trim()).filter(Boolean);
    const manifiesto = f.Manifiesto || String(f.Manifiestos || '').split(';').map(s => s.trim()).filter(Boolean).join(' / ') || null;
    let fechas = null;
    if (f.FechaRecepcion) fechas = dia(f.FechaRecepcion);
    else if (f.PrimerCierre && f.UltimoCierre) { const d1 = dia(f.PrimerCierre), d2 = dia(f.UltimoCierre); fechas = d1 === d2 ? d1 : `del ${d1} al ${d2}`; }
    const pub = estadoPublico(f); if (pub.estado === 'sin-firma') sinFirma++;
    const doc = {
        folio: f.Title, estado: pub.estado, generador: f.Generador ?? null, registro: f.GeneradorRegistro ?? null, direccion: f.GeneradorDireccion ?? null, pozo: f.Pozo ?? null,
        residuo, kg: f.Kg ?? null, manifiesto, ticket: f.TicketBascula ?? null, embarques: emb, fechas, transportista: f.Transportista ?? null,
        emitidoEl: f.EmitidoEl ?? null, sustituidoPor: pub.sustituidoPor, motivo: f.Estado === 'cancelado' ? (f.Motivo ?? null) : null, publicadoEl: ahora
    };
    const archivo = createHash('sha256').update(nombre, 'utf8').digest('hex') + '.json';
    const ruta = join(CARPETA, archivo);
    esperados.add(archivo);
    if (existsSync(ruta)) {
        try {
            const viejo = descifrar(readFileSync(ruta, 'utf8'), f.Title, f.Sufijo);
            const a = { ...doc, publicadoEl: viejo.publicadoEl };
            if (JSON.stringify(a) === JSON.stringify(viejo)) { iguales++; continue; }
        } catch (_) { /* ilegible: se reescribe */ }
    }
    writeFileSync(ruta, cifrar(doc, f.Title, f.Sufijo), 'utf8'); escritos++;
}
let borrados = 0;
for (const a of readdirSync(CARPETA)) if (a.endsWith('.json') && !esperados.has(a)) { unlinkSync(join(CARPETA, a)); borrados++; }
console.log(`certificados: ${renglones.length} renglones · ${esperados.size} publicables · ${escritos} escritos · ${iguales} sin cambio · ${borrados} borrados · ${sinSufijo} sin sufijo valido · ${sinFirma} vigentes SIN FIRMA publicados como no validos`);
