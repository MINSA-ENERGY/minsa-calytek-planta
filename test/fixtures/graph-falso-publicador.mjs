// Graph falso para test/publicador.test.mjs (v0.40.0, S-18 / C-39 / C-40): se carga con node --import antes del publicador
// y sustituye fetch. Cuatro certificados: firmado, doble vigente (el mas nuevo manda), firma de validador y sin firma.
const certs = [
  { _i: 11, Title: 'CT-26-0001', Sufijo: 'abcdefghjkmn', Estado: 'vigente', EmbarqueId: 5, EmitidoPor: 'G@minsa.mx', FechaRecepcion: '2026-09-23T02:30:00Z', Kg: 1000 },
  { _i: 12, Title: 'CT-26-0002', Sufijo: 'bcdefghjkmnp', Estado: 'vigente', EmbarqueId: 5, EmitidoPor: 'g@minsa.mx', FechaRecepcion: '2026-09-23T02:30:00Z', Kg: 1000 },
  { _i: 13, Title: 'CT-26-0003', Sufijo: 'cdefghjkmnpq', Estado: 'vigente', EmbarqueId: 6, EmitidoPor: 'v@minsa.mx', FechaRecepcion: '2026-09-22T15:00:00Z', Kg: 1000 },
  { _i: 14, Title: 'CT-26-0004', Sufijo: 'defghjkmnpqr', Estado: 'vigente', EmbarqueId: 7, EmitidoPor: 'g@minsa.mx', Kg: 1000 }
];
const firmas = [
  { _i: 1, Tipo: 'certificado', ObjetoId: 11, Firmante: 'g@minsa.mx' },
  { _i: 2, Tipo: 'certificado', ObjetoId: 12, Firmante: 'g@minsa.mx' },
  { _i: 3, Tipo: 'certificado', ObjetoId: 13, Firmante: 'v@minsa.mx' }
];
const roles = [{ _i: 1, Title: 'g@minsa.mx', Rol: 'gerencia', Activo: true }, { _i: 2, Title: 'v@minsa.mx', Rol: 'validador', Activo: true }];
const items = a => ({ value: a.map(({ _i, ...f }) => ({ id: String(_i), fields: f })) });
globalThis.fetch = async (url) => {
  const j = b => new Response(JSON.stringify(b), { status: 200 });
  if (url.includes('oauth2')) return j({ access_token: 't' });
  if (url.includes('/lists?')) return j({ value: [{ id: 'lc', name: 'PLANTA_Certificados' }, { id: 'lf', name: 'PLANTA_Firmas' }, { id: 'lr', name: 'PLANTA_Roles' }] });
  if (url.includes('/lists/lc/')) return j(items(certs));
  if (url.includes('/lists/lf/')) return j(items(process.env.FALSO_SIN_FIRMAS === '1' ? [] : firmas));   // el publicador sin permiso sobre PLANTA_Firmas
  if (url.includes('/lists/lr/')) return j(items(roles));
  return j({ id: 'sitio' });
};
