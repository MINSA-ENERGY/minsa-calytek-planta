// node test/reglas.test.js — las reglas de la puerta contra los casos de la verificacion del plan.
import assert from 'node:assert/strict';
import { compuerta, siguienteFolio, avisoNeto, placaNormal, slug, rolDe, evaluarVigencia, lista, accionCorreccion, PUEDE, prealtaSinMovimiento, horaMexico, aIsoDia, fechaCorta, autoformatoFecha, plural, limpiar, paraPatch, tipoDeArchivo, lunesDe, sumarDias, esLoteDeLaApp, residuoDe, sufijoVerificacion, resumenCertificado, urlVerificacion, toneladas } from '../reglas.js';

const hoy = new Date('2026-10-15T12:00:00Z');
const en = dias => new Date(hoy.getTime() + dias * 86400000).toISOString();

// Datos de demostracion (manifiesto MINSA/RME/000/2026): placas 55XY9K / 66ZW3M, chofer Juan Perez Demo,
// carrier de demostracion, autorizacion 00-ASEA-T-RME-00-00 vigente a 2028-12-05.
const carrier = { id: 1, Title: 'TRANSPORTES DEMO SA DE CV', AutorizacionASEA: '00-ASEA-T-RME-00-00',
    VigenciaASEA: '2028-12-05T00:00:00Z', Corrientes: 'base-aceite; base-agua', CSFVigencia: en(200) };
const unidad = { id: 10, Title: '55XY9K', PlacaPlana: '66ZW3M', CarrierId: 1, CapacidadKg: 25000,
    FolioOficio: 'UGI-DEMO-0000-2025', TarjetaVigencia: en(300), PolizaVigencia: en(90) };
const chofer = { id: 20, Title: 'JUAN PEREZ DEMO', CarrierId: 1, LicenciaVigencia: en(60) };
const prealta = { id: 5, Estado: 'firmada', FirmadaPor: 'validador@example.invalid', CarrierId: 1,
    UnidadesIds: '10', ChoferesIds: '20', Corriente: 'base-aceite' };
const base = { prealta, carrier, unidad, chofer, placaTractor: '77-AN-5C', placaPlana: '66ZW3M',
    corriente: 'base-aceite', manifiesto: 'MINSA/RME/000/2026', avisoDias: 30, hoy };

// Caso 3 del plan: todo en orden -> pasa.
{
    const r = compuerta(base);
    assert.equal(r.resultado, 'pasa', JSON.stringify(r.hallazgos));
    assert.ok(r.hallazgos.every(h => h.clase !== 'legal' && h.clase !== 'comercial'));
}

// Caso 4: placa no transcrita -> rechazo legal, y el folio es R-, sin consumir E-.
{
    const r = compuerta({ ...base, unidad: null });
    assert.equal(r.resultado, 'rechazo-legal');
    assert.ok(r.hallazgos.find(h => h.regla === 'Placa' && h.clase === 'legal'));
    assert.equal(siguienteFolio('R', ['E-26-00003', 'R-26-0001'], hoy), 'R-26-0002');
    assert.equal(siguienteFolio('E', ['E-26-00003', 'R-26-0001', 'R-26-0002'], hoy), 'E-26-00004');
}

// Caso 5: poliza vencida -> excepcion comercial (dispensable con motivo y gerencia).
{
    const r = compuerta({ ...base, unidad: { ...unidad, PolizaVigencia: en(-3) } });
    assert.equal(r.resultado, 'excepcion-comercial');
    const h = r.hallazgos.find(x => x.regla === 'Póliza de la unidad');
    assert.equal(h.clase, 'comercial');
    assert.match(h.detalle, /vencida hace 3/);
}

// Decision 4: pre-alta sin firmar -> legal.
{
    const r = compuerta({ ...base, prealta: { ...prealta, Estado: 'borrador' } });
    assert.equal(r.resultado, 'rechazo-legal');
    assert.ok(r.hallazgos.find(h => h.regla === 'Pre-alta' && h.clase === 'legal'));
}

// Padron seccion 2: corriente no amparada -> legal (el caso MANC: solo base agua).
{
    const r = compuerta({ ...base, carrier: { ...carrier, Corrientes: 'base-agua' } });
    assert.equal(r.resultado, 'rechazo-legal');
    assert.ok(r.hallazgos.find(h => h.regla === 'Corriente' && h.clase === 'legal'));
}

// Autorizacion ASEA vencida -> legal; a 20 dias -> solo aviso.
{
    assert.equal(compuerta({ ...base, carrier: { ...carrier, VigenciaASEA: en(-1) } }).resultado, 'rechazo-legal');
    const r = compuerta({ ...base, carrier: { ...carrier, VigenciaASEA: en(20) } });
    assert.equal(r.resultado, 'pasa');
    assert.ok(r.hallazgos.find(h => h.regla === 'Autorización ASEA del carrier' && h.clase === 'aviso'));
}

// Placa de la plana distinta a la del padron -> legal (dos placas, dos veces).
{
    const r = compuerta({ ...base, placaPlana: '00XX00' });
    assert.equal(r.resultado, 'rechazo-legal');
}

// Chofer fuera de la pre-alta -> comercial.
{
    const r = compuerta({ ...base, prealta: { ...prealta, ChoferesIds: '99' } });
    assert.equal(r.resultado, 'excepcion-comercial');
}

// Folio: primer folio del anio, y no se reutiliza aunque haya huecos.
assert.equal(siguienteFolio('E', [], hoy), 'E-26-00001');
assert.equal(siguienteFolio('E', ['E-26-00001', 'E-26-00007'], hoy), 'E-26-00008');
assert.equal(siguienteFolio('E', ['E-25-00099'], hoy), 'E-26-00001');
assert.equal(siguienteFolio('L', [], hoy), 'L-26-001');
assert.equal(siguienteFolio('L', ['L-26-001', 'L-26-002'], hoy), 'L-26-003');

// Tolerancia del neto: boleta 13372 de Global Water (44,600 / 23,380 -> 21,220) en una gondola de 25 t.
assert.equal(avisoNeto(44600, 23380, 25000, { minFraccion: 0.3, maxFraccion: 1.25 }), null);
assert.match(avisoNeto(60000, 23380, 25000, { minFraccion: 0.3, maxFraccion: 1.25 }), /supera/);
assert.match(avisoNeto(24000, 23380, 25000, { minFraccion: 0.3, maxFraccion: 1.25 }), /menos del 30/);
assert.match(avisoNeto(20000, 23380, 25000, { minFraccion: 0.3, maxFraccion: 1.25 }), /no es mayor/);

// Utilerias.
assert.equal(placaNormal(' 55-xy-9k '), '55XY9K');
assert.deepEqual(lista(' a; b ;;c '), ['a', 'b', 'c']);
assert.equal(slug('Bruto E-26-00001 · góndola'), 'bruto-e-26-00001-gondola');
assert.equal(rolDe('Validador@example.invalid', [{ Title: 'validador@example.invalid', Rol: 'validador' }]), 'validador');
assert.equal(rolDe('nadie@example.invalid', []), 'lectura');
assert.equal(evaluarVigencia('x', null, 'legal', 30, hoy).ok, false);

// Corregir una captura (2026-09-05): sin folio se elimina; con folio se anula; anulado no se toca.
assert.equal(accionCorreccion({ Etapa: 'compuerta', Title: '' }), 'eliminar');
assert.equal(accionCorreccion({ Etapa: 'compuerta', Title: undefined, Compuerta: 'excepcion-comercial' }), 'eliminar');
assert.equal(accionCorreccion({ Etapa: 'rechazado', Title: 'R-26-0001' }), 'anular');
assert.equal(accionCorreccion({ Etapa: 'bruto', Title: 'E-26-00001' }), 'anular');
assert.equal(accionCorreccion({ Etapa: 'cerrado', Title: 'E-26-00001' }), 'anular');
assert.equal(accionCorreccion({ Etapa: 'anulado', Title: 'E-26-00001' }), null);
assert.equal(accionCorreccion(null), null);
// El folio de un anulado sigue contando: el siguiente no lo reutiliza.
assert.equal(siguienteFolio('E', ['E-26-00001', 'E-26-00002'], hoy), 'E-26-00003');
assert.ok(PUEDE.corregir('trazabilidad') && PUEDE.corregir('gerencia') && !PUEDE.corregir('validador') && !PUEDE.corregir('lectura'));

// Pre-alta firmada sin movimiento (v0.19.13): senala, no cierra.
{
    const pa = { id: 7, Estado: 'firmada', GondolasEsperadas: 3, FechaEstimada: en(-40), FirmadaEl: en(-45) };
    const emb = (dias, etapa = 'cerrado') => ({ PreAltaId: 7, Etapa: etapa, Arribo: en(dias) });
    assert.equal(prealtaSinMovimiento({ ...pa, Estado: 'borrador' }, [], 15, hoy), null, 'un borrador no se juzga');
    assert.equal(prealtaSinMovimiento({ ...pa, FechaEstimada: en(5), FirmadaEl: en(-1) }, [], 15, hoy), null, 'recien firmada y con fecha futura: nada');
    assert.match(prealtaSinMovimiento(pa, [], 15, hoy).motivo, /sin un solo arribo en 40 días/, 'sin arribos: cuenta desde la fecha estimada (la mas reciente)');
    assert.equal(prealtaSinMovimiento(pa, [emb(-3)], 15, hoy), null, 'arribo hace 3 dias: viva');
    assert.match(prealtaSinMovimiento(pa, [emb(-20)], 15, hoy).motivo, /sin arribos desde hace 20 días/);
    assert.equal(prealtaSinMovimiento(pa, [emb(-20), emb(-2, 'anulado')], 15, hoy).motivo.includes('20'), true, 'un anulado no cuenta como movimiento');
    assert.match(prealtaSinMovimiento(pa, [emb(-1), emb(-1), emb(-1)], 15, hoy).motivo, /ya recibió sus 3/, 'completa aunque sea reciente');
    assert.equal(prealtaSinMovimiento({ ...pa, FechaEstimada: null, FirmadaEl: null }, [], 15, hoy), null, 'sin fechas no se juzga');
}

// C-10 (v0.22.0): la hora de Mexico nunca imprime «24:05» a medianoche (hourCycle h23, no hour12:false) y los tres modos
// salen del mismo helper. 06:05Z = 00:05 en Mexico (UTC-6 todo el anio desde 2022).
{
    const medianoche = '2026-03-10T06:05:00Z';
    assert.equal(horaMexico(medianoche, 'hora'), '00:05');
    assert.ok(horaMexico(medianoche).endsWith('00:05') && horaMexico(medianoche).includes('10/03/2026'), horaMexico(medianoche));
    assert.ok(!horaMexico(medianoche, 'completa').includes('24:') && horaMexico(medianoche, 'completa').includes('00:05:00'), horaMexico(medianoche, 'completa'));
    assert.equal(horaMexico('2026-03-10T18:30:00Z', 'hora'), '12:30', 'mediodia sin PM');
    assert.equal(horaMexico(''), '—');
    assert.equal(horaMexico('no es fecha'), 'no es fecha');
}

console.log('reglas: ok');

// C-25 (v0.28.0): la frontera de fechas que teclea la caseta. Antes vivia en app.js y la E2E solo pegaba ISO.
{
    const dia = iso => iso.slice(0, 10);
    // Ano de dos cifras = 20aa (fotos de Carlos, 2026-09-06): «16/03/26» es el 16 de marzo de 2026.
    assert.equal(dia(new Date(aIsoDia('16/03/26')).toLocaleDateString('sv-SE')), '2026-03-16');
    assert.equal(dia(new Date(aIsoDia('16/03/2026')).toLocaleDateString('sv-SE')), '2026-03-16');
    assert.equal(dia(new Date(aIsoDia('16.03.2026')).toLocaleDateString('sv-SE')), '2026-03-16');
    // Lo que teclea la gente («05092026») pasa por el autoformato y de ahi a aIsoDia.
    assert.equal(autoformatoFecha('05092026'), '05/09/2026');
    assert.equal(autoformatoFecha('0509'), '05/09');
    assert.equal(autoformatoFecha('05'), '05');
    assert.equal(autoformatoFecha('2026-09-05'), '05/09/2026');   // ISO pegado se muestra dd/mm/aaaa
    assert.equal(dia(new Date(aIsoDia(autoformatoFecha('05092026'))).toLocaleDateString('sv-SE')), '2026-09-05');
    // ISO pegado (pruebas y pegados) pasa igual.
    assert.equal(dia(new Date(aIsoDia('2026-09-05')).toLocaleDateString('sv-SE')), '2026-09-05');
    // Vacio = null; nunca una fecha adivinada.
    assert.equal(aIsoDia(''), null); assert.equal(aIsoDia(null), null); assert.equal(aIsoDia('   '), null);
    // Fecha que no existe: lanza con el mensaje que ve quien captura.
    assert.throws(() => aIsoDia('31/04/2026'), /no existe/);
    assert.throws(() => aIsoDia('29/02/2027'), /no existe/);
    assert.throws(() => aIsoDia('00/01/2026'), /no existe/);
    assert.equal(dia(new Date(aIsoDia('29/02/2028')).toLocaleDateString('sv-SE')), '2028-02-29');   // bisiesto si existe
    // Cualquier otra cosa: «no valida», con la forma esperada.
    assert.throws(() => aIsoDia('hoy'), /no v\u00e1lida.*dd\/mm\/aaaa/);
    assert.throws(() => aIsoDia('2026/09/05'), /no v\u00e1lida/);
    assert.throws(() => aIsoDia('5-9'), /no v\u00e1lida/);
    // fechaCorta: ISO -> dd/mm/aaaa; vacio -> raya; lo que no es ISO sale tal cual.
    assert.equal(fechaCorta('2026-09-05T18:00:00.000Z'), '05/09/2026');
    assert.equal(fechaCorta(''), '\u2014'); assert.equal(fechaCorta(null), '\u2014');
    assert.equal(fechaCorta('05/09/2026'), '05/09/2026');
    // plural (U-36), limpiar (POST) y paraPatch (PATCH) — las tres puras que viajaban con app.js.
    assert.equal(plural(1, 'g\u00f3ndola'), '1 g\u00f3ndola'); assert.equal(plural(3, 'g\u00f3ndola'), '3 g\u00f3ndolas');
    assert.equal(plural(2, 'anulado o rechazado', 'anulados o rechazados'), '2 anulados o rechazados');
    assert.deepEqual(limpiar({ a: 1, b: '', c: null, d: undefined, e: 0, f: false }), { a: 1, e: 0, f: false });
    assert.deepEqual(paraPatch({ a: 1, b: '', c: null, d: undefined, e: 0 }), { a: 1, b: null, c: null, d: null, e: 0 });
    // tipoDeArchivo (v0.33.0, Archivos): por nombre, con lo que la app y /archivar-calytek escriben de verdad.
    assert.equal(tipoDeArchivo('2026-09-05_CALYTEK_Foto_bruto-e-26-00001-44600-kg-01.jpg'), 'foto');
    assert.equal(tipoDeArchivo('_lote.json'), 'lote');
    assert.equal(tipoDeArchivo('2026-09-20_CALYTEK_Ticket_E-26-00012_XRT-402-A.pdf'), 'ticket');
    assert.equal(tipoDeArchivo('2026-09-20_MINSA_Manifiesto_RME-147-2026.pdf'), 'manifiesto');
    assert.equal(tipoDeArchivo('2025-03-14_ASEA_Oficio-UGI-DEMO-0212-2025_Transportes-Demo.pdf'), 'oficio');
    assert.equal(tipoDeArchivo('2026-01-09_SAT_CSF_Transportes-Demo.pdf'), 'csf');
    assert.equal(tipoDeArchivo('Constancia de Situaci\u00f3n Fiscal.pdf'), 'csf');
    assert.equal(tipoDeArchivo('2026-09-01_CALYTEK_Recibo_telmex.pdf'), 'otro');
    assert.equal(tipoDeArchivo('foto-del-indicador.PNG'), 'foto');
    assert.equal(tipoDeArchivo(''), 'otro'); assert.equal(tipoDeArchivo(null), 'otro');
    // C-32 (v0.34.0): el ORDEN de los patrones decide el filtro «Tipo» cuando un nombre casa dos: ticket > manifiesto > csf > oficio > foto.
    assert.equal(tipoDeArchivo('2026-09-20_CALYTEK_Ticket_E-26-00012_con-Manifiesto.pdf'), 'ticket');
    assert.equal(tipoDeArchivo('2026-01-09_ASEA_Oficio_anexo-CSF_Transportes-Demo.pdf'), 'csf');
    assert.equal(tipoDeArchivo('2025-03-14_ASEA_Oficio-UGI-0212-2025_escaneado.jpg'), 'oficio');
    // C-30 (v0.34.0): un solo «lunes» (sobre la fecha ya cortada en hora de México; aritmética UTC, sin zona del dispositivo).
    assert.equal(lunesDe('2026-09-20'), '2026-09-14');   // domingo → el lunes anterior
    assert.equal(lunesDe('2026-09-14'), '2026-09-14');   // lunes → él mismo
    assert.equal(lunesDe('2026-09-15'), '2026-09-14');
    assert.equal(lunesDe('2026-01-01'), '2025-12-29');   // cruza el año
    assert.equal(sumarDias('2026-09-14', 6), '2026-09-20'); assert.equal(sumarDias('2026-03-01', -1), '2026-02-28');
    // S-15 (v0.34.0): en el buzón solo se listan los lotes de la app (AAAA-MM-DD_<etiqueta>_…), no lo que Carlos deposita.
    assert.equal(esLoteDeLaApp('2026-09-20_Embarque_bruto-e-26-00012-44600-kg', 'Embarque'), true);
    assert.equal(esLoteDeLaApp('2026-09-20_Embarque_bruto-e-26-00012-44600-kg 2', 'Embarque'), true);   // el «2» que Graph agrega al chocar
    assert.equal(esLoteDeLaApp('2026-09-18_MINSA_Contrato_Transportes-Demo.pdf', 'Embarque'), false);
    assert.equal(esLoteDeLaApp('Embarque_sin-fecha', 'Embarque'), false); assert.equal(esLoteDeLaApp('', 'Embarque'), false);
}

// v0.35.0: certificado de tratamiento (CT-AA-NNNN, un certificado por programa).
assert.equal(siguienteFolio('C', [], hoy), 'CT-26-0001');
assert.equal(siguienteFolio('C', ['CT-26-0003', 'E-26-00009', 'CT-25-0040'], hoy), 'CT-26-0004');
assert.equal(siguienteFolio('E', ['CT-26-0003', 'E-26-00009'], hoy), 'E-26-00010', 'el CT no contamina el consecutivo E');
assert.equal(residuoDe('base-aceite'), 'RECORTES DE PERFORACIÓN');   // v0.38.0: sin la base (Carlos)
assert.equal(residuoDe('fluidos-base-agua'), 'RECORTES DE PERFORACIÓN');
assert.equal(residuoDe(''), 'RECORTES DE PERFORACIÓN');
{
    const suf = sufijoVerificacion(a => { for (let i = 0; i < a.length; i++) a[i] = (i * 37 + 5) % 256; return a; });
    assert.match(suf, /^[a-z2-9]{12}$/, 'sufijo de 12 sin ambiguos: ' + suf);
    assert.equal(sufijoVerificacion(a => a.fill(0)), 'aaaaaaaaaaaa', 'determinista con el aleatorio inyectado');
    assert.doesNotMatch(sufijoVerificacion(a => a.fill(255)), /[01lIO]/);
    assert.equal(urlVerificacion('https://planta.minsaenergy.com/certificado/', { Title: 'CT-26-0001', Sufijo: suf }), `https://planta.minsaenergy.com/certificado/?f=CT-26-0001-${suf}`);
    assert.equal(urlVerificacion('https://planta.minsaenergy.com/certificado', { Title: 'CT-26-0001', Sufijo: suf }).startsWith('https://planta.minsaenergy.com/certificado/?f='), true, 'agrega la barra');
    assert.equal(urlVerificacion('https://x/', { Title: 'CT-26-0001' }), null, 'sin sufijo no hay URL (no se pinta QR)');
}
{
    const p5 = { id: 5 };
    const embs = [
        { id: 1, PreAltaId: 5, Etapa: 'cerrado', NetoKg: 21220, TaraHora: '2026-09-02T15:00:00Z', Title: 'E-26-00002', Manifiesto: 'MINSA/RME/002/2026' },
        { id: 2, PreAltaId: 5, Etapa: 'cerrado', NetoKg: 19800, TaraHora: '2026-09-01T15:00:00Z', Title: 'E-26-00001', Manifiesto: 'MINSA/RME/001/2026' },
        { id: 3, PreAltaId: 5, Etapa: 'anulado', NetoKg: 50000, TaraHora: '2026-09-03T15:00:00Z', Title: 'E-26-00003' },
        { id: 4, PreAltaId: 5, Etapa: 'bruto', BrutoKg: 40000, Title: 'E-26-00004' },
        { id: 6, PreAltaId: 7, Etapa: 'cerrado', NetoKg: 1000, TaraHora: '2026-09-01T10:00:00Z', Title: 'E-26-00005' },
        { id: 7, PreAltaId: '5', Etapa: 'cerrado', NetoKg: 0, TaraHora: '2026-09-04T10:00:00Z', Title: 'E-26-00006' }
    ];
    const r = resumenCertificado(p5, embs);
    assert.deepEqual(r.folios, ['E-26-00001', 'E-26-00002'], 'solo cerrados con neto, del programa, en orden de cierre');
    assert.equal(r.kg, 41020);
    assert.deepEqual(r.manifiestos, ['MINSA/RME/001/2026', 'MINSA/RME/002/2026']);
    assert.equal(r.primerCierre, '2026-09-01T15:00:00Z'); assert.equal(r.ultimoCierre, '2026-09-02T15:00:00Z');
    assert.deepEqual(resumenCertificado({ id: 99 }, embs), { embarques: [], kg: 0, primerCierre: null, ultimoCierre: null, folios: [], manifiestos: [] });
    assert.equal(toneladas(41020), '41.02'); assert.equal(toneladas(1234567), '1,234.57'); assert.equal(toneladas(0), '0.00');
}
