// node test/reglas.test.js — las reglas de la puerta contra los casos de la verificacion del plan.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { compuerta, siguienteFolio, avisoNeto, placaNormal, slug, rolDe, evaluarVigencia, lista, accionCorreccion, PUEDE, prealtaSinMovimiento, horaMexico, aIsoDia, fechaCorta, autoformatoFecha, plural, limpiar, paraPatch, tipoDeArchivo, lunesDe, sumarDias, esLoteDeLaApp, residuoDe, sufijoVerificacion, datosCertificado, urlVerificacion, toneladas, siguientePaso, yaCapturado, CORRIENTES, etiquetaCorriente, palabraCompuerta, subpasoDeRegla, PANTALLA_DE_REGLA, clienteDe, sha256Hex, textoHuellaPrealta, huellaPrealta, firmaAmparaPrealta, basesRecientes, clientesPrealta, fechaDePestana, claveMes, mesesPrealtas } from '../reglas.js';

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
    { let n = 0; assert.equal(sufijoVerificacion(a => a.fill(n++ ? 0 : 255)), 'aaaaaaaaaaaa', 'C-47: los bytes >= 248 se descartan y se piden otros'); }
    { let n = 0; assert.equal(sufijoVerificacion(a => a.fill(n++ ? 0 : 247)), '999999999999', 'C-47: 247 aun vale (31*8 = 248)'); }
    assert.equal(urlVerificacion('https://planta.minsaenergy.com/certificado/', { Title: 'CT-26-0001', Sufijo: suf }), `https://planta.minsaenergy.com/certificado/?f=CT-26-0001-${suf}`);
    assert.equal(urlVerificacion('https://planta.minsaenergy.com/certificado', { Title: 'CT-26-0001', Sufijo: suf }).startsWith('https://planta.minsaenergy.com/certificado/?f='), true, 'agrega la barra');
    assert.equal(urlVerificacion('https://x/', { Title: 'CT-26-0001' }), null, 'sin sufijo no hay URL (no se pinta QR)');
}
// v0.39.0: un certificado por GONDOLA. datosCertificado congela UN embarque cerrado, no la suma del programa.
{
    const p5 = { id: 5 };
    const cerrado = { id: 1, PreAltaId: 5, Etapa: 'cerrado', NetoKg: 21220, TaraHora: '2026-09-02T15:00:00Z',
        Title: 'E-26-00002', Manifiesto: 'MINSA/RME/002/2026', TicketBascula: 'B-004512' };
    const d = datosCertificado(p5, cerrado);
    assert.equal(d.ok, true); assert.equal(d.motivo, null);
    assert.equal(d.kg, 21220); assert.equal(d.folio, 'E-26-00002');
    assert.equal(d.manifiesto, 'MINSA/RME/002/2026'); assert.equal(d.ticket, 'B-004512');
    assert.equal(d.fechaRecepcion, '2026-09-02T15:00:00Z'); assert.equal(d.tipoBulto, 'gondola');
    // Sin ticket de bascula SI se puede emitir (el operador pudo no tenerlo): el papel sale sin ese renglon.
    assert.equal(datosCertificado(p5, { ...cerrado, TicketBascula: null }).ok, true);
    // Lo que NO se certifica, con su motivo.
    assert.equal(datosCertificado(p5, { ...cerrado, Etapa: 'anulado' }).motivo, 'la góndola está en anulado, no cerrada');
    assert.equal(datosCertificado(p5, { ...cerrado, Etapa: 'bruto', NetoKg: null }).motivo, 'la góndola está en bruto, no cerrada');
    assert.equal(datosCertificado(p5, { ...cerrado, NetoKg: 0 }).motivo, 'la góndola no tiene neto');
    assert.equal(datosCertificado(p5, { ...cerrado, PreAltaId: 7 }).motivo, 'la góndola es de otro programa');
    assert.equal(datosCertificado(null, cerrado).motivo, 'la góndola no trae programa');
    // El PreAltaId puede llegar como cadena desde Graph.
    assert.equal(datosCertificado(p5, { ...cerrado, PreAltaId: '5' }).ok, true);
    assert.equal(toneladas(41020), '41.02'); assert.equal(toneladas(1234567), '1,234.57'); assert.equal(toneladas(0), '0.00');
}

// Rediseño tanda 3 (v0.48.0): el siguiente paso de cada góndola, en texto, para la lista de Góndolas.
{
    assert.deepEqual(siguientePaso({ Etapa: 'compuerta', Compuerta: 'pasa' }), { paso: 3, texto: 'falta peso bruto', tono: 'acc' });
    assert.deepEqual(siguientePaso({ Etapa: 'compuerta', Compuerta: 'excepcion-comercial' }), { paso: 2, texto: 'espera autorización de gerencia', tono: 'warn' });
    assert.deepEqual(siguientePaso({ Etapa: 'compuerta', Compuerta: 'excepcion-comercial' }, true), { paso: 3, texto: 'autorizada · falta peso bruto', tono: 'acc' });
    // La descarga empieza con el bruto: la hora va en hora de México (15:42 UTC = 09:42 CST).
    assert.deepEqual(siguientePaso({ Etapa: 'bruto', BrutoHora: '2026-10-15T15:42:00Z' }), { paso: 4, texto: 'falta tara · descarga desde 09:42', tono: 'info' });
    assert.equal(siguientePaso({ Etapa: 'bruto' }).texto, 'falta tara');
    assert.equal(siguientePaso({ Etapa: 'cerrado' }).tono, 'ok');
    assert.equal(siguientePaso({ Etapa: 'rechazado' }).texto, 'no entró');
    assert.equal(siguientePaso({ Etapa: 'anulado' }).paso, 0);
    assert.equal(siguientePaso({ Etapa: 'descargando' }).texto, 'etapa descargando');   // la etapa retirada el 7-sep no revienta
    assert.equal(siguientePaso(null).texto, 'sin etapa');
}

// Rediseño tanda 5 (v0.50.0, decisión 11): gana el primero — qué dejó capturado la otra sesión, y quién.
{
    assert.equal(yaCapturado({ Etapa: 'bruto', BrutoKg: 41080, BrutoHora: '2026-10-15T17:12:00Z', Title: 'E-26-00043', _por: 'Daniel Hipólito' }),
        'ya tiene peso bruto: 41080 kg a las 11:12, folio E-26-00043; lo guardó Daniel Hipólito');
    assert.equal(yaCapturado({ Etapa: 'cerrado', TaraKg: 17620, NetoKg: 23460, TaraHora: '2026-10-15T18:40:00Z' }), 'ya se cerró a las 12:40: tara 17620 kg, neto 23460 kg');
    // La anulación dice quién por su columna (AnuladoPor, correo → nombre), no por lastModifiedBy.
    assert.equal(yaCapturado({ Etapa: 'anulado', AnuladoPor: 'g@x.mx', AnuladoMotivo: 'placa mal' }, c => (c === 'g@x.mx' ? 'Gerencia' : c)), 'la anuló Gerencia («placa mal»)');
    assert.equal(yaCapturado({ Etapa: 'anulado' }), 'la anuló otra sesión');
    assert.equal(yaCapturado({ Etapa: 'rechazado', Title: 'R-26-0004' }), 'quedó como rechazo R-26-0004');
    assert.equal(yaCapturado(null), 'está en otra etapa');
}

// C-55 (v0.52.0): un solo catálogo de corrientes y una sola traducción de la compuerta.
{
    assert.equal(CORRIENTES.length, 4);
    assert.equal(etiquetaCorriente('base-aceite'), 'Recorte base aceite');
    assert.equal(etiquetaCorriente('fluidos-base-agua'), 'Fluido agotado base agua');
    assert.equal(etiquetaCorriente('otra-cosa'), 'otra-cosa');   // renglón viejo fuera del catálogo: tal cual
    assert.equal(etiquetaCorriente(undefined), '');
    assert.deepEqual([palabraCompuerta('pasa').palabra, palabraCompuerta('rechazo-legal').palabra, palabraCompuerta('excepcion-comercial').palabra], ['Pasa', 'No entra', 'Espera']);
    assert.equal(palabraCompuerta('excepcion-comercial', true).corta, 'excepción ok');
    assert.equal(palabraCompuerta('excepcion-comercial', false).corta, 'espera');
    assert.equal(palabraCompuerta('pasa', true).corta, 'pasa');   // la firma solo cuenta en la excepción
    const r = compuerta({ ...base, carrier: { ...carrier, Corrientes: 'base-agua' } });
    assert.ok(r.hallazgos.find(h => h.regla === 'Corriente').detalle.includes('no ampara «Recorte base aceite» (ampara: Recorte base agua)'));
}

// C-56 (v0.52.0): cada nombre de regla que emite compuerta() —en verde y en falla— tiene pantalla en la tabla. Antes era un
// regex sobre el prefijo: una regla renombrada caía callada en la pantalla 2 y «Corregir lo capturado» llevaba a otra.
{
    const vencido = en(-5);
    const casos = [
        base,
        { ...base, prealta: null, carrier: null, unidad: null, chofer: null, manifiesto: '' },
        { ...base, prealta: { ...prealta, Estado: 'borrador', CarrierId: 99, UnidadesIds: '77', ChoferesIds: '88' }, corriente: '',
          carrier: { ...carrier, AutorizacionASEA: '', VigenciaASEA: vencido, CSFVigencia: vencido },
          unidad: { ...unidad, FolioOficio: '', PlacaPlana: 'OTRA', TarjetaVigencia: vencido, PolizaVigencia: vencido },
          chofer: { ...chofer, LicenciaVigencia: vencido } }
    ];
    const nombres = new Set(casos.flatMap(c => compuerta(c).hallazgos.map(h => h.regla)));
    nombres.add('Art. 79');   // lo agrega app.js (evaluarPuerta) como aviso
    assert.ok(nombres.size >= 17, `solo ${nombres.size} nombres: ${[...nombres]}`);
    for (const n of nombres) assert.ok(n in PANTALLA_DE_REGLA, `la regla «${n}» no tiene pantalla en PANTALLA_DE_REGLA`);
    for (const n of Object.keys(PANTALLA_DE_REGLA)) assert.ok(nombres.has(n), `PANTALLA_DE_REGLA trae «${n}», que ninguna regla emite`);
    assert.equal(subpasoDeRegla('CSF del carrier'), 1);
    assert.equal(subpasoDeRegla('Licencia del chofer'), 2);
    assert.equal(subpasoDeRegla('Corriente'), 3);
    assert.equal(subpasoDeRegla('Art. 79'), 3);
}

// Rediseno de Pre-altas, tanda 1 (v0.53.0): el cliente sale de la columna Cliente (esquema v8) o, en los renglones
// anteriores que la traen vacia, del inicio del Title CLIENTE-POZO-AÑO.
assert.equal(clienteDe({ Cliente: 'gsm', Title: 'LATINA-IXACHI 1-2026' }), 'GSM', 'la columna manda sobre el titulo');
assert.equal(clienteDe({ Cliente: '', Title: 'LATINA-IXACHI 1052-2026' }), 'LATINA', 'sin columna: el inicio del titulo');
assert.equal(clienteDe({ Title: 'CLIENTE DEMO-POZO 1-2026' }), 'CLIENTE DEMO');
assert.equal(clienteDe(null), '');

// C-66 (v0.58.0): basesRecientes y clientesPrealta, que vinieron de app.js para poder probarlas sin la E2E.
{
    const pa = [
        { id: 1, Estado: 'cerrada', Cliente: 'GSM', Corriente: 'base-aceite', Title: 'GSM-A-2026' },
        { id: 2, Estado: 'firmada', Cliente: '', Corriente: 'base-aceite', Title: 'gsm-B-2026' },   // sin columna: cliente del titulo, en mayusculas
        { id: 3, Estado: 'borrador', Cliente: 'LATINA', Corriente: 'base-agua', Title: 'LATINA-C-2026' },
        { id: 4, Estado: 'firmada', Cliente: 'GSM', Corriente: 'base-agua', Title: 'GSM-D-2026' },
        { id: 5, Estado: 'cerrada', Cliente: 'PEMEX', Corriente: 'base-aceite', Title: 'PEMEX-E-2026' },
        { id: 6, Estado: 'firmada', Cliente: 'ZETA', Corriente: 'base-aceite', Title: 'ZETA-F-2026' },
        { id: 7, Estado: 'firmada', Cliente: 'LATINA', Corriente: '', Title: 'LATINA-G-2026' }
    ];
    const ids = xs => xs.map(p => p.id);
    assert.deepEqual(ids(basesRecientes(pa)), [7, 6, 5], 'las tres mas nuevas, firmadas o cerradas');
    assert.deepEqual(ids(basesRecientes(pa, 10)), [7, 6, 5, 4, 2], 'borrador fuera; GSM|base-aceite una sola vez (la mas nueva, id 2)');
    assert.deepEqual(ids(basesRecientes([...pa].reverse(), 10)), [7, 6, 5, 4, 2], 'no depende del orden de entrada');
    assert.deepEqual(basesRecientes([]), []);
    assert.deepEqual(basesRecientes(undefined), [], 'sin lista cargada no truena');
    assert.equal(pa[0].id, 1, 'no reordena la lista que recibe');

    const cl = clientesPrealta([...pa, { id: 8, Estado: 'borrador', Cliente: '', Title: '' }]);
    assert.deepEqual(cl.map(c => [c.clave, c.n, c.ultimo.id]), [['GSM', 3, 4], ['LATINA', 2, 7], ['PEMEX', 1, 5], ['ZETA', 1, 6]],
        'por numero de programas, empate alfabetico; ultimo = el id mas alto; sin cliente no cuenta; los borradores si');
    assert.deepEqual(clientesPrealta(undefined), []);
}

// S-27 (v0.56.0): la huella de la firma de pre-alta. SHA-256 propio contra node:crypto; el orden de los ids no la cambia;
// cambiar una unidad, el carrier o la corriente la rompe; una firma sin huella (antes de v0.56.0) sigue valiendo.
for (const t of ['', 'abc', 'a'.repeat(55), 'a'.repeat(56), 'a'.repeat(64), 'Pozo=IXACHI ñ — ✓', 'x'.repeat(1000)])
    assert.equal(sha256Hex(t), createHash('sha256').update(t, 'utf8').digest('hex'), `sha256 de ${t.length} caracteres`);
const pf = { Title: 'GSM-IXACHI 15-2026', Cliente: 'GSM', CarrierId: 5, UnidadesIds: '12;3', ChoferesIds: '7', Corriente: 'base-aceite', Generador: 'GEN DEMO' };
const hf = huellaPrealta(pf), hoyF = '2026-09-30T18:00:00Z';
assert.ok(hf.startsWith('huella v2 ') && hf.length < 255, 'cabe en la columna de texto Motivo');
assert.ok(firmaAmparaPrealta(hf, { ...pf, UnidadesIds: '3; 12', CarrierId: '5', Notas: 'otra', GondolasEsperadas: 9 }, hoyF), 'orden de ids, tipo del id y notas no cuentan');
assert.ok(firmaAmparaPrealta(hf, pf), 'la v2 vale sin importar cuándo se creó la firma');
assert.ok(!firmaAmparaPrealta(hf, { ...pf, UnidadesIds: '3;12;99' }, hoyF), 'una unidad agregada tras firmar');
assert.ok(!firmaAmparaPrealta(hf, { ...pf, CarrierId: 6 }, hoyF), 'otro carrier tras firmar');
assert.ok(!firmaAmparaPrealta(hf, { ...pf, Corriente: 'base-agua' }, hoyF), 'otra corriente tras firmar');
// S-29 (v0.63.0): sin huella vale solo si SharePoint la creó antes del push de la v0.56.0; sin Created, falla cerrado.
assert.ok(firmaAmparaPrealta('', { ...pf, CarrierId: 6 }, '2026-09-20T12:00:00Z'), 'firma sin huella de antes de v0.56.0: transitorio');
assert.ok(!firmaAmparaPrealta('', pf, '2026-09-24T12:00:00Z'), 'firma sin huella creada después del corte: no ampara');
assert.ok(!firmaAmparaPrealta('', pf), 'firma sin huella y sin Created: no ampara');
// S-30 (v0.63.0): la v1 se podía engañar moviendo texto entre campos contiguos; la v2 no. La v1 vale solo antes de su corte.
{
    const a = { ...pf, Pozo: 'P1\nCorriente=base-aceite', Corriente: '' }, b = { ...pf, Pozo: 'P1', Corriente: 'base-aceite\nCorriente=' };
    assert.equal(textoHuellaPrealta(a).split('\n').length, textoHuellaPrealta(b).split('\n').length);
    assert.notEqual(huellaPrealta(a), huellaPrealta(b), 'v2: mover texto entre campos cambia la huella');
    const v1 = 'huella v1 ' + sha256Hex(textoHuellaPrealta(pf));
    assert.ok(firmaAmparaPrealta(v1, pf, '2026-09-24T12:00:00Z'), 'v1 de una firma de antes del corte: vale');
    assert.ok(!firmaAmparaPrealta(v1, pf, '2026-09-26T12:00:00Z'), 'v1 escrita después del corte: no ampara');
    assert.ok(!firmaAmparaPrealta(v1, { ...pf, CarrierId: 6 }, '2026-09-24T12:00:00Z'), 'v1 vieja pero con el renglón cambiado: no ampara');
}

// Pantalla principal de Pre-altas, tanda 2 (v0.61.0): los programas de una pestaña en bloques por mes.
{
    assert.equal(fechaDePestana({ FechaEstimada: '2026-09-30', FirmadaEl: '2026-09-01T18:00:00Z' }, 'borrador'), '2026-09-30');
    assert.equal(fechaDePestana({ FirmadaEl: '2026-09-01T18:00:00Z', CerradaEl: null }, 'cerrada'), '2026-09-01T18:00:00Z', 'cerrada sin sello cae en su firma');
    assert.equal(claveMes('2026-10-01'), '2026-10', 'fecha de calendario tal cual');
    assert.equal(claveMes('2026-10-01T00:00:00Z'), '2026-09', 'C-73: un sello a medianoche UTC es de reloj: las 18:00 del 30-sep en México');
    assert.equal(claveMes('2026-10-01T03:00:00Z'), '2026-09', 'la de reloj va en hora de Mexico: las 21:00 del 30-sep');
    assert.equal(claveMes(null), '');
    const ps = [
        { id: 1, FirmadaEl: '2026-08-20T17:00:00Z' },
        { id: 2, FirmadaEl: '2026-09-02T17:00:00Z' },
        { id: 3, FirmadaEl: null },
        { id: 4, FirmadaEl: '2026-09-15T17:00:00Z' },
        { id: 5, FirmadaEl: '2026-09-15T17:00:00Z' }
    ];
    const b = mesesPrealtas(ps, 'firmada');
    assert.deepEqual(b.map(x => x.mes), ['Septiembre 2026', 'Agosto 2026', 'Sin fecha']);
    assert.deepEqual(b.map(x => x.ps.map(p => p.id)), [[5, 4, 2], [1], [3]], 'la fecha mas nueva primero; empate por id');
    assert.deepEqual(mesesPrealtas([], 'firmada'), []);
}
