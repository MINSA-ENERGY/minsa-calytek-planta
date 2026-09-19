// node test/reglas.test.js — las reglas de la puerta contra los casos de la verificacion del plan.
import assert from 'node:assert/strict';
import { compuerta, siguienteFolio, avisoNeto, placaNormal, slug, rolDe, evaluarVigencia, lista, accionCorreccion, PUEDE, prealtaSinMovimiento, horaMexico } from '../reglas.js';

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
    const h = r.hallazgos.find(x => x.regla === 'Poliza de la unidad');
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
    assert.ok(r.hallazgos.find(h => h.regla === 'Autorizacion ASEA del carrier' && h.clase === 'aviso'));
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
