// Pruebas con datos ficticios. Ejecutar: node --test tests/*.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const ex = require('../js/extraccion.js');
const fmt = require('../js/format.js');

const docs = [
  { tipo: 'FE', numero: '100', monto: 1000 },
  { tipo: 'ND', numero: '200', monto: 50 },
];

test('asignar abona si el saldo no cubre el documento', () => {
  const r = ex.asignar([{ saldo: 400, periodo: 'actual' }], docs);
  assert.deepEqual(r.aplicaciones, [[{ docto: 'Abonar FEE 100', monto: 400 }]]);
  assert.deepEqual(r.sinCubrir.map((p) => [p.d.numero, p.resta]), [['100', 600], ['200', 50]]);
});

test('asignar paga varios documentos y deja saldo sin usar', () => {
  const r = ex.asignar([{ saldo: 2000, periodo: 'actual' }], docs);
  assert.deepEqual(r.aplicaciones[0], [{ docto: 'Pagar FEE 100', monto: 1000 }, { docto: 'Pagar ND N° 200', monto: 50 }]);
  assert.equal(r.sinCubrir.length, 0);
});

test('asignar reparte un documento entre dos IPE', () => {
  const r = ex.asignar([{ saldo: 600, periodo: 'anterior' }, { saldo: 500, periodo: 'actual' }], docs);
  assert.deepEqual(r.aplicaciones, [
    [{ docto: 'Abonar FEE 100', monto: 600 }],
    [{ docto: 'Pagar FEE 100', monto: 400 }, { docto: 'Pagar ND N° 200', monto: 50 }],
  ]);
});

test('asignar omite documentos sin monto y usa NDE en años anteriores', () => {
  const r = ex.asignar([{ saldo: 100, periodo: 'anterior' }], [{ tipo: 'FE', numero: '1', monto: null }, { tipo: 'ND', numero: '2', monto: 100 }]);
  assert.deepEqual(r.aplicaciones[0], [{ docto: 'Pagar NDE 2', monto: 100 }]);
});

test('gruposDelCliente busca por RUT, luego por razón social, años anteriores primero', () => {
  const grupos = [
    { rut: '10000000', dv: '8', razon: 'EMPRESA UNO S.A.', periodo: 'actual' },
    { rut: '10000000', dv: '8', razon: 'EMPRESA UNO S.A.', periodo: 'anterior' },
    { rut: '5', dv: '1', razon: 'OTRA', periodo: 'actual' },
  ];
  assert.deepEqual(ex.gruposDelCliente({ rut_cliente: '10.000.000-8' }, grupos).map((g) => g.periodo), ['anterior', 'actual']);
  assert.equal(ex.gruposDelCliente({ rut_cliente: '', razon_social: 'empresa uno s.a' }, grupos).length, 2);
  assert.equal(ex.gruposDelCliente({ rut_cliente: '9-9', razon_social: 'NADIE' }, grupos).length, 0);
});

test('avisosDocumentos advierte RUT, saldo distinto, faltantes e imagen cortada', () => {
  const g = { rut: '10000000', dv: '8' };
  const lectura = { contiene_documentos: true, rut_cliente: '999-9', saldo_a_favor: 10, documentos: [{ tipo: 'FE', numero: '1', monto: null }], imagen_cortada: true, observaciones: '' };
  const a = ex.avisosDocumentos(lectura, g, 20, [{ d: { tipo: 'FE', numero: '2' }, resta: 5 }], fmt.formatCLP);
  assert.equal(a.length, 5);
  const sinAvisos = { contiene_documentos: true, rut_cliente: '10000000-8', saldo_a_favor: 20, documentos: [], imagen_cortada: false, observaciones: '' };
  assert.deepEqual(ex.avisosDocumentos(sinAvisos, g, 20, [], fmt.formatCLP), []);
});

test('bancoSelector reconoce nombres y deja el resto en Otro', () => {
  assert.equal(ex.bancoSelector('Banco Internacional').banco, 'Banco Internacional');
  assert.equal(ex.bancoSelector('BANCO DEL ESTADO DE CHILE').banco, 'BancoEstado');
  assert.equal(ex.bancoSelector('Banco de Chile').banco, 'Banco de Chile');
  assert.equal(ex.bancoSelector('Banco de Crédito e Inversiones').banco, 'Banco BCI');
  assert.equal(ex.bancoSelector('Itaú').banco, 'Banco Itaú');
  assert.deepEqual(ex.bancoSelector('Banco Nuevo'), { banco: 'Otro', bancoOtro: 'Banco Nuevo' });
  assert.deepEqual(ex.bancoSelector(''), { banco: '', bancoOtro: '' });
});

test('aplicarBanco conserva lo no leído y excluye correos DGAC', () => {
  const actual = { tipoCuenta: 'Cuenta Corriente', numeroCuenta: '1', banco: 'Tenpo', bancoOtro: '', correos: 'a@b.cl' };
  const r = ex.aplicarBanco(actual, { tipo_cuenta: 'Cuenta Vista', numero_cuenta: '', banco: '', correos: ['X@Cliente.cl', 'y@dgac.gob.cl', 'x@cliente.cl', 'z@sub.dgac.cl'] });
  assert.deepEqual(r, { tipoCuenta: 'Cuenta Vista', numeroCuenta: '1', banco: 'Tenpo', bancoOtro: '', correos: 'x@cliente.cl' });
});

test('avisosBanco advierte beneficiario distinto y datos faltantes', () => {
  const g = { rut: '12345678', dv: '5' };
  const completa = { contiene_datos_bancarios: true, beneficiario: 'Inversiones', rut_beneficiario: '12.345.678-5', banco: 'X', tipo_cuenta: 'Cuenta Corriente', numero_cuenta: '1', observaciones: '' };
  assert.equal(ex.rutValido('12.345.678-5'), true);
  assert.deepEqual(ex.avisosBanco(completa, g), []);
  assert.match(ex.avisosBanco(completa, { rut: '1', dv: '9' })[0], /distinto del cliente/);
  assert.equal(ex.avisosBanco({ ...completa, numero_cuenta: '', banco: '', tipo_cuenta: '' }, g).length, 3);
});
