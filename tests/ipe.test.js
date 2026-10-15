// Pruebas con datos ficticios. Ejecutar: node --test tests/
const test = require('node:test');
const assert = require('node:assert/strict');
const fmt = require('../js/format.js');
const parser = require('../js/parser.js');
const memo = require('../js/memo.js');
const devoluciones = require('../js/devoluciones.js');
const resumen = require('../js/resumen.js');

const HEAD = ['', 'MIS APUNTES', 'Responsable', 'Concesión', 'C.F.', 'COD. AUX.', 'DV', 'RAZÓN SOCIAL', 'DOC.', 'NRO.', 'TIPO',
  'COMPROB.', 'FECHA', 'GLOSA AUXILIAR', '  DEBE  ', '  HABER  ', '  SALDO  ', 'XX', 'Justificación Ingreso anticipado', ' Gestión realizada ',
  'DOCTO A APLICAR', 'MONTO A APLICAR'];
const row = (estado, rut, dv, razon, doc, nro, fecha, saldo, xx, just = '', gest = '', docto = '', monto = '') =>
  ['', estado, 'Ana', '', 51, rut, dv, razon, doc, nro, 'I', 1, fecha, '', 0, saldo, saldo, xx, just, gest, docto, monto];

const AOA = [
  HEAD,
  row('APLICAR ', 11111111, 1, 'EMPRESA UNO SPA', 'EFE', 100, '05-01-2026', 1000, 'ACREEDOR', 'Pago duplicado', '1) Correo 2) Cliente autoriza', 'Pagar FEE 1', 5000),
  row('APLICAR', 11111111, 1, 'EMPRESA UNO SPA', 'NBCH', 200, '06-01-2026', 300, 'DEUDOR'),
  row('APLICAR', 22222222, 'k', 'EMPRESA DOS LTDA', 'EFE', 300, new Date(2026, 5, 30), 2500, 'ACREEDOR', '', '', 'Pagar FEE 2', '2.500'),
  row('DEVOLVER', 33333333, 3, 'EMPRESA TRES', 'EFE', 400, '07-07-2026', 40, 'ACREEDOR', 'x', '1) correo 2) se procederá a devolver'),
  row('DEVOLVER', 33333333, 3, 'EMPRESA TRES', 'EFE', 401, '07-07-2026', 10, 'DEUDOR'),
  row('correo enviado 01-08-2026', 44444444, 4, 'EMPRESA CUATRO', 'EFE', 500, '01-07-2026', 70, 'ACREEDOR', '', '1) correo'),
  row('????', 55555555, 5, 'EMPRESA CINCO', 'EFE', 600, '01-07-2026', 1, 'ACREEDOR'),
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
];

test('formato RUT, montos y fecha', () => {
  assert.equal(fmt.formatRut(96755500, 2), '96.755.500-2');
  assert.equal(fmt.formatRut('89428000', 'k'), '89.428.000-K');
  assert.equal(fmt.formatCLP(631142), '$ 631.142');
  assert.equal(fmt.toNumber(' $  7.926 '), 7926);
  assert.equal(fmt.toNumber(' $ - '), 0);
  assert.equal(fmt.toNumber('$ -24.570'), -24570);
  assert.equal(fmt.formatMontoLibre(701483), '701.483');
  assert.equal(fmt.formatMontoLibre('Saldo total'), 'Saldo total');
  assert.equal(fmt.fechaMemo(new Date(2026, 9, 4)), '04.OCT.2026');
  assert.deepEqual(fmt.parseFecha('21.01.2026'), new Date(2026, 0, 21));
  assert.deepEqual(fmt.parseFecha(46022), new Date(2025, 11, 31));
});

test('parser mapea columnas por nombre y omite filas vacías', () => {
  const { rows, tieneColumnasAplicar } = parser.parseAoa([['Reporte IPE'], ...AOA]);
  assert.equal(rows.length, 7);
  assert.ok(tieneColumnasAplicar);
  assert.equal(rows[0].fila, 3);
  assert.equal(rows[0].categoria, 'APLICAR');
  assert.equal(rows[0].rutFmt, '11.111.111-1');
  assert.equal(rows[1].saldoNeto, -300);
  assert.equal(rows[5].categoria, 'CORREO ENVIADO');
  assert.equal(rows[6].categoria, 'OTRO');
});

test('parser sin columnas de aplicación y con columnas reordenadas', () => {
  const idx = [7, 1, 5, 6, 8, 9, 16, 17];
  const aoa = AOA.slice(0, 3).map((r) => idx.map((i) => r[i]));
  const { rows, tieneColumnasAplicar } = parser.parseAoa(aoa);
  assert.equal(tieneColumnasAplicar, false);
  assert.equal(rows[0].razon, 'EMPRESA UNO SPA');
  assert.equal(rows[0].saldo, 1000);
});

test('parser exige columnas obligatorias', () => {
  assert.throws(() => parser.parseAoa([['RAZÓN SOCIAL', 'SALDO'], ['X', 1]]), /Faltan columnas/);
});

test('memos agrupan APLICAR por empresa y excluyen DEUDOR por defecto', () => {
  const { rows } = parser.parseAoa(AOA);
  const groups = memo.buildGroups(rows);
  assert.equal(groups.length, 2);
  const uno = groups.find((g) => g.razon === 'EMPRESA UNO SPA');
  assert.deepEqual(uno.rows.map((r) => r.incluir), [true, false]);
  const data = memo.memoData(uno, { fecha: new Date(2026, 9, 14), iniciales: 'AB/cd' });
  assert.equal(data.fecha, '14.OCT.2026');
  assert.equal(data.rut, '11.111.111-1');
  assert.deepEqual(data.items, [{ doc: 'EFE', nro: '100', monto: '$ 1.000', docto_aplicar: 'Pagar FEE 1', monto_aplicar: '$ 5.000' }]);
  assert.deepEqual(data.motivos, [{ br: false, doc: 'EFE', nro: '100', justificacion: 'Pago duplicado' }]);
  assert.deepEqual(data.notas, []);
  assert.match(memo.warnings(uno).join(), /DEUDOR/);
  assert.match(memo.warnings(uno).join(), /supera el saldo en \$ 4\.000/);
  const dos = groups.find((g) => g.razon === 'EMPRESA DOS LTDA');
  assert.match(memo.warnings(dos).join(), /sin justificación/);
  assert.equal(memo.fileName(dos), 'MEMO_IPE_EMPRESA_DOS_LTDA_22222222.docx');
});

test('varias aplicaciones descuentan el saldo y generan nota de diferencia (memo Nº 176)', () => {
  const aoa = [HEAD, row('APLICAR', 100300, 3, 'AEROLINEA X', 'DOL', 805, '15-09-2026', 11314, 'ACREEDOR', 'Saldo depósito',
    '', 'Pagar FEE 263576; Pagar FEE 262729', '2.005; 2327')];
  const [g] = memo.buildGroups(parser.parseAoa(aoa).rows);
  assert.deepEqual(g.rows[0].aplicaciones, [{ docto: 'Pagar FEE 263576', monto: 2005 }, { docto: 'Pagar FEE 262729', monto: 2327 }]);
  const data = memo.memoData(g, { fecha: new Date(2026, 9, 13) });
  assert.deepEqual(data.items.map((i) => [i.monto, i.docto_aplicar, i.monto_aplicar]), [
    ['$ 11.314', 'Pagar FEE 263576', '$ 2.005'],
    ['$ 9.309', 'Pagar FEE 262729', '$ 2.327'],
  ]);
  assert.deepEqual(data.notas, [{ doc: 'DOL', nro: '805', remanente: '$ 6.982' }]);
  assert.deepEqual(memo.warnings(g), []);

  // aplicar todo el saldo no genera nota
  g.rows[0].aplicaciones.push({ docto: 'Pagar FEE 3', monto: 6982 });
  assert.deepEqual(memo.memoData(g, { fecha: new Date() }).notas, []);
});

test('documentos con memo en el historial salen de pendientes', () => {
  const { rows } = parser.parseAoa(AOA);
  const uno = memo.buildGroups(rows).find((g) => g.razon === 'EMPRESA UNO SPA');
  const entry = memo.historialEntry(uno, memo.memoData(uno, { fecha: new Date(2026, 9, 14) }), new Date(2026, 9, 14, 10));
  assert.equal(entry.documentos.length, 1); // la fila DEUDOR no se incluyó
  const idx = memo.indiceHistorial([entry]);
  const pendientes = memo.buildGroups(rows, idx);
  // Saldo sin cambio: EMPRESA UNO solo conserva su fila DEUDOR, por lo que ya no tiene memo pendiente
  assert.deepEqual(pendientes.map((g) => g.razon), ['EMPRESA DOS LTDA']);
});

test('IPE con memo previo vuelve a pendientes si su saldo cambia (saldo dinámico)', () => {
  const mes = (saldo) => parser.parseAoa([HEAD, row('APLICAR', 77777777, 7, 'EMPRESA X', 'EFE', 9, '01-09-2026', saldo, 'ACREEDOR', 'Pago en exceso')]).rows;
  // Mes 1: saldo 100, memo aplica 90 → quedan 10
  const [g1] = memo.buildGroups(mes(100));
  g1.rows[0].aplicaciones = [{ docto: 'Pagar FEE 1', monto: 90 }];
  const historial = [memo.historialEntry(g1, memo.memoData(g1, { fecha: new Date(2026, 8, 30) }), new Date(2026, 8, 30))];
  let idx = memo.indiceHistorial(historial);

  // Mes 2, Contabilidad no ha procesado el memo: saldo sigue en 100 → oculto
  assert.equal(memo.buildGroups(mes(100), idx).length, 0);
  // Mes 2, memo procesado: saldo 10 → vuelve con referencia al memo previo
  const [g2] = memo.buildGroups(mes(10), idx);
  assert.equal(g2.rows[0].previo.saldo, 100);
  assert.equal(g2.rows[0].previo.aplicado, 90);

  // Mes 2 se hace memo por los 10 sin aplicar nada aún; mes 3 el cliente suma 20 → saldo 30 vuelve a aparecer
  historial.push(memo.historialEntry(g2, memo.memoData(g2, { fecha: new Date(2026, 9, 30) }), new Date(2026, 9, 30)));
  idx = memo.indiceHistorial(historial);
  assert.equal(memo.buildGroups(mes(10), idx).length, 0);
  const [g3] = memo.buildGroups(mes(30), idx);
  assert.equal(g3.rows[0].previo.saldo, 10); // compara contra el memo más reciente
});

test('devoluciones netean DEUDOR y exportan resumen/detalle', () => {
  const { rows } = parser.parseAoa(AOA);
  const groups = devoluciones.buildGroups(rows);
  assert.equal(groups.length, 1);
  assert.equal(groups[0].total, 30);
  const { resumen: res, detalle } = devoluciones.sheets(groups);
  assert.deepEqual(res[1].slice(0, 4), ['33.333.333-3', 'EMPRESA TRES', 2, 30]);
  assert.equal(res[res.length - 1][3], 30);
  assert.equal(detalle.length, 3);
  assert.equal(detalle[1][11], 'se procederá a devolver');
});

test('resumen: estados, antigüedad y alertas', () => {
  const { rows } = parser.parseAoa(AOA);
  const r = resumen.compute(rows, new Date(2026, 9, 14), true);
  assert.equal(r.cantidad, 7);
  assert.equal(r.empresas, 5);
  assert.equal(r.antiguedad.reduce((s, a) => s + a.cantidad, 0), 7);
  const tipos = r.alertas.map((a) => a.tipo);
  assert.ok(tipos.includes('Memo'));
  assert.ok(tipos.includes('Estado'));
  assert.ok(r.alertas.some((a) => a.tipo === 'Seguimiento' && a.razon === 'EMPRESA CUATRO'));
});

test('ultimaGestion toma el último paso numerado', () => {
  assert.equal(parser.ultimaGestion('1) Correo a cliente 2 ) Cliente responde. 12)Correo final'), 'Correo final');
  assert.equal(parser.ultimaGestion(''), '');
});
