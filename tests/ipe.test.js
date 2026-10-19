// Pruebas con datos ficticios. Ejecutar: node --test tests/*.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const fmt = require('../js/format.js');
const parser = require('../js/parser.js');
const glosas = require('../js/glosas.js');
const movimientos = require('../js/movimientos.js');
const memo = require('../js/memo.js');
const devoluciones = require('../js/devoluciones.js');
const resumen = require('../js/resumen.js');

// Estructura final del Excel: A clave, B TIPO DE IPE, C MI GESTION, ...
const HEAD = ['', 'TIPO DE IPE', 'MI GESTION', 'Responsable', 'Concesión', 'C.F.', 'COD. AUX.', 'DV', 'RAZÓN SOCIAL', 'DOC.', 'NRO.', 'TIPO',
  'COMPROB.', 'FECHA', 'GLOSA AUXILIAR', '  DEBE  ', '  HABER  ', '  SALDO  ', 'XX', 'Justificación Ingreso anticipado', ' Gestión realizada '];
const row = (tipo, estado, rut, dv, razon, doc, nro, fecha, saldo, xx, just = '', gest = '', comprob = 1) =>
  ['', tipo, estado, 'Ana', '', 51, rut, dv, razon, doc, nro, 'I', comprob, fecha, '', 0, saldo, saldo, xx, just, gest];

const AOA = [
  HEAD,
  row('Pago en exceso', 'APLICAR ', 11111111, 1, 'EMPRESA UNO SPA', 'EFE', 100, '05-01-2025', 100, 'ACREEDOR', 'Pago duplicado', '1) Correo'),
  row('', 'APLICAR', 22222222, 'k', 'EMPRESA DOS LTDA', 'NBCH', 440146, '22-01-2026', 542760, 'ACREEDOR', 'NCE N°18549', '', 10),
  row('', 'APLICAR', 22222222, 'k', 'EMPRESA DOS LTDA', 'NBCH', 440146, '22-01-2026', 529798, 'DEUDOR', 'NCE N°18549', '', 11),
  row('Pago duplicado', 'DEVOLVER', 33333333, 3, 'EMPRESA TRES', 'EFE', 400, '07-07-2026', 40, 'ACREEDOR', 'x', '1) correo'),
  row('', 'correo enviado 01-08-2026', 44444444, 4, 'EMPRESA CUATRO', 'EFE', 500, '01-07-2026', 70, 'ACREEDOR', '', '1) correo'),
  row('', '', 55555555, 5, 'EMPRESA CINCO', 'EFE', 600, '01-03-2026', 1, 'ACREEDOR'),
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
];
const HOY = new Date(2026, 9, 19);
// En el Word los montos llevan espacio no separable entre "$" y la cifra.
const sp = (v) => JSON.parse(JSON.stringify(v).replace(/\u00a0/g, ' '));
const leer = () => parser.consolidar(parser.parseAoa(AOA).rows);

test('formato: RUT, montos, fechas y banderas', () => {
  assert.equal(fmt.formatRut(96755500, 2), '96.755.500-2');
  assert.equal(fmt.formatRut('89428000', 'k'), '89.428.000-K');
  assert.equal(fmt.formatCLP(631142), '$ 631.142');
  assert.equal(fmt.toNumber(' $  7.926 '), 7926);
  assert.equal(fmt.toNumber(' $ - '), 0);
  assert.equal(fmt.toNumber('$ -24.570'), -24570);
  assert.equal(fmt.fechaMemo(new Date(2026, 9, 4)), '04.OCT.2026');
  assert.equal(fmt.fechaMesAnio(new Date(2026, 9, 4)), 'OCT.2026');
  assert.deepEqual(fmt.parseFecha('21.01.2026'), new Date(2026, 0, 21));
  assert.equal(fmt.bandera(new Date(2025, 11, 31), HOY), 'roja');
  assert.equal(fmt.bandera(new Date(2026, 0, 1), HOY), 'amarilla');
  assert.equal(fmt.bandera(new Date(2026, 5, 30), HOY), 'amarilla');
  assert.equal(fmt.bandera(new Date(2026, 6, 1), HOY), 'verde');
  assert.equal(fmt.bandera(null, HOY), null);
});

test('parser: estructura final con TIPO DE IPE y MI GESTION', () => {
  const { rows, tieneEstado, tieneTipo } = parser.parseAoa([['Reporte IPE'], ...AOA]);
  assert.ok(tieneEstado && tieneTipo);
  assert.equal(rows.length, 6);
  assert.equal(rows[0].fila, 3);
  assert.equal(rows[0].categoria, 'APLICAR');
  assert.equal(rows[0].tipoIpe, 'Pago en exceso');
  assert.equal(rows[0].rutFmt, '11.111.111-1');
  assert.equal(rows[4].categoria, 'CORREO ENVIADO');
  assert.equal(rows[5].categoria, 'SIN ESTADO');
});

test('parser: acepta MIS APUNTES (formato anterior) y exige columnas obligatorias', () => {
  const viejo = AOA.map((r, i) => (i === 0 ? r.map((h) => (h === 'MI GESTION' ? 'MIS APUNTES' : h)) : r));
  assert.equal(parser.parseAoa(viejo).rows[0].categoria, 'APLICAR');
  assert.throws(() => parser.parseAoa([['RAZÓN SOCIAL', 'SALDO'], ['X', 1]]), /Faltan columnas/);
});

test('consolida DEUDOR y ACREEDOR del mismo IPE en un solo registro acreedor', () => {
  const rows = leer();
  assert.equal(rows.length, 5);
  const dos = rows.find((r) => r.razon === 'EMPRESA DOS LTDA');
  assert.equal(dos.saldo, 12962);
  assert.equal(dos.naturaleza, 'ACREEDOR');
  assert.equal(dos.comprob, '10');
  assert.deepEqual(dos.consolidado, { filas: [3, 4], acreedor: 542760, deudor: 529798 });
});

test('glosas: tipos del documento GLOSAS y datos a completar resaltados', () => {
  assert.equal(glosas.GLOSAS.length, 9);
  assert.equal(glosas.porTipo('PAGO EN EXCESO').id, 'pago-exceso');
  assert.equal(glosas.porTipo('Deposito sin identificar').id, 'deposito-sin-identificar');
  assert.equal(glosas.porTipo('Error en pago').id, 'error-pago-regularizacion');
  assert.equal(glosas.porTipo('Error en pago 2').id, 'error-pago-fee');
  assert.equal(glosas.porTipo('Error en pago 3').id, 'error-pago-recibo');
  assert.equal(glosas.porTipo('otro'), null);
  assert.equal(glosas.porTipo(''), null);
  assert.equal(glosas.pendientes(glosas.porTipo('Pago en exceso').texto), 7);
  assert.deepEqual(glosas.partes('Factura N°[[N° factura]] pagada.'), [{ n: true, t: 'Factura N°' }, { m: true, t: '[N° factura]' }, { n: true, t: ' pagada.' }]);
  assert.deepEqual(glosas.partes('a <completar por el usuario>'), [{ n: true, t: 'a ' }, { m: true, t: '[completar]' }]);
});

test('memo: MOTIVO desde TIPO DE IPE, fecha mes.año y aplicación parcial que deja saldo pendiente', () => {
  const rows = leer();
  let [uno] = memo.buildGroups(rows).filter((g) => g.razon === 'EMPRESA UNO SPA');
  assert.equal(uno.rows[0].glosaId, 'pago-exceso');
  assert.match(uno.rows[0].texto, /^Corresponde a pago en exceso/);
  uno.rows[0].aplicaciones = [{ docto: 'Pagar FEE 1', monto: 20 }];
  const fecha = new Date(2026, 9, 14);
  const data = memo.memoData(uno, { fecha, iniciales: 'AB/cd' });
  assert.equal(data.fecha, 'OCT.2026');
  assert.deepEqual(data.items, [{ doc: 'EFE', nro: '100', monto: '$ 100', docto_aplicar: 'Pagar FEE 1', monto_aplicar: '$ 20' }]);
  assert.deepEqual(sp(data.notas), [{ doc: 'EFE', nro: '100', remanente: '$ 80' }]);
  const historial = [memo.historialEntry(uno, data, new Date(2026, 9, 14, 10), { fecha })];
  historial[0].numero = '310';

  // Mismo Excel: Juan sigue pendiente por $80; el nuevo MOTIVO menciona la aplicación anterior
  let idx = movimientos.indice(historial, []);
  [uno] = memo.buildGroups(rows, idx).filter((g) => g.razon === 'EMPRESA UNO SPA');
  assert.equal(uno.rows[0].saldoInicial, 80);
  uno.rows[0].aplicaciones = [{ docto: 'Pagar FEE 2', monto: 10 }];
  const data2 = memo.memoData(uno, { fecha: new Date(2026, 9, 20) });
  const texto2 = data2.motivos[0].justificacion.map((p) => p.t).join('');
  assert.match(texto2, /Asimismo, con fecha \[día\]\.OCT\.2026, se aplicaron \$\s20, según MEMO N°310, de fecha \[día\]\.OCT\.2026\.$/);
  assert.equal(data2.motivos[0].justificacion.filter((p) => p.t === '[día]').length, 2);
  historial[0].numero = '';
  const sinNumero = memo.memoData(memo.buildGroups(rows, movimientos.indice(historial, [])).find((g) => g.razon === 'EMPRESA UNO SPA'), { fecha: HOY });
  assert.ok(sinNumero.motivos[0].justificacion.some((p) => p.m && p.t === '[N° memo]'));
  assert.equal(data2.items[0].monto, '$ 80');
  historial.push(memo.historialEntry(uno, data2, new Date(2026, 9, 20, 10), { fecha: new Date(2026, 9, 20) }));
  idx = movimientos.indice(historial, []);
  assert.equal(memo.buildGroups(rows, idx).find((g) => g.razon === 'EMPRESA UNO SPA').rows[0].saldoInicial, 70);

  // Mes siguiente: el Excel ya refleja ambas aplicaciones (saldo 70) → pendiente 70, no 40
  const mesSig = rows.map((r) => (r.razon === 'EMPRESA UNO SPA' ? { ...r, saldo: 70, saldoNeto: 70 } : r));
  assert.equal(memo.buildGroups(mesSig, idx).find((g) => g.razon === 'EMPRESA UNO SPA').rows[0].saldoInicial, 70);

  // Aplicado el total, el IPE sale de pendientes
  const uno3 = memo.buildGroups(mesSig, idx).find((g) => g.razon === 'EMPRESA UNO SPA');
  uno3.rows[0].aplicaciones = [{ docto: 'Pagar FEE 3', monto: 70 }];
  historial.push(memo.historialEntry(uno3, memo.memoData(uno3, { fecha: HOY }), new Date(2026, 10, 2), { fecha: HOY }));
  assert.equal(memo.buildGroups(mesSig, movimientos.indice(historial, [])).filter((g) => g.razon === 'EMPRESA UNO SPA').length, 0);
});

test('memo: sin TIPO DE IPE usa la justificación del Excel', () => {
  const dos = memo.buildGroups(leer()).find((g) => g.razon === 'EMPRESA DOS LTDA');
  assert.equal(dos.rows[0].glosaId, '');
  assert.equal(dos.rows[0].texto, 'NCE N°18549');
  assert.equal(memo.textoInicial({ tipoIpe: '', justificacion: '' }).texto, '[[motivo]]');
  assert.equal(dos.rows[0].saldoInicial, 12962);
});

test('devolución: observaciones por TIPO DE IPE, datos bancarios, devolución parcial e historial', () => {
  const rows = leer();
  const [tres] = devoluciones.buildGroups(rows);
  assert.equal(tres.razon, 'EMPRESA TRES');
  assert.equal(tres.rows[0].montoDevolver, 40);
  assert.match(tres.rows[0].texto, /^Corresponde al pago duplicado/);
  tres.rows[0].montoDevolver = 30;
  Object.assign(tres.banco, { referencia: 'Correo de fecha 01.10.2026', numeroCuenta: '123', banco: 'Otro', bancoOtro: 'Banco Prueba', correos: 'a@x.cl\nb@x.cl' });
  const data = devoluciones.devolucionData(tres, { fecha: HOY, iniciales: 'ERS/yaa' });
  assert.equal(data.fecha, 'OCT.2026');
  assert.equal(data.monto_total, '$ 30');
  assert.deepEqual(data.items[0].tipo_docto, 'EFE 400');
  assert.deepEqual(data.correos, [{ primero: true, correo: 'a@x.cl' }, { primero: false, correo: 'b@x.cl' }]);
  assert.deepEqual(data.banco, [{ n: true, t: 'Banco Prueba' }]);
  assert.deepEqual(data.referencia, [{ n: true, t: 'Correo de fecha 01.10.2026' }]);
  assert.ok(data.items[0].observaciones.some((p) => p.m));
  assert.deepEqual(devoluciones.warnings(tres), []);
  const entry = devoluciones.historialEntry(tres, data, new Date(2026, 9, 19, 9), { fecha: HOY });
  assert.equal(entry.documentos[0].devuelto, 30);
  const idx = movimientos.indice([], [entry]);
  const [pend] = devoluciones.buildGroups(rows, idx);
  assert.equal(pend.rows[0].saldoInicial, 10);
  assert.match(memo.textoCompleto(pend.rows[0]), /se devolvieron \$\s30, según OF\. \(O\) N°\[\[N° oficio\]\]/);
});

test('devolución: sin datos bancarios deja marcas para completar', () => {
  const [tres] = devoluciones.buildGroups(leer());
  const data = devoluciones.devolucionData(tres, { fecha: HOY });
  assert.deepEqual(data.referencia.filter((p) => p.m).map((p) => p.t), ['[correo del cliente]', '[fecha]', '[hora]']);
  assert.deepEqual(data.banco, [{ m: true, t: '[banco]' }]);
  assert.ok(devoluciones.BANCOS.includes('BancoEstado'));
  assert.ok(devoluciones.warnings(tres).some((w) => w.includes('REF.')));
});

test('resumen: empresas total, en memos, en devoluciones, pendientes y listado de pendientes', () => {
  const rows = leer();
  const r = resumen.compute(rows, memo.buildGroups(rows), devoluciones.buildGroups(rows), HOY);
  assert.equal(r.empresasTotal, 5);
  assert.equal(r.empresasMemos, 2);
  assert.equal(r.empresasDevoluciones, 1);
  assert.equal(r.empresasPendientes, 2);
  assert.deepEqual(r.listado.map((x) => [x.razon, x.bandera]), [['EMPRESA CUATRO', 'verde'], ['EMPRESA CINCO', 'amarilla']]);
});

test('elige la hoja de trabajo IPE entre varias hojas y tolera la falta de MI GESTION', () => {
  const mayor = [['', 'MAYOR AUXILIAR - RUT'], [], ['CENTRO FINANCIERO', '', 'X']];
  const sinEstado = AOA.map((r) => [r[0], ...r.slice(3)]);
  let r = parser.elegirHoja([{ nombre: 'Mayor', inicio: mayor }, { nombre: 'Sin', inicio: sinEstado }, { nombre: 'Trabajo', inicio: AOA }]);
  assert.equal(r.elegida.nombre, 'Trabajo');
  r = parser.elegirHoja([{ nombre: 'Mayor', inicio: mayor }, { nombre: 'Sin', inicio: sinEstado }]);
  assert.equal(r.elegida.nombre, 'Sin');
  assert.equal(parser.parseAoa(sinEstado).tieneEstado, false);
  assert.equal(parser.elegirHoja([{ nombre: 'Mayor', inicio: mayor }]).elegida, null);
});

test('copia MI GESTION y TIPO DE IPE del mes anterior', () => {
  const anterior = parser.parseAoa(AOA).rows;
  const nuevo = AOA.map((r) => [r[0], ...r.slice(3)]);
  nuevo.push(['', 'Ana', '', 51, 88888888, 8, 'EMPRESA NUEVA', 'EFE', 700, 'I', 1, '01-10-2026', '', 0, 50, 50, 'ACREEDOR', '', '']);
  const res = parser.copiarEstados(parser.parseAoa(nuevo).rows, anterior);
  assert.equal(res.exactos, 5); // EMPRESA CINCO no tenía MI GESTION en el mes anterior
  assert.equal(res.sinCoincidencia, 2);
  assert.equal(res.rows[0].tipoIpe, 'Pago en exceso');
  assert.equal(res.rows[0].categoria, 'APLICAR');
  assert.equal(res.rows.at(-1).categoria, 'SIN ESTADO');
});

test('reconoce la columna ACREEDOR/DEUDOR aunque no tenga encabezado', () => {
  const sinXX = AOA.map((r, i) => (i === 0 ? r.map((h) => (h === 'XX' ? '' : h)) : r));
  const { rows, columnas } = parser.parseAoa(sinXX);
  assert.equal(columnas.naturaleza, 18);
  assert.equal(rows[2].naturaleza, 'DEUDOR');
});

test('js/plantilla.js corresponde a las plantillas .docx (ejecutar scripts/embed_template.py si falla)', () => {
  const js = fs.readFileSync(path.join(__dirname, '../js/plantilla.js'), 'utf8');
  for (const [nombre, archivo] of [['plantillaB64', 'memo_template.docx'], ['plantillaDevolucionB64', 'devolucion_template.docx']]) {
    const b64 = js.match(new RegExp(`${nombre} = "([^"]+)"`))[1];
    assert.ok(Buffer.from(b64, 'base64').equals(fs.readFileSync(path.join(__dirname, '../templates', archivo))), archivo);
  }
});

test('frase "Asimismo" con varios movimientos previos', () => {
  const mov = (tipo, monto, fecha) => ({ tipo, monto, fecha, numero: '' });
  assert.deepEqual(movimientos.frasesAsimismo([]), []);
  assert.deepEqual(sp(movimientos.frasesAsimismo([mov('memo', 20000, '19.OCT.2026'), mov('memo', 10000, '20.OCT.2026')])), [
    'Asimismo, se aplicaron $ 30.000 en total, según el siguiente detalle: $ 20.000 según MEMO N°[[N° memo]] de fecha [[día]].OCT.2026; '
    + 'y $ 10.000 según MEMO N°[[N° memo]] de fecha [[día]].OCT.2026.',
  ]);
  const tres = sp(movimientos.frasesAsimismo([mov('memo', 1, '01.SEP.2026'), mov('memo', 2, '02.OCT.2026'), mov('memo', 3, '03.OCT.2026')]))[0];
  assert.match(tres, /^Asimismo, se aplicaron \$ 6 en total, según el siguiente detalle: \$ 1 .*SEP\.2026; \$ 2 .*; y \$ 3 .*OCT\.2026\.$/);
  assert.equal(sp(movimientos.frasesAsimismo([mov('memo', 20, '19.OCT.2026'), mov('devolucion', 10, '20.OCT.2026')]))[0],
    'Asimismo, este ingreso registra los siguientes movimientos: aplicación de $ 20 según MEMO N°[[N° memo]] de fecha [[día]].OCT.2026; '
    + 'y devolución de $ 10 según OF. (O) N°[[N° oficio]] de fecha [[día]].OCT.2026.');
});

test('ultimaGestion toma el último paso numerado', () => {
  assert.equal(parser.ultimaGestion('1) Correo a cliente 2 ) Cliente responde. 12)Correo final'), 'Correo final');
});
