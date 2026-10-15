// Planilla de devoluciones: consolida los IPE con estado DEVOLVER por empresa para Tesorería.
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const parser = typeof require === 'function' ? require('./parser.js') : root.IPE.parser;
  const memo = typeof require === 'function' ? require('./memo.js') : root.IPE.memo;
  const api = factory(fmt, parser, memo);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).devoluciones = api;
})(typeof self !== 'undefined' ? self : this, function (fmt, parser, memo) {
  function buildGroups(rows) {
    return memo.groupByEmpresa(rows.filter((r) => r.categoria === 'DEVOLVER')).map((g) => ({
      ...g,
      total: g.rows.reduce((s, r) => s + r.saldoNeto, 0),
    }));
  }

  /** Hojas (arreglos de filas) para exportar con SheetJS. */
  function sheets(groups) {
    const resumen = [
      ['RUT', 'Razón social', 'Cant. IPE', 'Total a devolver', 'Banco', 'Tipo de cuenta', 'Nº de cuenta', 'Correo de contacto', 'Observación'],
      ...groups.map((g) => [g.rutFmt, g.razon, g.rows.length, g.total, '', '', '', '', '']),
      ['', 'TOTAL', groups.reduce((s, g) => s + g.rows.length, 0), groups.reduce((s, g) => s + g.total, 0), '', '', '', '', ''],
    ];
    const detalle = [
      ['RUT', 'Razón social', 'C.F.', 'Doc.', 'Nro.', 'Comprob.', 'Fecha', 'Saldo', 'XX', 'Saldo neto', 'Justificación', 'Última gestión', 'Fila Excel'],
    ];
    for (const g of groups) {
      for (const r of g.rows) {
        detalle.push([g.rutFmt, g.razon, r.cf, r.doc, r.nro, r.comprob, fmt.fechaCorta(r.fecha), r.saldo,
          r.naturaleza, r.saldoNeto, r.justificacion, parser.ultimaGestion(r.gestion), r.fila]);
      }
    }
    return { resumen, detalle };
  }

  return { buildGroups, sheets };
});
