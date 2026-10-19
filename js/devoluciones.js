// Devoluciones: agrupa los IPE con estado DEVOLVER por empresa y rellena la plantilla del oficio de devolución
// (templates/devolucion_template.docx). Igual que los memos, parte del saldo pendiente según el historial y
// permite devolver una parte; OBSERVACIONES usa la glosa del TIPO DE IPE.
(function (root, factory) {
  const req = typeof require === 'function';
  const fmt = req ? require('./format.js') : root.IPE.format;
  const glosas = req ? require('./glosas.js') : root.IPE.glosas;
  const memo = req ? require('./memo.js') : root.IPE.memo;
  const api = factory(fmt, glosas, memo);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).devoluciones = api;
})(typeof self !== 'undefined' ? self : this, function (fmt, glosas, memo) {
  const C = glosas.PLACEHOLDER;
  const TIPOS_CUENTA = ['Cuenta Corriente', 'Cuenta Vista', 'Cuenta de Ahorro', 'Cuenta RUT'];
  const REFERENCIA_BASE = `Correo electrónico ${C} de fecha ${C} a las ${C} hrs.`;

  /** Grupos de devolución: empresas con IPE en DEVOLVER y saldo pendiente. */
  function buildGroups(rows, idx = new Map()) {
    return memo.prepararGrupos(rows, 'DEVOLVER', idx, (r, pendiente) => ({ montoDevolver: pendiente }))
      .map((g) => ({ ...g, banco: { referencia: '', tipoCuenta: TIPOS_CUENTA[0], numeroCuenta: '', banco: '', correos: '' } }));
  }

  function total(group) {
    return group.rows.filter((r) => r.incluir).reduce((s, r) => s + (Number(r.montoDevolver) || 0), 0);
  }

  function correos(texto) {
    return String(texto ?? '').split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean);
  }

  function warnings(group) {
    const w = [];
    const sel = group.rows.filter((r) => r.incluir);
    if (!sel.length) w.push('No hay documentos seleccionados.');
    for (const r of sel) {
      const id = `${r.doc} ${r.nro}`;
      if (!r.texto) w.push(`${id}: sin OBSERVACIONES (elige el TIPO DE IPE o escribe la justificación).`);
      const faltan = glosas.pendientes(r.texto);
      if (faltan) w.push(`${id}: OBSERVACIONES tiene ${faltan} dato(s) por completar (<completar por el usuario>).`);
      if (!(Number(r.montoDevolver) > 0)) w.push(`${id}: falta el monto a devolver.`);
      if (Number(r.montoDevolver) > r.saldoInicial) w.push(`${id}: el monto a devolver supera el saldo pendiente (${fmt.formatCLP(r.saldoInicial)}).`);
    }
    const b = group.banco;
    if (!b.referencia.trim()) w.push('Falta la REF. (correo del cliente con fecha y hora).');
    if (!b.numeroCuenta.trim() || !b.banco.trim()) w.push('Faltan datos bancarios (banco y número de cuenta).');
    if (!correos(b.correos).length) w.push('Falta al menos un correo electrónico del cliente.');
    return w;
  }

  /** Datos que consume la plantilla templates/devolucion_template.docx */
  function devolucionData(group, opts) {
    const sel = group.rows.filter((r) => r.incluir);
    const b = group.banco;
    const lista = correos(b.correos);
    return {
      fecha: fmt.fechaMesAnio(opts.fecha),
      referencia: b.referencia.trim() || REFERENCIA_BASE,
      rut: group.rutFmt,
      razon_social: group.razon,
      items: sel.map((r) => ({
        cf: r.cf,
        tipo_docto: `${r.doc} ${r.nro}`,
        monto: fmt.formatCLP(Number(r.montoDevolver) || 0),
        observaciones: memo.textoCompleto(r),
      })),
      tipo_cuenta: b.tipoCuenta || C,
      numero_cuenta: b.numeroCuenta.trim() || C,
      banco: b.banco.trim() || C,
      monto_total: fmt.formatCLP(total(group)),
      correos: (lista.length ? lista : [C]).map((correo, i) => ({ primero: i === 0, correo })),
      iniciales: opts.iniciales || '',
    };
  }

  function fileName(group, fecha) {
    return memo.nombreArchivo('DEVOLUCION_IPE', group, fecha);
  }

  /** Registro para el historial de devoluciones. */
  function historialEntry(group, data, generadoEn, opts) {
    const sel = group.rows.filter((r) => r.incluir);
    return {
      id: `${generadoEn.getTime()}-${group.key}`,
      tipo: 'devolucion',
      generado: generadoEn.toISOString(),
      fechaDocumento: data.fecha,
      fechaCompleta: fmt.fechaMemo(opts.fecha),
      numero: '',
      empresa: group.razon,
      rut: group.rutFmt,
      archivo: fileName(group, opts.fecha),
      banco: { ...group.banco },
      documentos: sel.map((r) => {
        const devuelto = Number(r.montoDevolver) || 0;
        return {
          key: r.key, doc: r.doc, nro: r.nro,
          saldoExcel: r.saldo, saldoInicial: r.saldoInicial, devuelto, remanente: r.saldoInicial - devuelto,
        };
      }),
      data,
    };
  }

  function render(templateBuffer, data, PizZip, Docxtemplater) {
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
    doc.render(data);
    return doc.getZip();
  }

  return { TIPOS_CUENTA, REFERENCIA_BASE, buildGroups, total, correos, warnings, devolucionData, fileName, historialEntry, render };
});
