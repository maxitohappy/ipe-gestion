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
  const TIPOS_CUENTA = ['Cuenta Corriente', 'Cuenta Vista', 'Cuenta de Ahorro', 'Cuenta RUT'];
  // Bancos e instituciones con cuentas para transferencias en Chile ("Otro" permite escribir el nombre).
  const BANCOS = [
    'Banco de Chile', 'Banco Edwards', 'BancoEstado', 'Banco Santander', 'Banco BCI', 'Scotiabank', 'Banco Itaú',
    'Banco BICE', 'Banco Security', 'Banco Falabella', 'Banco Ripley', 'Banco Consorcio', 'Banco Internacional',
    'HSBC Bank Chile', 'Banco BTG Pactual', 'JP Morgan Chase Bank', 'Banco de la Nación Argentina', 'Bank of China',
    'China Construction Bank', 'MUFG Bank', 'Tenpo', 'Coopeuch', 'Mercado Pago', 'Prepago Los Héroes',
  ];

  /** Grupos de devolución: empresas con IPE en DEVOLVER y saldo pendiente. */
  function buildGroups(rows, idx = new Map()) {
    return memo.prepararGrupos(rows, 'DEVOLVER', idx, (r, pendiente) => ({ montoDevolver: pendiente }))
      .map((g) => ({ ...g, banco: { tipoCuenta: TIPOS_CUENTA[0], numeroCuenta: '', banco: '', bancoOtro: '', correos: '' } }));
  }

  function total(group) {
    return group.rows.filter((r) => r.incluir).reduce((s, r) => s + (Number(r.montoDevolver) || 0), 0);
  }

  /** Banco elegido (o escrito en "Otro"). */
  function nombreBanco(b) {
    return (b.banco === 'Otro' ? b.bancoOtro : b.banco || '').trim();
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
      if (!(Number(r.montoDevolver) > 0)) w.push(`${id}: falta el monto a devolver.`);
      if (Number(r.montoDevolver) > r.saldoInicial) w.push(`${id}: el monto a devolver supera el saldo pendiente (${fmt.formatCLP(r.saldoInicial)}).`);
    }
    const b = group.banco;
    if (!b.numeroCuenta.trim() || !nombreBanco(b)) w.push('Faltan datos bancarios (banco y número de cuenta).');
    if (!correos(b.correos).length) w.push('Falta al menos un correo electrónico del cliente.');
    return w;
  }

  /** Datos que consume la plantilla templates/devolucion_template.docx */
  function devolucionData(group, opts) {
    const per = group.periodo || 'actual';
    // Montos como en cada formato: "$90.666.-" (año en curso) o "$ 158.182" (años anteriores).
    const montoDev = (n) => (per === 'anterior' ? `$ ${fmt.miles(n)}` : `$${fmt.miles(n)}.-`);
    const sel = group.rows.filter((r) => r.incluir);
    const b = group.banco;
    const lista = correos(b.correos);
    return {
      periodo: per,
      fecha: fmt.fechaMesAnio(opts.fecha),
      rut: group.rutFmt,
      razon_social: group.razon,
      items: sel.map((r) => ({
        cf: r.cf,
        tipo_docto: `${r.doc} ${r.nro}`,
        monto: montoDev(Number(r.montoDevolver) || 0),
        observaciones: glosas.partes(memo.textoCompleto(r)),
      })),
      tipo_cuenta: glosas.partes(b.tipoCuenta || '[[tipo de cuenta]]'),
      numero_cuenta: glosas.partes(b.numeroCuenta.trim() || '[[N° de cuenta]]'),
      banco: glosas.partes(nombreBanco(b) || '[[banco]]'),
      monto_total: montoDev(total(group)),
      // Año en curso: un correo por línea. Años anteriores: en una línea separados por "; ".
      correos: (lista.length ? lista : ['[correo]']).map((correo, i) => ({ primero: i === 0, correo })),
      correos_linea: lista.length ? lista.join('; ') : '[correo]',
    };
  }

  function fileName(group, fecha) {
    return memo.nombreArchivo('ANEXO_DEVOLUCION', group, fecha);
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
    // Los oficios guardados con versiones anteriores traen estos campos como texto: se dividen en partes.
    const aPartes = (v) => (typeof v === 'string' ? glosas.partes(v) : v);
    doc.render({
      ...data,
      tipo_cuenta: aPartes(data.tipo_cuenta),
      numero_cuenta: aPartes(data.numero_cuenta),
      banco: aPartes(data.banco),
      items: (data.items || []).map((it) => ({ ...it, observaciones: aPartes(it.observaciones) })),
    });
    return doc.getZip();
  }

  return { TIPOS_CUENTA, BANCOS, buildGroups, total, nombreBanco, correos, warnings, devolucionData, fileName, historialEntry, render };
});
