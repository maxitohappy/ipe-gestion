// Memos de aplicación: agrupa los IPE con estado APLICAR por empresa y rellena la plantilla Word.
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).memo = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  function groupByEmpresa(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = `${r.rut}-${r.dv}`;
      if (!map.has(key)) map.set(key, { key, rut: r.rut, dv: r.dv, rutFmt: r.rutFmt, razon: r.razon, rows: [] });
      map.get(key).rows.push(r);
    }
    return [...map.values()].sort((a, b) => a.razon.localeCompare(b.razon, 'es'));
  }

  /** Grupos de memo: un grupo por empresa con filas APLICAR. Las filas DEUDOR quedan desmarcadas por defecto. */
  function buildGroups(rows) {
    return groupByEmpresa(rows.filter((r) => r.categoria === 'APLICAR')).map((g) => ({
      ...g,
      rows: g.rows.map((r) => ({ ...r, incluir: r.naturaleza !== 'DEUDOR' })),
    }));
  }

  function warnings(group) {
    const w = [];
    const sel = group.rows.filter((r) => r.incluir);
    if (!sel.length) w.push('No hay documentos seleccionados.');
    if (group.rows.some((r) => r.naturaleza === 'DEUDOR'))
      w.push('Tiene filas DEUDOR (aplicaciones ya registradas): revisar si corresponde incluirlas.');
    if (sel.some((r) => !r.justificacion)) w.push('Hay documentos sin justificación (el MOTIVO quedará solo con el documento).');
    if (sel.some((r) => !r.doctoAplicar)) w.push('Falta "Docto a aplicar" en algún documento.');
    if (sel.some((r) => r.montoAplicar === '' || r.montoAplicar === null || r.montoAplicar === undefined))
      w.push('Falta "Monto a aplicar" en algún documento.');
    return w;
  }

  /** Datos que consume la plantilla templates/memo_template.docx */
  function memoData(group, opts) {
    const sel = group.rows.filter((r) => r.incluir);
    return {
      fecha: fmt.fechaMemo(opts.fecha),
      razon_social: group.razon,
      rut: group.rutFmt,
      iniciales: opts.iniciales || '',
      firmante: opts.firmante || '',
      cargo_firmante: opts.cargo || '',
      motivos: sel.map((r, i) => ({ br: i > 0, doc: r.doc, nro: r.nro, justificacion: r.justificacion })),
      items: sel.map((r) => ({
        doc: r.doc,
        nro: r.nro,
        monto: fmt.formatCLP(r.saldo),
        docto_aplicar: r.doctoAplicar,
        monto_aplicar: fmt.formatMontoLibre(r.montoAplicar),
      })),
    };
  }

  function fileName(group) {
    const base = group.razon.normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '');
    return `MEMO_IPE_${base}_${group.rut}.docx`;
  }

  /** Genera el .docx. PizZip y Docxtemplater se reciben como parámetro (navegador o Node). */
  function render(templateBuffer, data, PizZip, Docxtemplater) {
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
    doc.render(data);
    return doc.getZip();
  }

  return { buildGroups, groupByEmpresa, warnings, memoData, fileName, render };
});
