// Memos de aplicación: agrupa los IPE con estado APLICAR por empresa y rellena la plantilla Word.
// Cada IPE puede aplicarse a varios documentos; el saldo se va descontando en cada aplicación y,
// si queda remanente, el memo agrega una "Nota: Diferencia de ..." (formato memo Nº 176).
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).memo = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  /** Identificador estable de un IPE (empresa + documento + comprobante). */
  function rowKey(r) {
    return `${r.rut}-${r.dv}|${r.doc}|${r.nro}|${r.comprob}`;
  }

  function groupByEmpresa(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = `${r.rut}-${r.dv}`;
      if (!map.has(key)) map.set(key, { key, rut: r.rut, dv: r.dv, rutFmt: r.rutFmt, razon: r.razon, rows: [] });
      map.get(key).rows.push(r);
    }
    return [...map.values()].sort((a, b) => a.razon.localeCompare(b.razon, 'es'));
  }

  /**
   * Aplicaciones iniciales desde las columnas del Excel. Admite varias separadas por ";" o salto de línea:
   * DOCTO A APLICAR "Pagar FEE 263576; Pagar FEE 262729" y MONTO A APLICAR "2005; 2327".
   */
  function parseAplicaciones(docto, monto) {
    const split = (v) => (typeof v === 'number' ? [v] : String(v ?? '').split(/[;\n]+/).map((s) => s.trim()).filter(Boolean));
    const doctos = split(docto);
    const montos = split(monto);
    const n = Math.max(doctos.length, montos.length);
    const out = [];
    for (let i = 0; i < n; i++) out.push({ docto: doctos[i] ?? '', monto: montos[i] === undefined ? '' : fmt.toNumber(montos[i]) });
    return out;
  }

  /**
   * Índice del historial: por cada IPE, el registro de su memo más reciente
   * ({ saldo, aplicado, remanente, fechaMemo, generado }).
   */
  function indiceHistorial(historial) {
    const idx = new Map();
    for (const e of historial) {
      for (const d of e.documentos) {
        const prev = idx.get(d.key);
        if (!prev || e.generado > prev.generado) idx.set(d.key, { ...d, fechaMemo: e.fechaMemo, generado: e.generado });
      }
    }
    return idx;
  }

  /**
   * Un IPE con memo previo queda oculto solo si su saldo no cambió (Contabilidad aún no procesa el memo).
   * Si el saldo cambió (se aplicó una parte, o el cliente generó más saldo), vuelve a quedar pendiente.
   */
  function ocultoPorHistorial(r, previo) {
    return Boolean(previo) && Math.round(previo.saldo) === Math.round(r.saldo);
  }

  /** Grupos de memo: empresas con filas APLICAR pendientes. `idx` es el índice del historial. */
  function buildGroups(rows, idx = new Map()) {
    const pendientes = rows.filter((r) => r.categoria === 'APLICAR' && !ocultoPorHistorial(r, idx.get(rowKey(r))));
    return groupByEmpresa(pendientes)
      .map((g) => ({
        ...g,
        rows: g.rows.map((r) => ({
          ...r,
          key: rowKey(r),
          previo: idx.get(rowKey(r)) || null,
          incluir: r.naturaleza !== 'DEUDOR',
          aplicaciones: parseAplicaciones(r.doctoAplicar, r.montoAplicar),
        })),
      }))
      // Si solo quedan filas DEUDOR (aplicaciones ya registradas), no hay memo pendiente.
      .filter((g) => g.rows.some((r) => r.naturaleza !== 'DEUDOR'));
  }

  /** Saldo antes/después de cada aplicación y remanente final. */
  function calcular(r) {
    let saldo = r.saldo;
    const pasos = r.aplicaciones.map((a) => {
      const monto = Number(a.monto) || 0;
      const paso = { docto: a.docto, monto, saldoAntes: saldo, saldoDespues: saldo - monto };
      saldo -= monto;
      return paso;
    });
    return { pasos, aplicado: r.saldo - saldo, remanente: saldo };
  }

  function warnings(group) {
    const w = [];
    const sel = group.rows.filter((r) => r.incluir);
    if (!sel.length) w.push('No hay documentos seleccionados.');
    if (group.rows.some((r) => r.naturaleza === 'DEUDOR'))
      w.push('Tiene filas DEUDOR (aplicaciones ya registradas): revisar si corresponde incluirlas.');
    for (const r of sel) {
      const id = `${r.doc} ${r.nro}`;
      const { remanente } = calcular(r);
      if (!r.justificacion) w.push(`${id}: sin justificación (el MOTIVO quedará solo con el documento).`);
      if (!r.aplicaciones.length) w.push(`${id}: no tiene documentos a aplicar.`);
      if (r.aplicaciones.some((a) => !a.docto || a.monto === '' || !(Number(a.monto) > 0)))
        w.push(`${id}: hay una aplicación sin documento o sin monto.`);
      if (remanente < 0) w.push(`${id}: lo aplicado supera el saldo en ${fmt.formatCLP(-remanente)}.`);
    }
    return w;
  }

  /** Datos que consume la plantilla templates/memo_template.docx */
  function memoData(group, opts) {
    const sel = group.rows.filter((r) => r.incluir);
    const items = [];
    const notas = [];
    for (const r of sel) {
      const { pasos, remanente } = calcular(r);
      if (!pasos.length) {
        items.push({ doc: r.doc, nro: r.nro, monto: fmt.formatCLP(r.saldo), docto_aplicar: '', monto_aplicar: '' });
        continue;
      }
      for (const p of pasos) {
        items.push({
          doc: r.doc,
          nro: r.nro,
          monto: fmt.formatCLP(p.saldoAntes),
          docto_aplicar: p.docto,
          monto_aplicar: p.monto ? fmt.formatCLP(p.monto) : '',
        });
      }
      if (remanente > 0) notas.push({ doc: r.doc, nro: r.nro, remanente: fmt.formatCLP(remanente) });
    }
    return {
      fecha: fmt.fechaMemo(opts.fecha),
      razon_social: group.razon,
      rut: group.rutFmt,
      iniciales: opts.iniciales || '',
      motivos: sel.map((r, i) => ({ primero: i === 0, doc: r.doc, nro: r.nro, justificacion: r.justificacion })),
      items,
      notas,
    };
  }

  function fileName(group) {
    const base = group.razon.normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '');
    return `MEMO_IPE_${base}_${group.rut}.docx`;
  }

  /** Registro para el historial: guarda lo necesario para volver a descargar el memo. */
  function historialEntry(group, data, generadoEn) {
    const sel = group.rows.filter((r) => r.incluir);
    return {
      id: `${generadoEn.getTime()}-${group.key}`,
      generado: generadoEn.toISOString(),
      fechaMemo: data.fecha,
      empresa: group.razon,
      rut: group.rutFmt,
      archivo: fileName(group),
      documentos: sel.map((r) => {
        const { aplicado, remanente } = calcular(r);
        return { key: r.key, doc: r.doc, nro: r.nro, saldo: r.saldo, aplicado, remanente,
          aplicaciones: r.aplicaciones.map((a) => `${a.docto} (${fmt.formatCLP(Number(a.monto) || 0)})`).join('; ') };
      }),
      data,
    };
  }

  /** Genera el .docx. PizZip y Docxtemplater se reciben como parámetro (navegador o Node). */
  function render(templateBuffer, data, PizZip, Docxtemplater) {
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
    doc.render(data);
    return doc.getZip();
  }

  return { rowKey, indiceHistorial, ocultoPorHistorial, buildGroups, groupByEmpresa, parseAplicaciones, calcular, warnings, memoData, fileName, historialEntry, render };
});
