// Memos de aplicación: agrupa los IPE con estado APLICAR por empresa y rellena la plantilla Word.
// Cada IPE puede aplicarse a varios documentos; el saldo se va descontando en cada aplicación y,
// si queda remanente, el memo agrega una "Nota: Diferencia de ..." (formato memo Nº 176).
// Si el IPE ya tuvo memos (historial), se parte del saldo pendiente y el MOTIVO lo menciona ("Asimismo, ...").
(function (root, factory) {
  const req = typeof require === 'function';
  const fmt = req ? require('./format.js') : root.IPE.format;
  const glosas = req ? require('./glosas.js') : root.IPE.glosas;
  const movimientos = req ? require('./movimientos.js') : root.IPE.movimientos;
  const api = factory(fmt, glosas, movimientos);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).memo = api;
})(typeof self !== 'undefined' ? self : this, function (fmt, glosas, movimientos) {
  const { rowKey } = movimientos;

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
   * Aplicaciones iniciales desde las columnas del Excel (opcionales). Admite varias separadas por ";":
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
   * Texto del MOTIVO / OBSERVACIONES según el TIPO DE IPE del Excel: la glosa correspondiente o, si no hay
   * tipo, la justificación del Excel; si tampoco hay justificación, un dato [motivo] a completar en el Word.
   */
  function textoInicial(r) {
    const g = glosas.porTipo(r.tipoIpe);
    return { glosaId: g ? g.id : '', texto: g ? g.texto : (r.justificacion || '[[motivo]]') };
  }

  /**
   * Prepara los IPE de una categoría con saldo pendiente > 0, agrupados por empresa.
   * Compartido por memos (APLICAR) y devoluciones (DEVOLVER).
   */
  function prepararGrupos(rows, categoria, idx, extra) {
    const pendientes = [];
    for (const r of rows) {
      if (r.categoria !== categoria || r.naturaleza === 'DEUDOR') continue;
      const { pendiente, previos } = movimientos.pendiente(r, idx);
      if (pendiente <= 0) continue;
      const { glosaId, texto } = textoInicial(r);
      pendientes.push({ ...r, key: rowKey(r), incluir: true, saldoInicial: pendiente, previos, glosaId, texto, ...extra(r, pendiente) });
    }
    return groupByEmpresa(pendientes);
  }

  /** Grupos de memo: empresas con IPE en APLICAR y saldo pendiente. `idx` es movimientos.indice(...). */
  function buildGroups(rows, idx = new Map()) {
    return prepararGrupos(rows, 'APLICAR', idx, (r) => ({ aplicaciones: parseAplicaciones(r.doctoAplicar, r.montoAplicar) }));
  }

  /** Saldo antes/después de cada aplicación y remanente final, partiendo del saldo pendiente. */
  function calcular(r) {
    let saldo = r.saldoInicial ?? r.saldo;
    const inicio = saldo;
    const pasos = r.aplicaciones.map((a) => {
      const monto = Number(a.monto) || 0;
      const paso = { docto: a.docto, monto, saldoAntes: saldo, saldoDespues: saldo - monto };
      saldo -= monto;
      return paso;
    });
    return { pasos, aplicado: inicio - saldo, remanente: saldo };
  }

  /** MOTIVO completo: texto del tipo de IPE + frases de movimientos previos. */
  function textoCompleto(r) {
    return [r.texto, ...movimientos.frasesAsimismo(r.previos || [])].filter(Boolean).join(' ');
  }

  function warnings(group) {
    const w = [];
    const sel = group.rows.filter((r) => r.incluir);
    if (!sel.length) w.push('No hay documentos seleccionados.');
    for (const r of sel) {
      const id = `${r.doc} ${r.nro}`;
      const { remanente } = calcular(r);
      if (!r.aplicaciones.length) w.push(`${id}: no tiene documentos a aplicar.`);
      if (r.aplicaciones.some((a) => !a.docto || a.monto === '' || !(Number(a.monto) > 0)))
        w.push(`${id}: hay una aplicación sin documento o sin monto.`);
      if (remanente < 0) w.push(`${id}: lo aplicado supera el saldo pendiente en ${fmt.formatCLP(-remanente)}.`);
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
        items.push({ doc: r.doc, nro: r.nro, monto: fmt.formatCLP(r.saldoInicial ?? r.saldo), docto_aplicar: '', monto_aplicar: '' });
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
      if (remanente > 0) notas.push({ doc: r.doc, nro: r.nro, remanente: fmt.formatCLP(remanente).replace(' ', '\u00A0') });
    }
    return {
      fecha: fmt.fechaMesAnio(opts.fecha),
      razon_social: group.razon,
      rut: group.rutFmt,
      iniciales: opts.iniciales || '',
      motivos: sel.map((r, i) => ({ primero: i === 0, doc: r.doc, nro: r.nro, justificacion: glosas.partes(textoCompleto(r)) })),
      items,
      notas,
    };
  }

  /**
   * Nombre del .docx: prefijo, empresa, RUT, fecha del documento y hora de generación
   * (dos memos de la misma empresa en el mismo día no quedan con igual nombre).
   */
  function nombreArchivo(prefijo, group, fecha = new Date(), ahora = new Date()) {
    const base = group.razon.normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '');
    const dos = (n) => String(n).padStart(2, '0');
    const f = `${fecha.getFullYear()}-${dos(fecha.getMonth() + 1)}-${dos(fecha.getDate())}`;
    return `${prefijo}_${base}_${group.rut}_${f}_${dos(ahora.getHours())}${dos(ahora.getMinutes())}${dos(ahora.getSeconds())}.docx`;
  }

  function fileName(group, fecha) {
    return nombreArchivo('MEMO_IPE', group, fecha);
  }

  /** Registro para el historial de memos: guarda lo necesario para volver a descargarlo y calcular saldos. */
  function historialEntry(group, data, generadoEn, opts) {
    const sel = group.rows.filter((r) => r.incluir);
    return {
      id: `${generadoEn.getTime()}-${group.key}`,
      tipo: 'memo',
      generado: generadoEn.toISOString(),
      fechaMemo: data.fecha,
      fechaCompleta: fmt.fechaMemo(opts.fecha),
      numero: '',
      empresa: group.razon,
      rut: group.rutFmt,
      archivo: fileName(group, opts.fecha),
      documentos: sel.map((r) => {
        const { aplicado, remanente } = calcular(r);
        return {
          key: r.key, doc: r.doc, nro: r.nro,
          saldo: r.saldo, saldoExcel: r.saldo, saldoInicial: r.saldoInicial ?? r.saldo, aplicado, remanente,
          aplicaciones: r.aplicaciones.map((a) => `${a.docto} (${fmt.formatCLP(Number(a.monto) || 0)})`).join('; '),
        };
      }),
      data,
    };
  }

  /** Genera el .docx. PizZip y Docxtemplater se reciben como parámetro (navegador o Node). */
  function render(templateBuffer, data, PizZip, Docxtemplater) {
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
    // "primero" marca el documento que lleva el rótulo "MOTIVO :". Se recalcula siempre para que los
    // memos guardados en el historial con versiones anteriores (que no lo traían) se generen completos.
    // Los memos guardados con versiones anteriores traen la justificación como texto: se divide en partes.
    const motivos = (data.motivos || []).map((m, i) => ({
      ...m, primero: i === 0, justificacion: typeof m.justificacion === 'string' ? glosas.partes(m.justificacion) : m.justificacion,
    }));
    doc.render({ ...data, motivos, notas: data.notas || [] });
    return doc.getZip();
  }

  return {
    rowKey, groupByEmpresa, parseAplicaciones, textoInicial, prepararGrupos, buildGroups, calcular, textoCompleto,
    warnings, memoData, nombreArchivo, fileName, historialEntry, render,
  };
});
