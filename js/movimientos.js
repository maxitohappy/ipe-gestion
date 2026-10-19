// Movimientos registrados en los historiales (aplicaciones por memo y devoluciones) y saldo pendiente de cada IPE.
//
// Cada documento de un historial guarda el saldo que tenía el IPE en el Excel al generarse (saldoExcel) y el
// monto aplicado o devuelto. Al cargar un Excel:
//   - Los movimientos cuyo saldoExcel es igual al saldo actual aún no se reflejan en el Excel (Contabilidad o
//     Tesorería no los han procesado): se descuentan para obtener el saldo pendiente.
//   - Los de saldo distinto ya están reflejados (o el saldo cambió por otros motivos): no se descuentan.
// Así, un IPE de $100 con un memo por $20 queda pendiente por $80; con otro por $10, por $70.
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).movimientos = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  /** Identificador estable de un IPE (empresa + documento + número + comprobante). */
  function rowKey(r) {
    return `${r.rut}-${r.dv}|${r.doc}|${r.nro}|${r.comprob}`;
  }

  /** Índice clave de IPE -> movimientos (de ambos historiales), del más antiguo al más reciente. */
  function indice(historialMemos = [], historialDevoluciones = []) {
    const idx = new Map();
    const agregar = (e, tipo) => {
      for (const d of e.documentos || []) {
        const mov = {
          tipo,
          id: e.id,
          generado: e.generado,
          // Fecha completa del documento (dd.MMM.yyyy). Registros antiguos solo traían fechaMemo.
          fecha: e.fechaCompleta || e.fechaMemo || '',
          numero: e.numero || '',
          saldoExcel: d.saldoExcel ?? d.saldo,
          monto: tipo === 'devolucion' ? (d.devuelto ?? 0) : (d.aplicado ?? 0),
        };
        if (!idx.has(d.key)) idx.set(d.key, []);
        idx.get(d.key).push(mov);
      }
    };
    historialMemos.forEach((e) => agregar(e, 'memo'));
    historialDevoluciones.forEach((e) => agregar(e, 'devolucion'));
    for (const lista of idx.values()) lista.sort((a, b) => String(a.generado).localeCompare(String(b.generado)));
    return idx;
  }

  /** Saldo pendiente del IPE y movimientos previos registrados para él. */
  function pendiente(r, idx) {
    const previos = idx.get(rowKey(r)) || [];
    const sinProcesar = previos.filter((m) => Math.round(m.saldoExcel) === Math.round(r.saldo));
    const descontado = sinProcesar.reduce((s, m) => s + m.monto, 0);
    return { pendiente: r.saldo - descontado, previos, descontado };
  }

  /** Fecha del documento con el día por completar en el Word: "[día].OCT.2026". */
  function fechaSinDia(m) {
    const mesAnio = String(m.fecha || '').replace(/^\d{1,2}\./, '');
    return mesAnio ? `[[día]].${mesAnio}` : '[[fecha]]';
  }

  const documento = (m) => (m.tipo === 'devolucion'
    ? `OF. (O) N°${m.numero || '[[N° oficio]]'}`
    : `MEMO N°${m.numero || '[[N° memo]]'}`);

  // Monto con espacio no separable: en el Word "$" no queda en una línea y la cifra en la siguiente.
  const monto = (n) => fmt.formatCLP(n).replace(' ', '\u00A0');

  /** "a; b; y c" */
  function enumerar(items) {
    return items.length < 2 ? items.join('') : `${items.slice(0, -1).join('; ')}; y ${items[items.length - 1]}`;
  }

  /**
   * Frase que complementa el MOTIVO / OBSERVACIONES con los movimientos previos del IPE. El N° del memo u
   * oficio y el día se completan en el Word (quedan resaltados).
   * - Un movimiento: "Asimismo, con fecha [día].OCT.2026, se aplicaron $ 20, según MEMO N°[N° memo], de fecha [día].OCT.2026."
   * - Varios del mismo tipo: "Asimismo, se aplicaron $ 30 en total, según el siguiente detalle:
   *   $ 20 según MEMO N°[N° memo] de fecha [día].OCT.2026; y $ 10 según MEMO N°[N° memo] de fecha [día].OCT.2026."
   * - Aplicaciones y devoluciones: "Asimismo, este ingreso registra los siguientes movimientos:
   *   aplicación de $ 20 según MEMO N°… de fecha …; y devolución de $ 10 según OF. (O) N°… de fecha …."
   */
  function frasesAsimismo(previos) {
    const movs = previos.filter((m) => m.monto > 0);
    if (!movs.length) return [];
    if (movs.length === 1) {
      const [m] = movs;
      const verbo = m.tipo === 'devolucion' ? 'se devolvieron' : 'se aplicaron';
      return [`Asimismo, con fecha ${fechaSinDia(m)}, ${verbo} ${monto(m.monto)}, según ${documento(m)}, de fecha ${fechaSinDia(m)}.`];
    }
    const detalle = (m, conTipo) => `${conTipo ? (m.tipo === 'devolucion' ? 'devolución de ' : 'aplicación de ') : ''}${monto(m.monto)} según ${documento(m)} de fecha ${fechaSinDia(m)}`;
    const tipos = new Set(movs.map((m) => m.tipo));
    if (tipos.size > 1) {
      return [`Asimismo, este ingreso registra los siguientes movimientos: ${enumerar(movs.map((m) => detalle(m, true)))}.`];
    }
    const total = movs.reduce((s, m) => s + m.monto, 0);
    const verbo = movs[0].tipo === 'devolucion' ? 'se devolvieron' : 'se aplicaron';
    return [`Asimismo, ${verbo} ${monto(total)} en total, según el siguiente detalle: ${enumerar(movs.map((m) => detalle(m, false)))}.`];
  }

  return { rowKey, indice, pendiente, frasesAsimismo };
});
