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

  /**
   * Frases que complementan el MOTIVO / OBSERVACIONES con los movimientos previos del IPE:
   * "Asimismo, con fecha [día].OCT.2026, se aplicaron $20, según MEMO N°[N° memo], de fecha [día].OCT.2026."
   * El N° del memo u oficio se completa en el Word (queda resaltado).
   */
  function frasesAsimismo(previos) {
    return previos
      .filter((m) => m.monto > 0)
      .map((m) => {
        // El día queda por completar en el Word; se indican solo mes y año: "[día].OCT.2026".
        const mesAnio = String(m.fecha || '').replace(/^\d{1,2}\./, '');
        const fecha = mesAnio ? `[[día]].${mesAnio}` : '[[fecha]]';
        return m.tipo === 'devolucion'
          ? `Asimismo, con fecha ${fecha}, se devolvieron ${fmt.formatCLP(m.monto)}, según OF. (O) N°${m.numero || '[[N° oficio]]'}, de fecha ${fecha}.`
          : `Asimismo, con fecha ${fecha}, se aplicaron ${fmt.formatCLP(m.monto)}, según MEMO N°${m.numero || '[[N° memo]]'}, de fecha ${fecha}.`;
      });
  }

  return { rowKey, indice, pendiente, frasesAsimismo };
});
