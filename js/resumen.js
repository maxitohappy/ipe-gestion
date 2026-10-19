// Resumen: cantidad de empresas (total, con memos, con devoluciones, pendientes) y banderas por antigüedad.
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).resumen = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  const nombre = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toUpperCase();

  /**
   * Empresas contadas por razón social. "Pendientes" son las que no tienen IPE en Memos (APLICAR) ni en
   * Devoluciones; equivale a total − memos − devoluciones cuando ninguna empresa está en ambas listas.
   */
  function compute(rows, memoGroups, devGroups, hoy = new Date()) {
    const total = new Set(rows.map((r) => nombre(r.razon)));
    const enMemos = new Set(memoGroups.map((g) => nombre(g.razon)));
    const enDev = new Set(devGroups.map((g) => nombre(g.razon)));
    const gestionadas = new Set([...enMemos, ...enDev]);
    const banderas = { roja: 0, amarilla: 0, verde: 0 };
    const listado = rows.map((r) => {
      const b = fmt.bandera(r.fecha, hoy);
      if (b) banderas[b] += 1;
      return { ...r, bandera: b };
    });
    return {
      empresasTotal: total.size,
      empresasMemos: enMemos.size,
      empresasDevoluciones: enDev.size,
      empresasPendientes: [...total].filter((n) => !gestionadas.has(n)).length,
      banderas,
      listado,
    };
  }

  return { compute };
});
