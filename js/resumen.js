// Resumen: cantidad de empresas (total, con memos, con devoluciones, pendientes) y listado de las pendientes.
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
   * El listado incluye solo los ingresos anticipados de las empresas pendientes, con su categoría (bandera).
   */
  function compute(rows, memoGroups, devGroups, hoy = new Date()) {
    const total = new Set(rows.map((r) => nombre(r.razon)));
    const enMemos = new Set(memoGroups.map((g) => nombre(g.razon)));
    const enDev = new Set(devGroups.map((g) => nombre(g.razon)));
    const gestionada = (r) => enMemos.has(nombre(r.razon)) || enDev.has(nombre(r.razon));
    const pendientes = new Set([...total].filter((n) => !enMemos.has(n) && !enDev.has(n)));
    return {
      empresasTotal: total.size,
      empresasMemos: enMemos.size,
      empresasDevoluciones: enDev.size,
      empresasPendientes: pendientes.size,
      listado: rows.filter((r) => !gestionada(r)).map((r) => ({ ...r, bandera: fmt.bandera(r.fecha, hoy) })),
    };
  }

  return { compute };
});
