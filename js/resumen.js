// Panel de resumen: totales por estado, antigüedad y alertas de calidad de datos.
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).resumen = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  const TRAMOS = [
    { label: '0–30 días', max: 30 },
    { label: '31–90 días', max: 90 },
    { label: '91–180 días', max: 180 },
    { label: '181–365 días', max: 365 },
    { label: 'Más de 1 año', max: Infinity },
  ];
  const DIAS_REITERAR = 30;

  function compute(rows, hoy, tieneColumnasAplicar, tieneEstado = true) {
    const porEstado = new Map();
    for (const r of rows) {
      const e = porEstado.get(r.categoria) || { categoria: r.categoria, cantidad: 0, monto: 0 };
      e.cantidad += 1;
      e.monto += r.saldoNeto;
      porEstado.set(r.categoria, e);
    }

    const antiguedad = TRAMOS.map((t) => ({ label: t.label, cantidad: 0, monto: 0 }));
    for (const r of rows) {
      if (!r.fecha) continue;
      const dias = fmt.diasEntre(r.fecha, hoy);
      const i = TRAMOS.findIndex((t) => dias <= t.max);
      antiguedad[i].cantidad += 1;
      antiguedad[i].monto += r.saldoNeto;
    }

    const alertas = [];
    const add = (r, tipo, mensaje) => alertas.push({ fila: r.fila, razon: r.razon, doc: `${r.doc} ${r.nro}`, tipo, mensaje });
    const vistos = new Map();
    for (const r of rows) {
      const k = `${r.rut}|${r.doc}|${r.nro}|${r.saldo}|${r.naturaleza}`;
      if (vistos.has(k)) add(r, 'Duplicado', `Mismo documento, saldo y naturaleza que la fila ${vistos.get(k)}: verificar si es duplicado`);
      else vistos.set(k, r.fila);
    }
    for (const r of rows) {
      if (r.categoria === 'APLICAR' && tieneColumnasAplicar && (!r.doctoAplicar || r.montoAplicar === ''))
        add(r, 'Memo', 'APLICAR sin "Docto a aplicar" o "Monto a aplicar"');
      if ((r.categoria === 'APLICAR' || r.categoria === 'DEVOLVER') && !r.justificacion)
        add(r, 'Datos', 'Sin justificación del ingreso anticipado');
      if (tieneEstado && (r.categoria === 'OTRO' || r.categoria === 'SIN ESTADO'))
        add(r, 'Estado', `Estado no estándar en MIS APUNTES: "${r.estado || '(vacío)'}"`);
      if (!r.gestion) add(r, 'Gestión', 'Sin gestiones registradas');
      if (r.categoria === 'CORREO ENVIADO') {
        const f = fmt.parseFecha(r.estado);
        if (f && fmt.diasEntre(f, hoy) > DIAS_REITERAR)
          add(r, 'Seguimiento', `Correo enviado hace ${fmt.diasEntre(f, hoy)} días sin cambio de estado: reiterar`);
      }
      if (r.saldo === 0) add(r, 'Datos', 'Saldo cero: evaluar eliminar de la planilla');
    }

    const total = rows.reduce((s, r) => s + r.saldoNeto, 0);
    return {
      total,
      cantidad: rows.length,
      empresas: new Set(rows.map((r) => `${r.rut}-${r.dv}`)).size,
      porEstado: [...porEstado.values()].sort((a, b) => b.monto - a.monto),
      antiguedad,
      alertas,
    };
  }

  return { compute, TRAMOS };
});
