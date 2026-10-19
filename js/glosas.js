// Glosas por TIPO DE IPE (documento GLOSAS). Se usan como MOTIVO del memo y como OBSERVACIONES
// de la devolución. Los datos marcados con <completar por el usuario> los completa el usuario.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).glosas = api;
})(typeof self !== 'undefined' ? self : this, function () {
  const C = '<completar por el usuario>';

  const GLOSAS = [
    {
      id: 'pago-exceso', tipo: 'Pago en exceso', variante: '',
      texto: `Corresponde a pago en exceso de la factura N°${C} de fecha ${C} por $${C}, realizado a través de depósito y/ o transferencia Banco ${C} el día ${C}, por un monto de $${C}, quedando un saldo a favor de $${C}.`,
    },
    {
      id: 'pago-duplicado', tipo: 'Pago duplicado', variante: '',
      texto: `Corresponde al pago duplicado de la factura Nº ${C} de fecha ${C} por un monto de $${C}, realizado por segunda vez a través de depósito y/ o transferencia Banco ${C} el día ${C} por un monto de $${C}.`,
    },
    {
      id: 'deposito-sin-identificar', tipo: 'Depósito sin identificar', variante: '',
      texto: `Corresponde a transferencia de fecha ${C} realizada en el Banco ${C} por un monto de $${C}. Usuaria transfirió por error a la cuenta corriente de la DGAC, no existe documento de cobro.`,
    },
    {
      id: 'nota-credito', tipo: 'Nota de Crédito', variante: '',
      texto: `Corresponde a nota de crédito N°${C} de fecha ${C} por $${C}, asociada a la factura N°${C} de fecha ${C} por $${C}, por ${C}. Factura pagada con fecha ${C} en el Banco ${C} por un monto de $${C}.`,
    },
    {
      id: 'error-pago-regularizacion', tipo: 'Error en pago', variante: 'no aplicado a factura de regularización',
      texto: `Corresponde a transferencia de fecha ${C} realizado en el Banco Chile por un monto de $${C}. Por error, pago no fue aplicado a la factura de regularización N°${C} y se envió a ingresos percibidos en exceso en el sistema CGU.`,
    },
    {
      id: 'error-pago-fee', tipo: 'Error en pago', variante: 'Unidad realiza FEE para regularizar',
      texto: `Corresponde a transferencia de fecha ${C} realizada en ${C} por un monto de $${C}. La Unidad de ${C} realiza FEE N°${C} para regularizar depósito.`,
    },
    {
      id: 'error-pago-recibo', tipo: 'Error en pago', variante: 'reemplazo por recibo de tasas (error en pasajeros)',
      texto: `Corresponde al pago de ${C}, según ${C} N°${C} de fecha ${C} por un monto de $${C}, pagado a través de transferencia Banco Chile el día ${C}, el cual fue anulado del sistema SAC y reemplazado por recibo de tasas aeronáuticas N°${C} de fecha ${C} por $${C}, por error en la cantidad de pasajeros, quedando un saldo a favor del cliente de $${C}.`,
    },
    {
      id: 'regularizacion', tipo: 'Regularización', variante: '',
      texto: `Corresponde a transferencia de fecha ${C} realizada en el Banco Chile por un monto de $${C}.- Con fecha ${C} la sección Ventas informó mediante correo electrónico emisión de la factura N°${C}, que regulariza dicha transferencia.`,
    },
    {
      id: 'anulacion-rle', tipo: 'Anulación RLE', variante: '',
      texto: `Corresponde al pago de licencia piloto transporte línea aérea, según recibo de tasas aeronáuticas N°${C} de fecha ${C} por un monto de $${C}, pagado en forma electrónica a través del Controlador de Pago, el cual fue anulado del sistema SAC, ya que el usuario realizó el pago de forma incorrecta, quedando un saldo a favor de $${C}.`,
    },
  ];

  const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

  /** Glosas que corresponden al TIPO DE IPE escrito en el Excel (puede haber varias variantes). */
  function porTipo(tipo) {
    const t = norm(tipo);
    if (!t) return [];
    return GLOSAS.filter((g) => norm(g.tipo) === t);
  }

  function porId(id) {
    return GLOSAS.find((g) => g.id === id) || null;
  }

  function etiqueta(g) {
    return g.variante ? `${g.tipo} — ${g.variante}` : g.tipo;
  }

  /** Cantidad de datos que faltan por completar en un texto. */
  function pendientes(texto) {
    return (String(texto ?? '').match(/<completar por el usuario>/g) || []).length;
  }

  return { GLOSAS, PLACEHOLDER: C, porTipo, porId, etiqueta, pendientes };
});
