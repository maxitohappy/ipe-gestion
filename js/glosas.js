// Glosas por TIPO DE IPE (documento GLOSAS). Se usan como MOTIVO del memo y como OBSERVACIONES
// de la devolución. Los datos que completa el usuario se escriben [[así]] y en el Word aparecen
// como [así] con fondo amarillo.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).glosas = api;
})(typeof self !== 'undefined' ? self : this, function () {
  const GLOSAS = [
    {
      id: 'pago-exceso', tipo: 'Pago en exceso', variante: '',
      texto: 'Corresponde a pago en exceso de la factura N°[[N° factura]] de fecha [[fecha factura]] por $[[monto factura]], realizado a través de depósito y/ o transferencia Banco [[banco]] el día [[fecha pago]], por un monto de $[[monto pagado]], quedando un saldo a favor de $[[saldo a favor]].',
    },
    {
      id: 'pago-duplicado', tipo: 'Pago duplicado', variante: '',
      texto: 'Corresponde al pago duplicado de la factura Nº [[N° factura]] de fecha [[fecha factura]] por un monto de $[[monto factura]], realizado por segunda vez a través de depósito y/ o transferencia Banco [[banco]] el día [[fecha pago]] por un monto de $[[monto pagado]].',
    },
    {
      id: 'deposito-sin-identificar', tipo: 'Depósito sin identificar', variante: '',
      texto: 'Corresponde a transferencia de fecha [[fecha transferencia]] realizada en el Banco [[banco]] por un monto de $[[monto]]. Usuaria transfirió por error a la cuenta corriente de la DGAC, no existe documento de cobro.',
    },
    {
      id: 'nota-credito', tipo: 'Nota de Crédito', variante: '',
      texto: 'Corresponde a nota de crédito N°[[N° nota de crédito]] de fecha [[fecha nota de crédito]] por $[[monto nota de crédito]], asociada a la factura N°[[N° factura]] de fecha [[fecha factura]] por $[[monto factura]], por [[motivo nota de crédito]]. Factura pagada con fecha [[fecha pago]] en el Banco [[banco]] por un monto de $[[monto pagado]].',
    },
    {
      id: 'error-pago-regularizacion', tipo: 'Error en pago', variante: '1',
      texto: 'Corresponde a transferencia de fecha [[fecha transferencia]] realizado en el Banco Chile por un monto de $[[monto]]. Por error, pago no fue aplicado a la factura de regularización N°[[N° factura]] y se envió a ingresos percibidos en exceso en el sistema CGU.',
    },
    {
      id: 'error-pago-fee', tipo: 'Error en pago', variante: '2',
      texto: 'Corresponde a transferencia de fecha [[fecha transferencia]] realizada en [[banco]] por un monto de $[[monto]]. La Unidad de [[unidad]] realiza FEE N°[[N° FEE]] para regularizar depósito.',
    },
    {
      id: 'error-pago-recibo', tipo: 'Error en pago', variante: '3',
      texto: 'Corresponde al pago de [[concepto]], según [[tipo de documento]] N°[[N° documento]] de fecha [[fecha documento]] por un monto de $[[monto documento]], pagado a través de transferencia Banco Chile el día [[fecha pago]], el cual fue anulado del sistema SAC y reemplazado por recibo de tasas aeronáuticas N°[[N° recibo]] de fecha [[fecha recibo]] por $[[monto recibo]], por error en la cantidad de pasajeros, quedando un saldo a favor del cliente de $[[saldo a favor]].',
    },
    {
      id: 'regularizacion', tipo: 'Regularización', variante: '',
      texto: 'Corresponde a transferencia de fecha [[fecha transferencia]] realizada en el Banco Chile por un monto de $[[monto]].- Con fecha [[fecha correo Ventas]] la sección Ventas informó mediante correo electrónico emisión de la factura N°[[N° factura]], que regulariza dicha transferencia.',
    },
    {
      id: 'anulacion-rle', tipo: 'Anulación RLE', variante: '',
      texto: 'Corresponde al pago de licencia piloto transporte línea aérea, según recibo de tasas aeronáuticas N°[[N° recibo]] de fecha [[fecha recibo]] por un monto de $[[monto]], pagado en forma electrónica a través del Controlador de Pago, el cual fue anulado del sistema SAC, ya que el usuario realizó el pago de forma incorrecta, quedando un saldo a favor de $[[saldo a favor]].',
    },
  ];

  const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

  /**
   * Glosa que corresponde al TIPO DE IPE escrito en el Excel. "Error en pago" tiene tres redacciones:
   * "Error en pago" o "Error en pago 1" usan la primera; "Error en pago 2" y "Error en pago 3", las otras.
   */
  function porTipo(tipo) {
    const t = norm(tipo);
    if (!t) return null;
    const m = t.match(/^(.*?)\s*(\d)$/);
    const base = m ? m[1] : t;
    const variantes = GLOSAS.filter((g) => norm(g.tipo) === base);
    if (!variantes.length) return null;
    return (m && variantes.find((g) => g.variante === m[2])) || variantes[0];
  }

  /** Divide un texto en partes normales (n) y datos a completar (m), para resaltarlos en el Word. */
  function partes(texto) {
    const out = [];
    const re = /\[\[([^\]]+)\]\]|<completar por el usuario>/g;
    const s = String(texto ?? '');
    let ultimo = 0;
    let m;
    while ((m = re.exec(s))) {
      if (m.index > ultimo) out.push({ n: true, t: s.slice(ultimo, m.index) });
      out.push({ m: true, t: `[${m[1] || 'completar'}]` });
      ultimo = re.lastIndex;
    }
    if (ultimo < s.length) out.push({ n: true, t: s.slice(ultimo) });
    return out;
  }

  /** Texto plano con los datos a completar como [dato] (para vistas previas). */
  function plano(texto) {
    return partes(texto).map((p) => p.t).join('');
  }

  /** Cantidad de datos que faltan por completar en un texto. */
  function pendientes(texto) {
    return partes(texto).filter((p) => p.m).length;
  }

  return { GLOSAS, porTipo, partes, plano, pendientes };
});
