// Interpreta lo que Claude leyó de las imágenes de correos: documentos a aplicar y datos bancarios.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).extraccion = api;
})(typeof self !== 'undefined' ? self : this, function () {
  const normalizar = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
  const rutLimpio = (s) => String(s ?? '').toUpperCase().replace(/[^0-9K]/g, '');
  const rutDeGrupo = (g) => rutLimpio(`${g.rut}${g.dv}`);

  function rutValido(rut) {
    const limpio = rutLimpio(rut);
    if (limpio.length < 2) return false;
    let suma = 0;
    let factor = 2;
    for (let i = limpio.length - 2; i >= 0; i--) {
      suma += Number(limpio[i]) * factor;
      factor = factor === 7 ? 2 : factor + 1;
    }
    const dv = 11 - (suma % 11);
    return limpio.slice(-1) === (dv === 11 ? '0' : dv === 10 ? 'K' : String(dv));
  }

  // ---------- documentos a aplicar ----------
  /** Nombre del documento como en los formatos: "FEE 464037", "ND N° 54549" (año en curso) o "NDE 54195" (años anteriores). */
  function nombreDocumento(d, periodo) {
    const t = String(d.tipo).toUpperCase();
    if (t === 'FE' || t === 'FEE') return `FEE ${d.numero}`;
    if (t === 'ND' || t === 'NDE') return periodo === 'anterior' ? `NDE ${d.numero}` : `ND N° ${d.numero}`;
    return `${t} ${d.numero}`;
  }

  /**
   * Asigna los documentos a los IPE, en orden: cada IPE se va descontando documento por documento.
   * "Pagar" si cubre lo que falta del documento; "Abonar" si solo alcanza para una parte (lo que falta pasa al IPE
   * siguiente). `ipes`: [{ saldo, periodo }]. Devuelve las aplicaciones de cada IPE y lo que no alcanzó a cubrirse.
   */
  function asignar(ipes, docs) {
    const pendientes = docs.filter((d) => Number(d.monto) > 0).map((d) => ({ d, resta: Number(d.monto) }));
    let i = 0;
    const aplicaciones = ipes.map((ipe) => {
      let saldo = ipe.saldo;
      const lista = [];
      while (saldo > 0 && i < pendientes.length) {
        const p = pendientes[i];
        const monto = Math.min(saldo, p.resta);
        const cubre = monto === p.resta;
        lista.push({ docto: `${cubre ? 'Pagar' : 'Abonar'} ${nombreDocumento(p.d, ipe.periodo)}`, monto });
        saldo -= monto;
        p.resta -= monto;
        if (cubre) i += 1;
      }
      return lista;
    });
    return { aplicaciones, sinCubrir: pendientes.slice(i) };
  }

  /** Tarjetas del cliente del correo: por RUT y, si no se leyó, por razón social. Años anteriores primero. */
  function gruposDelCliente(lectura, grupos) {
    const rut = rutLimpio(lectura.rut_cliente);
    const razon = normalizar(lectura.razon_social);
    let hallados = rut ? grupos.filter((g) => rutDeGrupo(g) === rut) : [];
    if (!hallados.length && razon) hallados = grupos.filter((g) => normalizar(g.razon) === razon);
    const peso = (g) => (g.periodo === 'anterior' ? 0 : 1);
    return [...hallados].sort((a, b) => peso(a) - peso(b));
  }

  /** Advertencias de una lectura de documentos, ya asignada. `saldoIpe`: saldo sumado de los IPE incluidos. */
  function avisosDocumentos(lectura, destino, saldoIpe, sinCubrir, fmtCLP) {
    const a = [];
    if (!lectura.contiene_documentos) a.push('La imagen no parece un correo con documentos por pagar.');
    if (lectura.rut_cliente && rutLimpio(lectura.rut_cliente) !== rutDeGrupo(destino)) a.push(`El RUT del correo (${lectura.rut_cliente}) no coincide con el de este cliente.`);
    if (lectura.saldo_a_favor !== null && lectura.saldo_a_favor !== undefined && lectura.saldo_a_favor !== saldoIpe) {
      a.push(`El correo informa un saldo a favor de ${fmtCLP(lectura.saldo_a_favor)}; el saldo de los IPE incluidos es ${fmtCLP(saldoIpe)}.`);
    }
    const sinMonto = (lectura.documentos || []).filter((d) => !(Number(d.monto) > 0));
    if (sinMonto.length) a.push(`No se leyó el monto de: ${sinMonto.map((d) => `${d.tipo} ${d.numero}`).join(', ')}.`);
    if (sinCubrir.length) a.push(`El saldo no alcanzó para: ${sinCubrir.map((p) => `${p.d.tipo} ${p.d.numero} (falta ${fmtCLP(p.resta)})`).join(', ')}.`);
    if (lectura.imagen_cortada) a.push('La imagen parece cortada: revisa si faltan documentos.');
    if (lectura.observaciones) a.push(lectura.observaciones);
    return a;
  }

  // ---------- datos bancarios ----------
  // [expresión sobre el nombre normalizado, nombre en el selector de bancos]
  const BANCOS = [
    [/\bINTERNACIONAL\b/, 'Banco Internacional'],
    [/\bESTADO\b|\bBANCOESTADO\b/, 'BancoEstado'],
    [/\bEDWARDS\b/, 'Banco Edwards'],
    [/\bNACION ARGENTINA\b/, 'Banco de la Nación Argentina'],
    [/\bBANK OF CHINA\b/, 'Bank of China'],
    [/\bCHINA CONSTRUCTION\b/, 'China Construction Bank'],
    [/\bBANCO DE CHILE\b|^CHILE$/, 'Banco de Chile'],
    [/\bSANTANDER\b/, 'Banco Santander'],
    [/\bBCI\b|\bCREDITO E INVERSIONES\b/, 'Banco BCI'],
    [/\bSCOTIA/, 'Scotiabank'],
    [/\bITAU\b/, 'Banco Itaú'],
    [/\bBICE\b/, 'Banco BICE'],
    [/\bSECURITY\b/, 'Banco Security'],
    [/\bFALABELLA\b/, 'Banco Falabella'],
    [/\bRIPLEY\b/, 'Banco Ripley'],
    [/\bCONSORCIO\b/, 'Banco Consorcio'],
    [/\bHSBC\b/, 'HSBC Bank Chile'],
    [/\bBTG\b/, 'Banco BTG Pactual'],
    [/\bJP ?MORGAN\b/, 'JP Morgan Chase Bank'],
    [/\bMUFG\b/, 'MUFG Bank'],
    [/\bTENPO\b/, 'Tenpo'],
    [/\bCOOPEUCH\b/, 'Coopeuch'],
    [/\bMERCADO ?PAGO\b/, 'Mercado Pago'],
    [/\bLOS HEROES\b/, 'Prepago Los Héroes'],
  ];

  /** Valor del selector de bancos para un nombre leído; si no está en la lista queda en "Otro". */
  function bancoSelector(nombre) {
    const n = normalizar(nombre);
    if (!n) return { banco: '', bancoOtro: '' };
    const hallado = BANCOS.find(([re]) => re.test(n));
    return hallado ? { banco: hallado[1], bancoOtro: '' } : { banco: 'Otro', bancoOtro: String(nombre).trim() };
  }

  function correosCliente(lista) {
    const limpios = (lista || []).map((c) => String(c).trim().toLowerCase())
      .filter((c) => c.includes('@') && !/@(?:[\w-]+\.)*dgac\.(?:gob\.)?cl$/.test(c));
    return [...new Set(limpios)].join('\n');
  }

  /** Lo leído convertido a los campos de la tarjeta; los campos no leídos conservan lo que ya había. */
  function aplicarBanco(actual, lectura) {
    const sel = bancoSelector(lectura.banco);
    return {
      ...actual,
      tipoCuenta: lectura.tipo_cuenta || actual.tipoCuenta,
      numeroCuenta: lectura.numero_cuenta || actual.numeroCuenta,
      banco: sel.banco || actual.banco,
      bancoOtro: sel.banco ? sel.bancoOtro : actual.bancoOtro,
      correos: correosCliente(lectura.correos) || actual.correos,
    };
  }

  function avisosBanco(lectura, g) {
    const a = [];
    if (!lectura.contiene_datos_bancarios) a.push('La imagen no parece un correo con datos bancarios.');
    if (!lectura.numero_cuenta) a.push('No se leyó el número de cuenta.');
    if (!lectura.tipo_cuenta) a.push('No se leyó el tipo de cuenta: revisa el que aparece seleccionado.');
    if (!lectura.banco) a.push('No se leyó el banco.');
    const rut = lectura.rut_beneficiario;
    if (rut && rutLimpio(rut) !== rutDeGrupo(g)) a.push(`El beneficiario (${lectura.beneficiario || 'sin nombre'}, RUT ${rut}) es distinto del cliente del IPE.`);
    if (rut && !rutValido(rut)) a.push(`El RUT ${rut} no tiene un dígito verificador válido.`);
    if (lectura.observaciones) a.push(lectura.observaciones);
    return a;
  }

  return { rutValido, nombreDocumento, asignar, gruposDelCliente, avisosDocumentos, bancoSelector, correosCliente, aplicarBanco, avisosBanco };
});
