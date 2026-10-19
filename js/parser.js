// Lectura del Excel IPE: ubica la fila de encabezados y mapea columnas por nombre (no por posición).
(function (root, factory) {
  const fmt = typeof require === 'function' ? require('./format.js') : root.IPE.format;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).parser = api;
})(typeof self !== 'undefined' ? self : this, function (fmt) {
  // clave interna -> encabezados aceptados (ya normalizados con fmt.normalizeHeader)
  const COLUMNS = {
    estado: ['MIGESTION', 'MISAPUNTES', 'MISAPUNTE', 'APUNTES', 'ESTADO'],
    tipoIpe: ['TIPODEIPE', 'TIPOIPE'],
    responsable: ['RESPONSABLE'],
    concesion: ['CONCESION'],
    cf: ['CF'],
    rut: ['CODAUX', 'RUT'],
    dv: ['DV'],
    razon: ['RAZONSOCIAL'],
    doc: ['DOC'],
    nro: ['NRO'],
    tipo: ['TIPO'],
    comprob: ['COMPROB'],
    fecha: ['FECHA'],
    glosa: ['GLOSAAUXILIAR', 'GLOSA'],
    debe: ['DEBE'],
    haber: ['HABER'],
    saldo: ['SALDO'],
    naturaleza: ['XX'],
    justificacion: ['JUSTIFICACIONINGRESOANTICIPADO', 'JUSTIFICACION'],
    gestion: ['GESTIONREALIZADA'],
    doctoAplicar: ['DOCTOAPLICAR', 'DOCTOAAPLICAR', 'DOCUMENTOAAPLICAR'],
    montoAplicar: ['MONTOAPLICAR', 'MONTOAAPLICAR'],
  };
  // MIS APUNTES no es obligatoria: sin ella se muestran resumen y alertas, pero no memos ni devoluciones.
  const REQUIRED = ['rut', 'dv', 'razon', 'doc', 'nro', 'saldo'];
  const NOMBRES = { rut: 'COD. AUX.', dv: 'DV', razon: 'RAZÓN SOCIAL', doc: 'DOC.', nro: 'NRO.', saldo: 'SALDO', estado: 'MI GESTION' };

  /** Estado normalizado de la columna MIS APUNTES. */
  function categoria(estado) {
    const s = String(estado ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/\s+/g, ' ').trim();
    if (!s) return 'SIN ESTADO';
    if (s === 'APLICAR') return 'APLICAR';
    if (s === 'DEVOLVER') return 'DEVOLVER';
    if (s.startsWith('CORREO ENVIADO')) return 'CORREO ENVIADO';
    if (s.startsWith('REVISAR CORREO')) return 'REVISAR CORREO';
    if (s.startsWith('SIN CONTACTO')) return 'SIN CONTACTO';
    if (s.startsWith('MEMO')) return 'MEMO EMITIDO';
    if (s.startsWith('DEPOSITO SIN IDENTIFICAR')) return 'DEPÓSITO SIN IDENTIFICAR';
    return 'OTRO';
  }

  function str(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmt.fechaCorta(v);
    return String(v).replace(/\s+/g, ' ').trim();
  }

  function findHeaderRow(aoa) {
    for (let i = 0; i < Math.min(aoa.length, 30); i++) {
      const norm = (aoa[i] || []).map(fmt.normalizeHeader);
      if (norm.includes('RAZONSOCIAL') && norm.includes('SALDO')) return i;
    }
    return -1;
  }

  /**
   * aoa: arreglo de filas (cada fila un arreglo de celdas), como lo entrega
   * XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true }).
   */
  function mapColumns(headerRow) {
    const headers = headerRow.map(fmt.normalizeHeader);
    const idx = {};
    for (const [key, names] of Object.entries(COLUMNS)) {
      const i = headers.findIndex((x) => names.includes(x));
      if (i >= 0) idx[key] = i;
    }
    return idx;
  }

  /**
   * Si la columna ACREEDOR/DEUDOR no tiene encabezado (en algunos meses viene sin "XX"),
   * se reconoce por sus valores en las filas siguientes al encabezado.
   */
  function detectarNaturaleza(aoa, h, idx) {
    if (idx.naturaleza !== undefined) return;
    const usadas = new Set(Object.values(idx));
    const muestra = aoa.slice(h + 1, h + 101);
    const ancho = Math.max(0, ...muestra.map((r) => (r || []).length));
    for (let c = 0; c < ancho; c++) {
      if (usadas.has(c) || String((aoa[h] || [])[c] ?? '').trim()) continue;
      const valores = muestra.map((r) => String((r || [])[c] ?? '').trim().toUpperCase()).filter(Boolean);
      if (valores.length && valores.every((v) => v === 'ACREEDOR' || v === 'DEUDOR')) {
        idx.naturaleza = c;
        return;
      }
    }
  }

  /**
   * Evalúa si una hoja tiene la estructura IPE mirando solo sus primeras filas.
   * Devuelve { valida, tieneEstado, trabajo, faltan }.
   */
  function analizarHoja(aoaInicio) {
    const h = findHeaderRow(aoaInicio);
    if (h < 0) return { valida: false, tieneEstado: false, trabajo: 0, faltan: ['RAZÓN SOCIAL', 'SALDO'] };
    const idx = mapColumns(aoaInicio[h]);
    detectarNaturaleza(aoaInicio, h, idx);
    const faltan = REQUIRED.filter((k) => idx[k] === undefined).map((k) => NOMBRES[k]);
    // Columnas propias de la hoja de trabajo (no del mayor contable): sirven para desempatar.
    const trabajo = ['estado', 'tipoIpe', 'responsable', 'justificacion', 'gestion', 'naturaleza', 'doctoAplicar', 'montoAplicar']
      .filter((k) => idx[k] !== undefined).length;
    return { valida: faltan.length === 0, tieneEstado: idx.estado !== undefined, trabajo, faltan };
  }

  /**
   * Elige la hoja de trabajo entre todas las del libro: primero las que tienen la estructura IPE
   * y la columna MIS APUNTES; si ninguna la tiene, la primera con estructura IPE.
   * hojas: [{ nombre, inicio }] donde inicio son las primeras filas (aoa).
   */
  function elegirHoja(hojas) {
    const evaluadas = hojas.map((h) => ({ nombre: h.nombre, ...analizarHoja(h.inicio) }));
    const validas = evaluadas.filter((h) => h.valida);
    // Orden estable: primero con MIS APUNTES, luego más columnas de trabajo, luego el orden del libro.
    const ordenadas = validas
      .map((h, i) => ({ h, i }))
      .sort((a, b) => (b.h.tieneEstado - a.h.tieneEstado) || (b.h.trabajo - a.h.trabajo) || (a.i - b.i))
      .map((x) => x.h);
    return { elegida: ordenadas[0] || null, candidatas: ordenadas, evaluadas };
  }

  function parseAoa(aoa) {
    const h = findHeaderRow(aoa);
    if (h < 0) throw new Error('No se encontró la fila de encabezados (debe contener "RAZÓN SOCIAL" y "SALDO").');
    const idx = mapColumns(aoa[h]);
    detectarNaturaleza(aoa, h, idx);
    const missing = REQUIRED.filter((k) => idx[k] === undefined);
    if (missing.length) throw new Error(`Faltan columnas obligatorias en el Excel: ${missing.map((k) => NOMBRES[k]).join(', ')}`);

    const rows = [];
    for (let r = h + 1; r < aoa.length; r++) {
      const line = aoa[r] || [];
      const get = (k) => (idx[k] === undefined ? undefined : line[idx[k]]);
      const rut = str(get('rut'));
      const razon = str(get('razon'));
      if (!rut && !razon) continue;
      const naturaleza = str(get('naturaleza')).toUpperCase();
      const saldo = Math.abs(fmt.toNumber(get('saldo')));
      rows.push({
        fila: r + 1, // número de fila en Excel
        estado: str(get('estado')),
        categoria: categoria(get('estado')),
        tipoIpe: str(get('tipoIpe')),
        responsable: str(get('responsable')),
        cf: str(get('cf')),
        rut,
        dv: str(get('dv')).toUpperCase(),
        rutFmt: fmt.formatRut(rut, get('dv')),
        razon,
        doc: str(get('doc')),
        nro: str(get('nro')),
        tipo: str(get('tipo')),
        comprob: str(get('comprob')),
        fecha: fmt.parseFecha(get('fecha')),
        glosa: str(get('glosa')),
        debe: fmt.toNumber(get('debe')),
        haber: fmt.toNumber(get('haber')),
        saldo,
        naturaleza,
        // ACREEDOR suma a favor del cliente; DEUDOR resta.
        saldoNeto: naturaleza === 'DEUDOR' ? -saldo : saldo,
        justificacion: str(get('justificacion')),
        gestion: str(get('gestion')),
        doctoAplicar: str(get('doctoAplicar')),
        montoAplicar: get('montoAplicar') ?? '',
      });
    }
    return {
      rows,
      filaEncabezado: h,
      columnas: idx,
      tieneEstado: idx.estado !== undefined,
      tieneTipo: idx.tipoIpe !== undefined,
      tieneColumnasAplicar: idx.doctoAplicar !== undefined && idx.montoAplicar !== undefined,
    };
  }

  /**
   * Copia MIS APUNTES desde los registros de otro mes (normalmente el anterior).
   * Coincidencia exacta por RUT + documento + número + comprobante; si no hay, por RUT + documento + número.
   * Devuelve { rows, exactos, porDocumento, sinCoincidencia } sin modificar los registros de entrada.
   */
  function copiarEstados(destino, origen) {
    const clave = (r) => `${r.rut}-${r.dv}|${r.doc}|${r.nro}|${r.comprob}`;
    const claveDoc = (r) => `${r.rut}-${r.dv}|${r.doc}|${r.nro}`;
    const conEstado = origen.filter((r) => r.estado);
    const exacta = new Map(conEstado.map((r) => [clave(r), r]));
    const porDoc = new Map(conEstado.map((r) => [claveDoc(r), r]));
    let exactos = 0;
    let porDocumento = 0;
    const rows = destino.map((r) => {
      let previo = exacta.get(clave(r));
      if (previo) exactos++;
      else if ((previo = porDoc.get(claveDoc(r)))) porDocumento++;
      if (!previo) return { ...r, estadoCopiado: false };
      return { ...r, estado: previo.estado, categoria: categoria(previo.estado), tipoIpe: r.tipoIpe || previo.tipoIpe || '', estadoCopiado: true };
    });
    return { rows, exactos, porDocumento, sinCoincidencia: rows.length - exactos - porDocumento };
  }

  /**
   * Cuando un mismo IPE (empresa + documento + número) tiene movimientos DEUDOR y ACREEDOR, se muestra
   * un solo registro con el resultado: suma ACREEDOR menos suma DEUDOR. Los datos (estado, tipo,
   * justificación, comprobante, fecha) se toman del primer registro ACREEDOR.
   */
  function consolidar(rows) {
    const grupos = new Map();
    for (const r of rows) {
      const k = `${r.rut}-${r.dv}|${r.doc}|${r.nro}`;
      if (!grupos.has(k)) grupos.set(k, []);
      grupos.get(k).push(r);
    }
    const salida = [];
    const hechos = new Set();
    for (const r of rows) {
      const k = `${r.rut}-${r.dv}|${r.doc}|${r.nro}`;
      const g = grupos.get(k);
      const acreedores = g.filter((x) => x.naturaleza !== 'DEUDOR');
      const deudores = g.filter((x) => x.naturaleza === 'DEUDOR');
      if (!deudores.length || !acreedores.length) { salida.push(r); continue; }
      if (hechos.has(k)) continue;
      hechos.add(k);
      const base = acreedores[0];
      const primero = (campo) => (acreedores.find((x) => x[campo]) || g.find((x) => x[campo]) || base)[campo];
      const neto = acreedores.reduce((s, x) => s + x.saldo, 0) - deudores.reduce((s, x) => s + x.saldo, 0);
      salida.push({
        ...base,
        estado: primero('estado'),
        categoria: categoria(primero('estado')),
        tipoIpe: primero('tipoIpe'),
        justificacion: primero('justificacion'),
        gestion: primero('gestion'),
        saldo: Math.abs(neto),
        saldoNeto: neto,
        naturaleza: neto < 0 ? 'DEUDOR' : 'ACREEDOR',
        consolidado: { filas: g.map((x) => x.fila), acreedor: neto + deudores.reduce((s, x) => s + x.saldo, 0), deudor: deudores.reduce((s, x) => s + x.saldo, 0) },
      });
    }
    return salida;
  }

  /** Último paso numerado de "Gestión realizada" (ej. "19) Correo enviado..."). */
  function ultimaGestion(texto) {
    const s = String(texto ?? '').trim();
    if (!s) return '';
    const partes = s.split(/(?:^|\s)(?=\d{1,2}\s?\)\s?)/).map((x) => x.trim()).filter(Boolean);
    return partes[partes.length - 1].replace(/^\d{1,2}\s?\)\s?/, '');
  }

  return { parseAoa, analizarHoja, elegirHoja, copiarEstados, consolidar, categoria, ultimaGestion, COLUMNS };
});
