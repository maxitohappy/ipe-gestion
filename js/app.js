// Interfaz: carga del Excel, pestañas, documentos Word (memos y devoluciones) e historiales.
(function () {
  const { format: fmt, parser, memo, devoluciones, resumen, movimientos } = window.IPE;
  const $ = (id) => document.getElementById(id);
  const CONFIG_KEY = 'ipe-config';
  const HIST_KEY = 'ipe-historial';
  const HIST_DEV_KEY = 'ipe-historial-devoluciones';
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const state = { rowsBase: [], rows: [], memoGroups: [], devGroups: [], historial: [], historialDev: [], plantillas: {} };

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clp = (n) => fmt.formatCLP(n);
  const hoyISO = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const BANDERAS = { roja: 'Ingreso de años anteriores', amarilla: 'Ingreso del 1er semestre del año en curso', verde: 'Ingreso del 2º semestre del año en curso' };
  const flagHtml = (fecha) => {
    const b = fmt.bandera(fecha);
    return b ? `<span class="flag ${b}" title="${BANDERAS[b]}" role="img" aria-label="Bandera ${b}: ${BANDERAS[b]}"></span>` : '';
  };

  // ---------- configuración e historiales (se guardan en este navegador) ----------
  function readStore(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch (e) { return fallback; }
  }
  function writeStore(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
  }
  function loadConfig() {
    const cfg = readStore(CONFIG_KEY, {});
    $('cfg-iniciales').value = cfg.iniciales ?? 'ERS/yaa';
    $('cfg-fecha').value = hoyISO();
  }
  function saveConfig() {
    writeStore(CONFIG_KEY, { iniciales: $('cfg-iniciales').value });
  }
  function docOpts() {
    const [y, m, d] = $('cfg-fecha').value.split('-').map(Number);
    return { fecha: y ? new Date(y, m - 1, d) : new Date(), iniciales: $('cfg-iniciales').value.trim() };
  }
  function loadHistoriales() {
    const h = readStore(HIST_KEY, []);
    const hd = readStore(HIST_DEV_KEY, []);
    state.historial = Array.isArray(h) ? h : [];
    state.historialDev = Array.isArray(hd) ? hd : [];
  }
  function saveHistoriales() {
    const ok = writeStore(HIST_KEY, state.historial) && writeStore(HIST_DEV_KEY, state.historialDev);
    if (!ok) alert('No se pudo guardar el historial en este navegador. Descarga un respaldo para no perderlo.');
  }

  // ---------- carga del Excel ----------
  const MSG_LEYENDO = 'Leyendo el Excel… (los libros con muchas hojas pueden tardar unos segundos)';
  const pausa = () => new Promise((r) => setTimeout(r, 50)); // deja que el navegador muestre el mensaje

  /** Lee un libro: revisa las primeras filas de cada hoja y ordena las hojas con estructura IPE. */
  async function leerLibro(file) {
    const buf = await file.arrayBuffer();
    const cabeceras = XLSX.read(buf, { cellDates: true, sheetRows: 30 });
    const hojas = cabeceras.SheetNames.map((nombre) => ({
      nombre,
      inicio: XLSX.utils.sheet_to_json(cabeceras.Sheets[nombre], { header: 1, raw: true, defval: '' }),
    }));
    const { elegida, candidatas, evaluadas } = parser.elegirHoja(hojas);
    if (!elegida) {
      const detalle = evaluadas.map((h) => `"${h.nombre}" (faltan: ${h.faltan.join(', ')})`).join('; ');
      throw new Error(`Ninguna hoja del Excel "${file.name}" tiene la estructura IPE. Hojas revisadas: ${detalle}.`);
    }
    return { buf, nombreArchivo: file.name, elegida, candidatas, cache: {} };
  }

  /** Lectura completa de una sola hoja (los libros pueden traer hojas con más de un millón de filas). */
  function leerHoja(libro, nombre) {
    if (!libro.cache[nombre]) {
      const wb = XLSX.read(libro.buf, { cellDates: true, dense: true, sheets: nombre });
      libro.cache[nombre] = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: '', blankrows: false });
    }
    return libro.cache[nombre];
  }

  function mostrarError(err) {
    console.error(err);
    $('estado-carga').innerHTML = `<span style="color:var(--danger)">${esc(err.message)}</span>`;
  }

  async function handleFile(file) {
    if (!file) return;
    $('estado-carga').textContent = MSG_LEYENDO;
    await pausa();
    try {
      state.libro = await leerLibro(file);
      cargarHoja(state.libro.elegida.nombre);
    } catch (err) {
      $('selector-hoja').hidden = true;
      $('aviso-estado').hidden = true;
      $('info-copia').hidden = true;
      mostrarError(err);
    }
  }

  function cargarHoja(nombre) {
    const parsed = parser.parseAoa(leerHoja(state.libro, nombre));
    state.hoja = nombre;
    state.parsed = parsed;
    state.copia = null;
    $('estado-carga').innerHTML = `<span class="ok">✔</span> ${esc(state.libro.nombreArchivo)} — hoja "${esc(nombre)}", ${parsed.rows.length} registros.`;
    $('sel-hoja').innerHTML = state.libro.candidatas.map((h) =>
      `<option value="${esc(h.nombre)}" ${h.nombre === nombre ? 'selected' : ''}>${esc(h.nombre)}${h.tieneEstado ? '' : ' (sin MI GESTION)'}</option>`).join('');
    $('selector-hoja').hidden = state.libro.candidatas.length < 2;
    usarFilas(parsed.rows, parsed.tieneEstado);
  }

  /** Copia MI GESTION y TIPO DE IPE desde el Excel de otro mes (el que sí tiene esas columnas). */
  async function copiarApuntes(file) {
    if (!file) return;
    const previo = $('estado-carga').innerHTML;
    $('estado-carga').textContent = MSG_LEYENDO;
    await pausa();
    try {
      const libro = await leerLibro(file);
      const hoja = libro.candidatas.find((h) => h.tieneEstado);
      if (!hoja) throw new Error(`"${file.name}" no tiene la columna MI GESTION: elige el Excel del mes anterior en que la completaste.`);
      const origen = parser.parseAoa(leerHoja(libro, hoja.nombre)).rows;
      const res = parser.copiarEstados(state.parsed.rows, origen);
      state.copia = { archivo: file.name, hoja: hoja.nombre, ...res };
      $('estado-carga').innerHTML = previo;
      usarFilas(res.rows, true);
    } catch (err) {
      mostrarError(err);
    }
  }

  function usarFilas(rowsBase, tieneEstado) {
    state.rowsBase = rowsBase;
    // Un mismo IPE con movimientos DEUDOR y ACREEDOR se muestra como un solo registro con el saldo neto.
    state.rows = parser.consolidar(rowsBase);
    state.tieneEstado = tieneEstado;
    reconstruir();

    $('aviso-estado').hidden = tieneEstado;
    $('info-copia').hidden = !state.copia;
    if (state.copia) {
      const c = state.copia;
      $('info-copia-texto').innerHTML =
        `Se copiaron <b>MI GESTION</b> y <b>TIPO DE IPE</b> de <b>${c.exactos + c.porDocumento}</b> de ${rowsBase.length} registros desde "${esc(c.archivo)}" ` +
        `(${c.exactos} con el mismo comprobante y ${c.porDocumento} por documento). <b>${c.sinCoincidencia}</b> registros son nuevos y quedan sin estado.`;
    }
    $('contenido').hidden = false;
    renderTodo();
  }

  function reconstruir() {
    const idx = movimientos.indice(state.historial, state.historialDev);
    state.memoGroups = memo.buildGroups(state.rows, idx);
    state.devGroups = devoluciones.buildGroups(state.rows, idx);
  }

  function renderTodo() {
    renderResumen();
    renderMemos();
    renderDevoluciones();
    renderHistorial();
    renderHistorialDev();
  }

  /** Descarga la hoja de trabajo con TIPO DE IPE (B) y MI GESTION (C), para seguir trabajando en Excel. */
  function descargarHojaTrabajo() {
    const aoa = leerHoja(state.libro, state.hoja);
    const { filaEncabezado: h, columnas } = state.parsed;
    const porFila = new Map(state.rowsBase.map((r) => [r.fila, r]));
    const campos = [{ k: 'tipoIpe', titulo: 'TIPO DE IPE' }, { k: 'estado', titulo: 'MI GESTION' }];
    const faltan = campos.filter((c) => columnas[c.k] === undefined);
    const valor = (i, c) => (i === h ? c.titulo : i > h ? (porFila.get(i + 1)?.[c.k] ?? '') : '');
    const salida = aoa.map((fila, i) => {
      const f = [...fila];
      for (const c of campos) if (columnas[c.k] !== undefined && i >= h) f[columnas[c.k]] = valor(i, c);
      f.splice(1, 0, ...faltan.map((c) => valor(i, c)));
      return f;
    });
    const mover = (c) => (c < 1 ? c : c + faltan.length);
    if (columnas.naturaleza !== undefined && !String(salida[h][mover(columnas.naturaleza)] ?? '').trim()) salida[h][mover(columnas.naturaleza)] = 'XX';
    const ws = XLSX.utils.aoa_to_sheet(salida, { cellDates: true, dateNF: 'dd-mm-yyyy' });
    for (const k of ['debe', 'haber', 'saldo']) {
      if (columnas[k] === undefined) continue;
      for (let r = h + 1; r < salida.length; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c: mover(columnas[k]) })];
        if (cell && cell.t === 'n') cell.z = '"$" #,##0';
      }
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, state.hoja.trim().slice(0, 31) || 'IPE');
    const base = state.libro.nombreArchivo.replace(/\.xls[xm]?$/i, '').replace(/[.\s]+$/, '');
    XLSX.writeFile(wb, `${base} - hoja de trabajo.xlsx`);
  }

  // ---------- resumen ----------
  function renderResumen() {
    const r = resumen.compute(state.rows, state.memoGroups, state.devGroups);
    state.listado = r.listado;
    const kpis = [
      ['Empresas (total, por razón social)', r.empresasTotal],
      ['Empresas en Memos (APLICAR)', r.empresasMemos],
      ['Empresas en Devoluciones', r.empresasDevoluciones],
      ['Empresas pendientes', r.empresasPendientes],
    ];
    $('kpis').innerHTML = kpis.map(([l, v]) => `<div class="kpi"><div class="l">${esc(l)}</div><div class="v">${esc(v)}</div></div>`).join('');
    renderListado();
  }

  function renderListado() {
    const fb = $('filtro-bandera').value;
    const q = $('buscar').value.trim().toLowerCase();
    const lista = (state.listado || []).filter((r) => (!fb || r.bandera === fb)
      && (!q || `${r.razon} ${r.rutFmt} ${r.rut} ${r.doc} ${r.nro}`.toLowerCase().includes(q)));
    $('listado').innerHTML = lista.length
      ? `<table><thead><tr><th class="n">N°</th><th>MI GESTION</th><th>Categoría</th><th>Empresa</th><th>RUT</th><th>Doc.</th><th>Nro.</th><th>Fecha</th><th class="n">Saldo</th></tr></thead><tbody>
        ${lista.map((r, i) => `<tr class="${r.naturaleza === 'DEUDOR' ? 'deudor' : ''}">
          <td class="n">${i + 1}</td><td>${esc(r.estado)}</td><td>${flagHtml(r.fecha)}</td><td>${esc(r.razon)}</td>
          <td class="nowrap">${esc(r.rutFmt)}</td><td>${esc(r.doc)}</td><td>${esc(r.nro)}</td>
          <td class="nowrap">${esc(fmt.fechaCorta(r.fecha))}</td>
          <td class="n">${clp(r.saldo)}${r.consolidado ? `<div class="prev" title="Saldo acreedor menos saldo deudor del mismo IPE">neto: ${clp(r.consolidado.acreedor)} − ${clp(r.consolidado.deudor)}</div>` : ''}</td></tr>`).join('')}
        </tbody></table>`
      : '<p class="muted">No hay ingresos anticipados de empresas pendientes con ese filtro.</p>';
  }

  // ---------- piezas comunes de memos y devoluciones ----------
  /** TIPO DE IPE tal como viene en la planilla (no editable). */
  function tipoHtml(r) {
    if (!r.tipoIpe) return '<span class="muted">vacío</span>';
    const aviso = r.glosaId ? '' : '<div class="prev">no está en las glosas: el Word usa la justificación del Excel</div>';
    return `${esc(r.tipoIpe)}${aviso}`;
  }

  const netoHtml = (r) => (r.consolidado ? `<div class="prev">neto: ${clp(r.consolidado.acreedor)} − ${clp(r.consolidado.deudor)}</div>` : '');

  /** Buscador por RUT o empresa: oculta las tarjetas que no coinciden. */
  function filtrarTarjetas(contenedor, grupos, q, campo) {
    const t = q.trim().toLowerCase();
    const digitos = t.replace(/[^0-9k]/g, '');
    const lista = $(contenedor);
    // Al ocultar tarjetas la página se acorta y el navegador sube la vista. Se mantiene un alto mínimo
    // mientras hay búsqueda y se conserva la posición del buscador en pantalla.
    const antes = campo ? campo.getBoundingClientRect().top : null;
    lista.style.minHeight = t ? `${window.innerHeight}px` : '';
    let visibles = 0;
    document.querySelectorAll(`#${contenedor} [data-g]`).forEach((card) => {
      const g = state[grupos][Number(card.dataset.g)];
      const rut = `${g.rut}${g.dv}`.toLowerCase();
      const ok = !t || g.razon.toLowerCase().includes(t) || g.rutFmt.toLowerCase().includes(t) || (digitos.length >= 3 && rut.includes(digitos));
      card.hidden = !ok;
      if (ok) visibles += 1;
    });
    if (campo) window.scrollBy(0, campo.getBoundingClientRect().top - antes);
    return visibles;
  }

  function alertError(err) {
    console.error(err);
    const detalle = err.properties && err.properties.errors ? err.properties.errors.map((e) => e.properties.explanation).join('\n') : '';
    alert(`No se pudo generar el documento: ${err.message}\n${detalle}`);
  }

  // Las plantillas vienen incrustadas en js/plantilla.js: funcionan en línea y abriendo index.html sin internet.
  function plantilla(nombre) {
    if (!state.plantillas[nombre]) {
      const b64 = window.IPE[nombre];
      if (!b64) throw new Error('No se encontró la plantilla Word (js/plantilla.js).');
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      state.plantillas[nombre] = bytes.buffer;
    }
    return state.plantillas[nombre];
  }
  const renderMemo = (data) => memo.render(plantilla('plantillaB64'), data, PizZip, window.docxtemplater);
  const renderDevolucion = (data) => devoluciones.render(plantilla('plantillaDevolucionB64'), data, PizZip, window.docxtemplater);

  // ---------- memos ----------
  function renderMemos() {
    const groups = state.memoGroups;
    $('n-memos').textContent = groups.length;
    $('memos-resumen').textContent = `${groups.length} empresa(s) con IPE pendientes de aplicación.`;
    $('btn-zip').disabled = !groups.length;
    $('memos').innerHTML = groups.map((g, gi) => memoCardHtml(g, gi)).join('') || '<p class="muted">No hay memos pendientes.</p>';
    groups.forEach((g, gi) => refreshMemo(gi));
    filtrarTarjetas('memos', 'memoGroups', $('buscar-memos').value);
  }

  function montoInput(v) {
    return v === '' || v === null || v === undefined ? '' : fmt.miles(Number(v));
  }

  function memoCardHtml(g, gi) {
    return `
      <div class="card" data-g="${gi}">
        <div class="memo-head">
          <h3>${esc(g.razon)} <span class="muted">· RUT ${esc(g.rutFmt)}</span></h3>
          <button class="secondary" data-download="${gi}">Descargar Word</button>
        </div>
        <div class="scroll">
        <table>
          <thead><tr><th>Incluir</th><th>Categoría</th><th>Doc.</th><th>Nro.</th><th>Fecha</th><th class="n">Saldo pendiente</th><th>TIPO DE IPE</th><th>Documentos a aplicar · monto · saldo que queda</th><th class="n">Saldo final</th></tr></thead>
          <tbody>
            ${g.rows.map((r, ri) => `
              <tr>
                <td><input type="checkbox" data-f="incluir" data-r="${ri}" ${r.incluir ? 'checked' : ''} aria-label="Incluir ${esc(r.doc)} ${esc(r.nro)}"></td>
                <td>${flagHtml(r.fecha)}</td>
                <td>${esc(r.doc)}</td><td>${esc(r.nro)}</td><td class="nowrap">${esc(fmt.fechaCorta(r.fecha))}</td>
                <td class="n">${clp(r.saldoInicial)}${netoHtml(r)}</td>
                <td>${tipoHtml(r)}</td>
                <td class="apls">
                  ${r.aplicaciones.map((a, ai) => `
                    <div class="apl">
                      <input type="text" data-f="docto" data-r="${ri}" data-a="${ai}" value="${esc(a.docto)}" placeholder="Pagar FEE …" aria-label="Documento a aplicar">
                      <input type="text" data-f="monto" data-r="${ri}" data-a="${ai}" value="${esc(montoInput(a.monto))}" placeholder="Monto" aria-label="Monto a aplicar" inputmode="numeric">
                      <span class="rest" data-rest="${ri}-${ai}"></span>
                      <button class="icon" data-del="${ri}-${ai}" title="Quitar" aria-label="Quitar aplicación">×</button>
                    </div>`).join('')}
                  <button class="link" data-add="${ri}">+ Agregar documento</button>
                </td>
                <td class="n" data-final="${ri}"></td>
              </tr>`).join('')}
          </tbody>
        </table>
        </div>
        <div class="preview" data-preview="${gi}"></div>
        <ul class="warnings" data-warn="${gi}"></ul>
      </div>`;
  }

  /** Actualiza saldos, vista previa y advertencias sin redibujar los campos (no se pierde el foco). */
  function refreshMemo(gi) {
    const g = state.memoGroups[gi];
    const card = document.querySelector(`#memos [data-g="${gi}"]`);
    g.rows.forEach((r, ri) => {
      const { pasos, remanente } = memo.calcular(r);
      pasos.forEach((p, ai) => {
        const el = card.querySelector(`[data-rest="${ri}-${ai}"]`);
        el.textContent = `queda ${clp(p.saldoDespues)}`;
        el.classList.toggle('neg', p.saldoDespues < 0);
      });
      const fin = card.querySelector(`[data-final="${ri}"]`);
      fin.textContent = clp(remanente);
      fin.classList.toggle('neg', remanente < 0);
    });
    const data = memo.memoData(g, docOpts());
    const notas = data.notas.map((n) => `Nota: Diferencia de ${esc(n.doc)} ${esc(n.nro)} por ${esc(n.remanente)}.- mantener en la cuenta…`).join('<br>');
    card.querySelector(`[data-preview="${gi}"]`).innerHTML =
      `<div><b>Fecha en el memo:</b> SANTIAGO, ${esc(data.fecha)} · <b>Filas en la tabla:</b> ${data.items.length}</div>${notas ? `<div>${notas}</div>` : ''}`;
    card.querySelector(`[data-warn="${gi}"]`).innerHTML = memo.warnings(g).map((w) => `<li>${esc(w)}</li>`).join('');
  }

  function registrarMemos(entries) {
    state.historial.push(...entries);
    saveHistoriales();
    reconstruir();
    renderTodo();
  }

  async function downloadMemo(gi) {
    const g = state.memoGroups[gi];
    if (!g.rows.some((r) => r.incluir)) return alert('Selecciona al menos un documento.');
    try {
      const opts = docOpts();
      const data = memo.memoData(g, opts);
      saveAs(renderMemo(data).generate({ type: 'blob', mimeType: DOCX_MIME }), memo.fileName(g, opts.fecha));
      registrarMemos([memo.historialEntry(g, data, new Date(), opts)]);
    } catch (err) { alertError(err); }
  }

  async function downloadMemos() {
    try {
      const out = new PizZip();
      const entries = [];
      const ahora = new Date();
      const opts = docOpts();
      for (const g of state.memoGroups) {
        if (!g.rows.some((r) => r.incluir)) continue;
        const data = memo.memoData(g, opts);
        out.file(memo.fileName(g, opts.fecha), renderMemo(data).generate({ type: 'uint8array' }));
        entries.push(memo.historialEntry(g, data, ahora, opts));
      }
      if (!entries.length) return;
      saveAs(out.generate({ type: 'blob' }), `Memos_IPE_${$('cfg-fecha').value}.zip`);
      registrarMemos(entries);
    } catch (err) { alertError(err); }
  }

  // ---------- devoluciones ----------
  function renderDevoluciones() {
    const groups = state.devGroups;
    $('n-dev').textContent = groups.length;
    $('dev-resumen').textContent = `${groups.length} empresa(s) con IPE pendientes de devolución.`;
    $('btn-zip-dev').disabled = !groups.length;
    $('devoluciones').innerHTML = groups.map((g, gi) => devCardHtml(g, gi)).join('') || '<p class="muted">No hay devoluciones pendientes.</p>';
    groups.forEach((g, gi) => refreshDev(gi));
    filtrarTarjetas('devoluciones', 'devGroups', $('buscar-dev').value);
  }

  function devCardHtml(g, gi) {
    const b = g.banco;
    return `
      <div class="card" data-g="${gi}">
        <div class="memo-head">
          <h3>${esc(g.razon)} <span class="muted">· RUT ${esc(g.rutFmt)}</span></h3>
          <button class="secondary" data-download="${gi}">Descargar Word</button>
        </div>
        <div class="scroll">
        <table>
          <thead><tr><th>Incluir</th><th>Categoría</th><th>C.F.</th><th>Tipo docto.</th><th>Fecha</th><th class="n">Saldo pendiente</th><th>Monto a devolver</th><th>TIPO DE IPE</th></tr></thead>
          <tbody>
            ${g.rows.map((r, ri) => `
              <tr>
                <td><input type="checkbox" data-f="incluir" data-r="${ri}" ${r.incluir ? 'checked' : ''} aria-label="Incluir ${esc(r.doc)} ${esc(r.nro)}"></td>
                <td>${flagHtml(r.fecha)}</td>
                <td>${esc(r.cf)}</td><td>${esc(r.doc)} ${esc(r.nro)}</td><td class="nowrap">${esc(fmt.fechaCorta(r.fecha))}</td>
                <td class="n">${clp(r.saldoInicial)}${netoHtml(r)}</td>
                <td><input type="text" class="monto" data-f="montoDevolver" data-r="${ri}" value="${esc(montoInput(r.montoDevolver))}" inputmode="numeric" aria-label="Monto a devolver"><div class="prev" data-rest="${ri}"></div></td>
                <td>${tipoHtml(r)}</td>
              </tr>`).join('')}
          </tbody>
        </table>
        </div>
        <div class="form banco">
          <label class="ancho">REF.<input type="text" data-b="referencia" value="${esc(b.referencia)}" placeholder="Correo electrónico … de fecha dd.mm.aaaa a las hh:mm hrs."></label>
          <label>Tipo de cuenta<select data-b="tipoCuenta">${devoluciones.TIPOS_CUENTA.map((t) => `<option ${t === b.tipoCuenta ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
          <label>Número de cuenta<input type="text" data-b="numeroCuenta" value="${esc(b.numeroCuenta)}"></label>
          <label>Banco<select data-b="banco"><option value="">Selecciona el banco</option>${devoluciones.BANCOS.map((n) => `<option ${n === b.banco ? 'selected' : ''}>${esc(n)}</option>`).join('')}<option ${b.banco === 'Otro' ? 'selected' : ''}>Otro</option></select>
            <input type="text" data-b="bancoOtro" value="${esc(b.bancoOtro)}" placeholder="Nombre del banco" aria-label="Nombre del banco" ${b.banco === 'Otro' ? '' : 'hidden'}></label>
          <label class="ancho">Correos electrónicos (uno por línea)<textarea data-b="correos" rows="2">${esc(b.correos)}</textarea></label>
        </div>
        <div class="preview" data-preview="${gi}"></div>
        <ul class="warnings" data-warn="${gi}"></ul>
      </div>`;
  }

  function refreshDev(gi) {
    const g = state.devGroups[gi];
    const card = document.querySelector(`#devoluciones [data-g="${gi}"]`);
    g.rows.forEach((r, ri) => {
      const queda = r.saldoInicial - (Number(r.montoDevolver) || 0);
      const el = card.querySelector(`[data-rest="${ri}"]`);
      el.textContent = queda ? `queda pendiente ${clp(queda)}` : 'devolución total';
      el.classList.toggle('neg', queda < 0);
    });
    const data = devoluciones.devolucionData(g, docOpts());
    card.querySelector(`[data-preview="${gi}"]`).innerHTML =
      `<div><b>Fecha en el oficio:</b> SANTIAGO, ${esc(data.fecha)} · <b>Monto de devolución:</b> ${esc(data.monto_total)}</div>`;
    card.querySelector(`[data-warn="${gi}"]`).innerHTML = devoluciones.warnings(g).map((w) => `<li>${esc(w)}</li>`).join('');
  }

  function registrarDevoluciones(entries) {
    state.historialDev.push(...entries);
    saveHistoriales();
    reconstruir();
    renderTodo();
  }

  async function downloadDev(gi) {
    const g = state.devGroups[gi];
    if (!g.rows.some((r) => r.incluir)) return alert('Selecciona al menos un documento.');
    try {
      const opts = docOpts();
      const data = devoluciones.devolucionData(g, opts);
      saveAs(renderDevolucion(data).generate({ type: 'blob', mimeType: DOCX_MIME }), devoluciones.fileName(g, opts.fecha));
      registrarDevoluciones([devoluciones.historialEntry(g, data, new Date(), opts)]);
    } catch (err) { alertError(err); }
  }

  async function downloadDevs() {
    try {
      const out = new PizZip();
      const entries = [];
      const ahora = new Date();
      const opts = docOpts();
      for (const g of state.devGroups) {
        if (!g.rows.some((r) => r.incluir)) continue;
        const data = devoluciones.devolucionData(g, opts);
        out.file(devoluciones.fileName(g, opts.fecha), renderDevolucion(data).generate({ type: 'uint8array' }));
        entries.push(devoluciones.historialEntry(g, data, ahora, opts));
      }
      if (!entries.length) return;
      saveAs(out.generate({ type: 'blob' }), `Devoluciones_IPE_${$('cfg-fecha').value}.zip`);
      registrarDevoluciones(entries);
    } catch (err) { alertError(err); }
  }

  // ---------- historiales ----------
  function tablaHistorial(lista, columnaMonto, montoDoc) {
    return `<table><thead><tr><th>Generado</th><th>Fecha</th><th>Empresa</th><th>RUT</th><th>Documentos (${columnaMonto} / queda pendiente)</th><th></th></tr></thead><tbody>
      ${lista.map((e) => `<tr>
        <td>${esc(new Date(e.generado).toLocaleString('es-CL'))}</td>
        <td class="nowrap">${esc(e.fechaCompleta || e.fechaMemo || e.fechaDocumento)}</td>
        <td>${esc(e.empresa)}</td><td class="nowrap">${esc(e.rut)}</td>
        <td>${e.documentos.map((d) => `<b>${esc(d.doc)} ${esc(d.nro)}</b>: ${clp(montoDoc(d))} / ${clp(d.remanente)}`).join('<br>')}</td>
        <td><div class="actions">
          <button class="link" data-redl="${esc(e.id)}">Descargar de nuevo</button>
          <button class="link" data-undo="${esc(e.id)}">Volver a pendientes</button>
        </div></td></tr>`).join('')}
      </tbody></table>`;
  }

  const ordenar = (lista) => [...lista].sort((a, b) => b.generado.localeCompare(a.generado));

  function renderHistorial() {
    const h = ordenar(state.historial);
    $('n-hist').textContent = h.length;
    $('hist-resumen').textContent = `${h.length} memo(s) generados.`;
    $('historial').innerHTML = h.length ? tablaHistorial(h, 'aplicado', (d) => d.aplicado) : '<p class="muted">Aún no se han generado memos.</p>';
  }

  function renderHistorialDev() {
    const h = ordenar(state.historialDev);
    $('n-hist-dev').textContent = h.length;
    $('hist-dev-resumen').textContent = `${h.length} devolución(es) generadas.`;
    $('historial-dev').innerHTML = h.length ? tablaHistorial(h, 'devuelto', (d) => d.devuelto) : '<p class="muted">Aún no se han generado devoluciones.</p>';
  }

  function eventosHistorial(contenedor, lista, nombre) {
    $(contenedor).addEventListener('click', (ev) => {
      const t = ev.target.closest('button');
      if (!t) return;
      const e = state[lista].find((x) => x.id === (t.dataset.redl || t.dataset.undo));
      if (!e) return;
      if (t.dataset.redl) {
        try {
          const zip = e.tipo === 'devolucion' ? renderDevolucion(e.data) : renderMemo(e.data);
          saveAs(zip.generate({ type: 'blob', mimeType: DOCX_MIME }), e.archivo);
        } catch (err) { alertError(err); }
      } else if (confirm(`¿Quitar del historial ${nombre} de ${e.empresa}? Sus montos volverán a quedar pendientes.`)) {
        state[lista] = state[lista].filter((x) => x.id !== e.id);
        saveHistoriales();
        reconstruir();
        renderTodo();
      }
    });
  }

  function exportarHistorialXlsx(lista, hoja, archivo, cabecera, fila) {
    const aoa = [cabecera];
    for (const e of lista) for (const d of e.documentos) aoa.push(fila(e, d));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = cabecera.map((c) => ({ wch: Math.max(10, c.length + 4) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, hoja);
    XLSX.writeFile(wb, `${archivo}_${hoyISO()}.xlsx`);
  }

  function backup() {
    const blob = new Blob([JSON.stringify({ tipo: 'ipe-historial', version: 2, historial: state.historial, historialDevoluciones: state.historialDev }, null, 1)], { type: 'application/json' });
    saveAs(blob, `respaldo_historial_IPE_${hoyISO()}.json`);
  }

  async function importar(file) {
    if (!file) return;
    try {
      const json = JSON.parse(await file.text());
      const memos = Array.isArray(json) ? json : json.historial || [];
      const devs = Array.isArray(json) ? [] : json.historialDevoluciones || [];
      const valido = (l) => Array.isArray(l) && l.every((e) => e.id && Array.isArray(e.documentos));
      if (!valido(memos) || !valido(devs)) throw new Error('El archivo no es un respaldo válido del historial.');
      const sumar = (destino, nuevos) => {
        const ids = new Set(state[destino].map((e) => e.id));
        const add = nuevos.filter((e) => !ids.has(e.id));
        state[destino].push(...add);
        return add.length;
      };
      const nm = sumar('historial', memos);
      const nd = sumar('historialDev', devs);
      saveHistoriales();
      reconstruir();
      renderTodo();
      alert(`Se importaron ${nm} memo(s) y ${nd} devolución(es).`);
    } catch (err) {
      alert(`No se pudo importar: ${err.message}`);
    }
  }

  // ---------- eventos ----------
  /** Edición de campos en tarjetas de memos o devoluciones. */
  function eventosTarjetas(contenedor, grupos, refresh, rerender) {
    $(contenedor).addEventListener('input', (e) => {
      const card = e.target.closest('[data-g]');
      if (!card) return;
      const gi = Number(card.dataset.g);
      const g = state[grupos][gi];
      if (e.target.dataset.b) {
        g.banco[e.target.dataset.b] = e.target.value;
        if (e.target.dataset.b === 'banco') {
          const otro = card.querySelector('[data-b="bancoOtro"]');
          otro.hidden = e.target.value !== 'Otro';
          if (!otro.hidden) otro.focus();
        }
        return refresh(gi);
      }
      const f = e.target.dataset.f;
      if (!f) return;
      const ri = Number(e.target.dataset.r);
      const row = g.rows[ri];
      if (f === 'incluir') row.incluir = e.target.checked;
      else if (f === 'montoDevolver') row.montoDevolver = e.target.value.trim() === '' ? '' : fmt.toNumber(e.target.value);
      else {
        const apl = row.aplicaciones[Number(e.target.dataset.a)];
        if (f === 'docto') apl.docto = e.target.value;
        else apl.monto = e.target.value.trim() === '' ? '' : fmt.toNumber(e.target.value);
      }
      refresh(gi);
    });
    $(contenedor).addEventListener('click', (e) => {
      const card = e.target.closest('[data-g]');
      const t = e.target.closest('button');
      if (!card || !t) return;
      const gi = Number(card.dataset.g);
      const g = state[grupos][gi];
      if (t.dataset.download !== undefined) rerender.download(gi);
      else if (t.dataset.add !== undefined) {
        g.rows[Number(t.dataset.add)].aplicaciones.push({ docto: '', monto: '' });
        rerender.card(gi);
      } else if (t.dataset.del !== undefined) {
        const [ri, ai] = t.dataset.del.split('-').map(Number);
        g.rows[ri].aplicaciones.splice(ai, 1);
        rerender.card(gi);
      }
    });
  }

  function init() {
    const local = location.protocol === 'file:';
    $('modo-local').hidden = !local;
    $('modo-web').hidden = local;
    loadConfig();
    loadHistoriales();
    $('archivo').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const drop = $('carga');
    drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); handleFile(e.dataTransfer.files[0]); });

    document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => {
      document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-selected', String(x === b)));
      document.querySelectorAll('.tab').forEach((t) => { t.hidden = t.id !== `tab-${b.dataset.tab}`; });
    }));

    $('sel-hoja').addEventListener('change', (e) => {
      try { cargarHoja(e.target.value); } catch (err) { mostrarError(err); }
    });
    $('archivo-anterior').addEventListener('change', (e) => { copiarApuntes(e.target.files[0]); e.target.value = ''; });
    $('btn-hoja-trabajo').addEventListener('click', () => {
      try { descargarHojaTrabajo(); } catch (err) { mostrarError(err); }
    });
    $('cfg-iniciales').addEventListener('change', saveConfig);
    $('cfg-fecha').addEventListener('change', () => {
      state.memoGroups.forEach((g, gi) => refreshMemo(gi));
      state.devGroups.forEach((g, gi) => refreshDev(gi));
    });
    $('filtro-bandera').addEventListener('change', renderListado);
    $('buscar-memos').addEventListener('input', (e) => filtrarTarjetas('memos', 'memoGroups', e.target.value, e.target));
    $('buscar-dev').addEventListener('input', (e) => filtrarTarjetas('devoluciones', 'devGroups', e.target.value, e.target));
    $('buscar').addEventListener('input', renderListado);

    eventosTarjetas('memos', 'memoGroups', refreshMemo, {
      download: downloadMemo,
      card: (gi) => {
        document.querySelector(`#memos [data-g="${gi}"]`).outerHTML = memoCardHtml(state.memoGroups[gi], gi);
        refreshMemo(gi);
      },
    });
    eventosTarjetas('devoluciones', 'devGroups', refreshDev, { download: downloadDev, card: () => {} });
    $('btn-zip').addEventListener('click', downloadMemos);
    $('btn-zip-dev').addEventListener('click', downloadDevs);

    eventosHistorial('historial', 'historial', 'el memo');
    eventosHistorial('historial-dev', 'historialDev', 'la devolución');
    $('btn-hist-xlsx').addEventListener('click', () => exportarHistorialXlsx(state.historial, 'Historial memos', 'Historial_memos_IPE',
      ['Generado', 'Fecha', 'Empresa', 'RUT', 'Doc.', 'Nro.', 'Saldo pendiente', 'Aplicaciones', 'Total aplicado', 'Queda pendiente'],
      (e, d) => [new Date(e.generado).toLocaleString('es-CL'), e.fechaCompleta || e.fechaMemo, e.empresa, e.rut, d.doc, d.nro, d.saldoInicial ?? d.saldo, d.aplicaciones, d.aplicado, d.remanente]));
    $('btn-hist-dev-xlsx').addEventListener('click', () => exportarHistorialXlsx(state.historialDev, 'Historial devoluciones', 'Historial_devoluciones_IPE',
      ['Generado', 'Fecha', 'Empresa', 'RUT', 'Doc.', 'Nro.', 'Saldo pendiente', 'Devuelto', 'Queda pendiente', 'Banco', 'Tipo de cuenta', 'N° cuenta'],
      (e, d) => [new Date(e.generado).toLocaleString('es-CL'), e.fechaCompleta, e.empresa, e.rut, d.doc, d.nro, d.saldoInicial, d.devuelto, d.remanente, (e.banco?.banco === 'Otro' ? e.banco?.bancoOtro : e.banco?.banco) || '', e.banco?.tipoCuenta || '', e.banco?.numeroCuenta || '']));
    document.querySelectorAll('[data-backup]').forEach((b) => b.addEventListener('click', backup));
    document.querySelectorAll('[data-import]').forEach((i) => i.addEventListener('change', (e) => { importar(e.target.files[0]); e.target.value = ''; }));
  }

  init();
})();
