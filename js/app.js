// Interfaz: carga del Excel, pestañas, descargas e historial de memos.
(function () {
  const { format: fmt, parser, memo, devoluciones, resumen } = window.IPE;
  const $ = (id) => document.getElementById(id);
  const CONFIG_KEY = 'ipe-config';
  const HIST_KEY = 'ipe-historial';

  const state = { rows: [], memoGroups: [], devGroups: [], tieneColumnasAplicar: false, template: null, historial: [], alertas: [] };

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clp = (n) => fmt.formatCLP(n);
  const hoyISO = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

  // ---------- configuración e historial (se guardan en este navegador) ----------
  function readStore(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch (e) { return fallback; }
  }
  function writeStore(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
  }

  function loadConfig() {
    const cfg = readStore(CONFIG_KEY, {});
    $('cfg-iniciales').value = cfg.iniciales ?? 'ERS/yaa';
    const hoy = new Date();
    $('cfg-fecha').value = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}-${String(hoy.getDate()).padStart(2, '0')}`;
  }
  function saveConfig() {
    writeStore(CONFIG_KEY, { iniciales: $('cfg-iniciales').value });
  }
  function memoOpts() {
    const [y, m, d] = $('cfg-fecha').value.split('-').map(Number);
    return { fecha: y ? new Date(y, m - 1, d) : new Date(), iniciales: $('cfg-iniciales').value.trim() };
  }

  function loadHistorial() {
    const h = readStore(HIST_KEY, []);
    state.historial = Array.isArray(h) ? h : [];
  }
  function saveHistorial() {
    if (!writeStore(HIST_KEY, state.historial))
      alert('No se pudo guardar el historial en este navegador. Descarga un respaldo para no perderlo.');
  }
  // ---------- carga ----------
  async function handleFile(file) {
    if (!file) return;
    $('estado-carga').textContent = 'Leyendo…';
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
      const parsed = parser.parseAoa(aoa);
      state.rows = parsed.rows;
      state.tieneColumnasAplicar = parsed.tieneColumnasAplicar;
      state.devGroups = devoluciones.buildGroups(state.rows);
      rebuildMemoGroups();
      $('estado-carga').innerHTML = `<span class="ok">✔</span> ${esc(file.name)} — hoja "${esc(wb.SheetNames[0])}", ${state.rows.length} registros.`;
      $('contenido').hidden = false;
      renderResumen();
      renderMemos();
      renderDevoluciones();
      renderHistorial();
    } catch (err) {
      console.error(err);
      $('estado-carga').innerHTML = `<span style="color:var(--danger)">${esc(err.message)}</span>`;
    }
  }

  function rebuildMemoGroups() {
    const idx = memo.indiceHistorial(state.historial);
    state.memoGroups = memo.buildGroups(state.rows, idx);
    state.ocultos = state.rows.filter((r) => r.categoria === 'APLICAR' && memo.ocultoPorHistorial(r, idx.get(memo.rowKey(r)))).length;
  }

  // ---------- resumen ----------
  function renderResumen() {
    const r = resumen.compute(state.rows, new Date(), state.tieneColumnasAplicar);
    const cat = (c) => r.porEstado.find((e) => e.categoria === c) || { cantidad: 0, monto: 0 };
    const kpis = [
      ['Saldo neto IPE', clp(r.total)],
      ['Registros', r.cantidad],
      ['Empresas', r.empresas],
      ['A aplicar', `${cat('APLICAR').cantidad} · ${clp(cat('APLICAR').monto)}`],
      ['A devolver', `${cat('DEVOLVER').cantidad} · ${clp(cat('DEVOLVER').monto)}`],
    ];
    $('kpis').innerHTML = kpis.map(([l, v]) => `<div class="kpi"><div class="l">${esc(l)}</div><div class="v">${esc(v)}</div></div>`).join('');
    $('por-estado').innerHTML = table(['Estado', 'Cant.', 'Saldo neto'], r.porEstado.map((e) => [esc(e.categoria), e.cantidad, clp(e.monto)]), [false, true, true]);
    $('antiguedad').innerHTML = table(['Desde la fecha del IPE', 'Cant.', 'Saldo neto'], r.antiguedad.map((a) => [esc(a.label), a.cantidad, clp(a.monto)]), [false, true, true]);

    state.alertas = r.alertas;
    const tipos = [...new Set(r.alertas.map((a) => a.tipo))];
    $('filtro-alertas').innerHTML = ['<option value="">Todas</option>', ...tipos.map((t) => `<option>${esc(t)}</option>`)].join('');
    renderAlertas();
  }

  function renderAlertas() {
    const f = $('filtro-alertas').value;
    const list = state.alertas.filter((a) => !f || a.tipo === f);
    $('n-alertas').textContent = list.length;
    $('alertas').innerHTML = list.length
      ? table(['Fila', 'Tipo', 'Empresa', 'Documento', 'Detalle'], list.map((a) => [a.fila, esc(a.tipo), esc(a.razon), esc(a.doc), esc(a.mensaje)]), [true])
      : '<p class="muted">Sin alertas.</p>';
  }

  function table(headers, rows, numeric = []) {
    const th = headers.map((h, i) => `<th class="${numeric[i] ? 'n' : ''}">${esc(h)}</th>`).join('');
    const tr = rows.map((r) => `<tr>${r.map((c, i) => `<td class="${numeric[i] ? 'n' : ''}">${c}</td>`).join('')}</tr>`).join('');
    return `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table>`;
  }

  // ---------- memos ----------
  function renderMemos() {
    const groups = state.memoGroups;
    $('n-memos').textContent = groups.length;
    $('aviso-columnas').hidden = state.tieneColumnasAplicar;
    $('memos-resumen').textContent = `${groups.length} empresa(s) con IPE pendientes de memo.`;
    $('ocultos-hist').textContent = state.ocultos ? `${state.ocultos} documento(s) APLICAR no se muestran: ya tienen memo en el Historial y su saldo no ha cambiado (Contabilidad aún no lo procesa).` : '';
    $('btn-zip').disabled = !groups.length;
    $('memos').innerHTML = groups.map((g, gi) => cardHtml(g, gi)).join('') || '<p class="muted">No hay memos pendientes.</p>';
    groups.forEach((g, gi) => refreshGroup(gi));
  }

  function montoInput(v) {
    return v === '' || v === null || v === undefined ? '' : fmt.miles(Number(v));
  }

  function cardHtml(g, gi) {
    return `
      <div class="card" data-g="${gi}">
        <div class="memo-head">
          <h3>${esc(g.razon)} <span class="muted">· RUT ${esc(g.rutFmt)}</span></h3>
          <button class="secondary" data-download="${gi}">Descargar Word</button>
        </div>
        <div class="scroll">
        <table>
          <thead><tr><th>Incluir</th><th>Doc.</th><th>Nro.</th><th>Fecha</th><th class="n">Saldo IPE</th><th>XX</th><th>Justificación (MOTIVO)</th><th>Documentos a aplicar · monto · saldo que queda</th><th class="n">Saldo final</th></tr></thead>
          <tbody>
            ${g.rows.map((r, ri) => `
              <tr class="${r.naturaleza === 'DEUDOR' ? 'deudor' : ''}">
                <td><input type="checkbox" data-f="incluir" data-r="${ri}" ${r.incluir ? 'checked' : ''} aria-label="Incluir ${esc(r.doc)} ${esc(r.nro)}"></td>
                <td>${esc(r.doc)}</td><td>${esc(r.nro)}${r.previo ? `<div class="prev">Memo previo ${esc(r.previo.fechaMemo)}: saldo entonces ${clp(r.previo.saldo)}, aplicado ${clp(r.previo.aplicado)}. Saldo actual ${clp(r.saldo)}.</div>` : ''}</td><td>${esc(fmt.fechaCorta(r.fecha))}</td>
                <td class="n">${clp(r.saldo)}</td><td>${esc(r.naturaleza)}</td>
                <td><input type="text" data-f="justificacion" data-r="${ri}" value="${esc(r.justificacion)}"></td>
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

  function rerenderCard(gi) {
    const card = document.querySelector(`[data-g="${gi}"]`);
    card.outerHTML = cardHtml(state.memoGroups[gi], gi);
    refreshGroup(gi);
  }

  /** Actualiza saldos, vista previa y advertencias sin redibujar los campos (no se pierde el foco). */
  function refreshGroup(gi) {
    const g = state.memoGroups[gi];
    const card = document.querySelector(`[data-g="${gi}"]`);
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
    const data = memo.memoData(g, memoOpts());
    const motivo = data.motivos.map((m) => `<b>${esc(m.doc)} ${esc(m.nro)}</b> ${esc(m.justificacion)}`).join('<br>');
    const notas = data.notas.map((n) => `Nota: Diferencia de ${esc(n.doc)} ${esc(n.nro)} por ${esc(n.remanente)}.- mantener en la cuenta…`).join('<br>');
    card.querySelector(`[data-preview="${gi}"]`).innerHTML =
      `<div><b>MOTIVO:</b><br>${motivo || '—'}</div><div>Filas en la tabla del memo: ${data.items.length}</div>${notas ? `<div>${notas}</div>` : ''}`;
    card.querySelector(`[data-warn="${gi}"]`).innerHTML = memo.warnings(g).map((w) => `<li>${esc(w)}</li>`).join('');
  }

  async function getTemplate() {
    if (!state.template) {
      const res = await fetch('templates/memo_template.docx', { cache: 'no-cache' });
      if (!res.ok) throw new Error('No se pudo cargar la plantilla del memo.');
      state.template = await res.arrayBuffer();
    }
    return state.template;
  }

  async function renderDocx(data) {
    const tpl = await getTemplate();
    return memo.render(tpl, data, PizZip, window.docxtemplater);
  }

  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  function registrar(entries) {
    state.historial.push(...entries);
    saveHistorial();
    rebuildMemoGroups();
    renderMemos();
    renderHistorial();
  }

  async function downloadOne(gi) {
    const g = state.memoGroups[gi];
    if (!g.rows.some((r) => r.incluir)) return alert('Selecciona al menos un documento.');
    try {
      const data = memo.memoData(g, memoOpts());
      const zip = await renderDocx(data);
      saveAs(zip.generate({ type: 'blob', mimeType: DOCX_MIME }), memo.fileName(g));
      registrar([memo.historialEntry(g, data, new Date())]);
    } catch (err) { alertError(err); }
  }

  async function downloadAll() {
    try {
      const out = new PizZip();
      const entries = [];
      const ahora = new Date();
      for (const g of state.memoGroups) {
        if (!g.rows.some((r) => r.incluir)) continue;
        const data = memo.memoData(g, memoOpts());
        const zip = await renderDocx(data);
        out.file(memo.fileName(g), zip.generate({ type: 'uint8array' }));
        entries.push(memo.historialEntry(g, data, ahora));
      }
      if (!entries.length) return;
      saveAs(out.generate({ type: 'blob' }), `Memos_IPE_${$('cfg-fecha').value}.zip`);
      registrar(entries);
    } catch (err) { alertError(err); }
  }

  function alertError(err) {
    console.error(err);
    const detalle = err.properties && err.properties.errors ? err.properties.errors.map((e) => e.properties.explanation).join('\n') : '';
    alert(`No se pudo generar el memo: ${err.message}\n${detalle}`);
  }

  // ---------- historial ----------
  function renderHistorial() {
    const h = [...state.historial].sort((a, b) => b.generado.localeCompare(a.generado));
    $('n-hist').textContent = h.length;
    $('hist-resumen').textContent = `${h.length} memo(s) generados · ${h.reduce((s, e) => s + e.documentos.length, 0)} documento(s).`;
    $('historial').innerHTML = h.length
      ? `<table><thead><tr><th>Generado</th><th>Fecha memo</th><th>Empresa</th><th>RUT</th><th>Documentos (aplicado / queda)</th><th></th></tr></thead><tbody>
        ${h.map((e) => `<tr>
          <td>${esc(new Date(e.generado).toLocaleString('es-CL'))}</td>
          <td>${esc(e.fechaMemo)}</td><td>${esc(e.empresa)}</td><td>${esc(e.rut)}</td>
          <td>${e.documentos.map((d) => `<b>${esc(d.doc)} ${esc(d.nro)}</b>: ${clp(d.aplicado)} / ${clp(d.remanente)}`).join('<br>')}</td>
          <td><div class="actions">
            <button class="link" data-redl="${esc(e.id)}">Descargar de nuevo</button>
            <button class="link" data-undo="${esc(e.id)}">Volver a pendientes</button>
          </div></td></tr>`).join('')}
        </tbody></table>`
      : '<p class="muted">Aún no se han generado memos.</p>';
  }

  async function redownload(id) {
    const e = state.historial.find((x) => x.id === id);
    if (!e) return;
    try {
      const zip = await renderDocx(e.data);
      saveAs(zip.generate({ type: 'blob', mimeType: DOCX_MIME }), e.archivo);
    } catch (err) { alertError(err); }
  }

  function undo(id) {
    const e = state.historial.find((x) => x.id === id);
    if (!e || !confirm(`¿Quitar del historial el memo de ${e.empresa}? Sus documentos volverán a "Memos" como pendientes.`)) return;
    state.historial = state.historial.filter((x) => x.id !== id);
    saveHistorial();
    rebuildMemoGroups();
    renderMemos();
    renderHistorial();
  }

  function exportHistorialXlsx() {
    const aoa = [['Generado', 'Fecha memo', 'Empresa', 'RUT', 'Doc.', 'Nro.', 'Saldo IPE', 'Aplicaciones', 'Total aplicado', 'Saldo que queda', 'Archivo']];
    for (const e of state.historial) {
      for (const d of e.documentos) {
        aoa.push([new Date(e.generado).toLocaleString('es-CL'), e.fechaMemo, e.empresa, e.rut, d.doc, d.nro, d.saldo, d.aplicaciones, d.aplicado, d.remanente, e.archivo]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [18, 12, 40, 14, 6, 12, 14, 50, 14, 14, 45].map((wch) => ({ wch }));
    for (let r = 1; r < aoa.length; r++) for (const c of [6, 8, 9]) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (cell) cell.z = '"$" #,##0'; }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Historial memos');
    XLSX.writeFile(wb, `Historial_memos_IPE_${hoyISO()}.xlsx`);
  }

  function backupHistorial() {
    const blob = new Blob([JSON.stringify({ tipo: 'ipe-historial', version: 1, historial: state.historial }, null, 1)], { type: 'application/json' });
    saveAs(blob, `respaldo_historial_IPE_${hoyISO()}.json`);
  }

  async function importHistorial(file) {
    if (!file) return;
    try {
      const json = JSON.parse(await file.text());
      const lista = Array.isArray(json) ? json : json.historial;
      if (!Array.isArray(lista) || lista.some((e) => !e.id || !Array.isArray(e.documentos))) throw new Error('El archivo no es un respaldo válido del historial.');
      const ids = new Set(state.historial.map((e) => e.id));
      const nuevos = lista.filter((e) => !ids.has(e.id));
      state.historial.push(...nuevos);
      saveHistorial();
      rebuildMemoGroups();
      renderMemos();
      renderHistorial();
      alert(`Se importaron ${nuevos.length} memo(s) al historial.`);
    } catch (err) {
      alert(`No se pudo importar: ${err.message}`);
    }
  }

  // ---------- devoluciones ----------
  function renderDevoluciones() {
    const groups = state.devGroups;
    const total = groups.reduce((s, g) => s + g.total, 0);
    $('n-dev').textContent = groups.length;
    $('dev-resumen').textContent = `${groups.length} empresa(s), ${groups.reduce((s, g) => s + g.rows.length, 0)} IPE en estado DEVOLVER · total neto ${clp(total)}`;
    $('btn-dev').disabled = !groups.length;
    $('devoluciones').innerHTML = groups.length
      ? table(['RUT', 'Razón social', 'Documentos', 'Total a devolver', 'Última gestión'],
        groups.map((g) => [esc(g.rutFmt), esc(g.razon), esc(g.rows.map((r) => `${r.doc} ${r.nro}`).join(', ')), clp(g.total),
          esc(parser.ultimaGestion(g.rows[g.rows.length - 1].gestion))]), [false, false, false, true])
      : '<p class="muted">No hay registros con estado DEVOLVER.</p>';
  }

  function downloadDevoluciones() {
    const { resumen: res, detalle } = devoluciones.sheets(state.devGroups);
    const wb = XLSX.utils.book_new();
    const add = (aoa, name, widths, moneyCols) => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = widths.map((wch) => ({ wch }));
      for (let r = 1; r < aoa.length; r++) {
        for (const c of moneyCols) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell && cell.t === 'n') cell.z = '"$" #,##0';
        }
      }
      XLSX.utils.book_append_sheet(wb, ws, name);
    };
    add(res, 'Resumen', [14, 40, 9, 16, 18, 16, 18, 28, 30], [3]);
    add(detalle, 'Detalle', [14, 40, 6, 7, 12, 10, 11, 12, 10, 12, 35, 60, 8], [7, 9]);
    XLSX.writeFile(wb, `Devoluciones_IPE_${hoyISO()}.xlsx`);
  }

  // ---------- eventos ----------
  function init() {
    loadConfig();
    loadHistorial();
    $('archivo').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const drop = $('carga');
    drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); handleFile(e.dataTransfer.files[0]); });

    document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => {
      document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-selected', String(x === b)));
      document.querySelectorAll('.tab').forEach((t) => { t.hidden = t.id !== `tab-${b.dataset.tab}`; });
    }));

    $('filtro-alertas').addEventListener('change', renderAlertas);
    $('cfg-iniciales').addEventListener('change', saveConfig);
    $('cfg-fecha').addEventListener('change', () => state.memoGroups.forEach((g, gi) => refreshGroup(gi)));

    $('memos').addEventListener('input', (e) => {
      const card = e.target.closest('[data-g]');
      const f = e.target.dataset.f;
      if (!card || !f) return;
      const gi = Number(card.dataset.g);
      const row = state.memoGroups[gi].rows[Number(e.target.dataset.r)];
      if (f === 'incluir') row.incluir = e.target.checked;
      else if (f === 'justificacion') row.justificacion = e.target.value;
      else {
        const apl = row.aplicaciones[Number(e.target.dataset.a)];
        if (f === 'docto') apl.docto = e.target.value;
        else apl.monto = e.target.value.trim() === '' ? '' : fmt.toNumber(e.target.value);
      }
      refreshGroup(gi);
    });
    $('memos').addEventListener('click', (e) => {
      const card = e.target.closest('[data-g]');
      if (!card) return;
      const gi = Number(card.dataset.g);
      const g = state.memoGroups[gi];
      const t = e.target.closest('button');
      if (!t) return;
      if (t.dataset.download !== undefined) downloadOne(gi);
      else if (t.dataset.add !== undefined) {
        g.rows[Number(t.dataset.add)].aplicaciones.push({ docto: '', monto: '' });
        rerenderCard(gi);
      } else if (t.dataset.del !== undefined) {
        const [ri, ai] = t.dataset.del.split('-').map(Number);
        g.rows[ri].aplicaciones.splice(ai, 1);
        rerenderCard(gi);
      }
    });
    $('btn-zip').addEventListener('click', downloadAll);
    $('btn-dev').addEventListener('click', downloadDevoluciones);

    $('historial').addEventListener('click', (e) => {
      const t = e.target.closest('button');
      if (!t) return;
      if (t.dataset.redl) redownload(t.dataset.redl);
      if (t.dataset.undo) undo(t.dataset.undo);
    });
    $('btn-hist-xlsx').addEventListener('click', exportHistorialXlsx);
    $('btn-hist-backup').addEventListener('click', backupHistorial);
    $('hist-import').addEventListener('change', (e) => { importHistorial(e.target.files[0]); e.target.value = ''; });
  }

  init();
})();
