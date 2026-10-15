// Interfaz: carga del Excel, pestañas y descargas.
(function () {
  const { format: fmt, parser, memo, devoluciones, resumen } = window.IPE;
  const $ = (id) => document.getElementById(id);
  const CONFIG_KEY = 'ipe-config';
  const DEFAULT_CARGO = 'ENCARGADA SECCIÓN GESTIÓN DE INGRESOS';

  const state = { rows: [], memoGroups: [], devGroups: [], tieneColumnasAplicar: false, template: null, nombreArchivo: '' };

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clp = (n) => fmt.formatCLP(n);

  // ---------- configuración (se recuerda en el navegador) ----------
  function loadConfig() {
    let cfg = {};
    try { cfg = JSON.parse(localStorage.getItem(CONFIG_KEY)) || {}; } catch (e) { /* almacenamiento no disponible */ }
    $('cfg-iniciales').value = cfg.iniciales ?? 'ERS/yaa';
    $('cfg-firmante').value = cfg.firmante ?? '';
    $('cfg-cargo').value = cfg.cargo ?? DEFAULT_CARGO;
    const hoy = new Date();
    $('cfg-fecha').value = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}-${String(hoy.getDate()).padStart(2, '0')}`;
  }
  function saveConfig() {
    const cfg = { iniciales: $('cfg-iniciales').value, firmante: $('cfg-firmante').value, cargo: $('cfg-cargo').value };
    try { localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg)); } catch (e) { /* sin persistencia */ }
  }
  function memoOpts() {
    const [y, m, d] = $('cfg-fecha').value.split('-').map(Number);
    return {
      fecha: y ? new Date(y, m - 1, d) : new Date(),
      iniciales: $('cfg-iniciales').value.trim(),
      firmante: $('cfg-firmante').value.trim().toUpperCase(),
      cargo: $('cfg-cargo').value.trim().toUpperCase(),
    };
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
      state.memoGroups = memo.buildGroups(state.rows);
      state.devGroups = devoluciones.buildGroups(state.rows);
      state.nombreArchivo = file.name;
      $('estado-carga').innerHTML = `<span class="ok">✔</span> ${esc(file.name)} — hoja "${esc(wb.SheetNames[0])}", ${state.rows.length} registros.`;
      $('contenido').hidden = false;
      renderAll();
    } catch (err) {
      console.error(err);
      $('estado-carga').innerHTML = `<span style="color:var(--danger)">${esc(err.message)}</span>`;
    }
  }

  function renderAll() {
    renderResumen();
    renderMemos();
    renderDevoluciones();
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
    $('memos-resumen').textContent = `${groups.length} empresa(s) con IPE en estado APLICAR → ${groups.length} memo(s).`;
    $('btn-zip').disabled = !groups.length;
    $('memos').innerHTML = groups.map((g, gi) => `
      <div class="card" data-g="${gi}">
        <div class="memo-head">
          <h3>${esc(g.razon)} <span class="muted">· RUT ${esc(g.rutFmt)}</span></h3>
          <button class="secondary" data-download="${gi}">Descargar Word</button>
        </div>
        <div class="scroll">
        <table>
          <thead><tr><th>Incluir</th><th>Doc.</th><th>Nro.</th><th>Fecha</th><th class="n">Saldo</th><th>XX</th><th>Justificación (MOTIVO)</th><th>Docto a aplicar</th><th>Monto a aplicar</th></tr></thead>
          <tbody>
            ${g.rows.map((r, ri) => `
              <tr class="${r.naturaleza === 'DEUDOR' ? 'deudor' : ''}">
                <td><input type="checkbox" data-f="incluir" data-r="${ri}" ${r.incluir ? 'checked' : ''} aria-label="Incluir ${esc(r.doc)} ${esc(r.nro)}"></td>
                <td>${esc(r.doc)}</td><td>${esc(r.nro)}</td><td>${esc(fmt.fechaCorta(r.fecha))}</td>
                <td class="n">${clp(r.saldo)}</td><td>${esc(r.naturaleza)}</td>
                <td><input type="text" data-f="justificacion" data-r="${ri}" value="${esc(r.justificacion)}"></td>
                <td><input type="text" data-f="doctoAplicar" data-r="${ri}" value="${esc(r.doctoAplicar)}" placeholder="Pagar FEE …"></td>
                <td><input type="text" data-f="montoAplicar" data-r="${ri}" value="${esc(r.montoAplicar)}"></td>
              </tr>`).join('')}
          </tbody>
        </table>
        </div>
        <div class="preview" data-preview="${gi}"></div>
        <ul class="warnings" data-warn="${gi}"></ul>
      </div>`).join('') || '<p class="muted">No hay registros con estado APLICAR.</p>';
    groups.forEach((g, gi) => refreshGroup(gi));
  }

  function refreshGroup(gi) {
    const g = state.memoGroups[gi];
    const sel = g.rows.filter((r) => r.incluir);
    const total = sel.reduce((s, r) => s + r.saldoNeto, 0);
    const motivo = sel.map((r) => `<b>${esc(r.doc)} ${esc(r.nro)}</b> ${esc(r.justificacion)}`).join('<br>');
    document.querySelector(`[data-preview="${gi}"]`).innerHTML =
      `<div><b>MOTIVO:</b><br>${motivo || '—'}</div><div>Seleccionados: ${sel.length} · saldo neto ${clp(total)}</div>`;
    document.querySelector(`[data-warn="${gi}"]`).innerHTML = memo.warnings(g).map((w) => `<li>${esc(w)}</li>`).join('');
  }

  async function getTemplate() {
    if (!state.template) {
      const res = await fetch('templates/memo_template.docx');
      if (!res.ok) throw new Error('No se pudo cargar la plantilla del memo.');
      state.template = await res.arrayBuffer();
    }
    return state.template;
  }

  async function renderMemo(g) {
    const tpl = await getTemplate();
    return memo.render(tpl, memo.memoData(g, memoOpts()), PizZip, window.docxtemplater);
  }

  async function downloadOne(gi) {
    const g = state.memoGroups[gi];
    try {
      const zip = await renderMemo(g);
      saveAs(zip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), memo.fileName(g));
    } catch (err) { alertError(err); }
  }

  async function downloadAll() {
    try {
      const out = new PizZip();
      for (const g of state.memoGroups) {
        if (!g.rows.some((r) => r.incluir)) continue;
        const zip = await renderMemo(g);
        out.file(memo.fileName(g), zip.generate({ type: 'uint8array' }));
      }
      saveAs(out.generate({ type: 'blob' }), `Memos_IPE_${$('cfg-fecha').value}.zip`);
    } catch (err) { alertError(err); }
  }

  function alertError(err) {
    console.error(err);
    const detalle = err.properties && err.properties.errors ? err.properties.errors.map((e) => e.properties.explanation).join('\n') : '';
    alert(`No se pudo generar el memo: ${err.message}\n${detalle}`);
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
    XLSX.writeFile(wb, `Devoluciones_IPE_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // ---------- eventos ----------
  function init() {
    loadConfig();
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
    ['cfg-iniciales', 'cfg-firmante', 'cfg-cargo'].forEach((id) => $(id).addEventListener('change', saveConfig));

    $('memos').addEventListener('input', (e) => {
      const card = e.target.closest('[data-g]');
      if (!card || !e.target.dataset.f) return;
      const gi = Number(card.dataset.g);
      const row = state.memoGroups[gi].rows[Number(e.target.dataset.r)];
      row[e.target.dataset.f] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
      refreshGroup(gi);
    });
    $('memos').addEventListener('click', (e) => {
      const b = e.target.closest('[data-download]');
      if (b) downloadOne(Number(b.dataset.download));
    });
    $('btn-zip').addEventListener('click', downloadAll);
    $('btn-dev').addEventListener('click', downloadDevoluciones);
  }

  init();
})();
