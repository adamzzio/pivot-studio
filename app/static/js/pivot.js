(function () {
  'use strict';
  const { $, $$, el, icon, toast, request, openDialog, debounce, downloadBlob, setLoading, bindDropzone } = window.UI;

  const AGGS = [
    { v: 'SUM', t: 'SUM', hint: 'Jumlah nilai numerik' },
    { v: 'AVERAGE', t: 'AVERAGE', hint: 'Rata-rata nilai numerik' },
    { v: 'COUNT', t: 'COUNT (semua baris)', hint: 'Menghitung seluruh baris, termasuk yang kosong' },
    { v: 'COUNT_NON_NULL', t: 'COUNT NON-NULL', hint: 'Hanya menghitung nilai yang tidak kosong' },
    { v: 'COUNT_DISTINCT', t: 'COUNT DISTINCT', hint: 'Jumlah nilai unik, kosong diabaikan' },
    { v: 'VAR', t: 'VAR (population)', hint: 'Varians populasi (VAR.P)' },
    { v: 'STD', t: 'STD (population)', hint: 'Standar deviasi populasi (STD.P)' },
    { v: 'MEDIAN', t: 'MEDIAN', hint: 'Nilai tengah' },
    { v: 'MODE', t: 'MODE', hint: 'Nilai paling sering; jika seri dipilih yang terkecil' },
  ];
  const INT_AGGS = new Set(['COUNT', 'COUNT_NON_NULL', 'COUNT_DISTINCT']);
  const TYPE_TXT = { number: '123', text: 'Abc', date: 'Tgl' };
  const ZONES = [
    { key: 'filters', title: 'Filters' },
    { key: 'columns', title: 'Columns' },
    { key: 'rows', title: 'Rows' },
    { key: 'values', title: 'Values' },
  ];
  const PAGE = 500;

  const S = {
    files: [], hidden: new Set(), source: null, columns: [], typeOverrides: {},
    totalRows: 0, totalCols: 0, previewOffset: 0,
    rows: [], columns_: [], values: [], filters: [],
    opts: { decimals: 2, subtotals: true, grandTotal: true, sortBy: 'label', sortDir: 'asc', sortMetric: 0, colDir: 'asc' },
    result: null, resultError: '', resultLoading: false, loaded: 0, totalBody: 0, seq: 0, abort: null,
    ai: { state: 'idle', insight: null, raw: '', error: '', stale: false, confirmed: false, truncated: false },
  };
  const arr = (key) => (key === 'columns' ? S.columns_ : S[key]);
  const colType = (name) => (S.columns.find((c) => c.name === name) || {}).type || 'text';

  // ---------------------------------------------------------------- Upload dan sumber
  async function loadFiles() {
    try { const d = await request('/api/pivot/files', { method: 'GET' }); S.files = d.files; renderFiles(); } catch (_) { /* sesi baru */ }
  }

  async function upload(files) {
    const bad = files.filter((f) => !f.name.toLowerCase().endsWith('.xlsx'));
    if (bad.length) toast({ title: 'Format tidak didukung', desc: `${bad.map((f) => f.name).join(', ')} bukan file .xlsx.`, variant: 'error' });
    const good = files.filter((f) => f.name.toLowerCase().endsWith('.xlsx'));
    if (!good.length) return;
    if (good.some((f) => f.size > 100 * 1048576)) toast({ title: 'File besar', desc: 'Pembacaan bisa lama dan memakai banyak memori.', variant: 'warning' });
    const fd = new FormData(); good.forEach((f) => fd.append('files', f));
    const bar = $('#upload-progress'); bar.hidden = false;
    try {
      const d = await request('/api/pivot/upload', { form: fd, slowMs: 5000, slowMsg: 'Unggahan dan pembacaan file masih berjalan.' });
      S.files.push(...d.files);
      d.errors.forEach((e) => toast({ title: 'File ditolak', desc: e.error, variant: 'error', duration: 8000 }));
      if (d.files.length) toast({ title: 'Upload berhasil', desc: `${d.files.length} file ditambahkan.`, variant: 'success' });
      renderFiles();
      if (!S.source && d.files.length === 1 && d.files[0].sheets.length === 1) selectSource(d.files[0], d.files[0].sheets[0].name);
    } catch (e) { toast({ title: 'Upload gagal', desc: e.message, variant: 'error', duration: 8000 }); }
    finally { bar.hidden = true; }
  }

  function renderFiles() {
    const box = $('#files-list'); box.replaceChildren();
    $('#btn-clear-all').hidden = !S.files.length;
    if (!S.files.length) { box.append(el('div', { class: 'empty', text: 'Belum ada file. Unggah file XLSX untuk memulai.' })); return; }
    S.files.forEach((f) => {
      const sheets = f.sheets.filter((s) => !S.hidden.has(`${f.file_id}::${s.name}`));
      const rows = sheets.map((s) => {
        const active = S.source && S.source.file_id === f.file_id && S.source.sheet === s.name;
        const dims = s.exact ? `${s.rows.toLocaleString('id-ID')} baris, ${s.cols} kolom` : (s.rows ? `sekitar ${s.rows.toLocaleString('id-ID')} baris, ${s.cols || '?'} kolom` : 'ukuran belum diketahui');
        return el('div', { class: `sheet-row ${active ? 'active' : ''}` },
          el('div', { class: 'sheet-name' }, el('strong', { text: s.name }), el('span', { class: 'muted small', text: dims })),
          el('div', { class: 'chip-row' },
            el('button', { class: `btn btn-sm ${active ? 'btn-primary' : 'btn-outline'}`, type: 'button', 'aria-pressed': String(active), text: active ? 'Dipakai' : 'Gunakan', onclick: () => selectSource(f, s.name) }),
            el('button', { class: 'btn btn-ghost btn-icon btn-sm', type: 'button', 'aria-label': `Sembunyikan sheet ${s.name}`, title: 'Hapus sheet dari workspace', html: icon('x'), onclick: () => hideSheet(f, s.name) })));
      });
      box.append(el('div', { class: 'file-card' },
        el('div', { class: 'file-card-head' },
          el('div', { class: 'chip-row', style: 'min-width:0' }, el('span', { html: icon('file') }), el('span', { class: 'file-name', text: f.name }), el('span', { class: 'badge', text: `${sheets.length} sheet` })),
          el('button', { class: 'btn btn-ghost btn-sm', type: 'button', text: 'Hapus file', onclick: () => removeFile(f) })),
        rows.length ? rows : el('div', { class: 'sheet-row muted', text: 'Semua sheet disembunyikan.' })));
    });
  }

  function hideSheet(f, name) {
    S.hidden.add(`${f.file_id}::${name}`);
    if (S.source && S.source.file_id === f.file_id && S.source.sheet === name) clearSource();
    renderFiles();
  }
  async function removeFile(f) {
    try { await request(`/api/pivot/file/${f.file_id}`, { method: 'DELETE' }); } catch (e) { toast({ title: 'Gagal menghapus', desc: e.message, variant: 'error' }); return; }
    S.files = S.files.filter((x) => x.file_id !== f.file_id);
    if (S.source && S.source.file_id === f.file_id) clearSource();
    renderFiles();
  }
  async function clearAll() {
    try { await request('/api/pivot/clear', {}); } catch (e) { toast({ title: 'Gagal', desc: e.message, variant: 'error' }); return; }
    S.files = []; S.hidden.clear(); clearSource(); renderFiles();
  }
  function clearSource() {
    S.source = null; S.columns = []; S.rows = []; S.columns_ = []; S.values = []; S.filters = []; S.typeOverrides = {};
    S.result = null; S.ai.insight = null; S.ai.raw = ''; S.ai.state = 'idle'; S.ai.stale = false;
    ['#preview-card', '#builder-card', '#result-card', '#ai-card'].forEach((s) => { $(s).hidden = true; });
  }

  async function selectSource(f, sheet) {
    const prev = S.source;
    S.source = { file_id: f.file_id, sheet, name: f.name };
    S.typeOverrides = {}; S.rows = []; S.columns_ = []; S.values = []; S.filters = [];
    S.result = null; S.resultError = ''; S.ai = { ...S.ai, state: 'idle', insight: null, raw: '', error: '', stale: false };
    S.previewOffset = 0;
    renderFiles();
    try {
      await loadPreview(0);
      const sh = f.sheets.find((s) => s.name === sheet);
      if (sh) { sh.rows = S.totalRows; sh.cols = S.totalCols; sh.exact = true; }
      ['#preview-card', '#builder-card', '#result-card', '#ai-card'].forEach((s) => { $(s).hidden = false; });
      $('#prev-sub').textContent = `${f.name} / ${sheet}`;
      renderFiles(); renderBuilder(); renderResult(); renderAI(); updateSummary();
    } catch (e) {
      S.source = prev; renderFiles();
      toast({ title: 'Sheet tidak dapat dipakai', desc: e.message, variant: 'error', duration: 8000 });
    }
  }

  // ---------------------------------------------------------------- Preview dan tipe kolom
  async function loadPreview(offset) {
    const d = await request('/api/pivot/preview', { json: { file_id: S.source.file_id, sheet: S.source.sheet, types: S.typeOverrides, offset, limit: 25 }, slowMs: 5000, slowMsg: 'Membaca sheet masih berjalan.' });
    S.columns = d.columns; S.totalRows = d.total_rows; S.totalCols = d.total_cols; S.previewOffset = d.offset;
    renderPreview(d); renderTypes();
  }

  function renderPreview(d) {
    const t = $('#preview-table'); t.replaceChildren();
    t.append(el('thead', {}, el('tr', {}, d.columns.map((c) => el('th', {}, c.name, el('span', { class: 'type-badge', text: TYPE_TXT[c.type] }))))));
    t.append(el('tbody', {}, d.rows.map((r) => el('tr', {}, r.map((v) => el('td', { text: v }))))));
    const end = Math.min(d.offset + d.rows.length, d.total_rows);
    $('#prev-info').textContent = `Baris ${d.total_rows ? d.offset + 1 : 0} sampai ${end} dari ${d.total_rows.toLocaleString('id-ID')} (${d.total_cols} kolom)`;
    $('#prev-prev').disabled = d.offset === 0;
    $('#prev-next').disabled = end >= d.total_rows;
  }

  function renderTypes() {
    const box = $('#types-list'); box.replaceChildren();
    S.columns.forEach((c) => {
      const sel = el('select', { class: 'select', 'aria-label': `Tipe kolom ${c.name}` },
        el('option', { value: '', text: `Otomatis (${TYPE_TXT[c.inferred]})` }),
        el('option', { value: 'number', text: 'Angka' }), el('option', { value: 'text', text: 'Teks' }), el('option', { value: 'date', text: 'Tanggal' }));
      sel.value = S.typeOverrides[c.name] || '';
      sel.addEventListener('change', () => changeType(c.name, sel.value));
      box.append(el('div', { class: 'type-row' }, el('div', { style: 'min-width:0' }, el('div', { class: 'file-name', text: c.name }), el('div', { class: 'muted small', text: `${c.nulls.toLocaleString('id-ID')} kosong` })), sel));
    });
  }

  async function changeType(name, v) {
    if (v) S.typeOverrides[name] = v; else delete S.typeOverrides[name];
    const before = S.filters.length;
    S.filters = S.filters.filter((f) => f.field !== name);
    try { await loadPreview(S.previewOffset); } catch (e) { toast({ title: 'Gagal memperbarui tipe', desc: e.message, variant: 'error' }); }
    if (S.filters.length !== before) toast({ title: 'Filter dihapus', desc: `Filter pada '${name}' dihapus karena tipe kolom berubah.`, variant: 'info' });
    renderBuilder(); configChanged();
  }

  // ---------------------------------------------------------------- Builder
  let drag = null;

  function renderBuilder() { renderFieldList(); renderZones(); renderOptions(); }

  function renderFieldList() {
    const q = $('#field-search').value.trim().toLowerCase();
    const box = $('#field-list'); box.replaceChildren();
    S.columns.filter((c) => c.name.toLowerCase().includes(q)).forEach((c) => {
      const n = el('div', { class: 'fchip', title: c.name }, el('span', { class: 'grip', html: icon('grip') }), el('span', { class: 'name file-name', text: c.name }), el('span', { class: 'tbadge', text: TYPE_TXT[c.type] }));
      makeDraggable(n, { field: c.name, from: null, index: -1 });
      box.append(n);
    });
    if (!box.children.length) box.append(el('div', { class: 'muted small', text: 'Tidak ada field yang cocok.' }));
  }

  function makeDraggable(node, payload) {
    node.draggable = true;
    node.addEventListener('dragstart', (e) => { drag = payload; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', payload.field); node.classList.add('dragging'); });
    node.addEventListener('dragend', () => { drag = null; node.classList.remove('dragging'); $$('.zone.over').forEach((z) => z.classList.remove('over')); });
  }

  function dropIndex(body, e) {
    const kids = $$(':scope > .chip', body);
    for (let i = 0; i < kids.length; i++) {
      const r = kids[i].getBoundingClientRect();
      if (e.clientY < r.top + r.height / 2) return i;
    }
    return kids.length;
  }

  function renderZones() {
    const root = $('#zones'); root.replaceChildren();
    ZONES.forEach(({ key, title }) => {
      const list = arr(key);
      const body = el('div', { class: 'zone-body' });
      if (!list.length) body.append(el('div', { class: 'zone-empty', text: key === 'values' ? 'Seret field angka atau kategori ke sini' : 'Seret field ke sini' }));
      list.forEach((item, i) => body.append(chip(key, item, i)));

      const add = el('select', { class: 'select', 'aria-label': `Tambah field ke ${title}` }, el('option', { value: '', text: '+ Tambah field' }));
      S.columns.forEach((c) => {
        const blocked = (key === 'rows' && (S.rows.includes(c.name) || S.columns_.includes(c.name))) || (key === 'columns' && (S.rows.includes(c.name) || S.columns_.includes(c.name))) || (key === 'filters' && S.filters.some((f) => f.field === c.name));
        if (!blocked) add.append(el('option', { value: c.name, text: c.name }));
      });
      add.addEventListener('change', () => { if (add.value) addTo(key, add.value, null); });

      const zone = el('div', { class: 'zone', 'data-zone': key }, el('div', { class: 'zone-head' }, el('strong', { text: `${title}${list.length ? ` (${list.length})` : ''}` }), add), body);
      zone.addEventListener('dragover', (e) => { if (drag) { e.preventDefault(); zone.classList.add('over'); } });
      zone.addEventListener('dragleave', (e) => { if (!zone.contains(e.relatedTarget)) zone.classList.remove('over'); });
      zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('over'); if (drag) moveField(drag, key, dropIndex(body, e)); drag = null; });
      root.append(zone);
    });
    updateSummary();
  }

  function mvButtons(key, i, n) {
    return [
      el('button', { class: 'btn btn-ghost btn-icon btn-sm', type: 'button', 'aria-label': 'Pindah ke atas', disabled: i === 0, html: icon('up'), onclick: () => { reorder(key, i, i - 1); } }),
      el('button', { class: 'btn btn-ghost btn-icon btn-sm', type: 'button', 'aria-label': 'Pindah ke bawah', disabled: i === n - 1, html: icon('down'), onclick: () => { reorder(key, i, i + 2); } }),
    ];
  }
  const rmBtn = (key, i, name) => el('button', { class: 'btn btn-ghost btn-icon btn-sm', type: 'button', 'aria-label': `Hapus ${name}`, html: icon('x'), onclick: () => { arr(key).splice(i, 1); changed(); } });

  function chip(key, item, i) {
    const list = arr(key);
    if (key === 'rows' || key === 'columns') {
      const n = el('div', { class: 'chip', tabindex: 0 }, el('span', { class: 'grip', html: icon('grip') }), el('span', { class: 'name', text: item, title: item }), mvButtons(key, i, list.length), rmBtn(key, i, item));
      makeDraggable(n, { field: item, from: key, index: i });
      return n;
    }
    if (key === 'values') {
      const type = colType(item.field);
      const sel = el('select', { class: 'select', 'aria-label': `Fungsi agregasi untuk ${item.field}`, title: AGGS.find((a) => a.v === item.agg).hint },
        AGGS.map((a) => el('option', { value: a.v, text: a.t, selected: a.v === item.agg })));
      sel.addEventListener('change', () => { item.agg = sel.value; changed(); });
      const lbl = el('input', { class: 'input', type: 'text', maxlength: 80, placeholder: 'Nama label (opsional)', 'aria-label': `Label untuk ${item.field}`, value: item.label || '' });
      lbl.addEventListener('input', () => { item.label = lbl.value; changedSoft(); });
      const n = el('div', { class: 'chip metric' },
        el('div', { class: 'chip-row' }, el('span', { class: 'grip', html: icon('grip') }), el('span', { class: 'name', text: item.field, title: item.field }), el('span', { class: 'tbadge', text: TYPE_TXT[type] }), mvButtons(key, i, list.length), rmBtn(key, i, item.field)),
        sel, lbl);
      makeDraggable(n, { field: item.field, from: key, index: i });
      return n;
    }
    // filters
    const n = el('div', { class: 'chip filter', tabindex: 0, role: 'button', title: 'Klik untuk mengubah filter' },
      el('span', { html: icon('filter') }), el('span', { class: 'name' }, item.field, el('div', { class: 'sum', text: filterSummary(item) })), rmBtn(key, i, `filter ${item.field}`));
    n.addEventListener('click', async (e) => { if (e.target.closest('button')) return; const f = await editFilter(item.field, item); if (f) { list[i] = f; changed(); } });
    n.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target === n) n.click(); });
    makeDraggable(n, { field: item.field, from: key, index: i });
    return n;
  }

  function filterSummary(f) {
    if (f.kind === 'values') return `${(f.values || []).length} nilai dipilih`;
    if (f.kind === 'range') return `${f.min ?? '...'} s/d ${f.max ?? '...'}`;
    return `${f.start || '...'} s/d ${f.end || '...'}`;
  }

  function reorder(key, from, to) {
    const list = arr(key); const item = list.splice(from, 1)[0];
    if (to > from) to -= 1;
    list.splice(Math.max(0, Math.min(to, list.length)), 0, item);
    changed();
  }

  function defaultAgg(field) { return colType(field) === 'number' ? 'SUM' : 'COUNT'; }

  function moveField(d, to, idx) {
    if (d.from === to) { reorder(to, d.index, idx); return; }
    if (d.from) { const src = arr(d.from); if (src[d.index] !== undefined) src.splice(d.index, 1); }
    addTo(to, d.field, idx, true);
  }

  async function addTo(key, field, idx) {
    const list = arr(key);
    const at = idx === null || idx === undefined ? list.length : Math.min(idx, list.length);
    if (key === 'rows' || key === 'columns') {
      S.rows = S.rows.filter((f) => f !== field); S.columns_ = S.columns_.filter((f) => f !== field);
      const target = arr(key); target.splice(Math.min(at, target.length), 0, field);
    } else if (key === 'values') {
      list.splice(at, 0, { field, agg: defaultAgg(field), label: '' });
    } else if (key === 'filters') {
      const existing = S.filters.find((f) => f.field === field);
      const f = await editFilter(field, existing || null);
      if (f) { if (existing) S.filters[S.filters.indexOf(existing)] = f; else S.filters.splice(Math.min(at, S.filters.length), 0, f); }
    }
    changed();
  }

  // ---------------------------------------------------------------- Dialog filter
  async function editFilter(field, existing) {
    const t = colType(field);
    let info = null;
    try { info = await request('/api/pivot/field-values', { json: { file_id: S.source.file_id, sheet: S.source.sheet, types: S.typeOverrides, field, search: '', limit: 300 }, slowMs: 5000 }); }
    catch (e) { toast({ title: 'Gagal memuat nilai filter', desc: e.message, variant: 'error' }); return null; }
    let out = null;
    const body = el('div', { class: 'stack-sm' });

    if (t === 'number') {
      const lo = el('input', { class: 'input', type: 'number', step: 'any', placeholder: info.min !== undefined ? `min ${info.min}` : 'minimum', value: existing && existing.min !== null && existing.min !== undefined ? existing.min : '' });
      const hi = el('input', { class: 'input', type: 'number', step: 'any', placeholder: info.max !== undefined ? `maks ${info.max}` : 'maksimum', value: existing && existing.max !== null && existing.max !== undefined ? existing.max : '' });
      body.append(el('div', { class: 'options' }, el('label', { class: 'field' }, el('span', { text: 'Minimum' }), lo), el('label', { class: 'field' }, el('span', { text: 'Maksimum' }), hi)));
      const ok = await openDialog({ title: `Filter rentang: ${field}`, body, confirmText: 'Terapkan filter', onConfirm: () => {
        if (lo.value === '' && hi.value === '') { toast({ title: 'Isi minimal satu batas', variant: 'warning' }); return false; }
        out = { field, kind: 'range', min: lo.value === '' ? null : Number(lo.value), max: hi.value === '' ? null : Number(hi.value) };
      } });
      return ok ? out : null;
    }
    if (t === 'date') {
      const a = el('input', { class: 'input', type: 'date', min: info.min, max: info.max, value: (existing && existing.start) || '' });
      const b = el('input', { class: 'input', type: 'date', min: info.min, max: info.max, value: (existing && existing.end) || '' });
      body.append(el('div', { class: 'options' }, el('label', { class: 'field' }, el('span', { text: 'Dari tanggal' }), a), el('label', { class: 'field' }, el('span', { text: 'Sampai tanggal' }), b)));
      const ok = await openDialog({ title: `Filter tanggal: ${field}`, body, confirmText: 'Terapkan filter', onConfirm: () => {
        if (!a.value && !b.value) { toast({ title: 'Isi minimal satu tanggal', variant: 'warning' }); return false; }
        out = { field, kind: 'date_range', start: a.value || null, end: b.value || null };
      } });
      return ok ? out : null;
    }

    // kategorikal
    const selected = new Set(existing && existing.values ? existing.values : (info.total_unique <= 300 ? info.values.map((v) => v.label) : []));
    const search = el('input', { class: 'input', type: 'search', placeholder: 'Cari nilai...', 'aria-label': 'Cari nilai filter' });
    const list = el('div', { class: 'check-list' });
    const status = el('div', { class: 'muted small' });
    let shown = info.values;
    const paint = () => {
      list.replaceChildren();
      shown.forEach((v) => {
        const cb = el('input', { type: 'checkbox', checked: selected.has(v.label) });
        cb.addEventListener('change', () => { if (cb.checked) selected.add(v.label); else selected.delete(v.label); updStatus(); });
        list.append(el('label', {}, cb, el('span', { text: v.label }), el('span', { class: 'cnt', text: v.count.toLocaleString('id-ID') })));
      });
      if (!shown.length) list.append(el('div', { class: 'muted small', text: 'Tidak ada nilai yang cocok.' }));
      updStatus();
    };
    const updStatus = () => { status.textContent = `${selected.size} dipilih. Menampilkan ${shown.length} dari ${info.matched ?? shown.length} nilai cocok (${info.total_unique} unik).`; };
    search.addEventListener('input', debounce(async () => {
      try { const d = await request('/api/pivot/field-values', { json: { file_id: S.source.file_id, sheet: S.source.sheet, types: S.typeOverrides, field, search: search.value, limit: 300 }, slowMs: 5000 }); shown = d.values; info.matched = d.matched; paint(); }
      catch (e) { toast({ title: 'Gagal mencari', desc: e.message, variant: 'error' }); }
    }, 250));
    const allBtn = el('button', { class: 'btn btn-outline btn-sm', type: 'button', text: 'Pilih yang tampil', onclick: () => { shown.forEach((v) => selected.add(v.label)); paint(); } });
    const noneBtn = el('button', { class: 'btn btn-outline btn-sm', type: 'button', text: 'Kosongkan pilihan', onclick: () => { selected.clear(); paint(); } });
    body.append(search, el('div', { class: 'chip-row' }, allBtn, noneBtn), list, status);
    paint();
    const ok = await openDialog({ title: `Filter nilai: ${field}`, body, confirmText: 'Terapkan filter', onConfirm: () => {
      if (!selected.size) { toast({ title: 'Pilih minimal satu nilai', variant: 'warning' }); return false; }
      out = { field, kind: 'values', values: Array.from(selected) };
    } });
    return ok ? out : null;
  }

  // ---------------------------------------------------------------- Opsi
  function renderOptions() {
    const o = S.opts;
    $('#opt-decimals').value = o.decimals; $('#opt-sub').checked = o.subtotals; $('#opt-grand').checked = o.grandTotal;
    $('#opt-sort-by').value = o.sortBy; $('#opt-sort-dir').value = o.sortDir; $('#opt-col-dir').value = o.colDir;
    const sm = $('#opt-sort-metric'); sm.replaceChildren();
    S.values.forEach((v, i) => sm.append(el('option', { value: i, text: v.label || `${v.agg} ${v.field}`, selected: i === o.sortMetric })));
    $('#wrap-sort-metric').hidden = o.sortBy !== 'value' || !S.values.length;
  }

  function updateSummary() {
    const r = S.result && S.result.meta;
    const parts = `Rows ${S.rows.length}, Columns ${S.columns_.length}, Values ${S.values.length}, Filters ${S.filters.length}`;
    $('#config-summary').textContent = r ? `${parts}. Record dianalisis: ${r.records.toLocaleString('id-ID')} dari ${r.total_records.toLocaleString('id-ID')}.` : `${parts}.`;
    const fc = $('#filter-chips');
    fc.textContent = S.filters.length ? `Filter aktif: ${S.filters.map((f) => `${f.field} (${filterSummary(f)})`).join('; ')}` : '';
  }

  // ---------------------------------------------------------------- Perubahan konfigurasi
  function markStale() { if (S.ai.insight || S.ai.raw) { S.ai.stale = true; renderAI(); } }
  function changed() { renderZones(); renderOptions(); markStale(); scheduleCalc(); }
  function changedSoft() { markStale(); scheduleCalc(); }
  function configChanged() { markStale(); scheduleCalc(); }

  function buildConfig() {
    return {
      source: { file_id: S.source.file_id, sheet: S.source.sheet },
      types: S.typeOverrides,
      rows: S.rows, columns: S.columns_,
      values: S.values.map((v) => ({ field: v.field, agg: v.agg, label: v.label || '' })),
      filters: S.filters,
      options: { subtotals: S.opts.subtotals, grand_total: S.opts.grandTotal, row_sort: { by: S.opts.sortBy, dir: S.opts.sortDir, metric: S.opts.sortMetric }, col_sort: { dir: S.opts.colDir } },
    };
  }

  const scheduleCalc = debounce(() => calculate(true), 300);

  async function calculate(reset) {
    if (!S.source) return;
    const seq = ++S.seq;
    if (S.abort) S.abort.abort();
    if (!S.values.length) { S.result = null; S.resultError = ''; S.resultLoading = false; renderResult(); updateSummary(); return; }
    S.abort = new AbortController();
    S.resultLoading = true; if (reset) { S.resultError = ''; } renderResult();
    try {
      const d = await request('/api/pivot/calculate', { json: { ...buildConfig(), page: { offset: reset ? 0 : S.loaded, limit: PAGE } }, signal: S.abort.signal, slowMs: 5000, slowMsg: 'Perhitungan pivot masih berjalan.' });
      if (seq !== S.seq) return;
      if (reset || !S.result) { S.result = d.result; S.loaded = d.result.body.length; }
      else { S.result.body.push(...d.result.body); S.loaded += d.result.body.length; }
      S.totalBody = d.result.total_rows; S.resultError = '';
    } catch (e) {
      if (e.name === 'AbortError' || seq !== S.seq) return;
      S.resultError = e.message; if (reset) S.result = null;
    }
    S.resultLoading = false; renderResult(); updateSummary();
  }

  // ---------------------------------------------------------------- Tampilan hasil
  function fmt(v, mi) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    const agg = S.result.header.metrics[mi].agg;
    const d = INT_AGGS.has(agg) ? 0 : S.opts.decimals;
    return v.toLocaleString('id-ID', { minimumFractionDigits: d, maximumFractionDigits: d });
  }

  function renderResult() {
    const body = $('#res-body'), notes = $('#res-notes'); notes.replaceChildren();
    $('#btn-export').disabled = !S.result;
    $('#btn-ai').disabled = !S.result || S.ai.state === 'loading';
    $('#res-more').hidden = true;
    if (S.resultError) { body.replaceChildren(el('div', { class: 'alert alert-error', role: 'alert', text: S.resultError })); $('#res-sub').textContent = ''; return; }
    if (!S.values.length) { body.replaceChildren(el('div', { class: 'empty', text: 'Tambahkan minimal satu field ke area Values untuk melihat pivot.' })); $('#res-sub').textContent = ''; return; }
    if (S.resultLoading && !S.result) { body.replaceChildren(el('div', {}, el('div', { class: 'skeleton' }), el('div', { class: 'skeleton' }), el('div', { class: 'skeleton' }))); return; }
    if (!S.result) return;
    const r = S.result, h = r.header, M = h.metrics.length;
    $('#res-sub').textContent = `${r.meta.records.toLocaleString('id-ID')} record dianalisis dari ${r.meta.total_records.toLocaleString('id-ID')}${S.resultLoading ? ' (memperbarui...)' : ''}`;
    r.meta.notes.forEach((n) => notes.append(el('div', { class: 'alert alert-warn small', text: n })));
    if (!r.body.length || r.meta.records === 0) { body.replaceChildren(el('div', { class: 'empty', text: 'Tidak ada data yang cocok dengan filter saat ini.' })); return; }

    const C = h.col_fields.length, nrc = Math.max(1, h.row_fields.length);
    const headRows = C === 0 ? 1 : C + (M > 1 ? 1 : 0);
    const thead = el('thead');
    const rowsEl = Array.from({ length: headRows }, () => el('tr'));
    for (let i = 0; i < nrc; i++) rowsEl[0].append(el('th', { rowspan: headRows, text: h.row_fields[i] || 'Total' }));
    if (C === 0) {
      h.metrics.forEach((m) => rowsEl[0].append(el('th', { text: m.label })));
    } else {
      for (let l = 0; l < C; l++) {
        let i = 0;
        while (i < h.col_paths.length) {
          let j = i;
          while (j < h.col_paths.length && h.col_paths[j].slice(0, l + 1).join('\u0001') === h.col_paths[i].slice(0, l + 1).join('\u0001')) j++;
          rowsEl[l].append(el('th', { colspan: (j - i) * M, text: h.col_paths[i][l], title: h.col_fields[l] }));
          i = j;
        }
      }
      if (h.grand_col) rowsEl[0].append(el('th', { colspan: M, rowspan: C, text: 'Grand Total' }));
      if (M > 1) {
        const blocks = h.col_paths.length + (h.grand_col ? 1 : 0);
        for (let b = 0; b < blocks; b++) h.metrics.forEach((m) => rowsEl[C].append(el('th', { text: m.label })));
      }
    }
    rowsEl.forEach((r2) => thead.append(r2));

    const tbody = el('tbody');
    r.body.forEach((entry) => {
      const tr = el('tr', { class: entry.kind === 'detail' ? '' : entry.kind });
      for (let i = 0; i < nrc; i++) {
        let t = '';
        if (entry.kind === 'grand') t = i === 0 ? 'Grand Total' : '';
        else if (i < entry.labels.length) t = entry.labels[i];
        else if (entry.kind === 'subtotal' && i === entry.labels.length) t = 'Subtotal';
        tr.append(el('th', { scope: 'row', text: t }));
      }
      entry.values.forEach((v, j) => tr.append(el('td', { class: typeof v === 'number' ? 'num' : '', text: fmt(v, j % M) })));
      tbody.append(tr);
    });
    body.replaceChildren(el('div', { class: 'table-wrap' }, el('table', { class: 'pivot' }, thead, tbody)));

    if (S.loaded < S.totalBody) {
      $('#res-more').hidden = false;
      $('#res-more-info').textContent = `Menampilkan ${S.loaded.toLocaleString('id-ID')} dari ${S.totalBody.toLocaleString('id-ID')} baris`;
    }
  }

  // ---------------------------------------------------------------- Export
  async function exportXlsx() {
    const btn = $('#btn-export'); setLoading(btn, true, 'Mengekspor...');
    try {
      const { blob } = await request('/api/pivot/export', { json: { ...buildConfig(), decimals: S.opts.decimals }, as: 'blob', slowMs: 5000, slowMsg: 'Ekspor masih berjalan, mohon tunggu.' });
      downloadBlob(blob, 'pivot_table_export.xlsx');
      toast({ title: 'Ekspor berhasil', desc: 'pivot_table_export.xlsx', variant: 'success' });
    } catch (e) { toast({ title: 'Ekspor gagal', desc: e.message, variant: 'error', duration: 8000 }); }
    finally { setLoading(btn, false); btn.disabled = !S.result; }
  }

  // ---------------------------------------------------------------- AI Insights
  function section(title, items) {
    if (!items || !items.length) return null;
    return el('div', {}, el('h3', { text: title }), el('ul', {}, items.map((t) => el('li', { text: t }))));
  }

  function renderAI() {
    const box = $('#ai-body'), a = S.ai;
    $('#ai-stale').hidden = !a.stale;
    $('#btn-ai').disabled = !S.result || a.state === 'loading';
    $('#btn-ai').textContent = a.insight || a.raw ? 'Generate Insights lagi' : 'Generate Insights';
    if (a.state === 'loading') { box.replaceChildren(el('div', {}, el('div', { class: 'skeleton' }), el('div', { class: 'skeleton' }), el('div', { class: 'skeleton' }), el('p', { class: 'muted small', text: 'Menganalisis hasil pivot...' }))); return; }
    if (a.state === 'error') {
      box.replaceChildren(el('div', { class: 'alert alert-error', role: 'alert' }, el('div', { class: 'stack-sm' }, el('div', { text: a.error }), el('div', {}, el('button', { class: 'btn btn-outline btn-sm', type: 'button', text: 'Coba lagi', onclick: generate })))));
      return;
    }
    if (a.insight) {
      const i = a.insight;
      box.replaceChildren(el('div', { class: 'insight' },
        el('div', { class: 'title', text: i.title }),
        i.summary ? el('p', { text: i.summary }) : null,
        section('Temuan utama', i.findings), section('Nilai tertinggi dan terendah', i.extremes), section('Perbandingan antar kategori', i.comparisons),
        section('Anomali', i.anomalies), section('Rekomendasi', i.recommendations), section('Catatan dan keterbatasan', i.limitations),
        a.truncated ? el('p', { class: 'muted small', text: 'Data dipotong sebelum dikirim ke AI karena ukurannya besar.' }) : null));
      return;
    }
    if (a.raw) { box.replaceChildren(el('div', { class: 'insight' }, el('pre', { text: a.raw }))); return; }
    box.replaceChildren(el('div', { class: 'empty', text: S.result ? 'Tekan Generate Insights untuk membuat ringkasan dan rekomendasi dari pivot ini.' : 'Susun pivot terlebih dahulu. Tombol Generate Insights aktif setelah hasil pivot tersedia.' }));
  }

  async function generate() {
    if (!S.result) return;
    if (!S.ai.confirmed) {
      const ok = await openDialog({ title: 'Kirim data pivot ke Groq?', description: 'Ringkasan hasil pivot yang sedang tampil (bukan seluruh raw data) akan dikirim ke Groq untuk dianalisis. Pastikan data tidak mengandung informasi yang tidak boleh dibagikan.', confirmText: 'Kirim dan analisis' });
      if (!ok) return;
      S.ai.confirmed = true;
    }
    S.ai.state = 'loading'; renderAI();
    try {
      const d = await request('/api/ai/insights', { json: { ...buildConfig(), options: { language: $('#ai-lang').value, length: $('#ai-len').value } }, slowMs: 10000, slowMsg: 'Analisis AI masih berjalan, mohon tunggu.' });
      S.ai.insight = d.insight || null; S.ai.raw = d.raw_text || ''; S.ai.truncated = !!d.truncated; S.ai.stale = false; S.ai.state = 'done';
    } catch (e) { S.ai.state = 'error'; S.ai.error = e.message; }
    renderAI();
  }

  // ---------------------------------------------------------------- Event binding
  function resetConfig() {
    S.rows = []; S.columns_ = []; S.values = []; S.filters = [];
    S.opts = { decimals: 2, subtotals: true, grandTotal: true, sortBy: 'label', sortDir: 'asc', sortMetric: 0, colDir: 'asc' };
    changed(); toast({ title: 'Konfigurasi direset', variant: 'info', duration: 2500 });
  }

  function bind() {
    bindDropzone($('#dropzone'), $('#file-input'), upload);
    $('#btn-clear-all').addEventListener('click', clearAll);
    $$('[data-tab]').forEach((b) => b.addEventListener('click', () => {
      $$('[data-tab]').forEach((x) => { const on = x === b; x.classList.toggle('active', on); x.setAttribute('aria-selected', String(on)); });
      $('#tab-preview').hidden = b.dataset.tab !== 'preview'; $('#tab-types').hidden = b.dataset.tab !== 'types';
    }));
    const page = async (delta) => { try { await loadPreview(Math.max(0, S.previewOffset + delta)); } catch (e) { toast({ title: 'Gagal memuat preview', desc: e.message, variant: 'error' }); } };
    $('#prev-prev').addEventListener('click', () => page(-25));
    $('#prev-next').addEventListener('click', () => page(25));
    $('#field-search').addEventListener('input', renderFieldList);
    $('#btn-reset').addEventListener('click', resetConfig);
    $('#opt-decimals').addEventListener('input', (e) => { const v = Math.max(0, Math.min(8, parseInt(e.target.value, 10) || 0)); S.opts.decimals = v; renderResult(); });
    $('#opt-sub').addEventListener('change', (e) => { S.opts.subtotals = e.target.checked; changed(); });
    $('#opt-grand').addEventListener('change', (e) => { S.opts.grandTotal = e.target.checked; changed(); });
    $('#opt-sort-by').addEventListener('change', (e) => { S.opts.sortBy = e.target.value; changed(); });
    $('#opt-sort-metric').addEventListener('change', (e) => { S.opts.sortMetric = parseInt(e.target.value, 10) || 0; changed(); });
    $('#opt-sort-dir').addEventListener('change', (e) => { S.opts.sortDir = e.target.value; changed(); });
    $('#opt-col-dir').addEventListener('change', (e) => { S.opts.colDir = e.target.value; changed(); });
    $('#btn-more').addEventListener('click', () => calculate(false));
    $('#btn-export').addEventListener('click', exportXlsx);
    $('#btn-ai').addEventListener('click', generate);
  }

  bind();
  renderFiles();
  renderAI();
  loadFiles();
})();
