(function () {
  'use strict';
  const { $, $$, el, toast, request, formatBytes, bindDropzone, setLoading } = window.UI;

  const MODES = {
    'csv-to-xlsx': { ext: '.csv', outExt: '.xlsx', hint: 'Format yang diterima: .csv', label: 'CSV' },
    'xlsx-to-csv': { ext: '.xlsx', outExt: '.csv', hint: 'Format yang diterima: .xlsx', label: 'XLSX' },
  };
  let mode = 'csv-to-xlsx';
  let file = null;
  let downloadUrl = null;

  const dz = $('#dropzone'), input = $('#file-input');
  const info = $('#file-info'), opts = $('#options'), errBox = $('#error');
  const btnConvert = $('#btn-convert'), progress = $('#progress'), result = $('#result');

  function showError(msg) { errBox.textContent = msg; errBox.hidden = !msg; }
  function clearResult() {
    result.hidden = true;
    if (downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = null; }
  }

  function setMode(m) {
    mode = m;
    $$('.seg').forEach((b) => { const on = b.dataset.mode === m; b.classList.toggle('active', on); b.setAttribute('aria-selected', String(on)); });
    input.accept = MODES[m].ext;
    $('#dz-hint').textContent = MODES[m].hint;
    $('#opts-csv').hidden = m !== 'csv-to-xlsx';
    $('#opts-xlsx').hidden = m !== 'xlsx-to-csv';
    $('#xlsx-note').hidden = m !== 'xlsx-to-csv';
    resetFile();
  }

  function resetFile() {
    file = null; input.value = '';
    info.hidden = true; opts.hidden = true; showError(''); clearResult();
    btnConvert.disabled = true; progress.hidden = true;
  }

  async function pick(files) {
    const f = files[0];
    showError(''); clearResult();
    if (!f.name.toLowerCase().endsWith(MODES[mode].ext)) {
      showError(`Format file tidak sesuai. Untuk mode ini unggah file ${MODES[mode].ext}.`);
      return;
    }
    if (f.size === 0) { showError('File kosong.'); return; }
    if (f.size > 100 * 1048576) toast({ title: 'File besar', desc: 'Proses bisa memakan waktu lebih lama dan memori lebih besar.', variant: 'warning' });
    file = f;
    $('#fi-name').textContent = f.name;
    $('#fi-meta').textContent = `${formatBytes(f.size)}`;
    $('#fi-detect').textContent = 'Memeriksa file...';
    info.hidden = false; opts.hidden = false; btnConvert.disabled = true;
    try {
      const fd = new FormData(); fd.append('file', f); fd.append('kind', mode);
      const d = await request('/api/convert/inspect', { form: fd, slowMs: 5000 });
      if (mode === 'csv-to-xlsx') {
        $('#fi-detect').textContent = `Terdeteksi: CSV, delimiter ${d.info.delimiter_name}, encoding ${d.info.encoding}`;
      } else {
        const sel = $('#o-sheet'); sel.replaceChildren();
        d.info.sheets.forEach((s) => sel.append(el('option', { value: s.name, text: s.name })));
        $('#fi-detect').textContent = `Terdeteksi: XLSX dengan ${d.info.sheets.length} sheet`;
      }
      btnConvert.disabled = false;
    } catch (e) {
      info.hidden = true; opts.hidden = true; file = null;
      showError(e.message);
    }
  }

  async function convert() {
    if (!file) return;
    showError(''); clearResult();
    const fd = new FormData(); fd.append('file', file);
    if (mode === 'csv-to-xlsx') {
      fd.append('delimiter', $('#o-delim').value); fd.append('encoding', $('#o-enc').value); fd.append('type_mode', $('#o-type').value);
    } else {
      fd.append('sheet', $('#o-sheet').value); fd.append('delimiter', $('#o-delim2').value); fd.append('encoding', $('#o-enc2').value);
    }
    setLoading(btnConvert, true, 'Mengonversi...'); progress.hidden = false;
    try {
      const { blob, headers } = await request(`/api/convert/${mode}`, { form: fd, as: 'blob', slowMs: 5000, slowMsg: 'Konversi masih berjalan, mohon tunggu.' });
      let notes = [];
      try { notes = JSON.parse(decodeURIComponent(headers.get('X-Notes') || '[]')); } catch (_) { /* abaikan */ }
      const stem = file.name.replace(/\.[^.]+$/, '') || 'converted';
      downloadUrl = URL.createObjectURL(blob);
      const a = $('#btn-download');
      a.href = downloadUrl; a.download = stem + MODES[mode].outExt;
      $('#result-notes').replaceChildren(...notes.map((n) => el('div', { text: n })));
      result.hidden = false;
      toast({ title: 'Konversi berhasil', desc: a.download, variant: 'success' });
    } catch (e) {
      showError(e.message);
      toast({ title: 'Konversi gagal', desc: e.message, variant: 'error' });
    } finally {
      setLoading(btnConvert, false); progress.hidden = true;
    }
  }

  bindDropzone(dz, input, pick);
  $$('.seg').forEach((b) => b.addEventListener('click', () => setMode(b.dataset.mode)));
  $('#btn-remove').addEventListener('click', resetFile);
  btnConvert.addEventListener('click', convert);
  setMode('csv-to-xlsx');
})();
