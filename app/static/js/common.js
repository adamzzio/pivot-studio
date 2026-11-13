/* Utilitas UI bersama: DOM helper, toast, dialog, request, sidebar. */
(function () {
  'use strict';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  const ICONS = {
    x: '<path d="M18 6 6 18M6 6l12 12"/>',
    up: '<path d="m18 15-6-6-6 6"/>',
    down: '<path d="m6 9 6 6 6-6"/>',
    trash: '<path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="m19 6-1 14H6L5 6"/>',
    check: '<path d="M20 6 9 17l-5-5"/>',
    alert: '<circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/>',
    info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
    grip: '<circle cx="9" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="15" cy="18" r="1"/>',
    file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>',
    sparkles: '<path d="m12 3 1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9z"/>',
    filter: '<path d="M22 3H2l8 9.46V19l4 2v-8.54z"/>',
  };
  function icon(name) {
    return `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ''}</svg>`;
  }

  const PROPS = new Set(['value', 'checked', 'disabled', 'selected', 'hidden', 'draggable', 'tabIndex']);
  function el(tag, attrs, ...kids) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') n.className = v;
      else if (k === 'text') n.textContent = v;
      else if (k === 'html') n.innerHTML = v; // hanya untuk markup statis internal (ikon)
      else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
      else if (PROPS.has(k)) n[k] = v;
      else n.setAttribute(k, v === true ? '' : v);
    }
    const add = (c) => {
      if (c === null || c === undefined || c === false) return;
      if (Array.isArray(c)) c.forEach(add);
      else n.append(c instanceof Node ? c : document.createTextNode(String(c)));
    };
    kids.forEach(add);
    return n;
  }

  function toast({ title, desc, variant = 'info', duration = 4500, persist = false, spinner = false }) {
    const box = $('#toasts');
    const ic = spinner ? '<span class="spinner"></span>' : icon(variant === 'success' ? 'check' : variant === 'error' || variant === 'warning' ? 'alert' : 'info');
    let timer = null;
    const node = el('div', { class: `toast toast-${variant}`, role: variant === 'error' ? 'alert' : 'status' },
      el('span', { html: ic }),
      el('div', { class: 'toast-body' },
        title ? el('div', { class: 'toast-title', text: title }) : null,
        desc ? el('div', { class: 'toast-desc', text: desc }) : null),
      el('button', { class: 'btn btn-ghost btn-icon btn-sm', type: 'button', 'aria-label': 'Tutup notifikasi', html: icon('x'), onclick: () => dismiss() }));
    function dismiss() { clearTimeout(timer); node.remove(); }
    box.append(node);
    if (!persist) timer = setTimeout(dismiss, duration);
    return { dismiss };
  }

  class ApiError extends Error { constructor(m, s) { super(m); this.status = s; } }

  async function request(url, { method = 'POST', form, json, signal, slowMs = 5000, slowMsg = 'Proses masih berjalan, mohon tunggu sebentar.', as = 'json' } = {}) {
    const opts = { method, signal, headers: {} };
    if (json !== undefined) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(json); }
    else if (form) opts.body = form;
    let slow = null;
    const timer = setTimeout(() => { slow = toast({ title: 'Masih diproses', desc: slowMsg, variant: 'info', persist: true, spinner: true }); }, slowMs);
    try {
      const res = await fetch(url, opts);
      if (!res.ok) {
        let msg = `Terjadi kesalahan (${res.status}).`;
        try { const d = await res.json(); if (d && d.error) msg = d.error; } catch (_) { /* abaikan */ }
        throw new ApiError(msg, res.status);
      }
      if (as === 'blob') return { blob: await res.blob(), headers: res.headers };
      return await res.json();
    } catch (e) {
      if (e.name === 'AbortError' || e instanceof ApiError) throw e;
      throw new ApiError('Tidak dapat terhubung ke server. Pastikan aplikasi masih berjalan.', 0);
    } finally {
      clearTimeout(timer);
      if (slow) slow.dismiss();
    }
  }

  function openDialog({ title, description, body, confirmText = 'Simpan', cancelText = 'Batal', danger = false, onConfirm }) {
    return new Promise((resolve) => {
      let result = false;
      const dlg = el('dialog', { class: 'dialog', 'aria-labelledby': 'dlg-title' });
      const ok = el('button', { class: `btn ${danger ? 'btn-destructive' : 'btn-primary'}`, type: 'button', text: confirmText });
      const cancel = el('button', { class: 'btn btn-outline', type: 'button', text: cancelText, onclick: () => dlg.close() });
      ok.addEventListener('click', async () => {
        const r = onConfirm ? await onConfirm() : true;
        if (r !== false) { result = true; dlg.close(); }
      });
      dlg.append(
        el('div', { class: 'dlg-head' }, el('h2', { id: 'dlg-title', text: title }), description ? el('p', { class: 'muted', text: description }) : null),
        el('div', { class: 'dlg-body' }, body || null),
        el('div', { class: 'dlg-foot' }, cancel, ok));
      dlg.addEventListener('close', () => { dlg.remove(); resolve(result); });
      dlg.addEventListener('click', (e) => { if (e.target === dlg) dlg.close(); });
      document.body.append(dlg);
      dlg.showModal();
    });
  }

  function debounce(fn, ms) {
    let t;
    return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  }
  function formatBytes(n) {
    if (n < 1024) return `${n} B`;
    if (n < 1048576) return `${(n / 1024).toFixed(1)} KB`;
    if (n < 1073741824) return `${(n / 1048576).toFixed(1)} MB`;
    return `${(n / 1073741824).toFixed(2)} GB`;
  }
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }
  function setLoading(btn, on, label) {
    if (on) { btn.dataset.label = btn.textContent; btn.classList.add('loading'); btn.disabled = true; btn.innerHTML = `<span class="spinner"></span> ${label || 'Memproses...'}`; }
    else { btn.classList.remove('loading'); btn.disabled = false; btn.textContent = btn.dataset.label || btn.textContent; }
  }

  // Dropzone generik: klik, keyboard, dan drag-and-drop.
  function bindDropzone(zone, input, onFiles) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    ['dragenter', 'dragover'].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove('over'); }));
    zone.addEventListener('drop', (e) => { if (e.dataTransfer && e.dataTransfer.files.length) onFiles(Array.from(e.dataTransfer.files)); });
    input.addEventListener('change', () => { if (input.files.length) onFiles(Array.from(input.files)); input.value = ''; });
  }

  // Sidebar: ciutkan di desktop, drawer di mobile.
  function initSidebar() {
    const app = $('#app'), btn = $('#menu-btn'), back = $('#backdrop');
    if (!app || !btn) return;
    const mobile = () => window.matchMedia('(max-width: 900px)').matches;
    try { if (localStorage.getItem('sidebar-collapsed') === '1') app.classList.add('collapsed'); } catch (_) { /* abaikan */ }
    const sync = () => btn.setAttribute('aria-expanded', String(mobile() ? app.classList.contains('drawer-open') : !app.classList.contains('collapsed')));
    const closeDrawer = () => { app.classList.remove('drawer-open'); back.hidden = true; sync(); };
    btn.addEventListener('click', () => {
      if (mobile()) { const o = app.classList.toggle('drawer-open'); back.hidden = !o; }
      else { const c = app.classList.toggle('collapsed'); try { localStorage.setItem('sidebar-collapsed', c ? '1' : '0'); } catch (_) { /* abaikan */ } }
      sync();
    });
    back.addEventListener('click', closeDrawer);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && app.classList.contains('drawer-open')) closeDrawer(); });
    window.addEventListener('resize', () => { if (!mobile()) closeDrawer(); });
    sync();
  }
  initSidebar();

  window.UI = { $, $$, el, icon, toast, request, ApiError, openDialog, debounce, formatBytes, downloadBlob, setLoading, bindDropzone };
})();
