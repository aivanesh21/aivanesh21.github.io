/* GSTR-1 Converter — page controller. All processing happens in a Web Worker in the browser. */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var root = $('g1');
  if (!root) return;

  var STATES = [
    ['01', 'Jammu & Kashmir'], ['02', 'Himachal Pradesh'], ['03', 'Punjab'], ['04', 'Chandigarh'], ['05', 'Uttarakhand'],
    ['06', 'Haryana'], ['07', 'Delhi'], ['08', 'Rajasthan'], ['09', 'Uttar Pradesh'], ['10', 'Bihar'], ['11', 'Sikkim'],
    ['12', 'Arunachal Pradesh'], ['13', 'Nagaland'], ['14', 'Manipur'], ['15', 'Mizoram'], ['16', 'Tripura'], ['17', 'Meghalaya'],
    ['18', 'Assam'], ['19', 'West Bengal'], ['20', 'Jharkhand'], ['21', 'Odisha'], ['22', 'Chhattisgarh'], ['23', 'Madhya Pradesh'],
    ['24', 'Gujarat'], ['26', 'Dadra & Nagar Haveli & Daman & Diu'], ['27', 'Maharashtra'], ['29', 'Karnataka'], ['30', 'Goa'],
    ['31', 'Lakshdweep'], ['32', 'Kerala'], ['33', 'Tamil Nadu'], ['34', 'Puducherry'], ['35', 'Andaman & Nicobar Islands'],
    ['36', 'Telangana'], ['37', 'Andhra Pradesh'], ['38', 'Ladakh'], ['97', 'Other Territory']
  ];
  var STATE_NAME = {}; STATES.forEach(function (s) { STATE_NAME[s[0]] = s[1]; }); STATE_NAME['96'] = 'Foreign Country';
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  var fmt = new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  var fmtInt = new Intl.NumberFormat('en-IN');
  function money(n) { return n == null ? '—' : fmt.format(n); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  /* ---------- worker plumbing (falls back to an inline Blob worker for file:// use) ---------- */
  var worker = null, seq = 0, pending = {};
  function makeWorker() {
    if (window.GSTR1_WORKER_SOURCE) {
      var url = URL.createObjectURL(new Blob([window.GSTR1_WORKER_SOURCE], { type: 'text/javascript' }));
      return new Worker(url);
    }
    return new Worker(root.getAttribute('data-worker'));
  }
  function call(cmd, payload, transfer) {
    if (!worker) {
      worker = makeWorker();
      worker.onmessage = function (e) { var p = pending[e.data.id]; if (p) { delete pending[e.data.id]; e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error)); } };
      worker.onerror = function (e) { Object.keys(pending).forEach(function (k) { pending[k].reject(new Error(e.message || 'Worker failed to start')); delete pending[k]; }); };
    }
    var id = ++seq;
    return new Promise(function (resolve, reject) {
      pending[id] = { resolve: resolve, reject: reject };
      var msg = Object.assign({ id: id, cmd: cmd }, payload || {});
      worker.postMessage(msg, transfer || []);
    });
  }

  /* ---------- elements ---------- */
  var drop = $('g1-drop'), fileInput = $('g1-file'), status = $('g1-status');
  var fGstin = $('g1-gstin'), fState = $('g1-state'), fMonth = $('g1-month'), fYear = $('g1-year');
  var fLimit = $('g1-limit'), fZero = $('g1-zero'), fQty = $('g1-qty'), fMerge = $('g1-merge'), fGaps = $('g1-gaps'), fDesc = $('g1-desc'), fVer = $('g1-ver');
  var out = $('g1-out');
  var fileName = '', loaded = false, lastView = null;

  fState.innerHTML = '<option value="">Select state…</option>' + STATES.map(function (s) { return '<option value="' + s[0] + '">' + s[0] + ' – ' + esc(s[1]) + '</option>'; }).join('');
  fMonth.innerHTML = MONTHS.map(function (m, i) { var v = (i < 9 ? '0' : '') + (i + 1); return '<option value="' + v + '">' + m + '</option>'; }).join('');
  var now = new Date();
  fMonth.value = ((now.getMonth() || 12) < 10 ? '0' : '') + (now.getMonth() || 12);
  fYear.value = now.getMonth() ? now.getFullYear() : now.getFullYear() - 1;

  function setStatus(kind, html) { status.setAttribute('data-state', kind || ''); status.innerHTML = html || ''; }

  /* ---------- GSTIN check ---------- */
  var CH = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinOk(g) {
    if (!/^\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/.test(g)) return false;
    var s = 0; for (var i = 0; i < 14; i++) { var v = CH.indexOf(g[i]) * (i % 2 ? 2 : 1); s += Math.floor(v / 36) + v % 36; }
    return CH[(36 - s % 36) % 36] === g[14];
  }
  function gstinState() {
    var g = fGstin.value.trim().toUpperCase();
    fGstin.value = g;
    var err = $('g1-gstin-err');
    if (!g) { fGstin.removeAttribute('aria-invalid'); err.setAttribute('data-show', 'false'); fState.disabled = false; return true; }
    var ok = gstinOk(g);
    fGstin.setAttribute('aria-invalid', ok ? 'false' : 'true');
    err.setAttribute('data-show', ok ? 'false' : 'true');
    if (ok && STATE_NAME[g.slice(0, 2)]) { fState.value = g.slice(0, 2); fState.disabled = true; }
    else fState.disabled = false;
    return ok;
  }

  function options() {
    return {
      supplierGstin: gstinOk(fGstin.value.trim().toUpperCase()) ? fGstin.value.trim().toUpperCase() : '',
      homeState: fState.value || '27',
      period: fMonth.value + fYear.value,
      b2clLimit: +fLimit.value || 100000,
      zeroRateAsNil: fZero.checked, reduceQtyOnCN: fQty.checked, mergeNoteUnits: fMerge.checked,
      gapsAsCancelled: fGaps.checked, hsnDescInJson: fDesc.checked, jsonVersion: fVer.value.trim() || 'GST3.2.3'
    };
  }

  /* ---------- file loading ---------- */
  function handleFile(file) {
    if (!file) return;
    fileName = file.name;
    loaded = false;
    out.hidden = true;
    setStatus('info', '<span class="g1-spin" aria-hidden="true"></span> Reading <strong>' + esc(file.name) + '</strong> (' + (file.size / 1048576).toFixed(1) + ' MB)…');
    file.arrayBuffer().then(function (buf) {
      return call('load', { buf: buf, name: file.name }, [buf]);
    }).then(function (r) {
      var info = r.info;
      loaded = true;
      $('g1-filemeta').innerHTML = '<strong>' + esc(fileName) + '</strong> · sheet “' + esc(info.sheet) + '” · ' + fmtInt.format(info.lines) + ' lines · header on row ' + info.headerRow;
      drop.classList.add('is-loaded');
      if (!fState.value && info.guessState) fState.value = info.guessState;
      return run(true);
    }).catch(function (err) {
      setStatus('err', '<strong>Could not read the file.</strong> ' + esc(err.message));
    });
  }
  fileInput.addEventListener('change', function () { handleFile(fileInput.files[0]); });
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('is-over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('is-over'); }); });
  drop.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });

  /* ---------- build ---------- */
  var timer = null, firstRun = true;
  function run(first) {
    if (!loaded) return Promise.resolve();
    var opt = options();
    if (first && firstRun) opt.period = ''; // let the engine detect the period from the data
    setStatus('info', '<span class="g1-spin" aria-hidden="true"></span> Segregating documents and preparing summaries…');
    return call('build', { opt: opt }).then(function (r) {
      var v = r.view;
      if (first && firstRun && v.period) {
        firstRun = false;
        fMonth.value = v.period.slice(0, 2); fYear.value = v.period.slice(2);
        return run(false);
      }
      lastView = v;
      render(v);
      var errs = v.issues.filter(function (i) { return i.level === 'error'; }).reduce(function (s, i) { return s + i.count; }, 0);
      setStatus(errs ? 'err' : 'ok', errs
        ? '<strong>Converted with ' + fmtInt.format(errs) + ' error' + (errs > 1 ? 's' : '') + '.</strong> Review the checks below before filing — the offline tool or portal will reject these rows.'
        : '<strong>Converted.</strong> ' + fmtInt.format(v.stats.docs) + ' documents segregated. Review the summary and checks, then download.');
      out.hidden = false;
    }).catch(function (err) { setStatus('err', '<strong>Conversion failed.</strong> ' + esc(err.message)); });
  }
  function schedule() { if (!loaded) return; clearTimeout(timer); timer = setTimeout(function () { run(false); }, 250); }
  [fState, fMonth, fYear, fLimit, fZero, fQty, fMerge, fGaps, fDesc, fVer].forEach(function (el) { el.addEventListener('change', schedule); });
  fGstin.addEventListener('input', function () { if (fGstin.value.trim().length >= 15 || !fGstin.value.trim()) { gstinState(); schedule(); } });
  fGstin.addEventListener('blur', function () { gstinState(); schedule(); });

  /* ---------- render ---------- */
  function amtCells(a, neg) {
    return ['txval', 'iamt', 'camt', 'samt', 'csamt'].map(function (k) {
      var v = a[k];
      return '<td class="num' + (v < 0 ? ' g1-neg' : '') + '">' + money(v) + '</td>';
    }).join('');
  }
  function render(v) {
    var S = v.summary;
    var tax = S.net.iamt + S.net.camt + S.net.samt;
    $('g1-period-label').textContent = MONTHS[+v.period.slice(0, 2) - 1] + ' ' + v.period.slice(2);
    $('g1-k-docs').textContent = fmtInt.format(v.stats.docs);
    $('g1-k-docs-sub').textContent = fmtInt.format(v.stats.byKind.INV) + ' invoices · ' + fmtInt.format(v.stats.byKind.CN) + ' credit · ' + fmtInt.format(v.stats.byKind.DN) + ' debit notes';
    $('g1-k-tx').textContent = '₹' + money(S.net.txval);
    $('g1-k-tax').textContent = '₹' + money(tax);
    $('g1-k-tax-sub').textContent = 'IGST ' + money(S.net.iamt) + ' · CGST ' + money(S.net.camt) + ' · SGST ' + money(S.net.samt);
    var errs = 0, warns = 0;
    v.issues.forEach(function (i) { if (i.level === 'error') errs += i.count; else if (i.level === 'warn') warns += i.count; });
    $('g1-k-chk').textContent = errs ? fmtInt.format(errs) + ' errors' : 'No errors';
    $('g1-k-chk').className = 'g1-kpi__v ' + (errs ? 'g1-bad' : 'g1-good');
    $('g1-k-chk-sub').textContent = fmtInt.format(warns) + ' warnings to review';

    // sections
    var rows = S.rows.map(function (r) {
      var docs = r.docs == null ? '—' : fmtInt.format(r.docs) + (r.rowsLabel ? ' rows' : '');
      var empty = !r.docs && !r.amt.txval;
      return '<tr' + (empty ? ' class="g1-dim"' : '') + '><th scope="row">' + esc(r.label) + '<span class="g1-sheet">' + esc(r.sheet) + '</span></th><td class="num">' + docs + '</td>' + amtCells(r.amt) + '</tr>';
    }).join('');
    rows += '<tr class="g1-total"><th scope="row">Net outward supplies</th><td></td>' + amtCells(S.net) + '</tr>';
    $('g1-sections').innerHTML = rows;

    // reconciliation
    $('g1-recon').innerHTML = S.recon.map(function (r) {
      var ok = Math.abs(r.diff) <= 1;
      return '<tr><th scope="row">' + esc(r.label) + '</th><td class="num">' + money(r.expected) + '</td><td class="num">' + money(r.hsn) + '</td><td class="num ' + (ok ? 'g1-good' : 'g1-bad') + '">' + (ok ? '✓ ' : '') + money(r.diff) + '</td></tr>';
    }).join('');

    renderHsn(v.hsn, currentHsnTab);
    $('g1-hsn-b2b-n').textContent = v.hsn.b2b.length;
    $('g1-hsn-b2c-n').textContent = v.hsn.b2c.length;

    // docs
    $('g1-docs').innerHTML = v.docs.map(function (d) {
      return '<tr><td>' + esc(d.natureLabel) + '</td><td class="g1-mono">' + esc(d.from) + '</td><td class="g1-mono">' + esc(d.to) + '</td><td class="num">' + fmtInt.format(d.total) + '</td><td class="num">' + fmtInt.format(d.cancel) + '</td><td class="num">' + fmtInt.format(d.net) + '</td><td class="num g1-muted">' + (d.gaps ? fmtInt.format(d.gaps) : '—') + '</td></tr>';
    }).join('') + '<tr class="g1-total"><td>Total</td><td></td><td></td><td class="num">' + fmtInt.format(S.docsTotal) + '</td><td class="num">' + fmtInt.format(S.docsCancelled) + '</td><td class="num">' + fmtInt.format(S.docsTotal - S.docsCancelled) + '</td><td></td></tr>';

    // issues
    var lv = { error: 'Error', warn: 'Check', info: 'Note' };
    $('g1-issues').innerHTML = v.issues.length ? v.issues.map(function (i) {
      return '<li class="g1-issue g1-issue--' + i.level + '"><details><summary><span class="g1-badge">' + lv[i.level] + '</span><span class="g1-issue__msg">' + esc(i.msg) + '</span><span class="g1-issue__n">' + fmtInt.format(i.count) + '</span></summary>' +
        (i.refs.length ? '<p class="g1-refs">' + i.refs.map(esc).join(' · ') + (i.count > i.refs.length ? ' …' : '') + '</p>' : '') + '</details></li>';
    }).join('') : '<li class="g1-issue g1-issue--ok"><span class="g1-badge">OK</span> No issues found.</li>';
  }

  var currentHsnTab = 'b2b';
  function renderHsn(h, tab) {
    currentHsnTab = tab;
    var list = h[tab];
    $('g1-hsn').innerHTML = list.map(function (r) {
      return '<tr><td class="g1-mono">' + esc(r.hsn) + '</td><td class="g1-desc">' + esc(r.desc) + '</td><td>' + esc(r.uqc ? r.uqc.split('-')[0] : 'NA') + '</td><td class="num' + (r.qty < 0 ? ' g1-neg' : '') + '">' + (r.uqc ? fmtInt.format(r.qty) : '—') + '</td><td class="num">' + r.rate + '%</td>' + amtCells(r) + '</tr>';
    }).join('') || '<tr><td colspan="10" class="g1-muted">No rows.</td></tr>';
    Array.prototype.forEach.call(document.querySelectorAll('[data-hsn-tab]'), function (b) { b.setAttribute('aria-selected', b.getAttribute('data-hsn-tab') === tab ? 'true' : 'false'); });
  }
  Array.prototype.forEach.call(document.querySelectorAll('[data-hsn-tab]'), function (b) {
    b.addEventListener('click', function () { if (lastView) renderHsn(lastView.hsn, b.getAttribute('data-hsn-tab')); });
  });

  /* ---------- downloads ---------- */
  function save(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function baseName() {
    var o = options();
    return 'GSTR1_' + (o.supplierGstin || ('State' + o.homeState)) + '_' + o.period;
  }
  function busy(btn, on) { btn.disabled = on; btn.setAttribute('aria-busy', on ? 'true' : 'false'); }

  $('g1-dl-xlsx').addEventListener('click', function () {
    var b = this; busy(b, true);
    call('excel').then(function (r) {
      save(new Blob([r.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '.xlsx');
    }).catch(function (e) { setStatus('err', esc(e.message)); }).then(function () { busy(b, false); });
  });
  $('g1-dl-json').addEventListener('click', function () {
    var b = this;
    var g = fGstin.value.trim().toUpperCase();
    if (!gstinOk(g)) {
      gstinState();
      $('g1-gstin-err').setAttribute('data-show', 'true');
      fGstin.setAttribute('aria-invalid', 'true');
      fGstin.focus();
      setStatus('err', '<strong>Supplier GSTIN required.</strong> The portal JSON must carry your own GSTIN — enter a valid one in step 2.');
      return;
    }
    busy(b, true);
    call('json', { maxBytes: 4800000 }).then(function (r) {
      if (r.parts.length === 1) save(new Blob([r.parts[0]], { type: 'application/json' }), baseName() + '.json');
      else r.parts.forEach(function (p, i) { setTimeout(function () { save(new Blob([p], { type: 'application/json' }), baseName() + '_part' + (i + 1) + 'of' + r.parts.length + '.json'); }, i * 400); });
      $('g1-json-note').textContent = r.parts.length > 1
        ? 'The return exceeded the 5 MB upload limit and was split into ' + r.parts.length + ' files — upload each one in turn on the portal.'
        : 'JSON size: ' + (r.parts[0].length / 1048576).toFixed(2) + ' MB (portal limit 5 MB).';
    }).catch(function (e) { setStatus('err', esc(e.message)); }).then(function () { busy(b, false); });
  });
  $('g1-dl-csv').addEventListener('click', function () {
    call('csv').then(function (r) { save(new Blob(['﻿' + r.text], { type: 'text/csv' }), baseName() + '_checks.csv'); });
  });
})();
