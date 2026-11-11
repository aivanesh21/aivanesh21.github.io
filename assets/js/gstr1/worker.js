/* GSTR-1 Converter — Web Worker wrapper around engine.js.
   Keeps the parsed source and the last result in memory so settings can be re-applied quickly. */
/* global GSTR1, importScripts */
if (typeof GSTR1 === 'undefined') importScripts('../vendor/fflate-0.8.2.min.js', 'engine.js?v=4');

var src = null, res = null;

function view(r) {
  return {
    period: r.period, detectedPeriod: r.detectedPeriod, home: r.home, source: r.source,
    stats: r.stats, summary: r.summary, issues: r.issues,
    hsn: r.hsn, docs: r.docs, b2cs: r.b2cs, nil: r.nil,
    b2cl: r.b2cl.map(function (x) { return { no: x.d.no, pos: x.d.pos, val: x.val }; }),
    exp: r.exp.map(function (x) { return { no: x.d.no, typ: x.d.expType, val: x.val }; })
  };
}

self.onmessage = function (e) {
  var m = e.data, id = m.id;
  try {
    if (m.cmd === 'load') {
      src = GSTR1.loadSource(new Uint8Array(m.buf), m.name);
      self.postMessage({ id: id, ok: true, info: { sheet: src.sheet, headerRow: src.headerRow, lines: src.lines.length,
        columns: src.columns, guessState: GSTR1.guessStateFromName(m.name) } });
    } else if (m.cmd === 'build') {
      res = GSTR1.build(src, m.opt);
      self.postMessage({ id: id, ok: true, view: view(res) });
    } else if (m.cmd === 'excel') {
      var bytes = GSTR1.buildExcel(res);
      self.postMessage({ id: id, ok: true, bytes: bytes }, [bytes.buffer]);
    } else if (m.cmd === 'json') {
      var parts = GSTR1.splitJson(GSTR1.buildJson(res), m.maxBytes).map(function (p) { return JSON.stringify(p); });
      self.postMessage({ id: id, ok: true, parts: parts });
    } else if (m.cmd === 'errreport') {
      self.postMessage({ id: id, ok: true, report: GSTR1.parseErrorReport(new Uint8Array(m.buf), m.name) });
    } else if (m.cmd === 'hsninfo') {
      var want = {}, info = {};
      (m.codes || []).forEach(function (c) { want[c] = 1; info[c] = { code: c, lines: 0, value: 0, reg: 0, unreg: 0, descs: {} }; });
      (src ? src.lines : []).forEach(function (L) {
        var h = String(L.hsn == null ? '' : L.hsn).replace(/\.0+$/, '').replace(/\s/g, '');
        if (!want[h]) return;
        var o = info[h], tx = +L.txval || 0, ds = String(L.desc || '').trim();
        o.lines++; o.value += Math.abs(tx); if (L.gstin) o.reg++; else o.unreg++;
        if (ds) o.descs[ds] = (o.descs[ds] || 0) + 1;
      });
      self.postMessage({ id: id, ok: true, info: Object.keys(info).map(function (k) {
        var o = info[k];
        o.descs = Object.keys(o.descs).sort(function (x, y) { return o.descs[y] - o.descs[x]; }).slice(0, 3);
        return o;
      }) });
    } else if (m.cmd === 'csv') {
      self.postMessage({ id: id, ok: true, text: GSTR1.issuesCsv(res) });
    }
  } catch (err) {
    self.postMessage({ id: id, ok: false, error: (err && err.message) || String(err) });
  }
};
