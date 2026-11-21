/* GSTR-2B matcher (gstr2b.niveshsahayak.in): the user opens the GSTR-2B JSON or Excel downloaded from the GST portal; it stays in
   the browser and is compared with the purchase bills saved in the data home (NSBiz purchase bills and returns).
   Result: matched, amount or date differs, number differs (same supplier and amount), in 2B but not in your books, in your books
   but not in 2B, ITC not available in 2B, and suppliers with nothing in 2B. The 2B file is read by field names, not fixed positions,
   so small format changes on the portal do not break it. Purchase register download in the layout of the portal's Matching Offline
   Tool v2.9 template (PurchaseRegister_Template_v2.4: GSTIN C1, year E1, period E2, headers row 5, data from row 6). Screen words: NS.t('g2.*'). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, t = NS.t, money = NS.money, B = window.NSBooks;
  var root = $('#g2');
  if (!root) return;
  var SHEETJS = ['/assets/vendor/xlsx-0.20.3.full.min.js', ''];   /* reads the portal's 2B Excel (ExcelJS 4.4.0 has unfixed 2026 flaws when reading files) */
  var XLSX_JS = ['https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js', 'sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz'];
  var MONF = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function key(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function nkey(no) { return String(no || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, ''); }
  function g15(g) { return String(g || '').trim().toUpperCase(); }
  /* dates in 2B: dd-mm-yyyy, dd/mm/yyyy, dd-Mon-yyyy or Excel dates */
  var MS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function iso(v) {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    var s = String(v || '').trim(), m = s.match(/^(\d{1,2})[-\/.](\d{1,2}|[A-Za-z]{3})[-\/.](\d{2,4})$/);
    if (m) { var mo = isNaN(+m[2]) ? MS[m[2].toLowerCase()] : +m[2], y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + m[1]).slice(-2); }
    return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
  }
  function show(isoD) { return isoD ? NS.fmtDate(isoD) : ''; }

  var DOCS = [], BIZ = [], TWOB = null, R = null;

  /* ---------- reading GSTR-2B ---------- */
  function sumItems(x, f) { return (x.items || x.itms || []).reduce(function (a, i) { var d = i.itm_det || i; return a + num(d[f]); }, 0); }
  function money5(x) {
    var o = {};
    ['txval', 'igst', 'cgst', 'sgst', 'cess'].forEach(function (f) { o[f] = x[f] != null ? num(x[f]) : r2(sumItems(x, f) || sumItems(x, { txval: 'txval', igst: 'iamt', cgst: 'camt', sgst: 'samt', cess: 'csamt' }[f])); });
    return o;
  }
  function fromJson(j) {
    var out = { gstin: '', period: '', rows: [] };
    (function walk(o, path) {
      if (!o || typeof o !== 'object') return;
      if (Array.isArray(o)) { o.forEach(function (x) { walk(x, path); }); return; }
      if (o.gstin && !out.gstin && typeof o.gstin === 'string') out.gstin = o.gstin;
      if (o.rtnprd && !out.period) out.period = String(o.rtnprd);
      if (o.ctin && (o.inv || o.nt)) {
        var sec = /b2ba|cdnra/i.test(path) ? 'amend' : /isd/i.test(path) ? 'isd' : '';
        (o.inv || []).forEach(function (v) { out.rows.push(row(o, v, 'inv', v.inum, v.typ, sec)); });
        (o.nt || []).forEach(function (v) { out.rows.push(row(o, v, v.typ === 'D' || v.ntty === 'D' ? 'dn' : 'cn', v.ntnum || v.nt_num, v.suptyp || v.typ, sec)); });
        return;
      }
      Object.keys(o).forEach(function (k) { if (k !== 'itcsumm' && k !== 'cpsumm') walk(o[k], path + '/' + k); });
    })(j, '');
    return out;
  }
  function row(s, v, kind, no, typ, sec) {
    var m = money5(v);
    return { gstin: g15(s.ctin), name: s.trdnm || s.lglnm || '', kind: kind, no: String(no || ''), date: iso(v.dt || v.idt || v.ntdt), value: num(v.val), txv: m.txval, igst: m.igst, cgst: m.cgst, sgst: m.sgst, cess: m.cess,
      itc: String(v.itcavl || 'Y').toUpperCase() !== 'N', rsn: v.rsn || '', rev: String(v.rev || 'N').toUpperCase() === 'Y', filed: s.supfildt || '', sprd: s.supprd || '', sec: sec };
  }
  var COLS = [['gstin', /gstin of (the )?supplier/i], ['name', /trade.*legal|legal.*name|trade name/i], ['no', /(invoice|note|document) (number|no)/i], ['ntype', /note type/i],
    ['date', /(invoice|note|document) date/i], ['value', /(invoice|note|document) value/i], ['txv', /taxable value/i], ['igst', /integrated/i], ['cgst', /central/i], ['sgst', /state/i], ['cess', /^cess/i],
    ['itc', /itc avail/i], ['rsn', /reason/i], ['rev', /reverse charge/i]];
  /* an Excel date number (days since 30 Dec 1899) as yyyy-mm-dd, worked out in UTC so no time zone moves the day */
  function xlDate(v) { return typeof v === 'number' && v > 20000 && v < 80000 ? new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10) : v; }
  function fromSheets(sheets) {   /* sheets: [{ name, rows: [[cell, ...], ...] }] (rows from row 1, cells from column 1) */
    var out = { gstin: '', period: '', rows: [] };
    sheets.forEach(function (sh) {
      var name = String(sh.name).toUpperCase().replace(/\s/g, ''), rows = sh.rows, txt = function (c) { return String(c == null ? '' : c).replace(/\s+/g, ' '); };
      var sec = /^B2B$/.test(name) ? 'inv' : /^B2B-CDNR$|^CDNR$/.test(name) ? 'note' : /^B2BA$|CDNRA/.test(name) ? 'amend' : '';
      if (!out.gstin) rows.slice(0, 11).forEach(function (r) { (r || []).forEach(function (c) { var m = String(c || '').match(/\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]\b/); if (m && !out.gstin) out.gstin = m[0]; }); });
      if (!sec) return;
      var hr = -1, map = {};
      for (var i = 0; i < Math.min(12, rows.length) && hr < 0; i++) if ((rows[i] || []).some(function (c) { return /gstin of (the )?supplier/i.test(txt(c)); })) hr = i;
      if (hr < 0) return;
      [hr, hr + 1].forEach(function (rn) {
        (rows[rn] || []).forEach(function (c, col) { if (c === '' || c == null) return; COLS.forEach(function (k) { if (map[k[0]] == null && k[1].test(txt(c))) map[k[0]] = col; }); });
      });
      if (map.gstin == null || map.no == null) return;
      rows.slice(hr + 2).forEach(function (r) {
        r = r || [];
        var v = function (k) { var c = map[k] != null ? r[map[k]] : ''; return c == null ? '' : c; };
        var g = g15(v('gstin'));
        if (!/^\d{2}[A-Z0-9]{13}$/.test(g)) return;
        var nt = String(v('ntype') || '').toLowerCase();
        out.rows.push({ gstin: g, name: String(v('name') || ''), kind: sec === 'note' || nt ? (/debit/.test(nt) ? 'dn' : 'cn') : 'inv', no: String(v('no') || ''), date: iso(xlDate(v('date'))), value: num(v('value')),
          txv: num(v('txv')), igst: num(v('igst')), cgst: num(v('cgst')), sgst: num(v('sgst')), cess: num(v('cess')), itc: !/^n/i.test(String(v('itc') || 'Yes')), rsn: String(v('rsn') || ''),
          rev: /^y/i.test(String(v('rev') || 'No')), filed: '', sprd: '', sec: sec === 'amend' ? 'amend' : '' });
      });
    });
    return out;
  }
  function fromExcel(buf) {
    return NS.loadJs(SHEETJS[0], SHEETJS[1]).then(function () {
      var X = window.XLSX, wb = X.read(new Uint8Array(buf), { type: 'array' });
      return fromSheets(wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) }; }));
    });
  }

  /* ---------- our purchase bills ---------- */
  function books() {
    var bz = $('#g2-biz').value;
    return DOCS.filter(function (d) { return (d.type === 'pbill' || d.type === 'note' && d.side === 'purchase') && key(d.seller && d.seller.name) === bz && !d.cancelled; }).map(function (d) {
      var pn = d.type === 'note', back = pn ? d.kind === 'debit' : d.kind === 'return', tt = B.itemTotals(d), s = back ? -1 : 1;
      return { id: d.id, gstin: g15(d.buyer && d.buyer.gstin), name: d.buyer && d.buyer.name || '', kind: back ? 'cn' : pn ? 'dn' : 'inv', no: d.no || '', date: d.date || '', value: s * tt.total,
        txv: s * tt.taxable, igst: s * tt.igst, cgst: s * tt.cgst, sgst: s * tt.sgst, cess: 0, noGst: pn ? d.withGst === false : !!d.noGst };
    }).filter(function (b) { return !b.noGst; });
  }
  function tax(x) { return r2(Math.abs(x.igst) + Math.abs(x.cgst) + Math.abs(x.sgst) + Math.abs(x.cess)); }
  function stax(x) { return x.kind === 'cn' ? -tax(x) : tax(x); }   /* a supplier credit note takes ITC away (F3-7) */
  function inRange(d) { var f = $('#g2-from').value, to = $('#g2-to').value; return (!f || d >= f) && (!to || d <= to); }

  function match() {
    var tol = num($('#g2-tol').value), mine = books(), two = TWOB ? TWOB.rows.filter(function (x) { return x.sec !== 'isd'; }) : [];
    var used = {}, res = { ok: [], diff: [], numdiff: [], only2b: [], onlybooks: [], noitc: [], nofile: [], t2b: 0, tbooks: 0, tok: 0 };
    var idx = {}; mine.forEach(function (b, i) { var k = b.gstin + '|' + b.kind + '|' + nkey(b.no); (idx[k] = idx[k] || []).push(i); });
    two.forEach(function (x) {
      res.t2b += x.itc ? stax(x) : 0;
      if (!x.itc) res.noitc.push(x);
      var k = x.gstin + '|' + x.kind + '|' + nkey(x.no), hit = (idx[k] || []).filter(function (i) { return !used[i]; })[0];
      if (hit == null) {   /* same supplier, same tax, different number */
        hit = mine.map(function (b, i) { return i; }).filter(function (i) { var b = mine[i]; return !used[i] && b.gstin === x.gstin && b.kind === x.kind && Math.abs(tax(b) - tax(x)) <= tol && Math.abs(Math.abs(b.txv) - Math.abs(x.txv)) <= tol; })[0];
        if (hit != null) { used[hit] = 1; res.numdiff.push({ b: mine[hit], x: x }); return; }
        res.only2b.push(x); return;
      }
      used[hit] = 1;
      var b = mine[hit], dt = Math.abs(Math.abs(b.txv) - Math.abs(x.txv)), dx = Math.abs(tax(b) - tax(x));
      if (dt <= tol && dx <= tol && Math.abs(Math.abs(b.igst) - Math.abs(x.igst)) <= tol && (b.date === x.date || !x.date)) { res.ok.push({ b: b, x: x }); res.tok += x.itc ? stax(x) : 0; }
      else res.diff.push({ b: b, x: x, why: [dt > tol || dx > tol ? t('g2.why.amt') : '', Math.abs(Math.abs(b.igst) - Math.abs(x.igst)) > tol && dx <= tol ? t('g2.why.head') : '', b.date !== x.date && x.date ? t('g2.why.date') : ''].filter(Boolean).join(', ') });
    });
    mine.forEach(function (b, i) { if (!used[i] && inRange(b.date)) { res.onlybooks.push(b); } if (inRange(b.date)) res.tbooks += stax(b); });
    /* suppliers in our books for the period with nothing at all in 2B */
    var in2b = {}; two.forEach(function (x) { in2b[x.gstin] = 1; });
    var sup = {};
    res.onlybooks.forEach(function (b) { if (b.gstin && !in2b[b.gstin]) { var s = sup[b.gstin] || (sup[b.gstin] = { gstin: b.gstin, name: b.name, n: 0, tax: 0 }); s.n++; s.tax = r2(s.tax + tax(b)); } });
    res.nofile = Object.keys(sup).map(function (k) { return sup[k]; });
    res.t2b = r2(res.t2b); res.tbooks = r2(res.tbooks); res.tok = r2(res.tok);
    return res;
  }

  /* ---------- screen ---------- */
  function m(n) { return money(n, false); }
  function list(id, items, head, cells, act) {
    return '<section class="g1-sec" data-sec="' + id + '"><div class="g1-sh"><h3>' + esc(t('g2.s.' + id)) + ' <small class="soft">(' + items.length + ')</small></h3>' +
      (items.length ? '<button type="button" class="btn sm" data-csv="' + id + '">CSV</button>' : '') + '</div><p class="soft g2-help">' + esc(t('g2.h.' + id)) + '</p>' +
      (items.length ? '<div class="g1-tw"><table class="g1-t"><thead><tr>' + head.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') + (act ? '<th class="gt-noprint"></th>' : '') + '</tr></thead><tbody>' +
        items.map(function (x, i) { return '<tr>' + cells(x).map(function (c) { return '<td' + (typeof c === 'number' ? ' class="n"' : '') + '>' + esc(typeof c === 'number' ? m(c) : c) + '</td>'; }).join('') + (act ? '<td class="gt-noprint">' + act(x, i) + '</td>' : '') + '</tr>'; }).join('') + '</tbody></table></div>' : '') + '</section>';
  }
  var CSVS = {};
  function draw() {
    if (!TWOB) return;
    R = match();
    var r = R, H = [t('g2.c.sup'), t('g2.c.no'), t('g2.c.date'), t('g2.c.txv'), t('g2.c.tax')];
    var one = function (x) { return [x.name ? x.name + ' · ' + x.gstin : x.gstin, x.no, show(x.date), r2(x.txv), tax(x)]; };
    var pair = function (p) { return [(p.x.name || p.b.name) + ' · ' + p.x.gstin, p.b.no + (nkey(p.b.no) !== nkey(p.x.no) ? ' / 2B: ' + p.x.no : ''), show(p.b.date) + (p.x.date && p.x.date !== p.b.date ? ' / 2B: ' + show(p.x.date) : ''), r2(p.b.txv), r2(p.x.txv), tax(p.b), tax(p.x), p.why || ''].filter(function (c, i) { return i < 7 || p.why !== undefined; }); };
    var PH = [t('g2.c.sup'), t('g2.c.no'), t('g2.c.date'), t('g2.c.txvb'), t('g2.c.txv2'), t('g2.c.taxb'), t('g2.c.tax2')];
    CSVS = {
      ok: [PH].concat(r.ok.map(pair)), diff: [PH.concat([t('g2.c.why')])].concat(r.diff.map(pair)), numdiff: [PH].concat(r.numdiff.map(pair)),
      only2b: [H.concat([t('g2.c.itc')])].concat(r.only2b.map(function (x) { return one(x).concat([x.itc ? 'Y' : 'N']); })), onlybooks: [H].concat(r.onlybooks.map(one)),
      noitc: [H.concat([t('g2.c.rsn')])].concat(r.noitc.map(function (x) { return one(x).concat([x.rsn]); })), nofile: [[t('g2.c.sup'), 'GSTIN', t('g2.c.bills'), t('g2.c.tax')]].concat(r.nofile.map(function (s) { return [s.name, s.gstin, String(s.n), s.tax]; }))
    };
    var h = '<div class="g1-print-hd"><b>' + esc(t('g2.for', { g: TWOB.gstin || '', p: TWOB.period || '' })) + '</b></div>' +
      '<div class="g1-tiles"><div><span>' + esc(t('g2.k.2b')) + '</span><b>' + m(r.t2b) + '</b></div><div><span>' + esc(t('g2.k.books')) + '</span><b>' + m(r.tbooks) + '</b></div>' +
      '<div class="hi"><span>' + esc(t('g2.k.ok')) + '</span><b>' + m(r.tok) + '</b></div><div><span>' + esc(t('g2.k.n')) + '</span><b>' + r.ok.length + ' / ' + (r.ok.length + r.diff.length + r.numdiff.length + r.only2b.length) + '</b></div></div>' +
      (TWOB.gstin && SELLER() && SELLER().gstin && g15(SELLER().gstin) !== g15(TWOB.gstin) ? '<p class="callout">' + esc(t('g2.w.gstin', { a: TWOB.gstin, b: SELLER().gstin })) + '</p>' : '') +
      list('diff', r.diff, PH.concat([t('g2.c.why')]), pair) + list('only2b', r.only2b, H, one, function (x, i) { return x.sec === 'impg' || x.sec === 'impgsez' || x.kind === 'dn' ? '' : '<button type="button" class="btn sm" data-add2b="' + i + '">' + esc(t('uf.add2b')) + '</button>'; }) + list('onlybooks', r.onlybooks, H, one) + list('nofile', r.nofile, [t('g2.c.sup'), 'GSTIN', t('g2.c.bills'), t('g2.c.tax')], function (s) { return [s.name, s.gstin, String(s.n), s.tax]; }) +
      list('numdiff', r.numdiff, PH, pair) + list('noitc', r.noitc, H.concat([t('g2.c.rsn')]), function (x) { return one(x).concat([x.rsn || '—']); }) + list('ok', r.ok, PH, pair);
    $('#g2-out').innerHTML = h;
    var off = $('#g2-off'); if (off) off.innerHTML = official(r);
    $('#g2-files').hidden = false;
    keep(r);
  }
  /* S-6 (ERP round 1, r1-gst): the match is kept for GSTR-3B in data home meta 'gstr2b' '<business key>|<YYYY-MM of the 2B period>' =
     { at, ok: [ids of bills and notes shown in 2B with ITC available], noitc: [ids shown with ITC not available], impg: { igst, cess } of
     import of goods (4(A)(1)) }. GSTR-3B then takes ITC of the bills this 2B shows, whatever their date (rule 36(4), s.16(2)(aa)).
     Kept again whenever the result changes; a bill matched in two months counts in the first. */
  var KEPT = '';
  function keep(r) {
    var p = String(TWOB && TWOB.period || ''), mm = p.match(/^(\d{2})(\d{4})$/), bz = $('#g2-biz').value; if (!mm || !bz) return;
    var ok = [], no = [], im = { igst: 0, cess: 0 };
    r.ok.concat(r.diff, r.numdiff).forEach(function (q) { if (q.b && q.b.id) (q.x.itc ? ok : no).push(q.b.id); });
    TWOB.rows.forEach(function (x) { if ((x.sec === 'impg' || x.sec === 'impgsez') && x.itc) { im.igst = r2(im.igst + num(x.igst)); im.cess = r2(im.cess + num(x.cess)); } });
    var k = bz + '|' + mm[2] + '-' + mm[1], rec = { ok: ok.sort(), noitc: no.sort(), impg: im }, sig = k + JSON.stringify(rec);
    if (sig === KEPT) return; KEPT = sig;
    NSStore.meta('gstr2b').then(function (M) { M = M || {}; M[k] = Object.assign({ at: Date.now() }, rec); return NSStore.meta('gstr2b', M); })
      .then(function () { var st = $('#g2-kept'); if (!st) { st = document.createElement('p'); st.id = 'g2-kept'; st.className = 'soft'; $('#g2-out').insertAdjacentElement('beforebegin', st); } st.textContent = t('g2.kept', { n: ok.length, p: mm[1] + '/' + mm[2] }); })
      .catch(function () {});
  }
  /* a bill found only in GSTR-2B: save it as a purchase bill (unpaid) in the data home, one line at the bill's tax rate */
  var RATES = [0, 0.1, 0.25, 1.5, 3, 5, 12, 18, 28, 40];
  function addBill(x, btn) {
    var s = SELLER(); if (!x || !s || !s.name) { NS.toast(t('g3.nobiz'), 'bad'); return; }
    var txv = Math.abs(x.txv), g = txv ? (Math.abs(x.igst) + Math.abs(x.cgst) + Math.abs(x.sgst)) / txv * 100 : 0;
    var near = RATES.reduce(function (a, r) { return Math.abs(r - g) < Math.abs(a - g) ? r : a; }, 18);
    if (!NS.okSync(t('uf.add2bq', { no: x.no, p: x.name || x.gstin }))) return;
    var sel = JSON.parse(JSON.stringify(s)); delete sel.id; delete sel.created; delete sel.updated;
    var d = { type: 'pbill', kind: x.kind === 'cn' ? 'return' : 'bill', app: 1, from2b: 1, seller: sel, buyer: { name: x.name || x.gstin, gstin: x.gstin, state: x.gstin.slice(0, 2) }, no: x.no, date: x.date || NS.today(), round: false, noGst: false,
      items: [{ desc: 'Purchase as per GSTR-2B', hsn: '', qty: 1, unit: 'Nos', rate: txv, gst: Math.abs(near - g) <= 0.1 ? near : r2(g) }] };
    if (d.kind === 'return') d.orig = { no: '' }; else d.due = '';
    var tt = B.itemTotals(d); d.tax = { taxable: tt.taxable, cgst: tt.cgst, sgst: tt.sgst, igst: tt.igst }; d.total = r2(tt.taxable + tt.cgst + tt.sgst + tt.igst);
    btn.disabled = true;
    NSStore.put('doc', d).then(function (rec) { DOCS.push(rec); NS.toast(t('uf.add2bok', { no: x.no })); draw(); }, function (err) { btn.disabled = false; NS.toast(err.message, 'bad'); });
  }
  /* the statement as Form GSTR-2B lays out its summary (rule 60(7) of the CGST Rules): ITC available (Part A), ITC reversal (Part B),
     ITC not available, each with the GSTR-3B table it goes to; then the match with your purchase bills */
  function official(r) {
    var G = NS.gstForm; if (!G || !TWOB) return '';
    var s = SELLER() || {}, L = TWOB.rows, R = function (a) { return { cells: a }; };
    var sum = function (f) { var o = { igst: 0, cgst: 0, sgst: 0, cess: 0 }; L.filter(f).forEach(function (x) { o.igst = r2(o.igst + num(x.igst)); o.cgst = r2(o.cgst + num(x.cgst)); o.sgst = r2(o.sgst + num(x.sgst)); o.cess = r2(o.cess + num(x.cess)); }); return [o.igst, o.cgst, o.sgst, o.cess]; };
    var imp = function (x) { return x.sec === 'impg' || x.sec === 'impgsez'; }, base = function (x) { return x.itc && !imp(x) && x.sec !== 'isd'; };
    var per = String(TWOB.period || ''), pl = /^\d{6}$/.test(per) ? ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+per.slice(0, 2) - 1] + ' ' + per.slice(2) : per;
    var y = /^\d{6}$/.test(per) ? +per.slice(2) : 0, mo = /^\d{6}$/.test(per) ? +per.slice(0, 2) : 4, fy = y ? (mo >= 4 ? y + '-' + String(y + 1).slice(2) : (y - 1) + '-' + String(y).slice(2)) : '';
    var h = G.head('FORM GSTR-2B', '[See rule 60(7)]', 'Auto-drafted ITC Statement', [['Year', fy], ['Month', pl]], [
      ['1.', 'GSTIN', TWOB.gstin || s.gstin || ''], ['2(a).', 'Legal name of the registered person', (s.gst && s.gst.legal) || s.legal || s.name || ''], ['2(b).', 'Trade name, if any', s.name || ''], ['2(c).', 'Date of generation', '']]);
    var C = ['Heading', 'GSTR-3B table', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess'];
    h += G.table('Part A', 'ITC Available: summary', C, [
      { sub: 'I. All other ITC: Supplies from registered persons' },
      R(['B2B: Invoices', '4(A)(5)'].concat(sum(function (x) { return base(x) && !x.rev && x.kind === 'inv' && x.sec !== 'amend'; }))),
      R(['B2B: Debit notes', '4(A)(5)'].concat(sum(function (x) { return base(x) && !x.rev && x.kind === 'dn' && x.sec !== 'amend'; }))),
      R(['B2B: Invoices and debit notes (Amendment)', '4(A)(5)'].concat(sum(function (x) { return base(x) && !x.rev && x.kind !== 'cn' && x.sec === 'amend'; }))),
      { sub: 'II. Inward Supplies from ISD' }, R(['ISD: Invoices', '4(A)(4)'].concat(sum(function (x) { return x.itc && x.sec === 'isd'; }))),
      { sub: 'III. Inward Supplies liable for reverse charge' }, R(['B2B: Invoices and debit notes', '3.1(d), 4(A)(3)'].concat(sum(function (x) { return base(x) && x.rev && x.kind !== 'cn'; }))),
      { sub: 'IV. Import of Goods' }, R(['IMPG: Import of goods from overseas and from SEZ', '4(A)(1)'].concat(sum(function (x) { return x.itc && imp(x); })))]);
    h += G.table('Part B', 'ITC Reversal: summary', C, [{ sub: 'I. Others' },
      R(['B2B: Credit notes', '4(B)(2)'].concat(sum(function (x) { return base(x) && !x.rev && x.kind === 'cn'; }))),
      R(['B2B: Credit notes (Reverse charge)', '4(B)(2)'].concat(sum(function (x) { return base(x) && x.rev && x.kind === 'cn'; })))]);
    h += G.table('', 'ITC Not Available: summary', C, [R(['Documents on which ITC is not available', '4(D)(2)'].concat(sum(function (x) { return !x.itc; })))]);
    h += G.table('', 'Match with your purchase bills (books)', ['Result', 'Bills', 'Tax'], [
      R(['Matched', String(r.ok.length), r.tok]), R(['Amount differs', String(r.diff.length), '']), R(['Bill number differs', String(r.numdiff.length), '']),
      R(['In GSTR-2B, not in your books', String(r.only2b.length), '']), R(['In your books, not in GSTR-2B', String(r.onlybooks.length), '']), { cells: ['Tax as per GSTR-2B / as per books', '', m(r.t2b) + ' / ' + m(r.tbooks)], tot: 1 }]);
    h += G.note('Summary worked out by NSBiz from the GSTR-2B file you opened (the portal generates the statement itself on the 14th). The details of each bill are under the working.');
    return G.wrap(h);
  }
  (function offBox() {
    var out = $('#g2-out'); if (!out || $('#g2-off')) return;
    var d = document.createElement('section'); d.id = 'g2-off'; d.className = 'off-box'; d.setAttribute('aria-label', 'Form GSTR-2B');
    if (NS.gstForm && NS.gstForm.inApp) { out.parentNode.insertBefore(d, out); NS.gstForm.fold(out, 'See every bill as the working'); } else out.parentNode.insertBefore(d, out.nextSibling);
  })();
  function SELLER() { var bz = $('#g2-biz').value; return BIZ.filter(function (b) { return key(b.name) === bz; })[0] || (DOCS.filter(function (d) { return key(d.seller && d.seller.name) === bz; })[0] || {}).seller; }
  function save(blob, name) { var a = document.createElement('a'), u = URL.createObjectURL(blob); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 4000); }
  function csvCell(v) { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  function csv(id) { save(new Blob(['﻿' + CSVS[id].map(function (r) { return r.map(csvCell).join(','); }).join('\r\n') + '\r\n'], { type: 'text/csv;charset=utf-8' }), 'GSTR-2B match ' + id + '.csv'); }
  /* purchase register for the portal's Matching Offline Tool */
  function register() {
    var mine = books().filter(function (b) { return inRange(b.date); }), s = SELLER() || {}, f = $('#g2-from').value || NS.today();
    var y = +f.slice(0, 4), mo = +f.slice(5, 7), fy = mo >= 4 ? y + '-' + String(y + 1).slice(2) : (y - 1) + '-' + String(y).slice(2);
    NS.toast(t('g1.making'));
    NS.loadJs(XLSX_JS[0], XLSX_JS[1]).then(function () {
      var wb = new window.ExcelJS.Workbook(), ws = wb.addWorksheet('Purchase Register');
      wb.creator = 'Nivesh Sahayak';
      ws.getCell('B1').value = 'GSTIN of recipient* : '; ws.getCell('C1').value = g15(s.gstin); ws.getCell('D1').value = 'Financial year* :'; ws.getCell('E1').value = fy;
      ws.getCell('B2').value = 'Trade/Legal name:'; ws.getCell('C2').value = s.name || ''; ws.getCell('D2').value = 'Tax period* : '; ws.getCell('E2').value = MONF[mo - 1];
      ws.getRow(5).values = ['GSTIN of Supplier/ECO *', 'Trade/Legal name', 'Type of inward supplies *', 'Document type *', 'Document number *', 'Document date *', 'Taxable value (₹) *', 'Integrated tax (₹) ', 'Central tax (₹) ', 'State/ UT tax (₹) ', 'Cess (₹) '];
      ws.getRow(5).font = { bold: true };
      mine.forEach(function (b, i) {
        var p = b.date.split('-');
        ws.getRow(6 + i).values = [b.gstin, b.name, 'B2B', b.kind === 'cn' ? 'Credit Note' : 'Invoice', b.no, p[2] + '/' + p[1] + '/' + p[0], Math.abs(b.txv), Math.abs(b.igst), Math.abs(b.cgst), Math.abs(b.sgst), 0];
      });
      [18, 30, 14, 14, 18, 14, 16, 14, 14, 14, 10].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
      return wb.xlsx.writeBuffer();
    }).then(function (buf) { save(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'Purchase register ' + (s.gstin || '') + ' ' + MONF[mo - 1] + ' ' + y + '.xlsx'); })
      .catch(function () { NS.toast(t('g1.xfail'), 'bad'); });
  }
  /* the department's "GSTR2 Excel Workbook Template" (inside the Returns Offline Tool): the purchase side laid out the way the
     Offline Tool imports it. One row per GST rate of each bill; registered suppliers in b2b, unregistered in b2bur, notes in
     cdnr, HSN summary in hsnsum. ITC eligibility from the bill or line (blocked / ineligible -> Ineligible, SAC 99 -> Input services). */
  /* units as the GST unit codes (same table as the GSTR-1 screen) */
  var UQC = { nos: 'NOS-NUMBERS', pcs: 'PCS-PIECES', kg: 'KGS-KILOGRAMS', gm: 'GMS-GRAMMES', ltr: 'LTR-LITRES', ml: 'MLT-MILILITRE', mtr: 'MTR-METERS', 'sq ft': 'SQF-SQUARE FEET',
    box: 'BOX-BOX', pack: 'PAC-PACKS', set: 'SET-SETS', pair: 'PRS-PAIRS', dozen: 'DOZ-DOZENS', bag: 'BAG-BAGS', ton: 'TON-TONNES', bottle: 'BTL-BOTTLES', roll: 'ROL-ROLLS', 'sq m': 'SQM-SQUARE METERS', quintal: 'QTL-QUINTAL', unit: 'UNT-UNITS' };
  var NREASON = { ret: '01-Sales Return', disc: '02-Post Sale Discount', price: '04-Correction in Invoice', tax: '04-Correction in Invoice', short: '07-Others', other: '07-Others' };
  function g2Rows(docs, me) {
    var r2 = function (n) { return Math.round((+n || 0) * 100) / 100; }, ms = String(me && (me.state || String(me.gstin || '').slice(0, 2)) || '');
    var out = { b2b: [], b2bur: [], cdnr: [], hsn: {} }, LINES = function (d) { return (d.items || []).concat(B.chargeLines ? B.chargeLines(d) : []); };
    var elig = function (d, it) { var v = it && it.itc || typeof d.itc === 'string' && d.itc || 'ok'; return /^(blocked|ineligible)$/.test(v) || d.itc === false ? 'Ineligible' : /^99/.test(String(it && it.hsn || '')) ? 'Input services' : 'Inputs'; };
    docs.forEach(function (d) {
      var sup = d.buyer || {}, gst = g15(sup.gstin), ss = String(sup.state || gst.slice(0, 2) || ms), inter = !!ss && !!ms && ss !== ms;
      var tt = B.itemTotals(d), val = r2(tt.total), rates = {};
      LINES(d).forEach(function (it) {
        var q = num(it.qty), rt = num(it.rate); if (!it.desc && !rt) return;
        var gross = q * rt, txv = r2(gross - gross * num(it.disc) / 100), g = d.noGst ? 0 : num(it.gst), e = elig(d, it), k = g + '|' + e, b = rates[k] || (rates[k] = { g: g, e: e, txv: 0, tax: 0 });
        b.txv = r2(b.txv + txv); b.tax = r2(b.tax + txv * g / 100);
        if (it.hsn) { var h = String(it.hsn).trim(), uq = /^99/.test(h) ? 'NA' : UQC[key(it.unit)] || 'OTH-OTHERS', hk = h + '|' + uq, x = out.hsn[hk] || (out.hsn[hk] = { hsn: h, desc: it.desc || '', uqc: uq, qty: 0, val: 0, txv: 0, igst: 0, cgst: 0, sgst: 0 });
          var hx = r2(txv * g / 100); x.qty = r2(x.qty + q); x.txv = r2(x.txv + txv); x.val = r2(x.val + txv + hx); if (inter) x.igst = r2(x.igst + hx); else { x.cgst = r2(x.cgst + hx / 2); x.sgst = r2(x.sgst + hx / 2); } }
      });
      var pos = ms, note = d.type === 'note', rc = d.rc || d.rcm ? 'Y' : 'N';
      Object.keys(rates).forEach(function (k) {
        var b = rates[k], ig = inter ? b.tax : 0, cg = inter ? 0 : r2(b.tax / 2), sg = inter ? 0 : r2(b.tax - r2(b.tax / 2)), av = b.e === 'Ineligible' ? [0, 0, 0, 0] : [ig, cg, sg, 0];
        var tax = [b.g, b.txv, ig, cg, sg, 0, b.e].concat(av);
        if (note) out.cdnr.push([gst, d.no || '', d.date, d.orig && d.orig.no || '', d.orig && d.orig.date || '', 'N', d.kind === 'debit' ? 'C' : 'D', NREASON[d.reason] || '07-Others', inter ? 'Inter State' : 'Intra State', val].concat(tax));
        else if (gst.length === 15) out.b2b.push([gst, d.no || '', d.date, val, pos, rc, 'Regular'].concat(tax));
        else out.b2bur.push([sup.name || '', d.no || '', d.date, val, pos, inter ? 'Inter State' : 'Intra State'].concat(tax));
      });
    });
    out.hsnsum = Object.keys(out.hsn).map(function (k) { var x = out.hsn[k]; return [x.hsn, x.desc, x.uqc, x.qty, x.val, x.txv, x.igst, x.cgst, x.sgst, 0]; });
    return out;
  }
  function fillG2(book, docs, me) {
    var X = NS.xlFill, need = ['b2b', 'b2bur', 'cdnr', 'hsnsum', 'imps', 'impg', 'cdnur', 'at', 'atadj', 'itcr'];
    if (need.some(function (n) { return book.sheets.indexOf(n) < 0; })) throw new Error('notg2');
    var o = g2Rows(docs, me);
    return book.sheet('master').then(function (M) { return M.listUnder('POS'); }, function () { return []; }).then(function (POS) {
      var fix = function (L, dateCols, posCol) { return L.map(function (r) { return r.map(function (v, j) { return dateCols.indexOf(j) >= 0 ? X.serial(v) : j === posCol ? X.posFrom(POS, v && String(v).length === 2 ? v + '-' : v) : v; }); }); };
      var put = { b2b: fix(o.b2b, [2], 4), b2bur: fix(o.b2bur, [2], 4), cdnr: fix(o.cdnr, [2, 4], -1), hsnsum: o.hsnsum, imps: [], impg: [], cdnur: [], at: [], atadj: [], itcr: [] };
      return Object.keys(put).reduce(function (p, n) { return p.then(function () { return book.sheet(n); }).then(function (S) { S.putRows(5, put[n]); }); }, Promise.resolve()).then(function () { return o; });
    });
  }
  function showTpl() { var e = $('[data-tpl="gstr2"]'); if (e && NS.xlFill) NS.xlFill.info('gstr2').then(function (k) { e.textContent = k ? ' Using ' + k.name : ''; }); }
  function fillG2Tpl(own) {
    if (!NS.xlFill) return;
    var me = SELLER() || {}, bz = $('#g2-biz').value, f0 = $('#g2-from').value, t0 = $('#g2-to').value;
    var docs = DOCS.filter(function (d) { return (d.type === 'pbill' || d.type === 'note' && d.side === 'purchase') && key(d.seller && d.seller.name) === bz && !d.cancelled && !d.deleted && inRange(d.date || ''); });
    if (!docs.length) { NS.toast('No purchase bills in the chosen dates.'); return; }
    var tp = null;
    NS.xlFill.template('gstr2', own).then(function (x) {
      tp = x; if (!tp) return; var f = tp.file;
      NS.toast('Filling ' + f.name + '…');
      return NS.xlFill.open(f).then(function (book) { return fillG2(book, docs, me).then(function (o) { NS.xlFill.good('gstr2', tp);
        var nm = 'GSTR2_' + (g15(me.gstin) || 'GSTIN') + '_' + (f0 || '') + (t0 && t0 !== f0 ? '_' + t0 : '') + ' NSBiz.xlsx';
        return book.save(nm).then(function (bl) { return NS.xlFill.toFolder(g15(me.gstin), 'GSTR-2 ' + (f0 || '') + (t0 && t0 !== f0 ? ' to ' + t0 : ''), nm, bl); }).then(function (inF) { NS.toast('Saved as "' + nm + '" (' + (o.b2b.length + o.b2bur.length + o.cdnr.length) + ' rows).' + (inF ? ' Also saved in NSBiz › GST Returns.' : '')); showTpl(); });
      }); });
    }).catch(function (e) {
      var w = e && e.message; if (/^(notg2|notxl)$/.test(w)) NS.xlFill.bad('gstr2', tp);
      NS.toast(w === 'notzip' ? 'Choose the zip downloaded from gst.gov.in, or the Excel template inside it.' : w === 'notinzip' ? 'That zip does not hold the GSTR-2 Excel template. Download the Returns Offline Tool from gst.gov.in › Downloads › Offline Tools.' : w === 'notg2' || w === 'notxl' ? 'That file is not the GSTR-2 Excel Workbook Template (it comes inside the Returns Offline Tool from gst.gov.in).' : 'Could not fill the template: ' + (w || e));
    });
  }
  /* ---------- IMS (Invoice Management System): actions suggested from the purchase bills ----------
     The records come from the portal: the person downloads them on the IMS dashboard and loads them into the department's IMS
     Offline Tool with its own button. NSBiz then reads that file and, for every supplier invoice, debit note and credit note still
     at "No Action", suggests: Accepted when the same supplier and document number is in the purchase bills with the tax within
     Re 1; Pending with a remark when the tax differs or the document is not in the bills. NSBiz never chooses Rejected: that is the
     person's decision. Rows already decided, and rows where Pending is not allowed (the tool marks them), are left as they are. */
  var IMS = [{ n: 'B2B', st: 'G', rem: 'N', np: 'X', kind: 'inv' }, { n: 'B2B-DN', st: 'G', rem: 'N', np: 'X', kind: 'dn' }, { n: 'B2B-CN', st: 'G', rem: 'S', np: 'AH', kind: 'cn' }];
  function imsAct(book) {
    if (IMS.some(function (x) { return book.sheets.indexOf(x.n) < 0; }) || book.sheets.indexOf('Home') < 0) return Promise.reject(new Error('notims'));
    var mine = books(), idx = {}, R = { acc: 0, pend: 0, diff: 0, kept: 0, rows: 0 };
    mine.forEach(function (b) { idx[b.gstin + '|' + b.kind + '|' + nkey(b.no)] = b; });
    var val = function (S, ref) { var c = S.get(ref); return c ? String(c.v == null ? '' : c.v).trim() : ''; };
    return IMS.reduce(function (p, x) {
      return p.then(function () { return book.sheet(x.n); }).then(function (S) {
        for (var r = 7; r < 20007; r++) {
          var g = g15(val(S, 'A' + r)); if (!g) break;
          R.rows++;
          var cur = val(S, x.st + r); if (cur && !/^no action$/i.test(cur)) { R.kept++; continue; }
          var no = val(S, 'C' + r), tax = ['J', 'K', 'L', 'M'].reduce(function (a, c) { return a + Math.abs(num(val(S, c + r))); }, 0);
          var b = idx[g + '|' + x.kind + '|' + nkey(no)], noPend = /^y$/i.test(val(S, x.np + r));
          if (b && Math.abs(Math.abs(tax) - Math.abs(tax2(b))) <= 1) { S.set(x.st + r, 'Accepted'); R.acc++; }
          else if (noPend) { R.kept++; }
          else {
            S.set(x.st + r, 'Pending');
            S.set(x.rem + r, b ? 'Tax in our bill ' + money(tax2(b), false) + ', here ' + money(tax, false) + ': check with the supplier' : 'Not in our purchase bills yet');
            if (b) R.diff++; else R.pend++;
          }
        }
      });
    }, Promise.resolve()).then(function () { return R; });
  }
  function tax2(b) { return Math.abs(b.igst) + Math.abs(b.cgst) + Math.abs(b.sgst) + Math.abs(b.cess || 0); }
  function imsUtil() {
    if (!NS.xlFill) return;
    NS.xlFill.pick('.xlsm').then(function (f) {
      if (!f) return;
      NS.toast('Reading ' + f.name + '…');
      return NS.xlFill.open(f).then(function (book) { return imsAct(book).then(function (R) {
        var me = SELLER() || {}, nm = f.name.replace(/\.xlsm$/i, '') + ' NSBiz.xlsm';
        return book.save(nm).then(function (bl) { return NS.xlFill.toFolder(g15(me.gstin), 'IMS ' + NS.today().slice(0, 7), nm, bl); }).then(function (inF) {
          NS.toast('Saved as "' + nm + '": ' + R.acc + ' accepted, ' + R.diff + ' with a different amount and ' + R.pend + ' not in your bills marked Pending, ' + R.kept + ' left as they were (of ' + R.rows + ').' + (inF ? ' Also saved in NSBiz › GST Returns.' : '') + ' Check each Pending one, change any you want to reject, then upload with the tool\'s own button.');
        });
      }); });
    }).catch(function (e) {
      var w = e && e.message;
      NS.toast(w === 'notims' || w === 'nosheet' || w === 'notxl' ? 'Choose the IMS Offline Tool file (.xlsm) into which you loaded the records downloaded from the IMS dashboard.' : 'Could not read the file: ' + (w || e), 'bad');
    });
  }
  function setRange(p) {   /* 2B period MMYYYY -> that month */
    var m2 = String(p || '').match(/^(\d{2})(\d{4})$/);
    if (!m2) return;
    var y = +m2[2], mo = +m2[1], pad = function (n) { return (n < 10 ? '0' : '') + n; };
    $('#g2-from').value = y + '-' + pad(mo) + '-01'; $('#g2-to').value = y + '-' + pad(mo) + '-' + pad(new Date(y, mo, 0).getDate());
  }
  function open(file) {
    var st = $('#g2-file-st');
    st.textContent = t('g2.reading', { f: file.name });
    var p = /\.json$/i.test(file.name) || file.type === 'application/json' ? file.text().then(function (s) { return fromJson(JSON.parse(s)); }) : file.arrayBuffer().then(fromExcel);
    p.then(function (o) {
      if (!o.rows.length) { st.textContent = t('g2.empty'); return; }
      TWOB = o; setRange(o.period);
      /* pick the business whose GSTIN the 2B file is for */
      var hit = o.gstin && BIZ.concat(DOCS.map(function (d) { return d.seller; })).filter(function (x) { return x && g15(x.gstin) === g15(o.gstin); })[0];
      if (hit && $('#g2-biz option[value="' + key(hit.name).replace(/"/g, '') + '"]')) $('#g2-biz').value = key(hit.name);
      st.textContent = t('g2.read', { n: o.rows.length, g: o.gstin || '—', p: o.period ? o.period.slice(0, 2) + '/' + o.period.slice(2) : '—' });
      draw();
    }).catch(function () { st.textContent = t('g2.bad'); });
  }
  function start() {
    var rg = $('[data-do="reg"]'); if (rg && !$('[data-do="g2tpl"]')) rg.insertAdjacentHTML('afterend', ' <button type="button" class="btn" data-do="g2tpl" data-ic="xls" title="Fills the official GSTR-2 Excel template of the Returns Offline Tool (gst.gov.in) from your purchase bills of these dates and saves it as a new file">Fill the GSTR-2 template (Excel)</button> <button type="button" class="btn sm" data-do="g2own" data-noic title="Choose a newer Returns Offline Tool download from gst.gov.in">Change template file</button><span class="soft xl-tpl" data-tpl="gstr2"></span> <button type="button" class="btn" data-do="ims" data-ic="xls" title="Choose your IMS Offline Tool file after loading the records downloaded from the IMS dashboard: NSBiz suggests Accepted or Pending for each supplier document from your purchase bills">Suggest IMS actions (Excel)</button>'); showTpl();
    document.addEventListener('ns-lang', function () { $('#g2-status').textContent = t('g2.loaded', { n: DOCS.filter(function (d) { return d.type === 'pbill'; }).length }); draw(); });   /* the language pack can arrive after the first drawing */
    NSStore.ready().then(function () { return Promise.all([NSStore.list('doc'), NSStore.list('biz')]); }).then(function (x) {
      DOCS = x[0] || []; BIZ = x[1] || [];
      var seen = {}, bl = [];
      BIZ.concat(DOCS.filter(function (d) { return d.type === 'pbill'; }).map(function (d) { return d.seller; })).forEach(function (s) { var k = key(s && s.name); if (k && !seen[k]) { seen[k] = 1; bl.push([k, s.name + (s.gstin ? ' (' + s.gstin + ')' : '')]); } });
      $('#g2-biz').innerHTML = bl.length ? bl.map(function (b) { return '<option value="' + esc(b[0]) + '">' + esc(b[1]) + '</option>'; }).join('') : '<option value="">' + esc(t('g3.nobiz')) + '</option>';
      var n = DOCS.filter(function (d) { return d.type === 'pbill'; }).length;
      $('#g2-status').textContent = t('g2.loaded', { n: n });
      if (!$('#g2-from').value) { var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); setRange(('0' + (d.getMonth() + 1)).slice(-2) + d.getFullYear()); }
      draw();
    }).catch(function () { $('#g2-status').textContent = t('g1.fail'); });
    $('#g2-file').addEventListener('change', function () { if (this.files[0]) open(this.files[0]); });
    var drop = $('#g2-drop');
    ['dragover', 'dragenter'].forEach(function (e) { drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.add('over'); }); });
    ['dragleave', 'drop'].forEach(function (e) { drop.addEventListener(e, function () { drop.classList.remove('over'); }); });
    drop.addEventListener('drop', function (ev) { ev.preventDefault(); if (ev.dataTransfer.files[0]) open(ev.dataTransfer.files[0]); });
    root.addEventListener('change', function (e) { if (e.target.closest('.g1-ctl')) draw(); });
    root.addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      if (b.dataset.csv) csv(b.dataset.csv);
      else if (b.dataset.add2b && R) addBill(R.only2b[+b.dataset.add2b], b);
      else if (b.dataset.do === 'reg') register();
      else if (b.dataset.do === 'g2tpl') fillG2Tpl(false);
      else if (b.dataset.do === 'g2own') fillG2Tpl(true);
      else if (b.dataset.do === 'ims') imsUtil();
      else if (b.dataset.do === 'print') window.print();
      else if (b.dataset.do === 'share' && R) NS.sharePdf([$('#g2-out')], 'GSTR-2B match ' + (TWOB.period || ''), 'a4', 'GSTR-2B match', 'GSTR-2B match');
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
