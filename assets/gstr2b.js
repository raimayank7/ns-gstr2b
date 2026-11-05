/* GSTR-2B matcher (gstr2b.niveshsahayak.in): the user opens the GSTR-2B JSON or Excel downloaded from the GST portal; it stays in
   the browser and is compared with the purchase bills saved in the data home (Nivesh Sahayak App purchase bills and returns).
   Result: matched, amount or date differs, number differs (same supplier and amount), in 2B but not in your books, in your books
   but not in 2B, ITC not available in 2B, and suppliers with nothing in 2B. The 2B file is read by field names, not fixed positions,
   so small format changes on the portal do not break it. Purchase register download in the layout of the portal's Matching Offline
   Tool v2.9 template (PurchaseRegister_Template_v2.4: GSTIN C1, year E1, period E2, headers row 5, data from row 6). Screen words: NS.t('g2.*'). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, t = NS.t, money = NS.money, B = window.NSBooks;
  var root = $('#g2');
  if (!root) return;
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
  function fromExcel(buf) {
    return NS.loadJs(XLSX_JS[0], XLSX_JS[1]).then(function () {
      var wb = new window.ExcelJS.Workbook();
      return wb.xlsx.load(buf).then(function () {
        var out = { gstin: '', period: '', rows: [] };
        wb.eachSheet(function (ws) {
          var name = ws.name.toUpperCase().replace(/\s/g, '');
          var sec = /^B2B$/.test(name) ? 'inv' : /^B2B-CDNR$|^CDNR$/.test(name) ? 'note' : /^B2BA$|CDNRA/.test(name) ? 'amend' : '';
          if (!out.gstin) ws.eachRow(function (r, n) { if (n < 12) r.eachCell(function (c) { var m = String(c.value || '').match(/\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]\b/); if (m && !out.gstin) out.gstin = m[0]; }); });
          if (!sec) return;
          var hr = 0, map = {};
          for (var i = 1; i <= Math.min(12, ws.rowCount) && !hr; i++) ws.getRow(i).eachCell(function (c) { if (/gstin of (the )?supplier/i.test(String(c.value || ''))) hr = i; });
          if (!hr) return;
          [hr, hr + 1].forEach(function (rn) {
            ws.getRow(rn).eachCell({ includeEmpty: false }, function (c, col) {
              var txt = String(c.value && c.value.richText ? c.value.richText.map(function (x) { return x.text; }).join('') : c.value || '').replace(/\s+/g, ' ');
              COLS.forEach(function (k) { if (map[k[0]] == null && k[1].test(txt)) map[k[0]] = col; });
            });
          });
          if (map.gstin == null || map.no == null) return;
          ws.eachRow(function (r, n) {
            if (n <= hr + 1) return;
            var v = function (k) { var c = map[k] != null ? r.getCell(map[k]).value : ''; return c && c.result != null ? c.result : c; };
            var g = g15(v('gstin'));
            if (!/^\d{2}[A-Z0-9]{13}$/.test(g)) return;
            var nt = String(v('ntype') || '').toLowerCase();
            out.rows.push({ gstin: g, name: String(v('name') || ''), kind: sec === 'note' || nt ? (/debit/.test(nt) ? 'dn' : 'cn') : 'inv', no: String(v('no') || ''), date: iso(v('date')), value: num(v('value')),
              txv: num(v('txv')), igst: num(v('igst')), cgst: num(v('cgst')), sgst: num(v('sgst')), cess: num(v('cess')), itc: !/^n/i.test(String(v('itc') || 'Yes')), rsn: String(v('rsn') || ''),
              rev: /^y/i.test(String(v('rev') || 'No')), filed: '', sprd: '', sec: sec === 'amend' ? 'amend' : '' });
          });
        });
        return out;
      });
    });
  }

  /* ---------- our purchase bills ---------- */
  function books() {
    var bz = $('#g2-biz').value;
    return DOCS.filter(function (d) { return (d.type === 'pbill' || d.type === 'note' && d.side === 'purchase') && key(d.seller && d.seller.name) === bz; }).map(function (d) {
      var pn = d.type === 'note', back = pn ? d.kind === 'debit' : d.kind === 'return', tt = B.itemTotals(d), s = back ? -1 : 1;
      return { id: d.id, gstin: g15(d.buyer && d.buyer.gstin), name: d.buyer && d.buyer.name || '', kind: back ? 'cn' : pn ? 'dn' : 'inv', no: d.no || '', date: d.date || '', value: s * tt.total,
        txv: s * tt.taxable, igst: s * tt.igst, cgst: s * tt.cgst, sgst: s * tt.sgst, cess: 0, noGst: pn ? d.withGst === false : !!d.noGst };
    }).filter(function (b) { return !b.noGst; });
  }
  function tax(x) { return r2(Math.abs(x.igst) + Math.abs(x.cgst) + Math.abs(x.sgst) + Math.abs(x.cess)); }
  function inRange(d) { var f = $('#g2-from').value, to = $('#g2-to').value; return (!f || d >= f) && (!to || d <= to); }

  function match() {
    var tol = num($('#g2-tol').value), mine = books(), two = TWOB ? TWOB.rows.filter(function (x) { return x.sec !== 'isd'; }) : [];
    var used = {}, res = { ok: [], diff: [], numdiff: [], only2b: [], onlybooks: [], noitc: [], nofile: [], t2b: 0, tbooks: 0, tok: 0 };
    var idx = {}; mine.forEach(function (b, i) { var k = b.gstin + '|' + b.kind + '|' + nkey(b.no); (idx[k] = idx[k] || []).push(i); });
    two.forEach(function (x) {
      res.t2b += x.itc ? tax(x) : 0;
      if (!x.itc) res.noitc.push(x);
      var k = x.gstin + '|' + x.kind + '|' + nkey(x.no), hit = (idx[k] || []).filter(function (i) { return !used[i]; })[0];
      if (hit == null) {   /* same supplier, same tax, different number */
        hit = mine.map(function (b, i) { return i; }).filter(function (i) { var b = mine[i]; return !used[i] && b.gstin === x.gstin && b.kind === x.kind && Math.abs(tax(b) - tax(x)) <= tol && Math.abs(Math.abs(b.txv) - Math.abs(x.txv)) <= tol; })[0];
        if (hit != null) { used[hit] = 1; res.numdiff.push({ b: mine[hit], x: x }); return; }
        res.only2b.push(x); return;
      }
      used[hit] = 1;
      var b = mine[hit], dt = Math.abs(Math.abs(b.txv) - Math.abs(x.txv)), dx = Math.abs(tax(b) - tax(x));
      if (dt <= tol && dx <= tol && Math.abs(Math.abs(b.igst) - Math.abs(x.igst)) <= tol && (b.date === x.date || !x.date)) { res.ok.push({ b: b, x: x }); res.tok += x.itc ? tax(x) : 0; }
      else res.diff.push({ b: b, x: x, why: [dt > tol || dx > tol ? t('g2.why.amt') : '', Math.abs(Math.abs(b.igst) - Math.abs(x.igst)) > tol && dx <= tol ? t('g2.why.head') : '', b.date !== x.date && x.date ? t('g2.why.date') : ''].filter(Boolean).join(', ') });
    });
    mine.forEach(function (b, i) { if (!used[i] && inRange(b.date)) { res.onlybooks.push(b); } if (inRange(b.date)) res.tbooks += tax(b); });
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
    $('#g2-files').hidden = false;
  }
  /* a bill found only in GSTR-2B: save it as a purchase bill (unpaid) in the data home, one line at the bill's tax rate */
  var RATES = [0, 0.1, 0.25, 1.5, 3, 5, 12, 18, 28, 40];
  function addBill(x, btn) {
    var s = SELLER(); if (!x || !s || !s.name) { NS.toast(t('g3.nobiz'), 'bad'); return; }
    var txv = Math.abs(x.txv), g = txv ? (Math.abs(x.igst) + Math.abs(x.cgst) + Math.abs(x.sgst)) / txv * 100 : 0;
    var near = RATES.reduce(function (a, r) { return Math.abs(r - g) < Math.abs(a - g) ? r : a; }, 18);
    if (!confirm(t('uf.add2bq', { no: x.no, p: x.name || x.gstin }))) return;
    var sel = JSON.parse(JSON.stringify(s)); delete sel.id; delete sel.created; delete sel.updated;
    var d = { type: 'pbill', kind: x.kind === 'cn' ? 'return' : 'bill', app: 1, from2b: 1, seller: sel, buyer: { name: x.name || x.gstin, gstin: x.gstin, state: x.gstin.slice(0, 2) }, no: x.no, date: x.date || NS.today(), round: false, noGst: false,
      items: [{ desc: 'Purchase as per GSTR-2B', hsn: '', qty: 1, unit: 'Nos', rate: txv, gst: Math.abs(near - g) <= 0.1 ? near : r2(g) }] };
    if (d.kind === 'return') d.orig = { no: '' }; else d.due = '';
    var tt = B.itemTotals(d); d.tax = { taxable: tt.taxable, cgst: tt.cgst, sgst: tt.sgst, igst: tt.igst }; d.total = r2(tt.taxable + tt.cgst + tt.sgst + tt.igst);
    btn.disabled = true;
    NSStore.put('doc', d).then(function (rec) { DOCS.push(rec); NS.toast(t('uf.add2bok', { no: x.no })); draw(); }, function (err) { btn.disabled = false; NS.toast(err.message, 'bad'); });
  }
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
      else if (b.dataset.do === 'print') window.print();
      else if (b.dataset.do === 'share' && R) NS.sharePdf([$('#g2-out')], 'GSTR-2B match ' + (TWOB.period || ''), 'a4', 'GSTR-2B match', 'GSTR-2B match');
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
