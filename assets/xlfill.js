/* NS.xlFill: writes figures into a COPY of an Excel workbook (the GST portal's offline utilities), keeping its macros, styles,
   drop-down lists and its own formulas. The file the user chose is only read; the result is a new file to save under another name.
   open(file) -> book: { names: {name: [refs]}, sheet(name) -> sheet: { label(text) -> ref, set(ref, v), clear(ref) }, save(fname) } */
(function () {
  var NS = window.NS = window.NS || {};
  var JSZIP = ['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'sha512-XMVd28F1oH/O71fzwBnV7HucLxVwtxf26XV8P4wPk26EDxuGZ91N8bsOttmnomcCD3CS5ZMRL50H0GgOHvegtg=='];
  var xe = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
  var un = function (s) { return String(s).replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); };
  var colN = function (c) { var n = 0; for (var i = 0; i < c.length; i++) n = n * 26 + c.charCodeAt(i) - 64; return n; };
  var colL = function (n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  var split = function (ref) { var m = /^([A-Z]+)(\d+)$/.exec(ref); return [m[1], +m[2]]; };
  /* every cell named in a range text like 'GSTR-3B'!$C$11,'GSTR-3B'!$D$11:$F$12 */
  function refs(txt) {
    var out = []; if (/[()]/.test(txt)) return out;   /* a named formula, not cells */
    String(txt).split(',').forEach(function (p) {
      var a = p.trim().replace(/^.*!/, '').replace(/\$/g, '').split(':');
      if (a.length > 2 || !/^[A-Z]+\d+$/.test(a[0]) || (a[1] && !/^[A-Z]+\d+$/.test(a[1]))) return;   /* named formulas (OFFSET(...)) are not cells */
      var s = split(a[0]), e = split(a[1] || a[0]);
      for (var r = s[1]; r <= e[1]; r++) for (var c = colN(s[0]); c <= colN(e[0]); c++) out.push(colL(c) + r);
    });
    return out;
  }
  function open(file) {
    var lib = window.JSZip ? Promise.resolve() : NS.loadJs(JSZIP[0], JSZIP[1]);
    return lib.then(function () { return window.JSZip.loadAsync(file); }).then(function (zip) {
      var rd = function (p) { var f = zip.file(p); return f ? f.async('string') : Promise.resolve(''); };
      return Promise.all([rd('xl/workbook.xml'), rd('xl/_rels/workbook.xml.rels'), rd('xl/sharedStrings.xml'), rd('[Content_Types].xml')]).then(function (x) {
        var wb = x[0], rels = x[1], CT = x[3], SST = (x[2].match(/<si>[\s\S]*?<\/si>/g) || []).map(un);
        if (!wb) throw new Error('notxl');
        var names = {}; (wb.match(/<definedName [^>]*>[^<]*<\/definedName>/g) || []).forEach(function (d) { var n = /name="([^"]+)"/.exec(d)[1]; if (!/localSheetId/.test(d) || !names[n]) names[n] = refs(un(d)); });
        var sheets = {}, open = {};
        (wb.match(/<sheet [^>]*\/>/g) || []).forEach(function (s) {
          var nm = un(/name="([^"]*)"/.exec(s)[1]).trim(), id = /r:id="([^"]+)"/.exec(s)[1], rel = new RegExp('<Relationship [^>]*Id="' + id + '"[^>]*>').exec(rels);
          var tg = rel && /Target="([^"]+)"/.exec(rel[0])[1]; if (tg) sheets[nm] = tg.charAt(0) === '/' ? tg.slice(1) : 'xl/' + tg.replace(/^\.\//, '');
        });
        function sheet(name) {
          var path = sheets[name]; if (!path) return Promise.reject(new Error('nosheet'));
          if (open[name]) return Promise.resolve(open[name]);
          return rd(path).then(function (xml) {
            var S = { xml: xml, path: path };
            var cellRe = function (ref) { return new RegExp('<c r="' + ref + '"(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</c>)'); };
            S.get = function (ref) { var m = cellRe(ref).exec(S.xml); if (!m) return null; var v = /<v>([^<]*)<\/v>/.exec(m[0]), t = /\st="([^"]+)"/.exec(m[0]); return { raw: m[0], f: (/<f[^>]*>([^<]*)<\/f>/.exec(m[0]) || [])[1], v: v ? (t && t[1] === 's' ? SST[+v[1]] : un(v[1])) : (/<t[^>]*>([^<]*)<\/t>/.exec(m[0]) || [])[1] || '' }; };
            /* the cell holding exactly this text (labels such as "GSTIN", "Year") */
            S.label = function (text) {
              var want = String(text).trim().toLowerCase(), re = /<c r="([A-Z]+\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, m;
              while ((m = re.exec(S.xml))) { var c = S.get(m[1]); if (c && String(c.v).trim().toLowerCase() === want) return m[1]; }
              return null;
            };
            /* write a value; the workbook's own copy formulas (a cell that just repeats another cell, like SGST = CGST) are kept */
            S.set = function (ref, val) {
              var c = S.get(ref); if (c && c.f && /^\$?[A-Z]+\$?\d+$/.test(c.f.trim())) return false;
              var st = c ? (/\ss="(\d+)"/.exec(c.raw) || [])[0] || '' : '';
              var cell = val === '' || val == null ? '<c r="' + ref + '"' + st + '/>' : typeof val === 'number' ? '<c r="' + ref + '"' + st + '><v>' + (Math.round(val * 100) / 100) + '</v></c>' : '<c r="' + ref + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + xe(val) + '</t></is></c>';
              if (c) { S.xml = S.xml.replace(c.raw, function () { return cell; }); return true; }
              var p = split(ref), row = new RegExp('<row r="' + p[1] + '"[^>]*?(?:/>|>[\\s\\S]*?</row>)').exec(S.xml);
              if (!row) {   /* no such row yet: add it in order */
                var after = null, rr = /<row r="(\d+)"/g, mm; while ((mm = rr.exec(S.xml))) if (+mm[1] < p[1]) after = mm; else break;
                var ins = '<row r="' + p[1] + '">' + cell + '</row>';
                if (after) { var end = S.xml.indexOf('</row>', after.index); var se = S.xml.indexOf('/>', after.index); var cut = (se > -1 && se < S.xml.indexOf('>', after.index) + 1) ? se + 2 : end + 6; S.xml = S.xml.slice(0, cut) + ins + S.xml.slice(cut); }
                else S.xml = S.xml.replace(/<sheetData>/, '<sheetData>' + ins).replace(/<sheetData\/>/, '<sheetData>' + ins + '</sheetData>');
                return true;
              }
              var r0 = row[0], body = r0.replace(/\/>$/, '>') , cells = body.match(/<c r="[A-Z]+\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || [], at = -1;
              cells.forEach(function (x, i) { if (at < 0 && colN(split(/r="([A-Z]+\d+)"/.exec(x)[1])[0]) > colN(p[0])) at = i; });
              var nr = at < 0 ? body.replace(/(<\/row>)?$/, '') + cell + '</row>' : body.replace(cells[at], function () { return cell + cells[at]; });
              if (!/<\/row>$/.test(nr)) nr += '</row>';
              S.xml = S.xml.replace(r0, function () { return nr; }); return true;
            };
            S.clear = function (ref) { return S.set(ref, ''); };
            /* write a table from row `start` down in one pass: every row from `start` on is emptied first (templates come with example
               rows), each cell keeps the formatting of its column, and the rows of data are written in order. data = [[v, v, ...], ...] */
            S.putRows = function (start, data) {
              var m = /<sheetData>([\s\S]*?)<\/sheetData>|<sheetData\s*\/>/.exec(S.xml); if (!m) return;
              var rows = [], re = /<row r="(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, x, colStyle = {};
              while ((x = re.exec(m[1] || ''))) {
                var cells = [], cr = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g, y;
                while ((y = cr.exec(x[3] || ''))) { var st = (/\ss="(\d+)"/.exec(y[2]) || [])[1]; cells.push({ col: y[1], s: st, raw: y[0] }); if (+x[1] === start && st) colStyle[y[1]] = st; }
                rows.push({ r: +x[1], attr: x[2].replace(/\sspans="[^"]*"/, ''), cells: cells });
              }
              var byR = {}; rows.forEach(function (r) { byR[r.r] = r; });
              rows.forEach(function (r) { if (r.r >= start) r.cells = r.cells.map(function (c) { return { col: c.col, s: c.s, raw: '<c r="' + c.col + r.r + '"' + (c.s ? ' s="' + c.s + '"' : '') + '/>' }; }); });
              (data || []).forEach(function (vals, i) {
                var rn = start + i, row = byR[rn] || (byR[rn] = { r: rn, attr: '', cells: [] }); if (rows.indexOf(row) < 0) rows.push(row);
                var have = {}; row.cells.forEach(function (c) { have[c.col] = c; });
                vals.forEach(function (v, j) {
                  var col = colL(j + 1), st = have[col] && have[col].s || colStyle[col], sa = st ? ' s="' + st + '"' : '';
                  var raw = v === '' || v == null ? '<c r="' + col + rn + '"' + sa + '/>' : typeof v === 'number' ? '<c r="' + col + rn + '"' + sa + '><v>' + (Math.round(v * 1000) / 1000) + '</v></c>'
                    : '<c r="' + col + rn + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + xe(v) + '</t></is></c>';
                  if (have[col]) have[col].raw = raw; else row.cells.push(have[col] = { col: col, s: st, raw: raw });
                });
                row.cells.sort(function (a, b) { return colN(a.col) - colN(b.col); });
              });
              rows.sort(function (a, b) { return a.r - b.r; });
              var body = rows.map(function (r) { return r.cells.length ? '<row r="' + r.r + '"' + r.attr + '>' + r.cells.map(function (c) { return c.raw; }).join('') + '</row>' : '<row r="' + r.r + '"' + r.attr + '/>'; }).join('');
              S.xml = S.xml.slice(0, m.index) + '<sheetData>' + body + '</sheetData>' + S.xml.slice(m.index + m[0].length);
            };
            /* the values under a heading cell (a template's own drop-down list, for example its POS list) */
            S.listUnder = function (heading) {
              var ref = S.label(heading); if (!ref) return [];
              var c = /^[A-Z]+/.exec(ref)[0], r = +/\d+$/.exec(ref)[0], out = [];
              for (var i = r + 1; i < r + 200; i++) { var v = S.get(c + i); if (!v || v.v === '') break; out.push(String(v.v)); }
              return out;
            };
            open[name] = S; return S;
          });
        }
        function save(fname) {
          Object.keys(open).forEach(function (n) { zip.file(open[n].path, open[n].xml); });
          /* Excel works the formulas out again on opening; its saved order of formulas is dropped so it never asks to repair the file */
          zip.file('xl/workbook.xml', /<calcPr/.test(wb) ? wb.replace(/<calcPr([^>]*?)\s*\/>/, function (a, b) { return '<calcPr' + b.replace(/\sfullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>'; }) : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'));
          if (zip.file('xl/calcChain.xml')) {
            zip.remove('xl/calcChain.xml');
            zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship [^>]*calcChain[^>]*\/>/g, ''));
            zip.file('[Content_Types].xml', CT.replace(/<Override [^>]*calcChain[^>]*\/>/g, ''));
          }
          var mac = /\.xlsm$/i.test(fname);
          return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: mac ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }).then(function (b) {
            var a = document.createElement('a'), u = URL.createObjectURL(b); a.href = u; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 4000); return b;
          });
        }
        return { names: names, sheets: Object.keys(sheets), sheet: sheet, save: save, refs: refs };
      });
    });
  }
  /* ask for the utility file; resolves with the File (or null) */
  function pick(accept) {
    return new Promise(function (res) { var i = document.createElement('input'); i.type = 'file'; i.accept = accept || '.xlsm,.xlsx'; i.onchange = function () { res(i.files && i.files[0] || null); }; i.click(); });
  }
  /* Excel's day number for a date written 2026-10-15 or 15-Oct-2026 (templates format the cell themselves); other text unchanged */
  var MON3 = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function serial(v) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v), y, mo, d;
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; } else if ((m = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec(v)) && MON3[m[2].toLowerCase()]) { d = +m[1]; mo = MON3[m[2].toLowerCase()]; y = +m[3]; } else return v;
    return Math.round((Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);
  }
  /* "07-Delhi" spelled exactly as the template's own list has it, chosen by the state code */
  function posFrom(list, text) { var c = String(text || '').slice(0, 2); if (!/^\d\d$/.test(c)) return text; for (var i = 0; i < list.length; i++) if (list[i].slice(0, 2) === c) return list[i]; return text; }
  /* ---------- the GST portal's templates, downloaded once by the customer and kept on this device ----------
     KINDS: what each template is, where gst.gov.in offers it, and how to find it inside the downloaded zip */
  var KINDS = {
    gstr1: { title: 'GSTR-1 Excel template', tool: 'Returns Offline Tool', page: 'https://www.gst.gov.in/download/returns', zip: 'gst_offline_tool.zip, about 29 MB',
      file: /GSTR1_Excel_Workbook_Template[^\/]*\.xlsx$/i, ext: '.xlsx' },
    gstr2: { title: 'GSTR-2 Excel template', tool: 'Returns Offline Tool', page: 'https://www.gst.gov.in/download/returns', zip: 'gst_offline_tool.zip, about 29 MB',
      file: /GSTR2_Excel_Workbook_Template[^\/]*\.xlsx$/i, ext: '.xlsx' },
    itc04: { title: 'ITC-04 Offline Utility', tool: 'ITC-04 Offline Tool', page: 'https://www.gst.gov.in/download/itc04', zip: 'ITC_04_Offline_Utility.zip, about 0.25 MB',
      file: /ITC_04_Offline_Utility[^\/]*\.xlsm$/i, ext: '.xlsm' },
    gstr4a: { title: 'GSTR-4 (annual) Offline Utility', tool: 'GSTR-4 Offline Tool (Annual)', page: 'https://www.gst.gov.in/download/gstr4x', zip: 'GSTR_4_Annual_Offline_Utility.zip, about 3 MB',
      file: /GSTR_4_Annual_Offline_Utility[^\/]*\.xlsm$/i, ext: '.xlsm' },
    gstr9: { title: 'GSTR-9 Offline Tool', tool: 'GSTR-9 Offline Tool', page: 'https://www.gst.gov.in/download/gstr9', zip: 'GSTR_9_Offline_Utility.zip, about 1 MB',
      file: /GSTR-?9 Offline Tool[^\/]*\.xlsm$/i, ext: '.xlsm' },
    itc01: { title: 'ITC-01 Offline Tool', tool: 'ITC-01 Offline Tool', page: 'https://www.gst.gov.in/download/itc01', zip: 'ITC_01_Offline_v1.3.zip, about 1.7 MB', file: /ITC_01_Offline[^\/]*\.xls$/i, ext: '.xls' },
    itc03: { title: 'ITC-03 Offline Tool', tool: 'ITC-03 Offline Tool', page: 'https://www.gst.gov.in/download/itc03', zip: 'ITC_03_Offline_v2.0.zip, about 0.3 MB', file: /ITC_03_Offline[^\/]*\.xls$/i, ext: '.xls' },
    gstr10: { title: 'GSTR-10 Offline Tool', tool: 'GSTR10 Offline Tool', page: 'https://www.gst.gov.in/download/gstr10', zip: 'GSTR10_Offline_Utility.zip, about 0.4 MB', file: /GSTR10_Offline_Utility[^\/]*\.xlsb$/i, ext: '.xlsb' },
    gstr7: { title: 'GSTR-7 Offline Utility', tool: 'GSTR7 Offline Utility', page: 'https://www.gst.gov.in/download/gstr7', zip: 'GSTR7_Offline_Utility.zip, about 0.6 MB', file: /GSTR7_Offline_Utility[^\/]*\.xlsm$/i, ext: '.xlsm' },
    gstr6: { title: 'GSTR-6 Offline Tool', tool: 'GSTR 6 Offline Tool', page: 'https://tutorial.gst.gov.in/offlineutilities/returns/GSTR_6_Offline_Utility.zip', zip: 'GSTR_6_Offline_Utility.zip, about 4 MB', file: /GSTR-?6_Offline_Utility[^\/]*\.xlsm$/i, ext: '.xlsm' },
    gstr3b: { title: 'GSTR-3B Offline Utility', tool: 'GSTR3B Offline Utility', page: 'https://www.gst.gov.in/download/gstr3b', zip: 'GSTR3B_Excel_Utility.zip, about 0.2 MB',
      file: /GSTR3B_Excel_Utility[^\/]*\.xlsm$/i, ext: '.xlsm' }
  };
  var DB = 'ns-gst-templates';
  function db() { return new Promise(function (res, rej) { var r = indexedDB.open(DB, 1); r.onupgradeneeded = function () { r.result.createObjectStore('t'); }; r.onsuccess = function () { var d = r.result; d.onversionchange = function () { d.close(); }; res(d); }; r.onerror = function () { rej(r.error); }; }); }
  function kept(kind) { return db().then(function (d) { return new Promise(function (res) { var q = d.transaction('t').objectStore('t').get(kind); q.onsuccess = function () { d.close(); res(q.result || null); }; q.onerror = function () { d.close(); res(null); }; }); }).catch(function () { return null; }); }
  function keep(kind, rec) { return db().then(function (d) { return new Promise(function (res, rej) { var tx = d.transaction('t', 'readwrite'); rec ? tx.objectStore('t').put(rec, kind) : tx.objectStore('t').delete(kind); tx.oncomplete = function () { d.close(); res(); }; tx.onerror = function () { d.close(); rej(tx.error); }; }); }); }
  /* version number in a file name: "..._V2.2.xlsx" -> [2, 2]; the highest wins when a zip holds several */
  function ver(n) { var m = /v\s*(\d+(?:\.\d+)*)/i.exec(String(n).replace(/\.xls[xm]?$/i, '')); return m ? m[1].split('.').map(Number) : [0]; }
  function newer(a, b) { var x = ver(a), y = ver(b); for (var i = 0; i < Math.max(x.length, y.length); i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); } return false; }
  /* the template out of what the customer chose: the file itself, or the downloaded zip (also a zip inside the zip) */
  function extract(kind, f) {
    var K = KINDS[kind], name = f.name || '';
    if (K.file.test(name)) return Promise.resolve({ name: name, data: f });
    if (!/\.zip$/i.test(name)) {
      if (/\.xls[xmb]?$/i.test(name)) return Promise.resolve({ name: name, data: f });   /* a renamed copy: checked when it is filled */
      return Promise.reject(new Error('notzip'));
    }
    var lib = window.JSZip ? Promise.resolve() : NS.loadJs(JSZIP[0], JSZIP[1]);
    var look = function (zip, depth) {
      var best = null, inner = [];
      zip.forEach(function (p, e) { if (e.dir) return; var b = p.split('/').pop(); if (K.file.test(b) && (!best || newer(b, best.name))) best = { name: b, e: e }; else if (/\.zip$/i.test(b)) inner.push(e); });
      if (best) return best.e.async('blob').then(function (bl) { return { name: best.name, data: bl }; });
      if (depth > 1 || !inner.length) return Promise.reject(new Error('notinzip'));
      return inner.reduce(function (p, e) { return p.catch(function () { return e.async('arraybuffer').then(function (ab) { return window.JSZip.loadAsync(ab); }).then(function (z) { return look(z, depth + 1); }); }); }, Promise.reject(new Error('notinzip')));
    };
    return lib.then(function () { return window.JSZip.loadAsync(f); }).then(function (z) { return look(z, 0); });
  }
  /* the first-time guide: where to download it, then choose the file (or the zip as downloaded) */
  function guide(kind, current) {
    var K = KINDS[kind];
    return new Promise(function (res) {
      var d = document.createElement('dialog'); d.className = 'b2-dlg xl-guide';
      d.innerHTML = '<form method="dialog"><h3>' + (current ? 'Change the ' : 'The official ') + K.title + '</h3>' +
        (current ? '<p class="soft">Now using <b></b>, kept on this device since <span></span>.</p>' : '<p>NSBiz fills the government\'s own file, so you open and upload it exactly as usual. You need it only once: NSBiz keeps it on this device and uses it from then on, also without internet.</p>') +
        '<ol><li>Open <a target="_blank" rel="noopener"></a> and download the <b>' + K.tool + '</b>: the file ' + K.zip + '.</li>' +
        '<li>Choose the downloaded file below. The zip is fine as it is: there is no need to unzip it.</li></ol>' +
        '<p class="row"><button type="button" class="btn primary" data-x="pick" data-ic="folder">Choose the downloaded file</button> <button type="button" class="btn" data-x="close">Cancel</button></p></form>';
      var a = d.querySelector('a'); a.href = K.page; a.textContent = K.page.replace(/^https:\/\//, '');
      if (current) { d.querySelector('p.soft b').textContent = current.name; d.querySelector('p.soft span').textContent = new Date(current.saved).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' }); }
      var done = function (v) { try { d.close(); } catch (e) {} d.remove(); res(v); };
      d.addEventListener('click', function (e) { var x = e.target.closest('[data-x]'); if (!x) return; if (x.dataset.x === 'close') done(null); else (NS.xlFill.pick || pick)('.zip,' + K.ext).then(function (f) { if (f) done(f); }); });
      d.addEventListener('cancel', function () { done(null); });
      document.body.appendChild(d); d.showModal();
    });
  }
  /* the template to fill: the one kept on this device, or (first time / "change") the customer chooses it after the guide.
     Resolves { file, name, fresh } or null; fresh = just chosen (kept only after it was filled without problems: see good/bad) */
  function template(kind, change) {
    var inFolder = function () { return NS.folderFile ? NS.folderFile('find', ['GST Templates'], '', KINDS[kind].file.source) : Promise.resolve(null); };
    return kept(kind).then(function (k) {
      if (k && !change) return { file: new File([k.data], k.name), name: k.name, fresh: false };
      if (!change) return inFolder().then(function (x) {   /* kept in the NSBiz folder (another browser, or this one's data was cleared) */
        if (x) return keep(kind, { name: x.name, data: x.blob, saved: Date.now(), size: x.blob.size }).catch(function () {}).then(function () { return { file: new File([x.blob], x.name), name: x.name, fresh: false }; });
        return null;
      }).then(function (t) { return t || ask(k); });
      return ask(k);
    });
    function ask(k) {
      return guide(kind, k).then(function (f) {
        if (!f) return null;
        return extract(kind, f).then(function (x) { return { file: new File([x.data], x.name), name: x.name, fresh: true, rec: { name: x.name, data: x.data, saved: Date.now(), size: x.data.size } }; });
      });
    }
  }
  function good(kind, t) {
    if (!t || !t.fresh) return Promise.resolve();
    if (NS.folderFile) NS.folderFile('put', ['GST Templates'], t.rec.name, t.rec.data);   /* a copy in NSBiz/GST Templates when the folder is on */
    return keep(kind, t.rec).catch(function () {});
  }
  /* a filled return: also into NSBiz/GST Returns/<GSTIN>/<period> when the folder is on. Resolves true when it is there */
  function toFolder(gstin, period, name, blob) { return NS.folderFile && blob ? NS.folderFile('put', ['GST Returns', gstin || 'No GSTIN', period], name, blob) : Promise.resolve(false); }
  function bad(kind, t) { return t && !t.fresh ? keep(kind, null).catch(function () {}) : Promise.resolve(); }   /* a kept file that no longer fills: forget it */
  function info(kind) { return kept(kind).then(function (k) { return k ? { name: k.name, saved: k.saved } : null; }); }
  NS.xlFill = { open: open, pick: pick, refs: refs, serial: serial, posFrom: posFrom, template: template, good: good, bad: bad, info: info, toFolder: toFolder, KINDS: KINDS };
})();
