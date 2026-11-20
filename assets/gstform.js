/* The GST and TDS tools in the official format (owner 20 Nov 2026: "it should follow the official format").
   Inside the App (the tool shown in an App frame: html.embed) the website parts go away (the sales line, the guides below the tool)
   and each tool shows its return the way the government form lays it out: the form name and rule, the header (GSTIN, legal name,
   trade name, period) and the tables under their own numbers, with the verification at the end. The same sheet is what Print / PDF
   gives. On the public website the official sheet is shown under the working.
   Shared bits only: each tool (gstr3b.js, cmp08.js, gstr1.js, gstr2b.js) builds its own tables with these helpers. */
(function () {
  'use strict';
  if (!window.NS) return;
  var esc = NS.esc, inApp = document.documentElement.classList.contains('embed');
  if (inApp) document.documentElement.classList.add('gst-app');
  function money0(v) { v = Math.round((+v || 0) * 100) / 100; return NS.money ? NS.money(v, false) : v.toFixed(2); }
  /* the form's title block and Part 1 / 2: [[number, label, value], ...] */
  function head(form, rule, title, period, rows) {
    return '<div class="off-hd"><b class="off-form">' + esc(form) + '</b><span class="off-rule">' + esc(rule) + '</span>' + (title ? '<span class="off-title">' + esc(title) + '</span>' : '') + '</div>' +
      (period ? '<table class="off-t off-per"><tr>' + period.map(function (p) { return '<th>' + esc(p[0]) + '</th><td>' + esc(p[1] || '') + '</td>'; }).join('') + '</tr></table>' : '') +
      '<table class="off-t off-id"><tbody>' + rows.map(function (r) { return '<tr><td class="off-no">' + esc(r[0]) + '</td><th>' + esc(r[1]) + '</th><td>' + esc(r[2] || '') + '</td></tr>'; }).join('') + '</tbody></table>';
  }
  /* a numbered table: no + caption, column heads, rows of [label, values...]; numbers right aligned with two decimals */
  function table(no, caption, cols, rows, opt) {
    opt = opt || {};
    return '<h4 class="off-cap"><span>' + esc(no) + '</span> ' + esc(caption) + '</h4><table class="off-t"><thead><tr>' + cols.map(function (c, i) { return '<th' + (i ? ' class="n"' : '') + '>' + esc(c) + '</th>'; }).join('') + '</tr></thead><tbody>' +
      (rows.length ? rows.map(function (r) {
        if (r.sub) return '<tr class="off-sub"><th colspan="' + cols.length + '">' + esc(r.sub) + '</th></tr>';
        return '<tr' + (r.tot ? ' class="off-tot"' : '') + '>' + r.cells.map(function (c, i) { return i ? '<td class="n">' + (c === null || c === '' ? '' : typeof c === 'number' ? money0(c) : esc(c)) + '</td>' : '<td>' + esc(c) + '</td>'; }).join('') + '</tr>';
      }).join('') : '<tr><td colspan="' + cols.length + '" class="off-nil">' + esc(opt.nil || 'Nil') + '</td></tr>') + '</tbody></table>';
  }
  function verify() {
    return '<div class="off-ver"><h4 class="off-cap">Verification</h4><p>I hereby solemnly affirm and declare that the information given herein above is true and correct to the best of my knowledge and belief and nothing has been concealed therefrom.</p>' +
      '<table class="off-t off-sign"><tr><th>Signature of Authorised Signatory</th><td></td></tr><tr><th>Name of Authorised Signatory</th><td></td></tr><tr><th>Designation / Status</th><td></td></tr><tr><th>Date</th><td></td></tr><tr><th>Place</th><td></td></tr></table></div>';
  }
  function note(t) { return '<p class="off-note">' + esc(t) + '</p>'; }
  function wrap(html) { return '<div class="off-sheet" dir="ltr" lang="en-IN">' + html + '</div>'; }
  /* inside the App: the working and the boxes fold away under the official sheet */
  function fold(box, label) {
    if (!inApp || !box || box.closest('details.off-work')) return;
    var d = document.createElement('details'); d.className = 'off-work'; d.innerHTML = '<summary class="btn">' + esc(label || 'Change the figures and see the working') + '</summary>';
    box.parentNode.insertBefore(d, box); d.appendChild(box);
  }
  NS.gstForm = { inApp: inApp, head: head, table: table, verify: verify, note: note, wrap: wrap, fold: fold, money: money0 };
})();
