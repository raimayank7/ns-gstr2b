/* "Is GST credit blocked?" checker (section 17(5) of the CGST Act 2017), added below the main tool of gstr2b. and gstr3b.
   Read on 27 Nov 2026 (taxinformation.cbic.gov.in, CGST Act as amended): (a) motor vehicles for persons, 13 seats or fewer; (aa) vessels
   and aircraft; (ab) their insurance, servicing, repair; (b) food, beverages, outdoor catering, beauty, health, cosmetic and plastic surgery,
   renting of such vehicles, life and health insurance, club and fitness membership, leave travel - with the proviso that credit is
   available where an employer is obliged by law to provide it; (c) works contract and (d) own construction of immovable property other
   than plant and machinery (Finance Act 2025 made "plant and machinery" apply from 1 July 2017); (e) composition tax; (fa) CSR; (g) personal
   consumption; (h) goods lost, stolen, destroyed, written off, gifts and free samples; (i) tax paid under section 74 for periods up to
   2023-24, and sections 129 and 130. Screen words: NS.t('gb.*'). */
(function () {
  'use strict';
  var NS = window.NS, host = document.getElementById('g2') || document.getElementById('g3');
  if (!NS || !host) return;
  function T(k, v) { return NS.t(k, v); }
  var esc = NS.esc;
  /* id, status (b blocked, a allowed, d depends), law */
  var L = [['car', 'b', '17(5)(a)'], ['carservice', 'b', '17(5)(ab)'], ['truck', 'a', '17(5)(a)'], ['food', 'b', '17(5)(b)(i)'], ['beauty', 'b', '17(5)(b)(i)'], ['rentcar', 'b', '17(5)(b)(i)'], ['ins', 'b', '17(5)(b)(i)'],
    ['club', 'b', '17(5)(b)(ii)'], ['ltc', 'b', '17(5)(b)(iii)'], ['building', 'b', '17(5)(c), (d)'], ['pm', 'a', '17(5)(c), (d)'], ['gifts', 'b', '17(5)(h)'], ['lost', 'b', '17(5)(h)'], ['personal', 'b', '17(5)(g)'],
    ['csr', 'b', '17(5)(fa)'], ['comp', 'b', '17(5)(e)'], ['demand', 'd', '17(5)(i)']];
  var box = document.createElement('section'); box.className = 'wrap tf-box'; box.setAttribute('aria-labelledby', 'gb-h');
  host.insertAdjacentElement('afterend', box);
  var cur = 'car';
  function draw() {
    var x = L.filter(function (r) { return r[0] === cur; })[0];
    box.innerHTML = '<h2 id="gb-h">' + esc(T('gb.h')) + '</h2><p>' + esc(T('gb.lead')) + '</p><div class="f"><label for="gb-q">' + esc(T('gb.q')) + '</label><select id="gb-q">' + L.map(function (r) { return '<option value="' + r[0] + '"' + (r[0] === cur ? ' selected' : '') + '>' + esc(T('gb.i.' + r[0])) + '</option>'; }).join('') + '</select></div>' +
      '<div class="gb-res gb-' + x[1] + '" aria-live="polite"><p class="gb-v"><b>' + esc(T(x[1] === 'b' ? 'gb.blocked' : x[1] === 'a' ? 'gb.allowed' : 'gb.depends')) + '</b> · ' + esc(T('gb.law')) + ': ' + esc(x[2]) + '</p><p>' + esc(T('gb.w.' + x[0])) + '</p>' +
      (x[1] !== 'a' ? '<p><b>' + esc(T('gb.exc')) + ':</b> ' + esc(T('gb.e.' + x[0])) + '</p>' : '') + '</div><p class="note">' + esc(T('gb.src')) + '</p>';
  }
  box.addEventListener('change', function (e) { if (e.target.id === 'gb-q') { cur = e.target.value; draw(); var s = box.querySelector('#gb-q'); if (s) s.focus(); } });
  document.addEventListener('ns-lang', draw);
  draw();
})();
