(function () {
  var $ = function (id) { return document.getElementById(id); };
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var state = { parsed: null, name: '', example: false, from: '', to: '', active: 0 };

  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function fd(iso) { if (!iso) return ''; var p = iso.split('-'); return p[2] + ' ' + MONTHS[+p[1] - 1] + ' ' + p[0]; }
  function fp(x) { return x == null ? '–' : (x * 100).toFixed(1) + '%'; }
  function fa(x) { return x == null ? '–' : String(+x.toFixed(2)); }

  /* ---------- example data (made up) ---------- */
  function exampleRows() {
    var heads = ['New submissions', 'Revised submissions', 'Resubmissions', 'Unsubmission', 'Accepted submissions', 'Transferred submissions', 'Authors', 'Editors', 'Reviewers', 'Other'];
    var w = 21, rows = [];
    var r0 = new Array(w).fill(null); r0[1] = 'Example Journal of Ecology'; r0[11] = 'Example Review of Optics';
    rows.push(r0); rows.push(new Array(w).fill(null));
    var r2 = ['DATE']; for (var j = 0; j < 2; j++) heads.forEach(function (h) { r2.push(h); }); rows.push(r2);
    var seed = 11; function rnd() { seed = (seed * 16807) % 2147483647; return seed / 2147483647; }
    var counter = [100, 400], pool = [[], []];
    for (var day = 1; day <= 6; day++) {
      var block = []; for (var i = 0; i < 6; i++) block.push(new Array(w).fill(null));
      block[0][0] = (day < 10 ? '0' : '') + day + '.11.2026';
      for (var jj = 0; jj < 2; jj++) {
        for (var ci = 0; ci < 10; ci++) {
          var count = ci === 0 ? 2 + Math.floor(rnd() * 3) : ci < 6 ? Math.floor(rnd() * 3) : Math.floor(rnd() * 4);
          for (var k = 0; k < count; k++) {
            var val;
            if (ci === 0) { counter[jj]++; val = (jj ? 'OPT' : 'ECO') + '-26-' + String(counter[jj]).padStart(5, '0'); pool[jj].push(val); }
            else if (ci < 6) { if (!pool[jj].length) continue; val = pool[jj][Math.floor(rnd() * pool[jj].length)]; }
            else { val = ['Status of manuscript', 'Reviewer reminder', 'Request for extension', 'Decision letter query'][Math.floor(rnd() * 4)] + ' ' + (1000 + Math.floor(rnd() * 8000)); }
            block[k][1 + jj * 10 + ci] = val;
          }
        }
      }
      block.forEach(function (b) { rows.push(b); });
    }
    return rows;
  }

  /* ---------- load ---------- */
  function showError(msg) { var e = $('err'); e.textContent = msg; e.hidden = !msg; }

  function setParsed(parsed, name, isExample) {
    state.parsed = parsed; state.name = name; state.example = isExample; state.active = 0;
    state.from = parsed.dates[0]; state.to = parsed.dates[parsed.dates.length - 1];
    $('dFrom').min = $('dTo').min = state.from; $('dFrom').max = $('dTo').max = state.to;
    renderAll();
  }

  function handleFile(file) {
    showError('');
    if (!file) return;
    var reader = new FileReader();
    reader.onerror = function () { showError('This file could not be read. Try saving it again as .xlsx.'); };
    reader.onload = function (ev) {
      try {
        if (typeof XLSX === 'undefined') throw new Error('The spreadsheet reader did not load. Check your connection and reload the page.');
        var wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        var lastErr = null, parsed = null;
        for (var i = 0; i < wb.SheetNames.length && !parsed; i++) {
          var ws = wb.Sheets[wb.SheetNames[i]];
          var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
          try { parsed = parseSheet(rows); } catch (e) { lastErr = e; }
        }
        if (!parsed) throw lastErr || new Error('No usable sheet found.');
        setParsed(parsed, file.name, false);
      } catch (e) { showError(e.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  /* ---------- render ---------- */
  function range() {
    var f = state.from, t = state.to;
    return { f: f || null, t: t || null };
  }

  function daysWithData(j, f, t) { return dailyRows(j, f, t).length; }

  function renderSource() {
    var p = state.parsed; if (!p) { $('source').innerHTML = ''; return; }
    var n = p.journals.length;
    $('source').innerHTML = (state.example ? '<span class="pill">Example data</span>' : '<span class="pill">Your sheet</span>') +
      '<span><b>' + esc(state.name) + '</b></span><span>' + n + ' journal' + (n === 1 ? '' : 's') + ' found</span>' +
      '<span>' + p.dates.length + ' dates, ' + fd(p.dates[0]) + ' to ' + fd(p.dates[p.dates.length - 1]) + '</span>' +
      (state.example ? '<span>Upload your own file to replace it.</span>' : '');
  }

  function renderOverview() {
    var p = state.parsed, r = range(), tot = { n: 0, r: 0, t: 0, s: 0, a: 0, u: 0, m: 0 };
    var body = p.journals.map(function (j, i) {
      var s = summarize(j, r.f, r.t), g = function (k) { return s.rows.filter(function (x) { return x.key === k; })[0]; };
      var vals = [g('new').unique, g('revised').unique, g('transferred').unique, g('resub').unique, s.acceptedAll, s.unsubAll, s.mailTotal];
      tot.n += vals[0]; tot.r += vals[1]; tot.t += vals[2]; tot.s += vals[3]; tot.a += vals[4]; tot.u += vals[5]; tot.m += vals[6];
      return '<tr><td><button class="link" data-open="' + i + '">' + esc(j.name) + '</button></td>' + vals.map(function (v) { return '<td>' + v + '</td>'; }).join('') + '</tr>';
    }).join('');
    $('overview').innerHTML = '<table><thead><tr><th>Journal</th><th>New</th><th>Revised</th><th>Transferred</th><th>Resubmissions (dedup.)</th><th>Accepted</th><th>Unsubmitted</th><th>E-mails</th></tr></thead><tbody>' + body +
      '<tr class="total"><td>All journals</td><td>' + tot.n + '</td><td>' + tot.r + '</td><td>' + tot.t + '</td><td>' + tot.s + '</td><td>' + tot.a + '</td><td>' + tot.u + '</td><td>' + tot.m + '</td></tr></tbody></table>';
  }

  function renderTabs() {
    var p = state.parsed;
    $('tabs').innerHTML = p.journals.map(function (j, i) {
      return '<button class="tab" role="tab" id="tab' + i + '" aria-selected="' + (i === state.active) + '" data-tab="' + i + '">' + esc(j.name) + '</button>';
    }).join('');
  }

  function pctCell(v, cls) {
    if (v == null) return '<td><span class="dash">–</span></td>';
    return '<td><span class="cell"><span>' + fp(v) + '</span><span class="bar ' + cls + '"><i style="width:' + Math.min(100, v * 100).toFixed(1) + '%"></i></span></span></td>';
  }

  function renderReport() {
    var p = state.parsed, j = p.journals[state.active], r = range();
    var s = summarize(j, r.f, r.t), days = dailyRows(j, r.f, r.t);
    var g = function (k) { return s.rows.filter(function (x) { return x.key === k; })[0]; };
    var kp = function (l, v, sub) { return '<div class="kpi"><div class="l">' + l + '</div><div class="v">' + v + '</div><div class="s">' + sub + '</div></div>'; };
    var h = '';
    h += '<div class="rhead"><div><h3>' + esc(j.name) + '</h3><div class="meta">' + fd(r.f) + ' to ' + fd(r.t) + ' · ' + days.length + ' day' + (days.length === 1 ? '' : 's') + ' with entries</div></div>' +
      '<button class="btn" id="copyOne" type="button">Copy this report</button></div>';
    if (s.total.unique + s.acceptedAll + s.unsubAll + s.mailTotal === 0) {
      $('report').innerHTML = h + '<div class="empty">No entries for this journal in the selected period.</div>';
      return;
    }
    h += '<div class="kpis">' +
      kp('New', g('new').unique, 'unique IDs') +
      kp('Revised', g('revised').unique, 'unique IDs') +
      kp('Transferred', g('transferred').unique, 'unique IDs') +
      kp('Resubmissions', g('resub').unique, 'deduplicated') +
      kp('Accepted', s.acceptedAll, fp(s.total.unique ? s.total.accepted / s.total.unique : null) + ' of submissions') +
      kp('Unsubmitted', s.unsubAll, fp(s.total.unique ? s.total.unsub / s.total.unique : null) + ' of submissions') + '</div>';

    h += '<div class="block"><h2>Submissions</h2><div class="tablewrap"><table><thead><tr><th>Category</th><th>Unique</th><th>Accepted</th><th>Accepted %</th><th>Unsubmitted</th><th>Unsubmitted %</th></tr></thead><tbody>';
    s.rows.forEach(function (x) {
      h += '<tr><td>' + esc(x.label) + '</td><td>' + x.unique + '</td><td>' + x.accepted + '</td>' + pctCell(x.acceptedPct, 'good') + '<td>' + x.unsub + '</td>' + pctCell(x.unsubPct, 'warn') + '</tr>';
    });
    h += '<tr class="total"><td>Total</td><td>' + s.total.unique + '</td><td>' + s.total.accepted + '</td><td>' + fp(s.total.acceptedPct) + '</td><td>' + s.total.unsub + '</td><td>' + fp(s.total.unsubPct) + '</td></tr>';
    h += '<tr class="soft"><td>Not matched to a submission in this period</td><td><span class="dash">–</span></td><td>' + s.unmatchedAccepted + '</td><td></td><td>' + s.unmatchedUnsub + '</td><td></td></tr>';
    h += '</tbody></table></div><p class="note">An ID is counted once per category. Accepted and unsubmitted IDs are matched to the category they were submitted under within the selected period. Resubmissions are IDs that do not also appear as new, revised or transferred. Unmatched IDs were submitted before the period or are missing from the sheet.</p></div>';

    h += '<div class="block"><h2>Resubmissions</h2><div class="tablewrap"><table><thead><tr><th>Resubmitted from</th><th>Unique in category</th><th>Times resubmitted</th><th>Unique IDs resubmitted</th><th>% resubmitted</th><th>Avg. per ID</th></tr></thead><tbody>';
    s.resub.from.forEach(function (x) {
      h += '<tr><td>' + esc(x.label) + '</td><td>' + x.unique + '</td><td>' + x.times + '</td><td>' + x.resubmittedUnique + '</td>' + pctCell(x.resubmittedPct, 'info') + '<td>' + fa(x.avg) + '</td></tr>';
    });
    h += '<tr><td>Resubmission only (no earlier record)</td><td>' + s.resub.standaloneUnique + '</td><td>' + s.resub.standaloneRows + '</td><td>' + s.resub.standaloneUnique + '</td><td><span class="dash">–</span></td><td>' + fa(s.resub.standaloneAvg) + '</td></tr>';
    h += '<tr class="total"><td>All resubmission entries</td><td></td><td>' + s.resub.totalRows + '</td><td></td><td></td><td></td></tr></tbody></table></div><p class="note">Times resubmitted counts every resubmission entry for IDs in that category. An ID that is listed in more than one category is counted under each, so the rows can add up to more than the total.</p></div>';

    h += '<div class="block"><h2>E-mails</h2><div class="tablewrap"><table><thead><tr><th>Sender</th><th>E-mails</th><th>Share</th></tr></thead><tbody>';
    s.mail.forEach(function (m) { h += '<tr><td>' + m.label + '</td><td>' + m.count + '</td>' + pctCell(m.share, 'info') + '</tr>'; });
    h += '<tr class="total"><td>Total</td><td>' + s.mailTotal + '</td><td></td></tr></tbody></table></div></div>';

    h += '<div class="block"><h2>Daily activity</h2>';
    if (!days.length) h += '<div class="empty">No daily entries in this period.</div>';
    else {
      h += '<div class="tablewrap"><table><thead><tr><th>Date</th><th>New</th><th>Revised</th><th>Resub.</th><th>Transferred</th><th>Accepted</th><th>Unsubmitted</th><th>E-mails</th></tr></thead><tbody>';
      days.forEach(function (d) { h += '<tr><td>' + fd(d.date) + '</td><td>' + d.new + '</td><td>' + d.revised + '</td><td>' + d.resub + '</td><td>' + d.transferred + '</td><td>' + d.accepted + '</td><td>' + d.unsub + '</td><td>' + d.mails + '</td></tr>'; });
      h += '</tbody></table></div><p class="note">Daily figures count unique IDs for that day, so they can add up to more than the period totals above.</p>';
    }
    h += '</div>';
    $('report').innerHTML = h;
  }

  function renderAll() {
    $('dFrom').value = state.from; $('dTo').value = state.to;
    var bad = state.from && state.to && state.from > state.to;
    $('rangeHint').textContent = bad ? 'The start date is after the end date. Swap them to see results.' :
      'All figures below use ' + fd(state.from) + ' to ' + fd(state.to) + '. Counts are unique IDs unless a column says otherwise.';
    renderSource(); renderOverview(); renderTabs(); renderReport();
  }

  /* ---------- copy as TSV ---------- */
  function tsvFor(j) {
    var r = range(), s = summarize(j, r.f, r.t), days = dailyRows(j, r.f, r.t), L = [];
    var P = function (x) { return x == null ? '' : (x * 100).toFixed(1) + '%'; };
    L.push(['Journal', j.name].join('\t')); L.push(['Period', fd(r.f) + ' to ' + fd(r.t)].join('\t')); L.push('');
    L.push('Submissions'); L.push(['Category', 'Unique', 'Accepted', 'Accepted %', 'Unsubmitted', 'Unsubmitted %'].join('\t'));
    s.rows.forEach(function (x) { L.push([x.label, x.unique, x.accepted, P(x.acceptedPct), x.unsub, P(x.unsubPct)].join('\t')); });
    L.push(['Total', s.total.unique, s.total.accepted, P(s.total.acceptedPct), s.total.unsub, P(s.total.unsubPct)].join('\t'));
    L.push(['Not matched in period', '', s.unmatchedAccepted, '', s.unmatchedUnsub, ''].join('\t')); L.push('');
    L.push('Resubmissions'); L.push(['Resubmitted from', 'Unique in category', 'Times resubmitted', 'Unique IDs resubmitted', '% resubmitted', 'Avg per ID'].join('\t'));
    s.resub.from.forEach(function (x) { L.push([x.label, x.unique, x.times, x.resubmittedUnique, P(x.resubmittedPct), x.avg == null ? '' : +x.avg.toFixed(2)].join('\t')); });
    L.push(['Resubmission only', s.resub.standaloneUnique, s.resub.standaloneRows, s.resub.standaloneUnique, '', s.resub.standaloneAvg == null ? '' : +s.resub.standaloneAvg.toFixed(2)].join('\t'));
    L.push(['All resubmission entries', '', s.resub.totalRows, '', '', ''].join('\t')); L.push('');
    L.push('E-mails'); L.push(['Sender', 'E-mails', 'Share'].join('\t'));
    s.mail.forEach(function (m) { L.push([m.label, m.count, P(m.share)].join('\t')); });
    L.push(['Total', s.mailTotal, ''].join('\t')); L.push('');
    L.push('Daily activity'); L.push(['Date', 'New', 'Revised', 'Resubmissions', 'Transferred', 'Accepted', 'Unsubmitted', 'E-mails'].join('\t'));
    days.forEach(function (d) { L.push([d.date, d.new, d.revised, d.resub, d.transferred, d.accepted, d.unsub, d.mails].join('\t')); });
    return L.join('\n');
  }

  var toastTimer;
  function toast(msg) { var t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 2600); }
  function copyText(text) {
    var fallback = function () {
      var ta = document.createElement('textarea'); ta.value = text; ta.setAttribute('readonly', '');
      ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0'; document.body.appendChild(ta); ta.select();
      var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta); return ok;
    };
    var done = function (ok) { toast(ok ? 'Copied. Paste into Excel with Ctrl+V.' : 'Copy was blocked. Select a table and copy it by hand.'); };
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { done(true); }, function () { done(fallback()); });
    else done(fallback());
  }

  /* ---------- events ---------- */
  $('fileInput').addEventListener('change', function (e) { handleFile(e.target.files[0]); e.target.value = ''; });
  $('chooseBtn').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileInput').click(); } });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (n) { drop.addEventListener(n, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (n) { drop.addEventListener(n, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { handleFile(e.dataTransfer.files[0]); });
  $('dFrom').addEventListener('change', function (e) { state.from = e.target.value; renderAll(); });
  $('dTo').addEventListener('change', function (e) { state.to = e.target.value; renderAll(); });
  $('resetRange').addEventListener('click', function () { if (!state.parsed) return; state.from = state.parsed.dates[0]; state.to = state.parsed.dates[state.parsed.dates.length - 1]; renderAll(); });
  $('copyAll').addEventListener('click', function () { if (state.parsed) copyText(state.parsed.journals.map(tsvFor).join('\n\n\n')); });
  $('tabs').addEventListener('click', function (e) { var b = e.target.closest('[data-tab]'); if (!b) return; state.active = +b.getAttribute('data-tab'); renderTabs(); renderReport(); });
  $('overview').addEventListener('click', function (e) { var b = e.target.closest('[data-open]'); if (!b) return; state.active = +b.getAttribute('data-open'); renderTabs(); renderReport(); $('rp-h').scrollIntoView({ behavior: 'smooth', block: 'start' }); });
  $('report').addEventListener('click', function (e) { if (e.target.id === 'copyOne') copyText(tsvFor(state.parsed.journals[state.active])); });

  setParsed(parseSheet(exampleRows()), 'Made-up example sheet', true);
})();