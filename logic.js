/* ---------- Parsing and metrics (pure functions) ---------- */
const SUB_CATS = ['new', 'revised', 'resub', 'unsub', 'accepted', 'transferred'];
const MAIL_CATS = ['authors', 'editors', 'reviewers', 'other'];

function catOf(t) {
  const s = String(t == null ? '' : t).toLowerCase().trim();
  if (!s) return null;
  if (/^un/.test(s)) return 'unsub';
  if (/^re-?submi/.test(s)) return 'resub';
  if (/^new/.test(s)) return 'new';
  if (/^revis/.test(s)) return 'revised';
  if (/accept/.test(s)) return 'accepted';
  if (/transfer/.test(s)) return 'transferred';
  if (/author/.test(s)) return 'authors';
  if (/editor/.test(s)) return 'editors';
  if (/reviewer/.test(s)) return 'reviewers';
  if (/^other/.test(s)) return 'other';
  return null;
}

function pad2(n) { return String(n).padStart(2, '0'); }

function parseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) {
    const d = +m[1], mo = +m[2];
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return m[3] + '-' + pad2(mo) + '-' + pad2(d);
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  return null;
}

function cleanCell(v, isId) {
  if (v == null) return null;
  let s = typeof v === 'number' ? String(Math.round(v) === v ? v : v) : String(v);
  s = s.trim();
  if (!s || /insert rows/i.test(s)) return null;
  return isId ? s.replace(/\s+/g, '').toUpperCase() : s;
}

function parseSheet(rows) {
  if (!rows || rows.length < 4) throw new Error('The sheet looks empty. Expected journal names in row 1, headings in row 3 and data below.');
  const r1 = rows[0] || [], r3 = rows[2] || [];
  let width = 0;
  rows.forEach(r => { if (r && r.length > width) width = r.length; });
  const starts = [];
  for (let c = 1; c < width; c++) {
    const v = r1[c];
    if (v != null && String(v).trim() !== '') starts.push({ name: String(v).trim(), col: c });
  }
  if (!starts.length) throw new Error('No journal names found in row 1. Put each journal name above its first column.');
  const colCat = {};
  let recognised = 0;
  for (let c = 1; c < width; c++) {
    const k = catOf(r3[c]);
    if (k) { colCat[c] = k; recognised++; }
  }
  if (!recognised) throw new Error('No column headings recognised in row 3 (New submissions, Revised submissions, Resubmissions, Unsubmission, Accepted submissions, Transferred submissions, Authors, Editors, Reviewers, Other).');

  const journals = starts.map((s, i) => ({
    name: s.name,
    from: s.col,
    to: (i + 1 < starts.length ? starts[i + 1].col : width) - 1,
    days: {}
  }));
  const seen = new Set();
  let cur = null;
  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] || [];
    const d = parseDate(row[0]);
    if (d) { cur = d; seen.add(d); }
    if (!cur) continue;
    journals.forEach(j => {
      for (let c = j.from; c <= j.to; c++) {
        const cat = colCat[c];
        if (!cat) continue;
        const val = cleanCell(row[c], SUB_CATS.indexOf(cat) >= 0);
        if (val == null) continue;
        const day = j.days[cur] || (j.days[cur] = {});
        (day[cat] || (day[cat] = [])).push(val);
      }
    });
  }
  const dates = [...seen].sort();
  if (!dates.length) throw new Error('No dates found in column A. Dates should look like 01.11.2026.');
  return { journals, dates };
}

function inRange(iso, from, to) { return (!from || iso >= from) && (!to || iso <= to); }

function collect(journal, from, to) {
  const out = {};
  SUB_CATS.concat(MAIL_CATS).forEach(k => out[k] = []);
  Object.keys(journal.days).sort().forEach(d => {
    if (!inRange(d, from, to)) return;
    const day = journal.days[d];
    Object.keys(day).forEach(k => { out[k] = out[k].concat(day[k]); });
  });
  return out;
}

function pct(a, b) { return b > 0 ? a / b : null; }

function summarize(journal, from, to) {
  const all = collect(journal, from, to);
  const S = {};
  ['new', 'revised', 'transferred', 'resub', 'unsub', 'accepted'].forEach(k => S[k] = new Set(all[k]));
  const resubDedup = new Set([...S.resub].filter(id => !S.new.has(id) && !S.revised.has(id) && !S.transferred.has(id)));
  const groups = [
    { key: 'new', label: 'New submissions', set: S.new },
    { key: 'revised', label: 'Revised submissions', set: S.revised },
    { key: 'transferred', label: 'Transferred submissions', set: S.transferred },
    { key: 'resub', label: 'Resubmissions (deduplicated)', set: resubDedup }
  ];
  const inter = (a, b) => [...a].filter(x => b.has(x)).length;
  const rows = groups.map(g => {
    const acc = inter(g.set, S.accepted), uns = inter(g.set, S.unsub);
    return { key: g.key, label: g.label, unique: g.set.size, accepted: acc, acceptedPct: pct(acc, g.set.size), unsub: uns, unsubPct: pct(uns, g.set.size) };
  });
  const matchedAcc = new Set(), matchedUns = new Set();
  groups.forEach(g => g.set.forEach(id => { if (S.accepted.has(id)) matchedAcc.add(id); if (S.unsub.has(id)) matchedUns.add(id); }));
  const tot = rows.reduce((t, r) => ({ unique: t.unique + r.unique, accepted: t.accepted + r.accepted, unsub: t.unsub + r.unsub }), { unique: 0, accepted: 0, unsub: 0 });
  tot.acceptedPct = pct(tot.accepted, tot.unique);
  tot.unsubPct = pct(tot.unsub, tot.unique);

  const resubFrom = ['new', 'revised', 'transferred'].map(k => {
    const times = all.resub.filter(id => S[k].has(id)).length;
    const uniq = inter(S[k], S.resub);
    return { key: k, label: groups.find(g => g.key === k).label, unique: S[k].size, times, resubmittedUnique: uniq,
      resubmittedPct: pct(uniq, S[k].size), avg: uniq ? times / uniq : null };
  });
  const standaloneRows = all.resub.filter(id => resubDedup.has(id)).length;
  const resub = {
    from: resubFrom,
    standaloneUnique: resubDedup.size,
    standaloneRows,
    standaloneAvg: resubDedup.size ? standaloneRows / resubDedup.size : null,
    totalRows: all.resub.length
  };
  const mail = MAIL_CATS.map(k => ({ key: k, label: k[0].toUpperCase() + k.slice(1), count: all[k].length }));
  const mailTotal = mail.reduce((t, m) => t + m.count, 0);
  mail.forEach(m => m.share = pct(m.count, mailTotal));
  return {
    rows, total: tot, resub, mail, mailTotal,
    unmatchedAccepted: [...S.accepted].filter(id => !matchedAcc.has(id)).length,
    unmatchedUnsub: [...S.unsub].filter(id => !matchedUns.has(id)).length,
    acceptedAll: S.accepted.size, unsubAll: S.unsub.size
  };
}

function dailyRows(journal, from, to) {
  return Object.keys(journal.days).sort().filter(d => inRange(d, from, to)).map(d => {
    const day = journal.days[d];
    const u = k => new Set(day[k] || []).size;
    const mails = MAIL_CATS.reduce((t, k) => t + (day[k] ? day[k].length : 0), 0);
    return { date: d, new: u('new'), revised: u('revised'), resub: u('resub'), transferred: u('transferred'), accepted: u('accepted'), unsub: u('unsub'), mails };
  }).filter(r => r.new + r.revised + r.resub + r.transferred + r.accepted + r.unsub + r.mails > 0);
}

if (typeof module !== 'undefined') module.exports = { parseSheet, summarize, dailyRows, parseDate, catOf };