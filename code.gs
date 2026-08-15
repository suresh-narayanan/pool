/**
 * 8-Ball Tournament — Google Sheets backend
 * ------------------------------------------------------------------
 * Paste into Extensions ▸ Apps Script on a Google Sheet, change the PIN,
 * then Deploy ▸ New deployment ▸ Web app
 *   Execute as:      Me
 *   Who has access:  Anyone      ← must be "Anyone", not "Anyone with Google account"
 * Copy the /exec URL into index.html (APP_CONFIG.API_URL).
 *
 * Tabs created in the sheet:
 *   _state     JSON source of truth (don't edit by hand)
 *   Standings  group tables, rewritten on every save
 *   Matches    every group game with balls remaining
 *   Knockout   the bracket: quarters, semis, final, 3rd place
 * ------------------------------------------------------------------
 */

// ====== CHANGE THIS ======
var PIN = '7070';
// =========================

var LETTERS = ['A', 'B', 'C', 'D'];
var GBO = 3;   // group matches: best of 3
var KBO = 5;   // semis, final, 3rd place: best of 5
var KO_ORDER = ['qf1', 'qf2', 'qf3', 'qf4', 'sf1', 'sf2', 'f', 'tp'];

/* ---------- entry point (JSONP over GET) ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try { out = handle(p); }
  catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  var body = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) { return doGet(e); }

function handle(p) {
  var action = p.action || 'state';

  if (action === 'state') return { ok: true, state: readState() };
  if (action === 'checkpin')
    return String(p.pin) === String(PIN) ? { ok: true } : { ok: false, error: 'Wrong PIN' };

  if (String(p.pin) !== String(PIN)) return { ok: false, error: 'Wrong PIN' };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var st = readState();

    if (action === 'setup') {
      var groups = JSON.parse(p.groups || '[]');
      if (!groups.length) throw new Error('No groups supplied');
      st = { v: 3, title: p.title || '8-Ball Tournament', groups: groups,
             results: {}, ko: {}, created: new Date().toISOString() };

    } else if (action === 'match') {
      if (!p.key) throw new Error('No match key');
      var gs = JSON.parse(p.games || '[]');
      st.results = st.results || {};
      if (gs.length) st.results[p.key] = gs; else delete st.results[p.key];

    } else if (action === 'ko') {
      if (!p.key || KO_ORDER.indexOf(p.key) < 0) throw new Error('Unknown knockout match');
      var kgs = JSON.parse(p.games || '[]');
      var players = JSON.parse(p.p || '[]');
      if (players.length !== 2) throw new Error('Knockout match needs two players');
      st.ko = st.ko || {};
      // The players are stored with the result. If an earlier round is later
      // changed, this record no longer matches the bracket and is ignored —
      // no stale scores, and nothing is silently deleted either.
      if (kgs.length) st.ko[p.key] = { p: players, gs: kgs }; else delete st.ko[p.key];

    } else if (action === 'reset') {
      st = { v: 3, title: '8-Ball Tournament', groups: [], results: {}, ko: {} };

    } else {
      throw new Error('Unknown action: ' + action);
    }

    st.updated = new Date().toISOString();
    writeState(st);
    try { mirror(st); } catch (e) { /* a mirror failure must never lose a score */ }
    return { ok: true, state: st };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- storage ---------- */
function book() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tab(name) {
  var ss = book(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function readState() {
  var raw = tab('_state').getRange('A1').getValue();
  if (!raw) return { v: 3, title: '8-Ball Tournament', groups: [], results: {}, ko: {} };
  try {
    var st = JSON.parse(raw);
    st.groups = st.groups || [];
    st.results = st.results || {};
    st.ko = st.ko || {};
    return st;
  } catch (e) {
    return { v: 3, title: '8-Ball Tournament', groups: [], results: {}, ko: {} };
  }
}
function writeState(st) {
  var sh = tab('_state');
  sh.getRange('A1').setValue(JSON.stringify(st));
  sh.getRange('A2').setValue('Do not edit this tab by hand — it is the app\'s source of truth.');
  sh.getRange('A3').setValue('Last updated: ' + new Date());
}

/* ---------- shared scoring logic ---------- */
function pairsOf(n) {
  var o = [];
  for (var d = 1; d < n; d++) for (var i = 0; i + d < n; i++) o.push([i, i + d]);
  return o;
}
function needWins(bo) { return Math.ceil(bo / 2); }
function tally(gs) {
  var a = 0, b = 0;
  (gs || []).forEach(function (x) { if (!x) return; if (x.w === 0) a++; else b++; });
  return [a, b];
}
function decided(gs, bo) { var t = tally(gs); return t[0] >= needWins(bo) || t[1] >= needWins(bo); }
function gamesFor(st, g, i, j) {
  var r = st.results[g + '-' + i + '-' + j];
  return Array.isArray(r) ? r : [];
}

var RKEYS = [
  function (r) { return -r.pts; },
  function (r) { return -(r.gw - r.gl); },
  function (r) { return -(r.bf - r.ba); },
  function (r) { return r.ba; }
];
function rankRows(rows, h2h) {
  return split(rows, 0);
  function split(arr, k) {
    if (arr.length <= 1) return arr;
    if (arr.length === 2 && k >= 2) {
      var h = h2h[arr[0].n + '|' + arr[1].n];
      if (h === 1) return [arr[0], arr[1]];
      if (h === -1) return [arr[1], arr[0]];
    }
    if (k >= RKEYS.length)
      return arr.slice().sort(function (a, b) { return a.n < b.n ? -1 : (a.n > b.n ? 1 : 0); });
    var f = RKEYS[k], bk = {}, ord = [];
    arr.forEach(function (r) { var v = f(r); if (!(v in bk)) { bk[v] = []; ord.push(v); } bk[v].push(r); });
    ord.sort(function (a, b) { return a - b; });
    var out = [];
    ord.forEach(function (v) { out = out.concat(split(bk[v], k + 1)); });
    return out;
  }
}
function standings(st, gi) {
  var names = st.groups[gi] || [];
  var rows = names.map(function (n) {
    return { n: n, p: 0, w: 0, l: 0, gw: 0, gl: 0, bf: 0, ba: 0, pts: 0 };
  });
  var h2h = {};
  pairsOf(names.length).forEach(function (pr) {
    var gs = gamesFor(st, gi, pr[0], pr[1]);
    if (!decided(gs, GBO)) return;
    var A = rows[pr[0]], B = rows[pr[1]];
    A.p++; B.p++;
    gs.forEach(function (x) {
      if (!x) return;
      var bl = (typeof x.b === 'number') ? x.b : 0;
      if (x.w === 0) { A.gw++; B.gl++; A.bf += bl; B.ba += bl; }
      else { B.gw++; A.gl++; B.bf += bl; A.ba += bl; }
    });
    var t = tally(gs);
    if (t[0] > t[1]) { A.w++; A.pts++; B.l++; h2h[A.n + '|' + B.n] = 1; h2h[B.n + '|' + A.n] = -1; }
    else { B.w++; B.pts++; A.l++; h2h[B.n + '|' + A.n] = 1; h2h[A.n + '|' + B.n] = -1; }
  });
  return rankRows(rows, h2h);
}
function groupsDone(st) {
  if (!st.groups || st.groups.length < 4) return false;
  for (var g = 0; g < 4; g++) {
    var names = st.groups[g], ps = pairsOf(names.length);
    for (var i = 0; i < ps.length; i++)
      if (!decided(gamesFor(st, g, ps[i][0], ps[i][1]), GBO)) return false;
  }
  return true;
}

/* ---------- knockout ---------- */
function kgames(st, tree, key) {
  var rec = (st.ko || {})[key], m = tree[key];
  if (!rec || !m || !rec.p) return [];
  if (rec.p[0] !== m.p[0] || rec.p[1] !== m.p[1]) return [];   // stale, bracket moved on
  return Array.isArray(rec.gs) ? rec.gs : [];
}
function kdone(st, tree, key) {
  var m = tree[key];
  return !!(m.p[0] && m.p[1]) && decided(kgames(st, tree, key), m.bo);
}
function kwin(st, tree, key) {
  if (!kdone(st, tree, key)) return null;
  var t = tally(kgames(st, tree, key));
  return t[0] > t[1] ? tree[key].p[0] : tree[key].p[1];
}
function klose(st, tree, key) {
  if (!kdone(st, tree, key)) return null;
  var t = tally(kgames(st, tree, key));
  return t[0] > t[1] ? tree[key].p[1] : tree[key].p[0];
}
function koTree(st) {
  if (!groupsDone(st)) return null;
  var q = [standings(st, 0), standings(st, 1), standings(st, 2), standings(st, 3)];
  function nm(gi, ix) { return (q[gi][ix] || {}).n || null; }
  var t = {};
  t.qf1 = { p: [nm(0, 0), nm(2, 1)], bo: GBO, lab: 'Quarter-final 1', note: 'A1 v C2' };
  t.qf2 = { p: [nm(2, 0), nm(0, 1)], bo: GBO, lab: 'Quarter-final 2', note: 'C1 v A2' };
  t.qf3 = { p: [nm(1, 0), nm(3, 1)], bo: GBO, lab: 'Quarter-final 3', note: 'B1 v D2' };
  t.qf4 = { p: [nm(3, 0), nm(1, 1)], bo: GBO, lab: 'Quarter-final 4', note: 'D1 v B2' };
  t.sf1 = { p: [kwin(st, t, 'qf1'), kwin(st, t, 'qf4')], bo: KBO, lab: 'Semi-final 1', note: 'QF1 v QF4 winners' };
  t.sf2 = { p: [kwin(st, t, 'qf2'), kwin(st, t, 'qf3')], bo: KBO, lab: 'Semi-final 2', note: 'QF2 v QF3 winners' };
  t.f   = { p: [kwin(st, t, 'sf1'), kwin(st, t, 'sf2')], bo: KBO, lab: 'Final', note: 'SF winners' };
  t.tp  = { p: [klose(st, t, 'sf1'), klose(st, t, 'sf2')], bo: KBO, lab: '3rd place playoff', note: 'SF losers' };
  return t;
}

/* ---------- readable mirror tabs ---------- */
function mirror(st) {
  // Standings
  var s = tab('Standings');
  s.clear();
  var out = [['Group', 'Pos', 'Player', 'P', 'W', 'L', 'Games won', 'Games lost',
              'Game +/-', 'Balls for', 'Balls against', 'Ball +/-', 'Points', 'Qualified']];
  for (var gi = 0; gi < (st.groups || []).length; gi++) {
    standings(st, gi).forEach(function (r, ix) {
      out.push([LETTERS[gi], ix + 1, r.n, r.p, r.w, r.l, r.gw, r.gl,
                r.gw - r.gl, r.bf, r.ba, r.bf - r.ba, r.pts,
                (groupsDone(st) && ix < 2) ? 'YES' : '']);
    });
  }
  writeBlock(s, out);

  // Matches
  var m = tab('Matches');
  m.clear();
  var head = ['Group', 'Player A', 'Player B', 'Score', 'Status'];
  for (var k = 1; k <= GBO; k++) head.push('G' + k + ' winner', 'G' + k + ' balls left');
  var mo = [head];
  for (var g = 0; g < (st.groups || []).length; g++) {
    var names = st.groups[g];
    pairsOf(names.length).forEach(function (pr) {
      var A = names[pr[0]], B = names[pr[1]];
      var gs = gamesFor(st, g, pr[0], pr[1]), t = tally(gs);
      var row = [LETTERS[g], A, B, gs.length ? (t[0] + '-' + t[1]) : '',
                 decided(gs, GBO) ? 'Final' : (gs.length ? 'In progress' : 'Not played')];
      for (var q = 0; q < GBO; q++) {
        var x = gs[q];
        row.push(x ? (x.w === 0 ? A : B) : '');
        row.push(x ? (typeof x.b === 'number' ? x.b : '') : '');
      }
      mo.push(row);
    });
  }
  writeBlock(m, mo);

  // Knockout
  var kSh = tab('Knockout');
  kSh.clear();
  var kh = ['Round', 'Player A', 'Player B', 'Score', 'Winner', 'Status'];
  for (var k2 = 1; k2 <= KBO; k2++) kh.push('G' + k2 + ' winner', 'G' + k2 + ' balls left');
  var ko = [kh];
  var tree = koTree(st);
  if (!tree) {
    ko.push(['Knockout opens once every group match is played', '', '', '', '', '']);
  } else {
    KO_ORDER.forEach(function (key) {
      var mm = tree[key], gs = kgames(st, tree, key), t = tally(gs);
      var row = [mm.lab, mm.p[0] || 'TBD', mm.p[1] || 'TBD',
                 gs.length ? (t[0] + '-' + t[1]) : '',
                 kwin(st, tree, key) || '',
                 (!mm.p[0] || !mm.p[1]) ? 'Awaiting players'
                   : (kdone(st, tree, key) ? 'Final' : (gs.length ? 'In progress' : 'Not played'))];
      for (var q2 = 0; q2 < KBO; q2++) {
        var x = gs[q2];
        row.push(x ? (x.w === 0 ? mm.p[0] : mm.p[1]) : '');
        row.push(x ? (typeof x.b === 'number' ? x.b : '') : '');
      }
      ko.push(row);
    });
    var champ = kwin(st, tree, 'f');
    if (champ) {
      ko.push([]);
      ko.push(['RESULT', '1st', champ, '', '', '']);
      ko.push(['RESULT', '2nd', klose(st, tree, 'f') || '', '', '', '']);
      ko.push(['RESULT', '3rd', kwin(st, tree, 'tp') || '', '', '', '']);
    }
  }
  writeBlock(kSh, ko);
}

function writeBlock(sh, rows) {
  var w = 0;
  rows.forEach(function (r) { if (r.length > w) w = r.length; });
  var norm = rows.map(function (r) {
    var c = r.slice();
    while (c.length < w) c.push('');
    return c;
  });
  sh.getRange(1, 1, norm.length, w).setValues(norm);
  sh.getRange(1, 1, 1, w).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, w);
}

/* ---------- run from the editor to sanity-check ---------- */
function selfTest() {
  var st = { v: 3, title: 'Test', results: {}, ko: {},
             groups: [['A1', 'A2', 'A3', 'A4'], ['B1', 'B2', 'B3', 'B4'],
                      ['C1', 'C2', 'C3', 'C4'], ['D1', 'D2', 'D3', 'D4']] };
  // give every group a full set of results
  for (var g = 0; g < 4; g++) {
    pairsOf(4).forEach(function (pr) {
      st.results[g + '-' + pr[0] + '-' + pr[1]] = [{ w: 0, b: 3 }, { w: 0, b: 2 }];
    });
  }
  Logger.log('groups complete: ' + groupsDone(st));
  var t = koTree(st);
  Logger.log('QF1 ' + t.qf1.p.join(' v ') + ' | QF2 ' + t.qf2.p.join(' v ') +
             ' | QF3 ' + t.qf3.p.join(' v ') + ' | QF4 ' + t.qf4.p.join(' v '));
}
