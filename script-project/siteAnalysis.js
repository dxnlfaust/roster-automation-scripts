/**
 * siteAnalysis.gs — turn the Calendar History tab into per-site planning data.
 *
 * Menu: Apunga Roster ▸ Build site analysis   (run "Build calendar history" first)
 *
 * Writes / refreshes:
 *   Calendar Analysis  one row per site: last financial year's hours and visits,
 *                      usual day, visit size, gaps, supervisor, hours per month.
 *                      A draft for management to confirm. Rebuilt every run.
 *   Calendar Unmatched calendar wording that matched no site, with its hours. Map
 *                      the ones that matter in "Calendar aliases", then rebuild.
 *   Profiles           monthly weights for the seasonal presets. Created once; only
 *                      the "Business pattern (last FY)" row is refreshed after that.
 *   Site Plan          annual hours + monthly grid per site. Created once; later runs
 *                      only add new sites and refresh "Last FY delivered". Typed
 *                      entries (contract dates, annual hours, profile, custom months)
 *                      are never overwritten.
 *
 * "Last FY" = the most recent complete July–June year in the history.
 */

var SA_HIST_TAB = 'Calendar History', SA_TAB = 'Calendar Analysis', SA_UNM_TAB = 'Calendar Unmatched';
var SA_PROF_TAB = 'Profiles', SA_PLAN_TAB = 'Site Plan';
var SA_MONTHS = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
// calendar wording for work that isn't one Sites row (shown as a group, not as unmatched)
var SA_GROUPS = { 'sqids': 'SQIDs (group)', 'syd park wetlands': 'Sydney Park Wetlands (group)',
  'sydpark wetlands': 'Sydney Park Wetlands (group)', 'sydney wetlands 3': 'Sydney Park Wetlands (group)',
  'glebe sites': 'Glebe (group)', 'glebe': 'Glebe (group)', 'shaes creek': 'Sheas Creek (not in Sites)',
  'sheas creek': 'Sheas Creek (not in Sites)', 'nursery': 'NURSERY', 'seed collecting': 'NURSERY' };
var SA_PROFILES = [   // placeholder weights until management confirm (Phase 5 plan doc)
  ['Even',                 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
  ['Spring/summer heavy',  0.6, 0.6, 1.4, 1.4, 1.4, 1.4, 1.4, 1.4, 1, 1, 1, 0.6],
  ['Winter heavy',         1.4, 1.4, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 1, 1, 1, 1.4]
];
var SA_PROFILE_CHOICES = ['From last year', 'Business pattern (last FY)', 'Even', 'Spring/summer heavy',
                          'Winter heavy', 'Custom'];

// ---------------------------------------------------------------------------------
// Pure analysis (testable outside Apps Script)
// ---------------------------------------------------------------------------------

function saFy_(iso) { var y = +iso.slice(0, 4), m = +iso.slice(5, 7); return m >= 7 ? y + 1 : y; }
function saMonIdx_(iso) { return (+iso.slice(5, 7) + 5) % 12; }               // Jul = 0 … Jun = 11
function saDays_(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 864e5); }
function saMedian_(xs) {
  if (!xs.length) return '';
  var s = xs.slice().sort(function (a, b) { return a - b; }), i = Math.floor(s.length / 2);
  return s.length % 2 ? s[i] : (s[i - 1] + s[i]) / 2;
}
function saQuant_(xs, q) {
  if (!xs.length) return '';
  var s = xs.slice().sort(function (a, b) { return a - b; }), p = (s.length - 1) * q, i = Math.floor(p);
  return i + 1 < s.length ? s[i] + (s[i + 1] - s[i]) * (p - i) : s[i];
}
function saMode_(xs) {
  var c = {}, best = null, n = 0;
  xs.forEach(function (x) { c[x] = (c[x] || 0) + 1; if (c[x] > n) { n = c[x]; best = x; } });
  return { value: best, count: n };
}

/** rows: Calendar History values (header first). today: 'yyyy-mm-dd'. */
function analyseHistory_(rows, today) {
  var H = rows[0].map(String), I = {}; H.forEach(function (h, i) { I[h] = i; });
  var recs = rows.slice(1).filter(function (r) { return r[I.Date]; }).map(function (r) {
    return { date: String(r[I.Date]).slice(0, 10), day: String(r[I.Day]), cal: String(r[I.Calendar]),
             kind: String(r[I.Kind]), site: String(r[I.Site]), text: String(r[I['Site text']]).trim(),
             hours: Number(r[I.Hours]) || 0, sup: String(r[I.Supervisor]), people: String(r[I.People]),
             matched: String(r[I.Matched]) === 'yes' };
  }).filter(function (r) { return r.date <= today; });

  // staff names: words seen 3+ times in crew lists
  var wc = {};
  recs.forEach(function (r) { if (r.kind === 'crew') r.people.split(/\s+/).forEach(function (w) { if (w) wc[w] = (wc[w] || 0) + 1; }); });
  var visits = [], crews = [];
  recs.forEach(function (r) {
    if (r.kind === 'crew' || (r.kind === 'visit' && !/\s/.test(r.text) && wc[r.text] >= 3)) crews.push(r);
    else if (r.kind === 'visit') visits.push(r);
  });
  var vkey = {}; visits.forEach(function (r) { vkey[r.date + '|' + r.cal] = true; });
  var unnamed = crews.filter(function (r) { return !vkey[r.date + '|' + r.cal]; });   // a crew with no site beside it

  visits.forEach(function (r) {
    var g = SA_GROUPS[r.text.toLowerCase()];
    r.plan = r.matched ? r.site : (g || '');
    r.how = r.matched ? 'matched' : g ? 'group' : 'unmatched';
    if (!r.sup) r.sup = r.people.split(/\s+/)[0] || '';
  });

  // last complete FY in the history
  var fyNow = saFy_(today), lastFy = fyNow - 1;
  var business = SA_MONTHS.map(function () { return 0; });
  visits.concat(unnamed).forEach(function (r) { if (saFy_(r.date) === lastFy) business[saMonIdx_(r.date)] += r.hours; });

  var bySite = {};
  visits.forEach(function (r) { if (r.plan && r.plan !== 'NURSERY') (bySite[r.plan] = bySite[r.plan] || []).push(r); });
  var sites = Object.keys(bySite).map(function (s) {
    var g = bySite[s].sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    var last = g.filter(function (r) { return saFy_(r.date) === lastFy; });
    var cur = g.filter(function (r) { return saFy_(r.date) === fyNow; });
    var dates = g.map(function (r) { return r.date; }).filter(function (d, i, a) { return a.indexOf(d) === i; });
    var gaps = dates.slice(1).map(function (d, i) { return saDays_(dates[i], d); });
    var day = saMode_(g.map(function (r) { return r.day; }));
    var sups = g.map(function (r) { return r.sup; }).filter(String), sup = saMode_(sups);
    var size = saMode_(g.map(function (r) { return Math.round(r.hours * 10) / 10; }));
    var hrs = g.map(function (r) { return r.hours; });
    var month = SA_MONTHS.map(function () { return 0; });
    last.forEach(function (r) { month[saMonIdx_(r.date)] += r.hours; });
    var hows = g.map(function (r) { return r.how; }).filter(function (h, i, a) { return a.indexOf(h) === i; });
    return {
      site: s, cal: saMode_(g.map(function (r) { return r.cal; })).value,
      lastHours: last.reduce(function (a, r) { return a + r.hours; }, 0), lastVisits: last.length,
      curHours: cur.reduce(function (a, r) { return a + r.hours; }, 0),
      day: day.value, dayShare: day.count / g.length, size: size.value,
      lo: saQuant_(hrs, 0.1), hi: saQuant_(hrs, 0.9),
      gap: saMedian_(gaps), maxGap: gaps.length ? Math.max.apply(null, gaps) : '',
      sup: sup.value || '', supShare: sups.length ? sup.count / sups.length : '',
      first: dates[0], last: dates[dates.length - 1], how: hows.join(', '),
      activeMonths: month.filter(function (h) { return h > 0; }).length, month: month
    };
  }).sort(function (a, b) { return b.lastHours - a.lastHours || a.site.localeCompare(b.site); });

  var un = {};
  visits.filter(function (r) { return r.how === 'unmatched'; }).forEach(function (r) {
    var k = r.text.toLowerCase() || '(no site name)';
    var u = un[k] = un[k] || { text: r.text || '(no site name)', hours: 0, lastHours: 0, n: 0, cals: [], last: '' };
    u.hours += r.hours; u.n++; u.cals.push(r.cal); if (r.date > u.last) u.last = r.date;
    if (saFy_(r.date) === lastFy) u.lastHours += r.hours;
  });
  var unmatched = Object.keys(un).map(function (k) { var u = un[k]; u.cal = saMode_(u.cals).value; return u; })
    .sort(function (a, b) { return b.hours - a.hours; });

  var cls = { matched: 0, group: 0, unmatched: 0, unnamed: 0, nursery: 0 };
  visits.forEach(function (r) {
    if (saFy_(r.date) !== lastFy) return;
    if (r.plan === 'NURSERY') cls.nursery += r.hours; else cls[r.how] += r.hours;
  });
  unnamed.forEach(function (r) { if (saFy_(r.date) === lastFy) cls.unnamed += r.hours; });
  return { lastFy: lastFy, sites: sites, unmatched: unmatched, business: business, coverage: cls };
}

// ---------------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------------

function buildSiteAnalysis() {
  var ss = SpreadsheetApp.getActive(), tz = Session.getScriptTimeZone();
  var hs = ss.getSheetByName(SA_HIST_TAB);
  if (!hs || hs.getLastRow() < 2) throw new Error('Run "Build calendar history" first.');
  var A = analyseHistory_(hs.getDataRange().getValues().map(function (r) {
    return r.map(function (c) { return c instanceof Date ? Utilities.formatDate(c, tz, 'yyyy-MM-dd') : c; });
  }), Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
  var fyLabel = 'FY' + String(A.lastFy - 1).slice(2) + '-' + String(A.lastFy).slice(2);   // FY25-26

  // ---- Calendar Analysis --------------------------------------------------------
  var head = ['Site', 'Main calendar', fyLabel + ' hours', fyLabel + ' visits', 'This FY so far (h)',
              'Usual day', 'Usual day share', 'Usual visit (h)', 'Typical range (h)', 'Typical gap (days)',
              'Longest gap (days)', 'Main supervisor', 'Supervisor share', 'First seen', 'Last visit',
              'Named as'].concat(SA_MONTHS.map(function (m) { return m + ' (' + fyLabel + ')'; }));
  var out = A.sites.map(function (s) {
    return [s.site, s.cal, s.lastHours, s.lastVisits, s.curHours, s.day, s.dayShare, s.size,
            s.lo === '' ? '' : Math.round(s.lo) + ' to ' + Math.round(s.hi), s.gap, s.maxGap, s.sup, s.supShare,
            s.first, s.last, s.how].concat(s.month);
  });
  var sh = saFresh_(ss, SA_TAB);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setWrap(true);
  if (out.length) {
    sh.getRange(2, 14, out.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, out.length, head.length).setValues(out);
    sh.getRange(2, 7, out.length, 1).setNumberFormat('0%');
    sh.getRange(2, 13, out.length, 1).setNumberFormat('0%');
    sh.getRange(2, 3, out.length, 3).setNumberFormat('#,##0');
    sh.getRange(2, 17, out.length, 12).setNumberFormat('#,##0');
  }
  sh.setFrozenRows(1); sh.setFrozenColumns(1); sh.setColumnWidth(1, 240);

  // ---- Calendar Unmatched -------------------------------------------------------
  var us = saFresh_(ss, SA_UNM_TAB);
  var uh = ['Calendar wording', fyLabel + ' hours', 'All hours', 'Visits', 'Main calendar', 'Last seen'];
  us.getRange(1, 1, 1, uh.length).setValues([uh]).setFontWeight('bold');
  if (A.unmatched.length) {
    us.getRange(2, 6, A.unmatched.length, 1).setNumberFormat('@');
    us.getRange(2, 1, A.unmatched.length, uh.length).setValues(A.unmatched.map(function (u) {
      return [u.text, u.lastHours, u.hours, u.n, u.cal, u.last];
    }));
  }
  us.setFrozenRows(1); us.setColumnWidth(1, 260);

  // ---- Profiles (create once; refresh the business row) ------------------------
  var mean = A.business.reduce(function (a, b) { return a + b; }, 0) / 12 || 1;
  var bizRow = ['Business pattern (last FY)'].concat(A.business.map(function (h) { return Math.round(h / mean * 100) / 100; }));
  var ps = ss.getSheetByName(SA_PROF_TAB);
  if (!ps) {
    ps = ss.insertSheet(SA_PROF_TAB);
    ps.getRange(1, 1, 1, 14).setValues([['Profile'].concat(SA_MONTHS).concat(['Notes'])]).setFontWeight('bold');
    var prof = [bizRow.concat(['Whole business, ' + fyLabel + ' calendar hours. Refreshed by Build site analysis.'])]
      .concat(SA_PROFILES.map(function (p) { return p.concat(['Placeholder weights: confirm with management']); }));
    ps.getRange(2, 1, prof.length, 14).setValues(prof);
    ps.getRange(prof.length + 3, 1, 2, 1).setValues([['"From last year" uses each site\'s own months in Calendar Analysis.'],
                                                       ['"Custom": type the hours straight into Site Plan.']]);
    ps.setFrozenRows(1); ps.setColumnWidth(1, 220);
  } else {
    var pv = ps.getRange(1, 1, ps.getLastRow(), 1).getValues();
    for (var i = 0; i < pv.length; i++) if (pv[i][0] === bizRow[0]) ps.getRange(i + 1, 1, 1, 13).setValues([bizRow]);
  }

  // ---- Site Plan (create once; add new sites; refresh Last FY) ----------------
  var plan = ss.getSheetByName(SA_PLAN_TAB), created = false;
  var ph = ['Site', 'Contract start', 'Contract end', 'Annual hours', 'Profile'].concat(SA_MONTHS)
    .concat(['Total', 'vs annual', 'Last FY delivered', 'Notes']);
  if (!plan) {
    plan = ss.insertSheet(SA_PLAN_TAB); created = true;
    plan.getRange(1, 1, 1, ph.length).setValues([ph]).setFontWeight('bold').setWrap(true);
    plan.setFrozenRows(1); plan.setFrozenColumns(1); plan.setColumnWidth(1, 240);
  }
  var existing = {}, last = plan.getLastRow();
  if (last > 1) plan.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { existing[r[0]] = i + 2; });
  var add = A.sites.filter(function (s) { return !existing[s.site]; });
  add.forEach(function (s, j) {
    var r = last + 1 + j;
    var f = '=IFERROR(IF(OR($D' + r + '="",$E' + r + '="",$E' + r + '="Custom"),"",LET(w,IF($E' + r +
      '="From last year",XLOOKUP($A' + r + ",'" + SA_TAB + "'!$A:$A,'" + SA_TAB + "'!$Q:$AB),XLOOKUP($E" + r +
      ',' + SA_PROF_TAB + '!$A:$A,' + SA_PROF_TAB + '!$B:$M)),ROUND($D' + r + '*w/SUM(w)))),"")';
    plan.getRange(r, 1, 1, 5).setValues([[s.site, '', '', '=T' + r, s.activeMonths >= 6 ? 'From last year' : 'Business pattern (last FY)']]);
    plan.getRange(r, 6).setFormula(f);
    plan.getRange(r, 18, 1, 2).setFormulas([['=SUM(F' + r + ':Q' + r + ')', '=IF(D' + r + '="","",R' + r + '-D' + r + ')']]);
  });
  // Last FY delivered: refreshed for every row (a value, so Annual hours can default to it)
  var lastByName = {}; A.sites.forEach(function (s) { lastByName[s.site] = s.lastHours; });
  var n = plan.getLastRow() - 1;
  if (n > 0) {
    var names = plan.getRange(2, 1, n, 1).getValues();
    plan.getRange(2, 20, n, 1).setValues(names.map(function (r) { return [lastByName[r[0]] || 0]; }));
    plan.getRange(2, 5, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(SA_PROFILE_CHOICES, true).setAllowInvalid(false).build());
    plan.getRange(2, 2, n, 2).setNumberFormat('d mmm yyyy');
    plan.getRange(2, 4, n, 17).setNumberFormat('#,##0');
  }

  var c = A.coverage, tot = c.matched + c.group + c.unmatched + c.unnamed + c.nursery;
  var pct = function (x) { return Math.round(100 * x / Math.max(1, tot)) + '%'; };
  SpreadsheetApp.getUi().alert('Site analysis (' + fyLabel + ')',
    A.sites.length + ' sites profiled. ' + fyLabel + ' calendar hours: ' + Math.round(tot).toLocaleString() + '\n\n' +
    '  matched to a site: ' + pct(c.matched) + '\n  site groups (SQIDs, wetlands, Glebe): ' + pct(c.group) +
    '\n  nursery: ' + pct(c.nursery) + '\n  wording not matched: ' + pct(c.unmatched) +
    ' (see Calendar Unmatched)\n  crew list with no site name: ' + pct(c.unnamed) +
    (created ? '\n\nSite Plan created. Annual hours start at last FY delivered until you type the real figure.'
             : '\n\nSite Plan: ' + add.length + ' new sites added; your entries were left alone.'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function saFresh_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); sh.clearConditionalFormatRules(); } else sh = ss.insertSheet(name);
  return sh;
}
