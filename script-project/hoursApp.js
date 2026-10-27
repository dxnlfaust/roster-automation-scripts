/**
 * HoursApp.gs — end-of-day hours logging (server side).
 *
 * Supervisors open a phone-friendly web page, pick their site, and confirm or
 * adjust each person's hours. The crew is pre-filled from the Master Roster, so
 * the common case ("everyone did their day") is two taps.
 *
 * Writes to the "Hours Log" tab (append-only — the newest submission for a
 * date+site wins), and "Hours Variance" summarises scheduled vs actual.
 *
 * DEPLOYING (see SETUP_HOURS_APP.md):
 *   Apps Script editor > Deploy > New deployment > Web app
 *     Execute as:      Me
 *     Who has access:  Anyone
 *   Copy the /exec URL and give it to supervisors. No sign-in required.
 *
 * NOTE: "Anyone" means anyone with the link can submit. That is a deliberate
 * trade for adoption. Submissions are rejected unless they match a real visit
 * on the Master Roster, every row records who claimed to submit it, and nothing
 * is ever deleted — so a bad entry is visible and reversible.
 */

var HOURS_TAB = 'Hours Log';
var VARIANCE_TAB = 'Hours Variance';
var HRS_SUP = 8;        // a supervisor's standard day
var HRS_CREW = 7.5;     // everyone else
var REASONS = ['Full day', 'Late', 'Left early', 'Sick', 'No show', 'Rained out',
               'Extra hours', 'Other'];

function doGet() {
  return HtmlService.createTemplateFromFile('HoursForm')
    .evaluate()
    .setTitle('Apunga — log hours')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Master Roster as visit objects. Shared with the client. */
function hoursVisits_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DATA_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0].map(String);
  var c = function (n) { return hdr.indexOf(n); };
  var staffCols = hdr.map(function (h, i) {
    return /^Staff\s*\d*$/i.test(String(h).trim()) ? i : -1;
  }).filter(function (i) { return i >= 0; });
  var tz = Session.getScriptTimeZone();
  var out = [];
  vals.slice(1).forEach(function (r) {
    var raw = r[c('Date')];
    if (!raw || !r[c('Site')]) return;
    var date = (raw instanceof Date)
      ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
      : isoFromUs_(String(raw).trim());
    var crew = [];
    var sup = String(r[c('Supervisor')] || '').trim();
    var atm = c('ATM') >= 0 ? String(r[c('ATM')] || '').trim() : '';
    if (sup) crew.push({ name: sup, role: 'Supervisor', sched: HRS_SUP });
    if (atm) crew.push({ name: atm, role: 'ATM', sched: HRS_CREW });
    staffCols.forEach(function (i) {
      var nm = String(r[i] || '').trim();
      if (nm) crew.push({ name: nm, role: 'Crew', sched: HRS_CREW });
    });
    out.push({
      date: date, site: String(r[c('Site')]).trim(),
      siteHours: c('Site Hours') >= 0 ? String(r[c('Site Hours')] || '') : '',
      notes: c('Notes') >= 0 ? String(r[c('Notes')] || '') : '',
      supervisor: sup, crew: crew
    });
  });
  return out;
}

function isoFromUs_(s) {
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);   // M/D/YYYY as written by the optimiser
  if (!m) return s.slice(0, 10);
  return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
}

/** Everything the form needs in one round trip. */
function hoursBootstrap() {
  var visits = hoursVisits_();
  var dates = {};
  visits.forEach(function (v) { dates[v.date] = 1; });
  return {
    visits: visits,
    dates: Object.keys(dates).sort(),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    allStaff: allStaffIds_(),
    reasons: REASONS,
    submitted: submittedKeys_(),
    summary: (function () {
      try { return summaryBootstrap_(); }
      catch (e) { return { councils: {}, lastSpecies: {}, summarised: {} }; }
    })()
  };
}

function allStaffIds_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Staff');
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  var hr = -1;
  for (var r = 0; r < Math.min(v.length, 15); r++) {
    if (v[r].map(String).map(function (x) { return x.trim(); }).indexOf('StaffID') >= 0) { hr = r; break; }
  }
  if (hr < 0) return [];
  var i = v[hr].map(String).map(function (x) { return x.trim(); }).indexOf('StaffID');
  var out = [];
  for (var r2 = hr + 1; r2 < v.length; r2++) {
    var id = String(v[r2][i] || '').trim();
    if (id && id.indexOf('Wk ') !== 0 && id.indexOf('TOTAL') !== 0) out.push(id);
  }
  return out.sort();
}

/** date|site keys already logged, so the form can show what's done. */
function submittedKeys_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(HOURS_TAB);
  if (!sh || sh.getLastRow() < 2) return {};
  var v = sh.getDataRange().getValues(), hdr = v[0].map(String);
  var iD = hdr.indexOf('Date'), iS = hdr.indexOf('Site'), iB = hdr.indexOf('Submitted by');
  var tz = Session.getScriptTimeZone();
  var out = {};
  v.slice(1).forEach(function (r) {
    if (!r[iD] || !r[iS]) return;
    var d = (r[iD] instanceof Date)
      ? Utilities.formatDate(r[iD], tz, 'yyyy-MM-dd') : String(r[iD]).trim();
    out[d + '|' + String(r[iS]).trim()] = String(r[iB] || '');
  });
  return out;
}

function hoursSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(HOURS_TAB);
  if (!sh) {
    sh = ss.insertSheet(HOURS_TAB);
    sh.getRange(1, 1, 1, 12).setValues([['Timestamp', 'Date', 'Site', 'Submitted by', 'Staff',
      'Role', 'Scheduled', 'Actual', 'Rope hrs', 'Chainsaw hrs', 'Reason', 'Note']]);
    sh.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#212121').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  // Sheets will happily turn '2026-08-05' into a Date, which then no longer matches
  // the roster's yyyy-MM-dd keys. Pin the Date column to plain text.
  try { sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@'); } catch (e) {}
  return sh;
}

/**
 * Called from the form. Validates against the Master Roster, then appends.
 * Returns a short confirmation string.
 */
function submitHours(payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Busy — please try again in a moment.');
  try {
    if (!payload || !payload.date || !payload.site) throw new Error('Missing date or site.');
    if (!payload.submittedBy) throw new Error('Please choose your name.');
    if (!payload.rows || !payload.rows.length) throw new Error('No staff to record.');

    var match = hoursVisits_().filter(function (v) {
      return v.date === payload.date && v.site === payload.site;
    });
    if (!match.length) throw new Error('That site is not on the roster for that date.');

    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var stamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    var rows = payload.rows.map(function (r) {
      var actual = Number(r.actual);
      if (isNaN(actual) || actual < 0 || actual > 24) actual = 0;
      return [stamp, payload.date, payload.site, payload.submittedBy,
              String(r.name || '').trim(), String(r.role || 'Crew'),
              Number(r.sched) || 0, actual,
              Number(r.rope) || 0, Number(r.chainsaw) || 0,
              String(r.reason || ''), String(r.note || '').trim()];
    });
    var sh = hoursSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);

    var total = rows.reduce(function (a, r) { return a + r[7]; }, 0);
    return rows.length + ' staff logged, ' + total + ' hours total.';
  } finally {
    lock.releaseLock();
  }
}

// ---------- variance report --------------------------------------------------------

/**
 * Compares the newest submission per date+site against the Master Roster.
 * Menu: Apunga Roster > Build hours variance report.
 */
function buildHoursVariance() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName(HOURS_TAB);
  if (!log || log.getLastRow() < 2) throw new Error('No submissions in "' + HOURS_TAB + '" yet.');

  var v = log.getDataRange().getValues(), hdr = v[0].map(String);
  var c = function (n) { return hdr.indexOf(n); };
  var tz = Session.getScriptTimeZone();
  // Dates may come back as strings OR as Date objects, depending on how the cell
  // was formatted. Normalise both to yyyy-MM-dd before matching.
  var asIso = function (x) {
    if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    var s = String(x).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    var d = new Date(s);
    return isNaN(d.getTime()) ? s : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  };

  // newest submission wins for each date+site
  var latest = {};
  v.slice(1).forEach(function (r) {
    var key = asIso(r[c('Date')]) + '|' + String(r[c('Site')]).trim();
    var ts = String(r[c('Timestamp')]);
    if (!latest[key] || ts > latest[key].ts) latest[key] = { ts: ts, rows: [] };
    if (latest[key].ts === ts) latest[key].rows.push(r);
  });

  var sched = {};
  hoursVisits_().forEach(function (vis) {
    sched[vis.date + '|' + vis.site] = vis;
  });

  var out = [['Date', 'Site', 'Submitted by', 'Staff', 'Scheduled', 'Actual', 'Difference',
              'Rope hrs', 'Chainsaw hrs', 'Reason', 'Note']];
  var flags = [];
  Object.keys(latest).sort().forEach(function (key) {
    var entry = latest[key], parts = key.split('|');
    var siteSched = 0, siteActual = 0;
    entry.rows.forEach(function (r) {
      var s = Number(r[c('Scheduled')]) || 0, a = Number(r[c('Actual')]) || 0;
      siteSched += s; siteActual += a;
      out.push([parts[0], parts[1], r[c('Submitted by')], r[c('Staff')], s, a,
                Math.round((a - s) * 100) / 100, r[c('Rope hrs')], r[c('Chainsaw hrs')],
                r[c('Reason')], r[c('Note')]]);
    });
    out.push([parts[0], parts[1], '', 'SITE TOTAL', siteSched, siteActual,
              Math.round((siteActual - siteSched) * 100) / 100, '', '', '', '']);
    if (Math.abs(siteActual - siteSched) > 0.01) flags.push(out.length);
  });

  // rostered visits with nothing logged
  Object.keys(sched).sort().forEach(function (key) {
    if (!latest[key]) {
      var p = key.split('|');
      out.push([p[0], p[1], '', 'NOT SUBMITTED', '', '', '', '', '', '', '']);
    }
  });

  var sh = ss.getSheetByName(VARIANCE_TAB);
  if (sh) sh.clear(); else sh = ss.insertSheet(VARIANCE_TAB);
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold')
    .setBackground('#212121').setFontColor('#ffffff');
  for (var i = 2; i <= out.length; i++) {
    var label = String(out[i - 1][3]);
    if (label === 'SITE TOTAL') sh.getRange(i, 1, 1, out[0].length).setFontWeight('bold')
      .setBackground('#f1f3f4');
    if (label === 'NOT SUBMITTED') sh.getRange(i, 1, 1, out[0].length)
      .setBackground('#fce8e6').setFontWeight('bold');
    var d = out[i - 1][6];
    if (typeof d === 'number' && Math.abs(d) > 0.01) {
      sh.getRange(i, 7).setBackground(d < 0 ? '#fce8e6' : '#e6f4ea').setFontWeight('bold');
    }
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 220);
  ss.setActiveSheet(sh);
  ss.toast('Variance report built: ' + (out.length - 1) + ' rows.', 'Hours', 8);
}

/** Shows the live web app URL so you can copy it to the supervisors. */
function showHoursAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = 'ERROR: ' + e.message; }
  var msg = (url && url.indexOf('ERROR') !== 0)
    ? 'Web app URL:\n\n' + url + '\n\nSend this to supervisors.'
    : 'No web app URL yet. Deploy > New deployment > Web app (Execute as: Me, ' +
      'Who has access: Anyone), then run this again.' +
      (url ? '\n\n' + url : '');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Hours logging', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { /* no UI context — the log has it */ }
}

/**
 * Read-only self-test. Run this from the editor (Run > checkHoursApp) and read
 * the Execution log. It checks each thing that can go wrong, in order.
 */
function checkHoursApp() {
  var out = [];

  // 1. is the HTML file actually inside this Apps Script project?
  try {
    var t = HtmlService.createTemplateFromFile('HoursForm').evaluate().getContent();
    out.push('HoursForm file: FOUND inside the script project (' + t.length + ' chars)');
  } catch (e) {
    out.push('HoursForm file: *** MISSING *** — in the Apps Script editor click + next to '
           + 'Files, choose HTML, name it exactly HoursForm (no .html), and paste the '
           + 'contents in. Uploading the .html to Drive does NOT work. [' + e.message + ']');
  }

  // 2. deployment
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; out.push('getUrl failed: ' + e.message); }
  out.push(url
    ? 'Web app URL: ' + url + (url.indexOf('/exec') > 0 ? '  (looks correct)' :
        '  *** this is not an /exec URL — deploy as a Web app ***')
    : 'Web app URL: *** none *** — Deploy > New deployment > Web app');

  // 3. the data the form needs
  try {
    var v = hoursVisits_();
    out.push('Master Roster: ' + v.length + ' visits readable'
             + (v.length ? ' (first: ' + v[0].date + ' ' + v[0].site + ')' : ''));
    if (!v.length) out.push('  -> the form will show no sites. Run the optimiser first.');
  } catch (e) { out.push('Master Roster: ERROR ' + e.message); }

  try { out.push('Staff list: ' + allStaffIds_().length + ' names'); }
  catch (e) { out.push('Staff list: ERROR ' + e.message); }

  try { out.push('Hours Log tab: ' + (SpreadsheetApp.getActive().getSheetByName(HOURS_TAB)
        ? 'exists' : 'not created yet (normal until the first submission)')); }
  catch (e) { out.push('Hours Log: ERROR ' + e.message); }

  var msg = out.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Hours app check', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) {}
  return msg;
}