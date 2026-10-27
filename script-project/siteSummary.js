/**
 * SiteSummary.gs — the daily site summary, stage 2 of the supervisor form.
 *
 * Storage is deliberately split:
 *   Hours Log     one row per PERSON  (payroll)
 *   Site Summary  one row per VISIT   (reporting, invoicing, monthly reports)
 *
 * Keeping visit-level data out of the person-level rows is what makes the
 * monthly report straightforward later — species, works description and WHS
 * belong to the visit, not to each individual.
 *
 * The summary is OPTIONAL. Hours save first and independently, so a supervisor
 * who loses signal or gets interrupted has already banked the payroll data.
 */

var SUMMARY_TAB = 'Site Summary';

var SUMMARY_COLS = ['Timestamp', 'Date', 'Site', 'Council / Land Manager', 'Submitted by',
  'Other supervisor', 'Maintenance hrs', 'Primary hrs', 'Total hrs',
  'Target species', 'Other species', 'Herbicide', 'Herbicide detail',
  'Description of works', 'Proposed future works', 'Fauna sightings',
  'Flora sightings of interest', 'WHS / equipment / materials', 'WHS detail'];

function summarySheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SUMMARY_TAB);
  if (!sh) {
    sh = ss.insertSheet(SUMMARY_TAB);
    sh.getRange(1, 1, 1, SUMMARY_COLS.length).setValues([SUMMARY_COLS]);
    sh.getRange(1, 1, 1, SUMMARY_COLS.length).setFontWeight('bold')
      .setBackground('#212121').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 200); sh.setColumnWidth(10, 300); sh.setColumnWidth(14, 320);
  }
  // Same trap as the Hours Log: Sheets will coerce '2026-08-05' into a Date and
  // then it no longer matches the roster keys. Pin the Date column to text.
  try { sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@'); } catch (e) {}
  return sh;
}

/** site -> council, so the form can fill it in without asking. */
function siteCouncilMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SITES_TAB);
  if (!sh) return {};
  var vals = sh.getDataRange().getValues(), hdr = vals[0].map(String);
  var iN = hdr.findIndex(function (h) { return String(h).indexOf('Roster site name') === 0; });
  var iL = hdr.indexOf('LGA');
  if (iN < 0) return {};
  var m = {};
  vals.slice(1).forEach(function (r) {
    var n = String(r[iN] || '').trim();
    if (n) m[n] = iL >= 0 ? String(r[iL] || '').trim() : '';
  });
  return m;
}

/**
 * site -> {date, species[]} from the most recent summary for that site.
 * This is what powers "last time at this site", which turns the 90-item species
 * list into two taps for a site the crew visits regularly.
 */
function lastSpeciesBySite_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SUMMARY_TAB);
  if (!sh || sh.getLastRow() < 2) return {};
  var vals = sh.getDataRange().getValues(), hdr = vals[0].map(String);
  var iT = hdr.indexOf('Timestamp'), iD = hdr.indexOf('Date'),
      iS = hdr.indexOf('Site'), iSp = hdr.indexOf('Target species');
  if (iS < 0 || iSp < 0) return {};
  var tz = Session.getScriptTimeZone();
  var out = {};
  vals.slice(1).forEach(function (r) {
    var site = String(r[iS] || '').trim();
    if (!site) return;
    var ts = String(r[iT] || '');
    if (out[site] && out[site].ts > ts) return;
    var d = r[iD];
    out[site] = {
      ts: ts,
      date: (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').trim(),
      species: String(r[iSp] || '').split(',').map(function (x) { return x.trim(); })
                                             .filter(function (x) { return x; })
    };
  });
  return out;
}

/** date|site keys that already have a summary. */
function summaryKeys_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SUMMARY_TAB);
  if (!sh || sh.getLastRow() < 2) return {};
  var vals = sh.getDataRange().getValues(), hdr = vals[0].map(String);
  var iD = hdr.indexOf('Date'), iS = hdr.indexOf('Site'), iB = hdr.indexOf('Submitted by');
  var tz = Session.getScriptTimeZone();
  var out = {};
  vals.slice(1).forEach(function (r) {
    if (!r[iD] || !r[iS]) return;
    var d = (r[iD] instanceof Date) ? Utilities.formatDate(r[iD], tz, 'yyyy-MM-dd')
                                    : String(r[iD]).trim();
    out[d + '|' + String(r[iS]).trim()] = String(r[iB] || '');
  });
  return out;
}

/** Extra data the summary half of the form needs. Merged into hoursBootstrap(). */
function summaryBootstrap_() {
  return {
    councils: siteCouncilMap_(),
    lastSpecies: lastSpeciesBySite_(),
    summarised: summaryKeys_()
  };
}

/**
 * Called from the form after hours are saved. Validated against the roster the
 * same way, so a summary can't be filed against a visit that never happened.
 */
function submitSiteSummary(p) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Busy — please try again in a moment.');
  try {
    if (!p || !p.date || !p.site) throw new Error('Missing date or site.');
    if (!p.submittedBy) throw new Error('Missing your name.');

    var ok = hoursVisits_().some(function (v) {
      return v.date === p.date && v.site === p.site;
    });
    if (!ok) throw new Error('That site is not on the roster for that date.');

    if (p.whs === 'Yes' && !String(p.whsDetail || '').trim()) {
      throw new Error('Please describe the WHS issue, broken equipment or material request.');
    }

    var num = function (x) { var n = Number(x); return (isNaN(n) || n < 0) ? 0 : n; };
    var maint = num(p.maintenance), prim = num(p.primary);
    var species = (p.species || []).join(', ');

    var row = [
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      p.date, p.site, String(p.council || ''), p.submittedBy,
      String(p.otherSupervisor || ''), maint, prim, Math.round((maint + prim) * 100) / 100,
      species, String(p.otherSpecies || '').trim(),
      String(p.herbicide || 'None'), String(p.herbicideDetail || '').trim(),
      String(p.works || '').trim(), String(p.futureWorks || '').trim(),
      String(p.fauna || '').trim(), String(p.flora || '').trim(),
      String(p.whs || 'No'), String(p.whsDetail || '').trim()
    ];
    var sh = summarySheet_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, SUMMARY_COLS.length).setValues([row]);

    return 'Site summary saved' + (species ? ' (' + (p.species || []).length + ' species)' : '') + '.';
  } finally {
    lock.releaseLock();
  }
}

/** Read-only check that the summary half is wired up. */
function checkSiteSummary() {
  var out = [];
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SUMMARY_TAB);
    out.push('Site Summary tab: ' + (sh ? (sh.getLastRow() - 1) + ' summaries recorded'
                                        : 'not created yet (normal until the first one)'));
  } catch (e) { out.push('Site Summary: ERROR ' + e.message); }
  try {
    var c = siteCouncilMap_();
    out.push('Councils mapped from SITES: ' + Object.keys(c).length + ' sites');
  } catch (e) { out.push('Councils: ERROR ' + e.message); }
  try {
    var l = lastSpeciesBySite_();
    out.push('Sites with previous species recorded: ' + Object.keys(l).length);
  } catch (e) { out.push('Last species: ERROR ' + e.message); }
  var msg = out.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Site summary', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) {}
  return msg;
}