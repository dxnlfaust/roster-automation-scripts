/**
 * RunOptimiser.gs — the "Run optimiser" button (Phase 6d).
 *
 * Gathers the four data tabs, sends them to the Cloud Run solver, and writes
 * the results back into "Master Roster" and "Run Log", then renders the
 * calendar and staff view. Managers never touch a terminal.
 *
 * No credentials are stored in this file. It reads three Script Properties:
 *   ROSTER_URL      https://roster-optimiser-472676263009.australia-southeast1.run.app
 *   ROSTER_SECRET   the ROSTER_SECRET env var set on the Cloud Run service
 *   INVOKER_SA      roster-invoker@apunga-rostering.iam.gserviceaccount.com
 *
 * Set them once: Apps Script editor > Project Settings > Script Properties,
 * or run setUpRosterOptimiser() below after filling in the values.
 *
 * The Cloud Run service is private. This script authenticates by impersonating
 * INVOKER_SA through the IAM Credentials API to mint a short-lived ID token —
 * no key file exists anywhere, which is why org key policies don't apply.
 *
 * REQUIRES (one-time, see SETUP_BUTTON.md):
 *   - the script attached to the standard GCP project (Project Settings)
 *   - appsscript.json listing the cloud-platform and external_request scopes
 *
 * NOTE: onOpen() lives in RenderRoster.gs and adds this to the menu. Do not
 * define a second onOpen() here — the last one loaded would win.
 */

var CFG_KEYS = ['ROSTER_URL', 'ROSTER_SECRET', 'INVOKER_SA'];
var TAB_STAFF = 'Staff', TAB_SITES = 'SITES', TAB_MATRIX = 'TravelMatrix', TAB_DEMAND = 'Demand';

/** Local reporter — deliberately not shared with the other files. */
function rosterSay_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'Run optimiser', 12); } catch (e) {}
}

/** One-time convenience: fill these in, run once, then blank them out again. */
function setUpRosterOptimiser() {
  var values = {
    ROSTER_URL:    '',   // e.g. https://roster-optimiser-472676263009.australia-southeast1.run.app
    ROSTER_SECRET: '',   // the ROSTER_SECRET env var on the Cloud Run service
    INVOKER_SA:    ''    // e.g. roster-invoker@apunga-rostering.iam.gserviceaccount.com
  };
  var missing = CFG_KEYS.filter(function (k) { return !values[k]; });
  if (missing.length) throw new Error('Fill in these values first: ' + missing.join(', '));
  PropertiesService.getScriptProperties().setProperties(values);
  rosterSay_('Roster optimiser configured. You can blank the values in setUpRosterOptimiser() again.');
}

function getConfig_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var missing = CFG_KEYS.filter(function (k) { return !props[k]; });
  if (missing.length) {
    throw new Error('Missing Script Properties: ' + missing.join(', ') +
      '. Set them in Project Settings > Script Properties (see SETUP_BUTTON.md).');
  }
  props.ROSTER_URL = String(props.ROSTER_URL).replace(/\/+$/, '');
  return props;
}

/**
 * Mint an ID token for the Cloud Run service by impersonating INVOKER_SA.
 * Apps Script's own identity token has the wrong audience for Cloud Run, so we
 * exchange the user's OAuth token for a service-account ID token instead.
 */
function getIdToken_(cfg) {
  var endpoint = 'https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/' +
                 encodeURIComponent(cfg.INVOKER_SA) + ':generateIdToken';
  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ audience: cfg.ROSTER_URL, includeEmail: true }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode(), body = res.getContentText();
  if (code !== 200) {
    if (code === 403) {
      throw new Error('Not allowed to impersonate ' + cfg.INVOKER_SA + '. You need the ' +
        'Service Account Token Creator role on it (owner is not enough), and the IAM ' +
        'Service Account Credentials API must be enabled. Details: ' + body.slice(0, 300));
    }
    throw new Error('Could not get an access token (HTTP ' + code + '): ' + body.slice(0, 300));
  }
  var token = JSON.parse(body).token;
  if (!token) throw new Error('Token response had no token: ' + body.slice(0, 200));
  return token;
}

function readTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('No "' + name + '" tab. Tabs present: ' +
      ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  }
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('The "' + name + '" tab looks empty.');
  // Dates cross as ISO yyyy-mm-dd: unambiguous whichever way the sheet displays
  // them (Away/Returning mix real dates with typed dd/mm text). Needs solver >= the
  // 23 Oct ISO fix; older solvers read ISO day-first in the Away column.
  var tz = Session.getScriptTimeZone();
  return vals.map(function (row) {
    return row.map(function (c) {
      if (c instanceof Date) return Utilities.formatDate(c, tz, 'yyyy-MM-dd');
      return c === null || c === undefined ? '' : String(c);
    });
  });
}

/** Wake the container so the real request isn't paying cold-start out of its 60s. */
function warmUp_(cfg, token) {
  try {
    UrlFetchApp.fetch(cfg.ROSTER_URL + '/', {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
  } catch (e) { /* best effort only */ }
}

function runOptimiser() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.getActive();
  var cfg = getConfig_();

  // Stop two managers running at once — the second would overwrite the first.
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Someone else is running the optimiser right now. Try again in a minute.');
  }

  try {
    // tidy finished leave out of the Staff tab first (copied to "Leave history")
    try {
      var lv = clearFinishedLeave_(ss);
      if (lv.length) ss.toast('Cleared finished leave: ' + lv.join(', '), 'Run optimiser', 8);
    } catch (e) { Logger.log('Leave cleanup skipped: ' + e); }

    ss.toast('Reading the data tabs…', 'Run optimiser', 30);
    var payload = {
      secret: cfg.ROSTER_SECRET,
      staff:  readTab_(ss, TAB_STAFF),     // raw, including the rows above the header
      sites:  readTab_(ss, TAB_SITES),
      matrix: readTab_(ss, TAB_MATRIX),
      demand: readTab_(ss, TAB_DEMAND)
    };

    ss.toast('Signing in to the solver…', 'Run optimiser', 30);
    var token = getIdToken_(cfg);
    warmUp_(cfg, token);

    ss.toast('Solving — this usually takes a few seconds…', 'Run optimiser', 60);
    var res = UrlFetchApp.fetch(cfg.ROSTER_URL + '/solve', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode(), text = res.getContentText();
    if (code === 401) throw new Error('The solver rejected the secret. Check ROSTER_SECRET in ' +
      'Script Properties matches the ROSTER_SECRET on the Cloud Run service.');
    if (code === 403) throw new Error('Not allowed to call the solver. ' + cfg.INVOKER_SA +
      ' needs roles/run.invoker on the Cloud Run service.');
    if (code !== 200) throw new Error('The solver returned HTTP ' + code + ': ' + text.slice(0, 300));

    var out;
    try { out = JSON.parse(text); }
    catch (e) { throw new Error('The solver returned something unreadable: ' + text.slice(0, 200)); }

    // A data problem still returns a run log — write it so the manager can see why.
    if (out.run_log) writeTab_(ss, RUNLOG_TAB, out.run_log.columns, out.run_log.rows);

    if (!out.ok) {
      try { formatRunLog(); } catch (e) { /* formatting is optional */ }
      ss.toast('Could not build a roster: ' + out.error, 'Run optimiser — see Run Log', 20);
      return;
    }

    // Protect whatever is in Master Roster now (possibly hand-edited) before
    // overwriting it, and record what changed.
    var before = null;
    try { before = snapshotMasterRoster_(); } catch (e) { Logger.log('snapshot failed: ' + e); }

    writeTab_(ss, DATA_TAB, out.master_roster.columns, out.master_roster.rows);

    var nChanges = 0;
    if (before) {                      // nothing to compare against on a first run
      try {
        var after = [out.master_roster.columns].concat(out.master_roster.rows);
        nChanges = writeChanges_(ss, before, after, 'Optimiser run');
      } catch (e) { Logger.log('change log failed: ' + e); }
    }

    ss.toast('Rendering…', 'Run optimiser', 20);
    renderCalendar();
    renderStaffView();
    try { formatRunLog(); } catch (e) { /* ignore */ }

    var s = out.stats || {};
    var secs = Math.round((Date.now() - t0) / 1000);
    var msg = (s.visits || '?') + ' visits rostered in ' + secs + 's' +
              (s.status && s.status !== 'OPTIMAL' ? ' (' + s.status + ' — not proven best)' : '') +
              '. ' + (s.actions || 0) + ' item(s) need action' +
              (before ? ', ' + nChanges + ' change(s) vs previous — see Roster Changes' : '') + '.';
    ss.toast(msg, 'Run optimiser — done', 15);
    Logger.log(msg + ' stats=' + JSON.stringify(s));

  } catch (err) {
    var m = String(err && err.message || err);
    if (m.indexOf('Timeout') >= 0 || m.indexOf('timed out') >= 0) {
      m = 'The solver took too long to answer (over ~60s). Lower SOLVE_SECONDS on the ' +
          'Cloud Run service to about 40, or try again — the first run of the day is slowest.';
    }
    Logger.log('runOptimiser failed: ' + m);
    ss.toast(m, 'Run optimiser — FAILED', 20);
    throw new Error(m);
  } finally {
    lock.releaseLock();
  }
}

/**
 * The optimiser writes dates as M/D/YYYY text. Typed into a cell, Sheets reads
 * that by the *user's* locale, so in an Australian account "11/2/2026" became
 * 11 February. ISO yyyy-mm-dd text is read the same way in every locale.
 */
function isoDate_(v) {
  var m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2) : v;
}

/** Replace a tab's contents with columns + rows. */
function writeTab_(ss, name, columns, rows) {
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  var dc = columns.indexOf('Date');
  var data = [columns].concat(rows.map(function (r) {
    var out = r.slice(0, columns.length);
    while (out.length < columns.length) out.push('');
    out = out.map(function (c) { return c === null || c === undefined ? '' : c; });
    if (dc >= 0) out[dc] = isoDate_(out[dc]);
    return out;
  }));
  sh.getRange(1, 1, data.length, columns.length).setValues(data);
  if (dc >= 0 && data.length > 1)
    sh.getRange(2, dc + 1, data.length - 1, 1).setNumberFormat('ddd d mmm yyyy');
  sh.getRange(1, 1, 1, columns.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/** Read-only check that everything is wired up. Changes nothing. */
function checkOptimiserConnection() {
  var cfg = getConfig_();
  var out = ['URL: ' + cfg.ROSTER_URL, 'Service account: ' + cfg.INVOKER_SA];
  var token = getIdToken_(cfg);
  out.push('Token: obtained OK');
  var res = UrlFetchApp.fetch(cfg.ROSTER_URL + '/', {
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  out.push('Health check: HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 400));
  var ss = SpreadsheetApp.getActive();
  [TAB_STAFF, TAB_SITES, TAB_MATRIX, TAB_DEMAND].forEach(function (t) {
    var sh = ss.getSheetByName(t);
    out.push(t + ': ' + (sh ? sh.getLastRow() + ' rows' : 'MISSING'));
  });
  rosterSay_(out.join(' | '));
  Logger.log(out.join('\n'));
}