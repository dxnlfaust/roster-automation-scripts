/**
 * BuildTravelMatrix.gs  v2  (incremental, cost-safe)
 * Builds / refreshes the staff -> site travel matrix via Routes API.
 *
 * KEY FIX vs v1: only genuinely NEW coordinate-pairs are sent to the paid API.
 * It caches every coord-pair already priced in the existing matrix, so:
 *   - Adding sites   -> bills only (origins x new sites).  ~$3 for 5 sites.
 *   - Adding a staffer in an EXISTING suburb -> $0 (rows written from cache).
 *   - Adding a staffer in a NEW suburb       -> ~$1.50 (one origin x all sites).
 *   - Moving a staff/site (coord change)     -> only that row's pairs.
 *
 * SAFETY
 *   - DRY_RUN (default true): prints the EXACT billable element count and $ for
 *     this run, with NO API calls. Always run this first after any change.
 *   - COST_CEILING_ELEMENTS: a non-dry run that would bill more than this
 *     ABORTS unless CONFIRM_LARGE = true — so you can't accidentally trigger a
 *     full rebuild. The first build (or a deliberate TM_FORCE rebuild) needs
 *     CONFIRM_LARGE = true.
 *   - Key read from Script Properties (ROUTES_API_KEY). Resumable. Paced under
 *     the per-minute quota.
 *
 * Run buildTravelMatrix().
 */

var TM_SITES_SHEET = 'SITES', TM_STAFF_SHEET = 'Staff', MATRIX_SHEET = 'TravelMatrix', KEY_PROP = 'ROUTES_API_KEY';
var TM_SITE_NAME_H = 'Roster site name (Official name)', SITE_LAT_H = 'Lat', SITE_LNG_H = 'Lng', SITE_CURRENT_H = 'Current Site';
var TM_STAFF_ID_H = 'StaffID', STAFF_LAT_H = 'Home Lat', STAFF_LNG_H = 'Home Lng', STAFF_MODE_H = 'Transport';

var TZ = 'Australia/Sydney';
var DEPART_HOUR = 6, DEPART_MIN = 20, ARRIVE_HOUR = 7, ARRIVE_MIN = 0;
var COORD_DP = 5, DRIVE_MAX_ELEMS = 625, TRANSIT_MAX_ELEMS = 100;
var ELEMENTS_PER_MIN = 2500, MAX_RUNTIME_MS = 5 * 60 * 1000, COST_PER_ELEMENT = 0.01;

var DRY_RUN = true;                 // <-- run this first; prints cost, no spend
var COST_CEILING_ELEMENTS = 1500;   // abort a real run above this unless CONFIRM_LARGE
var CONFIRM_LARGE = false;          // set true ONLY for a deliberate first build / full rebuild
var TM_FORCE = false;                  // true = recompute every pair (full rebuild; needs CONFIRM_LARGE)

var ENDPOINT = 'https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix';

function buildTravelMatrix() {
  var t0 = Date.now(), ss = SpreadsheetApp.getActiveSpreadsheet();
  var apiKey = PropertiesService.getScriptProperties().getProperty(KEY_PROP);
  if (!apiKey && !DRY_RUN) throw new Error('No ' + KEY_PROP + ' in Script Properties.');

  var staff = readStaff_(ss), sites = readSites_(ss);
  if (!staff.length || !sites.length) throw new Error('Missing staff coords or site coords.');

  // unique destinations
  var dckList = [], destByDck = {};
  sites.forEach(function (s) { if (!destByDck[s.dck]) { destByDck[s.dck] = { lat: s.lat, lng: s.lng }; dckList.push(s.dck); } });

  // origins grouped by (coord, mode)
  var origins = {};
  staff.forEach(function (p) {
    var k = p.ock + '|' + p.mode;
    if (!origins[k]) origins[k] = { ock: p.ock, mode: p.mode, lat: p.lat, lng: p.lng, staff: [] };
    origins[k].staff.push(p);
  });

  var sheet = ensureMatrixSheet_(ss);
  var ex = loadExisting_(sheet);     // {pairs: {staff|site:{row,ck}}, cache: {mode|ck:mins}, lastRow}
  var cache = ex.cache;

  // figure out what genuinely needs the API
  var needByOrigin = {}, totalNeeded = 0;
  Object.keys(origins).forEach(function (k) {
    var o = origins[k];
    var need = dckList.filter(function (dck) {
      return TM_FORCE || cache[o.mode + '|' + o.ock + '>' + dck] === undefined;
    });
    needByOrigin[k] = need; totalNeeded += need.length;
  });

  if (staff.needGeocode.length) {
    Logger.log('*** %s STAFF SKIPPED — no Home Lat/Lng. Run geocodeStaff() first, then re-run: %s',
      staff.needGeocode.length, staff.needGeocode.join(', '));
    ss.toast(staff.needGeocode.length + ' staff SKIPPED (no coords): ' + staff.needGeocode.join(', ') +
      '. Run geocodeStaff() first.', 'Travel matrix — action needed', 20);
  }
  if (staff.noMode.length)
    Logger.log('Note: blank Transport, defaulted to DRIVE: %s', staff.noMode.join(', '));
  if (staff.noHome.length)
    Logger.log('Note: no Location set (fixed-site supervisors, expected): %s', staff.noHome.join(', '));

  var modeChanged = [];
  staff.forEach(function (p) {
    for (var pk in ex.pairs) {
      if (pk.indexOf(p.id + '|') === 0) {
        if (ex.pairs[pk].mode && ex.pairs[pk].mode !== p.mode) modeChanged.push(p.id + ': ' + ex.pairs[pk].mode + ' -> ' + p.mode);
        break;
      }
    }
  });
  if (modeChanged.length)
    Logger.log('Transport mode changed since last build (their rows will be rewritten): %s', modeChanged.join(', '));

  Logger.log('Staff with coords %s, sites %s. Billable elements THIS RUN: %s (~$%s).',
    staff.length, sites.length, totalNeeded, (totalNeeded * COST_PER_ELEMENT).toFixed(2));
  if (DRY_RUN) {
    ss.toast('DRY RUN: ' + totalNeeded + ' new elements, ~$' + (totalNeeded * COST_PER_ELEMENT).toFixed(2) +
      '. No API calls.' + (staff.needGeocode.length ? ' WARNING: ' + staff.needGeocode.length +
      ' staff skipped for missing coords.' : ''), 'Travel matrix', 12);
    return;
  }
  if (totalNeeded > COST_CEILING_ELEMENTS && !CONFIRM_LARGE) {
    throw new Error('This run would bill ' + totalNeeded + ' elements (~$' +
      (totalNeeded * COST_PER_ELEMENT).toFixed(2) + '), above the ' + COST_CEILING_ELEMENTS +
      '-element safety ceiling. Review with DRY_RUN, or set CONFIRM_LARGE = true to proceed.');
  }

  var departISO = rfc3339_(DEPART_HOUR, DEPART_MIN, TZ), arriveISO = rfc3339_(ARRIVE_HOUR, ARRIVE_MIN, TZ);
  var budget = ELEMENTS_PER_MIN, winStart = Date.now(), billed = 0, written = 0, failures = [];

  var keys = Object.keys(origins);
  for (var i = 0; i < keys.length; i++) {
    if (Date.now() - t0 > MAX_RUNTIME_MS) { Logger.log('Stopped near time limit — re-run to continue.'); break; }
    var o = origins[keys[i]], need = needByOrigin[keys[i]], maxE = (o.mode === 'TRANSIT') ? TRANSIT_MAX_ELEMS : DRIVE_MAX_ELEMS;

    for (var c = 0; c < need.length; c += maxE) {
      var chunk = need.slice(c, c + maxE);
      if (budget - chunk.length < 0) {
        var wait = 60000 - (Date.now() - winStart); if (wait > 0) Utilities.sleep(wait);
        budget = ELEMENTS_PER_MIN; winStart = Date.now();
      }
      budget -= chunk.length;
      var res = callMatrix_(apiKey, o, chunk.map(function (d) { return destByDck[d]; }), o.mode, departISO, arriveISO);
      if (res.error) { failures.push(o.ock + ': ' + res.error); break; }
      res.elements.forEach(function (e) {
        cache[o.mode + '|' + o.ock + '>' + chunk[e.destinationIndex]] = (e.minutes == null ? '' : e.minutes);
      });
      billed += chunk.length; Utilities.sleep(150);
    }

    // write/upsert rows for every staffer at this origin x every site (from cache)
    var appends = [], now = new Date();
    o.staff.forEach(function (p) {
      sites.forEach(function (st) {
        var ck = o.ock + '>' + st.dck, mins = cache[o.mode + '|' + ck];
        if (mins === undefined) return;
        var pk = p.id + '|' + st.id, prev = ex.pairs[pk];
        var rowVals = [p.id, st.id, o.mode, mins, ck, now];
        if (TM_FORCE || !prev || prev.ck !== ck || prev.mode !== o.mode) {
          if (prev) { sheet.getRange(prev.row, 1, 1, 6).setValues([rowVals]); prev.ck = ck; prev.mode = o.mode; }
          else { appends.push(rowVals); }
          written++;
        }
      });
    });
    if (appends.length) {
      var start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, appends.length, 6).setValues(appends);
      for (var a = 0; a < appends.length; a++) ex.pairs[appends[a][0] + '|' + appends[a][1]] = { row: start + a, ck: appends[a][4], mode: appends[a][2] };
    }
  }

  Logger.log('Done: billed %s elements (~$%s), rows written/updated %s, failures %s, %ss.',
    billed, (billed * COST_PER_ELEMENT).toFixed(2), written, failures.length, ((Date.now() - t0) / 1000).toFixed(0));
  if (failures.length) Logger.log('Failures: ' + failures.join(' ; '));
  ss.toast('Billed ' + billed + ' elements (~$' + (billed * COST_PER_ELEMENT).toFixed(2) + '), rows ' + written +
    (failures.length ? ', FAILED ' + failures.length : ''), 'Travel matrix', 10);
}

function callMatrix_(apiKey, origin, dests, mode, departISO, arriveISO) {
  var body = {
    origins: [{ waypoint: { location: { latLng: { latitude: origin.lat, longitude: origin.lng } } } }],
    destinations: dests.map(function (d) { return { waypoint: { location: { latLng: { latitude: d.lat, longitude: d.lng } } } }; }),
    travelMode: mode
  };
  if (mode === 'TRANSIT') body.arrivalTime = arriveISO;
  else { body.routingPreference = 'TRAFFIC_AWARE'; body.departureTime = departISO; }
  var resp = UrlFetchApp.fetch(ENDPOINT, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true,
    headers: { 'X-Goog-Api-Key': apiKey, 'X-Goog-FieldMask': 'originIndex,destinationIndex,duration,condition,status' }
  });
  if (resp.getResponseCode() !== 200) return { error: 'HTTP ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 160) };
  var arr; try { arr = JSON.parse(resp.getContentText()); } catch (e) { return { error: 'bad JSON' }; }
  return { elements: arr.map(function (el) {
    var mins = null;
    if (el.condition === 'ROUTE_EXISTS' && el.duration) mins = Math.round(parseInt(String(el.duration).replace('s', ''), 10) / 60);
    return { destinationIndex: el.destinationIndex || 0, minutes: mins };
  }) };
}

function readSites_(ss) {
  var sh = ss.getSheetByName(TM_SITES_SHEET); if (!sh) throw new Error('No SITES sheet.');
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  var iN = h.indexOf(TM_SITE_NAME_H), iLa = h.indexOf(SITE_LAT_H), iLn = h.indexOf(SITE_LNG_H), iC = h.indexOf(SITE_CURRENT_H);
  if (iN < 0 || iLa < 0 || iLn < 0) throw new Error('Site headers not found.');
  var out = [];
  for (var r = 1; r < v.length; r++) {
    var n = String(v[r][iN]).trim(), la = v[r][iLa], ln = v[r][iLn];
    if (!n || la === '' || ln === '' || la == null || ln == null) continue;
    if (iC >= 0 && String(v[r][iC]).trim().toLowerCase() !== 'yes') continue;
    out.push({ id: n, lat: Number(la), lng: Number(ln), dck: ck_(la, ln) });
  }
  return out;
}
function readStaff_(ss) {
  var sh = ss.getSheetByName(TM_STAFF_SHEET); if (!sh) throw new Error('No Staff sheet.');
  var v = sh.getDataRange().getValues(), hr = -1;
  for (var r = 0; r < Math.min(v.length, 15); r++) {
    var row = v[r].map(function (x) { return String(x).trim(); });
    if (row.indexOf(TM_STAFF_ID_H) !== -1 && row.indexOf(STAFF_MODE_H) !== -1) { hr = r; break; }
  }
  if (hr < 0) throw new Error('Staff header row not found.');
  var h = v[hr].map(function (x) { return String(x).trim(); });
  var iId = h.indexOf(TM_STAFF_ID_H), iLa = h.indexOf(STAFF_LAT_H), iLn = h.indexOf(STAFF_LNG_H), iM = h.indexOf(STAFF_MODE_H);
  var iLoc = h.indexOf('Location');
  var out = [], needGeocode = [], noHome = [], noMode = [];
  for (var r2 = hr + 1; r2 < v.length; r2++) {
    var id = String(v[r2][iId]).trim(), la = v[r2][iLa], ln = v[r2][iLn];
    if (!id || id.indexOf('Wk ') === 0 || id.indexOf('TOTAL') === 0) continue;
    if (la === '' || ln === '' || la == null || ln == null) {
      // A Location with no coords means geocodeStaff() hasn't been run for this
      // person yet — actionable. No Location at all is normal for fixed-site
      // supervisors who never get crew travel.
      var loc = (iLoc >= 0) ? String(v[r2][iLoc]).trim() : '';
      if (loc) needGeocode.push(id + ' (' + loc + ')'); else noHome.push(id);
      continue;
    }
    var mdRaw = String(v[r2][iM]).trim();
    if (!mdRaw) noMode.push(id);
    var md = mdRaw.toLowerCase();
    out.push({ id: id, lat: Number(la), lng: Number(ln), ock: ck_(la, ln),
               mode: (md.indexOf('public') !== -1 || md === 'pt' || md === 'transit') ? 'TRANSIT' : 'DRIVE' });
  }
  out.needGeocode = needGeocode; out.noHome = noHome; out.noMode = noMode;
  return out;
}
function ensureMatrixSheet_(ss) {
  var sh = ss.getSheetByName(MATRIX_SHEET);
  if (!sh) { sh = ss.insertSheet(MATRIX_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['staff_id', 'site_id', 'mode', 'minutes', 'coord_key', 'updated_at']]); sh.setFrozenRows(1); }
  return sh;
}
function loadExisting_(sh) {
  var pairs = {}, cache = {}, last = sh.getLastRow();
  if (last < 2) return { pairs: pairs, cache: cache, lastRow: last };
  var v = sh.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) {
    var sid = String(v[i][0]).trim(), site = String(v[i][1]).trim(), mode = String(v[i][2]).trim(),
        mins = v[i][3], ck = String(v[i][4]).trim();
    if (!sid || !site) continue;
    pairs[sid + '|' + site] = { row: i + 2, ck: ck, mode: mode };
    if (ck) cache[mode + '|' + ck] = mins;     // cache the priced coord-pair
  }
  return { pairs: pairs, cache: cache, lastRow: last };
}
function ck_(lat, lng) { return Number(lat).toFixed(COORD_DP) + ',' + Number(lng).toFixed(COORD_DP); }
function rfc3339_(hour, min, tz) {
  var d = new Date(); d.setDate(d.getDate() + ((2 - d.getDay() + 7) % 7 || 7)); d.setHours(12, 0, 0, 0);
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd') + 'T' + ('0' + hour).slice(-2) + ':' + ('0' + min).slice(-2) + ':00' + Utilities.formatDate(d, tz, 'XXX');
}