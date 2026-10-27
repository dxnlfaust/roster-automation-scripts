/**
 * ResolvePins.gs
 * Resolves the maps.app.goo.gl pin hyperlinks in the SITES sheet to lat/lng.
 *
 * How it works:
 *  - Reads the hyperlink behind each cell in the "Meeting points" column.
 *  - Follows the short-link redirect chain (Google blocks following these
 *    from outside Workspace, but UrlFetchApp inside your account can).
 *  - Extracts coordinates from the resolved URL, preferring the exact dropped
 *    pin (!3d/!4d) over the map centre (@lat,lng) over a q= fallback.
 *  - Writes Lat / Lng / Pin status into three new columns.
 *
 * Safe to re-run: rows that already have a Lat are skipped unless PIN_FORCE = true.
 *
 * Setup: Extensions > Apps Script, paste this in, set PIN_SHEET_NAME if needed,
 * run resolvePins(). First run will ask for authorisation (UrlFetchApp + Sheets).
 */

var PIN_SHEET_NAME       = 'SITES';
var MEETING_HEADER   = 'Meeting points (link to google maps pin)';
var PIN_FORCE            = false;   // true = recompute rows that already have a Lat
var MAX_REDIRECTS    = 10;
var PIN_SLEEP_MS         = 200;     // gentle pause between requests

function resolvePins() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PIN_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + PIN_SHEET_NAME + '" not found.');

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var header  = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  var meetCol = header.indexOf(MEETING_HEADER) + 1;
  if (meetCol === 0) throw new Error('Meeting points column not found.');

  // Place output columns immediately after the meeting-points column,
  // creating them once and reusing them on later runs.
  var latCol = ensureColumn_(sh, header, 'Lat', meetCol + 1);
  var lngCol = ensureColumn_(sh, header, 'Lng', meetCol + 2);
  var stsCol = ensureColumn_(sh, header, 'Pin status', meetCol + 3);

  // Re-read header after possible insertions.
  header  = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  meetCol = header.indexOf(MEETING_HEADER) + 1;
  latCol  = header.indexOf('Lat') + 1;
  lngCol  = header.indexOf('Lng') + 1;
  stsCol  = header.indexOf('Pin status') + 1;

  var n = lastRow - 1;
  if (n <= 0) return;

  var richVals = sh.getRange(2, meetCol, n, 1).getRichTextValues();
  var latVals  = sh.getRange(2, latCol,  n, 1).getValues();
  var done = 0, failed = 0, skipped = 0;

  for (var i = 0; i < n; i++) {
    var row = i + 2;
    if (!PIN_FORCE && latVals[i][0] !== '' && latVals[i][0] !== null) { skipped++; continue; }

    var url = richVals[i][0] ? richVals[i][0].getLinkUrl() : null;
    if (!url) { sh.getRange(row, stsCol).setValue('no pin'); continue; }

    try {
      var coords = resolveUrlToCoords_(url);
      if (coords) {
        sh.getRange(row, latCol).setValue(coords.lat);
        sh.getRange(row, lngCol).setValue(coords.lng);
        sh.getRange(row, stsCol).setValue(coords.kind); // pin | centre | query
        done++;
      } else {
        sh.getRange(row, stsCol).setValue('no coords in redirect');
        failed++;
      }
    } catch (e) {
      sh.getRange(row, stsCol).setValue('error: ' + e.message);
      failed++;
    }
    Utilities.sleep(PIN_SLEEP_MS);
  }

  SpreadsheetApp.getActive().toast(
    'Resolved ' + done + ', failed ' + failed + ', skipped ' + skipped, 'Pins', 8);
}

/** Follow the redirect chain and pull coordinates from the first URL that has them. */
function resolveUrlToCoords_(startUrl) {
  var url = startUrl;
  for (var hop = 0; hop < MAX_REDIRECTS; hop++) {
    var c = extractCoords_(url);
    if (c) return c;
    var resp = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code >= 300 && code < 400) {
      var loc = resp.getHeaders()['Location'] || resp.getHeaders()['location'];
      if (!loc) break;
      url = loc;
    } else {
      // Landed on a real page; try the final URL, then the body as a last resort.
      var c2 = extractCoords_(url);
      if (c2) return c2;
      return extractCoords_(resp.getContentText());
    }
  }
  return extractCoords_(url);
}

/** Extract coordinates from a URL or text, preferring the exact pin. */
function extractCoords_(s) {
  if (!s) return null;
  var m;
  // 1. Exact dropped pin: !3d<lat>!4d<lng>
  m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]), kind: 'pin' };
  // 2. Map centre: @<lat>,<lng>
  m = s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]), kind: 'centre' };
  // 3. Query fallback: ?q=<lat>,<lng>  or  /search/<lat>,+<lng>
  m = s.match(/[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]), kind: 'query' };
  m = s.match(/\/search\/(-?\d+\.\d+),\+?(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]), kind: 'query' };
  return null;
}

/** Create a header column at the given position if it doesn't already exist. */
function ensureColumn_(sh, header, name, atCol) {
  var idx = header.indexOf(name);
  if (idx !== -1) return idx + 1;
  sh.insertColumnBefore(atCol);
  sh.getRange(1, atCol).setValue(name);
  header.splice(atCol - 1, 0, name);
  return atCol;
}