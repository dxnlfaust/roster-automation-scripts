/**
 * GeocodeStaff.gs
 * Geocodes the staff "Location" (suburb) column to coordinates and writes
 * two columns — Home Lat / Home Lng — appended at the FAR RIGHT of the sheet,
 * so the readable constraint columns in the middle stay uncluttered.
 *
 * Uses the built-in Apps Script Maps geocoder: no API key, no billing setup.
 * Every suburb is queried as "<suburb>, NSW, Australia" so same-named suburbs
 * in other states can't be picked by mistake.
 *
 * Safe to re-run: rows that already have a Home Lat are skipped unless GEO_FORCE.
 * Rows with a blank Location (e.g. fixed-site supervisors) are left untouched.
 *
 * Failures are reported via a toast + the execution log (View > Logs),
 * deliberately NOT written into the sheet, to keep it clean.
 *
 * Setup: paste into Extensions > Apps Script, set GEO_SHEET_NAME to your staff
 * tab name if it isn't "Staff", then run geocodeStaff().
 */

var GEO_SHEET_NAME      = 'Staff';     // <-- set to your staff tab name
var LOCATION_HEADER = 'Location';
var KEY_HEADER      = 'StaffID';   // used only to locate the header row
var REGION          = 'au';
var SUBURB_SUFFIX   = ', NSW, Australia';
var GEO_FORCE           = false;       // true = recompute rows that already have a Home Lat
var GEO_SLEEP_MS        = 150;

function geocodeStaff() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(GEO_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + GEO_SHEET_NAME + '" not found.');

  // Locate the header row (the one carrying StaffID + Location) and the
  // Location column. Tolerant of the title/blank rows above the real header.
  var scan = sh.getRange(1, 1, Math.min(sh.getLastRow(), 15), sh.getLastColumn()).getValues();
  var headerRow = -1, locCol = -1;
  for (var r = 0; r < scan.length; r++) {
    var cells = scan[r].map(function (v) { return String(v).trim(); });
    if (cells.indexOf(KEY_HEADER) !== -1 && cells.indexOf(LOCATION_HEADER) !== -1) {
      headerRow = r + 1;
      locCol    = cells.indexOf(LOCATION_HEADER) + 1;
      break;
    }
  }
  if (headerRow === -1) throw new Error('No header row containing ' + KEY_HEADER + ' and ' + LOCATION_HEADER + '.');

  // Find or create the two output columns at the far right.
  var headerCells = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(function (v) { return String(v).trim(); });
  var latCol = headerCells.indexOf('Home Lat') + 1;
  var lngCol = headerCells.indexOf('Home Lng') + 1;
  if (latCol === 0) { latCol = sh.getLastColumn() + 1; sh.getRange(headerRow, latCol).setValue('Home Lat'); }
  if (lngCol === 0) { lngCol = sh.getLastColumn() + 1; sh.getRange(headerRow, lngCol).setValue('Home Lng'); }

  var firstDataRow = headerRow + 1;
  var n = sh.getLastRow() - firstDataRow + 1;
  if (n <= 0) return;

  var locs = sh.getRange(firstDataRow, locCol, n, 1).getValues();
  var lats = sh.getRange(firstDataRow, latCol, n, 1).getValues();

  var geocoder = Maps.newGeocoder().setRegion(REGION);
  var done = 0, skipped = 0, blank = 0, failed = [];

  for (var i = 0; i < n; i++) {
    var row = firstDataRow + i;
    var raw = String(locs[i][0]).trim();
    if (!raw) { blank++; continue; }
    if (!GEO_FORCE && lats[i][0] !== '' && lats[i][0] !== null) { skipped++; continue; }

    var suburb = raw.replace(/\(.*?\)/g, '').trim(); // drop any parenthetical note
    if (!suburb) { blank++; continue; }

    try {
      var resp = geocoder.geocode(suburb + SUBURB_SUFFIX);
      if (resp.status === 'OK' && resp.results.length) {
        var loc = resp.results[0].geometry.location;
        sh.getRange(row, latCol).setValue(loc.lat);
        sh.getRange(row, lngCol).setValue(loc.lng);
        done++;
      } else {
        failed.push(suburb + ' (' + resp.status + ')');
      }
    } catch (e) {
      failed.push(suburb + ' (' + e.message + ')');
    }
    Utilities.sleep(GEO_SLEEP_MS);
  }

  Logger.log('Geocoded %s, skipped %s, blank %s, failed %s', done, skipped, blank, failed.length);
  if (failed.length) Logger.log('Failures: ' + failed.join('; '));
  SpreadsheetApp.getActive().toast(
    'Geocoded ' + done + ', skipped ' + skipped + ', blank ' + blank +
    (failed.length ? ', FAILED ' + failed.length + ' (see View > Logs)' : ''),
    'Staff geocode', 8);
}