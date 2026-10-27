/**
 * SetupDemandTab.gs  v3  (monthly, date-stamped)
 * Builds (or refreshes) the Demand tab for a whole month of visits.
 *
 * Columns:
 *   Date | Day | Site | Secondary site | Supervisor | Site Hours | Crew size |
 *   Required quals | Force ATM | Pin staff | Keep off | Lock crew
 *
 * Date is a real calendar date (so the optimiser can map each row to its
 * fortnight and weekday). onEdit auto-fills:
 *   - Day, derived from the Date (1 Jun 2026 -> Mon), so you never type both
 *   - Supervisor from the chosen site's default (Supervisor A)
 *   - Date AND Day carried down from the row above when blank (block in a run
 *     of days, only re-enter the Date when it changes)
 *   - Crew size from Site Hours:  ceil((hours - 8) / 7.5 + 1), min 2, editable
 *
 * Secondary site is recorded but not routed. Safe: only touches the Demand tab.
 * Run setupDemandTab() once (delete any old Demand tab first, since columns moved).
 *
 * NOTE: the optimiser's month/fortnight handling is a separate build — until
 * that's done it still reads by weekday only, so don't run a full month through
 * it yet. This tab is the input format (and Phase 5's future output format).
 */

var DEMAND_SHEET = 'Demand';
var DEM_SITES_SHEET  = 'SITES';
var DEM_STAFF_SHEET  = 'Staff';
var DEM_SITE_NAME_H  = 'Roster site name (Official name)';
var SITE_SUPA_H  = 'Supervisor A';
var DEM_STAFF_ID_H   = 'StaffID';

var DAYS    = ['Mon', 'Tue', 'Wed', 'Thur', 'Fri'];
var WEEKDAY = ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat'];  // getDay() index -> label
var QUALS   = ['chainsaw', 'whitecard', 'irata', 'wwc'];
var ROWS    = 400;
var SUP_HOURS = 8, CREW_HOURS = 7.5;
var DATE_FMT = 'd mmm yyyy';

var C_DATE = 1, C_DAY = 2, C_SITE = 3, C_SEC = 4, C_SUP = 5, C_HRS = 6, C_CREW = 7,
    C_QUAL = 8, C_ATM = 9, C_PIN = 10, C_OFF = 11, C_LOCK = 12;
var HEADERS = ['Date', 'Day', 'Site', 'Secondary site', 'Supervisor', 'Site Hours', 'Crew size',
               'Required quals', 'Force ATM', 'Pin staff', 'Keep off', 'Lock crew'];

function setupDemandTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEMAND_SHEET) || ss.insertSheet(DEMAND_SHEET);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);

  var siteRange  = headerColumnRange_(ss, DEM_SITES_SHEET, DEM_SITE_NAME_H, 1);
  var staffRange = staffIdRange_(ss);

  sh.getRange(2, C_DATE, ROWS, 1).setNumberFormat(DATE_FMT);
  setDate_(sh, C_DATE);
  setList_(sh, C_DAY, DAYS, true);          // auto-filled from Date; lenient
  setRange_(sh, C_SITE, siteRange, false);
  setRange_(sh, C_SEC, siteRange, false);
  setRange_(sh, C_SUP, staffRange, false);
  setNumberMin_(sh, C_HRS, 0);
  setNumber_(sh, C_CREW, 2, 8);
  setList_(sh, C_QUAL, QUALS, true);
  sh.getRange(2, C_ATM, ROWS, 1).insertCheckboxes();
  setRange_(sh, C_PIN, staffRange, true);
  setRange_(sh, C_OFF, staffRange, true);
  sh.getRange(2, C_LOCK, ROWS, 1).insertCheckboxes();

  sh.autoResizeColumns(1, HEADERS.length);
  ss.toast('Monthly Demand tab ready. Enter a Date; Day fills itself.', 'Setup', 6);
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== DEMAND_SHEET) return;
  var row = e.range.getRow(), c = e.range.getColumn();
  if (row < 2) return;

  if (c === C_DATE) setDayFromDate_(sh, row);

  if (c === C_SITE) {
    var site = e.range.getValue();
    if (site) {
      var supCell = sh.getRange(row, C_SUP);
      if (!supCell.getValue()) {
        var map = siteSupervisorMap_();
        if (map[site]) supCell.setValue(map[site]);
      }
      // carry Date down, then derive Day from it (fall back to copying Day)
      var dateCell = sh.getRange(row, C_DATE);
      if (!dateCell.getValue() && row > 2) {
        var aboveDate = sh.getRange(row - 1, C_DATE).getValue();
        if (aboveDate) dateCell.setValue(aboveDate);
      }
      var dayCell = sh.getRange(row, C_DAY);
      if (!dayCell.getValue()) {
        if (dateCell.getValue()) setDayFromDate_(sh, row);
        else if (row > 2) {
          var aboveDay = sh.getRange(row - 1, C_DAY).getValue();
          if (aboveDay) dayCell.setValue(aboveDay);
        }
      }
    }
  }

  if (c === C_HRS) {
    var hrs = parseFloat(e.range.getValue());
    if (!isNaN(hrs) && hrs > 0) {
      sh.getRange(row, C_CREW).setValue(Math.max(2, Math.ceil((hrs - SUP_HOURS) / CREW_HOURS + 1)));
    }
  }
}

function setDayFromDate_(sh, row) {
  var d = sh.getRange(row, C_DATE).getValue();
  if (d && Object.prototype.toString.call(d) === '[object Date]') {
    sh.getRange(row, C_DAY).setValue(WEEKDAY[d.getDay()]);
  }
}

// ---- helpers -----------------------------------------------------------------
function setDate_(sh, col) {
  sh.getRange(2, col, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
}
function setList_(sh, col, list, allowInvalid) {
  sh.getRange(2, col, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(!!allowInvalid).build());
}
function setRange_(sh, col, range, allowInvalid) {
  sh.getRange(2, col, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(!!allowInvalid).build());
}
function setNumber_(sh, col, lo, hi) {
  sh.getRange(2, col, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(lo, hi).setAllowInvalid(false).build());
}
function setNumberMin_(sh, col, lo) {
  sh.getRange(2, col, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThan(lo).setAllowInvalid(false).build());
}
function headerColumnRange_(ss, sheetName, headerName, headerRowIdx) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" not found.');
  var header = sh.getRange(headerRowIdx, 1, 1, sh.getLastColumn()).getValues()[0]
                 .map(function (v) { return String(v).trim(); });
  var col = header.indexOf(headerName) + 1;
  if (col === 0) throw new Error('"' + headerName + '" not found on ' + sheetName);
  return sh.getRange(headerRowIdx + 1, col, Math.max(1, sh.getLastRow() - headerRowIdx), 1);
}
function staffIdRange_(ss) {
  var sh = ss.getSheetByName(DEM_STAFF_SHEET);
  if (!sh) throw new Error('Sheet "' + DEM_STAFF_SHEET + '" not found.');
  var scan = sh.getRange(1, 1, Math.min(sh.getLastRow(), 15), sh.getLastColumn()).getValues();
  for (var r = 0; r < scan.length; r++) {
    var rowv = scan[r].map(function (v) { return String(v).trim(); });
    var c = rowv.indexOf(DEM_STAFF_ID_H);
    if (c !== -1) return sh.getRange(r + 2, c + 1, Math.max(1, sh.getLastRow() - r - 1), 1);
  }
  throw new Error('StaffID column not found on ' + DEM_STAFF_SHEET);
}
function siteSupervisorMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEM_SITES_SHEET);
  var vals = sh.getDataRange().getValues();
  var h = vals[0].map(function (v) { return String(v).trim(); });
  var iName = h.indexOf(DEM_SITE_NAME_H), iSup = h.indexOf(SITE_SUPA_H);
  var map = {};
  for (var r = 1; r < vals.length; r++) {
    var name = String(vals[r][iName]).trim(), sup = String(vals[r][iSup]).trim();
    if (name && sup) map[name] = sup;
  }
  return map;
}