/**
 * leaveCleanup.gs — clear leave that has finished from the Staff tab.
 *
 * A leave period is finished when its Returning date is before today (Returning
 * is the last day away). Its Upcoming Leave, Away and Returning cells are then
 * cleared, and the old values are copied to a "Leave history" tab first, so
 * nothing is lost.
 *
 * Runs automatically at the start of every "Run optimiser", and from the menu:
 * Apunga Roster ▸ Clear finished leave.
 *
 * Left alone: Away with no Returning (open-ended, may still be away), and
 * Upcoming Leave text with no dates (can't tell when it ends).
 */

var LV_STAFF_TAB = 'Staff';
var LV_HISTORY_TAB = 'Leave history';
var LV_COLS = ['Upcoming Leave', 'Away', 'Returning'];

/** Cell -> Date at midnight, or null. Typed text is read the Australian way (d/m/y). */
function lvDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/))) {
    var a = +m[1], b = +m[2], y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return b > 12 ? new Date(y, a - 1, b) : new Date(y, b - 1, a);   // m/d only if it can't be d/m
  }
  return null;
}

/** Clears finished leave. Returns the StaffIDs cleared. */
function clearFinishedLeave_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(LV_STAFF_TAB); if (!sh) return [];
  var v = sh.getDataRange().getValues(), h = -1;
  for (var i = 0; i < 15 && h < 0; i++) if (v[i].map(String).indexOf('StaffID') >= 0) h = i;
  if (h < 0) return [];
  var hd = v[h].map(function (x) { return String(x).trim(); });
  var cId = hd.indexOf('StaffID'), cUp = hd.indexOf('Upcoming Leave'), cAw = hd.indexOf('Away'), cRe = hd.indexOf('Returning');
  if (cAw < 0 || cRe < 0) return [];

  var now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var hist = [], cleared = [];
  for (var r = h + 1; r < v.length; r++) {
    var id = String(v[r][cId] || '').trim();
    if (!id) continue;
    var back = lvDate_(v[r][cRe]);
    if (!back || back >= today) continue;
    hist.push([today, id, cUp >= 0 ? String(v[r][cUp]) : '', v[r][cAw], v[r][cRe]]);
    [cUp, cAw, cRe].forEach(function (c) { if (c >= 0) sh.getRange(r + 1, c + 1).clearContent(); });
    cleared.push(id);
  }
  if (hist.length) {
    var hs = ss.getSheetByName(LV_HISTORY_TAB);
    if (!hs) {
      hs = ss.insertSheet(LV_HISTORY_TAB);
      hs.getRange(1, 1, 1, 5).setValues([['Cleared on', 'StaffID', 'Upcoming Leave', 'Away', 'Returning']])
        .setFontWeight('bold');
      hs.setFrozenRows(1);
    }
    var start = hs.getLastRow() + 1;
    hs.getRange(start, 1, hist.length, 5).setValues(hist);
    hs.getRange(start, 1, hist.length, 1).setNumberFormat('d mmm yyyy');
  }
  return cleared;
}

/** Menu: Apunga Roster ▸ Clear finished leave. */
function clearFinishedLeave() {
  var ids = clearFinishedLeave_();
  SpreadsheetApp.getActive().toast(ids.length
    ? 'Cleared finished leave for ' + ids.join(', ') + '. Old dates are in the "' + LV_HISTORY_TAB + '" tab.'
    : 'No finished leave to clear.', 'Staff leave', 10);
}
