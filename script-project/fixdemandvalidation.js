/**
 * FixDemandValidation.gs — one-run upgrade for the Demand tab.
 *
 * Run fixDemandTab() once from the Apps Script editor (Run ▸ fixDemandTab).
 * Then run checkDemandTab() any time to confirm the state (read-only).
 *
 * Fixes Dave's issue: the 'Pin staff', 'Keep off' and 'Required quals' dropdowns
 * REJECT anything that isn't a single list value, so you can't enter two names.
 * This switches those columns to warning-only validation: the dropdown still
 * offers single picks, but typing a comma-separated list ("Steve, MelL") is
 * accepted (small warning triangle only). The optimiser already splits on
 * commas, so multi-entries just work.
 *
 * Also adds a free-text 'Notes' column (before 'Pin staff' if missing).
 *
 * Safe to re-run. Reports via the execution log (View ▸ Logs) and a toast —
 * deliberately NOT SpreadsheetApp.getUi(), which throws "Cannot call
 * SpreadsheetApp.getUi() from this context" when the editor was opened from
 * script.google.com rather than from inside the spreadsheet.
 *
 * NOTE: no onOpen() here on purpose — this project already defines one in
 * RenderRoster.gs, and a second would silently replace the Apunga Roster menu.
 */
var DEMAND_TAB = 'Demand';
var SOFT_COLUMNS = ['Pin staff', 'Keep off', 'Required quals'];
var HELP_TEXT = 'Pick from the list, or type several separated by commas '
              + '(e.g. "Steve, MelL"). The warning triangle is fine.';

/** Report without ever touching getUi() — toast if we can, always log. */
function say_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'Demand tab', 10); } catch (e) { /* no UI context */ }
}

function fixDemandTab() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('No active spreadsheet. Open the sheet, then Extensions > Apps Script, '
                         + 'and run this from there.');
  var sh = ss.getSheetByName(DEMAND_TAB);
  if (!sh) throw new Error('No "' + DEMAND_TAB + '" tab found. Tabs present: '
                         + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));

  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var done = [];

  // 1) add a free-text Notes column if missing (before 'Pin staff', else at the end)
  if (hdr.indexOf('Notes') < 0) {
    var at = hdr.indexOf('Pin staff');
    if (at >= 0) { sh.insertColumnBefore(at + 1); sh.getRange(1, at + 1).setValue('Notes'); }
    else { sh.getRange(1, hdr.length + 1).setValue('Notes'); }
    hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    done.push('added Notes column');
  } else {
    done.push('Notes column already present');
  }

  // 2) soften validation on the multi-entry columns
  var nrows = Math.max(sh.getMaxRows() - 1, 1);
  SOFT_COLUMNS.forEach(function (name) {
    var cIdx = hdr.indexOf(name);
    if (cIdx < 0) { done.push(name + ': COLUMN NOT FOUND'); return; }
    var rng = sh.getRange(2, cIdx + 1, nrows, 1);
    var existing = rng.getDataValidations(), changed = 0;
    var softened = existing.map(function (row) {
      return row.map(function (dv) {
        if (!dv) return null;                     // no dropdown here: already free text
        changed++;
        return dv.copy().setAllowInvalid(true).setHelpText(HELP_TEXT).build();
      });
    });
    if (changed) { rng.setDataValidations(softened); done.push(name + ': ' + changed + ' cells softened'); }
    else { done.push(name + ': no dropdowns found (already free text)'); }
  });

  // 3) Notes column must have NO validation
  var nIdx = hdr.indexOf('Notes');
  if (nIdx >= 0) sh.getRange(2, nIdx + 1, nrows, 1).clearDataValidations();

  say_('Demand tab updated — ' + done.join('; ') + '.');
}

/** Read-only: confirms whether the fix is in place. Changes nothing. */
function checkDemandTab() {
  var sh = SpreadsheetApp.getActive().getSheetByName(DEMAND_TAB);
  if (!sh) throw new Error('No "' + DEMAND_TAB + '" tab found.');
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var out = ['Notes column: ' + (hdr.indexOf('Notes') >= 0 ? 'YES' : 'MISSING')];

  SOFT_COLUMNS.forEach(function (name) {
    var cIdx = hdr.indexOf(name);
    if (cIdx < 0) { out.push(name + ': column not found'); return; }
    var n = Math.max(Math.min(sh.getMaxRows() - 1, 200), 1);
    var dvs = sh.getRange(2, cIdx + 1, n, 1).getDataValidations();
    var strict = 0, lenient = 0, none = 0;
    dvs.forEach(function (row) {
      row.forEach(function (dv) {
        if (!dv) none++; else if (dv.getAllowInvalid()) lenient++; else strict++;
      });
    });
    out.push(name + ': ' + lenient + ' multi-entry OK, ' + strict + ' still STRICT, ' + none + ' no dropdown');
  });
  say_(out.join(' | '));
}