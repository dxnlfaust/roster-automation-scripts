/**
 * renameTool.gs — rename a site or a staff member everywhere at once.
 *
 * Names are the join keys between tabs, so a rename typed into one tab silently
 * breaks the others (tick rules, travel times, pins). This does them together.
 *
 * Use: fill the "Renames" tab (one row per rename), then
 *      Apunga Roster ▸ Apply renames (Renames tab).
 *
 *   Type  | Old name                 | New name        | Result
 *   Site  | Eastern (Eastern Reserve)| Eastern Reserve | (filled in by the script)
 *   Staff | JessB                    | Jess            |
 *
 * Old names are matched ignoring odd spaces and line breaks, so a copy-paste from
 * Sites works even when the cell has hidden characters. Rows with a Result that
 * starts "Done" are skipped, so the tab doubles as a record.
 *
 * SITE rename touches: Sites, Demand (Site, Secondary site), TravelMatrix (site_id,
 *   so no rebuild or API cost), Master Roster, Calendar aliases, and Staff
 *   Excluded sites / Preferred Sites entries that name the site.
 * STAFF rename touches: Staff (StaffID + name lists), TravelMatrix (staff_id; any
 *   rows already under the NEW name belonged to someone else and are removed
 *   first), Demand (Supervisor, Pin staff, Keep off), Sites (Supervisor,
 *   Preferred ATM), Master Roster. Hours Log is history and is left alone.
 */

var REN_TAB = 'Renames';
var REN_HDR = ['Type', 'Old name', 'New name', 'Result'];

function renNorm_(s) { return String(s == null ? '' : s).replace(/[ \s]+/g, ' ').trim().toLowerCase(); }
function renShort_(s) { var n = renNorm_(s), i = n.indexOf(' ('); return i > 0 ? n.slice(0, i).trim() : n; }

/** Creates the Renames tab with headers if it's missing. */
function setupRenamesTab() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(REN_TAB);
  if (!sh) {
    sh = ss.insertSheet(REN_TAB);
    sh.getRange(1, 1, 1, 4).setValues([REN_HDR]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A2:A200').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Site', 'Staff'], true).build());
    sh.setColumnWidths(1, 4, 240);
  }
  ss.setActiveSheet(sh);
  return sh;
}

/** Header-row helpers for a tab whose header may not be row 1 (Staff). */
function renTable_(ss, name, mustHave) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return null;
  var v = sh.getDataRange().getValues(), h = 0;
  if (mustHave) for (h = 0; h < Math.min(15, v.length); h++)
    if (v[h].map(function (x) { return String(x).trim(); }).indexOf(mustHave) >= 0) break;
  if (h >= v.length) return null;
  var hdr = v[h].map(function (x) { return String(x).trim(); });
  return { sh: sh, v: v, h: h, col: function (n) { return hdr.indexOf(n); }, changed: {} };
}
function renSet_(t, r, c, val) { t.v[r][c] = val; t.changed[r + ',' + c] = true; }
function renSave_(t) {                // writes only the changed cells, so formulas elsewhere survive
  if (!t) return;
  Object.keys(t.changed).forEach(function (k) {
    var rc = k.split(','); t.sh.getRange(+rc[0] + 1, +rc[1] + 1).setValue(t.v[+rc[0]][+rc[1]]);
  });
}

/** Whole-cell replace in named columns. Returns count. */
function renCells_(t, cols, match, to) {
  var n = 0; if (!t) return 0;
  cols.forEach(function (c) {
    var ci = t.col(c); if (ci < 0) return;
    for (var r = t.h + 1; r < t.v.length; r++) {
      if (String(t.v[r][ci]) !== '' && match(t.v[r][ci])) { renSet_(t, r, ci, to); n++; }
    }
  });
  return n;
}
/** Replace matching entries inside comma / line-break separated lists. Returns count. */
function renLists_(t, cols, match, to) {
  var n = 0; if (!t) return 0;
  cols.forEach(function (c) {
    var ci = t.col(c); if (ci < 0) return;
    for (var r = t.h + 1; r < t.v.length; r++) {
      var cell = String(t.v[r][ci]); if (!cell) continue;
      var parts = cell.split(/(\s*[,\n]\s*)/), hit = false;
      for (var i = 0; i < parts.length; i += 2) if (parts[i] && match(parts[i])) { parts[i] = to; hit = true; n++; }
      if (hit) renSet_(t, r, ci, parts.join(''));
    }
  });
  return n;
}

function applyRenames() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var rt = ss.getSheetByName(REN_TAB);
  if (!rt || rt.getLastRow() < 2) { setupRenamesTab(); ui.alert('Fill the Renames tab (Type, Old name, New name), then run this again.'); return; }
  var rows = rt.getRange(2, 1, rt.getLastRow() - 1, 4).getValues();
  var todo = rows.map(function (r, i) { return { i: i, type: String(r[0]).trim().toLowerCase(), old: String(r[1]), neu: String(r[2]).trim(), res: String(r[3]) }; })
    .filter(function (r) { return r.old.trim() && r.neu && !/^Done/.test(r.res); });
  if (!todo.length) { ui.alert('Nothing to do: every row is empty or already Done.'); return; }
  if (ui.alert('Apply ' + todo.length + ' rename(s)?',
      'Changes Sites, Staff, Demand, TravelMatrix, Master Roster and Calendar aliases. ' +
      'Backups of Sites, Staff and Demand are made first (hidden tabs).', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another change is running. Try again shortly.');
  try {
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
    ['Sites', 'Staff', 'Demand'].forEach(function (n) {
      var s = ss.getSheetByName(n); if (s) s.copyTo(ss).setName(n + ' backup ' + stamp).hideSheet();
    });

    var T = {
      sites: renTable_(ss, 'Sites', 'Roster site name (Official name)'),
      staff: renTable_(ss, 'Staff', 'StaffID'),
      demand: renTable_(ss, 'Demand', 'Date'),
      roster: renTable_(ss, 'Master Roster', 'Date'),
      aliases: renTable_(ss, IMP_ALIAS_TAB, null)
    };
    var mx = ss.getSheetByName('TravelMatrix');
    var M = mx && mx.getLastRow() > 1 ? mx.getRange(2, 1, mx.getLastRow() - 1, 6).getValues() : [];
    var mxDirty = false, mxRows = M.length;

    var results = {};
    todo.forEach(function (r) {
      var out = [];
      try {
        if (r.type === 'site') out = renameSite_(T, M, r);
        else if (r.type === 'staff') out = renameStaff_(T, M, r);
        else throw new Error('Type must be Site or Staff');
        if (out.matrix) mxDirty = true;
        if (out.matrixRows) M = out.matrixRows;
        results[r.i] = 'Done ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM') + ': ' + out.summary;
      } catch (e) {
        results[r.i] = 'NOT DONE: ' + e.message;
      }
    });

    Object.keys(T).forEach(function (k) { renSave_(T[k]); });
    if (mxDirty) {
      if (M.length) mx.getRange(2, 1, M.length, 6).setValues(M);
      if (M.length < mxRows) mx.getRange(2 + M.length, 1, mxRows - M.length, 6).clearContent();
    }
    Object.keys(results).forEach(function (i) { rt.getRange(+i + 2, 4).setValue(results[i]); });
  } finally { lock.releaseLock(); }

  var bad = Object.keys(results).filter(function (i) { return /^NOT/.test(results[i]); }).length;
  ui.alert('Renames applied', (todo.length - bad) + ' done' + (bad ? ', ' + bad + ' NOT done (see the Result column)' : '') +
    '.\n\nRun the optimiser again so the roster picks up the new names.', ui.ButtonSet.OK);
}

function renameSite_(T, M, r) {
  var siteCol = T.sites.col('Roster site name (Official name)');
  var all = T.sites.v.slice(T.sites.h + 1).map(function (x) { return String(x[siteCol]); }).filter(function (x) { return x.trim(); });
  var oldN = renNorm_(r.old), newN = renNorm_(r.neu);
  var hitOld = all.filter(function (n) { return renNorm_(n) === oldN; });
  var hitNew = all.filter(function (n) { return renNorm_(n) === newN; });
  if (!hitOld.length && !hitNew.length) throw new Error("no site called '" + r.old.trim() + "' in Sites");
  if (hitOld.length && hitNew.length) throw new Error("'" + r.neu + "' already exists in Sites");
  // short form of the old name ('Eastern' for 'Eastern (Eastern Reserve)'), only if no other site shares it
  var oldShort = renShort_(r.old);
  var shortUnique = all.filter(function (n) { return renShort_(n) === oldShort; }).length <= 1 && oldShort !== oldN;
  var exact = function (v) { return renNorm_(v) === oldN; };
  var listMatch = function (v) { var t = renNorm_(v); return t === oldN || (shortUnique && t === oldShort); };

  var n = {};
  n.sites = renCells_(T.sites, ['Roster site name (Official name)'], exact, r.neu);
  n.demand = renCells_(T.demand, ['Site', 'Secondary site'], exact, r.neu);
  n.roster = renCells_(T.roster, ['Site'], function (v) { return renNorm_(String(v).replace(/ \/ /g, '\n')) === oldN || exact(v); }, r.neu);
  n.aliases = 0;
  if (T.aliases) for (var a = 1; a < T.aliases.v.length; a++)
    if (exact(T.aliases.v[a][1])) { renSet_(T.aliases, a, 1, r.neu); n.aliases++; }
  n.staffLists = renLists_(T.staff, ['Excluded sites', 'Preferred Sites'], listMatch, r.neu);
  n.matrix = 0;
  for (var i = 0; i < M.length; i++) if (exact(M[i][1])) { M[i][1] = r.neu; n.matrix++; }
  return { matrix: n.matrix > 0, summary: 'Sites ' + n.sites + ', Demand ' + n.demand + ', TravelMatrix ' + n.matrix +
           ', Master Roster ' + n.roster + ', Staff lists ' + n.staffLists + ', aliases ' + n.aliases };
}

function renameStaff_(T, M, r) {
  var idCol = T.staff.col('StaffID');
  var ids = T.staff.v.slice(T.staff.h + 1).map(function (x) { return String(x[idCol]).trim(); });
  var oldId = r.old.trim(), neu = r.neu;
  var hasOld = ids.indexOf(oldId) >= 0, hasNew = ids.indexOf(neu) >= 0;
  if (!hasOld && !hasNew) throw new Error("no StaffID '" + oldId + "' in Staff");
  if (hasOld && hasNew) throw new Error("StaffID '" + neu + "' is already in use in Staff");
  var exact = function (v) { return String(v).trim() === oldId; };
  var ci = function (v) { return String(v).trim().toLowerCase() === oldId.toLowerCase(); };

  var n = {};
  n.staff = renCells_(T.staff, ['StaffID'], exact, neu);
  n.lists = renLists_(T.staff, ['Not work with Staff', 'Preferred Staff / Pairing', 'Carpool with'], ci, neu);
  n.demand = renCells_(T.demand, ['Supervisor'], exact, neu) + renLists_(T.demand, ['Pin staff', 'Keep off'], ci, neu);
  n.sites = renCells_(T.sites, ['Supervisor', 'Supervisor A', 'Preferred ATM'], exact, neu);
  var rcols = T.roster ? T.roster.v[T.roster.h].map(String).filter(function (h) { return /^(Supervisor|ATM|Staff\s*\d+)$/.test(h.trim()); }) : [];
  n.roster = renCells_(T.roster, rcols, exact, neu);
  // TravelMatrix: drop stale rows already under the new name (a former staff member), then rename
  var before = M.length;
  var kept = M.filter(function (row) { return String(row[0]).trim() !== neu; });
  var removed = before - kept.length, renamed = 0;
  kept.forEach(function (row) { if (String(row[0]).trim() === oldId) { row[0] = neu; renamed++; } });
  return { matrix: renamed > 0 || removed > 0, matrixRows: kept,
           summary: 'Staff ' + n.staff + ', Staff lists ' + n.lists + ', Demand ' + n.demand + ', Sites ' + n.sites +
                    ', Master Roster ' + n.roster + ', TravelMatrix ' + renamed + ' renamed' +
                    (removed ? ' + ' + removed + " old '" + neu + "' rows removed" : '') };
}
