/**
 * SnapshotDiff.gs — protects manual edits to the Master Roster (Phase 6e).
 *
 * Every time the optimiser is about to overwrite the Master Roster, the current
 * contents are copied to a hidden dated snapshot tab first, and the differences
 * are written to a "Roster Changes" tab in plain English:
 *
 *     Staff changed | Fri 24 Jul | BBC | Cam removed, Sasha added
 *
 * So a re-run can never silently destroy someone's 7am fix — the change is
 * visible, and "Restore last snapshot" puts it back.
 *
 * Menu items live in RenderRoster.gs (onOpen).
 */

var SNAP_PREFIX = 'Snap ';        // tab names: "Snap 2026-08-03 11-20"
var MAX_SNAPSHOTS = 6;            // older ones are deleted automatically
var CHANGES_TAB = 'Roster Changes';

/** Read a tab as a 2D array, or null if it's missing/empty. */
function snapRead_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return null;
  return sh.getDataRange().getValues().map(function (r) {
    return r.map(function (c) {
      if (c instanceof Date) {
        return Utilities.formatDate(c, Session.getScriptTimeZone(), 'yyyy-MM-dd');   // ISO: locale-proof
      }
      return c === null || c === undefined ? '' : String(c).trim();
    });
  });
}

/** Copy the current Master Roster to a dated hidden tab. Returns its data, or null. */
function snapshotMasterRoster_() {
  var ss = SpreadsheetApp.getActive();
  var data = snapRead_(ss, DATA_TAB);
  if (!data) return null;                       // nothing to protect yet

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  var name = SNAP_PREFIX + stamp;
  if (ss.getSheetByName(name)) ss.deleteSheet(ss.getSheetByName(name));   // same minute re-run
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
  sh.hideSheet();
  pruneSnapshots_(ss);
  return data;
}

function pruneSnapshots_(ss) {
  var snaps = ss.getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n.indexOf(SNAP_PREFIX) === 0; })
    .sort();                                     // name sorts chronologically
  while (snaps.length > MAX_SNAPSHOTS) {
    ss.deleteSheet(ss.getSheetByName(snaps.shift()));
  }
}

/** Newest snapshot tab name, or null. */
function latestSnapshotName_(ss) {
  var snaps = ss.getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n.indexOf(SNAP_PREFIX) === 0; })
    .sort();
  return snaps.length ? snaps[snaps.length - 1] : null;
}

// ---------- diffing ---------------------------------------------------------------

/** 2D array -> {key: visit} keyed on date|site (with #2, #3 for repeat visits). */
function indexVisits_(data) {
  if (!data) return {};
  var hdr = data[0].map(String);
  var c = function (n) { return hdr.indexOf(n); };
  var staffCols = hdr.map(function (h, i) {
    return /^Staff\s*\d*$/i.test(String(h).trim()) ? i : -1;
  }).filter(function (i) { return i >= 0; });
  var out = {}, seen = {};
  data.slice(1).forEach(function (r) {
    var date = String(r[c('Date')] || '').trim(), site = String(r[c('Site')] || '').trim();
    if (!date || !site) return;
    var base = date + '|' + site;
    seen[base] = (seen[base] || 0) + 1;
    var key = seen[base] > 1 ? base + '#' + seen[base] : base;
    out[key] = {
      date: date, site: site,
      day: c('Day') >= 0 ? r[c('Day')] : '',
      notes: c('Notes') >= 0 ? String(r[c('Notes')] || '') : '',
      hours: c('Site Hours') >= 0 ? String(r[c('Site Hours')] || '') : '',
      sup: c('Supervisor') >= 0 ? String(r[c('Supervisor')] || '') : '',
      atm: c('ATM') >= 0 ? String(r[c('ATM')] || '') : '',
      crew: staffCols.map(function (i) { return String(r[i] || '').trim(); })
                     .filter(function (s) { return s; }),
      short: c('Short') >= 0 ? String(r[c('Short')] || '') : ''
    };
  });
  return out;
}

function fmtDay_(v) {
  var s = String(v.date), m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/), d;
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) d = new Date(+iso[1], +iso[2] - 1, +iso[3]);
  else if (m) d = new Date(+m[3], +m[1] - 1, +m[2]);
  else return v.date;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE d MMM');
}

function shortName_(s) {
  var cut = String(s).indexOf(' (');
  return (cut > 0 ? String(s).slice(0, cut) : String(s));
}

/** before/after 2D arrays -> rows for the Roster Changes tab. */
function diffRosters_(before, after) {
  var A = indexVisits_(before), B = indexVisits_(after);
  var rows = [], keys = {};
  Object.keys(A).forEach(function (k) { keys[k] = 1; });
  Object.keys(B).forEach(function (k) { keys[k] = 1; });

  var counts = { added: 0, removed: 0, changed: 0, people: 0 };

  Object.keys(keys).sort().forEach(function (k) {
    var a = A[k], b = B[k];

    if (!a && b) {
      counts.added++;
      rows.push(['Visit added', fmtDay_(b), shortName_(b.site), '',
                 '', [b.sup ? 'sup ' + b.sup : '', b.atm ? 'ATM ' + b.atm : '']
                        .concat(b.crew).filter(String).join(', ')]);
      return;
    }
    if (a && !b) {
      counts.removed++;
      rows.push(['Visit removed', fmtDay_(a), shortName_(a.site), '',
                 [a.sup ? 'sup ' + a.sup : '', a.atm ? 'ATM ' + a.atm : '']
                    .concat(a.crew).filter(String).join(', '), '']);
      return;
    }

    var visitChanged = false;
    [['Supervisor', 'sup'], ['ATM', 'atm'], ['Site Hours', 'hours'], ['Notes', 'notes']]
      .forEach(function (f) {
        if (a[f[1]] !== b[f[1]]) {
          visitChanged = true;
          rows.push([f[0] + ' changed', fmtDay_(b), shortName_(b.site), '',
                     a[f[1]] || '(none)', b[f[1]] || '(none)']);
        }
      });

    var gone = a.crew.filter(function (s) { return b.crew.indexOf(s) < 0; });
    var came = b.crew.filter(function (s) { return a.crew.indexOf(s) < 0; });
    if (gone.length || came.length) {
      visitChanged = true;
      counts.people += gone.length + came.length;
      rows.push(['Crew changed', fmtDay_(b), shortName_(b.site),
                 (came.length ? came.length + ' in' : '') +
                 (came.length && gone.length ? ', ' : '') +
                 (gone.length ? gone.length + ' out' : ''),
                 a.crew.join(', ') || '(none)', b.crew.join(', ') || '(none)']);
    }
    if (a.short !== b.short) {
      visitChanged = true;
      rows.push(['Shortfall changed', fmtDay_(b), shortName_(b.site), '',
                 a.short || '0', b.short || '0']);
    }
    if (visitChanged) counts.changed++;
  });

  return { rows: rows, counts: counts };
}

/** Per-person view: who gained or lost work. Useful now, essential for Phase 8. */
function perPersonChanges_(before, after) {
  var A = indexVisits_(before), B = indexVisits_(after);
  function assignments(idx) {
    var map = {};
    Object.keys(idx).forEach(function (k) {
      var v = idx[k];
      var all = v.crew.slice();
      if (v.sup) all.push(v.sup);
      if (v.atm) all.push(v.atm);
      all.forEach(function (p) {
        (map[p] = map[p] || {})[v.date + '|' + v.site] = v;
      });
    });
    return map;
  }
  var a = assignments(A), b = assignments(B), out = [];
  var people = {};
  Object.keys(a).forEach(function (p) { people[p] = 1; });
  Object.keys(b).forEach(function (p) { people[p] = 1; });
  Object.keys(people).sort().forEach(function (p) {
    var was = a[p] || {}, now = b[p] || {};
    var lost = Object.keys(was).filter(function (k) { return !now[k]; });
    var got = Object.keys(now).filter(function (k) { return !was[k]; });
    if (!lost.length && !got.length) return;
    var describe = function (keys, src) {
      return keys.map(function (k) {
        return fmtDay_(src[k]) + ' ' + shortName_(src[k].site);
      }).join('; ');
    };
    out.push(['Staff change', '', p,
              (got.length ? got.length + ' added' : '') +
              (got.length && lost.length ? ', ' : '') +
              (lost.length ? lost.length + ' removed' : ''),
              describe(lost, was), describe(got, now)]);
  });
  return out;
}

/** Write the Roster Changes tab. Returns the number of change rows. */
function writeChanges_(ss, before, after, note) {
  var d = diffRosters_(before, after);
  var people = perPersonChanges_(before, after);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE d MMM, h:mm a');

  var header = ['Change', 'When', 'Site / Staff', 'Detail', 'Was', 'Now'];
  var summary = [
    ['SUMMARY', stamp, note || 'Roster updated', '',
     '', d.counts.added + ' visits added, ' + d.counts.removed + ' removed, ' +
         d.counts.changed + ' changed'],
  ];
  if (!d.rows.length && !people.length) {
    summary.push(['OK', '', 'No differences', '', '', 'The new roster matches the previous one.']);
  }
  var data = [header].concat(summary, d.rows,
              people.length ? [['', '', '', '', '', '']].concat(people) : []);

  var sh = ss.getSheetByName(CHANGES_TAB);
  if (sh) sh.clear(); else sh = ss.insertSheet(CHANGES_TAB);
  sh.getRange(1, 1, data.length, header.length).setValues(data);

  var style = { 'Visit added': '#0b8043', 'Visit removed': '#d50000', 'Crew changed': '#f6bf26',
                'Supervisor changed': '#f6bf26', 'ATM changed': '#f6bf26',
                'Shortfall changed': '#f6bf26', 'Site Hours changed': '#e8eaed',
                'Notes changed': '#e8eaed', 'Staff change': '#3d85c6',
                'SUMMARY': '#212121', 'OK': '#0b8043' };
  var bgs = [], fgs = [];
  for (var r = 0; r < data.length; r++) {
    var tag = String(data[r][0]);
    var col = r === 0 ? HDR_BG : (style[tag] || '#ffffff');
    var dark = ['#0b8043', '#d50000', '#3d85c6', '#212121', HDR_BG].indexOf(col) >= 0;
    bgs.push([col, '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']);
    fgs.push([dark ? '#ffffff' : '#000000', '#000000', '#000000', '#000000', '#000000', '#000000']);
  }
  sh.getRange(1, 1, data.length, header.length).setBackgrounds(bgs).setFontColors(fgs)
    .setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 260); sh.setColumnWidth(6, 260);
  try { protectRendered_(sh, CHANGES_TAB); } catch (e) {}
  return d.rows.length + people.length;
}

// ---------- menu actions -----------------------------------------------------------

/** Take a snapshot on demand (e.g. before hand-editing). */
function snapshotRosterNow() {
  var data = snapshotMasterRoster_();
  if (!data) throw new Error('Nothing to snapshot — the "' + DATA_TAB + '" tab is empty.');
  SpreadsheetApp.getActive().toast('Snapshot saved (' + (data.length - 1) + ' visits). ' +
    'Use "Restore last snapshot" to undo changes made after this point.', 'Apunga Roster', 10);
}

/** Compare the current Master Roster against the newest snapshot. */
function showChangesSinceSnapshot() {
  var ss = SpreadsheetApp.getActive();
  var name = latestSnapshotName_(ss);
  if (!name) throw new Error('No snapshots yet. Run the optimiser, or use "Snapshot roster now".');
  var before = snapRead_(ss, name), after = snapRead_(ss, DATA_TAB);
  if (!after) throw new Error('The "' + DATA_TAB + '" tab is empty.');
  var n = writeChanges_(ss, before, after, 'Compared against ' + name);
  ss.setActiveSheet(ss.getSheetByName(CHANGES_TAB));
  ss.toast(n ? n + ' change(s) since ' + name : 'No changes since ' + name, 'Apunga Roster', 10);
}

/** Put the newest snapshot back into the Master Roster (undo an overwrite). */
function restoreLastSnapshot() {
  var ss = SpreadsheetApp.getActive();
  var name = latestSnapshotName_(ss);
  if (!name) throw new Error('No snapshots available.');
  var data = snapRead_(ss, name);
  if (!data) throw new Error('Snapshot "' + name + '" is empty.');

  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  if (ui) {
    var answer = ui.alert('Restore ' + name + '?',
      'This replaces the current Master Roster with the snapshot taken at ' +
      name.replace(SNAP_PREFIX, '') + ' (' + (data.length - 1) + ' visits). ' +
      'The current version will be snapshotted first, so this is reversible.',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }

  snapshotMasterRoster_();                       // so the restore itself is undoable
  var sh = ss.getSheetByName(DATA_TAB);
  if (sh) sh.clear(); else sh = ss.insertSheet(DATA_TAB);
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
  var dcol = data[0].indexOf('Date');
  if (dcol >= 0 && data.length > 1)
    sh.getRange(2, dcol + 1, data.length - 1, 1).setNumberFormat('ddd d mmm yyyy');
  sh.setFrozenRows(1);

  renderCalendar();
  renderStaffView();
  ss.toast('Restored ' + name + '. Calendar and staff view re-rendered.', 'Apunga Roster', 12);
}