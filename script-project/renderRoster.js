/**
 * RenderRoster.gs — Apunga roster write-back (Phase 6).
 *
 * Reads the wide "Master Roster" tab (one row per visit) and renders:
 *   1. "Master Calendar" — Mon–Fri day columns, one block per site visit,
 *      colour-coded by council (LGA), supervisor bold, crew listed, SHORT flags.
 *   2. "Staff View"      — 2-week rolling grid: one row per staff member,
 *      one column per workday, cell = site, plus a total-hours column.
 *
 * Workflow:
 *   - Apunga Roster > Run optimiser fills the "Master Roster" tab and renders
 *   - edit "Master Roster" by hand, then Apunga Roster > Render both to update
 *     the calendar and staff view from it
 *
 * Rendering only touches this spreadsheet — no new OAuth scopes needed
 * (spreadsheets.currentonly covers it). Never writes to SITES/Staff/Demand.
 */

const DATA_TAB  = 'Master Roster';
const RUNLOG_TAB = 'Run Log';     // plain-English report from the optimiser
const SITES_TAB = 'Sites';   // for council colours (adjust if your tab is named differently)
const CAL_TAB   = 'Master Calendar';
const STAFF_TAB = 'Staff View';
const DAY_ORDER = ['Mon', 'Tue', 'Wed', 'Thur', 'Fri'];
// Sat/Sun columns are added automatically, and only when weekend visits exist.
// Fixed council -> colour map (keyed to the old paper roster for continuity).
// Neighbours nudged apart slightly so similar hues stay legible.
const LGA_COLOURS = {
  'bayside':                  '#c1440e',  // burnt orange
  'blacktown':                '#1a3e72',  // dark blue
  'canada bay':               '#f6a04d',  // light orange
  'canterbury-bankstown':     '#e5528f',  // pink
  'city of sydney':           '#12a4a4',  // teal
  'cumberland':               '#c9a227',  // dark yellow
  'fairfield':                '#0091ff',  // electric blue
  'greater sydney parklands': '#6a4bcc',  // blue-purple
  'hornsby':                  '#7c4a2d',  // brown
  'hunters hill':             '#7fbfe0',  // light blue
  'inner west':               '#8e24aa',  // purple
  'ku-ring-gai':              '#2e9e4f',  // green
  'lane cove':                '#9e9e9e',  // grey
  'north sydney':             '#c26a35',  // brown-orange
  'northern beaches':         '#f57c00',  // orange
  'npws':                     '#1b5e20',  // dark green (all NPWS sub-regions)
  'sutherland':               '#3f6fd8',  // blue
  'waverley':                 '#aacc22',  // lime green
  'willoughby':               '#808000',  // olive
};
const LGA_ALIASES = {          // normalise data-sheet variants onto the keys above
  'blacktown council': 'blacktown', 'cumberland council': 'cumberland',
  'kuringai': 'ku-ring-gai', 'ku ring gai': 'ku-ring-gai',
  'waverly': 'waverley', "hunter's hill": 'hunters hill',
};
const LGA_FALLBACK = '#616161'; // anything unmapped (e.g. Private Contract, blank)

function lgaColour_(lga) {
  let k = String(lga || '').trim().toLowerCase();
  if (k.indexOf('npws') === 0) return LGA_COLOURS['npws'];   // NPWS, NPWS - Garigal, ...
  if (LGA_ALIASES[k]) k = LGA_ALIASES[k];
  return LGA_COLOURS[k] || LGA_FALLBACK;
}
const HDR_BG = '#212121', HDR_FG = '#ffffff', SHORT_BG = '#d50000';

/** black or white text depending on background luminance */
function textOn_(hex) {
  const n = parseInt(hex.slice(1), 16);
  const lum = 0.299*((n>>16)&255) + 0.587*((n>>8)&255) + 0.114*(n&255);
  return lum > 150 ? '#000000' : '#ffffff';
}
/** blend a hex colour toward white (0..1, 1 = white) for crew-row tints */
function tint_(hex, f) {
  const n = parseInt(hex.slice(1), 16);
  const ch = v => Math.round(v + (255 - v) * f).toString(16).padStart(2, '0');
  return '#' + ch((n>>16)&255) + ch((n>>8)&255) + ch(n&255);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Apunga Roster')
    .addItem('How to use this tool', 'showHelp')
    .addSeparator()
    .addItem('Import Demand from Google Calendar…', 'importDemandFromCalendar')
    .addItem('Apply Demand import', 'applyDemandImport')
    .addItem('Build calendar history', 'buildCalendarHistory')
    .addItem('Build site analysis', 'buildSiteAnalysis')
    .addItem('Apply renames (Renames tab)', 'applyRenames')
    .addItem('Clear finished leave', 'clearFinishedLeave')
    .addSeparator()
    .addItem('Run optimiser', 'runOptimiser')
    .addSeparator()
    .addItem('Render both', 'renderBoth')
    .addItem('Render master calendar', 'renderCalendar')
    .addItem('Render staff view', 'renderStaffView')
    .addItem('Format run log', 'formatRunLog')
    .addSeparator()
    .addItem('Snapshot roster now', 'snapshotRosterNow')
    .addItem('Show changes since snapshot', 'showChangesSinceSnapshot')
    .addItem('Restore last snapshot', 'restoreLastSnapshot')
    .addSeparator()
    .addItem('Sync roster to calendar', 'syncRosterToCalendar')
    .addSeparator()
    .addItem('Build hours variance report', 'buildHoursVariance')
    .addItem('Show hours app link', 'showHoursAppUrl')
    .addItem('Check hours app setup', 'checkHoursApp')
    .addItem('Check site summary setup', 'checkSiteSummary')
    .addItem('Check calendar sync', 'checkCalendarSync')
    .addItem('Clear roster events from calendar', 'clearRosterCalendar')
    .addSeparator()
    .addItem('Check optimiser connection', 'checkOptimiserConnection')
    .addToUi();
}
function renderBoth() {
  renderCalendar();
  renderStaffView();
  try { formatRunLog(); } catch (e) { /* no Run Log pasted yet — fine */ }
}

// ---------- run log --------------------------------------------------------------

const SEV_STYLE = {
  'ACTION NEEDED': { bg: '#d50000', fg: '#ffffff' },
  'CHECK':         { bg: '#f6bf26', fg: '#000000' },
  'INFO':          { bg: '#e8eaed', fg: '#000000' },
  'OK':            { bg: '#0b8043', fg: '#ffffff' },
  'SUMMARY':       { bg: '#3d85c6', fg: '#ffffff' },
};

/** Colour-codes and sizes the pasted Run Log so managers can scan it. */
function formatRunLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RUNLOG_TAB);
  if (!sh) throw new Error('No "' + RUNLOG_TAB + '" tab. Paste RunLog.csv into a tab with that name.');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('"' + RUNLOG_TAB + '" is empty — paste RunLog.csv in first.');

  const vals = sh.getRange(1, 1, last, 4).getValues();
  const bgs = [], fgs = [], wts = [];
  bgs.push([HDR_BG, HDR_BG, HDR_BG, HDR_BG]);
  fgs.push([HDR_FG, HDR_FG, HDR_FG, HDR_FG]);
  wts.push(['bold', 'bold', 'bold', 'bold']);

  for (let r = 1; r < last; r++) {
    const sev = String(vals[r][0]).trim();
    const st = SEV_STYLE[sev] || { bg: '#ffffff', fg: '#000000' };
    bgs.push([st.bg, '#ffffff', '#ffffff', '#ffffff']);
    fgs.push([st.fg, '#000000', '#000000', '#000000']);
    wts.push(['bold', 'normal', sev === 'ACTION NEEDED' ? 'bold' : 'normal', 'normal']);
  }
  const rng = sh.getRange(1, 1, last, 4);
  rng.setBackgrounds(bgs).setFontColors(fgs).setFontWeights(wts)
     .setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 380); sh.setColumnWidth(4, 420);
  sh.setFrozenRows(1);
  protectRendered_(sh, RUNLOG_TAB);

  const acts = vals.slice(1).filter(r => String(r[0]).trim() === 'ACTION NEEDED').length;
  ss.toast(acts ? acts + ' item(s) need action — see the Run Log tab.' : 'Run log clean.',
           'Apunga Roster', 8);
  ss.setActiveSheet(sh);
}

// ---------- data ----------------------------------------------------------------

function readVisits_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DATA_TAB);
  if (!sh) {
    throw new Error('No "' + DATA_TAB + '" tab found. Run Apunga Roster > Run optimiser to ' +
                    'create it.');
  }
  if (sh.getLastRow() < 2) {
    throw new Error('The "' + DATA_TAB + '" tab has no roster rows. Run Apunga Roster > ' +
                    'Run optimiser first.');
  }
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const c = n => hdr.indexOf(n);
  ['Date','Day','Site','Supervisor'].forEach(n => {
    if (c(n) < 0) throw new Error(DATA_TAB + ' is missing the "' + n + '" column.');
  });
  const staffCols = hdr.map((h, i) => /^Staff\s*\d*$/i.test(h.trim()) ? i : -1).filter(i => i >= 0);
  const lga = siteLgaMap_();
  // a site edited by hand ('Kimberley & Crete') isn't in Sites: colour it by the
  // first part that is, using the calendar-import matcher
  let ctx = null;
  const lgaOf = site => {
    if (lga[site]) return lga[site];
    try {
      ctx = ctx || impContext_(Object.keys(lga), [], impReadAliases_(SpreadsheetApp.getActive()));
      for (const part of site.split(/\s*(?:&|\+|\/| and )\s*/)) {
        const m = impMatchSite_(part, ctx);
        if (m.site && lga[m.site]) return lga[m.site];
      }
    } catch (e) { /* importer not available: leave uncoloured */ }
    return '';
  };
  return vals.slice(1).filter(r => r[c('Date')] !== '' && r[c('Site')] !== '').map(r => {
    const site = String(r[c('Site')]).trim();
    return {
      date: normDate_(r[c('Date')]), day: String(r[c('Day')]).trim(), site: site,
      lga: lgaOf(site), notes: String(c('Notes') >= 0 ? r[c('Notes')] || '' : '').trim(),
      siteHours: String(c('Site Hours') >= 0 ? r[c('Site Hours')] || '' : '').trim(),
      sup: String(r[c('Supervisor')] || '').trim(),
      atm: String(c('ATM') >= 0 ? r[c('ATM')] || '' : '').trim(),
      crew: staffCols.map(i => String(r[i] || '').trim()).filter(Boolean),
      short: Number(c('Short') >= 0 ? r[c('Short')] : 0) || 0,
    };
  });
}

/** site name -> LGA, read live from the SITES tab so Master Roster stays clean */
function siteLgaMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SITES_TAB);
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const iN = hdr.findIndex(h => h.indexOf('Roster site name') === 0), iL = hdr.indexOf('LGA');
  if (iN < 0 || iL < 0) return {};
  const m = {};
  vals.slice(1).forEach(r => {
    const n = String(r[iN] || '').trim();
    if (n) m[n] = String(r[iL] || '').trim();
  });
  return m;
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);       // M/D/YYYY from the sheet
  if (us) return us[3] + '-' + ('0' + us[1]).slice(-2) + '-' + ('0' + us[2]).slice(-2);
  return s.slice(0, 10);
}


function shortSite_(s) {
  const cut = s.indexOf(' (');
  return (cut > 0 ? s.slice(0, cut) : s).slice(0, 26);
}

function mondayOf_(iso) {              // ISO date string -> ISO date of that week's Monday
  const d = new Date(iso + 'T00:00:00');
  const off = (d.getDay() + 6) % 7;    // Mon=0 ... Sun=6
  d.setDate(d.getDate() - off);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function addDays_(iso, n) {
  const d = new Date(iso + 'T00:00:00'); d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function fmtDM_(iso) {
  return Utilities.formatDate(new Date(iso + 'T00:00:00'), Session.getScriptTimeZone(), 'EEE d MMM');
}

function freshSheet_(name, ss) {
  let sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  return sh;
}

function protectRendered_(sh, what) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sh.protect().setWarningOnly(true)
    .setDescription(what + ' is generated — edit the ' + DATA_TAB + ' tab and re-render instead.');
}


// ---------- master calendar ------------------------------------------------------

function renderCalendar() {
  const ss = SpreadsheetApp.getActive();
  const vis = readVisits_();
  const colours = {};
  vis.forEach(v => { if (v.lga) colours[v.lga] = lgaColour_(v.lga); });
  const sh = freshSheet_(CAL_TAB, ss);

  // legend — label on its own row, then councils wrapped LEGEND_COLS per row so
  // the block stays compact instead of running off to the right.
  const LEGEND_COLS = 5;
  const lgas = Object.keys(colours).sort();
  sh.getRange(1, 1).setValue('Council legend:').setFontWeight('bold').setFontSize(9);

  const nLegendRows = Math.max(1, Math.ceil(lgas.length / LEGEND_COLS));
  const lVals = [], lBg = [], lFg = [];
  for (let r = 0; r < nLegendRows; r++) {
    const vRow = [], bRow = [], fRow = [];
    for (let c = 0; c < LEGEND_COLS; c++) {
      const l = lgas[r * LEGEND_COLS + c];
      vRow.push(l || '');
      bRow.push(l ? colours[l] : '#ffffff');
      fRow.push(l ? textOn_(colours[l]) : '#000000');
    }
    lVals.push(vRow); lBg.push(bRow); lFg.push(fRow);
  }
  sh.getRange(2, 1, nLegendRows, LEGEND_COLS).setValues(lVals).setBackgrounds(lBg)
    .setFontColors(lFg).setFontSize(9).setWrap(true).setHorizontalAlignment('center');

  // week blocks
  const weeks = {};
  vis.forEach(v => { const w = mondayOf_(v.date); (weeks[w] = weeks[w] || []).push(v); });

  // show Sat/Sun columns only when the month actually has weekend visits
  let nCols = 5;
  vis.forEach(v => {
    const dow = new Date(v.date + 'T00:00:00').getDay();      // 0=Sun, 6=Sat
    if (dow === 6) nCols = Math.max(nCols, 6);
    if (dow === 0) nCols = 7;
  });
  let row = 2 + nLegendRows + 1;   // label + legend rows + a gap
  const mondays = Object.keys(weeks).sort();
  const rank = staffRank_();
  const office = mondays.length ? officeEntries_(mondays[0], addDays_(mondays[mondays.length - 1], 6)) : {};
  mondays.forEach(monday => {
    const wvis = weeks[monday];
    // build one cell-list per weekday column
    const cols = Array.from({ length: nCols }, (_, i) => {
      const dIso = addDays_(monday, i);
      const dv = wvis.filter(v => v.date === dIso)
                     .sort((a, b) => a.site.localeCompare(b.site));
      const cells = [{ t: fmtDM_(dIso), bg: HDR_BG, fg: HDR_FG, bold: true }];
      dv.forEach(v => {
        const base = lgaColour_(v.lga), light = tint_(base, 0.82);
        const title = v.site + (v.siteHours ? ' (' + v.siteHours + ')' : '');
        cells.push({ t: title, bg: base, fg: textOn_(base), bold: true });
        if (v.notes) cells.push({ t: '» ' + v.notes, bg: light });
        cells.push({ t: v.sup ? 'Sup: ' + v.sup : '(no supervisor)', bg: light, bold: true });
        if (v.atm) cells.push({ t: 'ATM: ' + v.atm, bg: light, bold: true });
        bySeniority_(v.crew, rank).forEach(nm => cells.push({ bg: light, t: nm }));
        if (v.short > 0) cells.push({ t: '⚠ SHORT ' + v.short, bg: SHORT_BG, fg: '#ffffff', bold: true });
        cells.push({ t: '' });                       // spacer between site blocks
      });
      // office days, training etc. from the x OFFICE calendar, under the visits
      (office[dIso] || []).forEach(t => cells.push({ t: '🏢 ' + t, bg: OFFICE_BG }));
      return cells;
    });
    const h = Math.max.apply(null, cols.map(cl => cl.length));
    const vals = [], bgs = [], wts = [], fgs = [];
    for (let r = 0; r < h; r++) {
      vals.push(cols.map(cl => (cl[r] ? cl[r].t : '')));
      bgs.push(cols.map(cl => (cl[r] && cl[r].bg) ? cl[r].bg : '#ffffff'));
      wts.push(cols.map(cl => (cl[r] && cl[r].bold) ? 'bold' : 'normal'));
      fgs.push(cols.map(cl => (cl[r] && cl[r].fg) ? cl[r].fg : '#000000'));
    }
    sh.getRange(row, 1, h, nCols)
      .setValues(vals).setBackgrounds(bgs).setFontWeights(wts).setFontColors(fgs)
      .setVerticalAlignment('middle').setWrap(true);
    row += h + 1;                                    // gap row between weeks
  });

  for (let cIdx = 1; cIdx <= nCols; cIdx++) sh.setColumnWidth(cIdx, 190);
  sh.setFrozenRows(0);
  sh.setHiddenGridlines(true);
  protectRendered_(sh, CAL_TAB);
  ss.setActiveSheet(sh);
}

/**
 * Office days, training etc. from management's "x OFFICE" Google Calendar:
 * {isoDate: [titles]}. Empty if the calendar can't be read (renders without them).
 */
const OFFICE_CAL_NAME = 'x OFFICE';
const RR_STAFF_SRC = 'Staff';          // the staff list (STAFF_TAB is the rendered Staff View)

/**
 * StaffID -> seniority rank for listing crews: spare supervisors, then ATMs, then
 * field staff by status (Qualified, Enrolled, Unqualified/New), unknown last.
 */
function staffRank_() {
  const rank = {};
  try {
    const v = SpreadsheetApp.getActive().getSheetByName(RR_STAFF_SRC).getDataRange().getValues();
    let h = -1;
    for (let i = 0; i < 15 && h < 0; i++) if (v[i].map(String).indexOf('StaffID') >= 0) h = i;
    if (h < 0) return rank;
    const hd = v[h].map(x => String(x).trim());
    const ci = hd.indexOf('StaffID'), cc = hd.indexOf('Category'), cs = hd.indexOf('Cat. Status');
    v.slice(h + 1).forEach(r => {
      const id = String(r[ci] || '').trim(); if (!id) return;
      const cat = String(r[cc] || '').trim().toLowerCase(), st = String(cs >= 0 ? r[cs] : '').trim().toLowerCase();
      rank[id] = cat === 'supervisor' ? 0 : cat === 'atm' ? 1
               : st.indexOf('qualified') === 0 ? 2 : st.indexOf('enrolled') === 0 ? 3 : 4;
    });
  } catch (err) { Logger.log('Staff ranks not read: ' + err); }
  return rank;
}
function bySeniority_(names, rank) {
  return names.map((n, i) => [n, i])
    .sort((a, b) => ((rank[a[0]] ?? 5) - (rank[b[0]] ?? 5)) || (a[1] - b[1]))
    .map(x => x[0]);
}
/** StaffIDs named as whole words in an office-calendar title ('ACT - Kim, Rosie, Riley(?)'). */
function namesIn_(title, ids) {
  return ids.filter(id => new RegExp('(^|[^A-Za-z])' + id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
                                     '($|[^A-Za-z])', 'i').test(title));
}
const OFFICE_BG = '#d9d9d9';
function officeEntries_(fromIso, toIso) {
  const out = {};
  try {
    const cals = CalendarApp.getCalendarsByName(OFFICE_CAL_NAME);
    if (!cals || !cals.length) return out;
    const tz = Session.getScriptTimeZone();
    const from = new Date(fromIso + 'T00:00:00'), to = new Date(toIso + 'T00:00:00');
    to.setDate(to.getDate() + 1);
    cals[0].getEvents(from, to).forEach(ev => {
      const s = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
      const e = ev.isAllDayEvent() ? ev.getAllDayEndDate() : new Date(s.getTime() + 1);
      for (let d = new Date(s.getFullYear(), s.getMonth(), s.getDate()); d < e; d.setDate(d.getDate() + 1)) {
        const iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        (out[iso] = out[iso] || []).push(ev.getTitle().trim());
      }
    });
  } catch (err) { Logger.log('Office calendar not read: ' + err); }
  return out;
}

// ---------- staff view (2-week rolling) -------------------------------------------

function renderStaffView() {
  const ss = SpreadsheetApp.getActive();
  const vis = readVisits_();
  const sh = freshSheet_(STAFF_TAB, ss);

  const allDates = [...new Set(vis.map(v => v.date))].sort();
  const start = allDates[0], end = addDays_(start, 13);      // 2-week rolling window
  const dates = allDates.filter(d => d <= end);
  const inWin = vis.filter(v => v.date <= end);

  const byKey = {}, hoursBy = {};                            // staff|date -> {t, sup, lga}
  function put(name, v, label, hrs, isSup) {
    byKey[name + '|' + v.date] = { t: shortSite_(v.site) + label, sup: isSup, lga: v.lga };
    hoursBy[name] = (hoursBy[name] || 0) + hrs;
  }
  inWin.forEach(v => {
    if (v.sup) put(v.sup, v, ' (S)', 8, true);
    if (v.atm) put(v.atm, v, ' (A)', 7.5, false);
    v.crew.forEach(nm => put(nm, v, '', 7.5, false));
  });
  // office days / training from the x OFFICE calendar, matched to StaffIDs by name
  const allIds = Object.keys(staffRank_());
  const office = officeEntries_(start, end), offBy = {};
  Object.keys(office).forEach(d => {
    if (d > end) return;
    office[d].forEach(t => namesIn_(t, allIds).forEach(id => {
      const label = /office/i.test(t) ? 'Office' + ((t.match(/\(([^)]*)\)/) || [])[1] ? ' (' + t.match(/\(([^)]*)\)/)[1] + ')' : '')
                                      : t.split(/\s+-\s+/)[0].slice(0, 24);
      (offBy[id + '|' + d] = offBy[id + '|' + d] || []).push(label);
    }));
  });
  const staffIds = [...new Set(Object.keys(byKey).concat(Object.keys(offBy)).map(k => k.split('|')[0]))].sort();

  const header = ['Staff'].concat(dates.map(fmtDM_)).concat(['Total hrs']);
  const vals = [header], wts = [header.map(() => 'bold')];
  const bgs = [header.map(() => HDR_BG)], fgs = [header.map(() => HDR_FG)];
  staffIds.forEach((s, i) => {
    const zebra = i % 2 ? '#f3f3f3' : '#ffffff';
    const line = [s], w = ['bold'], bg = [zebra], fg = ['#000000'];
    dates.forEach(d => {
      const c = byKey[s + '|' + d], off = offBy[s + '|' + d];
      if (c) {
        const col = lgaColour_(c.lga);
        line.push(c.t + (off ? '\n🏢 ' + off.join(', ') : '')); w.push(c.sup ? 'bold' : 'normal');
        bg.push(off ? SHORT_BG : col); fg.push(off ? '#ffffff' : textOn_(col));   // red = clash
      } else if (off) {
        line.push('🏢 ' + off.join(', ')); w.push('normal'); bg.push(OFFICE_BG); fg.push('#000000');
      } else {
        line.push(''); w.push('normal'); bg.push(zebra); fg.push('#000000');
      }
    });
    line.push(hoursBy[s] || 0); w.push('normal'); bg.push(zebra); fg.push('#000000');
    vals.push(line); wts.push(w); bgs.push(bg); fgs.push(fg);
  });

  const rng = sh.getRange(1, 1, vals.length, header.length);
  rng.setValues(vals).setFontWeights(wts).setBackgrounds(bgs).setFontColors(fgs).setWrap(true)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2, 1, Math.max(1, vals.length - 1), 1).setHorizontalAlignment('left');   // names
  if (vals.length > 1) sh.setRowHeights(2, vals.length - 1, 34);                    // breathing room
  sh.setRowHeight(1, 30);
  sh.getRange(1, 1, 1, header.length).setBackground(HDR_BG).setFontColor(HDR_FG);
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 90);
  for (let cIdx = 2; cIdx <= dates.length + 1; cIdx++) sh.setColumnWidth(cIdx, 130);
  protectRendered_(sh, STAFF_TAB);
  ss.setActiveSheet(sh);
}