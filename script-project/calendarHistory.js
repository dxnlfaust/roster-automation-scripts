/**
 * calendarHistory.gs — read past months of management's council calendars into a
 * "Calendar History" tab, one row per calendar entry, parsed.
 *
 * Used for Phase 5 (monthly planning): each site's usual day, visit size,
 * frequency, supervisor and hours per month. Read only: no calendar is changed.
 *
 * Menu: Apunga Roster ▸ Build calendar history
 * Reuses the Demand-import parser (importDemand.gs) and skips the crew entries
 * this project syncs in (tagged), so they're never counted as visits.
 *
 * Kinds written:
 *   visit  — "(38) Bungaribee - Rosie" (hours counted)
 *   crew   — a crew list entry beside a visit, "(23) Sonya Marinella Ingrid" (hours NOT counted)
 *   other  — no "(hours)" at the start: office days, leave, notes (not counted)
 */

var HIST_TAB = 'Calendar History';
var HIST_MONTHS_BACK = 16;           // Jul 2025 onwards when run in Oct 2026: a full FY + this one
var HIST_COLS = ['Date', 'Day', 'Calendar', 'Kind', 'Site', 'Second site', 'Site text', 'Hours',
                 'Supervisor', 'People', 'Matched', 'Title'];

function buildCalendarHistory() {
  var ss = SpreadsheetApp.getActive(), tz = Session.getScriptTimeZone();
  var now = new Date();
  var from = new Date(now.getFullYear(), now.getMonth() - HIST_MONTHS_BACK, 1);
  var to = new Date(now.getFullYear(), now.getMonth() + 3, 1);     // include what's already planned
  var ctx = impContext_(impReadSites_(ss), impReadStaff_(ss), impReadAliases_(ss));
  var rosterId = PropertiesService.getScriptProperties().getProperty('ROSTER_CALENDAR_ID');

  var rows = [], cals = 0;
  CalendarApp.getAllCalendars().forEach(function (cal) {
    var name = cal.getName();
    if (cal.getId() === rosterId || IMP_SKIP_CALENDARS.indexOf(name) >= 0 ||
        IMP_INFO_CALENDARS.indexOf(name) >= 0) return;
    var evs = impEvents_(cal, from, to);
    if (evs.length) cals++;
    evs.forEach(function (e) { rows.push(histRow_(e, ctx, tz)); });
  });
  rows.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[2].localeCompare(b[2]); });

  var sh = ss.getSheetByName(HIST_TAB) || ss.insertSheet(HIST_TAB);
  sh.clear();
  sh.getRange(1, 1, 1, HIST_COLS.length).setValues([HIST_COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, HIST_COLS.length).setNumberFormat('@');   // keep text as text
    sh.getRange(2, 1, rows.length, HIST_COLS.length).setValues(rows);
  }
  var visits = rows.filter(function (r) { return r[3] === 'visit'; });
  var matched = visits.filter(function (r) { return r[10] === 'yes'; });
  var hrs = function (list) { return list.reduce(function (a, r) { return a + (Number(r[7]) || 0); }, 0); };
  ss.toast(rows.length + ' entries from ' + cals + ' calendars (' +
    Utilities.formatDate(from, tz, 'MMM yyyy') + ' on). Visits: ' + visits.length + ', ' +
    Math.round(100 * hrs(matched) / Math.max(1, hrs(visits))) + '% of hours matched to a site.',
    'Calendar history', 15);
}

function histRow_(e, ctx, tz) {
  var date = Utilities.formatDate(e.date, tz, 'yyyy-MM-dd');
  var day = IMP_WEEKDAY[e.date.getDay()];
  var raw = String(e.title || '').replace(/ /g, ' ').trim();
  // last year's combined style: "(30.5) Chestnut - (30.5) Sonya Mel Claire Thomas*"
  var title = raw.replace(/\s-\s*\(\d[\d.+? ]*\)\s*/, ' - ');
  var p = impParseTitle_(title, ctx);
  if (p.kind === 'skip') return [date, day, e.calendar, 'other', '', '', '', '', '', '', '', raw];

  var lead = title.match(/^\s*\(([^)]*)\)/), lh = lead && impParseHours_(lead[1]);
  var leadHours = lh ? lh.hours : p.hours;      // a crew list's own total, not people's hours added up
  var hasSep = /\s-\s|\s-$|-\s/.test(title.replace(/^\s*\([^)]*\)\s*/, ''));
  var words = title.replace(/^\s*\([^)]*\)\s*/, '').replace(/\([^)]*\)/g, ' ').replace(/[*?]/g, ' ')
    .split(/[\s,&\/]+/).filter(Boolean);
  // a crew list on its own: no "Site - Supervisor" dash, no site found, every word a Name
  // ...or it starts with a staff member's name: "(30.5) Dan Marinella Sasha" is a crew, not Wandangalli
  if (!hasSep && words.length >= 2 && ctx.staff[words[0].toLowerCase()]) {
    return [date, day, e.calendar, 'crew', '', '', '', leadHours, '', words.join(' '), '', raw];
  }
  // ...or names with their own hours: "(38+2)Tim(8+2) Matthew Bronte Sasha"
  if (!hasSep && /[A-Z][a-z]+\s*\(\d/.test(title.replace(/^\s*\([^)]*\)\s*/, ''))) {
    return [date, day, e.calendar, 'crew', '', '', '', leadHours, '', words.join(' '), '', raw];
  }
  if (!hasSep && !p.site && words.length >= 2 && words.length <= 9 &&
      words.every(function (w) { return /^[A-Z][A-Za-z'.-]*$/.test(w); })) {
    return [date, day, e.calendar, 'crew', '', '', '', leadHours, '', words.join(' '), '', raw];
  }
  var people = (title.split(/\s+-\s*|\s*-\s+/).pop() || '').replace(/\([^)]*\)/g, ' ').replace(/[*?]/g, ' ')
    .split(/[\s,&\/]+/).filter(Boolean);
  return [date, day, e.calendar, 'visit', p.site || '', p.second || '', p.siteText || '', p.hours,
          p.sup || '', hasSep ? people.join(' ') : '', p.site ? 'yes' : 'no', raw];
}
