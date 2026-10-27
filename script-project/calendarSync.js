/**
 * CalendarSync.gs — push the roster's crew lists into Google Calendar (one way).
 *
 * Management's council calendars already hold each visit as "(hours) Site -
 * Supervisor" (the Demand import reads them). This adds ONE entry per visit, the
 * crew list, into the SAME council calendar as that visit, so it shares the
 * colour and sits beside the site entry:   Nina Matt Riley Zac
 * Visits with no matching council entry go to a separate "Apunga Roster" calendar.
 *
 * Every entry this script creates is tagged, and only tagged entries are ever
 * updated or deleted. Management's own entries are never touched. Creating an
 * entry sends no invitations (no guests).
 *
 * Menu: Apunga Roster > Sync roster to calendar
 *
 * REQUIRES the calendar scope in appsscript.json:
 *   "https://www.googleapis.com/auth/calendar"
 */

var CAL_NAME = 'Apunga Roster';
var CAL_PROP = 'ROSTER_CALENDAR_ID';
var CAL_TAG = 'apungaKey';          // marks events this script owns

/** Region -> one of Calendar's fixed colours. Sites with no region stay default. */
var CAL_COLOURS = {
  'north':   CalendarApp.EventColor.BLUE,
  'south':   CalendarApp.EventColor.GREEN,
  'east':    CalendarApp.EventColor.MAUVE,
  'west':    CalendarApp.EventColor.ORANGE,
  'central': CalendarApp.EventColor.RED
};

/** The dedicated calendar, created on first run. */
function rosterCalendar_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(CAL_PROP);
  if (id) {
    var existing = CalendarApp.getCalendarById(id);
    if (existing) return existing;
    props.deleteProperty(CAL_PROP);              // deleted in Calendar — make a new one
  }
  var found = CalendarApp.getCalendarsByName(CAL_NAME);
  var cal = found && found.length ? found[0] : CalendarApp.createCalendar(CAL_NAME, {
    summary: 'Roster generated from the Automation Ops sheet — do not edit here',
    color: '#3d85c6'
  });
  props.setProperty(CAL_PROP, cal.getId());
  return cal;
}

/** site name -> region, read from the SITES tab. */
function siteRegionMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SITES_TAB);
  if (!sh) return {};
  var vals = sh.getDataRange().getValues(), hdr = vals[0].map(String);
  var iN = hdr.findIndex(function (h) { return String(h).indexOf('Roster site name') === 0; });
  var iR = hdr.indexOf('Region');
  if (iN < 0 || iR < 0) return {};
  var m = {};
  vals.slice(1).forEach(function (r) {
    var n = String(r[iN] || '').trim();
    if (n) m[n] = String(r[iR] || '').trim().toLowerCase();
  });
  return m;
}

/** Master Roster -> the events we want, keyed so a re-sync can match them. */
/** Calendar's 11 event colours (current palette) -> id, for matching a council calendar's colour. */
var CAL_EVENT_HEX = { '1': '#7986cb', '2': '#33b679', '3': '#8e24aa', '4': '#e67c73', '5': '#f6bf26',
  '6': '#f4511e', '7': '#039be5', '8': '#616161', '9': '#3f51b5', '10': '#0b8043', '11': '#d50000' };
function nearestEventColour_(hex) {
  var h = String(hex || '').replace('#', ''); if (h.length !== 6) return null;
  var rgb = function (x) { return [0, 2, 4].map(function (i) { return parseInt(x.substr(i, 2), 16); }); };
  var a = rgb(h), best = null, bd = 1e9;
  Object.keys(CAL_EVENT_HEX).forEach(function (id) {
    var b = rgb(CAL_EVENT_HEX[id].slice(1));
    var d = Math.pow(a[0] - b[0], 2) + Math.pow(a[1] - b[1], 2) + Math.pow(a[2] - b[2], 2);
    if (d < bd) { bd = d; best = id; }
  });
  return best;
}
function calSiteKey_(s) { return String(s).replace(/\s*\n\s*/g, ' / ').replace(/[\u00a0\s]+/g, ' ').trim().toLowerCase(); }

/**
 * date|site -> the council calendar management booked that visit in, found by
 * reading their calendars with the same parser the Demand import uses.
 * Returns {map: {key: Calendar}, cals: [Calendar]} (cals = every calendar searched).
 */
function councilHomes_(fromIso, toIso) {
  var map = {}, cals = [];
  var ss = SpreadsheetApp.getActive();
  var ctx = impContext_(impReadSites_(ss), impReadStaff_(ss), impReadAliases_(ss));
  var rosterId = PropertiesService.getScriptProperties().getProperty(CAL_PROP);
  var from = calDate_(fromIso), to = calDate_(toIso); to.setDate(to.getDate() + 1);
  var tz = Session.getScriptTimeZone();
  CalendarApp.getAllCalendars().forEach(function (cal) {
    var name = cal.getName();
    if (cal.getId() === rosterId || IMP_SKIP_CALENDARS.indexOf(name) >= 0 ||
        IMP_INFO_CALENDARS.indexOf(name) >= 0) return;
    cals.push(cal);
    impEvents_(cal, from, to).forEach(function (e) {      // skips our own crew entries
      var p = impParseTitle_(e.title, ctx);
      if (p.kind !== 'visit') return;
      var day = Utilities.formatDate(e.date, tz, 'yyyy-MM-dd');
      [p.site, p.second, p.siteText,
       e.title.replace(/^\s*\([^)]*\)\s*/, '').replace(/\s+-\s*[^-]*$/, '')]   // raw wording
        .filter(Boolean).forEach(function (s) {
          var k = day + '|' + calSiteKey_(s);
          if (!map[k]) map[k] = cal;
        });
      // supervisor + day: catches a site renamed by hand in Master Roster
      if (p.sup) {
        var ks = 'sup:' + day + '|' + p.sup;
        map[ks] = (map[ks] === undefined) ? cal : null;     // null = ambiguous (two visits)
      }
    });
  });
  return { map: map, cals: cals };
}

/**
 * Master Roster -> the crew entries we want, keyed so a re-sync can match them.
 * Management's council calendars already carry "(hours) Site - Supervisor", so we
 * add only the crew list, IN THE SAME council calendar (same colour, sits beside
 * the site entry):   Nina Matt Riley Zac   <- supervisor, ATM, crew by seniority
 * A visit with no matching council entry goes to the "Apunga Roster" calendar.
 */
function desiredEvents_() {
  var vis = readVisits_();                        // shared with the renderer
  var regions = siteRegionMap_();
  var seen = {}, out = [], rank = staffRank_();
  var dates = vis.map(function (v) { return v.date; }).sort();
  var homes = dates.length ? councilHomes_(dates[0], dates[dates.length - 1]) : { map: {}, cals: [] };
  vis.forEach(function (v) {
    var base = v.date + '|' + v.site;
    seen[base] = (seen[base] || 0) + 1;
    var key = seen[base] > 1 ? base + '#' + seen[base] : base;

    var team = [v.sup, v.atm].concat(bySeniority_(v.crew, rank)).filter(Boolean).join(' ') +
               (v.short ? ' + ' + v.short + ' SHORT' : '');

    var lines = [];
    if (v.notes) lines.push('Note: ' + v.notes);
    if (v.sup) lines.push('Supervisor: ' + v.sup);
    if (v.atm) lines.push('ATM: ' + v.atm);
    if (v.crew.length) lines.push('Crew: ' + v.crew.join(', '));
    if (v.short) lines.push('SHORT ' + v.short + ' crew');
    lines.push('', 'Site: ' + v.site + (v.siteHours ? ' (' + v.siteHours + 'h)' : ''),
               'Crew list generated from the Master Roster. Edit the roster and re-sync, not this entry.');

    var home = homes.map[v.date + '|' + calSiteKey_(v.site)] ||
               (v.sup && homes.map['sup:' + v.date + '|' + v.sup]) || null;
    out.push({ key: key + '|team', date: v.date, title: team || '(no crew)', body: lines.join('\n'),
               cal: home,                          // null -> Apunga Roster calendar
               colour: home ? null : (CAL_COLOURS[regions[v.site] || ''] || null) });
  });
  out.cals = homes.cals;
  return out;
}

function calDate_(iso) { return new Date(iso + 'T00:00:00'); }

/** Every event this script tagged, across the given calendars and window: key -> {ev, calId}. */
function taggedEvents_(cals, from, to) {
  var mine = {}, dupes = [];
  cals.forEach(function (cal) {
    cal.getEvents(from, to).forEach(function (ev) {
      var k = ev.getTag(CAL_TAG);                 // untagged events are never touched
      if (!k) return;
      if (mine[k]) dupes.push(ev); else mine[k] = { ev: ev, calId: cal.getId() };
    });
  });
  return { mine: mine, dupes: dupes };
}

/**
 * Sync. Safe to run repeatedly: it updates rather than duplicates, and only ever
 * edits or deletes entries it created itself (tagged). Management's own entries
 * are never changed.
 */
function syncRosterToCalendar() {
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('A sync is already running. Try again shortly.');

  try {
    var want = desiredEvents_();
    if (!want.length) throw new Error('No roster rows to sync. Run the optimiser first.');

    var roster = rosterCalendar_();
    var dates = want.map(function (e) { return e.date; }).sort();
    var from = calDate_(dates[0]);
    var to = calDate_(dates[dates.length - 1]);
    to.setDate(to.getDate() + 1);

    var t = taggedEvents_(want.cals.concat([roster]), from, to), mine = t.mine;
    var created = 0, updated = 0, removed = 0, unchanged = 0, inCouncil = 0, failed = [];
    t.dupes.forEach(function (ev) { ev.deleteEvent(); removed++; });

    want.forEach(function (e) {
      var target = e.cal || roster, have = mine[e.key];
      if (e.cal) inCouncil++;
      if (have && have.calId !== target.getId()) {   // visit moved calendar: recreate
        have.ev.deleteEvent(); removed++; have = null;
      }
      delete mine[e.key];
      if (have) {
        var ev = have.ev, changed = false;
        if (ev.getTitle() !== e.title) { ev.setTitle(e.title); changed = true; }
        if (ev.getDescription() !== e.body) { ev.setDescription(e.body); changed = true; }
        if (e.colour && ev.getColor() !== String(e.colour)) {
          try { ev.setColor(e.colour); changed = true; } catch (err) {}
        }
        if (changed) updated++; else unchanged++;
      } else {
        try {
          var nev = target.createAllDayEvent(e.title, calDate_(e.date), { description: e.body });
          nev.setTag(CAL_TAG, e.key);
          if (e.colour) { try { nev.setColor(e.colour); } catch (err) {} }
          created++;
        } catch (err) {                            // no write access to that council calendar
          failed.push(target.getName());
          var fb = roster.createAllDayEvent(e.title, calDate_(e.date), { description: e.body });
          fb.setTag(CAL_TAG, e.key); created++;
        }
      }
    });

    // anything left in `mine` is a visit that no longer exists
    Object.keys(mine).forEach(function (k) { mine[k].ev.deleteEvent(); removed++; });

    var msg = created + ' added, ' + updated + ' updated, ' + removed + ' removed, ' +
              unchanged + ' unchanged (' + dates[0] + ' to ' + dates[dates.length - 1] + '). ' +
              inCouncil + ' of ' + want.length + ' crews are in council calendars; the rest are in "' +
              CAL_NAME + '" (no matching council entry).';
    if (failed.length) {
      var uniq = failed.filter(function (n, i) { return failed.indexOf(n) === i; });
      msg += ' Could not write to: ' + uniq.join(', ') + ' (put in "' + CAL_NAME + '" instead).';
    }
    Logger.log('Calendar sync: ' + msg);
    ss.toast(msg, 'Calendar sync', 20);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

/** Removes every crew entry this script created, in every calendar. Manual entries are left alone. */
function clearRosterCalendar() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  if (ui) {
    var a = ui.alert('Clear roster events?',
      'Deletes only the crew entries created by this sync, in the council calendars and in "' +
      CAL_NAME + '". Anything added by hand is left alone.', ui.ButtonSet.YES_NO);
    if (a !== ui.Button.YES) return;
  }
  var from = new Date(); from.setFullYear(from.getFullYear() - 1);
  var to = new Date();   to.setFullYear(to.getFullYear() + 1);
  var rosterId = rosterCalendar_().getId();
  var cals = CalendarApp.getAllCalendars().filter(function (c) {
    return c.getId() === rosterId || IMP_SKIP_CALENDARS.indexOf(c.getName()) < 0;
  });
  var n = 0;
  cals.forEach(function (cal) {
    cal.getEvents(from, to).forEach(function (ev) {
      if (ev.getTag(CAL_TAG)) { ev.deleteEvent(); n++; }
    });
  });
  SpreadsheetApp.getActive().toast(n + ' roster events removed.', 'Calendar sync', 10);
}

/** Read-only check: says which calendars will be used and what is there. */
function checkCalendarSync() {
  var out = [];
  try {
    var cal = rosterCalendar_();
    out.push('Fallback calendar: "' + cal.getName() + '"');
  } catch (e) {
    out.push('Calendar: ERROR ' + e.message);
  }
  try {
    var want = desiredEvents_();
    var inC = want.filter(function (e) { return e.cal; });
    out.push('Crew entries ready to sync: ' + want.length + ' (' + inC.length + ' into council calendars)');
    var miss = want.filter(function (e) { return !e.cal; }).slice(0, 8)
      .map(function (e) { return e.date + ' ' + e.key.split('|')[1]; });
    if (miss.length) out.push('No matching council entry (go to "' + CAL_NAME + '"): ' + miss.join('; '));
  } catch (e) { out.push('Master Roster: ERROR ' + e.message); }
  var msg = out.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Calendar sync', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) {}
  return msg;
}
