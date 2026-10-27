/**
 * importDemand.gs — fill the Demand tab from management's Google Calendars.
 *
 * Management block out each month as all-day events on the council calendars,
 * titled like:
 *     (38) Bungaribee - Rosie
 *     (23) Pyes Creek (7.5) Arundo Pyes - Sonya        <- second site
 *     (38+5) Wisemans - Sonya                          <- extra hours
 *     (30.5) Clifftops Bondi (7 Rps) - Lee             <- note in brackets
 *     (30.5) Riverview [WWC] - Kim                     <- tag
 *     (38) Sydney Park Wetlands - ?                    <- no supervisor yet
 *
 * Two menu items, so nothing reaches Demand without a look first:
 *   1. "Import Demand from Google Calendar…"  reads every calendar this account
 *      can see for the chosen month and writes a "Demand Import" review tab:
 *      one row per visit, Status OK / CHECK / FIX, with the reason.
 *   2. "Apply Demand import"  backs up the current Demand tab (hidden copy),
 *      then replaces its rows with the reviewed import.
 *
 * Fix a FIX row by choosing the Site from the dropdown in the review tab. When
 * applied, that calendar wording is remembered in the "Calendar aliases" tab,
 * so next month it matches automatically.
 *
 * Reads calendars only. Never writes to any calendar. Skips the "Apunga Roster"
 * calendar (the optimiser's own output) so the roster can't feed back into Demand.
 */

var IMP_TAB = 'Demand Import';
var IMP_ALIAS_TAB = 'Calendar aliases';
var IMP_DEMAND_TAB = 'Demand';
var IMP_SITES_TAB = 'Sites';
var IMP_STAFF_TAB = 'Staff';
var IMP_SITE_H = 'Roster site name (Official name)';
// calendars that are never site visits
var IMP_SKIP_CALENDARS = ['Apunga Roster', 'Holidays in Australia'];
// calendars read only to warn about clashes (office days, leave, training)
var IMP_INFO_CALENDARS = ['x OFFICE', 'Ze Holiday Calendar'];
var IMP_WEEKDAY = ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat'];
var IMP_SUP_HOURS = 8, IMP_CREW_HOURS = 7.5;

// calendar shorthand -> Sites wording (extend as new shorthand appears)
var IMP_SHORTHAND = { 'nth': 'north', 'sth': 'south', 'n.head': 'north head',
                      'bicent': 'bicentennial', 'vaulcluse': 'vaucluse' };
// shorthand that could mean two sites -> the one management usually mean
// (first name that exists in Sites wins, so it survives the Oct 2026 rename)
var IMP_DEFAULT_SITE = { 'quarry': ['Hornsby Quarry', 'Quarry (Hornsby Quarry)'] };

var IMP_COLS = ['Status', 'Date', 'Day', 'Site', 'Secondary site', 'Supervisor', 'Site Hours',
                'Crew size', 'Required quals', 'Pin staff', 'Notes', 'Calendar', 'Calendar entry',
                'Why', 'Calendar site text'];

// =============================================================================
// Parsing (pure functions: no Sheets/Calendar calls, so they can be tested)
// =============================================================================

function impNorm_(s) {
  return String(s == null ? '' : s).replace(/[ \s]+/g, ' ').replace(/[’']/g, '')
    .trim().toLowerCase();
}
function impShort_(name) {                        // 'Primrose (Primrose Park)' -> 'primrose'
  var n = impNorm_(name), i = n.indexOf(' (');
  return (i > 0 ? n.slice(0, i) : n).trim();
}
function impExpand_(t) {
  return t.split(' ').map(function (w) { return IMP_SHORTHAND.hasOwnProperty(w) ? IMP_SHORTHAND[w] : w; })
    .filter(Boolean).join(' ');
}
function impRatio_(a, b) {                        // 1 - edit distance / longer length
  if (a === b) return 1;
  var m = a.length, n = b.length; if (!m || !n) return 0;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= n; j++) prev[j] = j;
  for (i = 1; i <= m; i++) {
    cur = [i];
    for (j = 1; j <= n; j++)
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return 1 - prev[n] / Math.max(m, n);
}

/** Builds the lookup context once per import. */
function impContext_(siteNames, staffRows, aliases) {
  var sites = siteNames.filter(function (n) { return impNorm_(n); });
  var info = sites.map(function (n) {
    var sh = impShort_(n);
    return { name: n, full: impNorm_(n), short: sh,
             segs: sh.split('/').map(function (x) { return x.trim(); }).filter(Boolean) };
  });
  var staff = {};
  staffRows.forEach(function (r) { if (r.id) staff[r.id.toLowerCase()] = r; });
  var al = {};
  Object.keys(aliases || {}).forEach(function (k) { al[impNorm_(k)] = aliases[k]; });
  var byNorm = info.reduce(function (o, s) { o[s.full] = s.name; return o; }, {});
  // alias targets typed into a sheet lose trailing spaces / line breaks: resolve loosely
  Object.keys(al).forEach(function (k) { al[k] = byNorm[impNorm_(al[k])] || al[k]; });
  return { sites: info, staff: staff, aliases: al,
           byName: info.reduce(function (o, s) { o[s.name] = s; return o; }, {}) };
}

/** Tight matches only: alias, exact name, exact short name, exact '/'-segment. */
function impMatchExact_(t, ctx) {
  if (ctx.aliases[t] && ctx.byName[ctx.aliases[t]]) return { site: ctx.aliases[t], note: '' };
  var dflt = (IMP_DEFAULT_SITE[t] || []).filter(function (n) { return ctx.byName[n]; })[0];
  if (dflt) return { site: dflt, note: "'" + t + "' taken as " + dflt, quiet: true };
  var tiers = [
    function (s) { return s.full === t; },
    function (s) { return s.short === t; },
    function (s) { return s.segs.indexOf(t) >= 0; }
  ];
  for (var k = 0; k < tiers.length; k++) {
    var hit = ctx.sites.filter(tiers[k]);
    if (hit.length === 1) return { site: hit[0].name, note: '' };
    if (hit.length > 1) return { site: '', note: 'could be ' + hit.slice(0, 3).map(function (s) { return s.name; }).join(' | ') };
  }
  return null;
}

/** Calendar site wording -> {site, note, loose}. site '' means not matched. */
function impMatchSite_(text, ctx) {
  var t = impExpand_(impNorm_(text).replace(/\?+$/, '').trim());
  if (!t) return { site: '', note: 'no site in the calendar entry' };
  var r = impMatchExact_(t, ctx);
  if (r) return r;
  function one(list, how) {
    if (list.length === 1) return { site: list[0].name, note: how ? "matched '" + text.trim() + "' to " + list[0].name : '', loose: !!how };
    if (list.length > 1) return { site: '', note: 'could be ' + list.slice(0, 3).map(function (s) { return s.name; }).join(' | ') };
    return null;
  }
  // prefix either way ('Byles Creek' -> 'Byles', 'Girrahween' -> 'Girrahween Park')
  r = one(ctx.sites.filter(function (s) {
    return s.short.indexOf(t) === 0 || s.full.indexOf(t) === 0 || (t.indexOf(s.short + ' ') === 0);
  }), 'prefix');
  if (r) { r.loose = false; return r; }          // unambiguous: noted, not flagged
  // every word appears in exactly one site name ('Cooks Foreshore', 'Sutherland/Cobran')
  var words = t.split(/[\s\/&+]+/).filter(function (w) { return w.length >= 3; });
  if (words.length) {
    r = one(ctx.sites.filter(function (s) {
      return words.every(function (w) { return s.full.indexOf(w) >= 0; });
    }), 'words');
    if (r) return r;
  }
  // close spelling ('Bungaribee' -> 'Bungarribee', 'Galstone Gorge' -> 'Galston Gorge')
  var scored = ctx.sites.map(function (s) { return [Math.max(impRatio_(t, s.short), impRatio_(t, s.full)), s]; })
    .sort(function (a, b) { return b[0] - a[0]; });
  if (scored.length && scored[0][0] >= 0.85 && (scored.length < 2 || scored[1][0] < scored[0][0] - 0.05))
    return { site: scored[0][1].name, note: "spelling matched '" + text.trim() + "' to " + scored[0][1].name, loose: true };
  // plural ('Fishers' -> 'Fisher')
  if (/s$/.test(t) && t.length > 4) {
    var p = impMatchSite_(t.slice(0, -1), ctx);
    if (p.site) return { site: p.site, note: "matched '" + text.trim() + "' to " + p.site, loose: true };
  }
  var close = scored.filter(function (x) { return x[0] >= 0.72; }).slice(0, 2).map(function (x) { return x[1].name; });
  return { site: '', note: close.length ? 'not in Sites. Did you mean ' + close.join(' or ') + '?' : 'not in Sites' };
}

function impParseHours_(h) {
  var t = String(h).replace(/\s/g, ''), unsure = t.indexOf('?') >= 0;
  t = t.replace(/\?/g, '');
  if (!/^\d+(\.\d+)?(\+\d+(\.\d+)?)*$/.test(t)) return null;
  return { hours: t.split('+').reduce(function (a, x) { return a + parseFloat(x); }, 0),
           unsure: unsure, raw: h };
}

/** One calendar title -> {kind:'visit'|'skip', ...}. */
function impParseTitle_(title, ctx) {
  var raw = String(title || '').replace(/ /g, ' ').trim();
  var m = raw.match(/^\(([^)]*)\)\s*(.+)$/);
  if (!m || !impParseHours_(m[1])) return { kind: 'skip', why: 'no (hours) at the start, so not a site visit' };
  var status = 'OK', why = [], notes = [];
  function check(msg) { if (status === 'OK') status = 'CHECK'; why.push(msg); }
  function fix(msg) { status = 'FIX'; why.push(msg); }

  // split "site part - people part" on the LAST dash that has a space on either side
  var body = '(' + m[1] + ') ' + m[2], cut = -1, re = /\s+-\s*|\s*-\s+/g, mm;
  while ((mm = re.exec(body)) !== null) cut = mm.index;
  var sitePart = cut >= 0 ? body.slice(0, cut) : body;
  var people = cut >= 0 ? body.slice(cut).replace(/^\s*-\s*/, '').trim() : '';
  if (cut < 0) check('no " - Supervisor" part in the entry');

  // tags [WWC] [Water?] [Final?]
  var quals = [];
  sitePart = sitePart.replace(/\[([^\]]*)\]/g, function (_, tag) {
    if (/^\s*wwc\s*$/i.test(tag)) { quals.push('wwc'); notes.push('WWC'); }
    else notes.push(tag.trim());
    return ' ';
  });

  // segments: (hours) site text, (hours) site text, ...  Non-number brackets are notes.
  var segs = [], segRe = /\(([^)]*)\)/g, last = null, pos = 0, g;
  while ((g = segRe.exec(sitePart)) !== null) {
    var hrs = impParseHours_(g[1]);
    if (last) last.text += sitePart.slice(pos, g.index);
    if (hrs) { last = { hours: hrs, text: '' }; segs.push(last); }
    else {
      var n = g[1].trim();
      notes.push(/rps/i.test(n) ? 'ropes ' + n.replace(/\s*rps\s*/i, 'h ').trim() : n);
      if (/\?/.test(n) && /rps/i.test(n)) check('rope hours marked uncertain (' + n + ')');
    }
    pos = segRe.lastIndex;
  }
  if (last) last.text += sitePart.slice(pos);
  segs.forEach(function (s) { s.text = s.text.replace(/\s+/g, ' ').trim(); });

  var total = segs.reduce(function (a, s) { return a + s.hours.hours; }, 0);
  if (segs.some(function (s) { return s.hours.unsure; })) check('hours marked uncertain in the calendar');

  // resolve sites; "A & B" is two sites done together unless Sites has the joined name
  var resolved = [];
  segs.forEach(function (s, i) {
    var txt = s.text, r0 = impMatchExact_(impExpand_(impNorm_(txt).replace(/\?+$/, '')), ctx);
    if (!r0 && txt.indexOf('&') > 0) {
      txt.split('&').forEach(function (p, j) {
        var r = impMatchSite_(p, ctx);
        resolved.push({ text: p.trim(), r: r, hours: j === 0 ? s.hours.hours : 0, seg: i });
      });
    } else {
      resolved.push({ text: txt, r: r0 || impMatchSite_(txt, ctx), hours: s.hours.hours, seg: i });
    }
  });
  var bare = sitePart.replace(/\([^)]*\)/g, '').trim();
  if (/\?/.test(bare) && !/^\?+$/.test(bare)) check('site marked uncertain (?)');

  function noted(r) { if (!r.note) return; if (r.loose || /^could be/.test(r.note)) check(r.note); else why.push(r.note); }
  var primary = resolved[0] || { text: '', r: { site: '', note: 'no site in the calendar entry' } };
  if (!primary.r.site) fix(/^\?*$/.test(primary.text) ? "site not decided yet ('" + (primary.text || '?') + "')"
                                                      : "site '" + primary.text + "': " + primary.r.note);
  else noted(primary.r);
  var site = primary.r.site, second = '', extras = [];
  resolved.slice(1).forEach(function (x) {
    if (!x.r.site) { check("second site '" + x.text + "': " + x.r.note); notes.push('also: ' + x.text); return; }
    noted(x.r);
    if (x.r.site === site || x.r.site === second) { notes.push(x.text + ' (same site)'); return; }
    if (!second) second = x.r.site; else { extras.push(x.r.site); }
  });
  if (extras.length) { notes.push('also: ' + extras.join(', ')); check('more than two sites in one entry: ' + extras.join(', ') + ' (only one Secondary site column)'); }
  if (segs.length > 1)
    notes.push(segs.map(function (s) { return s.hours.raw + 'h ' + s.text; }).join(' + '));

  // people: first name is the supervisor; anyone else joined with & , / is pinned
  var sup = '', pins = [], ppl = people.replace(/\?/g, function () { return ' ? '; });
  var toks = ppl.split(/\s*(?:&|,|\/|\band\b)\s*|\s+/i).filter(function (x) { return x && x !== '?'; });
  if (!toks.length) check('no supervisor yet');
  else {
    var first = ctx.staff[toks[0].toLowerCase()];
    if (!first) check("supervisor '" + toks[0] + "' is not a StaffID, left blank");
    else {
      sup = first.id;
      if (/\?/.test(people)) check(sup + ' marked uncertain (?)');
      if (first.category && first.category !== 'Supervisor') check(sup + ' is listed as ' + first.category + ', not Supervisor');
    }
    toks.slice(1).forEach(function (tk) {
      var p = ctx.staff[tk.toLowerCase()];
      if (p) pins.push(p.id); else check("'" + tk + "' is not a StaffID");
    });
    if (pins.length) notes.push('also leading: ' + pins.join(', '));
  }

  var crewExact = (total - IMP_SUP_HOURS) / IMP_CREW_HOURS + 1;
  var crew = Math.max(2, Math.ceil(crewExact - 1e-9));
  if (Math.abs(crewExact - Math.round(crewExact)) > 0.1)
    check(total + ' hours is not a whole crew (' + crewExact.toFixed(1) + ' people), rounded up to ' + crew);

  return { kind: 'visit', status: status, why: why, site: site, second: second, sup: sup,
           hours: total, crew: crew, quals: quals, pins: pins, notes: notes,
           siteText: primary.text };
}

/** Clash notes for a supervisor from office/leave/training entries that day. */
function impClashes_(sup, infoTitles) {
  if (!sup) return [];
  var re = new RegExp('(^|[^A-Za-z])' + sup.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '($|[^A-Za-z])', 'i');
  return (infoTitles || []).filter(function (t) {
    return re.test(t) && !/birthday|\bback\b/i.test(t);
  });
}

// =============================================================================
// Sheets / Calendar glue
// =============================================================================

function impReadSites_(ss) {
  var sh = ss.getSheetByName(IMP_SITES_TAB);
  if (!sh) throw new Error('No "' + IMP_SITES_TAB + '" tab.');
  var v = sh.getDataRange().getValues(), c = v[0].map(String).indexOf(IMP_SITE_H);
  if (c < 0) throw new Error('Sites tab has no "' + IMP_SITE_H + '" column.');
  return v.slice(1).map(function (r) { return String(r[c]); }).filter(function (n) { return n.trim(); });
}

function impReadStaff_(ss) {
  var v = ss.getSheetByName(IMP_STAFF_TAB).getDataRange().getValues(), h = -1;
  for (var i = 0; i < 15 && h < 0; i++) if (v[i].map(String).indexOf('StaffID') >= 0) h = i;
  if (h < 0) throw new Error('Staff tab: no StaffID header found.');
  var hd = v[h].map(function (x) { return String(x).trim(); });
  var ci = hd.indexOf('StaffID'), cc = hd.indexOf('Category');
  return v.slice(h + 1).map(function (r) {
    return { id: String(r[ci]).trim(), category: cc >= 0 ? String(r[cc]).trim() : '' };
  }).filter(function (r) { return r.id && !/^(Wk|TOTAL)/.test(r.id); });
}

function impReadAliases_(ss) {
  var sh = ss.getSheetByName(IMP_ALIAS_TAB), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (String(r[0]).trim() && String(r[1]).trim()) out[String(r[0])] = String(r[1]);
  });
  return out;
}

function impYmd_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
// Dates go into the sheet as 'yyyy-mm-dd' text, which Sheets reads as that calendar
// day in its OWN time zone. Writing a JS Date instead shifts it a day, because the
// spreadsheet is on GMT and the script on Sydney time.
function impSheetDate_(d, tz) {
  return Object.prototype.toString.call(d) === '[object Date]'
    ? Utilities.formatDate(d, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd') : d;
}

/** Every all-day (or timed) event in [from, to), as {date, title, calendar}, one per day. */
function impEvents_(cal, from, to) {
  var out = [];
  cal.getEvents(from, to).forEach(function (ev) {
    var s, e;
    if (ev.isAllDayEvent()) { s = ev.getAllDayStartDate(); e = ev.getAllDayEndDate(); }
    else { s = ev.getStartTime(); e = new Date(s.getTime() + 1); }
    var days = 0;
    for (var d = new Date(s.getFullYear(), s.getMonth(), s.getDate()); d < e; d.setDate(d.getDate() + 1)) {
      if (d < from || d >= to) continue;
      out.push({ date: new Date(d), title: ev.getTitle(), calendar: cal.getName(), multi: false });
      days++;
    }
    if (days > 1) out.slice(-days).forEach(function (o) { o.multi = true; });
  });
  return out;
}

function impAskMonth_(ui) {
  var now = new Date(), next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var dflt = Utilities.formatDate(next, 'Australia/Sydney', 'MMM yyyy');
  var r = ui.prompt('Import Demand from Google Calendar',
    'Which month? (e.g. ' + dflt + '). Leave blank for ' + dflt + '.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var t = r.getResponseText().trim() || dflt;
  var mm = t.match(/^([A-Za-z]{3})[a-z]*\s+(\d{4})$/) || null;
  var months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var mi = mm ? months.indexOf(mm[1].toLowerCase()) : -1;
  if (mi < 0) { ui.alert('Could not read "' + t + '". Type it like ' + dflt + '.'); return null; }
  return new Date(+mm[2], mi, 1);
}

function importDemandFromCalendar() {
  var ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  var from = impAskMonth_(ui); if (!from) return;
  var to = new Date(from.getFullYear(), from.getMonth() + 1, 1);
  var ctx = impContext_(impReadSites_(ss), impReadStaff_(ss), impReadAliases_(ss));

  var rosterId = PropertiesService.getScriptProperties().getProperty('ROSTER_CALENDAR_ID');
  var visits = [], info = {}, calsRead = [];
  CalendarApp.getAllCalendars().forEach(function (cal) {
    var name = cal.getName();
    if (cal.getId() === rosterId || IMP_SKIP_CALENDARS.indexOf(name) >= 0) return;
    var evs = impEvents_(cal, from, to);
    if (IMP_INFO_CALENDARS.indexOf(name) >= 0) {
      evs.forEach(function (e) { (info[impYmd_(e.date)] = info[impYmd_(e.date)] || []).push(e.title); });
      return;
    }
    if (evs.length) calsRead.push(name + ' (' + evs.length + ')');
    visits = visits.concat(evs);
  });

  var rows = [], skipped = [], supDay = {};
  visits.sort(function (a, b) { return a.date - b.date || a.calendar.localeCompare(b.calendar); });
  visits.forEach(function (e) {
    var p = impParseTitle_(e.title, ctx);
    if (p.kind === 'skip') { skipped.push(e); return; }
    var dow = e.date.getDay(), key = impYmd_(e.date);
    if (dow === 0 || dow === 6) { p.status = p.status === 'FIX' ? 'FIX' : 'CHECK'; p.why.push('weekend visit: only pinned staff are rostered'); }
    if (e.multi) { p.status = p.status === 'FIX' ? 'FIX' : 'CHECK'; p.why.push('entry spans several days, imported once per day'); }
    impClashes_(p.sup, info[key]).forEach(function (t) {
      if (p.status === 'OK') p.status = 'CHECK';
      p.why.push(p.sup + " also on '" + t + "' that day");
    });
    if (p.sup) {
      var k2 = key + '|' + p.sup;
      if (supDay[k2]) {
        if (p.status === 'OK') p.status = 'CHECK';
        p.why.push(p.sup + ' already supervises ' + (supDay[k2] || 'another site') + ' that day');
      } else supDay[k2] = p.site || p.siteText;
    }
    rows.push([p.status, impSheetDate_(e.date), IMP_WEEKDAY[dow], p.site, p.second, p.sup, p.hours, p.crew,
               p.quals.join(', '), p.pins.join(', '), p.notes.join('; '), e.calendar,
               e.title.trim(), p.why.join('; '), p.siteText]);
  });

  // ---- write the review tab ------------------------------------------------------
  var sh = ss.getSheetByName(IMP_TAB) || ss.insertSheet(IMP_TAB);
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();
  var monthLabel = Utilities.formatDate(from, 'Australia/Sydney', 'MMMM yyyy');
  sh.getRange(1, 1, 1, IMP_COLS.length).setValues([IMP_COLS]).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, IMP_COLS.length).setValues(rows);
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('ddd d mmm');
    var siteRange = ss.getSheetByName(IMP_SITES_TAB).getRange('A2:A');
    var siteCol = IMP_COLS.indexOf('Site') + 1;
    sh.getRange(2, siteCol, rows.length, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(siteRange, true).setAllowInvalid(true).build());
    var colours = { OK: '#d9ead3', CHECK: '#fff2cc', FIX: '#f4cccc' };
    sh.getRange(2, 1, rows.length, 1).setBackgrounds(rows.map(function (r) { return [colours[r[0]]]; }));
  }
  var r0 = rows.length + 3;
  sh.getRange(r0, 1).setValue('Not imported (no "(hours)" at the start, so not site visits)').setFontWeight('bold');
  if (skipped.length)
    sh.getRange(r0 + 1, 1, skipped.length, 4).setValues(skipped.map(function (e) {
      return ['SKIPPED', impSheetDate_(e.date), e.calendar, e.title.trim()];
    })).setNumberFormats(skipped.map(function () { return ['@', 'ddd d mmm', '@', '@']; }));
  sh.setColumnWidths(1, IMP_COLS.length, 110);
  sh.setColumnWidth(IMP_COLS.indexOf('Site') + 1, 220);
  sh.setColumnWidth(IMP_COLS.indexOf('Calendar entry') + 1, 280);
  sh.setColumnWidth(IMP_COLS.indexOf('Why') + 1, 420);
  sh.getRange(1, 1, Math.max(1, sh.getLastRow()), IMP_COLS.length).setWrap(false);
  ss.setActiveSheet(sh);

  var n = { OK: 0, CHECK: 0, FIX: 0 };
  rows.forEach(function (r) { n[r[0]]++; });
  PropertiesService.getDocumentProperties().setProperty('IMP_MONTH', impYmd_(from));
  ui.alert('Calendar import: ' + monthLabel,
    rows.length + ' visits read from ' + calsRead.length + ' calendars.\n\n' +
    '  ' + n.OK + ' OK\n  ' + n.CHECK + ' CHECK (imported, worth a look)\n  ' + n.FIX +
    ' FIX (no site matched: pick one from the Site dropdown)\n  ' + skipped.length +
    ' non-visit entries skipped (listed at the bottom)\n\n' +
    'Nothing has changed in Demand yet. When the review tab looks right, use ' +
    'Apunga Roster ▸ Apply Demand import.', ui.ButtonSet.OK);
}

function applyDemandImport() {
  var ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(IMP_TAB);
  if (!sh || sh.getLastRow() < 2) { ui.alert('Run "Import Demand from Google Calendar…" first.'); return; }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, IMP_COLS.length).getValues(), rows = [];
  for (var i = 0; i < vals.length && String(vals[i][0]).trim(); i++) rows.push(vals[i]);
  var C = {}; IMP_COLS.forEach(function (h, j) { C[h] = j; });
  var good = rows.filter(function (r) { return String(r[C.Site]).trim(); });
  var noSite = rows.length - good.length;
  if (!good.length) { ui.alert('No rows with a Site to apply.'); return; }

  var dem = ss.getSheetByName(IMP_DEMAND_TAB);
  var oldRows = Math.max(0, dem.getLastRow() - 1);
  var ok = ui.alert('Apply Demand import',
    'Replace the ' + oldRows + ' rows on the Demand tab with ' + good.length + ' imported visits?' +
    (noSite ? '\n\n' + noSite + ' rows still have no Site and will be left out.' : '') +
    '\n\nThe current Demand tab is backed up to a hidden tab first.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  // backup (hidden), like the roster snapshots
  var stamp = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd HH-mm');
  dem.copyTo(ss).setName('Demand backup ' + stamp).hideSheet();

  var hdr = dem.getRange(1, 1, 1, Math.max(1, dem.getLastColumn())).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  if (hdr.indexOf('Notes') < 0) {                  // optimiser carries Notes to the roster
    var nc = hdr.filter(String).length + 1;
    dem.getRange(1, nc).setValue('Notes').setFontWeight('bold');
    hdr[nc - 1] = 'Notes';
  }
  var width = hdr.length;
  if (oldRows) dem.getRange(2, 1, oldRows, width).clearContent();

  var map = { 'Date': 'Date', 'Day': 'Day', 'Site': 'Site', 'Secondary site': 'Secondary site',
              'Supervisor': 'Supervisor', 'Site Hours': 'Site Hours', 'Crew size': 'Crew size',
              'Required quals': 'Required quals', 'Pin staff': 'Pin staff', 'Notes': 'Notes' };
  var out = good.map(function (r) {
    return hdr.map(function (h) {
      if (h === 'Date') return impSheetDate_(r[C.Date], ss.getSpreadsheetTimeZone());   // read back from the sheet
      return map[h] ? r[C[map[h]]] : '';
    });
  });
  dem.getRange(2, 1, out.length, width).setValues(out);
  var dc = hdr.indexOf('Date') + 1;
  if (dc) dem.getRange(2, dc, out.length, 1).setNumberFormat('M/d/yyyy');

  // remember hand-fixed site wording for next month
  var learned = 0, al = impReadAliases_(ss), known = {};
  Object.keys(al).forEach(function (k) { known[impNorm_(k)] = true; });
  var ash = ss.getSheetByName(IMP_ALIAS_TAB);
  good.forEach(function (r) {
    var txt = String(r[C['Calendar site text']]).trim(), why = String(r[C.Why]);
    if (!txt || known[impNorm_(txt)]) return;
    if (!/^site '|could be|not in Sites/.test(why)) return;      // only rows that needed fixing
    if (!ash) {
      ash = ss.insertSheet(IMP_ALIAS_TAB);
      ash.getRange(1, 1, 1, 2).setValues([['Calendar wording', 'Site']]).setFontWeight('bold');
    }
    ash.appendRow([txt, String(r[C.Site])]);
    known[impNorm_(txt)] = true; learned++;
  });

  ss.setActiveSheet(dem);
  ui.alert('Demand updated', good.length + ' visits written to Demand (backup: "Demand backup ' + stamp + '").' +
    (learned ? '\n' + learned + ' new calendar wordings remembered in "' + IMP_ALIAS_TAB + '".' : '') +
    '\n\nNext: Run optimiser.', ui.ButtonSet.OK);
}
