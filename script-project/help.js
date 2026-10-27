/**
 * Help.gs — the how-to guide, shown inside the spreadsheet.
 *
 * Apunga Roster ▸ How to use this tool. Rendering it as HTML rather than a
 * worksheet keeps it readable (headings, tables, colour) while still living
 * with the tool, so it can't drift off into a forgotten document.
 *
 * To edit the wording, change the HTML in helpHtml_() below.
 */

function showHelp() {
  var html = HtmlService.createHtmlOutput(helpHtml_())
    .setWidth(940)
    .setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, 'Apunga Roster — how to use this tool');
}

function helpHtml_() {
  return [
'<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>',
'  body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;font-size:14px;',
'       line-height:1.55;color:#202124;margin:0;padding:0 28px 40px}',
'  h1{font-size:21px;margin:24px 0 4px}',
'  h2{font-size:17px;margin:30px 0 8px;padding-bottom:6px;border-bottom:2px solid #eceff1}',
'  h3{font-size:14px;margin:18px 0 4px;color:#3c4043}',
'  p,li{margin:6px 0}  ul,ol{padding-left:22px}',
'  .lead{color:#5f6368;margin:2px 0 10px}',
'  table{border-collapse:collapse;width:100%;margin:10px 0;font-size:13px}',
'  th{background:#f1f3f4;text-align:left;padding:7px 9px;border:1px solid #dadce0;font-weight:600}',
'  td{padding:7px 9px;border:1px solid #dadce0;vertical-align:top}',
'  code{background:#f1f3f4;padding:1px 5px;border-radius:3px;font-size:12.5px}',
'  .step{background:#f8f9fa;border-left:4px solid #1a73e8;padding:10px 14px;margin:10px 0;border-radius:0 4px 4px 0}',
'  .warn{background:#fef7e0;border-left:4px solid #f6bf26;padding:10px 14px;margin:12px 0;border-radius:0 4px 4px 0}',
'  .danger{background:#fce8e6;border-left:4px solid #d50000;padding:10px 14px;margin:12px 0;border-radius:0 4px 4px 0}',
'  .ok{background:#e6f4ea;border-left:4px solid #0b8043;padding:10px 14px;margin:12px 0;border-radius:0 4px 4px 0}',
'  .pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;color:#fff}',
'  .menu{background:#202124;color:#fff;padding:1px 7px;border-radius:3px;font-size:12.5px;white-space:nowrap}',
'  .tabname{font-weight:600;color:#1a73e8}',
'</style></head><body>',

'<h1>How to use the roster tool</h1>',
'<p class="lead">Everything happens from the <span class="menu">Apunga Roster</span> menu at the top of this spreadsheet.</p>',

'<h2>The weekly routine</h2>',
'<div class="step"><b>1. Fill in the Demand tab.</b> Quickest: block the month out in the council ',
'calendars as usual, then <span class="menu">Apunga Roster ▸ Import Demand from Google Calendar…</span>. ',
'It writes a <b>Demand Import</b> tab to check (green OK, yellow CHECK, red FIX: pick the site from the ',
'dropdown), and <span class="menu">Apply Demand import</span> copies it into Demand. Or type rows by hand: ',
'one row per site visit: date, day, site, ',
'supervisor, and either site hours or crew size. Add anything special in <b>Notes</b> ',
'(e.g. "bring the auger"), and use <b>Pin staff</b> to lock specific people onto a visit.</div>',
'<div class="step"><b>2. Check the Staff tab is current.</b> Availability for the week, plus any ',
'new starters, leavers or changed notes.</div>',
'<div class="step"><b>3. Click <span class="menu">Apunga Roster ▸ Run optimiser</span>.</b> ',
'Takes about ten seconds. You will see progress messages in the bottom-right corner.</div>',
'<div class="step"><b>4. Read the <span class="tabname">Run Log</span> tab.</b> Anything in red ',
'needs a decision from you. See below.</div>',
'<div class="step"><b>5. Adjust the roster</b> in <span class="tabname">Master Roster</span> if ',
'needed, then click <span class="menu">Render both</span> to update the calendar and staff view.</div>',

'<h2>What each tab is for</h2>',
'<table><tr><th style="width:170px">Tab</th><th style="width:110px">You edit it?</th><th>What it is</th></tr>',
'<tr><td class="tabname">Demand</td><td><b>Yes</b></td><td>What work needs doing. This is your main input.</td></tr>',
'<tr><td class="tabname">Staff</td><td><b>Yes</b></td><td>Who is available, their quals, exclusions, minimum days.</td></tr>',
'<tr><td class="tabname">SITES</td><td>Occasionally</td><td>Site list with council, region, coordinates, tick flag.</td></tr>',
'<tr><td class="tabname">TravelMatrix</td><td>No</td><td>Travel times. Rebuilt only when someone moves or a site is added.</td></tr>',
'<tr><td class="tabname">Master Roster</td><td><b>Yes</b></td><td>The draft roster — <b>the one you edit</b> to make changes.</td></tr>',
'<tr><td class="tabname">Master Calendar</td><td>No</td><td>Calendar view for reading. Regenerated from Master Roster.</td></tr>',
'<tr><td class="tabname">Staff View</td><td>No</td><td>Two-week per-person view with total hours.</td></tr>',
'<tr><td class="tabname">Run Log</td><td>No</td><td>What the optimiser wants you to know.</td></tr>',
'<tr><td class="tabname">Roster Changes</td><td>No</td><td>What changed compared with the previous version.</td></tr>',
'</table>',
'<p>The read-only tabs will warn you if you type in them — they get rebuilt, so edits there are lost.</p>',

'<h2>Making changes to the roster</h2>',
'<p>Edit the <span class="tabname">Master Roster</span> tab directly — swap a name, clear a row for ',
'a cancellation, change a supervisor. Then <span class="menu">Apunga Roster ▸ Render both</span>.</p>',
'<ul>',
'<li>Need a sixth crew member? Add a column headed <code>Staff6</code> and type the name in.</li>',
'<li>Site names must match the SITES tab exactly, or the block shows grey on the calendar.</li>',
'<li>Cancelling a visit: delete the whole row.</li>',
'</ul>',
'<div class="danger"><b>Careful:</b> <span class="menu">Run optimiser</span> rebuilds the whole ',
'Master Roster, so hand edits are replaced. They are never lost — the previous version is ',
'snapshotted automatically and every difference is listed in <span class="tabname">Roster Changes</span>. ',
'To bring it back, use <span class="menu">Restore last snapshot</span>.</div>',
'<div class="warn"><b>Better approach:</b> if you want a change to survive future runs, make it in the ',
'<span class="tabname">Demand</span> tab instead — use <b>Pin staff</b> to force someone onto a visit, ',
'<b>Keep off</b> to exclude them, or <b>Lock crew</b> to freeze the whole crew. Then re-run.</div>',

'<h2>Reading the Run Log</h2>',
'<table><tr><th style="width:150px">Label</th><th>What to do</th></tr>',
'<tr><td><span class="pill" style="background:#d50000">ACTION NEEDED</span></td>',
'<td>Something is wrong or unfilled. Read the "What to do" column and fix it, then re-run.</td></tr>',
'<tr><td><span class="pill" style="background:#f6bf26;color:#000">CHECK</span></td>',
'<td>Probably fine, but worth a glance — e.g. a visit with no ATM, or a day with no spare staff.</td></tr>',
'<tr><td><span class="pill" style="background:#9aa0a6">INFO</span></td>',
'<td>Background detail: spare staff per day, carpool status, which sites an exclusion blocked.</td></tr>',
'<tr><td><span class="pill" style="background:#0b8043">OK</span></td>',
'<td>A check that passed. Nothing to do.</td></tr>',
'<tr><td><span class="pill" style="background:#3d85c6">SUMMARY</span></td>',
'<td>Headline numbers for the run.</td></tr></table>',
'<h3>Two worth understanding</h3>',
'<ul>',
'<li><b>Crew shortfall</b> — a visit could not be filled. Usually nobody eligible was free. ',
'Options: accept a smaller crew, move the visit, or free someone up.</li>',
'<li><b>Spare staff</b> — who was available but unrostered each day. This is your list when ',
'you need another body somewhere.</li>',
'</ul>',

'<h2>Pinning people onto a visit</h2>',
'<p>Type a name (or several, comma separated) into <b>Pin staff</b> on the Demand row and they ',
'will be on that visit. A pin <b>overrides the availability grid</b> — useful for ad-hoc staff, ',
'one-off arrangements and someone picking up an extra day.</p>',
'<p>What a pin will <i>not</i> override: required qualifications, tick sites for tick-sensitive ',
'staff, excluded sites and regions, and their maximum travel time. If a pin fails, the ',
'<span class="tabname">Run Log</span> says exactly which of those was the reason.</p>',
'<p>Pins that ignore availability are listed in the Run Log so nothing happens silently, and a ',
'pin involving a very long trip is flagged too — pinning bypasses travel balancing.</p>',

'<h2>Weekend work</h2>',
'<p>The Staff tab has no Saturday or Sunday availability, so nobody is rostered onto weekends ',
'automatically. Put the people who have agreed to work into <b>Pin staff</b> on that Demand row ',
'and they will appear. The supervisor comes across as normal.</p>',

'<h2>Undo and history</h2>',
'<table><tr><th style="width:250px">Menu item</th><th>What it does</th></tr>',
'<tr><td><span class="menu">Snapshot roster now</span></td><td>Saves the current roster before you start hand-editing.</td></tr>',
'<tr><td><span class="menu">Show changes since snapshot</span></td><td>Lists what has changed since the last snapshot, without running anything.</td></tr>',
'<tr><td><span class="menu">Restore last snapshot</span></td><td>Puts the previous roster back. Asks first, and is itself reversible.</td></tr>',
'</table>',
'<p>The six most recent snapshots are kept as hidden tabs (View ▸ Hidden sheets).</p>',

'<h2>When something goes wrong</h2>',
'<table><tr><th style="width:290px">Message</th><th>What it means</th></tr>',
'<tr><td>Someone else is running the optimiser</td><td>Two people clicked at once. Wait a minute and try again.</td></tr>',
'<tr><td>The solver took too long to answer</td><td>Usually the first run of the day. Try again.</td></tr>',
'<tr><td>Could not build a roster</td><td>A data problem — the <span class="tabname">Run Log</span> says which row.</td></tr>',
'<tr><td>Not allowed to call the solver / rejected the secret</td><td>A settings problem. Run <span class="menu">Check optimiser connection</span> and send the result to whoever maintains the tool.</td></tr>',
'<tr><td>Menu is missing</td><td>Reload the spreadsheet.</td></tr>',
'</table>',
'<div class="ok"><b>Safe to try:</b> nothing you click in this menu can damage the source data. ',
'The Staff, Sites and TravelMatrix tabs are only ever read, never written to. Demand only changes when ',
'you click <b>Apply Demand import</b>, and a hidden backup copy is made first.</div>',

'<h2>What the optimiser is actually doing</h2>',
'<p>It builds the roster that adds up to the least total travel time, while obeying the rules it ',
'cannot break: supervisors stay on their assigned site, people only go to sites they are ',
'qualified and permitted to attend, nobody is double-booked, tick-sensitive staff avoid tick ',
'sites, crews stay within safety rules, and pinned staff stay pinned.</p>',
'<p>It also balances travel fairly over each fortnight, spreads paid ATM shifts evenly, respects ',
'carpool pairings, honours minimum guaranteed days, and leans towards people\'s preferred sites ',
'and regions where it can do so cheaply.</p>',
'<p class="lead">It produces a <b>draft</b>. Your judgement still wins, but let me know if something is consistetly wrong! We can always make this tool better :)</p>', 

'</body></html>'].join('\n');
}