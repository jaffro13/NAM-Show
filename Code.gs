/**
 * MODEL SHOW VOTING - Apps Script backend
 * ----------------------------------------
 * Bind this script to your Google Sheet (Extensions > Apps Script from inside the Sheet).
 * It expects these sheet tabs to exist, matching the template Claude gave you:
 *   - "Class Setup"    columns: Class ID | Number | Title | Max Entries | Enabled?
 *                      (only classes with Enabled? = TRUE appear on the voting page - this is how you
 *                      switch between the show, monthly Model of the Month votes, or a scaled-down
 *                      test with just a few classes turned on. Edit Max Entries any time too - both
 *                      are read live, no redeploy needed.)
 *   - "Show Settings"  B4 = Voting Open? TRUE/FALSE
 *                      B6 = Voting Closes At (date/time, drives the voter-facing countdown)
 *                      B10 = Expected Entrants, B12 = Alert Email, B14 = Votes Cast So Far (formula),
 *                      B16 = All Votes Cast? (formula), B18 = Alert Sent? (managed by this script)
 *   - "Voting Codes"   columns: Code (6-char, no dash, e.g. K4H7M2) | Used? | Time Used - for entrants
 *   - "Spectator Codes" same layout as Voting Codes, but a SEPARATE pool - for the spectator link
 *   - "Form Responses" columns: Timestamp | <Number> 1st Choice | <Number> 2nd Choice (one pair per
 *                      class, matching Class Setup) | Special Award - People's Choice
 *
 * TWO LINKS FROM ONE DEPLOYMENT:
 *   - Entrant link (normal, e.g. .../exec) -> full ballot, all enabled classes + People's Choice,
 *     checked against Voting Codes.
 *   - Spectator link (add "?mode=spectator" to the same URL, e.g. .../exec?mode=spectator) -> People's
 *     Choice only, checked against the separate Spectator Codes pool. Make a second QR code pointing
 *     at this URL.
 *
 * Because class list, entry counts, and all Show Settings are read fresh from the Sheet on every
 * request, you can edit Class Setup / Show Settings at any time and it takes effect immediately -
 * no need to touch this file or redeploy.
 */

function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode;

  if (mode === 'entry') {
    var entryTmpl = HtmlService.createTemplateFromFile('EntryForm');
    entryTmpl.classesJson = JSON.stringify(getEntryFormClasses());
    entryTmpl.entriesOpenJson = JSON.stringify(isEntriesOpen());
    return entryTmpl.evaluate()
      .setTitle('NAM Competition Entry Form 2026')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (mode === 'motm-entry') {
    var motmEntryTmpl = HtmlService.createTemplateFromFile('MotmEntryForm');
    motmEntryTmpl.entriesOpenJson = JSON.stringify(isMotmEntriesOpen());
    return motmEntryTmpl.evaluate()
      .setTitle('Model of the Month - Entry')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (mode === 'motm-vote') {
    var motmVoteTmpl = HtmlService.createTemplateFromFile('MotmVotingForm');
    motmVoteTmpl.votingDataJson = JSON.stringify(getMotmVotingData());
    motmVoteTmpl.votingOpenJson = JSON.stringify(isMotmVotingOpen());
    return motmVoteTmpl.evaluate()
      .setTitle('Model of the Month - Vote')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var isSpectator = mode === 'spectator';
  var tmpl = HtmlService.createTemplateFromFile(isSpectator ? 'SpectatorIndex' : 'Index');
  tmpl.classesJson = JSON.stringify(getClassConfig());
  tmpl.votingOpenJson = JSON.stringify(isVotingOpen());
  tmpl.votingClosesAtJson = JSON.stringify(getVotingClosesAt());
  return tmpl.evaluate()
    .setTitle(isSpectator ? "NAM People's Choice - Spectator Voting" : 'NAM Digital Voting Slip 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Adds a menu so you can trigger the completion alert manually and bulk enable/disable classes. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NAM Voting')
    .addItem('Send completion alert now', 'sendManualCompletionAlert')
    .addSeparator()
    .addItem('Enable ALL show classes', 'enableAllShowClasses')
    .addItem('Disable ALL show classes', 'disableAllShowClasses')
    .addSeparator()
    .addItem('Print entry slips...', 'showEntryPrintDialog')
    .addItem('Print Simon Breust nominees', 'printSimonBreustNominees')
    .addItem('Print Best in Show shortlist', 'printBestInShowShortlist')
    .addItem('Print class rankings (full results)', 'printClassRankings')
    .addItem('Print final placings', 'printFinalPlacings')
    .addItem('Print sponsors list', 'printSponsorsList')
    .addSeparator()
    .addItem('Close entries (sync counts)', 'closeEntries')
    .addItem('Reopen entries', 'reopenEntries')
    .addToUi();
}

/** Opens the entry-slip printing tool as a dialog inside the Sheet (kept in-Sheet rather than a
 * public link, since it lists entrant names/emails/phones - this way it's only reachable by people
 * you've already shared the Sheet with). */
function showEntryPrintDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AdminPrint').setWidth(720).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print Entry Slips');
}

/** Wraps report HTML in consistent print-friendly styling and shows it as a modal dialog with a
 * Print button - the dialog itself is the print target, no pop-up window needed. */
function buildAndShowPrintDialog_(title, bodyHtml) {
  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#222;margin:20px;}' +
    'h1{color:#2E4053;font-size:20px;margin:0 0 4px 0;}' +
    'p.sub{color:#666;font-size:13px;margin:0 0 16px 0;}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:6px;}' +
    'th{background:#2E4053;color:#fff;padding:8px;text-align:left;font-size:12px;}' +
    'td{padding:7px 8px;border-bottom:1px solid #ddd;font-size:13px;vertical-align:top;}' +
    'tr:nth-child(even) td{background:#F4F6F7;}' +
    '.class-header{background:#EAF2F8;font-weight:bold;color:#2E4053;font-size:14px;border-radius:4px;}' +
    '#printBtn{margin-bottom:16px;padding:10px 20px;font-size:14px;background:#1E8449;color:#fff;border:none;border-radius:6px;cursor:pointer;}' +
    '@media print { #printBtn{display:none;} body{margin:0;} }' +
    '</style></head><body>' +
    '<button id="printBtn" onclick="window.print()">Print this page</button>' +
    '<h1>' + title + '</h1>' +
    bodyHtml +
    '</body></html>';
  var output = HtmlService.createHtmlOutput(html).setWidth(750).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(output, title);
}

/** Prints the Top 5 Simon Breust Memorial Trophy nominees from the Class Winners sheet. */
function printSimonBreustNominees() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class Winners');
  var data = ws.getRange(41, 1, 5, 6).getValues(); // Rank, Class#, Model#, Model Name, Entrant Name, Votes
  var rows = data.filter(function (r) { return r[3]; });

  var body = '<p class="sub">Australian-subject models, ranked by People\'s Choice votes received.</p>';
  if (rows.length === 0) {
    body += '<p>No Australian-subject models have been entered yet.</p>';
  } else {
    body += '<table><tr><th>Rank</th><th>Class</th><th>Model #</th><th>Model Name</th><th>Entrant</th><th>Votes</th></tr>';
    rows.forEach(function (r) {
      body += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td>' + r[2] + '</td><td>' + r[3] + '</td><td>' + r[4] + '</td><td>' + r[5] + '</td></tr>';
    });
    body += '</table>';
  }
  buildAndShowPrintDialog_('Simon Breust Memorial Trophy - Nominees', body);
}

/** Prints every class's 1st place entry as a shortlist, for choosing Best in Show. */
function printBestInShowShortlist() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class Winners');
  var data = ws.getRange(5, 1, 24, 9).getValues(); // A-I: Class#, Class, Model#, 1st, 2nd, Points, Status, ModelName, EntrantName

  var body = '<p class="sub">Every class\'s 1st place entry, for choosing Best in Show.</p>';
  body += '<table><tr><th>Class #</th><th>Class</th><th>Model #</th><th>Model Name</th></tr>';
  data.forEach(function (r) {
    body += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td>' + r[2] + '</td><td>' + (r[7] || '') + '</td></tr>';
  });
  body += '</table>';
  buildAndShowPrintDialog_('Best in Show Shortlist', body);
}

/**
 * Prints 1st/2nd/3rd place for every class, in class order (01-24), but with each class's
 * placings listed 3rd, then 2nd, then 1st - built for reading backwards at a presentation.
 */
/** Builds a lookup map "classNum-entryNum" -> {modelName, entrantName} from Entry Models, for
 * resolving Tally's raw entry numbers into readable results across the print reports. */
function buildEntryModelsLookup_() {
  var emSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entry Models');
  var emLastRow = emSheet.getLastRow();
  var lookup = {};
  if (emLastRow >= 2) {
    var emData = emSheet.getRange(2, 1, emLastRow - 1, 12).getValues();
    emData.forEach(function (row) {
      if (!row[0]) return;
      var key = String(row[2]).trim() + '-' + row[3];
      lookup[key] = { modelName: row[4], entrantName: row[11] };
    });
  }
  return lookup;
}

/**
 * Prints every class's FULL results, 1st place through last, ordered by points - not just the
 * top 3. Shows each entry's 1st-choice and 2nd-choice vote counts alongside its points, so the
 * full picture of how a class was judged is on record, not just who placed.
 */
function printClassRankings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tally = ss.getSheetByName('Tally');
  var classMap = getClassTitleMap();
  var emLookup = buildEntryModelsLookup_();

  var body = '';
  for (var classNum = 1; classNum <= 24; classNum++) {
    var numStr = pad2(classNum);
    var startRow = 6 + (classNum - 1) * 45;
    var blockData = tally.getRange(startRow, 2, 40, 6).getValues(); // B-G: Entry#,1st,2nd,Points,Key,Rank

    var entries = [];
    blockData.forEach(function (row) {
      var entryNum = row[0];
      var key = numStr + '-' + entryNum;
      var info = emLookup[key];
      if (!info) return; // this entry slot was never actually used
      entries.push({
        rank: row[5],
        entrantName: info.entrantName,
        modelName: info.modelName,
        firstVotes: row[1],
        secondVotes: row[2],
        points: row[3]
      });
    });
    if (entries.length === 0) continue; // no entries in this class - skip it entirely

    entries.sort(function (a, b) { return a.rank - b.rank; });

    var title = classMap[numStr] || ('Class ' + numStr);
    body += '<div class="class-header" style="padding:8px;margin-top:14px;">Class ' + numStr + ' - ' + title + '</div>';
    body += '<table><tr><th style="width:50px;">Rank</th><th>Entrant</th><th>Model</th><th>1st Choice</th><th>2nd Choice</th><th>Points</th></tr>';
    entries.forEach(function (e) {
      body += '<tr><td>' + e.rank + '</td><td>' + e.entrantName + '</td><td>"' + e.modelName + '"</td><td>' +
        e.firstVotes + '</td><td>' + e.secondVotes + '</td><td>' + e.points + '</td></tr>';
    });
    body += '</table>';
  }
  buildAndShowPrintDialog_('Class Rankings - Full Results', body);
}

/** Prints the class sponsors list. */
function printSponsorsList() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sponsors');
  var data = ws.getRange(5, 1, 24, 3).getValues(); // Class#, Class Name, Sponsor

  var body = '<p class="sub">Thank you to everyone sponsoring a class this show.</p>';
  body += '<table><tr><th>Class #</th><th>Class</th><th>Sponsor</th></tr>';
  data.forEach(function (r) {
    var sponsor = r[2] ? r[2] : '<em>No sponsor</em>';
    body += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td>' + sponsor + '</td></tr>';
  });
  body += '</table>';
  buildAndShowPrintDialog_('Class Sponsors', body);
}

function printFinalPlacings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tally = ss.getSheetByName('Tally');
  var classWinners = ss.getSheetByName('Class Winners');
  var classMap = getClassTitleMap();
  var emLookup = buildEntryModelsLookup_();

  // Read the Class Placement Tiebreaks table (rows 56-70): Class#, Place, Winning Model#.
  // Builds a lookup like "05-2nd" -> entry number, used to override a tied placing.
  var tieOverrides = {};
  var tieData = classWinners.getRange(56, 1, 15, 3).getValues();
  tieData.forEach(function (row) {
    if (!row[0] || !row[1] || !row[2]) return;
    var classKey = normalizeClassNum(row[0]);
    tieOverrides[classKey + '-' + row[1]] = row[2];
  });

  var placeLabels = { 1: '1st', 2: '2nd', 3: '3rd' };
  var body = '';
  for (var classNum = 1; classNum <= 24; classNum++) {
    var numStr = pad2(classNum);
    var startRow = 6 + (classNum - 1) * 45;
    var blockData = tally.getRange(startRow, 2, 40, 6).getValues(); // B-G: Entry#,1st,2nd,Points,Key,Rank

    var placings = { 1: [], 2: [], 3: [] };
    blockData.forEach(function (row) {
      var rank = row[5];
      if (rank === 1 || rank === 2 || rank === 3) {
        var key = numStr + '-' + row[0];
        var info = emLookup[key] || { modelName: '(unknown)', entrantName: '(unknown)' };
        placings[rank].push(info.entrantName + ' - "' + info.modelName + '"');
      }
    });

    var title = classMap[numStr] || ('Class ' + numStr);
    body += '<div class="class-header" style="padding:8px;margin-top:14px;">Class ' + numStr + ' - ' + title + '</div>';
    body += '<table><tr><th style="width:70px;">Place</th><th>Entrant - Model</th></tr>';
    [3, 2, 1].forEach(function (place) {
      var overrideKey = numStr + '-' + placeLabels[place];
      var names;
      if (tieOverrides.hasOwnProperty(overrideKey)) {
        var overrideEntryNum = tieOverrides[overrideKey];
        var overrideInfo = emLookup[numStr + '-' + overrideEntryNum] || { modelName: '(unknown)', entrantName: '(unknown)' };
        names = overrideInfo.entrantName + ' - "' + overrideInfo.modelName + '" (tiebreak confirmed)';
      } else {
        names = placings[place].length ? placings[place].join('<br>') : '-';
      }
      body += '<tr><td><strong>' + placeLabels[place] + '</strong></td><td>' + names + '</td></tr>';
    });
    body += '</table>';
  }

  // Append Simon Breust, People's Choice, Best in Show, in that order, using the Confirmed
  // Results section (rows 74-76). People's Choice falls back to the auto-computed winner if
  // left blank; Simon Breust and Best in Show always require a manual pick, since neither has
  // an automatic result.
  var trophyRows = classWinners.getRange(74, 1, 3, 5).getValues(); // Award, -, Entry Code, Model Name, Entrant Name
  var trophyLabels = ['Simon Breust Memorial Trophy', "People's Choice", 'Best in Show'];
  body += '<div class="class-header" style="padding:8px;margin-top:20px;">Special Awards</div>';
  body += '<table><tr><th>Award</th><th>Entrant - Model</th></tr>';

  for (var i = 0; i < 3; i++) {
    var award = trophyLabels[i];
    var entryCode = trophyRows[i][2];
    var display;
    if (entryCode) {
      display = trophyRows[i][4] + ' - "' + trophyRows[i][3] + '"';
    } else if (award === "People's Choice") {
      // No manual override - fall back to the auto-computed winner from the People's Choice box.
      var pcRow = classWinners.getRange(34, 1, 1, 5).getValues()[0]; // Class#,Class,Model#,ModelName,EntrantName
      display = pcRow[3] ? (pcRow[4] + ' - "' + pcRow[3] + '"') : 'Not yet determined';
    } else {
      display = '<em>Not yet selected - fill in the Confirmed Results section on Class Winners</em>';
    }
    body += '<tr><td><strong>' + award + '</strong></td><td>' + display + '</td></tr>';
  }
  body += '</table>';

  buildAndShowPrintDialog_('Final Placings', body);
}

/**
 * Closes online/walk-in entries and, in the same step, syncs the numbers that depend on final
 * entry counts:
 *   - Expected Entrants (Show Settings) is set to the actual number of entrants, so the
 *     "all votes cast" completion alert fires at the right number.
 *   - Class Setup's Max Entries per class is set to the actual number of models entered in that
 *     class, tightening the voting page's input validation to match reality.
 * Safe to run more than once - it just re-syncs from whatever's in Entrants/Entry Models right now.
 */
function closeEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName('Show Settings');
  var entrantsSheet = ss.getSheetByName('Entrants');
  var modelsSheet = ss.getSheetByName('Entry Models');
  var classSetupSheet = ss.getSheetByName('Class Setup');

  settings.getRange('B21').setValue(false); // Entries Open? = FALSE

  var entrantsData = entrantsSheet.getDataRange().getValues();
  var entrantCount = 0;
  for (var i = 1; i < entrantsData.length; i++) {
    if (entrantsData[i][0]) entrantCount++;
  }
  // Expected Entrants (B10) is now a live formula that always reflects the actual entrant count,
  // so it no longer needs setting here.

  var modelsData = modelsSheet.getDataRange().getValues();
  var countByClass = {};
  for (var j = 1; j < modelsData.length; j++) {
    var cls = String(modelsData[j][2]);
    if (!cls) continue;
    countByClass[cls] = (countByClass[cls] || 0) + 1;
  }

  var classData = classSetupSheet.getDataRange().getValues();
  var classesUpdated = 0;
  for (var k = 5; k < classData.length; k++) {
    var id = classData[k][0];
    if (!id || String(id) === 'motm') continue;
    var num = String(classData[k][1]);
    var actualCount = countByClass[num] || 0;
    classSetupSheet.getRange(k + 1, 4).setValue(Math.max(actualCount, 1)); // Max Entries, min 1
    classesUpdated++;
  }

  SpreadsheetApp.getUi().alert(
    'Entries closed.\n\n' +
    entrantCount + ' entrants, ' + modelsData.length + ' total model rows.\n' +
    'Max Entries synced for ' + classesUpdated + ' classes.'
  );
}

/** Reopens entries (e.g. if closed by mistake, or you want to allow a late walk-in exception). */
function reopenEntries() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Show Settings').getRange('B21').setValue(true);
  SpreadsheetApp.getUi().alert('Entries are open again.');
}

/** Sets Enabled? = TRUE for every real show class in Class Setup (leaves Model of the Month alone). */
function enableAllShowClasses() {
  setAllShowClassesEnabled(true);
}

/** Sets Enabled? = FALSE for every real show class in Class Setup (leaves Model of the Month alone). */
function disableAllShowClasses() {
  setAllShowClassesEnabled(false);
}

function setAllShowClassesEnabled(value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Class Setup');
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 5; i < data.length; i++) { // rows before this are title/subtitle/blank/header
    var id = data[i][0];
    if (!id || String(id) === 'motm') continue; // Model of the Month stays untouched - toggle it manually
    sheet.getRange(i + 1, 5).setValue(value); // column E = Enabled?
    count++;
  }
  SpreadsheetApp.getUi().alert((value ? 'Enabled' : 'Disabled') + ' all ' + count + ' show classes. Model of the Month was left as-is.');
}

/**
 * Reads the Class Setup sheet and returns ALL real show classes (ignores Enabled?, since that
 * switch controls the voting page, not entry eligibility - entries should always be able to pick
 * from the full class list regardless of what's currently enabled for voting). Excludes Model of
 * the Month. Returns [{number, title}, ...] sorted by class number.
 */
function getEntryFormClasses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Class Setup');
  var data = sheet.getDataRange().getValues();
  var classes = [];
  for (var i = 5; i < data.length; i++) {
    var id = data[i][0];
    if (!id) continue;
    if (String(id) === 'motm') continue; // Model of the Month isn't a competition class entrants enter
    classes.push({
      number: String(data[i][1]),
      title: data[i][2]
    });
  }
  return classes;
}

/** Calculates entry fee: $3 for the first 5 models, $1 each for the next 5, free after that, capped
 * at $20. Junior/Intermediate entrants are completely free regardless of model count. */
function calculateEntryFee(ageGroup, modelCount) {
  if (ageGroup === 'Junior/Intermediate (Under 18)') return 0;
  var fee = 0;
  for (var i = 1; i <= modelCount; i++) {
    if (i <= 5) fee += 3;
    else if (i <= 10) fee += 1;
    // 11th model onward is free
  }
  return Math.min(fee, 20);
}

/**
 * Called from the entry form (EntryForm.html) when an entrant submits.
 * entrant: { firstName, lastName, email, phone, ageGroup, agreedToRules, website (honeypot) }
 * models: [{ classNumber, modelName, brand, scale, australian }, ...]
 */
function submitEntry(entrant, models) {
  // Honeypot check - this field is invisible to real people, only bots fill it in
  if (entrant && entrant.website) {
    return { status: 'ok', totalFee: 0, entries: [] }; // silently pretend success, don't record anything
  }

  if (!isEntriesOpen()) {
    return { status: 'error', message: 'Entries are closed. Please see the show organisers if you need to enter.' };
  }
  if (!entrant || !entrant.firstName || !entrant.lastName || !entrant.email || !entrant.phone || !entrant.ageGroup) {
    return { status: 'error', message: "We're missing a few of your details - please make sure your name, email and phone number are filled in." };
  }
  if (!entrant.agreedToRules) {
    return { status: 'error', message: 'One last thing - please tick the box confirming you\'ve read the competition rules.' };
  }
  if (!models || models.length === 0) {
    return { status: 'error', message: "Don't forget to add at least one model before submitting!" };
  }
  for (var i = 0; i < models.length; i++) {
    var m = models[i];
    if (!m.classNumber || !m.modelName || !m.brand || !m.scale) {
      return { status: 'error', message: 'Looks like a model is missing some details - please make sure class, model name, brand and scale are all filled in for every model.' };
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var entrantsSheet = ss.getSheetByName('Entrants');
    var modelsSheet = ss.getSheetByName('Entry Models');

    // If this email address already has an entry, add these models to it instead of creating a
    // second separate entrant - covers people who submit again to add more models or fix a mistake.
    var entrantsData = entrantsSheet.getDataRange().getValues();
    var cleanEmail = String(entrant.email).trim().toLowerCase();
    var existingRowNum = -1; // 1-based sheet row
    var entrantId, votingCode, priorModelCount;
    for (var r = 1; r < entrantsData.length; r++) {
      if (String(entrantsData[r][4]).trim().toLowerCase() === cleanEmail) {
        existingRowNum = r + 1;
        entrantId = entrantsData[r][0];
        votingCode = entrantsData[r][9];
        priorModelCount = Number(entrantsData[r][7]) || 0;
        break;
      }
    }
    var isUpdate = existingRowNum > -1;

    if (!isUpdate) {
      entrantId = 'E' + new Date().getTime();
      priorModelCount = 0;
      // Assign a voting code now (not shown to the entrant - only printed on their slip at check-in)
      votingCode = assignNextVotingCode(ss);
    }

    // Work out the next free class-entry-number per class already used in Entry Models,
    // so numbers keep incrementing correctly across every entrant, not just this submission.
    var modelData = modelsSheet.getDataRange().getValues(); // includes header row
    var nextNumberByClass = {};
    for (var mr = 1; mr < modelData.length; mr++) {
      var cls = String(modelData[mr][2]);
      var num = Number(modelData[mr][3]) || 0;
      if (!nextNumberByClass[cls] || num >= nextNumberByClass[cls]) {
        nextNumberByClass[cls] = num + 1;
      }
    }

    var assignedEntries = [];
    var newRows = [];
    for (var j = 0; j < models.length; j++) {
      var model = models[j];
      var classNum = String(model.classNumber);
      var entryNum = nextNumberByClass[classNum] || 1;
      nextNumberByClass[classNum] = entryNum + 1;

      var modelId = entrantId + '-' + (priorModelCount + j + 1);
      newRows.push([modelId, entrantId, classNum, entryNum, model.modelName, model.brand, model.scale, !!model.australian]);
      assignedEntries.push({ classNumber: classNum, entryNumber: entryNum, modelName: model.modelName });
    }
    if (newRows.length > 0) {
      var startRow = modelsSheet.getLastRow() + 1;
      // Force Model Name/Brand/Scale columns to plain text BEFORE writing - otherwise Google Sheets
      // sometimes auto-interprets values that look date-like (e.g. a scale of "1/48") as an actual
      // date, silently corrupting the data. Setting the format first stops that from happening.
      modelsSheet.getRange(startRow, 5, newRows.length, 3).setNumberFormat('@');
      modelsSheet.getRange(startRow, 1, newRows.length, 8).setValues(newRows);
    }

    var newTotalModelCount = priorModelCount + models.length;
    var fee = calculateEntryFee(entrant.ageGroup, newTotalModelCount);

    if (isUpdate) {
      entrantsSheet.getRange(existingRowNum, 3).setValue(entrant.firstName);
      entrantsSheet.getRange(existingRowNum, 4).setValue(entrant.lastName);
      entrantsSheet.getRange(existingRowNum, 6).setValue(entrant.phone);
      entrantsSheet.getRange(existingRowNum, 7).setValue(entrant.ageGroup);
      entrantsSheet.getRange(existingRowNum, 8).setValue(newTotalModelCount);
      entrantsSheet.getRange(existingRowNum, 9).setValue(fee);
    } else {
      entrantsSheet.appendRow([
        entrantId, new Date(), entrant.firstName, entrant.lastName, entrant.email, entrant.phone,
        entrant.ageGroup, newTotalModelCount, fee, votingCode, false
      ]);
    }

    // Send a confirmation email listing their full, up-to-date entry (not just what was just
    // submitted) - wrapped so a mail failure never breaks the actual entry submission, but the
    // real error is logged (View > Executions in Apps Script) and surfaced in the response so it's
    // actually possible to diagnose if it fails.
    var emailErrorMessage = null;
    try {
      sendEntryConfirmationEmail(ss, entrant, entrantId, fee, newTotalModelCount, isUpdate);
    } catch (mailErr) {
      emailErrorMessage = mailErr.message;
      Logger.log('Confirmation email failed for ' + entrant.email + ': ' + mailErr.message);
    }

    var result = { status: 'ok', totalFee: fee, entries: assignedEntries, updated: isUpdate, totalModelCount: newTotalModelCount };
    if (!votingCode) {
      result.warning = 'Entry saved, but the voting code pool has run out - generate more codes before the show.';
    }
    if (emailErrorMessage) {
      result.warning = (result.warning ? result.warning + ' ' : '') +
        'Entry saved, but the confirmation email failed to send (' + emailErrorMessage + ').';
    }
    return result;
  } catch (err) {
    return { status: 'error', message: 'Something went wrong: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Returns every model on record for one entrant: [{classNum, entryNum, name}, ...], sorted by class then entry number. */
function getModelsForEntrant(ss, entrantId) {
  var modelsData = ss.getSheetByName('Entry Models').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < modelsData.length; i++) {
    if (modelsData[i][1] === entrantId) {
      list.push({
        classNum: String(modelsData[i][2]),
        entryNum: Number(modelsData[i][3]),
        name: safeText(modelsData[i][4])
      });
    }
  }
  list.sort(function (a, b) {
    if (a.classNum !== b.classNum) return a.classNum < b.classNum ? -1 : 1;
    return a.entryNum - b.entryNum;
  });
  return list;
}

/** Emails the entrant a confirmation listing their full current entry and total fee. */
function sendEntryConfirmationEmail(ss, entrant, entrantId, fee, totalModelCount, isUpdate) {
  if (!entrant.email || entrant.email.indexOf('@') === -1) return;
  var models = getModelsForEntrant(ss, entrantId);
  var classTitles = getClassTitleMap();
  var lines = models.map(function (m) {
    var className = classTitles[normalizeClassNum(m.classNum)] || ('Class ' + pad2(m.classNum));
    return '  - ' + className + ' - ' + m.name;
  });
  var feeText = entrant.ageGroup === 'Junior/Intermediate (Under 18)' ? 'FREE entry' : ('$' + fee + ' (payable on arrival)');
  var hasModifiedDiecast = models.some(function (m) { return normalizeClassNum(m.classNum) === '11'; });
  var modifiedDiecastNote = hasModifiedDiecast
    ? '\n\nA note on your Modified Diecast entry: modified diecast can be any diecast model that has ' +
      'been modified in some way (paint, addons, etc). If possible, please bring a photo of the ' +
      'original model before the modifications, to display alongside your model.'
    : '';

  var subject = 'Northern Area Modellers - Entry Confirmation';
  var intro = isUpdate
    ? 'You have successfully updated your entry, here is an updated list of your entries:'
    : 'Thanks for entering the Northern Area Modellers N.N.L 2026 show! Here is a summary of your entry:';
  var body =
    'Hi ' + entrant.firstName + ',\n\n' +
    intro + '\n\n' +
    'Models entered (' + totalModelCount + ' total):\n' +
    lines.join('\n') + '\n\n' +
    'Amount due: ' + feeText + modifiedDiecastNote + '\n\n' +
    'Upon entry to the show, you will be provided with a unique voting code and a QR code link to ' +
    'the electronic voting slip you can access with an internet connected device. If you don\'t have ' +
    'one of those, a club member will be able to help you submit your votes through our voting portal.\n\n' +
    'If you need to add more models or fix anything, just fill out the entry form again using this ' +
    'same email address - it will be added to this entry rather than creating a new one.\n\n' +
    'See you at the show!\n' +
    'Northern Area Modellers';

  var settings = ss.getSheetByName('Show Settings');
  var fromAddress = String(settings.getRange('B78').getValue() || '').trim();
  var bccAddress = String(settings.getRange('B80').getValue() || '').trim();

  var options = { name: 'Northern Area Modellers' };
  if (bccAddress && bccAddress.indexOf('@') > -1) options.bcc = bccAddress;

  if (fromAddress && fromAddress.indexOf('@') > -1) {
    // GmailApp lets you send as a different address, but ONLY if it's already set up as a verified
    // "Send mail as" alias on the Google account that owns this script - see the Instructions tab.
    // If that alias isn't verified yet, this throws an error, which is caught by the try/catch
    // around this whole call in submitEntry() - so the entry still saves fine either way, it just
    // won't get a confirmation email until the alias is set up.
    options.from = fromAddress;
    GmailApp.sendEmail(entrant.email, subject, body, options);
  } else {
    MailApp.sendEmail(entrant.email, subject, body, options);
  }
}

/**
 * Finds the first Voting Codes entry not already assigned to any entrant (checked against the
 * Entrants sheet's Voting Code column) and returns it. Returns '' if the pool has run out.
 * Doesn't mark anything as "Used?" - that column specifically tracks whether the code has actually
 * been used to cast a vote, which happens later, separately, when the entrant votes at the show.
 */
function assignNextVotingCode(ss) {
  var codesSheet = ss.getSheetByName('Voting Codes');
  var codesData = codesSheet.getDataRange().getValues();
  var entrantsSheet = ss.getSheetByName('Entrants');
  var entrantsData = entrantsSheet.getDataRange().getValues();

  var alreadyAssigned = {};
  for (var i = 1; i < entrantsData.length; i++) {
    var existing = entrantsData[i][9]; // column J = Voting Code
    if (existing) alreadyAssigned[String(existing).trim().toUpperCase()] = true;
  }

  for (var j = 4; j < codesData.length; j++) { // data rows start after the 4 header/title rows
    var code = String(codesData[j][0]).trim().toUpperCase();
    if (!code || alreadyAssigned[code]) continue;
    return code;
  }
  return ''; // ran out of codes
}

/** Returns a simple summary list of all entrants for the print dialog: [{id, name, modelCount, fee, ageGroup, votingCode}, ...] */
function getEntrantsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Entrants');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    list.push({
      id: row[0],
      name: row[2] + ' ' + row[3],
      ageGroup: row[6],
      modelCount: row[7],
      fee: row[8],
      votingCode: row[9] || ''
    });
  }
  return list;
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

/**
 * Guards against a specific Google Sheets gotcha: text that looks date-like (e.g. a scale of
 * "1/48") can get silently auto-converted into an actual date value when written. The write side
 * is now fixed to prevent this going forward (see submitEntry), but this catches any older rows
 * that got corrupted before that fix, so they show a clear flag instead of an ugly raw date string.
 */
function safeText(val) {
  if (val instanceof Date) {
    return '(check entry - format error)';
  }
  return val;
}

/** Builds a map of class number -> class title from Class Setup, e.g. {"05": "Armour - 1/35"}. */
function getClassTitleMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheetByName('Class Setup').getDataRange().getValues();
  var map = {};
  for (var i = 5; i < data.length; i++) {
    var id = data[i][0];
    if (!id) continue;
    map[normalizeClassNum(data[i][1])] = data[i][2];
  }
  return map;
}

/**
 * Google Sheets sometimes auto-converts numeric-looking text like "09" into the actual number 9,
 * stripping the leading zero, depending on how a cell was entered/imported. That silently broke
 * class-name lookups for classes 01-09 (e.g. "09" from one source not matching "9" from another).
 * This normalizes any class number - whether it arrives as "09", 9, or "9" - to the same "09" form,
 * so lookups always match regardless of which representation either side happens to be using.
 */
function normalizeClassNum(val) {
  var s = String(val).trim();
  return /^[0-9]+$/.test(s) ? pad2(s) : s; // leaves non-numeric values like "MOTM" untouched
}

/**
 * Builds the full printable HTML (main entry slip + one A6 slip per model) for one entrant.
 * Returned as a string to the print dialog, which opens it in a new browser tab for printing -
 * this way the personal details never touch a public URL, only the browser session that already
 * has the Sheet open.
 */
function getEntrantPrintHtml(entrantId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entrantsData = ss.getSheetByName('Entrants').getDataRange().getValues();
  var modelsData = ss.getSheetByName('Entry Models').getDataRange().getValues();
  var classTitles = getClassTitleMap();

  var entrant = null;
  for (var i = 1; i < entrantsData.length; i++) {
    if (entrantsData[i][0] === entrantId) {
      entrant = {
        firstName: entrantsData[i][2], lastName: entrantsData[i][3],
        ageGroup: entrantsData[i][6], modelCount: entrantsData[i][7],
        fee: entrantsData[i][8], votingCode: entrantsData[i][9] || 'NO CODE ASSIGNED'
      };
      break;
    }
  }
  if (!entrant) return '<p>Entrant not found.</p>';

  var models = [];
  for (var j = 1; j < modelsData.length; j++) {
    if (modelsData[j][1] === entrantId) {
      models.push({
        classNum: String(modelsData[j][2]),
        entryNum: Number(modelsData[j][3]),
        name: safeText(modelsData[j][4]),
        brand: safeText(modelsData[j][5]),
        scale: safeText(modelsData[j][6])
      });
    }
  }

  var votingUrl = ScriptApp.getService().getUrl();
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=' + encodeURIComponent(votingUrl);
  var feeText = entrant.ageGroup === 'Junior/Intermediate (Under 18)' ? 'FREE' : ('$' + entrant.fee);

  var css = `
    @page { size: 100mm 150mm; margin: 5mm; }
    * { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    body { font-family: Arial, Helvetica, sans-serif; margin: 0; }
    .slip { width: 100%; height: 140mm; page-break-after: always; position: relative; padding: 5mm; }
    .slip:last-child { page-break-after: auto; }
    .main-slip h1 { font-size: 20px; color: #2E4053; margin: 0 0 2px 0; text-align: center; }
    .main-slip h2 { font-size: 14px; color: #2E4053; margin: 0 0 8px 0; text-align: center; }
    .main-slip .name { font-size: 25px; font-weight: bold; text-align: center; margin: 6px 0; }
    .main-slip .row { display: flex; justify-content: space-between; font-size: 16px; margin: 4px 0; }
    .main-slip .model-list-title { font-weight: bold; color: #2E4053; margin: 6px 0 3px 0; }
    .main-slip .model-list { line-height: 1.4; margin: 0 0 6px 0; }
    .main-slip .fee { text-align: center; font-size: 30px; font-weight: bold; color: #C0392B; margin: 6px 0; }
    .main-slip .qr-row { display: flex; align-items: center; gap: 10px; margin-top: 6px; }
    .main-slip .qr-row img { width: 22mm; height: 22mm; }
    .main-slip .code-box { border: 3px solid #2E4053; border-radius: 6px; padding: 5px; text-align: center; flex: 1; }
    .main-slip .code-box .label { font-size: 11px; color: #666; }
    .main-slip .code-box .code { font-size: 22px; font-weight: bold; letter-spacing: 1px; }
    .main-slip .paid-line { margin-top: 8px; font-size: 14px; border-top: 1px dashed #999; padding-top: 6px; }

    .slip.model-slip { display: flex; flex-direction: column; }
    .model-slip .top-spacer { flex: 1 1 auto; min-height: 68mm; }
    .model-slip .show-name { font-size: 20px; font-weight: bold; color: #2E4053; margin-bottom: 6px; text-align: center; }
    .model-slip .show-name img { height: 14mm; display: block; margin: 0 auto 4px auto; }
    .model-slip .field-row { font-size: 22px; margin: 6px 0; display: flex; }
    .model-slip .field-row .field-label { font-weight: bold; color: #2E4053; width: 30mm; flex-shrink: 0; }
    .model-slip .field-row .field-value { font-weight: normal; color: #222; flex: 1; }
    .model-slip .field-row.split { gap: 6mm; }
    .model-slip .field-row.split .half { display: flex; flex: 1; }
    .model-slip .number-box { margin-top: 10px; display: flex; justify-content: flex-end; gap: 6px; }
    .model-slip .num-tile { border: 3px solid; border-radius: 6px; padding: 4px 10px; text-align: center; background: #fff; }
    .model-slip .num-tile .tile-label { font-size: 10px; display: block; font-weight: bold; }
    .model-slip .num-tile .tile-value { font-size: 32px; font-weight: bold; display: block; line-height: 1.1; }
    .model-slip .class-tile { border-color: #2E4053; }
    .model-slip .class-tile .tile-label, .model-slip .class-tile .tile-value { color: #2E4053; }
    .model-slip .entry-tile { border-color: #B8860B; }
    .model-slip .entry-tile .tile-label, .model-slip .entry-tile .tile-value { color: #B8860B; }
  `;

  var html = '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>';

  // --- Main entry slip ---
  // Bigger text for entrants with fewer models, automatically shrinking a little for entrants with
  // a lot of models, so short lists get nice big readable text but long lists still fit the page.
  var listFontSize = models.length <= 5 ? 17 : (models.length <= 8 ? 15 : 13);
  html += '<div class="slip main-slip">';
  html += '<h1>Northern Area Modellers</h1><h2>N.N.L Entry Slip 2026</h2>';
  html += '<div class="name">' + entrant.firstName + ' ' + entrant.lastName + '</div>';
  html += '<div class="row"><span>Models entered:</span><strong>' + entrant.modelCount + '</strong></div>';
  html += '<div class="model-list-title" style="font-size:' + (listFontSize + 1) + 'px;">Entries:</div>';
  html += '<div class="model-list" style="font-size:' + listFontSize + 'px;">';
  models.forEach(function (m) {
    var className = classTitles[normalizeClassNum(m.classNum)] || ('Class ' + pad2(m.classNum));
    html += pad2(m.classNum) + ' - ' + className + ' - ' + m.name + '<br>';
  });
  html += '</div>';
  html += '<div class="fee">' + feeText + '</div>';
  html += '<div class="qr-row"><img id="qrCodeImg" src="' + qrUrl + '"><div class="code-box"><div class="label">YOUR VOTING CODE</div><div class="code">' + entrant.votingCode + '</div></div></div>';
  html += '<div class="paid-line">Amount paid: _______  Received by: _______</div>';
  html += '</div>';

  // --- One A6 slip per model ---
  models.forEach(function (m) {
    var title = classTitles[normalizeClassNum(m.classNum)] || '';
    html += '<div class="slip model-slip">';
    html += '<div class="top-spacer"></div>';
    html += '<div class="show-name">Northern Area Modellers - N.N.L 2026</div>';
    html += '<div class="field-row"><span class="field-label">MODEL</span><span class="field-value">' + m.name + '</span></div>';
    html += '<div class="field-row split">';
    html += '<div class="half"><span class="field-label">BRAND</span><span class="field-value">' + m.brand + '</span></div>';
    html += '<div class="half"><span class="field-label">SCALE</span><span class="field-value">' + m.scale + '</span></div>';
    html += '</div>';
    html += '<div class="field-row"><span class="field-label">Class:</span><span class="field-value">' + title + '</span></div>';
    html += '<div class="number-box">';
    html += '<div class="num-tile class-tile"><span class="tile-label">CLASS</span><span class="tile-value">' + pad2(m.classNum) + '</span></div>';
    html += '<div class="num-tile entry-tile"><span class="tile-label">MODEL #</span><span class="tile-value">' + pad2(m.entryNum) + '</span></div>';
    html += '</div>';
    html += '</div>';
  });

  // Wait for the QR code image to actually finish loading before printing - a fixed timer isn't
  // reliable since it depends on network speed reaching an external image server. Falls back to
  // printing anyway after 4 seconds if the image is slow/blocked, so it's never stuck waiting forever.
  html += '<script>';
  html += 'var qrImg = document.getElementById("qrCodeImg");';
  html += 'var printed = false;';
  html += 'function doPrint() { if (printed) return; printed = true; window.focus(); window.print(); }';
  html += 'if (qrImg) {';
  html += '  if (qrImg.complete) { setTimeout(doPrint, 150); }';
  html += '  else { qrImg.onload = function () { setTimeout(doPrint, 150); }; qrImg.onerror = doPrint; }';
  html += '  setTimeout(doPrint, 4000);';
  html += '} else { setTimeout(doPrint, 150); }';
  html += '</script>';

  html += '</body></html>';
  return html;
}



/** Reads the Class Setup sheet and returns only ENABLED classes: [{id, number, name, maxEntries}, ...] */
function getClassConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Class Setup');
  var data = sheet.getDataRange().getValues();
  var classes = [];
  for (var i = 5; i < data.length; i++) { // rows before this are title/subtitle/blank/header
    var id = data[i][0];
    if (!id) continue;
    var enabled = data[i][4];
    if (!(enabled === true || enabled === 'TRUE')) continue; // skip disabled classes entirely
    var number = String(data[i][1]);
    var titleText = data[i][2];
    classes.push({
      id: String(id),
      number: number,
      name: number + ' - ' + titleText,
      maxEntries: Number(data[i][3]) || 1
    });
  }
  return classes;
}

/** Reads the Show Settings sheet's Voting Open? switch (cell B4). */
function isVotingOpen() {
  var val = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Show Settings').getRange('B4').getValue();
  return val === true || val === 'TRUE';
}

/** Reads the Show Settings sheet's Entries Open? switch (cell B21). */
function isEntriesOpen() {
  var val = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Show Settings').getRange('B21').getValue();
  return val === true || val === 'TRUE';
}

// ===================== MODEL OF THE MONTH (new sheet-local settings) =====================
// Row layout on the "Model of the Month" sheet: B5=Entries Open?, B6=Voting Open?, B7=Expected
// Entrants. Data rows start at row 13: A=Model Number, B=Entrant Name, C=Model Name,
// D=Has Voted?, E=1st Choice votes, F=2nd Choice votes, G=Points (formula), H=Tiebreak Key (formula).
var MOTM_FIRST_DATA_ROW = 13;

function getMotmSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Model of the Month');
}

/** Reads the Model of the Month sheet's Entries Open? switch (cell B5). */
function isMotmEntriesOpen() {
  var val = getMotmSheet_().getRange('B5').getValue();
  return val === true || val === 'TRUE';
}

/** Reads the Model of the Month sheet's Voting Open? switch (cell B6). */
function isMotmVotingOpen() {
  var val = getMotmSheet_().getRange('B6').getValue();
  return val === true || val === 'TRUE';
}

/** Reads every data row currently on the Model of the Month sheet as plain objects. */
function getMotmRows_() {
  var sheet = getMotmSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < MOTM_FIRST_DATA_ROW) return [];
  var data = sheet.getRange(MOTM_FIRST_DATA_ROW, 1, lastRow - MOTM_FIRST_DATA_ROW + 1, 6).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue; // skip fully blank rows (no entrant name)
    rows.push({
      sheetRow: MOTM_FIRST_DATA_ROW + i,
      modelNumber: String(data[i][0] || '').trim(),
      entrantName: String(data[i][1] || '').trim(),
      modelName: String(data[i][2] || '').trim(),
      hasVoted: data[i][3] === true,
      firstChoiceVotes: Number(data[i][4]) || 0,
      secondChoiceVotes: Number(data[i][5]) || 0
    });
  }
  return rows;
}

/**
 * Called from the Model of the Month entry form. Adds one row to the "Model of the Month" sheet.
 * If nothingToEnter is true, the person is registered as eligible to vote without an actual model -
 * their Model Number and Model Name are left blank, which automatically excludes them from the
 * "vote for a model" list (they can still vote for other people's models, and be identified by
 * name for the one-vote-per-person check).
 */
function submitMotmEntry(entrantName, modelName, modelNumber, nothingToEnter, website) {
  if (website) {
    return { status: 'ok' }; // honeypot tripped - silently pretend success, don't record anything
  }
  if (!isMotmEntriesOpen()) {
    return { status: 'error', message: 'Model of the Month entries are not open right now.' };
  }
  entrantName = (entrantName || '').trim();
  modelName = (modelName || '').trim();
  modelNumber = (modelNumber || '').toString().trim();

  if (!entrantName) {
    return { status: 'error', message: 'Please fill in your name.' };
  }
  if (!nothingToEnter && (!modelName || !modelNumber)) {
    return { status: 'error', message: "Please fill in the model name and the number on your model - or tick \"Nothing to enter this month\" if you're just here to vote." };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var rows = getMotmRows_();

    if (!nothingToEnter) {
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].modelNumber && rows[i].modelNumber === modelNumber) {
          return { status: 'error', message: 'Number ' + modelNumber + ' has already been entered this month - please double-check the number on your model.' };
        }
      }
    }

    var sheet = getMotmSheet_();
    var nextRow = MOTM_FIRST_DATA_ROW + rows.length;
    var rowModelNumber = nothingToEnter ? '' : modelNumber;
    var rowModelName = nothingToEnter ? '(Nothing to enter this month)' : modelName;

    sheet.getRange(nextRow, 1, 1, 3).setNumberFormat('@');
    sheet.getRange(nextRow, 1, 1, 6).setValues([[rowModelNumber, entrantName, rowModelName, false, 0, 0]]);

    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: 'Something went wrong: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns everything the voting page needs in one call:
 *  - voters: [{name, hasVoted}] for every registered person this month (including those with
 *    nothing to enter) - used for the "who are you" dropdown.
 *  - models: [{number, name, entrantName}] for entries that actually have a model - used for the
 *    1st/2nd choice dropdowns. The voting page itself filters out the selected voter's own model.
 */
function getMotmVotingData() {
  var rows = getMotmRows_();
  var voters = rows.map(function (r) { return { name: r.entrantName, hasVoted: r.hasVoted }; });
  var models = rows
    .filter(function (r) { return r.modelNumber; })
    .map(function (r) { return { number: r.modelNumber, name: r.modelName, entrantName: r.entrantName }; });
  return { voters: voters, models: models };
}

/**
 * Records one person's 1st and 2nd choice votes. Enforces: the voter must be a registered
 * entrant who hasn't voted yet this month, the two choices must be different real models, and
 * neither choice can be the voter's own model.
 */
function submitMotmVote(voterName, firstChoiceNumber, secondChoiceNumber) {
  if (!isMotmVotingOpen()) {
    return { status: 'error', message: 'Model of the Month voting is not open right now.' };
  }
  voterName = (voterName || '').trim();
  firstChoiceNumber = (firstChoiceNumber || '').toString().trim();
  secondChoiceNumber = (secondChoiceNumber || '').toString().trim();

  if (!voterName) {
    return { status: 'error', message: 'Please select your name.' };
  }
  if (!firstChoiceNumber || !secondChoiceNumber) {
    return { status: 'error', message: 'Please choose both a 1st and 2nd favourite.' };
  }
  if (firstChoiceNumber === secondChoiceNumber) {
    return { status: 'error', message: 'Your 1st and 2nd choices need to be two different models.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var rows = getMotmRows_();

    var voterRow = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].entrantName === voterName) { voterRow = rows[i]; break; }
    }
    if (!voterRow) {
      return { status: 'error', message: "We couldn't find your name on this month's entry list - please check with whoever's running Model of the Month." };
    }
    if (voterRow.hasVoted) {
      return { status: 'error', message: "It looks like you've already voted this month - only one vote per person." };
    }

    var firstRow = null, secondRow = null;
    for (var j = 0; j < rows.length; j++) {
      if (rows[j].modelNumber === firstChoiceNumber) firstRow = rows[j];
      if (rows[j].modelNumber === secondChoiceNumber) secondRow = rows[j];
    }
    if (!firstRow || !secondRow) {
      return { status: 'error', message: "One of those models wasn't found - please try again." };
    }
    if (firstRow.entrantName === voterName || secondRow.entrantName === voterName) {
      return { status: 'error', message: "You can't vote for your own model." };
    }

    var sheet = getMotmSheet_();
    sheet.getRange(firstRow.sheetRow, 5).setValue(firstRow.firstChoiceVotes + 1);
    sheet.getRange(secondRow.sheetRow, 6).setValue(secondRow.secondChoiceVotes + 1);
    sheet.getRange(voterRow.sheetRow, 4).setValue(true);

    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: 'Something went wrong: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Reads the Show Settings sheet's Voting Closes At (cell B6). Returns an ISO string, or null if blank. */
function getVotingClosesAt() {
  var val = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Show Settings').getRange('B6').getValue();
  if (!val || !(val instanceof Date)) return null;
  return val.toISOString();
}

/**
 * Called from the spectator page (SpectatorIndex.html) - People's Choice only. Uses its own separate
 * pool of one-time codes (Spectator Codes sheet) so spectators can't vote more than once, without
 * touching the entrant Voting Codes or the class 1st/2nd columns.
 * code: the string typed in from a spectator's slip
 * answers: object like { peoplesChoice: "16-2" }
 */
function submitSpectatorVote(code, answers) {
  if (!isVotingOpen()) {
    return { status: 'error', message: "Voting hasn't started just yet - hang tight for the announcement, then come on back!" };
  }
  if (!answers || !answers.peoplesChoice) {
    return { status: 'error', message: "Please enter a class number and model number for People's Choice." };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var codesSheet = ss.getSheetByName('Spectator Codes');
    var codesData = codesSheet.getDataRange().getValues();

    var cleanCode = String(code).trim().toUpperCase();
    var rowIndex = -1;
    for (var j = 4; j < codesData.length; j++) { // data rows start after the 4 header/title rows
      if (String(codesData[j][0]).trim().toUpperCase() === cleanCode) {
        rowIndex = j;
        break;
      }
    }

    if (rowIndex === -1) {
      return { status: 'error', message: "Hmm, we don't recognise that code - have a check for typos and try again." };
    }
    if (codesData[rowIndex][1] === true || codesData[rowIndex][1] === 'TRUE') {
      return { status: 'error', message: "Looks like this code's already been used to vote. If that doesn't seem right, please see the show organisers." };
    }

    codesSheet.getRange(rowIndex + 1, 2).setValue(true);
    codesSheet.getRange(rowIndex + 1, 3).setValue(new Date());

    var respSheet = ss.getSheetByName('Form Responses');
    var headerRow = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0];
    var newRow = new Array(headerRow.length).fill('');
    newRow[0] = new Date();
    var pcCol = headerRow.indexOf("Special Award - People's Choice");
    if (pcCol > -1) newRow[pcCol] = answers.peoplesChoice;
    respSheet.appendRow(newRow);

    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: 'Something went wrong: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Called from the browser (Index.html) when a voter submits.
 * code: the string typed in from their slip
 * answers: object like { class01_1st: "3", class01_2nd: "4", ..., peoplesChoice: "2C" }
 */
function submitVote(code, answers) {
  if (!isVotingOpen()) {
    return { status: 'error', message: "Voting hasn't started just yet - hang tight for the announcement, then come on back!" };
  }

  var classes = getClassConfig();
  for (var i = 0; i < classes.length; i++) {
    var cid = classes[i].id;
    if (!answers[cid + '_1st'] || !answers[cid + '_2nd']) {
      return { status: 'error', message: 'Please pick both a 1st and 2nd favourite in every class.' };
    }
    if (answers[cid + '_1st'] === answers[cid + '_2nd']) {
      return { status: 'error', message: 'Your 1st and 2nd favourite in ' + classes[i].name + ' cannot be the same entry.' };
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // wait up to 30s if another vote is being processed at the same instant
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var codesSheet = ss.getSheetByName('Voting Codes');
    var codesData = codesSheet.getDataRange().getValues();

    var cleanCode = String(code).trim().toUpperCase();
    var rowIndex = -1;
    for (var j = 4; j < codesData.length; j++) { // data rows start after the 4 header/title rows
      if (String(codesData[j][0]).trim().toUpperCase() === cleanCode) {
        rowIndex = j;
        break;
      }
    }

    if (rowIndex === -1) {
      return { status: 'error', message: "Hmm, we don't recognise that code - have a check for typos and try again." };
    }
    if (codesData[rowIndex][1] === true || codesData[rowIndex][1] === 'TRUE') {
      return { status: 'error', message: "Looks like this code's already been used to vote. If that doesn't seem right, please see the show organisers." };
    }

    // Mark the code as used
    codesSheet.getRange(rowIndex + 1, 2).setValue(true);
    codesSheet.getRange(rowIndex + 1, 3).setValue(new Date());

    // Record the vote - matched up by column header name, so row order in Class Setup doesn't matter
    var respSheet = ss.getSheetByName('Form Responses');
    var headerRow = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0];
    var newRow = new Array(headerRow.length).fill('');
    newRow[0] = new Date();

    classes.forEach(function (cls) {
      var col1st = headerRow.indexOf(cls.number + ' 1st Choice');
      var col2nd = headerRow.indexOf(cls.number + ' 2nd Choice');
      if (col1st > -1) newRow[col1st] = 'Entry ' + answers[cls.id + '_1st'];
      if (col2nd > -1) newRow[col2nd] = 'Entry ' + answers[cls.id + '_2nd'];
    });

    var pcCol = headerRow.indexOf("Special Award - People's Choice");
    if (pcCol > -1) newRow[pcCol] = answers.peoplesChoice;

    respSheet.appendRow(newRow);

    checkAndAlertIfComplete(ss);

    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: 'Something went wrong: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * After every vote, checks whether Votes Cast So Far has reached Expected Entrants (Show Settings
 * sheet). If so - and an alert hasn't already been sent - emails Alert Email once and marks
 * Alert Sent? as TRUE so it doesn't fire again on every subsequent vote.
 */
function checkAndAlertIfComplete(ss) {
  var settings = ss.getSheetByName('Show Settings');
  var expected = Number(settings.getRange('B10').getValue()) || 0;
  var castSoFar = Number(settings.getRange('B14').getValue()) || 0;
  var alreadySent = settings.getRange('B18').getValue();
  alreadySent = alreadySent === true || alreadySent === 'TRUE';

  if (expected > 0 && castSoFar >= expected && !alreadySent) {
    sendCompletionEmail(settings, castSoFar, expected, false);
    settings.getRange('B18').setValue(true);
  }
}

/**
 * Manual override - triggered from the "NAM Voting" menu (see onOpen). Sends the completion alert
 * immediately regardless of the current vote count, for cases like an entrant leaving early without
 * voting. Safe to click more than once.
 */
function sendManualCompletionAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName('Show Settings');
  var expected = Number(settings.getRange('B10').getValue()) || 0;
  var castSoFar = Number(settings.getRange('B14').getValue()) || 0;
  sendCompletionEmail(settings, castSoFar, expected, true);
  settings.getRange('B18').setValue(true);
  SpreadsheetApp.getUi().alert('Completion alert sent to ' + settings.getRange('B12').getValue() + '.');
}

function sendCompletionEmail(settings, castSoFar, expected, isManual) {
  var email = settings.getRange('B12').getValue();
  if (!email || String(email).indexOf('@') === -1) return; // no valid address configured, skip silently
  var subject = isManual
    ? 'NAM Voting - manually marked complete'
    : 'NAM Voting - all expected votes are in!';
  var body = isManual
    ? 'Voting was manually marked complete via the "Send completion alert now" menu.\n\n' +
      'Votes cast so far: ' + castSoFar + '\nExpected entrants: ' + expected
    : 'All expected votes have been cast.\n\n' +
      'Votes cast: ' + castSoFar + '\nExpected entrants: ' + expected;
  MailApp.sendEmail(email, subject, body);
}
