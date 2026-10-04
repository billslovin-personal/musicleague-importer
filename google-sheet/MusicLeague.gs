/**
 * @OnlyCurrentDoc
 *
 * Music League master sheet builder (Google Apps Script).
 *
 * Builds three tabs — Songs, Rounds, Points — from Music League export zip
 * files uploaded through the "Music League" menu. The raw export data is kept
 * in hidden tabs of this spreadsheet and de-duplicated on every upload, so you
 * can upload just the newest zip each time (re-uploading an old one is safe).
 *
 * Permissions: the @OnlyCurrentDoc tag above limits this script to THIS
 * spreadsheet. It never reads or writes anything else in Google Drive.
 *
 * Setup: see SETUP.md.
 */

var DEFAULT_LEAGUE_NAME = 'MFFL VIII';

// Hidden tabs that hold the de-duplicated raw export data.
var RAW = {
  competitors: { sheet: 'raw_competitors', cols: ['ID', 'Name'] },
  rounds: { sheet: 'raw_rounds', cols: ['ID', 'Created', 'Name', 'Description', 'Playlist URL'] },
  submissions: { sheet: 'raw_submissions', cols: ['Spotify URI', 'Title', 'Album', 'Artist(s)', 'Submitter ID', 'Round ID'] },
  votes: { sheet: 'raw_votes', cols: ['Spotify URI', 'Voter ID', 'Points Assigned', 'Round ID'] }
};
var SETTINGS_SHEET = 'settings';

// ---- Menu ----------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Music League')
    .addItem('Upload zip files…', 'showUploadDialog')
    .addItem('Rebuild tabs', 'rebuildFromStoredData')
    .addSeparator()
    .addItem('Set league name…', 'setLeagueName')
    .addItem('Clear all stored data…', 'clearAllData')
    .addToUi();
}

function showUploadDialog() {
  var html = HtmlService.createHtmlOutput(UPLOAD_HTML).setWidth(420).setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Music League export zips');
}

function setLeagueName() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('League name',
    'Current: ' + getLeagueName_() + '\n\nEnter the league name to show on the Rounds tab:',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim()) return;
  settingsSheet_().getRange('B1').setValue(res.getResponseText().trim());
  rebuildFromStoredData();
}

function clearAllData() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Clear all stored data?',
    'This removes every uploaded season from this sheet. You would need to upload the zips again.',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  Object.keys(RAW).forEach(function (k) { writeRaw_(RAW[k], []); });
  rebuildFromStoredData();
}

// ---- Upload / rebuild ----------------------------------------------------

/**
 * Called from the upload dialog.
 * @param {Array<{name: string, data: string}>} files  base64-encoded zip files
 */
function importZips(files) {
  files.sort(function (a, b) { return naturalCompare_(a.name, b.name); });
  var datasets = [readStoredData_()];
  files.forEach(function (f) {
    var blob = Utilities.newBlob(Utilities.base64Decode(f.data), 'application/zip', f.name);
    datasets.push(readZip_(blob, f.name));
  });
  var result = buildLeagueTables(datasets, getLeagueName_());
  saveRaw_(result.raw);
  writeSheets_(SpreadsheetApp.getActive(), result);
  return summary_(result, files.length + ' zip file(s) imported.');
}

function rebuildFromStoredData() {
  var result = buildLeagueTables([readStoredData_()], getLeagueName_());
  writeSheets_(SpreadsheetApp.getActive(), result);
  SpreadsheetApp.getActive().toast(summary_(result, 'Tabs rebuilt.'), 'Music League', 8);
}

function summary_(result, prefix) {
  var s = result.stats;
  return prefix + ' Now holding ' + s[4][1] + ' competitors, ' + s[3][1] + ' rounds, ' +
    s[2][1] + ' submissions.';
}

function readZip_(blob, name) {
  var tables = { name: name };
  Utilities.unzip(blob).forEach(function (b) {
    var m = b.getName().match(/(?:^|\/)(competitors|rounds|submissions|votes)\.csv$/i);
    if (m) tables[m[1].toLowerCase()] = parseCsvObjects(b.getDataAsString('UTF-8'));
  });
  Object.keys(RAW).forEach(function (t) {
    if (!tables[t]) throw new Error(name + ' is missing ' + t + '.csv — is it a Music League export?');
  });
  return tables;
}

// ---- Raw data storage (hidden tabs) ---------------------------------------

function readStoredData_() {
  var ss = SpreadsheetApp.getActive();
  var data = { name: 'stored' };
  Object.keys(RAW).forEach(function (k) {
    var def = RAW[k], sh = ss.getSheetByName(def.sheet), rows = [];
    if (sh && sh.getLastRow() > 1) {
      rows = sh.getRange(2, 1, sh.getLastRow() - 1, def.cols.length).getValues();
    }
    data[k] = rows.map(function (r) {
      var o = {};
      def.cols.forEach(function (c, i) { o[c] = r[i]; });
      return o;
    });
  });
  return data;
}

function saveRaw_(raw) {
  Object.keys(RAW).forEach(function (k) {
    var def = RAW[k];
    writeRaw_(def, raw[k].map(function (o) { return def.cols.map(function (c) { return o[c]; }); }));
  });
}

function writeRaw_(def, rows) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(def.sheet) || ss.insertSheet(def.sheet);
  sh.clear();
  var all = [def.cols].concat(rows);
  if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
  // Plain text so IDs and URIs are stored exactly as exported.
  sh.getRange(1, 1, all.length, def.cols.length).setNumberFormat('@').setValues(all);
  sh.hideSheet();
}

function settingsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange('A1:B1').setValues([['League name', DEFAULT_LEAGUE_NAME]]);
    sh.hideSheet();
  }
  return sh;
}

function getLeagueName_() {
  return String(settingsSheet_().getRange('B1').getValue() || DEFAULT_LEAGUE_NAME);
}

// ---- Core logic (no Google services; also runnable in Node for testing) ---

/**
 * @param {Array<{competitors, rounds, submissions, votes}>} datasets
 *   In processing order; each table is an array of row objects keyed by the
 *   export's column names. Later datasets win when the same record repeats.
 */
function buildLeagueTables(datasets, leagueName) {
  var competitors = new Map(); // ID -> row
  var rounds = new Map();      // Round ID -> row
  var submissions = new Map(); // Round ID|URI -> row
  var votes = new Map();       // Round ID|URI|Voter ID -> row

  datasets.forEach(function (t) {
    t.competitors.forEach(function (c) {
      var id = clean_(c.ID);
      if (id) competitors.set(id, { 'ID': id, 'Name': clean_(c.Name) });
    });
    t.rounds.forEach(function (r) {
      var id = clean_(r.ID);
      if (id) rounds.set(id, {
        'ID': id, 'Created': clean_(r.Created), 'Name': clean_(r.Name),
        'Description': clean_(r.Description), 'Playlist URL': clean_(r['Playlist URL'])
      });
    });
    t.submissions.forEach(function (s) {
      var roundId = clean_(s['Round ID']), uri = clean_(s['Spotify URI']);
      if (uri) submissions.set(roundId + '|' + uri, {
        'Spotify URI': uri, 'Title': clean_(s.Title), 'Album': clean_(s.Album),
        'Artist(s)': clean_(s['Artist(s)']), 'Submitter ID': clean_(s['Submitter ID']), 'Round ID': roundId
      });
    });
    t.votes.forEach(function (v) {
      var roundId = clean_(v['Round ID']), uri = clean_(v['Spotify URI']), voter = clean_(v['Voter ID']);
      if (uri) votes.set(roundId + '|' + uri + '|' + voter, {
        'Spotify URI': uri, 'Voter ID': voter,
        'Points Assigned': Number(clean_(v['Points Assigned'])) || 0, 'Round ID': roundId
      });
    });
  });

  // Points per submission (a song within a specific round).
  var subPoints = new Map();
  votes.forEach(function (v) {
    var key = v['Round ID'] + '|' + v['Spotify URI'];
    subPoints.set(key, (subPoints.get(key) || 0) + v['Points Assigned']);
  });

  function nameOf(id) { return competitors.has(id) ? competitors.get(id).Name : id; }

  var songs = [];
  submissions.forEach(function (s, key) {
    var round = rounds.get(s['Round ID']);
    songs.push({
      artist: s['Artist(s)'], title: s.Title, album: s.Album,
      competitor: nameOf(s['Submitter ID']), points: subPoints.get(key) || 0,
      round: round ? round.Name : '', submitterId: s['Submitter ID']
    });
  });
  songs.sort(function (a, b) {
    return ciCompare_(a.artist, b.artist) || ciCompare_(a.title, b.title) || ciCompare_(a.round, b.round);
  });

  var roundList = Array.from(rounds.values()).sort(function (a, b) {
    return a.Created < b.Created ? -1 : a.Created > b.Created ? 1 : 0;
  });

  var totals = new Map();
  competitors.forEach(function (_, id) { totals.set(id, 0); });
  songs.forEach(function (s) { totals.set(s.submitterId, (totals.get(s.submitterId) || 0) + s.points); });
  var points = [];
  totals.forEach(function (pts, id) { points.push([nameOf(id), pts]); });
  points.sort(function (a, b) { return b[1] - a[1] || ciCompare_(a[0], b[0]); });

  var artists = new Set(), uniqueSongs = new Set();
  songs.forEach(function (s) {
    artists.add(s.artist.toLowerCase());
    uniqueSongs.add((s.artist + '|' + s.title).toLowerCase());
  });

  return {
    songs: songs.map(function (s) { return [s.artist, s.title, s.album, s.competitor, s.points, s.round]; }),
    rounds: roundList.map(function (r) { return [r.Name, r.Description, r['Playlist URL'], leagueName]; }),
    points: points,
    stats: [
      ['Total Unique Artists', artists.size],
      ['Total Songs', uniqueSongs.size],
      ['Total Submissions', songs.length],
      ['Total Rounds', roundList.length],
      ['Total Competitors', competitors.size]
    ],
    raw: {
      competitors: Array.from(competitors.values()),
      rounds: Array.from(rounds.values()),
      submissions: Array.from(submissions.values()),
      votes: Array.from(votes.values())
    }
  };
}

/** RFC 4180 CSV parser (handles quoted commas, quotes and line breaks). */
function parseCsvObjects(text) {
  text = text.replace(/^﻿/, '');
  var rows = [], row = [], field = '', i = 0, inQuotes = false, n = text.length;
  while (i < n) {
    var ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === '"') { inQuotes = true; i++; continue; }
    if (ch === ',') { row.push(field); field = ''; i++; continue; }
    if (ch === '\r' || ch === '\n') {
      row.push(field); field = '';
      rows.push(row); row = [];
      if (ch === '\r' && text[i + 1] === '\n') i++;
      i++; continue;
    }
    field += ch; i++;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }

  var header = rows.shift() || [];
  return rows
    .filter(function (r) { return r.some(function (v) { return v !== ''; }); })
    .map(function (r) {
      var o = {};
      header.forEach(function (h, idx) { o[h.trim()] = r[idx] === undefined ? '' : r[idx]; });
      return o;
    });
}

function clean_(v) { return v == null ? '' : String(v).trim(); }

function ciCompare_(a, b) {
  a = a.toLowerCase(); b = b.toLowerCase();
  return a < b ? -1 : a > b ? 1 : 0;
}

function naturalCompare_(a, b) {
  var re = /(\d+)|(\D+)/g;
  var pa = a.toLowerCase().match(re) || [], pb = b.toLowerCase().match(re) || [];
  for (var i = 0; i < Math.min(pa.length, pb.length); i++) {
    var x = pa[i], y = pb[i];
    if (x === y) continue;
    var nx = /^\d/.test(x), ny = /^\d/.test(y);
    if (nx && ny) return Number(x) - Number(y);
    return x < y ? -1 : 1;
  }
  return pa.length - pb.length;
}

// ---- Sheet output --------------------------------------------------------

var TITLE_BG = '#000000', HEADER_BG = '#434343', HEADER_FG = '#ff0000';

function writeSheets_(ss, result) {
  var songs = writeTable_(ss, 'Songs', 'Artist / Song Master',
    ['Artist(s)', 'Title', 'Album', 'Competitors Name', 'Points Assigned', 'Rounds Name'],
    [220, 260, 260, 150, 110, 300], result.songs);

  // League stats beside the song list.
  songs.setColumnWidth(7, 30);
  songs.setColumnWidth(8, 170);
  songs.setColumnWidth(9, 80);
  songs.getRange('H1:I2').merge().setValue('League Stats')
    .setBackground(TITLE_BG).setFontColor(HEADER_FG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  songs.getRange(3, 8, result.stats.length, 2).setValues(result.stats).setFontWeight('bold');
  songs.getRange(3, 8, result.stats.length, 1).setFontColor('#0000ff');

  var rounds = writeTable_(ss, 'Rounds', 'Rounds Master',
    ['Name', 'Description', 'Playlist URL', 'League'],
    [280, 500, 380, 100], result.rounds);
  if (result.rounds.length) {
    var links = result.rounds.map(function (r) {
      var b = SpreadsheetApp.newRichTextValue().setText(r[2] || '');
      if (r[2]) b.setLinkUrl(r[2]);
      return [b.build()];
    });
    rounds.getRange(3, 3, links.length, 1).setRichTextValues(links);
  }

  writeTable_(ss, 'Points', '"Lifetime" League Points',
    ['Competitor Name', 'Total Points'], [200, 110], result.points);

  // Put the three tabs first, in order.
  ['Songs', 'Rounds', 'Points'].forEach(function (name, i) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName('Songs'));
}

function writeTable_(ss, name, title, headers, widths, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  var needRows = rows.length + 2;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());

  var cols = headers.length;
  sh.getRange(1, 1, 1, cols).merge().setValue(title)
    .setBackground(TITLE_BG).setFontColor(HEADER_FG).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, cols).setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(3, 1, rows.length, cols).setValues(rows)
      .setVerticalAlignment('top').setWrap(true);
  }
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);
  sh.getRange(2, 1, rows.length + 1, cols).createFilter();
  return sh;
}

// ---- Upload dialog -------------------------------------------------------

var UPLOAD_HTML = [
  '<style>body{font-family:Arial,sans-serif;font-size:14px}#msg{margin-top:12px;white-space:pre-wrap}',
  'button{margin-top:12px;padding:6px 14px}</style>',
  '<div>Choose one or more Music League export <b>.zip</b> files. Already-imported',
  ' data is skipped automatically.</div>',
  '<input type="file" id="files" accept=".zip" multiple style="margin-top:12px">',
  '<br><button id="go" onclick="go()">Import</button>',
  '<div id="msg"></div>',
  '<script>',
  'function read(f){return new Promise(function(ok,fail){var r=new FileReader();',
  ' r.onload=function(){ok({name:f.name,data:r.result.split(",")[1]})};r.onerror=fail;r.readAsDataURL(f);});}',
  'function go(){var fs=document.getElementById("files").files,m=document.getElementById("msg");',
  ' if(!fs.length){m.textContent="Pick at least one zip file.";return;}',
  ' document.getElementById("go").disabled=true;m.textContent="Importing "+fs.length+" file(s)… this can take a minute.";',
  ' Promise.all(Array.prototype.map.call(fs,read)).then(function(files){',
  '  google.script.run.withSuccessHandler(function(s){m.textContent="Done! "+s;',
  '   setTimeout(function(){google.script.host.close();},2500);})',
  '  .withFailureHandler(function(e){m.textContent="Error: "+e.message;document.getElementById("go").disabled=false;})',
  '  .importZips(files);});}',
  '</script>'
].join('\n');

// Allow the core logic to be tested in Node.
if (typeof module !== 'undefined') {
  module.exports = { buildLeagueTables: buildLeagueTables, parseCsvObjects: parseCsvObjects, naturalCompare_: naturalCompare_ };
}
