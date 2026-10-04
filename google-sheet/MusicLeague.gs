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
 * "Look up genres & years" also calls the public Deezer and MusicBrainz APIs
 * with each song's artist and title (nothing else is sent).
 *
 * Setup: see SETUP.md.
 */

var DEFAULT_LEAGUE_NAME = 'MFFL VIII';

// Hidden tabs that hold the de-duplicated raw export data.
var RAW = {
  competitors: { sheet: 'raw_competitors', cols: ['ID', 'Name', 'Name As Of'] },
  rounds: { sheet: 'raw_rounds', cols: ['ID', 'Created', 'Name', 'Description', 'Playlist URL'] },
  submissions: { sheet: 'raw_submissions', cols: ['Spotify URI', 'Title', 'Album', 'Artist(s)', 'Submitter ID', 'Round ID'] },
  votes: { sheet: 'raw_votes', cols: ['Spotify URI', 'Voter ID', 'Points Assigned', 'Round ID'] }
};
var SETTINGS_SHEET = 'settings';
// Looked-up genre/year per song, keyed by Spotify URI. Not cleared by "Clear all".
var SONG_INFO = { sheet: 'raw_song_info', cols: ['Spotify URI', 'Artist', 'Title', 'Genre', 'Year', 'Looked Up'] };
var LOOKUP_TIME_BUDGET_MS = 4.5 * 60 * 1000; // Apps Script stops scripts at 6 minutes

// ---- Menu ----------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Music League')
    .addItem('Upload zip files…', 'showUploadDialog')
    .addItem('Look up genres & years', 'lookUpGenresAndYears')
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
  // A song-info .csv (from a previous lookup) fills in genres/years without looking them up again.
  var csvs = files.filter(function (f) { return /\.csv$/i.test(f.name); });
  files = files.filter(function (f) { return !/\.csv$/i.test(f.name); });
  csvs.forEach(function (f) {
    importSongInfoCsv_(Utilities.newBlob(Utilities.base64Decode(f.data)).getDataAsString('UTF-8'), f.name);
  });

  files.sort(function (a, b) { return naturalCompare_(a.name, b.name); });
  var datasets = [readStoredData_()];
  files.forEach(function (f) {
    var blob = Utilities.newBlob(Utilities.base64Decode(f.data), 'application/zip', f.name);
    datasets.push(readZip_(blob, f.name));
  });
  var result = buildLeagueTables(datasets, getLeagueName_(), readSongInfo_());
  saveRaw_(result.raw);
  writeSheets_(SpreadsheetApp.getActive(), result);
  var missing = result.songsMissingInfo.length;
  return summary_(result, files.length + ' zip file(s)' + (csvs.length ? ' and song info' : '') + ' imported.') +
    (missing ? ' Run "Look up genres & years" to fill in ' + missing + ' new song(s).' : '');
}

function rebuildFromStoredData() {
  var result = buildLeagueTables([readStoredData_()], getLeagueName_(), readSongInfo_());
  writeSheets_(SpreadsheetApp.getActive(), result);
  SpreadsheetApp.getActive().toast(summary_(result, 'Tabs rebuilt.'), 'Music League', 8);
  return result;
}

function summary_(result, prefix) {
  var s = result.stats;
  return prefix + ' Now holding ' + s[4][1] + ' competitors, ' + s[3][1] + ' rounds, ' +
    s[2][1] + ' submissions.';
}

/**
 * Looks up genre and year for songs that don't have them yet, for as long as
 * Google allows one run, then rebuilds the tabs. Run again to continue.
 */
function lookUpGenresAndYears() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var pending = buildLeagueTables([readStoredData_()], getLeagueName_(), readSongInfo_()).songsMissingInfo;
  if (!pending.length) {
    ui.alert('Every song already has its genre and year looked up.');
    return;
  }
  ss.toast('Looking up ' + pending.length + ' song(s). This can take up to 5 minutes…', 'Music League', 30);

  var deadline = Date.now() + LOOKUP_TIME_BUDGET_MS;
  var sh = ss.getSheetByName(SONG_INFO.sheet);
  if (!sh) {
    sh = ss.insertSheet(SONG_INFO.sheet);
    sh.getRange(1, 1, 1, SONG_INFO.cols.length).setValues([SONG_INFO.cols]);
    sh.hideSheet();
  }
  var done = 0;
  for (var i = 0; i < pending.length && Date.now() < deadline; i += 10) {
    var rows = lookupSongBatch_(pending.slice(i, i + 10), deadline);
    if (rows.length) {
      // Save as we go so nothing is lost if Google stops the script.
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, SONG_INFO.cols.length)
        .setNumberFormat('@').setValues(rows);
      done += rows.length;
    }
  }

  var left = rebuildFromStoredData().songsMissingInfo.length;
  ui.alert(left
    ? 'Looked up ' + done + ' song(s). ' + left + ' still to go — run "Look up genres & years" again to continue.'
    : 'Done! Looked up ' + done + ' song(s). Every song now has its genre and year looked up.');
}

function importSongInfoCsv_(text, name) {
  var rows = parseCsvObjects(text);
  if (rows.length && !('Spotify URI' in rows[0] && 'Genre' in rows[0] && 'Year' in rows[0])) {
    throw new Error(name + " doesn't look like a song-info file (needs Spotify URI, Genre and Year columns).");
  }
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SONG_INFO.sheet) || ss.insertSheet(SONG_INFO.sheet);
  var merged = new Map();
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, SONG_INFO.cols.length).getValues()
      .forEach(function (r) { merged.set(String(r[0]), r); });
  }
  rows.forEach(function (o) {
    var uri = clean_(o['Spotify URI']);
    if (uri) merged.set(uri, SONG_INFO.cols.map(function (c) { return clean_(o[c]); }));
  });
  var all = [SONG_INFO.cols].concat(Array.from(merged.values()));
  sh.clear();
  if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
  sh.getRange(1, 1, all.length, SONG_INFO.cols.length).setNumberFormat('@').setValues(all);
  sh.hideSheet();
}

function readSongInfo_() {
  var info = new Map();
  var sh = SpreadsheetApp.getActive().getSheetByName(SONG_INFO.sheet);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, SONG_INFO.cols.length).getValues().forEach(function (r) {
      info.set(String(r[0]), { genre: String(r[3] || ''), year: r[4] ? Number(r[4]) : null });
    });
  }
  return info;
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

// ---- Genre / year lookup ---------------------------------------------------
//
// Genre comes from Deezer (album genre: Rock, Pop, Alternative, ...). Year is
// the earliest believable year from MusicBrainz (first release of any matching
// recording), the Deezer recording's ISRC year code, and the Deezer album date,
// because the album Spotify/Deezer link to is often a later reissue.

var MB_USER_AGENT = 'MusicLeagueSheet/1.0 (Google Sheets script for a private music league)';

/**
 * Looks up up to ~10 songs. Returns rows for SONG_INFO; songs that hit a
 * temporary error (rate limit, outage) are left out so a later run retries them.
 * @param {Array<{uri, artist, title}>} songs
 */
function lookupSongBatch_(songs, deadline) {
  var fetchOpts = { muteHttpExceptions: true };
  function searchDeezer(list, exact) {
    return UrlFetchApp.fetchAll(list.map(function (s) {
      return Object.assign({ url: deezerSearchUrl_(s.artist, s.title, exact) }, fetchOpts);
    })).map(function (res, i) {
      var json = parseJson_(res);
      if (!json || json.error) return { failed: true };
      return { track: pickDeezerTrack_(json, list[i].artist, list[i].title) };
    });
  }
  var tracks = searchDeezer(songs, false);
  // Plain search can rank covers/karaoke first; retry misses with an exact artist + title search.
  var retry = [];
  tracks.forEach(function (t, i) { if (!t.failed && !t.track) retry.push(i); });
  if (retry.length) {
    searchDeezer(retry.map(function (i) { return songs[i]; }), true).forEach(function (t, j) {
      if (t.failed || t.track) tracks[retry[j]] = t;
    });
  }

  var albumIds = [];
  tracks.forEach(function (t) {
    if (t.track && albumIds.indexOf(t.track.album.id) < 0) albumIds.push(t.track.album.id);
  });
  var albums = {};
  UrlFetchApp.fetchAll(albumIds.map(function (id) {
    return Object.assign({ url: 'https://api.deezer.com/album/' + id }, fetchOpts);
  })).forEach(function (res, i) {
    var json = parseJson_(res);
    albums[albumIds[i]] = json && !json.error ? json : null;
  });

  var rows = [];
  var today = new Date().toISOString().slice(0, 10);
  for (var i = 0; i < songs.length; i++) {
    if (Date.now() > deadline) break;
    var s = songs[i], t = tracks[i];
    var started = Date.now();
    var mb = fetchMusicBrainz_(s.artist, s.title);
    // MusicBrainz allows about one request per second.
    Utilities.sleep(Math.max(0, 1100 - (Date.now() - started)));
    if (t.failed || mb.failed) continue;

    var album = t.track ? albums[t.track.album.id] : null;
    if (t.track && !album) continue; // album lookup failed; retry later
    var year = earliestYear_([
      mbEarliestYear_(mb.json, s.artist, s.title),
      t.track ? isrcYear_(t.track.isrc) : null,
      album ? parseInt(String(album.release_date || '').slice(0, 4), 10) : null
    ]);
    rows.push([s.uri, s.artist, s.title, album ? deezerGenre_(album) : '', year || '', today]);
  }
  return rows;
}

function fetchMusicBrainz_(artist, title) {
  var opts = { muteHttpExceptions: true, headers: { 'User-Agent': MB_USER_AGENT, 'Accept': 'application/json' } };
  for (var attempt = 0; attempt < 2; attempt++) {
    var res = UrlFetchApp.fetch(musicBrainzUrl_(artist, title), opts);
    if (res.getResponseCode() === 200) return { json: parseJson_(res) };
    Utilities.sleep(2000);
  }
  return { failed: true };
}

function parseJson_(res) {
  if (res.getResponseCode() !== 200) return null;
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

// Pure helpers below are shared with the Node test harness.

function deezerSearchUrl_(artist, title, exact) {
  var a = primaryArtist_(artist), t = baseTitle_(title);
  var q = exact ? 'artist:"' + a.replace(/"/g, '') + '" track:"' + t.replace(/"/g, '') + '"' : a + ' ' + t;
  return 'https://api.deezer.com/search?limit=25&q=' + encodeURIComponent(q);
}

function musicBrainzUrl_(artist, title) {
  var esc = function (s) { return s.replace(/["\\]/g, '\\$&'); };
  var q = 'recording:"' + esc(baseTitle_(title)) + '" AND artist:"' + esc(primaryArtist_(artist)) + '"';
  return 'https://musicbrainz.org/ws/2/recording?fmt=json&limit=100&query=' + encodeURIComponent(q);
}

function pickDeezerTrack_(json, artist, title) {
  var t = matchKey_(baseTitle_(title));
  return (json.data || []).filter(function (x) {
    var xt = matchKey_(baseTitle_(x.title_short || x.title || ''));
    return x.artist && isArtist_(x.artist.name, artist) && x.album &&
      (xt === t || xt.indexOf(t) === 0 || t.indexOf(xt) === 0);
  })[0] || null;
}

/**
 * Spotify joins multiple artists with commas, but some names contain one
 * ("Tyler, The Creator"), so accept any leading run of comma-separated parts.
 */
function isArtist_(name, artists) {
  var key = matchKey_(name), parts = String(artists).split(',');
  for (var i = 1; i <= parts.length; i++) {
    if (matchKey_(parts.slice(0, i).join(',')) === key) return true;
  }
  return false;
}

function deezerGenre_(album) {
  var g = album.genres && album.genres.data && album.genres.data[0];
  return g ? g.name : '';
}

function mbEarliestYear_(json, artist, title) {
  if (!json) return null;
  var t = matchKey_(baseTitle_(title));
  var years = (json.recordings || []).filter(function (r) {
    return matchKey_(baseTitle_(r.title || '')) === t && (r['artist-credit'] || []).some(function (c) {
      return isArtist_(c.name || '', artist) || (c.artist && isArtist_(c.artist.name || '', artist));
    });
  }).map(function (r) { return parseInt(String(r['first-release-date'] || '').slice(0, 4), 10); });
  return earliestYear_(years);
}

/** ISRCs look like CC-XXX-YY-NNNNN; YY is the year the code was assigned. */
function isrcYear_(isrc) {
  var m = String(isrc || '').match(/^[A-Z]{2}[A-Z0-9]{3}(\d{2})\d{5}$/i);
  if (!m) return null;
  var yy = Number(m[1]), nowYY = new Date().getFullYear() % 100;
  return yy <= nowYY ? 2000 + yy : 1900 + yy;
}

function earliestYear_(years) {
  var max = new Date().getFullYear();
  var ok = years.filter(function (y) { return y && y >= 1900 && y <= max; });
  return ok.length ? Math.min.apply(null, ok) : null;
}

function primaryArtist_(artists) { return String(artists).split(',')[0].trim(); }

/** "Doctor My Eyes - Remastered" -> "Doctor My Eyes"; drops (Live), (feat. X), etc. */
function baseTitle_(title) {
  return String(title)
    .replace(/\s+-\s+.*$/, '')
    .replace(/\s*[\(\[][^\)\]]*(remaster|version|live|mono|stereo|edit|mix|feat|with|from)[^\)\]]*[\)\]]/ig, '')
    .trim() || String(title).trim();
}

function matchKey_(s) {
  var k = String(s).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/&/g, 'and').replace(/^the\s+/, '')
    .replace(/[\u00a0-\u00bf\u2000-\u206f]/g, '')   // curly quotes, dashes, etc.
    .replace(/[^a-z0-9\u00c0-\uffff]/g, '');        // keep non-Latin letters
  return k || String(s).toLowerCase().trim();
}

// ---- Core logic (no Google services; also runnable in Node for testing) ---

/**
 * @param {Array<{competitors, rounds, submissions, votes}>} datasets
 *   In processing order; each table is an array of row objects keyed by the
 *   export's column names. Later datasets win when the same record repeats,
 *   except competitor names: the name from the export with the most recent
 *   rounds wins, so upload order doesn't matter.
 * @param {Map<string, {genre, year}>} [songInfo]  looked-up info by Spotify URI
 */
function buildLeagueTables(datasets, leagueName, songInfo) {
  songInfo = songInfo || new Map();
  var competitors = new Map(); // ID -> row
  var rounds = new Map();      // Round ID -> row
  var submissions = new Map(); // Round ID|URI -> row
  var votes = new Map();       // Round ID|URI|Voter ID -> row

  datasets.forEach(function (t) {
    // An export's names are current as of its newest round.
    var exportAsOf = t.rounds.reduce(function (max, r) {
      var created = clean_(r.Created);
      return created > max ? created : max;
    }, '');
    t.competitors.forEach(function (c) {
      var id = clean_(c.ID);
      if (!id) return;
      var asOf = clean_(c['Name As Of']) || exportAsOf;
      var existing = competitors.get(id);
      if (existing && existing['Name As Of'] > asOf) return;
      competitors.set(id, { 'ID': id, 'Name': clean_(c.Name), 'Name As Of': asOf });
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

  var songs = [], songsMissingInfo = [], missingSeen = new Set();
  submissions.forEach(function (s, key) {
    var round = rounds.get(s['Round ID']);
    var uri = s['Spotify URI'], info = songInfo.get(uri);
    if (!info && !missingSeen.has(uri)) {
      missingSeen.add(uri);
      songsMissingInfo.push({ uri: uri, artist: s['Artist(s)'], title: s.Title });
    }
    songs.push({
      artist: s['Artist(s)'], title: s.Title, album: s.Album,
      competitor: nameOf(s['Submitter ID']), points: subPoints.get(key) || 0,
      round: round ? round.Name : '', roundId: s['Round ID'], submitterId: s['Submitter ID'],
      genre: info ? info.genre : '', year: info && info.year ? info.year : '', lookedUp: !!info
    });
  });
  songs.sort(function (a, b) {
    return ciCompare_(a.artist, b.artist) || ciCompare_(a.title, b.title) || ciCompare_(a.round, b.round);
  });

  var roundList = Array.from(rounds.values()).sort(function (a, b) {
    return a.Created < b.Created ? -1 : a.Created > b.Created ? 1 : 0;
  });

  // Finishing place of every song in its round (ties share a place: 1, 1, 3).
  // Rounds with no votes yet (still in progress) are skipped.
  var votedRounds = new Set();
  votes.forEach(function (v) { votedRounds.add(v['Round ID']); });
  var placeOf = new Map(), byRound = new Map();
  songs.forEach(function (s) {
    if (!votedRounds.has(s.roundId)) return;
    if (!byRound.has(s.roundId)) byRound.set(s.roundId, []);
    byRound.get(s.roundId).push(s);
  });
  byRound.forEach(function (list) {
    list.forEach(function (s) {
      placeOf.set(s, 1 + list.filter(function (o) { return o.points > s.points; }).length);
    });
  });

  // Per-competitor totals.
  var comp = new Map();
  competitors.forEach(function (_, id) { comp.set(id, { points: 0, songs: 0, wins: 0, top3: 0 }); });
  songs.forEach(function (s) {
    if (!comp.has(s.submitterId)) comp.set(s.submitterId, { points: 0, songs: 0, wins: 0, top3: 0 });
    var c = comp.get(s.submitterId), place = placeOf.get(s);
    c.points += s.points;
    c.songs++;
    if (place === 1) c.wins++;
    if (place <= 3) c.top3++;
  });
  var points = [];
  comp.forEach(function (c, id) {
    points.push([nameOf(id), c.points, c.songs, c.songs ? Math.round(c.points / c.songs * 100) / 100 : '', c.wins, c.top3]);
  });
  points.sort(function (a, b) { return b[1] - a[1] || ciCompare_(a[0], b[0]); });

  // All-time top 25 songs (ties at 25th place are all included).
  var ranked = songs.slice().sort(function (a, b) { return b.points - a.points || ciCompare_(a.artist, b.artist); });
  var topSongs = [], rank = 0;
  for (var i = 0; i < ranked.length; i++) {
    var s = ranked[i];
    if (!i || s.points !== ranked[i - 1].points) rank = i + 1;
    if (rank > 25) break;
    topSongs.push([rank, s.artist, s.title, s.competitor, s.round, s.points]);
  }

  // Biggest fan: the voter who has given each competitor the most points overall.
  var given = new Map(); // submitter ID -> Map(voter ID -> points)
  votes.forEach(function (v) {
    var sub = submissions.get(v['Round ID'] + '|' + v['Spotify URI']);
    if (!sub || sub['Submitter ID'] === v['Voter ID']) return;
    var id = sub['Submitter ID'];
    if (!given.has(id)) given.set(id, new Map());
    var m = given.get(id);
    m.set(v['Voter ID'], (m.get(v['Voter ID']) || 0) + v['Points Assigned']);
  });
  var fans = [];
  given.forEach(function (m, id) {
    var best = Math.max.apply(null, Array.from(m.values()));
    if (best <= 0) return;
    var who = [];
    m.forEach(function (pts, voter) { if (pts === best) who.push(nameOf(voter)); });
    fans.push([nameOf(id), who.sort(ciCompare_).join(' & '), best]);
  });
  fans.sort(function (a, b) { return b[2] - a[2] || ciCompare_(a[0], b[0]); });

  var artists = new Set(), uniqueSongs = new Set();
  songs.forEach(function (s) {
    artists.add(s.artist.toLowerCase());
    uniqueSongs.add((s.artist + '|' + s.title).toLowerCase());
  });

  // Share of submissions and average points by genre and by decade.
  function breakdown(labelOf, order) {
    var groups = new Map();
    songs.forEach(function (s) {
      var label = labelOf(s), g = groups.get(label) || { n: 0, pts: 0 };
      g.n++;
      g.pts += s.points;
      groups.set(label, g);
    });
    return Array.from(groups.entries()).map(function (e) {
      return [e[0], e[1].n, songs.length ? e[1].n / songs.length : 0, Math.round(e[1].pts / e[1].n * 100) / 100];
    }).sort(order);
  }
  var NOT_YET = 'Not looked up yet', UNKNOWN = 'Unknown';
  function lastIfUnknown(a, b) {
    var ua = a[0] === UNKNOWN || a[0] === NOT_YET, ub = b[0] === UNKNOWN || b[0] === NOT_YET;
    return ua !== ub ? (ua ? 1 : -1) : 0;
  }
  var byGenre = breakdown(function (s) {
    return s.genre || (s.lookedUp ? UNKNOWN : NOT_YET);
  }, function (a, b) { return lastIfUnknown(a, b) || b[1] - a[1] || ciCompare_(a[0], b[0]); });
  var byDecade = breakdown(function (s) {
    return s.year ? Math.floor(s.year / 10) * 10 + 's' : (s.lookedUp ? UNKNOWN : NOT_YET);
  }, function (a, b) { return lastIfUnknown(a, b) || ciCompare_(a[0], b[0]); });

  return {
    songs: songs.map(function (s) { return [s.artist, s.title, s.album, s.competitor, s.points, s.round, s.genre, s.year]; }),
    byGenre: byGenre,
    byDecade: byDecade,
    topSongs: topSongs,
    fans: fans,
    songsMissingInfo: songsMissingInfo,
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
    ['Artist(s)', 'Title', 'Album', 'Competitors Name', 'Points Assigned', 'Rounds Name', 'Genre', 'Year'],
    [220, 260, 260, 150, 110, 300, 120, 60], result.songs);

  // League stats beside the song list.
  songs.setColumnWidth(9, 30);
  songs.setColumnWidth(10, 170);
  songs.setColumnWidth(11, 80);
  songs.getRange('J1:K2').merge().setValue('League Stats')
    .setBackground(TITLE_BG).setFontColor(HEADER_FG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  songs.getRange(3, 10, result.stats.length, 2).setValues(result.stats).setFontWeight('bold');
  songs.getRange(3, 10, result.stats.length, 1).setFontColor('#0000ff');

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
    ['Competitor Name', 'Total Points', 'Songs', 'Avg Points / Song', 'Round Wins', 'Top 3 Finishes'],
    [200, 100, 70, 130, 100, 120], result.points);

  // Stats tab: side-by-side tables, each followed by a blank spacer column.
  var stats = ss.getSheetByName('Stats') || ss.insertSheet('Stats');
  stats.clear();
  stats.getRange(1, 1, stats.getMaxRows(), stats.getMaxColumns()).breakApart();
  var blocks = [
    ['All-Time Top 25 Songs', ['Rank', 'Artist(s)', 'Title', 'Competitor', 'Round', 'Points'],
      [50, 180, 220, 130, 220, 60], result.topSongs, []],
    ['Biggest Fans', ['Competitor', 'Biggest Fan', 'Points Given'],
      [150, 170, 95], result.fans, []],
    ['Songs by Genre', ['Genre', 'Songs', '% of Songs', 'Avg Points'],
      [140, 65, 90, 85], result.byGenre, [null, null, '0.0%', '0.00']],
    ['Songs by Decade Recorded', ['Decade', 'Songs', '% of Songs', 'Avg Points'],
      [140, 65, 90, 85], result.byDecade, [null, null, '0.0%', '0.00']]
  ];
  var needCols = blocks.reduce(function (n, b) { return n + b[1].length + 1; }, 0);
  if (stats.getMaxColumns() < needCols) stats.insertColumnsAfter(stats.getMaxColumns(), needCols - stats.getMaxColumns());
  var needRows = 2 + Math.max.apply(null, blocks.map(function (b) { return b[3].length; }));
  if (stats.getMaxRows() < needRows) stats.insertRowsAfter(stats.getMaxRows(), needRows - stats.getMaxRows());
  var col = 1;
  blocks.forEach(function (b) {
    writeBlock_(stats, col, b[0], b[1], b[3], b[4]);
    b[2].forEach(function (w, i) { stats.setColumnWidth(col + i, w); });
    stats.setColumnWidth(col + b[1].length, 30);
    col += b[1].length + 1;
  });
  stats.setFrozenRows(2);

  // Put the tabs first, in order.
  ['Songs', 'Rounds', 'Points', 'Stats'].forEach(function (name, i) {
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

/** Writes a titled table starting at column `col`; `formats` are per-column number formats. */
function writeBlock_(sh, col, title, headers, rows, formats) {
  var cols = headers.length;
  sh.getRange(1, col, 1, cols).merge().setValue(title)
    .setBackground(TITLE_BG).setFontColor(HEADER_FG).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, col, 1, cols).setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold').setWrap(true);
  if (!rows.length) return;
  sh.getRange(3, col, rows.length, cols).setValues(rows).setVerticalAlignment('top').setWrap(true);
  (formats || []).forEach(function (f, i) {
    if (f) sh.getRange(3, col + i, rows.length, 1).setNumberFormat(f);
  });
}

// ---- Upload dialog -------------------------------------------------------

var UPLOAD_HTML = [
  '<style>body{font-family:Arial,sans-serif;font-size:14px}#msg{margin-top:12px;white-space:pre-wrap}',
  'button{margin-top:12px;padding:6px 14px}</style>',
  '<div>Choose one or more Music League export <b>.zip</b> files (and optionally a',
  ' <b>song-info.csv</b>). Already-imported data is skipped automatically.</div>',
  '<input type="file" id="files" accept=".zip,.csv" multiple style="margin-top:12px">',
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
  module.exports = {
    buildLeagueTables: buildLeagueTables, parseCsvObjects: parseCsvObjects, naturalCompare_: naturalCompare_,
    lookupSongBatch_: lookupSongBatch_, baseTitle_: baseTitle_, matchKey_: matchKey_, isrcYear_: isrcYear_
  };
}
