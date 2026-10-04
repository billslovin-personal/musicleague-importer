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
 * Setup: see README.md in the google-sheet folder.
 */

// League name used if a zip's file name is somehow empty (see leagueFromFileName_).
var DEFAULT_LEAGUE_NAME = 'Music League';

// Hidden tabs that hold the de-duplicated raw export data.
var RAW = {
  competitors: { sheet: 'raw_competitors', cols: ['ID', 'Name', 'Name As Of'] },
  rounds: { sheet: 'raw_rounds', cols: ['ID', 'Created', 'Name', 'Description', 'Playlist URL', 'League'] },
  submissions: { sheet: 'raw_submissions', cols: ['Spotify URI', 'Title', 'Album', 'Artist(s)', 'Submitter ID', 'Round ID'] },
  votes: { sheet: 'raw_votes', cols: ['Spotify URI', 'Voter ID', 'Points Assigned', 'Round ID'] }
};
// Looked-up genre/year per song, keyed by Spotify URI.
var SONG_INFO = { sheet: 'raw_song_info', cols: ['Spotify URI', 'Artist', 'Title', 'Genre', 'Year', 'Looked Up', 'Lead Artist', 'Genre Source'] };
var LOOKUP_TIME_BUDGET_MS = 4.5 * 60 * 1000; // Apps Script stops scripts at 6 minutes

// ---- Menu ----------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Music League')
    .addItem('Upload zip files…', 'showUploadDialog')
    .addItem('Look up genres & years', 'lookUpGenresAndYears')
    .addItem('Rebuild tabs', 'rebuildFromStoredData')
    .addSeparator()
    .addItem('League name…', 'showLeagueNameDialog')
    .addSeparator()
    .addItem('Start a new league (erase everything)…', 'startNewLeague')
    .addToUi();
}

function showUploadDialog() {
  var html = HtmlService.createHtmlOutput(UPLOAD_HTML).setWidth(420).setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Music League export zips');
}

// ---- League name setting ---------------------------------------------------
//
// Uploaded zips' rounds are labelled either with the zip's file name (the
// default) or with a fixed name entered in the "League name…" dialog. The
// choice is saved with the spreadsheet and applies to every upload after it.

var LEAGUE_MODE_PROP = 'leagueNameMode';  // 'file' or 'custom'
var LEAGUE_NAME_PROP = 'leagueNameCustom';

function getLeagueSetting_() {
  var props = PropertiesService.getDocumentProperties();
  var name = props.getProperty(LEAGUE_NAME_PROP) || '';
  var mode = props.getProperty(LEAGUE_MODE_PROP) === 'custom' && name ? 'custom' : 'file';
  return { mode: mode, name: name };
}

/** Called from the League name dialog. */
function saveLeagueSetting(mode, name) {
  name = clean_(name);
  if (mode === 'custom' && !name) throw new Error('Enter a league name, or choose "Use the zip file name(s)".');
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(LEAGUE_MODE_PROP, mode === 'custom' ? 'custom' : 'file');
  if (name) props.setProperty(LEAGUE_NAME_PROP, name);
  return mode === 'custom'
    ? 'Zips you upload from now on will be labelled "' + name + '".'
    : 'Zips you upload from now on will be labelled with their file names.';
}

function showLeagueNameDialog() {
  var s = getLeagueSetting_();
  var html = HtmlService.createHtmlOutput(LEAGUE_HTML
    .replace('__MODE__', function () { return JSON.stringify(s.mode); })
    .replace('__NAME__', function () { return JSON.stringify(s.name).replace(/</g, '\\u003c'); }))
    .setWidth(420).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(html, 'League name');
}

/** League name for an uploaded zip's rounds, per the saved setting. */
function leagueForUpload_(fileName, setting) {
  return setting.mode === 'custom' ? setting.name : leagueFromFileName_(fileName);
}

/** Erases all league data and looked-up song info, leaving empty tabs. */
function startNewLeague() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Erase everything and start a new league?',
    'This permanently deletes ALL songs, rounds, votes, points, stats and looked-up ' +
    'genres/years from this spreadsheet.\n\n' +
    'Tip: make a backup first with File → Make a copy.\n\nThis cannot be undone. Continue?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActive();
  var hidden = Object.keys(RAW).map(function (k) { return RAW[k].sheet; }).concat([SONG_INFO.sheet]);
  hidden.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  // A custom name from the old league wouldn't fit the new one.
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty(LEAGUE_MODE_PROP);
  props.deleteProperty(LEAGUE_NAME_PROP);
  rebuildFromStoredData(); // leaves empty tabs

  ui.alert('Everything has been erased, and the league name is back to using file names. ' +
    'Upload the new league\'s zip files with Music League → Upload zip files… (name each zip ' +
    'after its league, e.g. "MFFL X.zip", or set a name with Music League → League name…).');
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
  var datasets = [readStoredData_()], leagueSetting = getLeagueSetting_();
  files.forEach(function (f) {
    var blob = Utilities.newBlob(Utilities.base64Decode(f.data), 'application/zip', f.name);
    datasets.push(readZip_(blob, f.name, leagueForUpload_(f.name, leagueSetting)));
  });
  var result = buildLeagueTables(datasets, readSongInfo_());
  saveRaw_(result.raw);
  writeSheets_(SpreadsheetApp.getActive(), result);
  var missing = result.songsMissingInfo.length;
  return summary_(result, files.length + ' zip file(s)' + (csvs.length ? ' and song info' : '') + ' imported.') +
    (missing ? ' Run "Look up genres & years" to fill in ' + missing + ' new song(s).' : '');
}

function rebuildFromStoredData() {
  var result = buildLeagueTables([readStoredData_()], readSongInfo_());
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
  var pending = buildLeagueTables([readStoredData_()], readSongInfo_()).songsMissingInfo;
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
      info.set(String(r[0]), {
        genre: String(r[3] || ''), year: r[4] ? Number(r[4]) : null, lead: String(r[6] || ''), source: String(r[7] || '')
      });
    });
  }
  return info;
}

function readZip_(blob, name, league) {
  var tables = { name: name, league: league };
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


// ---- Genre / year lookup ---------------------------------------------------
//
// Genre comes from MusicBrainz (fine-grained: "indie rock", "post-punk", ...),
// taking the top genre of the original recording, else its album, else its
// artist. If MusicBrainz has none, Deezer's broad album genre (Rock, Pop, ...)
// is used instead. Year is the earliest believable year from MusicBrainz (first
// release of any matching recording), the Deezer recording's ISRC year code,
// and the Deezer album date, because the album Spotify/Deezer link to is often
// a later reissue. The lead artist is Deezer's name for the matched track.

var MB_USER_AGENT = 'MusicLeagueSheet/1.0 (Google Sheets script for a private music league)';
var MB_MIN_INTERVAL_MS = 1100; // MusicBrainz allows about one request per second
var mbLastRequestAt_ = 0;
var mbArtistGenreCache_ = {};  // artist MBID -> genre, for the current run

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
    if (t.failed) continue;
    var album = t.track ? albums[t.track.album.id] : null;
    if (t.track && !album) continue; // album lookup failed; retry later
    var mb = fetchMusicBrainz_(musicBrainzUrl_(s.artist, s.title));
    if (mb.failed) continue;
    var mbGenre = musicBrainzGenre_(mb.json, s.artist, s.title);
    if (mbGenre.failed) continue;

    var year = earliestYear_([
      mbEarliestYear_(mb.json, s.artist, s.title),
      t.track ? isrcYear_(t.track.isrc) : null,
      album ? parseInt(String(album.release_date || '').slice(0, 4), 10) : null
    ]);
    var deezerGenre = album ? deezerGenre_(album) : '';
    var genre = mbGenre.genre || deezerGenre;
    var source = mbGenre.genre ? 'MusicBrainz' : deezerGenre ? 'Deezer' : '';
    // The matched Deezer track's main artist, spelled as Deezer has it.
    var lead = t.track ? t.track.artist.name : '';
    rows.push([s.uri, s.artist, s.title, genre, year || '', today, lead, source]);
  }
  return rows;
}

/**
 * Fine-grained genre from MusicBrainz for the song's original recording (the
 * matching recording released first): its top genre, else its first album's,
 * else its artist's. Returns {genre} ('' if none) or {failed: true}.
 */
function musicBrainzGenre_(searchJson, artist, title) {
  var rec = mbOriginalRecording_(searchJson, artist, title);
  if (!rec) return { genre: '' };
  var base = 'https://musicbrainz.org/ws/2/';

  var full = fetchMusicBrainz_(base + 'recording/' + rec.id + '?fmt=json&inc=genres+releases+release-groups+artists');
  if (full.failed) return full;
  var genre = topGenre_(full.json);
  if (genre) return { genre: genre };

  var releases = (full.json.releases || []).slice().sort(function (a, b) {
    return String(a.date || '9999').localeCompare(String(b.date || '9999'));
  });
  if (releases.length && releases[0]['release-group']) {
    var rg = fetchMusicBrainz_(base + 'release-group/' + releases[0]['release-group'].id + '?fmt=json&inc=genres');
    if (rg.failed) return rg;
    genre = topGenre_(rg.json);
    if (genre) return { genre: genre };
  }

  var credit = (full.json['artist-credit'] || [])[0];
  if (credit && credit.artist) {
    var id = credit.artist.id;
    if (!(id in mbArtistGenreCache_)) {
      var a = fetchMusicBrainz_(base + 'artist/' + id + '?fmt=json&inc=genres');
      if (a.failed) return a;
      mbArtistGenreCache_[id] = topGenre_(a.json);
    }
    return { genre: mbArtistGenreCache_[id] };
  }
  return { genre: '' };
}

/** Rate-limited MusicBrainz GET with one retry. Returns {json} or {failed: true}. */
function fetchMusicBrainz_(url) {
  var opts = { muteHttpExceptions: true, headers: { 'User-Agent': MB_USER_AGENT, 'Accept': 'application/json' } };
  for (var attempt = 0; attempt < 2; attempt++) {
    Utilities.sleep(Math.max(0, MB_MIN_INTERVAL_MS - (Date.now() - mbLastRequestAt_)));
    mbLastRequestAt_ = Date.now();
    var res = UrlFetchApp.fetch(url, opts);
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
 * Spotify joins multiple artists with ", ", but some names contain one
 * ("Tyler, The Creator", "Crosby, Stills & Nash"), so `name` matches if it
 * equals any leading run of the parts (i.e. it's the lead artist).
 */
function isArtist_(name, artists) {
  var key = matchKey_(name), parts = String(artists).split(', ');
  for (var i = 1; i <= parts.length; i++) {
    if (matchKey_(parts.slice(0, i).join(', ')) === key) return true;
  }
  return false;
}

/** Lead artist when Deezer didn't identify one: the first name in the export. */
function leadArtistGuess_(artists) { return String(artists).split(', ')[0].trim(); }

function deezerGenre_(album) {
  var g = album.genres && album.genres.data && album.genres.data[0];
  return g ? g.name : '';
}

/** MusicBrainz search results that are this song by this artist. */
function mbMatchingRecordings_(json, artist, title) {
  if (!json) return [];
  var t = matchKey_(baseTitle_(title));
  return (json.recordings || []).filter(function (r) {
    return matchKey_(baseTitle_(r.title || '')) === t && (r['artist-credit'] || []).some(function (c) {
      return isArtist_(c.name || '', artist) || (c.artist && isArtist_(c.artist.name || '', artist));
    });
  });
}

function mbEarliestYear_(json, artist, title) {
  return earliestYear_(mbMatchingRecordings_(json, artist, title).map(function (r) {
    return parseInt(String(r['first-release-date'] || '').slice(0, 4), 10);
  }));
}

/** The matching recording released first (most likely the original, not a live or compilation copy). */
function mbOriginalRecording_(json, artist, title) {
  return mbMatchingRecordings_(json, artist, title).sort(function (a, b) {
    return String(a['first-release-date'] || '9999').localeCompare(String(b['first-release-date'] || '9999'));
  })[0] || null;
}

/** Highest-voted genre on a MusicBrainz entity, title-cased ("indie rock" -> "Indie Rock"). */
function topGenre_(entity) {
  var genres = ((entity && entity.genres) || []).slice().sort(function (a, b) { return b.count - a.count; });
  return genres.length ? titleCaseGenre_(genres[0].name) : '';
}

/**
 * Deezer's broad genre names, renamed to match the MusicBrainz genres they
 * overlap with, so fallback genres merge with MusicBrainz ones. Deezer genres
 * not listed (Rock, Pop, Blues, ...) already use the same name.
 */
var DEEZER_TO_MUSICBRAINZ_GENRE = {
  'Alternative': 'Alternative Rock',
  'Rap/Hip Hop': 'Hip Hop',
  'Electro': 'Electronic',
  'Films/Games': 'Soundtrack',
  'Soul & Funk': 'Soul',
  'Singer & Songwriter': 'Singer-Songwriter',
  'Latin Music': 'Latin',
  'Kids': "Children's Music"
};

/** Genre as shown in the sheet. Only Deezer-sourced names are renamed. */
function displayGenre_(genre, source) {
  return source === 'Deezer' && DEEZER_TO_MUSICBRAINZ_GENRE[genre] || genre;
}

function titleCaseGenre_(name) {
  return String(name).split(/(\s+|-|\/)/).map(function (w, i) {
    if (/&/.test(w)) return w.toUpperCase(); // "r&b" -> "R&B"
    if (i > 0 && /^(and|of|the|n)$/.test(w)) return w; // "drum and bass", "rock n roll"
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join('');
}

/**
 * ISRCs look like CC-XXX-YY-NNNNN; YY is the year the code was assigned (some
 * labels use the original recording year). Two digits can't tell 1937 from a
 * typo, so years before 1950 are ignored.
 */
function isrcYear_(isrc) {
  var m = String(isrc || '').match(/^[A-Z]{2}[A-Z0-9]{3}(\d{2})\d{5}$/i);
  if (!m) return null;
  var yy = Number(m[1]), nowYY = new Date().getFullYear() % 100;
  var year = yy <= nowYY ? 2000 + yy : 1900 + yy;
  return year >= 1950 ? year : null;
}

function earliestYear_(years) {
  var max = new Date().getFullYear();
  var ok = years.filter(function (y) { return y && y >= 1900 && y <= max; });
  return ok.length ? Math.min.apply(null, ok) : null;
}

function primaryArtist_(artists) { return leadArtistGuess_(artists); }

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
 * @param {Array<{competitors, rounds, submissions, votes, league}>} datasets
 *   In processing order; each table is an array of row objects keyed by the
 *   export's column names. Later datasets win when the same record repeats,
 *   except competitor names: the name from the export with the most recent
 *   rounds wins, so upload order doesn't matter. `league` is the league name
 *   for a zip's rounds (stored rounds carry their own League column).
 * @param {Map<string, {genre, year, lead}>} [songInfo]  looked-up info by Spotify URI
 */
function buildLeagueTables(datasets, songInfo) {
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
        'Description': clean_(r.Description), 'Playlist URL': clean_(r['Playlist URL']),
        'League': t.league || clean_(r.League) || DEFAULT_LEAGUE_NAME
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

  // Points per submission (a song within a specific round), plus who gave it
  // points and the biggest single vote, for Music League's tie-breakers.
  var subPoints = new Map(), subVoters = new Map();
  votes.forEach(function (v) {
    var key = v['Round ID'] + '|' + v['Spotify URI'], pts = v['Points Assigned'];
    subPoints.set(key, (subPoints.get(key) || 0) + pts);
    if (!subVoters.has(key)) subVoters.set(key, { up: new Set(), maxVote: 0 });
    var sv = subVoters.get(key);
    if (pts > 0) sv.up.add(v['Voter ID']);
    if (pts > sv.maxVote) sv.maxVote = pts;
  });
  var NO_VOTERS = { up: new Set(), maxVote: 0 };

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
      lead: (info && info.lead) || leadArtistGuess_(s['Artist(s)']),
      competitor: nameOf(s['Submitter ID']), points: subPoints.get(key) || 0,
      round: round ? round.Name : '', roundId: s['Round ID'], submitterId: s['Submitter ID'],
      voters: subVoters.get(key) || NO_VOTERS,
      genre: info ? displayGenre_(info.genre, info.source) : '',
      year: info && info.year ? info.year : '', lookedUp: !!info
    });
  });
  songs.sort(function (a, b) {
    return ciCompare_(a.lead, b.lead) || ciCompare_(a.title, b.title) || ciCompare_(a.round, b.round);
  });

  var roundList = Array.from(rounds.values()).sort(function (a, b) {
    return a.Created < b.Created ? -1 : a.Created > b.Created ? 1 : 0;
  });

  // Finishing place of every song in its round. Ties on points are broken the
  // way Music League does: more unique voters giving it points, then the
  // highest single vote. Songs still tied after that share the place (1, 1, 3).
  // Rounds with no votes yet (still in progress) are skipped.
  function beats(o, s) {
    if (o.points !== s.points) return o.points > s.points;
    if (o.voters.up.size !== s.voters.up.size) return o.voters.up.size > s.voters.up.size;
    return o.voters.maxVote > s.voters.maxVote;
  }
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
      placeOf.set(s, 1 + list.filter(function (o) { return beats(o, s); }).length);
    });
  });

  // Per-competitor totals.
  var comp = new Map();
  function newComp(id) {
    return { name: nameOf(id), points: 0, songs: 0, wins: 0, top3: 0, voters: new Set() };
  }
  competitors.forEach(function (_, id) { comp.set(id, newComp(id)); });
  songs.forEach(function (s) {
    if (!comp.has(s.submitterId)) comp.set(s.submitterId, newComp(s.submitterId));
    var c = comp.get(s.submitterId), place = placeOf.get(s);
    c.points += s.points;
    c.songs++;
    if (place === 1) c.wins++;
    if (place <= 3) c.top3++;
    s.voters.up.forEach(function (v) { c.voters.add(v); });
  });
  // Standings tie-breaker: most unique voters who gave them points across all rounds.
  var points = Array.from(comp.values()).sort(function (a, b) {
    return b.points - a.points || b.voters.size - a.voters.size || ciCompare_(a.name, b.name);
  }).map(function (c) {
    return [c.name, c.points, c.songs, c.songs ? Math.round(c.points / c.songs * 100) / 100 : '', c.wins, c.top3];
  });

  // Running total of songs per artist, using the same lead artist as the Songs
  // tab. Spelling variants ("Belle and Sebastian" / "Belle & Sebastian") are
  // grouped, showing the most common spelling. Every submission counts.
  var byArtist = new Map();
  songs.forEach(function (s) {
    var key = matchKey_(s.lead);
    var round = rounds.get(s.roundId), created = round ? round.Created : '';
    var a = byArtist.get(key) || { spellings: new Map(), songs: 0, lastCreated: '', lastRound: '' };
    a.spellings.set(s.lead, (a.spellings.get(s.lead) || 0) + 1);
    a.songs++;
    if (created >= a.lastCreated) { a.lastCreated = created; a.lastRound = s.round; }
    byArtist.set(key, a);
  });
  var artistTotals = Array.from(byArtist.values()).map(function (a) {
    var name = '', best = 0;
    a.spellings.forEach(function (n, spelling) { if (n > best) { best = n; name = spelling; } });
    return [name, a.songs, a.lastRound];
  }).sort(function (a, b) { return b[1] - a[1] || ciCompare_(a[0], b[0]); });

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

  var taste = buildTaste_(songs, byGenre, byDecade, [UNKNOWN, NOT_YET]);

  return {
    taste: taste,
    songs: songs.map(function (s) { return [s.lead, s.title, s.album, s.competitor, s.points, s.round, s.genre, s.year]; }),
    byGenre: byGenre,
    byDecade: byDecade,
    topSongs: topSongs,
    fans: fans,
    artistTotals: artistTotals,
    songsMissingInfo: songsMissingInfo,
    rounds: roundList.map(function (r) { return [r.Name, r.Description, r['Playlist URL'], r.League]; }),
    points: points,
    stats: [
      ['Total Unique Artists', artists.size],
      ['Total Songs', uniqueSongs.size],
      ['Total Submissions', songs.length],
      ['Total Rounds', roundList.length],
      ['Total Competitors', competitors.size],
      ['Total Genres', new Set(songs.map(function (s) { return s.genre; }).filter(Boolean)).size]
    ],
    raw: {
      competitors: Array.from(competitors.values()),
      rounds: Array.from(rounds.values()),
      submissions: Array.from(submissions.values()),
      votes: Array.from(votes.values())
    }
  };
}

var TASTE_TOP_GENRES = 12; // league-wide genres shown as columns on the Taste tab

/**
 * Pivot of each competitor's songs by decade and by the league's most common
 * genres (as a share of that competitor's songs), plus their own top genres
 * and year range. Competitors are listed alphabetically.
 */
function buildTaste_(songs, byGenre, byDecade, unknownLabels) {
  function known(label) { return unknownLabels.indexOf(label) < 0; }
  var decades = byDecade.map(function (d) { return d[0]; }).filter(known);
  var genres = byGenre.map(function (g) { return g[0]; }).filter(known).slice(0, TASTE_TOP_GENRES);

  var people = new Map();
  songs.forEach(function (s) {
    var p = people.get(s.submitterId) || { name: s.competitor, songs: 0, years: [], genres: new Map(), decades: new Map() };
    p.songs++;
    if (s.year) {
      p.years.push(s.year);
      var d = Math.floor(s.year / 10) * 10 + 's';
      p.decades.set(d, (p.decades.get(d) || 0) + 1);
    }
    if (s.genre) p.genres.set(s.genre, (p.genres.get(s.genre) || 0) + 1);
    people.set(s.submitterId, p);
  });

  var rows = Array.from(people.values()).sort(function (a, b) { return ciCompare_(a.name, b.name); }).map(function (p) {
    var top = Array.from(p.genres.entries()).sort(function (a, b) { return b[1] - a[1] || ciCompare_(a[0], b[0]); })
      .slice(0, 3).map(function (e) { return e[0] + ' (' + e[1] + ')'; }).join(', ');
    var avg = p.years.length ? Math.round(p.years.reduce(function (a, b) { return a + b; }, 0) / p.years.length) : '';
    return [p.name, p.songs, avg,
      p.years.length ? Math.min.apply(null, p.years) : '', p.years.length ? Math.max.apply(null, p.years) : '', top]
      .concat(decades.map(function (d) { return (p.decades.get(d) || 0) / p.songs; }))
      .concat(genres.map(function (g) { return (p.genres.get(g) || 0) / p.songs; }));
  });

  return {
    summaryHeaders: ['Competitor', 'Songs', 'Avg Year', 'Oldest', 'Newest', 'Top Genres'],
    decades: decades,
    genres: genres,
    rows: rows
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

/**
 * League name from a zip's file name: the whole name without ".zip", e.g.
 * "MFFL VIII.zip" -> "MFFL VIII". A browser's duplicate-download suffix like
 * " (1)" is dropped. An empty name falls back to DEFAULT_LEAGUE_NAME.
 */
function leagueFromFileName_(fileName) {
  var base = String(fileName).replace(/^.*[\\\/]/, '').replace(/\.zip$/i, '').replace(/\s*\(\d+\)$/, '');
  return base.trim() || DEFAULT_LEAGUE_NAME;
}

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
    ['Artist', 'Title', 'Album', 'Competitors Name', 'Points Assigned', 'Rounds Name', 'Genre', 'Year'],
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

  writeTable_(ss, 'Artists', 'Songs by Artist',
    ['Artist', 'Songs', 'Last Round Used'], [260, 70, 320], result.artistTotals);

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

  writeTaste_(ss, result.taste);

  // Put the tabs first, in order.
  ['Songs', 'Rounds', 'Points', 'Artists', 'Stats', 'Taste'].forEach(function (name, i) {
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

/**
 * Taste tab: one row per competitor. A summary section, then the share of
 * their songs by decade and by the league's top genres, each with a color scale.
 */
function writeTaste_(ss, taste) {
  var sh = ss.getSheetByName('Taste') || ss.insertSheet('Taste');
  sh.setFrozenColumns(0); // merging titles across a frozen column fails on rebuilds
  sh.clear();
  sh.setConditionalFormatRules([]);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  var nSummary = taste.summaryHeaders.length, nDec = taste.decades.length, nGen = taste.genres.length;
  var needCols = nSummary + nDec + nGen, needRows = taste.rows.length + 2;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());

  function slice(from, n) { return taste.rows.map(function (r) { return r.slice(from, from + n); }); }
  var pct = '0%;-0%;'; // blank instead of 0%
  writeBlock_(sh, 1, 'Competitor Taste', taste.summaryHeaders, slice(0, nSummary));
  // The name column is frozen, and Sheets can't freeze part of a merged cell, so
  // un-merge this title; left-aligned, it still spills across the empty cells.
  sh.getRange(1, 1, 1, nSummary).breakApart();
  sh.getRange(1, 1).setHorizontalAlignment('left');
  var rules = [];
  [[nSummary + 1, '% of Their Songs by Decade', taste.decades],
   [nSummary + nDec + 1, '% of Their Songs in the League\'s Top Genres', taste.genres]].forEach(function (b) {
    if (!b[2].length) return;
    writeBlock_(sh, b[0], b[1], b[2], slice(b[0] - 1, b[2].length), b[2].map(function () { return pct; }));
    if (taste.rows.length) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpoint('#ffffff').setGradientMaxpoint('#57bb8a')
        .setRanges([sh.getRange(3, b[0], taste.rows.length, b[2].length)]).build());
    }
  });
  sh.setConditionalFormatRules(rules);

  [150, 60, 75, 65, 65, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  for (var c = nSummary + 1; c <= needCols; c++) sh.setColumnWidth(c, 80);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(1);
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

// ---- League name dialog ------------------------------------------------------

var LEAGUE_HTML = [
  '<style>body{font-family:Arial,sans-serif;font-size:14px}label{display:block;margin-top:10px}',
  '#name{margin:6px 0 0 24px;width:300px;padding:4px}#msg{margin-top:12px;white-space:pre-wrap}',
  'button{margin-top:14px;padding:6px 14px}</style>',
  '<div>Which league name should uploaded zips be labelled with on the Rounds tab?',
  ' This applies to zips you upload from now on.</div>',
  '<label><input type="radio" name="mode" value="file" id="file"> Use the zip file name(s)',
  ' <span style="color:#666">(e.g. "MFFL VIII.zip" → MFFL VIII)</span></label>',
  '<label><input type="radio" name="mode" value="custom" id="custom"> Use this name:</label>',
  '<input type="text" id="name" placeholder="League name">',
  '<br><button id="save" onclick="save()">Save</button>',
  '<div id="msg"></div>',
  '<script>',
  'var mode=__MODE__,name=__NAME__;',
  'document.getElementById(mode).checked=true;document.getElementById("name").value=name;',
  'document.getElementById("name").oninput=function(){document.getElementById("custom").checked=true;};',
  'function save(){var m=document.getElementById("msg"),b=document.getElementById("save");',
  ' var chosen=document.getElementById("custom").checked?"custom":"file";b.disabled=true;',
  ' google.script.run.withSuccessHandler(function(s){m.textContent=s;',
  '   setTimeout(function(){google.script.host.close();},2000);})',
  '  .withFailureHandler(function(e){m.textContent=e.message;b.disabled=false;})',
  '  .saveLeagueSetting(chosen,document.getElementById("name").value);}',
  '</script>'
].join('\n');

// Allow the core logic to be tested in Node.
if (typeof module !== 'undefined') {
  module.exports = {
    buildLeagueTables: buildLeagueTables, parseCsvObjects: parseCsvObjects, naturalCompare_: naturalCompare_,
    lookupSongBatch_: lookupSongBatch_, baseTitle_: baseTitle_, matchKey_: matchKey_, isrcYear_: isrcYear_,
    deezerSearchUrl_: deezerSearchUrl_, pickDeezerTrack_: pickDeezerTrack_, leagueFromFileName_: leagueFromFileName_,
    musicBrainzUrl_: musicBrainzUrl_, fetchMusicBrainz_: fetchMusicBrainz_, musicBrainzGenre_: musicBrainzGenre_,
    titleCaseGenre_: titleCaseGenre_, LEAGUE_HTML: LEAGUE_HTML, leagueForUpload_: leagueForUpload_
  };
}
