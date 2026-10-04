# Music League Importer

Turns the zip exports from [Music League](https://musicleague.com) into a
league spreadsheet that de-duplicates everything and keeps lifetime stats
across seasons.

There are two ways to use it:

| | Google Sheet (recommended) | Local Excel file |
|---|---|---|
| Where | [`google-sheet/`](google-sheet/) | [`build-league-sheet.js`](build-league-sheet.js) |
| Who runs it | Anyone with the sheet open, from a menu | Someone with Node.js installed |
| Sharing | Share the sheet link with the league | Send the `.xlsx` file around |
| Tabs | Songs, Rounds, Points, Artists, Stats | Songs, Rounds, Points |
| Genre, year & lead artist lookup | Yes | No |

## Files

| Path | What it is |
|---|---|
| [`google-sheet/MusicLeague.gs`](google-sheet/MusicLeague.gs) | The Google Apps Script. Paste it into the sheet's script editor. |
| [`google-sheet/SETUP.md`](google-sheet/SETUP.md) | Step-by-step setup and weekly instructions for the Google Sheet. |
| [`google-sheet/song-info.csv`](google-sheet/song-info.csv) | Genre, year and lead artist already looked up for every song in the existing exports. Upload it on the first import to skip about an hour of lookups. |
| [`build-league-sheet.js`](build-league-sheet.js) | The local Node.js script that writes an Excel file. |

## Google Sheet

A Google Apps Script that adds a **Music League** menu to a Google Sheet:

- **Upload zip files…:** import one or more Music League exports, plus
  optionally `song-info.csv`. Already imported data is skipped, so each week
  you only need to upload the new zip.
- **Look up genres & years:** fill in genre, year and lead artist for songs
  that don't have them yet. Each run stops after about 4½ minutes (Google's
  limit) and picks up where it left off next time.
- **Rebuild tabs:** regenerate the tabs from the stored data.
- **Start a new league (erase everything)…:** wipe the sheet for a brand-new
  league, after a confirmation.

**Setup:** follow [`google-sheet/SETUP.md`](google-sheet/SETUP.md). It takes
about 5 minutes and needs no installs.

**Each week:** download the new export from Music League, name it
`export-<League name>.zip`, upload it, then run **Look up genres & years**.

**Permissions:** the script uses `@OnlyCurrentDoc`, so it can only access its
own spreadsheet, not the rest of your Google Drive. The lookup sends each
song's artist and title (nothing else) to the public Deezer and MusicBrainz
APIs.

### What's in the tabs

- **Songs:** every submission with its lead artist, title, album, competitor,
  points, round, genre and year, plus league totals (unique artist credits,
  songs, submissions, rounds and competitors).
- **Rounds:** every round with its description, playlist link and league name.
- **Points:** lifetime points per competitor, songs submitted, average points
  per song, round wins and top-3 finishes.
- **Artists:** running total of songs submitted per lead artist (the same
  artist shown on the Songs tab), with the last round each was used in. Useful
  for rounds with "artist used only once" rules.
- **Stats:** all-time top 25 songs, each competitor's biggest fan (who has
  given them the most points), and the share of songs and average points by
  genre and by decade recorded.

## Local Excel file

Requires [Node.js](https://nodejs.org) 18 or newer.

```bash
npm install
```

Put the `export-*.zip` files in this folder, then:

```bash
node build-league-sheet.js
```

This writes `Music League.xlsx` with the Songs, Rounds and Points tabs. You can
pass a different output path as an argument. It doesn't do the genre/year/lead
artist lookup, so its Songs tab shows the full artist credit from the export.

## How the data is combined

Each Music League export holds `competitors.csv`, `rounds.csv`,
`submissions.csv` and `votes.csv`. Records from all exports are merged and
de-duplicated, so zips can be uploaded in any order and uploading one twice is
safe:

| Record | Matched by |
|---|---|
| Competitor | ID. The name comes from the export with the newest rounds, so upload order doesn't matter. |
| Round | Round ID |
| Song submission | Round ID + Spotify URI |
| Vote | Round ID + Spotify URI + voter ID |

**League name:** exports don't include it, so it comes from the zip's file
name, everything after the first "-": `export-MFFL VIII.zip` → **MFFL VIII**.
A file name without a "-" gives **Music League**. Each zip's rounds keep the
name of the file they came from; rename and re-upload a zip to change it.

**Points:** a song's points are the votes it got in the round it was submitted
to. A competitor's lifetime points are the total across all their songs.

**Ties** follow Music League's tie-breakers:
- *Round placings* (wins, top-3): points, then the most unique voters who gave
  the song points, then the highest single vote. Songs still tied share the
  place, and each counts as a win.
- *Points tab order:* total points, then the most unique voters who have given
  the competitor points across all rounds.

## Genre, year and lead artist (Google Sheet)

| Field | Source | Notes |
|---|---|---|
| Lead artist | Deezer's main artist for the matched track | Falls back to the first artist in the export (split on ", ") when Deezer doesn't have the song. |
| Genre | Deezer, the first genre listed for the track's album | Broad labels (Rock, Pop, Alternative, …). About 9% of current songs are *Unknown*, mostly because Deezer's catalog doesn't carry them. |
| Year | Earliest of MusicBrainz's first release, the recording's ISRC year code and the Deezer album date | Remasters and compilations still show the original year. Right for the large majority of songs, not all. |

Matching is strict: a Deezer result is used only when both the artist and the
title match, so covers, karaoke versions and tribute acts are skipped.

## Status

The data processing, lookups and stats are tested locally against eight
seasons of exports. The Google Sheet parts (menus, upload dialog, formatting,
and the lookup running from Google's servers) haven't been run inside Google
Sheets yet. In particular, MusicBrainz sometimes limits requests from Google's
servers; if years stop appearing for new songs, that's the likely cause.

## Privacy

Export files and generated spreadsheets contain league members' names, IDs and
comments. `.gitignore` keeps `*.zip` and `*.xlsx` files out of this repository.
Don't commit them. `song-info.csv` holds only song IDs, artists, titles,
genres and years.
