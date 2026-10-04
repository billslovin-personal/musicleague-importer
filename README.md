# Music League Importer

Turns the zip exports from [Music League](https://musicleague.com) into a
league spreadsheet that de-duplicates everything and keeps lifetime stats
across seasons. There are two versions, each with its own instructions:

| | [Google Sheet](google-sheet/) (recommended) | [Excel file](excel/) |
|---|---|---|
| Instructions | [`google-sheet/README.md`](google-sheet/README.md) | [`excel/README.md`](excel/README.md) |
| Who runs it | Anyone with the sheet open, from a menu | Someone with Node.js installed |
| Sharing | Share the sheet link with the league | Send the `.xlsx` file around |
| Tabs | Songs, Rounds, Points, Artists, Stats, Taste | Songs, Rounds, Points |
| Genre, year & lead artist lookup | Yes (MusicBrainz and Deezer) | No |

## Files

| Path | What it is |
|---|---|
| [`google-sheet/MusicLeague.gs`](google-sheet/MusicLeague.gs) | The Google Apps Script, pasted into the sheet's script editor. |
| [`google-sheet/song-info.csv`](google-sheet/song-info.csv) | Genre, year and lead artist already looked up for every song in the existing exports, uploaded on the first import. |
| [`excel/build-league-sheet.js`](excel/build-league-sheet.js) | The Node.js script that writes the Excel file (with its `package.json`). |

## Status

The data processing, lookups and stats are tested locally against eight
seasons of exports, and the Google Sheet has been run in Google Sheets. The
genre/year lookup running from Google's servers is the least-tested part:
MusicBrainz sometimes limits requests from Google, so lookups may need
re-running.

## Privacy

Export files and generated spreadsheets contain league members' names, IDs and
comments. `.gitignore` keeps `*.zip` and `*.xlsx` files out of this repository;
don't commit them. `song-info.csv` holds only song IDs, artists, titles,
genres, years and genre sources.
