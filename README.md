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
| Tabs | Songs, Rounds, Points, Stats | Songs, Rounds, Points |
| Genre & year lookup | Yes | No |

## Google Sheet

A Google Apps Script that adds a **Music League** menu to a Google Sheet:

- **Upload zip files…:** import one or more Music League exports. Already
  imported data is skipped, so each week you only need to upload the new zip.
- **Look up genres & years:** fill in genre and year for new songs.
- **Rebuild tabs:** regenerate the tabs from the stored data.
- **Set league name…:** the name shown on the Rounds tab.
- **Start a new league (erase everything)…:** wipe the sheet for a brand-new
  league, after a confirmation.

The script uses `@OnlyCurrentDoc`, so it can only access its own spreadsheet,
not the rest of your Google Drive. The genre/year lookup sends each song's
artist and title to the public Deezer and MusicBrainz APIs.

**Setup:** follow [`google-sheet/SETUP.md`](google-sheet/SETUP.md). It takes
about 5 minutes and needs no installs.

### What's in the tabs

- **Songs:** every submission with its competitor, points, round, genre and year,
  plus league totals (unique artists, songs, submissions, rounds, competitors).
- **Rounds:** every round with its description, playlist link and league name.
- **Points:** lifetime points per competitor, songs submitted, average points per
  song, round wins and top-3 finishes.
- **Stats:** all-time top 25 songs, each competitor's biggest fan, and the share
  of songs and average points by genre and by decade recorded.

## Local Excel file

Requires [Node.js](https://nodejs.org) 18 or newer.

```bash
npm install
```

Put the `export-*.zip` files in this folder, then:

```bash
node build-league-sheet.js
```

This writes `MFFL VIII Music League.xlsx`. You can pass a different output path
as an argument. The league name is set by `LEAGUE_NAME` at the top of the script.

## How the data is combined

Each Music League export holds `competitors.csv`, `rounds.csv`,
`submissions.csv` and `votes.csv`. Records from all exports are merged and
de-duplicated:

| Record | Matched by |
|---|---|
| Competitor | ID. The name comes from the export with the newest rounds, so upload order doesn't matter. |
| Round | Round ID |
| Song submission | Round ID + Spotify URI |
| Vote | Round ID + Spotify URI + voter ID |

A song's points are the votes it got in the round it was submitted to. Exports
don't include the league name, so it's configured separately.

## Privacy

Export files and generated spreadsheets contain league members' names, IDs and
comments. `.gitignore` keeps `*.zip` and `*.xlsx` files out of this repository.
Don't commit them.
