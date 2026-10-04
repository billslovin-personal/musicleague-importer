# Music League Excel File

A Node.js script that turns Music League export zip files into an Excel
workbook, de-duplicating everything across seasons. It's the simpler, offline
counterpart to the [Google Sheet version](../google-sheet/): no genre, year or
artist lookups, and three tabs instead of six.

## Requirements

[Node.js](https://nodejs.org) 18 or newer. Install the script's packages once,
from the repository folder:

```bash
npm --prefix excel install
```

## Usage

1. Put the league's export zip files in one folder, and rename each to its
   league's name, e.g. `MFFL VIII.zip`. Exports don't include the league's
   name, so the file name (without ".zip") is used for the League column.
2. From that folder, run:

```bash
node path/to/excel/build-league-sheet.js
```

It reads every `.zip` in the folder and writes `Music League.xlsx` there. To
write somewhere else, pass the output path as an argument. From the repository
folder, `npm --prefix excel run build` does the same.

## The tabs

- **Songs:** every submission with its artist credit (as listed in the
  export), title, album, competitor, points and round, sorted by artist, plus
  league totals (unique artists, songs, submissions, rounds and competitors).
- **Rounds:** every round with its description, playlist link and league name,
  oldest first.
- **Points:** lifetime points per competitor, highest first.

## How the data is combined

Each export holds `competitors.csv`, `rounds.csv`, `submissions.csv` and
`votes.csv`. Records from all the zips are merged and de-duplicated, so the
zips can be in any order and including one twice is safe:

| Record | Matched by |
|---|---|
| Competitor | ID. The name comes from the export with the newest rounds. |
| Round | Round ID |
| Song submission | Round ID + Spotify URI |
| Vote | Round ID + Spotify URI + voter ID |

A song's points are the votes it got in the round it was submitted to. A
competitor's lifetime points are the total across all their songs. The script
prints how many rows were read and how many remained after de-duplication.
