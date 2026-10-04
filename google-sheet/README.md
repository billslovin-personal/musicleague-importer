# Music League Google Sheet

This turns a Google Sheet into the league's master spreadsheet. You upload the
Music League export files from a menu inside the sheet, and it builds six tabs
of songs, rounds, points and stats. Uploads are de-duplicated, so it keeps
lifetime totals across every season.

## What you need

- A Google account.
- These two files from this folder:
  - [`MusicLeague.gs`](MusicLeague.gs): the script.
  - [`song-info.csv`](song-info.csv): genres, years and artist names already
    looked up for every existing song, which saves a few hours of lookups.
- The league's export zip files from Music League (one per season).

## The tabs

- **Songs:** every submission with its lead artist, title, album, competitor,
  points, round, genre and year, plus league totals (artists, songs,
  submissions, rounds, competitors and genres).
- **Rounds:** every round with its description, playlist link and league name.
- **Points:** lifetime points per competitor, plus songs submitted, average
  points per song, round wins and top-3 finishes.
- **Artists:** running total of songs submitted per artist, with the last round
  each artist was used in. Handy for rounds with "artist used only once" rules.
  Spelling variants like "Belle and Sebastian" / "Belle & Sebastian" are
  grouped.
- **Stats:** the all-time top 25 songs, each competitor's biggest fan (who has
  given them the most points), and the share of songs and average points by
  genre and by decade recorded.
- **Taste:** one row per competitor with their songs, average/oldest/newest
  song year and top 3 genres, plus the share of their songs from each decade
  and in each of the league's 12 most common genres (color-scaled).

**Privacy:** the script can only see and change *this one spreadsheet*. It has
no access to anything else in your Google Drive, Gmail, etc. To find genres
and years, it sends each song's artist and title (nothing else) to two free
public music databases, Deezer and MusicBrainz.

## One-time setup (about 5 minutes)

1. Go to [sheets.google.com](https://sheets.google.com) and create a **Blank
   spreadsheet**. Name it something like "MFFL Music League".
2. In the sheet, click **Extensions → Apps Script**. A code editor opens in a
   new tab.
3. Delete everything in the editor, then paste in the entire contents of
   `MusicLeague.gs`.
4. Click the **Save** icon (💾). Close the Apps Script tab.
5. Reload the spreadsheet tab. After a few seconds a **Music League** menu
   appears next to **Help**.

## First import

1. Rename each export zip to its league's name, e.g. `MFFL VIII.zip` (see
   [League name](#league-name) below).
2. Click **Music League → Upload zip files…**
3. The first time, Google asks for permission:
   - Click **Continue** and choose your Google account.
   - You may see **"Google hasn't verified this app"**. This appears because
     the script is your own copy and hasn't been published. Click
     **Advanced → Go to (project name) (unsafe)**.
   - The permissions listed should be only *"View and manage spreadsheets that
     this application has been installed in"*, *"Display and run third-party
     web content in prompts and sidebars"*, and *"Connect to an external
     service"* (used only for the genre/year lookup). Click **Allow**.
   - Then click **Music League → Upload zip files…** again.
4. Pick all the export zip files **and** `song-info.csv` (you can select
   several at once) and click **Import**. Larger uploads can take up to a
   minute.

After the first import you can delete the empty **Sheet1** tab.

## Each new week or season

1. Download the new export zip from Music League, rename it to the league's
   name, and upload it with **Music League → Upload zip files…**. Uploading
   only the new zip is enough, because earlier data is kept. Uploading a zip
   again is safe, since duplicates are removed automatically.
2. Click **Music League → Look up genres & years** to fill in the new songs.
   A week's songs take a couple of minutes. A run stops itself after about
   4½ minutes (Google's time limit); if songs are left, it says so, and you
   just run it again to continue.

## League name

Music League's exports don't include the league's name. By default the zip's
file name is used: `MFFL VIII.zip` shows as **MFFL VIII** in the League column
of the Rounds tab.

To use a name of your own instead, open **Music League → League name…**,
choose **Use this name**, type it and click **Save**. The choice is saved with
the sheet and applies to every zip uploaded after that, until you change it
back to **Use the zip file name(s)**. Rounds already uploaded keep their name;
to relabel a season, change the setting and upload its zip again (nothing gets
duplicated).

## Other menu items

- **Rebuild tabs:** regenerates all the tabs from the stored data.
- **Start a new league (erase everything)…:** for when a brand-new league
  begins. After you confirm, it permanently deletes all songs, rounds, votes,
  points, stats and looked-up genres/years, and resets the league name setting.
  It only touches the tabs this script creates; any tabs you added yourself are
  left alone. Consider **File → Make a copy** first to keep the old league's
  sheet.

## Sharing with the league

Click **Share** (top right) → under *General access* choose **Anyone with the
link** → role **Viewer** → **Copy link**, and send that link to everyone.
Viewers can see and filter the tabs but can't change data or run the import.

## How the numbers work

- **Points:** a song's points are the votes it got in the round it was
  submitted to. A competitor's lifetime points are the total across all their
  songs.
- **Round wins and top-3 finishes** use Music League's round tie-breakers:
  songs tied on points are ranked by the most unique voters who gave them
  points, then by the highest single vote. Songs still tied after that share
  the place, and each counts as a win. Rounds without any votes yet are
  skipped.
- **Points tab order** uses Music League's standings tie-breaker: competitors
  tied on total points are ranked by the most unique voters who have given
  them points across all rounds.
- **Duplicates:** competitors are matched by ID, rounds by ID, songs by song +
  round, and votes by song + round + voter, so zips can be uploaded in any
  order. If someone changed their display name, the name from the most recent
  season is used.

## Artists, genres and years

- **Artist** on the Songs tab is the song's lead artist, spelled as Deezer
  lists it. Until a song has been looked up, or if Deezer doesn't have it, the
  first artist named in the Music League export is shown instead. The "Total
  Unique Artists" stat counts each distinct artist credit from the export.
- **Genre** comes from MusicBrainz, which has fine-grained genres (Indie Rock,
  Post-Punk, Doo-Wop, …): the top genre of the song's original recording, else
  of its album, else of its artist. When MusicBrainz has none, Deezer's
  broader genre is used, renamed to the matching MusicBrainz genre where they
  overlap (Deezer's "Alternative" shows as "Alternative Rock", "Electro" as
  "Electronic"). Songs neither source has show as *Unknown*.
- **Year** is the earliest release of that recording found in MusicBrainz or
  Deezer, so remasters and compilations still show the original year. It's
  right for the large majority of songs, but not all of them.
- Matching is strict: a result is used only when both the artist and the
  title match, so covers, karaoke versions and tribute acts are skipped.

## Behind the scenes

The raw data from every upload is stored in hidden tabs (`raw_competitors`,
`raw_rounds`, `raw_submissions`, `raw_votes` and `raw_song_info`). Don't edit
them; the visible tabs are rebuilt from them.

## If something goes wrong

- **An error message appears:** note what it says and when (which menu item),
  and pass it on. Re-pasting a newer `MusicLeague.gs` and running
  **Rebuild tabs** fixes most problems without re-uploading anything.
- **New songs stay "Not looked up yet" after a lookup:** MusicBrainz sometimes
  limits requests coming from Google's servers. Songs whose lookup failed are
  retried the next time you run **Look up genres & years**, so try again later.
