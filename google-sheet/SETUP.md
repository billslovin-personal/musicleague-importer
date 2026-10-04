# Music League Google Sheet — Setup

This turns a Google Sheet into the league's master spreadsheet with five tabs:

- **Songs:** every submission with its lead artist, points, round, genre and
  year.
- **Rounds:** every round with its description and playlist link.
- **Points:** lifetime points per competitor, plus songs submitted, average
  points per song, round wins and top-3 finishes.
- **Artists:** running total of songs submitted per artist, with the last round
  each artist was used in. Handy for rounds with "artist used only once" rules.
  It uses the same lead artist as the Songs tab, grouping spelling variants
  like "Belle and Sebastian" / "Belle & Sebastian".
- **Stats:** the all-time top 25 songs, each competitor's biggest fan (who has
  given them the most points), and the share of songs and average points by
  genre and by decade recorded.

You upload the Music League export zip files from a menu inside the sheet;
everything else is automatic.

**Privacy:** the script can only see and change *this one spreadsheet*. It has
no access to anything else in your Google Drive, Gmail, etc. To find genres
and years, it sends each song's artist and title (nothing else) to two
free public music databases, Deezer and MusicBrainz.

## One-time setup (about 5 minutes)

1. Go to [sheets.google.com](https://sheets.google.com) and create a **Blank spreadsheet**.
   Name it something like "MFFL Music League".
2. In the sheet, click **Extensions → Apps Script**. A code editor opens in a new tab.
3. Delete everything in the editor, then paste in the entire contents of
   `MusicLeague.gs`.
4. Click the **Save** icon (💾). Close the Apps Script tab.
5. Reload the spreadsheet tab. After a few seconds a **Music League** menu
   appears next to **Help**.

## Importing data

1. Click **Music League → Upload zip files…**
2. The first time, Google asks for permission:
   - Click **Continue** and choose your Google account.
   - You may see **"Google hasn't verified this app"**. This appears because
     you wrote the script yourself and didn't publish it. Click **Advanced → Go to
     (project name) (unsafe)**.
   - The permissions listed should be only *"View and manage spreadsheets that
     this application has been installed in"*, *"Display and run third-party
     web content in prompts and sidebars"*, and *"Connect to an external
     service"* (used only for the genre/year lookup). Click **Allow**.
   - Then click **Music League → Upload zip files…** again.
3. Pick all the `export-*.zip` files **and** `song-info.csv` (you can select
   several at once) and click **Import**. Larger uploads can take up to a
   minute. `song-info.csv` holds genres and years already looked up for all
   the existing songs, which saves about an hour of lookups.

After the first import you can delete the empty **Sheet1** tab.

### Each new week / season
1. Download the new export zip from Music League and upload it the same way.
   Uploading only the new zip is enough because earlier data is kept.
   Uploading a zip again is safe, since duplicates are removed automatically.
2. Click **Music League → Look up genres & years** to fill in the new songs.
   A week's songs take under a minute.

### About artists, genres and years
- **Artist** on the Songs tab is the song's lead artist, spelled as Deezer
  lists it. Until a song has been looked up, or if Deezer doesn't have it, the
  first artist named in the Music League export is shown instead. The "Total
  Unique Artists" stat still counts each distinct artist credit from the export.
- **Genre** comes from Deezer and is a broad label (Rock, Pop, Alternative,
  Rap/Hip Hop, …). About 1 in 8 songs has no genre there and shows as *Unknown*.
- **Year** is the earliest release of that recording found in MusicBrainz or
  Deezer, so remasters and compilations still show the original year. It's
  right for the large majority of songs but not all of them.
- A lookup run stops itself after about 4½ minutes (Google's time limit). If
  there are more songs left, it says so — just run it again to continue.

## Other menu items

- **Set league name…:** the name shown in the League column of the Rounds
  tab. It defaults to "MFFL VIII".
- **Rebuild tabs:** regenerates Songs/Rounds/Points from the stored data.
- **Start a new league (erase everything)…:** for when a brand-new league
  begins. After you confirm, it permanently deletes all songs, rounds, votes,
  points, stats, looked-up genres/years and the league name, then asks for the
  new league's name. It only touches the tabs this script creates; any tabs you
  added yourself are left alone. Consider **File → Make a copy** first to keep
  the old league's sheet.

## Sharing with the league

Click **Share** (top right) → under *General access* choose **Anyone with the
link** → role **Viewer** → **Copy link**, and send that link to everyone.
Viewers can see and filter the tabs but can't change data or run the import.

## How the data works

- The raw data from every upload is stored in hidden tabs (`raw_competitors`,
  `raw_rounds`, `raw_submissions`, `raw_votes`, `settings`). Don't edit these.
- De-duplication: competitors by ID, rounds by ID, songs by song + round, and
  votes by song + round + voter. Zips can be uploaded in any order. If someone
  changed their display name, the name from the most recent season is used.
- A song's points are the votes it got in the round it was submitted to.
  A competitor's lifetime points are the total across all their songs.
- Round wins / top-3 finishes use Music League's round tie-breakers: songs
  tied on points are ranked by the most unique voters who gave them points,
  then by the highest single vote. Songs still tied after that share the
  place, and each counts as a win. Rounds that don't have any votes yet are
  skipped.
- The Points tab uses Music League's standings tie-breaker: competitors tied
  on total points are ranked by the most unique voters who have given them
  points across all rounds.
