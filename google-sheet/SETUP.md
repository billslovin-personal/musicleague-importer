# Music League Google Sheet — Setup

This turns a Google Sheet into the league's master spreadsheet with three tabs:
**Songs**, **Rounds**, and **Points**. You upload the Music League export zip
files from a menu inside the sheet; everything else is automatic.

**Privacy:** the script can only see and change *this one spreadsheet*. It has
no access to anything else in your Google Drive, Gmail, etc.

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
     this application has been installed in"* and *"Display and run
     third-party web content in prompts and sidebars"*. Click **Allow**.
   - Then click **Music League → Upload zip files…** again.
3. Pick one or more `export-*.zip` files (you can select several at once) and
   click **Import**. Larger uploads can take up to a minute.

After the first import you can delete the empty **Sheet1** tab.

### Each new week / season
Download the new export zip from Music League and upload it the same way.
Uploading only the new zip is enough because earlier data is kept. Uploading
a zip again is safe, since duplicates are removed automatically.

## Other menu items

- **Set league name…:** the name shown in the League column of the Rounds
  tab. It defaults to "MFFL VIII".
- **Rebuild tabs:** regenerates Songs/Rounds/Points from the stored data.
- **Clear all stored data…:** wipes everything so you can start over.

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
