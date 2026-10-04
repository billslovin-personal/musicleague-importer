// Builds a Music League master spreadsheet (Songs, Rounds, Points) from the
// Music League export-*.zip files in this directory.
//
// Usage:  node build-league-sheet.js [outputFile]
//
// Each zip holds competitors.csv, rounds.csv, submissions.csv and votes.csv.
// Zips are processed in file-name order; when the same record appears in more
// than one zip it is kept once (later zips win for changed values like names).

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// ---- Configuration -------------------------------------------------------
const LEAGUE_NAME = 'MFFL VIII';           // Not present in the export files
const INPUT_DIR = __dirname;
const ZIP_PATTERN = /^export-.*\.zip$/i;
const OUTPUT_FILE = process.argv[2] || path.join(INPUT_DIR, `${LEAGUE_NAME} Music League.xlsx`);
// --------------------------------------------------------------------------

const clean = (v) => (v == null ? '' : String(v).replace(/^﻿/, '').trim());

async function readZip(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const tables = {};
  for (const name of ['competitors', 'rounds', 'submissions', 'votes']) {
    const entry = zip.file(new RegExp(`(^|/)${name}\\.csv$`, 'i'))[0];
    if (!entry) throw new Error(`${path.basename(file)} is missing ${name}.csv`);
    const text = await entry.async('string');
    tables[name] = parse(text, { columns: true, skip_empty_lines: true, bom: true, relax_quotes: true });
  }
  return tables;
}

async function main() {
  const zips = fs.readdirSync(INPUT_DIR)
    .filter((f) => ZIP_PATTERN.test(f))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  if (!zips.length) throw new Error(`No export zip files found in ${INPUT_DIR}`);

  // Keyed maps give us de-duplication; insertion order preserves file order.
  const competitors = new Map(); // ID -> name
  const nameAsOf = new Map();    // ID -> newest round date the name came from
  const rounds = new Map();      // Round ID -> round
  const submissions = new Map(); // Round ID|URI -> submission
  const votes = new Map();       // Round ID|URI|Voter ID -> vote
  const raw = { competitors: 0, rounds: 0, submissions: 0, votes: 0 };

  for (const z of zips) {
    const t = await readZip(path.join(INPUT_DIR, z));
    // An export's names are current as of its newest round; the newest name wins.
    const asOf = t.rounds.reduce((max, r) => (clean(r.Created) > max ? clean(r.Created) : max), '');
    for (const c of t.competitors) {
      raw.competitors++;
      const id = clean(c.ID);
      if ((nameAsOf.get(id) || '') > asOf) continue;
      competitors.set(id, clean(c.Name));
      nameAsOf.set(id, asOf);
    }
    for (const r of t.rounds) {
      raw.rounds++;
      const id = clean(r.ID);
      const existing = rounds.get(id);
      rounds.set(id, {
        id,
        created: clean(r.Created),
        name: clean(r.Name),
        description: clean(r.Description),
        playlist: clean(r['Playlist URL']),
        source: existing ? existing.source : z,
      });
    }
    for (const s of t.submissions) {
      raw.submissions++;
      const roundId = clean(s['Round ID']);
      const uri = clean(s['Spotify URI']);
      submissions.set(`${roundId}|${uri}`, {
        uri,
        roundId,
        title: clean(s.Title),
        album: clean(s.Album),
        artist: clean(s['Artist(s)']),
        submitterId: clean(s['Submitter ID']),
      });
    }
    for (const v of t.votes) {
      raw.votes++;
      const key = `${clean(v['Round ID'])}|${clean(v['Spotify URI'])}|${clean(v['Voter ID'])}`;
      votes.set(key, Number(clean(v['Points Assigned'])) || 0);
    }
    console.log(`Read ${z}: ${t.rounds.length} rounds, ${t.submissions.length} submissions, ${t.votes.length} votes`);
  }

  // Points per submission (song within a specific round).
  const pointsBySubmission = new Map();
  for (const [key, pts] of votes) {
    const subKey = key.slice(0, key.lastIndexOf('|'));
    pointsBySubmission.set(subKey, (pointsBySubmission.get(subKey) || 0) + pts);
  }

  const nameOf = (id) => competitors.get(id) || id;
  const cmp = (a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' });

  const songRows = [...submissions.entries()].map(([key, s]) => ({
    artist: s.artist,
    title: s.title,
    album: s.album,
    competitor: nameOf(s.submitterId),
    points: pointsBySubmission.get(key) || 0,
    round: rounds.get(s.roundId)?.name || '',
    submitterId: s.submitterId,
  }));
  songRows.sort((a, b) => cmp(a.artist, b.artist) || cmp(a.title, b.title) || cmp(a.round, b.round));

  const roundRows = [...rounds.values()].sort((a, b) => a.created.localeCompare(b.created));

  const totals = new Map(); // competitor ID -> points
  for (const id of competitors.keys()) totals.set(id, 0);
  for (const s of songRows) totals.set(s.submitterId, (totals.get(s.submitterId) || 0) + s.points);
  const pointRows = [...totals.entries()]
    .map(([id, pts]) => ({ name: nameOf(id), points: pts }))
    .sort((a, b) => b.points - a.points || cmp(a.name, b.name));

  // ---- Write workbook ----------------------------------------------------
  const wb = new ExcelJS.Workbook();
  const font = { name: 'Arial' };
  const titleStyle = {
    font: { ...font, bold: true, color: { argb: 'FFFF0000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
    alignment: { horizontal: 'center', vertical: 'top' },
  };
  const headerStyle = {
    font: { ...font, bold: true, color: { argb: 'FFFF0000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF434343' } },
    alignment: { vertical: 'top', wrapText: true },
  };
  const bodyAlign = { vertical: 'top', wrapText: true };

  function addSheet(name, title, headers, widths, rows) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 2 }] });
    ws.columns = widths.map((width) => ({ width }));
    ws.mergeCells(1, 1, 1, headers.length);
    Object.assign(ws.getCell(1, 1), { value: title });
    for (let c = 1; c <= headers.length; c++) ws.getCell(1, c).style = titleStyle;
    headers.forEach((h, i) => { const cell = ws.getCell(2, i + 1); cell.value = h; cell.style = headerStyle; });
    rows.forEach((r, i) => {
      const row = ws.getRow(i + 3);
      row.values = r;
      row.eachCell((cell) => { cell.font = font; cell.alignment = bodyAlign; });
    });
    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2 + rows.length, column: headers.length } };
    return ws;
  }

  const songs = addSheet('Songs', 'Artist / Song Master',
    ['Artist(s)', 'Title', 'Album', 'Competitors Name', 'Points Assigned', 'Rounds Name'],
    [26, 34, 37, 18, 12, 40],
    songRows.map((s) => [s.artist, s.title, s.album, s.competitor, s.points, s.round]));

  // League stats beside the song list, as in the example workbook.
  songs.getColumn(7).width = 4;
  songs.getColumn(8).width = 22;
  songs.getColumn(9).width = 10;
  songs.mergeCells('H1:I2');
  songs.getCell('H1').value = 'League Stats';
  songs.getCell('H1').style = { ...titleStyle, alignment: { horizontal: 'center', vertical: 'middle' } };
  const stats = [
    ['Total Unique Artists', new Set(songRows.map((s) => s.artist.toLowerCase())).size],
    ['Total Songs', new Set(songRows.map((s) => `${s.artist}|${s.title}`.toLowerCase())).size],
    ['Total Submissions', songRows.length],
    ['Total Rounds', roundRows.length],
    ['Total Competitors', competitors.size],
  ];
  stats.forEach(([label, value], i) => {
    songs.getCell(3 + i, 8).value = label;
    songs.getCell(3 + i, 8).font = { ...font, bold: true, color: { argb: 'FF0000FF' } };
    songs.getCell(3 + i, 9).value = value;
    songs.getCell(3 + i, 9).font = { ...font, bold: true };
  });

  const roundsWs = addSheet('Rounds', 'Rounds Master',
    ['Name', 'Description', 'Playlist URL', 'League'],
    [40, 70, 55, 14],
    roundRows.map((r) => [r.name, r.description,
      r.playlist ? { text: r.playlist, hyperlink: r.playlist } : '', LEAGUE_NAME]));
  roundsWs.getColumn(3).eachCell((cell, rowNum) => {
    if (rowNum > 2 && cell.value) cell.font = { ...font, color: { argb: 'FF1155CC' }, underline: true };
  });

  addSheet('Points', '"Lifetime" League Points',
    ['Competitor Name', 'Total Points'],
    [26, 14],
    pointRows.map((p) => [p.name, p.points]));

  await wb.xlsx.writeFile(OUTPUT_FILE);

  console.log('\nDe-duplication summary (raw rows -> unique):');
  console.log(`  competitors: ${raw.competitors} -> ${competitors.size}`);
  console.log(`  rounds:      ${raw.rounds} -> ${rounds.size}`);
  console.log(`  submissions: ${raw.submissions} -> ${submissions.size}`);
  console.log(`  votes:       ${raw.votes} -> ${votes.size}`);
  console.log(`\nWrote ${OUTPUT_FILE}`);
}

main().catch((err) => { console.error(err.message || err); process.exit(1); });
