const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'assets', 'xlsx.full.min.js'));
const Core = require(path.join(__dirname, '..', 'assets', 'core.js'));

const file = process.argv[2];
let wb;
try {
  const buf = fs.readFileSync(file);
  wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
} catch (e) {
  console.log('READ ERROR:', e.message);
  process.exit(2);
}
console.log('sheets:', wb.SheetNames.join(','));
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: null });
console.log('rows:', rows.length);
console.log('header:', (rows[0] || []).slice(0, 8).join(' | '));

const ds = Core.buildDataset([{ name: wb.SheetNames[0], rows }]);
if (!ds.ok) { console.log('parse fail', ds.error); process.exit(1); }
console.log('students:', ds.students.length, 'absent:', ds.absent.length);
console.log('warnings:', ds.warnings.slice(0, 5));

Core.ensureTotals(ds.students);
Core.computeRanks(ds.students);

let checked = 0, mismatchG = 0, mismatchC = 0, nullSrc = 0;
const samples = [];
for (const s of ds.students) {
  for (const item of Core.ITEMS) {
    const src = s.srcRanks[item];
    if (!src || src.grade === null) { nullSrc++; continue; }
    const cur = s.ranks[item];
    checked++;
    if (cur.grade !== src.grade) { mismatchG++; if (samples.length < 6) samples.push([s.sid, item, 'grade', cur.grade, src.grade]); }
    if (src.cls !== null && cur.cls !== src.cls) { mismatchC++; if (samples.length < 6) samples.push([s.sid, item, 'cls', cur.cls, src.cls]); }
  }
}
console.log(JSON.stringify({ checked, mismatchG, mismatchC, nullSrc }));
console.log('samples:', JSON.stringify(samples));

const s0 = ds.students[0];
console.log('first:', s0.sid, s0.name, s0.clsLabel, 'total', s0.scores['总分'], 'rank', JSON.stringify(s0.ranks['总分']), '3in1', s0.scores['语数外总分']);
console.log('null total:', ds.students.filter(s => s.scores['总分'] === null).length);
