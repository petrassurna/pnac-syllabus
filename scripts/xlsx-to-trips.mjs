/**
 * Extract the syllabus from Syllabus.xlsx into trips.json (the app's data).
 *
 * Spreadsheet columns (no header row):
 *   A = Date        real date for a single day, OR a text range e.g. "27-28 July 2026"
 *                   or "30 November - 1 December 2026"
 *   B = Type        Saltwater | Freshwater | Surf/Estuary | Saltwater & Surf/Estuary,
 *                   or an event name (Working Bee, AGM, Presentation Night, ...)
 *   C = Location    place / venue
 *
 * Run:  npm run data
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'Syllabus.xlsx');
const OUT = path.join(ROOT, 'trips.json');

const MONTHS = ['january','february','march','april','may','june','july',
                'august','september','october','november','december'];
// Types that mean "fishing trip"; anything else is treated as a club event.
const TRIP_MARKERS = ['saltwater', 'freshwater', 'surf/estuary'];

const pad = (n) => String(n).padStart(2, '0');
const iso = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
const monthIdx = (name) => MONTHS.indexOf(String(name).toLowerCase());

function fromDisplay(text) {
  // "D Month - D Month YYYY"  (range spanning two months)
  let m = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s*-\s*\d{1,2}\s+[A-Za-z]+\s+(\d{4})$/);
  if (m) return { start: iso(+m[3], monthIdx(m[2]) + 1, +m[1]), display: text };
  // "D-D Month YYYY"  (range within one month)
  m = text.match(/^(\d{1,2})\s*-\s*\d{1,2}\s+([A-Za-z]+)\s+(\d{4})$/);
  if (m) return { start: iso(+m[3], monthIdx(m[2]) + 1, +m[1]), display: text };
  // "D Month YYYY"  (single day written as text)
  m = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (m) return { start: iso(+m[3], monthIdx(m[2]) + 1, +m[1]), display: text };
  throw new Error(`Unrecognised date text: "${text}"`);
}

function fromDate(dt) {
  const y = dt.getFullYear(), mo = dt.getMonth(), d = dt.getDate();
  return { start: iso(y, mo + 1, d), display: `${d} ${cap(MONTHS[mo])} ${y}` };
}
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

function kindOf(type) {
  const t = String(type).toLowerCase();
  return TRIP_MARKERS.some((m) => t.includes(m)) ? 'trip' : 'event';
}

const wb = XLSX.read(fs.readFileSync(SRC), { type: 'buffer', cellDates: true });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });

const trips = [];
for (const row of rows) {
  const [a, type, location] = row;
  if (a == null || type == null) continue;
  const { start, display } = (a instanceof Date) ? fromDate(a) : fromDisplay(String(a).trim());
  trips.push({ start, display, type: String(type).trim(), location: String(location ?? '').trim(), kind: kindOf(type) });
}

trips.sort((x, y) => (x.start < y.start ? -1 : x.start > y.start ? 1 : x.location.localeCompare(y.location)));

fs.writeFileSync(OUT, JSON.stringify(trips, null, 2) + '\n', 'utf8');
console.log(`Wrote ${trips.length} entries to ${path.relative(ROOT, OUT)}`);
