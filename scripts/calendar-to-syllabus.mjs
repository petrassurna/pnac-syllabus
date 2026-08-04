/**
 * Convert the club's calendar-style syllabus workbook (e.g. "Syllabus 2026-2027 ver4.xlsx")
 * into the flat 3-column Syllabus.xlsx that `npm run data` consumes.
 *
 * Source layout (first sheet), read top to bottom:
 *   - a month header row:            A = month name, B = year
 *   - one row per DAY of an outing:  A = day name, B = day number, C = description,
 *                                    D = donor, E = trophy value   (D/E are dropped)
 *   - blank rows separate outings
 *
 * Consecutive dated rows are folded into one entry; a gap in the dates starts a new
 * entry (the December block lists Christmas Tree 5-6 and the break-up on the 11th
 * without a blank row between them).
 *
 * Run:  npm run convert -- "Syllabus 2026-2027 ver4.xlsx"
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'Syllabus.xlsx');

const MONTHS = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];
const DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
// The workbook spells a few of these its own way.
const MONTH_ALIASES = { febuary: 'February', feburary: 'February', sept: 'September' };

/** Water-type prefixes used in column C. Order matters — longest first. */
const TYPE_PREFIXES = [
  [/^f\s*\/?\s*w\b[\s-]*/i, 'Freshwater'],
  [/^s\s*\/?\s*w\b[\s-]*/i, 'Saltwater'],
  [/^s\s*\/?\s*e\b[\s-]*/i, 'Surf/Estuary'],
];

/**
 * Club events (no water-type prefix). Keyed by the description squashed to
 * lowercase alphanumerics, so spacing/punctuation drift in the source is harmless.
 */
const EVENTS = {
  'agm':                                  ['AGM', 'Clubrooms'],
  'annualgeneralmeeting':                 ['AGM', 'Clubrooms'],
  'presentationnight':                    ['Presentation Night', 'Clubrooms'],
  'christmastree':                        ['Christmas Tree', 'Clubrooms'],
  'meetingclubbreakup':                   ['Club Break-up', 'Clubrooms'],
  'clubopens':                            ['Club Re-opens', 'Clubrooms'],
  'noclubmeeting':                        ['No Club Meeting', 'Clubrooms'],
  'bbqknockoutpoolcomp1st16':             ['Club Night', 'Clubrooms (BBQ & knockout pool comp, first 16)'],
  'howquawbee':                           ['Working Bee', 'Howqua'],
  'howquaworkingbee':                     ['Working Bee', 'Howqua'],
  'murrayroadworkingbee':                 ['Working Bee', 'Murray Road'],
  'howquaassociationdbackidsweekend':     ['Kids Weekend', 'Howqua (DBAC association weekend)'],
  'howquaassockidsweekendlodgeavailable': ['Kids Weekend', 'Howqua (lodge available to parents and kids)'],
  'howquaassocgonefishingdaynopoints':    ['Gone Fishing Day', 'Howqua (no points)'],
};

/**
 * Locations that need more than prefix-stripping — typos, duplicated text, and
 * notes carried up from the follow-on rows of the same outing. Keyed the same
 * squashed way as EVENTS, on the description *including* its water-type prefix.
 */
const LOCATIONS = {
  'fwhowquaareaincnillahcootie':            'Howqua Area (inc. Nillahcootie)',
  'swcliftonsprings':                       'Clifton Springs (one day, weather dependent)',
  'fwcamperdown':                           'Camperdown',
  'swbarwonheads':                          'Barwon Heads',
  'fwgoulburntributariesdaycomp':           'Goulburn & tributaries (day comp — Eildon Pondage to Hume Hwy bridge)',
  'fwhowquajohnkirkmansheildhowqauassoc':   'Howqua (John Kirkman Shield — Howqua Assoc.)',
  'swwerribeeweatherdependant':             'Werribee (weather dependent)',
  'sewerribeeweatherdependant':             'Werribee (weather dependent)',
  'fwlakewartookgrampians':                 'Lake Wartook, Grampians',
  'segoanywherewicr7pm':                    'Go Anywhere (WICR 7pm; BBQ clubrooms Sunday 7pm)',
  'swfgoanywherewicr7pm':                   'Go Anywhere (WICR 7pm; BBQ clubrooms Sunday 7pm)',
  'fwgoanywherewicr7pm':                    'Go Anywhere (WICR 7pm; BBQ clubrooms Sunday 7pm)',
  'fwhowquaincnillahcoote':                 'Howqua (inc. Nillahcootie)',
  'fwdartmouthpeterswampypattersonsheild':  'Dartmouth (Peter "Swampy" Patterson Shield)',
  'semarloinccorringleslipsseportalbert':   'Marlo (inc. Corringle Slips) & Port Albert',
  'swppbstkilda24hourbag':                  'Port Phillip Bay — St Kilda (24 hour bag, measured fish)',
  'fwhowquanativeclassicstarts400pm':       'Howqua (Native Classic, starts 4pm Friday)',
  'fweppalock':                             'Lake Eppalock',
  'fweppalockweekend':                      'Lake Eppalock',
  'semarloinccorringle':                    'Marlo (inc. Corringle)',
  'fwgoanywheremeasuredfishfwhowquainc':    'Go Anywhere (measured fish, inc. Howqua)',
  'swportalbertinterclub':                  'Port Albert (interclub)',
  'seportalbertinterclub':                  'Port Albert (interclub)',
  'segippslandlakesportalbert':             'Gippsland Lakes / Port Albert',
  'swlakesentranceportalbert':              'Lakes Entrance / Port Albert',
  'swcliftonspringsbellarinelodge':         'Clifton Springs (Bellarine Lodge?)',
  'fwhowqualespenrosesheild':               'Howqua (Les Penrose Shield)',
  'swcorinellaweatherdependant':            'Corinella (weather dependent)',
  'fweildonpondageonly':                    'Eildon Pondage only',
  'semaribyrnong3pmweighinatessendonanglers': 'Maribyrnong (3pm weigh-in at Essendon Anglers)',
  'fwgoanywherewicr700pm':                  'Go Anywhere (WICR 7pm; return to clubrooms to be eligible — BBQ)',
};

const squash = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
const tidy = (s) => String(s).replace(/\s+/g, ' ').trim();
const pad = (n) => String(n).padStart(2, '0');

function monthIdx(name) {
  const key = tidy(name).toLowerCase();
  const canonical = MONTH_ALIASES[key] ?? key;
  return MONTHS.findIndex((m) => m.toLowerCase() === canonical.toLowerCase());
}

/** A month header row: month name in A, 4-digit year in B, nothing else. */
function asMonthHeader(row) {
  const [a, b] = row;
  if (typeof a !== 'string' || typeof b !== 'number') return null;
  const mi = monthIdx(a);
  if (mi < 0 || b < 2000 || b > 2100) return null;
  return { month: mi, year: b };
}

/** Split "C" into a type + location, using the water-type prefix or the event table. */
function classify(desc) {
  const raw = tidy(desc);
  const event = EVENTS[squash(raw)];
  if (event) return { type: event[0], location: event[1], kind: 'event' };

  for (const [re, type] of TYPE_PREFIXES) {
    if (!re.test(raw)) continue;
    const location = LOCATIONS[squash(raw)] ?? tidy(raw.replace(re, ''));
    return { type, location, kind: 'trip' };
  }
  return null; // not an outing — headers, notes, trailing prose
}

function displayRange(a, b) {
  const [ay, am, ad] = a, [by, bm, bd] = b;
  if (ay === by && am === bm && ad === bd) return `${ad} ${MONTHS[am]} ${ay}`;
  if (ay === by && am === bm) return `${ad}-${bd} ${MONTHS[am]} ${ay}`;
  if (ay !== by) throw new Error(`Range spans two years (${a} -> ${b}); xlsx-to-trips cannot parse that display form`);
  return `${ad} ${MONTHS[am]} - ${bd} ${MONTHS[bm]} ${by}`;
}

// ---------------------------------------------------------------------------

const src = process.argv[2];
if (!src) {
  console.error('Usage: node scripts/calendar-to-syllabus.mjs "<calendar workbook>.xlsx"');
  process.exit(1);
}
const SRC = path.resolve(ROOT, src);

const wb = XLSX.read(fs.readFileSync(SRC), { type: 'buffer', cellDates: true });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, raw: true });

const warnings = [];
const entries = [];
let header = null;   // {month, year} from the most recent month header row
let cursor = null;   // where the current outing has got to; reset from `header` per outing
let block = [];      // dated rows of the outing being accumulated

/**
 * Resolve a day number against the current outing's cursor. Outings that run past
 * the end of the month (Sat 31 Oct, Sun 1, Mon 2...) stay under their month header,
 * so a day number that doesn't move forward means the next month.
 */
function resolveDay(dayNum, prev) {
  if (!cursor) cursor = { ...header };
  let { month, year } = cursor;
  let date = [year, month, dayNum];
  if (prev && (date[0] < prev[0] || (date[0] === prev[0] && (date[1] < prev[1] || (date[1] === prev[1] && date[2] <= prev[2]))))) {
    month += 1;
    if (month > 11) { month = 0; year += 1; }
    cursor = { month, year };
    date = [year, month, dayNum];
  }
  return date;
}

const dayAfter = ([y, m, d]) => {
  const n = new Date(y, m, d + 1);
  return [n.getFullYear(), n.getMonth(), n.getDate()];
};
const sameDate = (a, b) => a[0] === b[0] && a[1] === b[1] && a[2] === b[2];

function flush() {
  cursor = null;
  if (!block.length) return;
  const dated = block.filter((r) => r.date);
  const named = block.find((r) => r.info);
  block = [];
  if (!dated.length || !named) return;

  const first = dated[0].date, last = dated[dated.length - 1].date;
  entries.push({
    sort: `${first[0]}-${pad(first[1] + 1)}-${pad(first[2])}`,
    display: displayRange(first, last),
    type: named.info.type,
    location: named.info.location,
  });
}

for (const row of rows) {
  const cells = (row ?? []).map((c) => (c == null ? null : c));
  if (!cells.some((c) => c != null && String(c).trim() !== '')) { flush(); continue; }

  const monthHeader = asMonthHeader(cells);
  if (monthHeader) { flush(); header = monthHeader; continue; }

  const [dayName, dayNum, desc] = cells;
  const hasDay = typeof dayNum === 'number' && dayNum >= 1 && dayNum <= 31;
  const info = desc != null && String(desc).trim() !== '' ? classify(desc) : null;

  if (!hasDay) {
    // Undated line: a note attached to the current outing, or stray prose.
    if (block.length && info) block.push({ date: null, info });
    continue;
  }
  if (!header) { warnings.push(`Dated row before any month header: ${JSON.stringify(cells)}`); continue; }

  const prevDated = [...block].reverse().find((r) => r.date);
  const date = resolveDay(dayNum, prevDated?.date);

  // A jump in the dates means a new outing even without a blank row between them.
  if (prevDated && !sameDate(date, dayAfter(prevDated.date))) flush();

  if (typeof dayName === 'string') {
    const expected = DAYS[new Date(date[0], date[1], date[2]).getDay()];
    const stated = tidy(dayName).split(/[\s(]/)[0];
    if (stated && stated.toLowerCase() !== expected.toLowerCase()) {
      warnings.push(`${date[2]} ${MONTHS[date[1]]} ${date[0]} is a ${expected}, spreadsheet says "${tidy(dayName)}"`);
    }
  }
  block.push({ date, info });
}
flush();

// The source repeats a few single-day notices (e.g. "NO CLUB MEETING") once per
// parallel trip column; the app only needs one card.
const seen = new Set();
const unique = entries.filter((e) => {
  const key = [e.sort, e.display, e.type, e.location].join('|');
  if (seen.has(key)) return false;
  seen.add(key);
  return true;
});

unique.sort((a, b) => (a.sort < b.sort ? -1 : a.sort > b.sort ? 1 : a.location.localeCompare(b.location)));

const aoa = unique.map((e) => [e.display, e.type, e.location]);
const out = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1');
XLSX.writeFile(out, OUT);

for (const w of warnings) console.warn(`  ! ${w}`);
console.log(`Wrote ${unique.length} rows to ${path.relative(ROOT, OUT)} (from ${path.basename(SRC)})`);
for (const e of unique) console.log(`  ${e.display.padEnd(32)} ${e.type.padEnd(14)} ${e.location}`);
