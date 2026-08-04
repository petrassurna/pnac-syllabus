/**
 * Extract the syllabus from the committee's workbook into trips.json (the app's data).
 *
 * The workbook is a month-by-month calendar grid, read top to bottom:
 *   - a month header row:            A = month name, B = year
 *   - one row per DAY of an outing:  A = day name, B = day number, C = description,
 *                                    D = donor, E = trophy value   (D/E are dropped)
 *   - blank rows separate outings
 *
 * Column C carries the water type as an F/W, S/W or S/E prefix on the venue, or is
 * the name of a club event. Consecutive dated rows fold into one entry; a gap in the
 * dates starts a new one (the December block lists Christmas Tree 5-6 and the
 * break-up on the 11th without a blank row between them).
 *
 * Run:  npm run data              # newest Syllabus*.xlsx in the project root
 *       npm run data -- <file>    # a specific workbook
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'trips.json');

const MONTHS = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];
const DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
// The workbook spells a few of these its own way.
const MONTH_ALIASES = { febuary: 'February', feburary: 'February', sept: 'September' };

/** Water-type prefixes used in column C. These are the app's three trip types. */
const TYPE_PREFIXES = [
  [/^f\s*\/?\s*w\b[\s-]*/i, 'Freshwater'],
  [/^s\s*\/?\s*w\b[\s-]*/i, 'Saltwater'],
  [/^s\s*\/?\s*e\b[\s-]*/i, 'Surf/Estuary'],
];

/**
 * Club events (no water-type prefix), as [type, location]. The app coloursand
 * filters on type, so each event description has to be split into the two.
 * Keyed by the description squashed to lowercase alphanumerics, so spacing and
 * punctuation drift in the source is harmless.
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
  'howquaassocgonefishingdaynopoints':    ['Gone Fishing Day', 'Howqua (no points)'],
  // Kids weekends keep the committee's own wording, verbatim.
  'howquaassociationdbackidsweekend':     ['Kids Weekend', 'Howqua Association DBAC kids weekend'],
  'howquaassockidsweekendlodgeavailable': ['Kids Weekend', 'Howqua Assoc Kids weekend Lodge available to parents and kids'],
};

/**
 * Venues that need more than prefix-stripping — typos, duplicated text, and notes
 * carried up from the follow-on rows of the same outing. Keyed the same squashed
 * way as EVENTS, on the description *including* its water-type prefix. Anything
 * not listed here just has its prefix stripped and its whitespace tidied.
 */
const LOCATIONS = {
  'fwhowquaareaincnillahcootie':            'Howqua Area (inc. Nillahcootie)',
  'swcliftonsprings':                       'Clifton Springs (one day, weather dependent)',
  'swbarwonheads':                          'Barwon Heads',
  'fwgoulburntributariesdaycomp':           'Goulburn & tributaries (day comp)',
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
  'semaribyrnong3pmweighinatessendonanglers': 'Maribyrnong (3pm weigh-in at Essendon Anglers)',
  'fwgoanywherewicr700pm':                  'Go Anywhere (WICR 7pm; return to clubrooms to be eligible — BBQ)',
};

/**
 * Venue for one specific occurrence, overriding the EVENTS table. Keyed by the
 * entry's start date plus the squashed description, since the same function recurs
 * each season and the two occurrences can differ.
 */
const VENUE_BY_DATE = {
  // Not settled at the time ver4 went out. The workbook names no venue for any
  // club function; the "Clubrooms" in EVENTS is inferred, so say so where it isn't.
  '2027-08-07|presentationnight': 'Venue to be confirmed',
};

/**
 * Evening functions that occupy one day only. The workbook pads them out with a
 * blank-description row for the Sunday, which otherwise reads as a second day —
 * unlike a working bee or a trip, where that same padding really does mean the
 * outing runs the whole weekend.
 */
const SINGLE_DAY = new Set([
  'christmastree',
  'presentationnight',
]);

/**
 * Descriptions that begin a NEW outing rather than annotating the one above them,
 * as [type, location]. Consecutive dated rows normally fold into a single entry, so
 * a second comp on the Sunday of a weekend — written without a blank row above it,
 * and without a water-type prefix to give it away — would otherwise be swallowed as
 * a note on the Saturday's comp. There is no reliable way to tell the two apart
 * automatically, so each one is listed here explicitly.
 */
const SPLITS = {
  // ver4 runs two separate day comps across the first weekend of September 2026:
  // the Goulburn on the Saturday, Eildon Pondage on the Sunday.
  'eildonpondagetohumehwybridge': ['Freshwater', 'Eildon Pondage to Hume Hwy Bridge'],
};

/**
 * Dates the workbook states wrongly, confirmed with the club. Keyed by the date as
 * written plus the squashed description, so a correction can only ever fire on the
 * exact row it was meant for. The day-name cross-check below is what surfaces these;
 * once corrected the row stops warning, which is the signal the fix is right.
 *
 * The workbook is left untouched — it stays exactly as the committee sent it.
 */
const DATE_CORRECTIONS = {
  // ver4 says "Thursday 24 March 2027", but the 24th is a Wednesday and club
  // nights are Thursdays. Confirmed with the club: it should be the 25th.
  '2027-03-24|noclubmeeting': '2027-03-25',
};

const squash = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
const tidy = (s) => String(s).replace(/\s+/g, ' ').trim();
const pad = (n) => String(n).padStart(2, '0');
const isoOf = ([y, m, d]) => `${y}-${pad(m + 1)}-${pad(d)}`;

function monthIdx(name) {
  const key = tidy(name).toLowerCase();
  const canonical = MONTH_ALIASES[key] ?? key;
  return MONTHS.findIndex((m) => m.toLowerCase() === canonical.toLowerCase());
}

/** A month header row: month name in A, 4-digit year in B. */
function asMonthHeader(row) {
  const [a, b] = row;
  if (typeof a !== 'string' || typeof b !== 'number') return null;
  const mi = monthIdx(a);
  if (mi < 0 || b < 2000 || b > 2100) return null;
  return { month: mi, year: b };
}

/** Split column C into the app's type + location, or null if it isn't an outing. */
function classify(desc) {
  const raw = tidy(desc);
  const key = squash(raw);
  const event = EVENTS[key];
  if (event) return { type: event[0], location: event[1], kind: 'event', key };

  for (const [re, type] of TYPE_PREFIXES) {
    if (!re.test(raw)) continue;
    return { type, location: LOCATIONS[key] ?? tidy(raw.replace(re, '')), kind: 'trip', key };
  }
  return null; // column headers, stray notes, trailing prose
}

function displayRange(a, b) {
  const [ay, am, ad] = a, [by, bm, bd] = b;
  if (ay === by && am === bm && ad === bd) return `${ad} ${MONTHS[am]} ${ay}`;
  if (ay === by && am === bm) return `${ad}-${bd} ${MONTHS[am]} ${ay}`;
  if (ay !== by) return `${ad} ${MONTHS[am]} ${ay} - ${bd} ${MONTHS[bm]} ${by}`;
  return `${ad} ${MONTHS[am]} - ${bd} ${MONTHS[bm]} ${by}`;
}

/** Newest Syllabus*.xlsx in the project root, so next season's file drops straight in. */
function findWorkbook() {
  const candidates = fs.readdirSync(ROOT)
    .filter((f) => /^syllabus.*\.xlsx$/i.test(f) && !f.startsWith('~$'))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  if (!candidates.length) throw new Error(`No Syllabus*.xlsx found in ${ROOT}`);
  return candidates[candidates.length - 1];
}

// ---------------------------------------------------------------------------

const SRC = path.resolve(ROOT, process.argv[2] ?? findWorkbook());

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

  const first = dated[0].date;
  const last = SINGLE_DAY.has(named.info.key) ? first : dated[dated.length - 1].date;
  const start = isoOf(first);
  entries.push({
    start,
    display: displayRange(first, last),
    type: named.info.type,
    location: VENUE_BY_DATE[`${start}|${named.info.key}`] ?? named.info.location,
    kind: named.info.kind,
  });
}

for (const row of rows) {
  const cells = (row ?? []).map((c) => (c == null ? null : c));
  if (!cells.some((c) => c != null && String(c).trim() !== '')) { flush(); continue; }

  const monthHeader = asMonthHeader(cells);
  if (monthHeader) { flush(); header = monthHeader; continue; }

  const [dayName, dayNum, desc] = cells;
  const hasDay = typeof dayNum === 'number' && dayNum >= 1 && dayNum <= 31;
  const split = desc != null ? SPLITS[squash(desc)] : null;
  const info = split
    ? { type: split[0], location: split[1], kind: split[2] ?? 'trip' }
    : (desc != null && String(desc).trim() !== '' ? classify(desc) : null);

  if (!hasDay) {
    // Undated line: a note attached to the current outing, or stray prose.
    if (block.length && info) block.push({ date: null, info });
    continue;
  }
  if (!header) { warnings.push(`Dated row before any month header: ${JSON.stringify(cells)}`); continue; }

  const prevDated = [...block].reverse().find((r) => r.date);
  let date = resolveDay(dayNum, prevDated?.date);

  const fix = DATE_CORRECTIONS[`${isoOf(date)}|${squash(desc ?? '')}`];
  if (fix) {
    const [fy, fm, fd] = fix.split('-').map(Number);
    date = [fy, fm - 1, fd];
    cursor = { month: date[1], year: date[0] };
  }

  // A jump in the dates, or a row listed in SPLITS, starts a new outing even
  // without a blank row between them.
  if (split || (prevDated && !sameDate(date, dayAfter(prevDated.date)))) flush();

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
const trips = entries.filter((e) => {
  const key = [e.start, e.display, e.type, e.location].join('|');
  if (seen.has(key)) return false;
  seen.add(key);
  return true;
});

/**
 * The workbook lists Saltwater and Surf/Estuary as separate blocks even when they
 * run the same dates at the same venue, because each carries its own trophies. To a
 * member reading the app that's just the same weekend printed twice, so fold those
 * pairs into one card. The app understands the combined type: it draws a two-tone
 * accent, and its filter is a substring test, so the card still appears under both
 * Saltwater and Surf/Estuary. Different venues stay separate.
 */
const merged = [];
for (const trip of trips) {
  const twin = merged.find((m) => m.start === trip.start && m.display === trip.display &&
    m.location === trip.location &&
    ((m.type === 'Saltwater' && trip.type === 'Surf/Estuary') ||
     (m.type === 'Surf/Estuary' && trip.type === 'Saltwater')));
  if (twin) twin.type = 'Saltwater & Surf/Estuary';
  else merged.push(trip);
}
trips.length = 0;
trips.push(...merged);

trips.sort((x, y) => (x.start < y.start ? -1 : x.start > y.start ? 1 : x.location.localeCompare(y.location)));

fs.writeFileSync(OUT, JSON.stringify(trips, null, 2) + '\n', 'utf8');

for (const w of warnings) console.warn(`  ! ${w}`);
console.log(`Wrote ${trips.length} entries to ${path.relative(ROOT, OUT)} (from ${path.basename(SRC)})`);
