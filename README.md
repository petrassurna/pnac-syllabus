# Preston Northcote Angling Club — Syllabus 2026/27

A mobile-first, installable **PWA** that shows the club's fishing trips and events
for the 2026/27 season. It opens with a short intro animation, then fades to a
colour-coded, filterable listing that auto-scrolls to the next upcoming date.

It's plain HTML/CSS/JS — no framework, no build step for the app itself. The only
tooling is a small script that turns the syllabus spreadsheet into the app's data.

## Features

- **Locked filter bar** — All · Saltwater · Freshwater · Surf/Estuary · Events
- **Next-up highlight** — auto-scrolls to and badges the next upcoming date
- **Installable** to the iOS / Android home screen (standalone, offline-capable)
- **Offline** via a service worker that caches the app shell and data
- Dark, glassy UI with month section headers

## Project layout

| Path | Role |
|------|------|
| `index.html` | The whole app (HTML + CSS + JS inline) |
| `trips.json` | The app's data — **generated** (do not hand-edit) |
| `Syllabus 2026-2027 ver4.xlsx` | **Source of truth** — the committee's calendar-style workbook |
| `Syllabus.xlsx` | Flattened intermediate (Date, Type, Location) — **generated** |
| `scripts/calendar-to-syllabus.mjs` | Converts the committee workbook → `Syllabus.xlsx` |
| `scripts/xlsx-to-trips.mjs` | Converts `Syllabus.xlsx` → `trips.json` |
| `manifest.webmanifest`, `sw.js` | PWA manifest + service worker |
| `icon-*.png`, `apple-touch-icon.png`, `favicon-32.png` | App icons |

## Data pipeline (Excel → app)

The app reads `trips.json`, which is generated in two steps:

```
Syllabus 2026-2027 ver4.xlsx  --npm run convert-->  Syllabus.xlsx  --npm run data-->  trips.json
```

The committee maintains the syllabus as a month-by-month calendar grid (one row per
*day* of each outing). `npm run convert` folds that into the flat three-column
`Syllabus.xlsx`; `npm run data` turns that into the app's JSON.

```bash
npm install                                       # once, installs the xlsx reader
npm run convert -- "Syllabus 2026-2027 ver4.xlsx" # calendar workbook -> Syllabus.xlsx
npm run data                                      # Syllabus.xlsx      -> trips.json
```

`npm run convert` prints every row it produced, plus warnings where a date and its
day name disagree — worth reading, since those are usually mistakes in the source.

### Committee workbook format (first sheet)

| Column | Meaning |
|--------|---------|
| A — Day name | `Saturday`, `Monday (Labour Day)`, … (used only to sanity-check the date) |
| B — Day number | Day of the month; the enclosing month header row supplies month + year |
| C — Description | `F/W`, `S/W` or `S/E` prefix + venue, **or** an event name (`AGM`, `Howqua W/Bee`, …) |
| D, E | Donor and trophy value — ignored |

Month header rows are `<Month name>` in A and `<year>` in B. Blank rows separate
outings; consecutive dated rows fold into one entry, and a gap in the dates starts a
new one. Venue wording that needs tidying (typos, notes on follow-on rows) lives in
the `EVENTS` / `LOCATIONS` tables at the top of `scripts/calendar-to-syllabus.mjs`.

### Intermediate format (`Syllabus.xlsx`, first sheet, no header row)

| Column | Meaning | Examples |
|--------|---------|----------|
| A — Date | A real date for a single day, **or** a text range for a weekend/long weekend | `2026-10-26`  ·  `27-28 July 2026`  ·  `30 November - 1 December 2026` |
| B — Type | Water type or event name | `Saltwater`, `Freshwater`, `Surf/Estuary`, `Saltwater & Surf/Estuary`, `Working Bee`, `AGM`, `Presentation Night`, `Kids Night` … |
| C — Location | Place / venue | `Port Albert`, `Howqua`, `Clubrooms` … |

The script derives, for each row:
- `start` — an ISO date used for sorting and "next up" detection (parsed from the range text when A is a range);
- `display` — the human date shown on the card;
- `kind` — `trip` if Type contains Saltwater/Freshwater/Surf/Estuary, otherwise `event`.

**Trips at two different venues on the same weekend should be separate rows** — the
app shows them as separate cards.

## Run locally

```bash
npm run serve        # http://localhost:8731
```

(Any static server works. A service worker needs a secure context; `localhost` counts.)

## Deploy (for phone install)

PWA install requires **HTTPS**. Host this folder on any static host —
GitHub Pages, Netlify, Cloudflare Pages, or your own web server — then:

- **Android / Chrome:** an "Install app" button appears automatically.
- **iOS / Safari:** use Share → *Add to Home Screen* (the app shows a hint).

## Credits

App icon: brown trout (*Salmo trutta*) illustration by **Duane Raver / U.S. Fish and
Wildlife Service** — a work of the U.S. federal government, in the **public domain**.
It has been recomposed here (background removed, rotated into a leaping pose over water).
Source: <https://commons.wikimedia.org/wiki/File:Brown_trout_FWS_white_background.jpg>
