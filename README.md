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
| `Syllabus 2026-2027 ver4.xlsx` | **The single source of truth** — the committee's workbook |
| `scripts/xlsx-to-trips.mjs` | Converts the workbook → `trips.json` |
| `manifest.webmanifest`, `sw.js` | PWA manifest + service worker |
| `icon-*.png`, `apple-touch-icon.png`, `favicon-32.png` | App icons |

## Data pipeline (Excel → app)

**The committee's workbook is the only source of trips.** The app reads `trips.json`,
which is generated from it in one step:

```
Syllabus 2026-2027 ver4.xlsx  --npm run data-->  trips.json
```

So when the syllabus changes, update the workbook and regenerate:

```bash
npm install          # once, installs the xlsx reader
npm run data         # newest Syllabus*.xlsx in this folder -> trips.json
```

`npm run data` picks the highest-numbered `Syllabus*.xlsx` in the project root, so
next season's file drops straight in — it prints which one it used. Pass a path to
override: `npm run data -- "some other file.xlsx"`.

It also cross-checks every date against the day name written beside it and warns on
any disagreement. **Read those warnings** — they're mistakes in the source. A clean
run means every date in the workbook is internally consistent.

Where the club has confirmed the correct date, record it in the `DATE_CORRECTIONS`
table at the top of the script rather than editing the workbook, which stays exactly
as the committee sent it. A corrected row stops warning, so the check doubles as
confirmation the fix is right. One is recorded so far: ver4 lists the March 2027
"no club meeting" notice as Thursday the 24th, which is a Wednesday — it's the 25th.

### Workbook format (first sheet)

| Column | Meaning |
|--------|---------|
| A — Day name | `Saturday`, `Monday (Labour Day)`, … (used only to sanity-check the date) |
| B — Day number | Day of the month; the enclosing month header row supplies month + year |
| C — Description | `F/W`, `S/W` or `S/E` prefix + venue, **or** an event name (`AGM`, `Howqua W/Bee`, …) |
| D, E | Donor and trophy value — ignored |

Month header rows are `<Month name>` in A and `<year>` in B. Blank rows separate
outings; consecutive dated rows fold into one entry, and a gap in the dates starts a
new one.

**Watch for two comps run back-to-back over one weekend.** Written without a blank
row between them, and with the second lacking a water-type prefix, they look
identical to a comp with a note on its second day — compare the Goulburn/Eildon
Pondage weekend (two comps) with `SW Clifton Springs` / `(One day weather dependant)`
(one comp plus a note). Nothing in the file distinguishes them, so each second comp
has to be listed in the `SPLITS` table in the script.

**Trips at two different venues on the same weekend should be separate blocks** — the
app shows them as separate cards. Saltwater and Surf/Estuary blocks that share a
venue *and* dates are folded into one `Saltwater & Surf/Estuary` card, which still
appears under both filter chips; showing the same weekend twice just reads as a
duplicate. Different venues stay separate.

A blank-description row after an outing normally extends it to a second day. Evening
functions listed that way (`CHRISTMAS TREE` on the Saturday, blank Sunday) are one
day only, and are named in the `SINGLE_DAY` set.

### What the script produces

For each outing it writes:
- `start` — ISO date of the first day, used for sorting and "next up" detection;
- `display` — the human date shown on the card (`25-27 September 2026`);
- `type` — `Saltwater` / `Freshwater` / `Surf/Estuary`, or an event name;
- `location` — the venue shown as the card's title;
- `kind` — `trip` or `event`, which drives the colour and the filter chips.

The app needs `type` and `location` as separate fields, so each description in column
C has to be split into the two. That split lives in two tables at the top of
`scripts/xlsx-to-trips.mjs`:

- **`EVENTS`** — maps a club event's description to its `[type, location]`
  (`Howqua W/Bee` → `Working Bee` at `Howqua`).
- **`LOCATIONS`** — venue wording that needs more than having its `F/W`/`S/W`/`S/E`
  prefix stripped: typos, duplicated text, and notes stranded on an outing's
  follow-on rows. Anything not listed is just prefix-stripped and whitespace-tidied.
- **`VENUE_BY_DATE`** — venue for one specific occurrence, overriding `EVENTS`, since
  the same function recurs each season and the two can differ.

**The workbook names no venue for any club function.** The `Clubrooms` in `EVENTS` is
an inference, not sourced — worth checking against the committee before each season.

**If a venue reads oddly in the app, edit those tables — not `trips.json`**, which is
overwritten on every run. Keys are the workbook's own text squashed to lowercase
letters and digits, so spacing and punctuation drift in the source doesn't break them.

## Run locally

```bash
npm run serve        # http://localhost:8731
```

(Any static server works. A service worker needs a secure context; `localhost` counts.)

## Deploy

Live at **<https://pnac-syllabus.vercel.app/>**, deployed by Vercel from `main` —
**pushing to GitHub is what publishes to members' phones.** A build lands in a minute
or two. There is nothing to run by hand.

```bash
npm run data                        # regenerate after editing the workbook
git add -A && git commit -m "..."
git push origin main                # this is the deploy
```

**Verify against the live URL afterwards — don't assume the push worked.** Give the
build a minute, then (PowerShell):

```powershell
$d = (Invoke-WebRequest "https://pnac-syllabus.vercel.app/trips.json" `
      -Headers @{'Cache-Control'='no-cache'} -UseBasicParsing).Content | ConvertFrom-Json
"$($d.Count) entries; first $($d[0].display), last $($d[-1].display)"
```

Compare that count against what `npm run data` printed locally. They should match.

PWA install needs HTTPS, which Vercel provides. Android/Chrome shows an "Install app"
button; iOS/Safari uses Share → *Add to Home Screen* (the app shows a hint).

### How updates reach phones that already have the app

This is the part that bites. The service worker serves `trips.json` **network-first**,
so *data* changes land on the next online launch by themselves. Everything else —
`index.html`, the icons — is **cache-first and precached**, so an installed copy is
pinned to whatever shell it downloaded until the cache *name* changes.

> **⚠️ Bump `CACHE` in `sw.js` whenever you change `index.html`.**
> Forget, and members keep the old app indefinitely. The `activate` handler deletes
> every cache that isn't the current name, so a new name is what evicts the old one.
> A data-only change does **not** need a bump.

Two supporting pieces, already in place in `index.html`:

- `register('sw.js', { updateViaCache: 'none' })` — stops `sw.js` *itself* being
  served from the HTTP cache, the classic reason an update silently never arrives.
- a `controllerchange` listener that reloads once when a new worker takes over, so
  the update shows on the launch that fetches it rather than the one after. It is
  attached only when a controller already exists, or a first install would reload
  on its own `clients.claim()`.

If someone is ever truly wedged, deleting and re-adding the home-screen icon is the
reset.

## Open questions for the committee

Judgement calls in the current data that nobody has confirmed:

- **Presentation Night is treated as one day.** The club confirmed Christmas Tree is;
  Presentation Night has the identical shape in the workbook (Saturday named, blank
  Sunday) and a *night* can't span two days — but it wasn't explicitly confirmed.
- **Working bees and trips with that same blank-Sunday padding are left as two days**,
  on the assumption they really do run the weekend.
- **Every `Clubrooms` is inferred**, since the workbook names no venue for any club
  function. The club meets at 270 Murray Road, Preston, so it's a fair inference —
  but it isn't sourced.
- **Venues are lightly tidied** (typos fixed, notes merged in) *except* the two kids
  weekends, which the club asked to keep verbatim. That inconsistency is deliberate
  but unresolved — see `LOCATIONS`.

## Credits

App icon: brown trout (*Salmo trutta*) illustration by **Duane Raver / U.S. Fish and
Wildlife Service** — a work of the U.S. federal government, in the **public domain**.
It has been recomposed here (background removed, rotated into a leaping pose over water).
Source: <https://commons.wikimedia.org/wiki/File:Brown_trout_FWS_white_background.jpg>
