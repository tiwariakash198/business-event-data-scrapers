# Event Data Scrapers

Two Python scrapers that collect business and economics event listings into Excel for analysis.

| Scraper | Source | Collects | Output |
|---|---|---|---|
| `scrape_conferencealerts.py` | conferencealerts.co.in | Conferences for chosen months and countries: name, date, venue, organizer, themes, registration link, enquiry email, contact | `Business_Economics_Events_Nov-Dec_2026_scraped.xlsx` |
| `scrape_ficci_events.py` | ficci.in/event | FICCI's forthcoming events across all states in one list, with details from each event page | `FICCI_all_forthcoming_events.xlsx` |

Both use Playwright for browser automation and pandas for the Excel output.

## conferencealerts scraper

**Features**
- **Resumable.** Every saved event is appended to a progress CSV. Re-running skips what is already saved.
- **Link cache.** Event links are gathered once and stored, so later runs go straight to the event pages.
- **Non-event memory.** Pages that turn out not to be events are recorded and not visited again.
- **Retries and pacing.** Each page is tried twice, with a longer pause after repeated failures and randomized delays between pages.
- **Automatic Excel export.** The workbook is rewritten every 250 events and at the end of the run.

**Output columns**

`Event Date, Event Name, Organizer, Event Theme, Registration Link, Email for Enquiry, Venue, Contact, Country, Website Link, Source URL`

The venue is the city and country shown on the site. If the registration button opens a pop-up form
instead of a link, the event page URL is used.

**Settings** are constants at the top of `scrape_conferencealerts.py`:

| Setting | Meaning | Default |
|---|---|---|
| `MONTHS` | months to collect | `november`, `december` |
| `COUNTRIES` | countries to include | 17 countries across Asia and the Middle East |
| `TOPIC`, `SUBTOPICS` | topic filters | Business and Economics, and 26 subtopics |
| `HEADLESS` | hide the browser window | `False` |
| `BROWSER_CHANNEL` | use an installed browser such as `"msedge"` | bundled Chromium |
| `MAX_PAGES_PER_MONTH` | page limit when gathering links | 40 |
| `DELAY` | seconds between event pages | 1.5 to 3.5 |
| `PROGRESS_CSV`, `OUTPUT_XLSX` | output file names | `events_progress.csv`, `Business_Economics_Events_Nov-Dec_2026_scraped.xlsx` |

Keep the delays as they are.

## FICCI scraper

Reads FICCI's "Forthcoming Events" list, scrolls until the list ends, and optionally opens each event's
detail page.

**What it does**
- Parses each event's dates (single day or range), title and venue from the list.
- Skips events that have already ended (`MIN_DATE`), with an optional month filter (`ONLY_MONTHS`).
- Opens each detail page to collect contact emails and phone numbers, a registration link, the organizer, a short summary and the start of the page text for checking.
- Sorts the result by start date and writes one Excel sheet.

**Inferred columns.** `Event Type`, `State` and `Country` are inferred by keyword rules from the event title
and venue text (`TYPE_RULES`, `STATE_KEYWORDS`, `COUNTRIES`). Anything that matches no rule is labelled
`Unknown - check` or `Outside India (country unclear)`. Treat these as a first pass and spot-check them.

**Output columns**

`Event Date, Event Name, Event Type, Organizer, Country, State, Venue / City, Event Details Link, Source,
Contact Email, Contact Phone, Registration Link, Organizer (detail page), Summary, Detail text (check)`

**Settings** are constants at the top of `scrape_ficci_events.py`:

| Setting | Meaning | Default |
|---|---|---|
| `LIST_URL` | events list page | `https://ficci.in/event` |
| `ONLY_MONTHS` | restrict to months, for example `{11, 12}` | `None` (all forthcoming events) |
| `MIN_DATE` | skip events ending before this date | today |
| `FETCH_DETAILS` | open each event's detail page | `True` |
| `HEADLESS`, `BROWSER_CHANNEL` | browser visibility and choice | `False`, bundled Chromium |
| `OUT` | output file | `FICCI_all_forthcoming_events.xlsx` |

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate          # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Run

```bash
python scrape_conferencealerts.py
python scrape_ficci_events.py
```

To restart the conferencealerts crawl from scratch, delete `links_cache.json`, `events_progress.csv` and
`non_events.txt`. To re-gather its links only, delete `links_cache.json`.

## Generated files

Running the scrapers creates `events_progress.csv`, `links_cache.json`, `non_events.txt` and the Excel
workbooks. They include organizer contact details, so they are excluded from this repository
(see `.gitignore`).

## Limitations

- Parsing depends on each site's page layout and text labels. A redesign will break it.
- The conferencealerts link discovery is heuristic: it scrolls, clicks "next" or "load more" buttons, and falls back to `?page=N`.
- The FICCI scraper is not resumable. It writes the Excel file at the end of a run.
- FICCI phone numbers and emails are extracted with pattern matching and can include noise. Check them before use.
- Listings change over time, so two runs on different days will not match exactly.
