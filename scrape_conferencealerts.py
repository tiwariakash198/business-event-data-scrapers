"""
Collect business & economics conferences (Nov + Dec) from conferencealerts.co.in
and save organizer / email / contact / theme / venue details to Excel.

SETUP (once):
    pip install playwright pandas openpyxl
    playwright install chromium

RUN:
    python scrape_conferencealerts.py

Notes
- Link discovery is heuristic (scrolling, "next"/"load more" buttons and a ?page=N fallback).
- Runs a visible browser by default (HEADLESS = False). It is slower and more polite.
- Progress is saved to events_progress.csv after every event; re-running resumes.
- Check the site's Terms / robots.txt and keep the delays. Do not lower them.
"""
import csv
import os
import random
import re
from urllib.parse import quote_plus, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright

HEADLESS = False
BROWSER_CHANNEL = None   # set to "msedge" to use your installed Edge instead of bundled Chromium
MONTHS = ["november", "december"]
MAX_PAGES_PER_MONTH = 40
DELAY = (1.5, 3.5)  # seconds between event pages
PROGRESS_CSV = "events_progress.csv"
OUTPUT_XLSX = "Business_Economics_Events_Nov-Dec_2026_scraped.xlsx"

COUNTRIES = ["Bahrain", "Hong Kong", "India", "Indonesia", "Japan", "Kuwait", "Maldives",
             "Oman", "Qatar", "Saudi Arabia", "Singapore", "Sri Lanka", "Taiwan",
             "Thailand", "Turkey", "UAE", "Vietnam"]
TOPIC = "Business and Economics"
SUBTOPICS = ["Accounting", "Banking", "Banking and Finance", "Business", "Business Ethics",
             "Circular Economy", "Data Analytics", "Digital Transformation", "E-commerce",
             "Economics", "Entrepreneurship", "Finance", "Fintech", "Human Resources",
             "Innovation", "Insurance", "International Business", "Investment Strategies",
             "Logistics", "Management", "Marketing", "Procurement", "Quality Management",
             "Real Estate", "Risk Management", "Supply Chain"]

DOMAIN = "conferencealerts.co.in"
SKIP_FIRST_SEGMENTS = {"virtual", "add-event", "subscribe", "blog", "journal", "contact-us",
                       "contact", "login", "searchresult", "privacy-policy", "faq",
                       "terms-conditions", "terms", "about", "sitemap"}
NEXT_SELECTORS = ["a[rel=next]", "a:has-text('Next')", "button:has-text('Next')",
                  "li.next a", ".pagination .next a", "button:has-text('Load More')",
                  "a:has-text('Load More')", "a:has-text('›')", "a:has-text('»')"]
COLUMNS = ["Event Date", "Event Name", "Organizer", "Event Theme", "Registration Link",
           "Email for Enquiry", "Venue", "Contact", "Country", "Website Link", "Source URL"]


def go(page, url):
    """Load a page without waiting for 'networkidle' (some sites never go idle)."""
    page.goto(url, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(3000)       # let scripts render the content


def search_url(month, page_no=None):
    q = (f"countries={quote_plus(','.join(COUNTRIES))}&topics={quote_plus(TOPIC)}"
         f"&subtopics={quote_plus(','.join(SUBTOPICS))}&month={month}")
    if page_no:
        q += f"&page={page_no}"
    return f"https://{DOMAIN}/searchresult?{q}"


def candidates(page):
    hrefs = page.eval_on_selector_all(
        "a[href]",
        "els => els.filter(e => !e.closest('header,footer,nav')).map(e => e.href)")
    out = []
    for h in hrefs:
        u = urlparse(h)
        if u.netloc.replace("www.", "") != DOMAIN:
            continue
        seg = u.path.strip("/")
        if not seg or seg.split("/")[0] in SKIP_FIRST_SEGMENTS:
            continue
        out.append(h.split("#")[0])
    return out


def gather_links(page, month):
    seen, idle = [], 0
    seen_set = set()
    print(f"[{month}] opening results page...", flush=True)
    go(page, search_url(month))
    print(f"[{month}] loaded: '{page.title()}'", flush=True)
    for rnd in range(1, MAX_PAGES_PER_MONTH + 1):
        for _ in range(5):                       # scroll for lazy loading
            page.mouse.wheel(0, 4000)
            page.wait_for_timeout(600)
        new = 0
        for h in candidates(page):
            if h not in seen_set:
                seen_set.add(h)
                seen.append(h)
                new += 1
        clicked = False
        for sel in NEXT_SELECTORS:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible() and loc.is_enabled():
                try:
                    loc.click()
                    page.wait_for_timeout(2500)
                    clicked = True
                    break
                except Exception:
                    continue
        idle = 0 if new else idle + 1
        print(f"  [{month}] round {rnd}: {new} new links, {len(seen)} total, next-clicked={clicked}", flush=True)
        if idle >= 2:
            break
    if len(seen) < 25:                           # fallback: ?page=N pagination
        for n in range(2, MAX_PAGES_PER_MONTH):
            go(page, search_url(month, n))
            new = [h for h in candidates(page) if h not in seen_set]
            if not new:
                break
            for h in new:
                seen_set.add(h)
                seen.append(h)
    print(f"[{month}] candidate links: {len(seen)}")
    return seen


def field(text, label):
    m = re.search(rf"{re.escape(label)}\s*\n\s*(.+)", text)
    return m.group(1).strip() if m else ""


def get_registration(page, links, url):
    """Registration link from 'Attend as Listener/Presenter'. Falls back to the event page."""
    # 1) plain link with a real address
    for l in links:
        if (re.search(r"attend as (listener|presenter)", l["t"], re.I) and l["h"]
                and not l["h"].rstrip("/").endswith("#") and l["h"] != url):
            return l["h"]
    # 2) it is a button: click it and see where it leads
    for label in ("Attend as Listener", "Attend as Presenter"):
        loc = page.get_by_text(label, exact=False).first
        if not loc.count():
            continue
        before = page.url
        try:
            with page.context.expect_page(timeout=4000) as pop:
                loc.click()
            newp = pop.value
            newp.wait_for_load_state()
            dest = newp.url
            newp.close()
            return dest
        except Exception:
            if page.url != before:                 # navigated in the same tab
                dest = page.url
                try:
                    page.go_back(wait_until="domcontentloaded")
                except Exception:
                    pass
                return dest
    # 3) probably a pop-up form on the event page itself
    return url


def parse_event(page, url):
    go(page, url)
    text = page.inner_text("body")
    if "Event Organizer Details" not in text:
        return None                              # not an event page (category page etc.)
    try:
        title = page.locator("h1").first.inner_text().strip()
    except Exception:
        title = text.strip().splitlines()[0]

    date_line, location = "", ""
    for line in text.splitlines()[:40]:
        if re.match(r"^\s*\d{1,2}\s+[A-Za-z]+\s+\d{4}", line):
            parts = re.split(r"\s{2,}|\t", line.strip(), maxsplit=1)
            date_line = parts[0]
            location = parts[1] if len(parts) > 1 else ""
            break

    agenda = re.search(r"Agenda Overview\s*\n\s*\d+\s+Days?\s*\n\s*(.+)", text)
    event_date = agenda.group(1).strip() if agenda else (field(text, "Event Start Date") or date_line)

    email = field(text, "Event Enquiries Email ID")
    if not email:
        found = [e for e in re.findall(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)
                 if "conferencealerts" not in e]
        email = found[0] if found else ""

    tracks = ""
    m = re.search(r"Tracks Include:\s*\n(.*?)(?:\n\+\d+\s+More|\nVIEW ALL)", text, re.S)
    if m:
        tracks = "; ".join(t.strip() for t in m.group(1).splitlines() if t.strip())

    links = page.eval_on_selector_all(
        "a", "els => els.map(e => ({t: (e.innerText||'').trim(), h: e.href}))")
    website = next((l["h"] for l in links if l["t"].lower() == "click here"), "")
    reg = get_registration(page, links, url)

    country = location.split(",")[-1].strip() if location else ""
    return {
        "Event Date": event_date, "Event Name": title,
        "Organizer": field(text, "Organized By"),
        "Event Theme": tracks or title,
        "Registration Link": reg,
        "Email for Enquiry": email,
        "Venue": location,                       # site shows city + country only
        "Contact": field(text, "Contact Person"),
        "Country": country, "Website Link": website, "Source URL": url,
    }


LINKS_CACHE = "links_cache.json"       # delete this file to re-gather links from scratch
NON_EVENTS_FILE = "non_events.txt"


def load_done():
    if not os.path.exists(PROGRESS_CSV):
        return set()
    return set(pd.read_csv(PROGRESS_CSV)["Source URL"].dropna())


def save_excel():
    try:
        if os.path.exists(PROGRESS_CSV):
            df = pd.read_csv(PROGRESS_CSV).drop_duplicates(subset="Source URL")
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"Saved {len(df)} events to {OUTPUT_XLSX}", flush=True)
    except Exception as e:                       # e.g. file open in Excel
        print(f"Could not write Excel ({e}). The data is safe in {PROGRESS_CSV}.", flush=True)


def main():
    import json
    done = load_done()
    non_events = set()
    if os.path.exists(NON_EVENTS_FILE):
        non_events = set(open(NON_EVENTS_FILE, encoding="utf-8").read().split())
    new_file = not os.path.exists(PROGRESS_CSV)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, **({"channel": BROWSER_CHANNEL} if BROWSER_CHANNEL else {}))
        page = browser.new_context().new_page()
        page.set_default_timeout(60000)
        if os.path.exists(LINKS_CACHE):
            links = json.load(open(LINKS_CACHE, encoding="utf-8"))
            print(f"Loaded {len(links)} saved links from {LINKS_CACHE} (delete it to re-gather).", flush=True)
        else:
            links = []
            for month in MONTHS:
                for h in gather_links(page, month):
                    if h not in links:
                        links.append(h)
            json.dump(links, open(LINKS_CACHE, "w", encoding="utf-8"))
        todo = [u for u in links if u not in done and u not in non_events]
        print(f"Total candidate links: {len(links)} | still to visit: {len(todo)}", flush=True)
        consecutive_errors = 0
        try:
            with open(PROGRESS_CSV, "a", newline="", encoding="utf-8") as f, \
                 open(NON_EVENTS_FILE, "a", encoding="utf-8") as ne:
                w = csv.DictWriter(f, fieldnames=COLUMNS)
                if new_file:
                    w.writeheader()
                saved_now = 0
                for i, url in enumerate(todo, 1):
                    row, failed = None, False
                    for attempt in (1, 2):
                        try:
                            row = parse_event(page, url)
                            failed = False
                            break
                        except Exception as e:
                            failed = True
                            print(f"  ! attempt {attempt} failed on {url}: {str(e)[:120]}", flush=True)
                            page.wait_for_timeout(5000)
                    if failed:
                        consecutive_errors += 1
                        if consecutive_errors >= 3:
                            print("  ! Several failures in a row - internet down? Waiting 60 s...", flush=True)
                            page.wait_for_timeout(60000)
                            consecutive_errors = 0
                        continue                  # not saved, so a re-run will retry it
                    consecutive_errors = 0
                    if row is None:
                        ne.write(url + "\n"); ne.flush()
                    else:
                        w.writerow(row); f.flush()
                        saved_now += 1
                        print(f"  [{i}/{len(todo)}] {row['Event Date']} | {row['Event Name'][:60]}", flush=True)
                        if saved_now % 250 == 0:
                            save_excel()
                    page.wait_for_timeout(int(random.uniform(*DELAY) * 1000))
        finally:
            browser.close()
            save_excel()


if __name__ == "__main__":
    main()
