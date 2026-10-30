"""
Scrape FICCI's Forthcoming Events (all states in one list) into Excel.
"""
import re
from datetime import date

import pandas as pd
from playwright.sync_api import sync_playwright

HEADLESS = False
BROWSER_CHANNEL = None           # "msedge" to use installed Edge
LIST_URL = "https://ficci.in/event"
ONLY_MONTHS = None             # None = ALL forthcoming events; e.g. {11, 12} to restrict to Nov-Dec
MIN_DATE = date.today()        # skip events that have already ended
FETCH_DETAILS = True
OUT = "FICCI_all_forthcoming_events.xlsx"
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

DATE_RE = re.compile(
    r"^(?P<m1>[A-Z][a-z]{2})\s+(?P<d1>\d{1,2})(?:\s*-\s*(?P<m2>[A-Z][a-z]{2})\s+(?P<d2>\d{1,2}))?,\s*(?P<y>\d{4})$")

# city/keyword -> state (lower-case keywords). Extend freely.
STATE_KEYWORDS = {
    "Delhi": ["delhi", "federation house", "noida"], "Maharashtra": ["mumbai", "pune", "nagpur", "maharashtra", "nariman"],
    "Rajasthan": ["jaipur", "jodhpur", "udaipur", "rajasthan"], "Assam": ["guwahati", "assam"],
    "Telangana": ["hyderabad", "telangana"], "Karnataka": ["bengaluru", "bangalore", "mysuru", "karnataka"],
    "Andhra Pradesh": ["visakhapatnam", "vijayawada", "amaravati", "andhra"], "Uttar Pradesh": ["lucknow", "kanpur", "agra", "varanasi", "uttar pradesh", "ghaziabad"],
    "Tamil Nadu": ["chennai", "coimbatore", "madurai", "tamil nadu"], "West Bengal": ["kolkata", "west bengal"],
    "Gujarat": ["ahmedabad", "gandhinagar", "surat", "vadodara", "gujarat"], "Madhya Pradesh": ["bhopal", "indore", "madhya pradesh"],
    "Punjab/Haryana/Chandigarh": ["chandigarh", "ludhiana", "gurugram", "gurgaon", "faridabad", "haryana", "punjab"],
    "Kerala": ["kochi", "trivandrum", "thiruvananthapuram", "kerala"], "Odisha": ["bhubaneswar", "odisha"],
    "Bihar": ["patna", "bihar"], "Uttarakhand": ["dehradun", "uttarakhand"], "Goa": ["goa", "panaji"],
    "Jharkhand": ["ranchi", "jamshedpur", "jharkhand"], "Chhattisgarh": ["raipur", "chhattisgarh"],
    "Himachal Pradesh": ["shimla", "himachal"], "Jammu & Kashmir": ["srinagar", "jammu"],
}
FOREIGN = ["japan", "russia", "moscow", "nigeria", "lagos", "turkey", "istanbul", "taiwan", "taipei", "sweden", "las vegas",
           " us", "usa", "uzbekistan", "dubai", "uae", "singapore", "china", "germany", "france", "italy", "uk", "london",
           "hong kong", "muscat", "oman", "qatar", "doha", "bahrain", "kuwait", "bangladesh", "sri lanka", "nepal", "vietnam", "thailand", "indonesia", "korea", "kenya", "egypt", "saudi"]


def infer_state(venue):
    v = " " + venue.lower().strip() + " "
    for st, keys in STATE_KEYWORDS.items():
        if any(k in v for k in keys):
            return st
    if any(f" {f.strip()} " in v or f.strip() in v.split(",")[-1] for f in FOREIGN):
        return "Outside India"
    return "Unknown - check"



COUNTRIES = {"Japan": ["japan"], "Russia": ["russia", "moscow"], "Nigeria": ["nigeria", "lagos"], "Turkey": ["turkey", "istanbul"],
             "Taiwan": ["taiwan", "taipei"], "Sweden": ["sweden", "stockholm"], "USA": ["las vegas", " us", "usa", "new york"],
             "Uzbekistan": ["uzbekistan", "tashkent"], "UAE": ["dubai", "uae", "abu dhabi"], "Oman": ["oman", "muscat"],
             "Hong Kong": ["hong kong"], "Singapore": ["singapore"], "China": ["china", "shanghai"], "Germany": ["germany", "frankfurt", "munich"],
             "France": ["france", "paris"], "UK": ["london", " uk"], "Italy": ["italy", "milan"], "Vietnam": ["vietnam"], "Thailand": ["thailand", "bangkok"],
             "Indonesia": ["indonesia", "jakarta"], "Bangladesh": ["bangladesh", "dhaka"], "Sri Lanka": ["sri lanka", "colombo"], "Kenya": ["kenya"],
             "Egypt": ["egypt"], "Saudi Arabia": ["saudi", "riyadh"], "Qatar": ["qatar", "doha"], "South Korea": ["korea", "seoul"]}


def infer_country(venue, state):
    if state != "Outside India":
        return "India" if state != "Unknown - check" else "Unknown - check"
    v = " " + venue.lower() + " "
    for c, keys in COUNTRIES.items():
        if any(k in v for k in keys):
            return c
    return "Outside India (country unclear)"


TYPE_RULES = [
    ("Awards call", r"call for application"),
    ("Meeting / interactive session", r"interactive session|meeting|roundtable|round table"),
    ("Business delegation / mission", r"delegation|business mission"),
    ("Roadshow", r"roadshow"),
    ("Trade-show pavilion", r"pavilion"),
    ("Workshop / training", r"workshop|programme|training|masterclass|webinar"),
    ("Summit / conference", r"summit|conclave|conference|symposium|forum|\bheal\b|congress|conclav"),
    ("Expo / trade show", r"\bexpo\b|\btex\b|\bfair\b|\bshow\b|interdye|inter food|pharmtech"),
    ("Awards", r"awards"),
]


def event_type(title):
    t = title.lower()
    for name, rx in TYPE_RULES:
        if re.search(rx, t):
            return name
    return "Other"


def parse_date_line(line):
    m = DATE_RE.match(line.strip())
    if not m:
        return None
    y = int(m["y"])
    m1 = MONTHS.get(m["m1"].lower())
    m2 = MONTHS.get((m["m2"] or m["m1"]).lower())
    if not m1 or not m2:
        return None
    start = date(y, m1, int(m["d1"]))
    end = date(y, m2, int(m["d2"])) if m["d2"] else start
    return start, end


def fmt_range(s, e):
    if s == e:
        return s.strftime("%d %b %Y")
    if s.month == e.month:
        return f"{s.day}-{e.day} {e.strftime('%b %Y')}"
    return f"{s.strftime('%d %b')} - {e.strftime('%d %b %Y')}"


def parse_blocks(text):
    """Turn page/card text into [{title, start, end, venue}]. Title = line before the date line."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    out = []
    for i, l in enumerate(lines):
        d = parse_date_line(l)
        if not d or i == 0:
            continue
        venue = lines[i + 1] if i + 1 < len(lines) else ""
        if venue.startswith("FICCI: forthcoming") or venue.lower().startswith(("event details", "add to calendar")):
            venue = ""
        out.append({"title": lines[i - 1], "start": d[0], "end": d[1], "venue": venue})
    return out


def extract_detail(page, url):
    page.goto(url, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(3500)
    text = page.inner_text("body")
    links = page.eval_on_selector_all("a[href]", "els => els.map(a => ({t: a.innerText.trim(), h: a.href}))")
    emails = {e.lower() for e in re.findall(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)}
    emails |= {l["h"].split("mailto:")[1].split("?")[0].lower() for l in links if l["h"].startswith("mailto:")}
    emails = sorted(e for e in emails if not e.startswith(("info@", "webmaster@")))
    phones = []
    for p in re.findall(r"(\+?\d[\d\s\-()]{8,17}\d)", text):
        digits = re.sub(r"\D", "", p)
        if 9 <= len(digits) <= 13 and not digits.startswith("20"):
            phones.append(p.strip())
    reg = next((l["h"] for l in links if re.search(r"regist|apply|enrol|book|buy ticket", l["t"], re.I)
                or "regist" in l["h"].lower()), "")
    org = re.search(r"Organi[sz](?:ed by|er)\s*[:\-]?\s*(.+)", text)
    summary = re.search(r"(?:Event Overview|About the Event|Overview|About)[^\n]*\n+(.{40,400})", text)
    return {"Contact Email": "; ".join(emails[:4]), "Contact Phone": "; ".join(dict.fromkeys(phones))[:80],
            "Registration Link": reg, "Organizer (detail page)": org.group(1).strip() if org else "",
            "Summary": summary.group(1).strip() if summary else "",
            "Detail text (check)": re.sub(r"\s+", " ", text)[:1200]}


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS,
                                    **({"channel": BROWSER_CHANNEL} if BROWSER_CHANNEL else {}))
        page = browser.new_context().new_page()
        page.set_default_timeout(60000)
        page.goto(LIST_URL, wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(4000)
        for i in range(80):                                   # scroll until list ends
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1200)
            if "No more content available" in page.inner_text("body"):
                break
        print(f"Scrolled {i + 1} times", flush=True)

        cards = page.eval_on_selector_all(
            "a", """els => els.filter(a => /event details/i.test(a.innerText)).map(a => {
                 let n = a, txt = '';
                 for (let i = 0; i < 6 && n; i++) {
                   n = n.parentElement;
                   if (n && /[A-Z][a-z]{2} \\d{1,2}(\\s*-\\s*[A-Z][a-z]{2} \\d{1,2})?, \\d{4}/.test(n.innerText)) { txt = n.innerText; break; }
                 }
                 return {href: a.href, text: txt};
               })""")
        events = []
        for c in cards:
            blocks = parse_blocks(c["text"])
            if blocks:
                b = blocks[0]
                b["href"] = c["href"]
                events.append(b)
        if not events:                                        # fallback: plain page text, no links
            events = [dict(b, href="") for b in parse_blocks(page.inner_text("body"))]
        print(f"Events on page: {len(events)}", flush=True)

        keep = [e for e in events if e["end"] >= MIN_DATE and (ONLY_MONTHS is None or e["start"].month in ONLY_MONTHS)]
        print(f"Forthcoming events kept: {len(keep)}", flush=True)
        rows = []
        for n, e in enumerate(keep, 1):
            row = {"Event Date": fmt_range(e["start"], e["end"]), "Start": e["start"], "Event Name": e["title"],
                   "Event Type": event_type(e["title"]), "Organizer": "FICCI", "Country": infer_country(e["venue"], infer_state(e["venue"])),
                   "State": infer_state(e["venue"]), "Venue / City": e["venue"],
                   "Event Details Link": e["href"], "Source": "ficci.in/event"}
            if FETCH_DETAILS and e["href"]:
                try:
                    row.update(extract_detail(page, e["href"]))
                except Exception as ex:
                    print(f"  ! detail failed for {e['title'][:50]}: {str(ex)[:80]}", flush=True)
                page.wait_for_timeout(2000)
            rows.append(row)
            print(f"  [{n}/{len(keep)}] {row['Event Date']} | {row['Country']} | {row['State']} | {e['title'][:55]}", flush=True)
        browser.close()
    if not rows:
        print("No events saved - send me the console output.")
        return
    df = pd.DataFrame(rows).sort_values("Start").drop(columns=["Start"])
    df.to_excel(OUT, index=False)
    print(f"Saved {len(df)} events to {OUT}")


if __name__ == "__main__":
    main()
