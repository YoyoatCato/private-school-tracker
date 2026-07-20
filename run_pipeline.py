#!/usr/bin/env python3
"""
Private Schooling Status Tracker — daily sweep (Google Alerts / Google News edition)
====================================================================================

This is the "sweep." It reads the same searches you have in Google Alerts, pulls the
matching news items, sorts each into an OPENING or a CLOSURE, drops anything already in
your master workbook, and writes the two files the dashboard reads:

    data/candidates.json   (the review-queue cards)
    data/meta.json         (the "Last sweep" timestamp)

The dashboard already knows how to read these on load, every few hours, and when you
click "Run sweep now" — and it ADDS new items to the queue without wiping what's there.

------------------------------------------------------------------------------------
QUICK START
------------------------------------------------------------------------------------
1) Put this file (and, optionally, All_Closures___Openings.xlsx for de-duping) in the
   SAME folder as index.html.
2) (optional, for de-dupe)   pip install openpyxl
3) Run it:                   python3 run_pipeline.py
   -> creates data/candidates.json and data/meta.json
4) Serve the folder (fetch() does NOT work from a file:// path — it must be http):
       python3 -m http.server 8000
   then open  http://localhost:8000/index.html
5) In the dashboard, click "Run sweep now" (or just reload). New candidates appear
   in the review queue and accumulate until you approve/reject them.

------------------------------------------------------------------------------------
RUN IT AUTOMATICALLY EVERY WEEKDAY AT 8:00 AM  (cron)
------------------------------------------------------------------------------------
    crontab -e
    # minute hour * * days-of-week   (1-5 = Mon-Fri).  Use your machine's local time.
    0 8 * * 1-5  cd /full/path/to/the/folder && /usr/bin/python3 run_pipeline.py >> sweep.log 2>&1
(On a server, keep `python3 -m http.server` running, or host the folder on any static
 web host and point cron at it.)

------------------------------------------------------------------------------------
USE YOUR ACTUAL GOOGLE ALERTS FEEDS (exact parity)
------------------------------------------------------------------------------------
By default this queries Google News RSS for the QUERIES below (same idea as your
alerts, nothing to set up). To use your real alerts instead: in Google Alerts, edit an
alert -> "Deliver to" -> "RSS feed", copy the feed URL, and paste it into ALERT_FEEDS.
When ALERT_FEEDS is non-empty, those feeds are used instead of the Google News search.
"""

import json, re, html, urllib.parse, urllib.request
from datetime import datetime, timezone
from xml.etree import ElementTree as ET
from pathlib import Path

# ------------------------------------------------------------------ CONFIG
# Your Google Alerts search terms, exactly as you use them:
QUERIES = [
    "Catholic school closing",
    "COVID school closing",
    "New private school opening",
    "new school opening",
    "Private academy closures",
    "Private school closures",
    "School permanent closing",
]

# OPTIONAL: paste your Google Alerts RSS feed URLs here for exact parity.
# (Google Alerts -> edit alert -> Deliver to: RSS feed -> copy the URL.)
ALERT_FEEDS = [
    # "https://www.google.com/alerts/feeds/00000000000000000000/00000000000000000000",
]

WORKBOOK  = "All_Closures___Openings.xlsx"   # used only for de-duplication (optional)
OUT_DIR   = Path("data")                      # dashboard reads data/candidates.json
DAYS_BACK = 21                                # ignore items older than this many days
UA        = {"User-Agent": "Mozilla/5.0 (school-tracker sweep)"}

# ------------------------------------------------------------------ keyword rules
SCHOOL_HINTS = ("school", "academy", "prep", "montessori", "microschool", "micro-school",
                "micro school", "learning center", "learning academy")
CLOSE_HINTS  = ("clos", "shut", "shutter", "will close", "to close", "permanently")
OPEN_HINTS   = ("open", "opening", "opens", "launch", "new school", "new private",
                "new campus", "new academy", "expand", "expands", "to open", "set to open",
                "will open", "plans to open", "coming to", "under construction",
                "microschool", "micro-school", "micro school")
# almost never what we want (unless clearly a private K-12 school)
STOP_HINTS   = ("university", "board of education", "public school district",
                "school district", "charter school", "community college", "junior college")

# U.S.-only scope: skip items that clearly name a non-U.S. place.
# (Heuristic — it removes obvious foreign results; a human still reviews the rest,
#  since a foreign item with no country word in the headline can slip through.)
US_ONLY = True
NON_US_HINTS = (
    "india","delhi","mumbai","bengaluru","bangalore","odisha","kerala","punjab","uttar",
    "uae","dubai","sharjah","abu dhabi","qatar","doha","saudi","riyadh","kuwait","bahrain",
    "united kingdom"," u.k","uk school","london","england","scotland","wales","britain","ireland","dublin",
    "canada","canadian","ontario","toronto","quebec","alberta","manitoba","vancouver","calgary","quinte","napanee",
    "australia","sydney","melbourne","brisbane","new zealand","auckland",
    "nigeria","lagos","abuja","uganda","kenya","nairobi","ghana","accra","south africa","zimbabwe",
    "pakistan","karachi","lahore","bangladesh","dhaka","sri lanka","nepal","afghanistan",
    "singapore","malaysia","kuala lumpur","philippines","manila","indonesia","jakarta","vietnam","thailand",
    "hong kong","china","beijing","shanghai","japan","tokyo","korea","germany","france","spain","italy",
)


def fetch(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=45) as r:
        return r.read()


def gnews_url(query):
    return ("https://news.google.com/rss/search?q="
            + urllib.parse.quote(query) + "&hl=en-US&gl=US&ceid=US:en")


def parse_rss(xml_bytes):
    out = []
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return out
    for it in root.iter("item"):
        def g(tag):
            e = it.find(tag)
            return (e.text or "").strip() if e is not None and e.text else ""
        src = it.find("source")
        out.append({
            "title":  html.unescape(g("title")),
            "link":   g("link"),
            "pub":    g("pubDate"),
            "source": (src.text.strip() if src is not None and src.text else ""),
        })
    return out


def parse_date(s):
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z", "%a, %d %b %Y %H:%M:%S GMT"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


def clean_title(title):
    # Google News formats titles as "Headline - Source"; drop the trailing source.
    return re.sub(r"\s+-\s+[^-]+$", "", title).strip() if " - " in title else title.strip()


def classify(title):
    t = title.lower()
    if not any(h in t for h in SCHOOL_HINTS):
        return None
    if any(h in t for h in STOP_HINTS) and not any(k in t for k in ("private", "prep", "academy", "catholic")):
        return None
    is_close = any(h in t for h in CLOSE_HINTS)
    is_open  = any(h in t for h in OPEN_HINTS)
    if is_close and not is_open:
        return "closure"
    if is_open and not is_close:
        return "opening"
    if is_close:
        return "closure"   # tie-break: closures are the higher-signal case
    if is_open:
        return "opening"
    return None


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def load_workbook_names(path):
    names = set()
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        for sheet, col in (("All_closures", 1), ("All_openings", 1)):
            if sheet in wb.sheetnames:
                for row in wb[sheet].iter_rows(min_row=2, values_only=True):
                    if row[col]:
                        names.add(norm(str(row[col])))
    except Exception as e:
        print(f"  (workbook de-dupe skipped: {e})")
    return {n for n in names if len(n) >= 6}   # ignore very short/ambiguous names


# blank record templates matching the dashboard's column schema
OPEN_KEYS  = ["name","link","town","state","region","type","date_reported","opening_date",
              "exp_enrollment","source_links","notes","tuition","tuition_notes","link_b",
              "full_address","tract_mfi","data_year","microschool"]
CLOSE_KEYS = ["name","covid","merging","new_school","town","state","region","type","date_reported",
              "enrollment","enroll_source","year_founded","notes","prefin","link","pct_finaid","link2",
              "tuition","link3","pct_black","enr_black","pct_hispanic","enr_hispanic","pct_asian",
              "enr_asian","pct_white","enr_white","pct_other","enr_other","link4","full_address",
              "tract_mfi","data_year"]


def blank(keys):
    return {k: "" for k in keys}


def main():
    started = datetime.now()
    print("Sweep started:", started.isoformat(timespec="seconds"))

    known = load_workbook_names(WORKBOOK)
    print(f"  workbook: {len(known)} known school names for de-dupe")

    feeds = ALERT_FEEDS if ALERT_FEEDS else [gnews_url(q) for q in QUERIES]
    print(f"  reading {len(feeds)} feed(s): {'your Google Alerts' if ALERT_FEEDS else 'Google News for your queries'}")

    raw = []
    for url in feeds:
        try:
            raw += parse_rss(fetch(url))
        except Exception as e:
            print(f"  feed error: {e}")

    cutoff = datetime.now(timezone.utc).timestamp() - DAYS_BACK * 86400
    seen, candidates, seq = set(), [], 0

    for it in raw:
        title = clean_title(it["title"])
        kind = classify(title)
        if not title or not kind:
            continue
        if US_ONLY:
            hay = (title + " " + it["source"]).lower()
            if any(h in hay for h in NON_US_HINTS):
                print(f"  skip (outside U.S.): {title}")
                continue

        d = parse_date(it["pub"])
        if d:
            ts = (d if d.tzinfo else d.replace(tzinfo=timezone.utc)).timestamp()
            if ts < cutoff:
                continue
            date_reported = d.astimezone().strftime("%Y-%m-%d")
        else:
            date_reported = started.strftime("%Y-%m-%d")

        nn = norm(title)
        if nn in seen:
            continue
        seen.add(nn)

        if any(k in nn for k in known):        # already in the master workbook
            print(f"  skip (already tracked): {title}")
            continue

        seq += 1
        note = f"From the {it['source'] or 'news'} sweep — verify the details before approving."
        if kind == "opening":
            data = blank(OPEN_KEYS)
            data.update(name=title, link=it["link"], date_reported=date_reported,
                        source_links=it["source"], notes=note)
        else:
            data = blank(CLOSE_KEYS)
            data.update(name=title, link=it["link"], date_reported=date_reported,
                        enroll_source=it["source"], notes=note)
        candidates.append({"id": f"sw{started:%Y%m%d}{seq:03d}", "kind": kind,
                           "auto": ["name", "link", "date_reported"], "data": data})

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "candidates.json").write_text(json.dumps(candidates, indent=2))

    hour12 = started.strftime("%I").lstrip("0") or "12"
    last_run = started.strftime(f"%a, %b %d %Y &middot; {hour12}:%M %p") + " ET"
    (OUT_DIR / "meta.json").write_text(json.dumps({"last_run": last_run}))

    opens = sum(1 for c in candidates if c["kind"] == "opening")
    closes = len(candidates) - opens
    print(f"  wrote {len(candidates)} candidates ({opens} openings, {closes} closures) -> {OUT_DIR}/candidates.json")
    print("Sweep done.")


if __name__ == "__main__":
    main()
