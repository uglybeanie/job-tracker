#!/usr/bin/env python3
"""
Job tracker for community organizing / advocacy / nonprofit communications roles.

Run:  python job_tracker.py
Pulls from Adzuna (aggregator) + public ATS boards (Greenhouse, Lever, Ashby)
for the organizations in config.json, scores and de-duplicates the results,
and merges them into jobs.xlsx WITHOUT overwriting Status / Notes you've typed.

Setup: see README.md
"""
import hashlib
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
CONFIG_PATH = HERE / "config.json"
OUTPUT_PATH = HERE / "jobs.xlsx"
TIMEOUT = 20
HEADERS = {"User-Agent": "personal-job-tracker/1.0"}

# ---------------------------------------------------------------- targets ----
TARGET_TITLES = [
    "community organizer", "field organizer", "civic engagement coordinator",
    "outreach coordinator", "outreach specialist", "coalition coordinator",
    "advocacy coordinator", "grassroots campaign associate",
    "communications coordinator", "communications specialist",
    "social media coordinator", "social media specialist",
    "digital communications specialist", "content coordinator",
    "content specialist", "community engagement coordinator",
    "public affairs coordinator", "stakeholder relations coordinator",
    "program coordinator", "program associate",
    "communications & outreach coordinator", "programs & communications associate",
]

# A job title must contain one of these to be kept at all.
TITLE_KEYWORDS = [
    "organizer", "organizing", "civic engagement", "outreach", "coalition",
    "advocacy", "grassroots", "communications", "social media", "digital",
    "content", "community engagement", "public affairs", "stakeholder",
    "program coordinator", "program associate", "programs", "engagement",
]
# ...and none of these (too senior).
EXCLUDE_TITLE = [
    "director", "vp", "vice president", "chief", "head of", "senior", "sr.",
    "sr ", "manager", "principal", "lead", "president", "counsel", "attorney",
    "engineer", "developer", "nurse", "physician",
]
# Bonus points when found in the title or description.
MISSION_KEYWORDS = [
    "advocacy", "civic engagement", "voter", "community organizing", "grassroots",
    "coalition", "nonprofit", "non-profit", "social justice", "equity", "campaign",
    "policy", "labor", "union", "public interest", "movement",
]
# Description patterns that signal too much experience required.
SENIOR_EXP = re.compile(r"\b([5-9]|1\d)\+?\s*(?:or more\s*)?years", re.I)


# ----------------------------------------------------------------- helpers ----
def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text((HERE / "config.example.json").read_text())
        print("Created config.json from config.example.json")
    return json.loads(CONFIG_PATH.read_text())


def job_id(org, title, location):
    key = "|".join(re.sub(r"\W+", " ", s.lower()).strip() for s in (org, title, location))
    return hashlib.sha1(key.encode()).hexdigest()[:12]


def get_json(url, **kw):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
        if r.status_code == 404:      # board doesn't exist at this ATS: expected, stay quiet
            return None
        r.raise_for_status()
        return r.json()
    except Exception as e:  # one failing source should never kill the run
        print(f"  ! {url[:80]}... failed: {e}")
        return None


def make_job(source, org, title, location, url, posted, description="", salary="", remote=False):
    return {
        "source": source, "org": (org or "").strip(), "title": (title or "").strip(),
        "location": (location or "").strip(), "url": url, "posted": posted,
        "description": re.sub(r"<[^>]+>", " ", description or "")[:3000],
        "salary": salary, "remote": bool(remote)
        or "remote" in (location or "").lower(),
    }


# ----------------------------------------------------------------- sources ----
# Adzuna free limits: 250 calls/day, 1,000/week, 2,500/month. Keep calls few.
ADZUNA_QUERIES = ["organizer", "civic engagement", "outreach coordinator", "advocacy coordinator",
                  "communications coordinator", "social media coordinator", "community engagement",
                  "program coordinator", "program associate", "coalition coordinator"]
ADZUNA_REMOTE_QUERIES = ["organizer", "communications coordinator", "outreach coordinator", "advocacy coordinator"]


def adzuna_due(cfg):
    """On GitHub (hourly runs) only hit Adzuna at a few set hours (UTC) to stay under the free limit."""
    if not os.getenv("GITHUB_ACTIONS"):
        return True
    return datetime.now(timezone.utc).hour in cfg.get("adzuna_hours_utc", [12, 17, 22])


def fetch_adzuna(cfg):
    app_id = os.getenv("ADZUNA_APP_ID") or cfg.get("adzuna_app_id")
    app_key = os.getenv("ADZUNA_APP_KEY") or cfg.get("adzuna_app_key")
    if not (app_id and app_key):
        print("Adzuna: no keys set, skipping (watchlist still runs)")
        return []
    if not adzuna_due(cfg):
        print("Adzuna: skipped this run (saving free quota)")
        return []
    base = {"app_id": app_id, "app_key": app_key, "results_per_page": 50,
            "max_days_old": cfg.get("max_days_old", 21), "sort_by": "date"}
    searches = [(q, {"where": w, "distance": cfg.get("radius_miles", 40)})
                for w in cfg["locations"] for q in ADZUNA_QUERIES]
    if cfg.get("include_remote", True):
        searches += [(q + " remote", {}) for q in ADZUNA_REMOTE_QUERIES]
    jobs = []
    for q, extra in searches:
        data = get_json("https://api.adzuna.com/v1/api/jobs/us/search/1",
                        params={**base, **extra, "what_phrase": q})
        for r in (data or {}).get("results", []):
            sal = ""
            if r.get("salary_min"):
                sal = f"${int(r['salary_min']):,}-${int(r.get('salary_max') or r['salary_min']):,}"
            jobs.append(make_job(
                "Adzuna", r.get("company", {}).get("display_name"), r.get("title"),
                r.get("location", {}).get("display_name") or "", r.get("redirect_url"),
                (r.get("created") or "")[:10], r.get("description"), sal,
                remote=q.endswith("remote")))
    print(f"Adzuna: {len(searches)} calls made")
    return jobs


def fetch_greenhouse(token, org):
    data = get_json(f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs")
    return [make_job("Greenhouse", org, j["title"], j.get("location", {}).get("name"),
                     j["absolute_url"], (j.get("updated_at") or "")[:10])
            for j in (data or {}).get("jobs", [])]


def fetch_lever(slug, org):
    data = get_json(f"https://api.lever.co/v0/postings/{slug}", params={"mode": "json"})
    out = []
    if not isinstance(data, list):
        return out
    for j in data:
        posted = datetime.fromtimestamp(j["createdAt"] / 1000, timezone.utc).strftime("%Y-%m-%d") \
            if j.get("createdAt") else ""
        out.append(make_job("Lever", org, j["text"], j.get("categories", {}).get("location"),
                            j["hostedUrl"], posted, j.get("descriptionPlain", "")))
    return out


def fetch_ashby(slug, org):
    data = get_json(f"https://api.ashbyhq.com/posting-api/job-board/{slug}")
    return [make_job("Ashby", org, j["title"], j.get("location"), j["jobUrl"],
                     (j.get("publishedAt") or "")[:10], j.get("descriptionPlain", ""),
                     remote=j.get("isRemote"))
            for j in (data or {}).get("jobs", [])]


def candidate_slugs(org):
    """Board names to try for an org: any listed in config, plus guesses from its name."""
    base = re.sub(r"[^a-z0-9 ]", "", org["name"].lower())
    guesses = [base.replace(" ", ""), base.replace(" ", "-")]
    return list(dict.fromkeys(org.get("slugs", []) + guesses))


def fetch_org(org):
    """Try every candidate board name on Greenhouse, Lever and Ashby; keep whatever exists."""
    out, found = [], []
    for slug in candidate_slugs(org):
        for kind, fn in (("Greenhouse", fetch_greenhouse), ("Lever", fetch_lever), ("Ashby", fetch_ashby)):
            res = fn(slug, org["name"])
            if res:
                out += res
                found.append(f"{kind}:{slug}")
    return org["name"], out, found


def fetch_all(cfg):
    from concurrent.futures import ThreadPoolExecutor
    jobs = fetch_adzuna(cfg)
    orgs = cfg.get("watchlist", [])
    with ThreadPoolExecutor(max_workers=8) as pool:
        results = list(pool.map(fetch_org, orgs))
    hit = [r for r in results if r[2]]
    for name, out, found in results:
        jobs += out
    print(f"Watchlist: found a job board for {len(hit)} of {len(orgs)} organizations")
    missing = [r[0] for r in results if not r[2]]
    if missing:
        print("  No Greenhouse/Lever/Ashby board found for: " + "; ".join(missing))
    return jobs


# --------------------------------------------------------- filter and score ----
def location_ok(job, cfg):
    loc = job["location"].lower()
    wanted = [l.lower().split(",")[0] for l in cfg["locations"]] + \
             [s.lower() for s in cfg.get("states", ["kentucky", ", ky"])]
    if any(w in loc for w in wanted):
        return True
    return cfg.get("include_remote", True) and job["remote"]


def score(job, cfg):
    t, text = job["title"].lower(), (job["title"] + " " + job["description"]).lower()
    if not any(k in t for k in TITLE_KEYWORDS):
        return None
    if any(re.search(rf"\b{re.escape(x.strip())}\b", t) for x in EXCLUDE_TITLE):
        return None
    if not location_ok(job, cfg):
        return None
    s = 40
    if any(tt in t for tt in TARGET_TITLES):
        s += 25                                   # exact title match
    s += min(20, 4 * sum(k in text for k in MISSION_KEYWORDS))
    if any(l.lower().split(",")[0] in job["location"].lower() for l in cfg["locations"]):
        s += 10                                   # local beats remote
    if SENIOR_EXP.search(job["description"]):
        s -= 25
    if job["posted"] and job["posted"] >= (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d"):
        s += 5                                    # fresh posting
    return max(0, min(100, s))


# --------------------------------------------------------------- spreadsheet ----
COLUMNS = ["Score", "Status", "Title", "Organization", "Location", "Remote", "Salary",
           "Posted", "Deadline", "Link", "Source", "First Seen", "Last Seen", "Notes", "ID"]
STATUSES = ["New", "Interested", "Applied", "Interview", "Offer", "Rejected", "Pass"]
USER_COLS = ("Status", "Deadline", "Notes")  # never overwritten by a refresh


def read_existing():
    rows = {}
    if OUTPUT_PATH.exists():
        ws = load_workbook(OUTPUT_PATH)["Jobs"]
        hdr = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            d = dict(zip(hdr, r))
            if d.get("ID"):
                rows[d["ID"]] = d
    return rows


def merge(existing, found, today):
    seen = set()
    for j in found:
        jid = job_id(j["org"], j["title"], j["location"])
        seen.add(jid)
        row = existing.get(jid, {})
        is_new = not row
        row.update({
            "Score": j["score"], "Title": j["title"], "Organization": j["org"],
            "Location": j["location"], "Remote": "Yes" if j["remote"] else "",
            "Salary": j["salary"] or row.get("Salary"), "Posted": j["posted"],
            "Link": j["url"], "Source": j["source"], "Last Seen": today, "ID": jid,
        })
        row.setdefault("First Seen", today)
        row.setdefault("Status", "New")
        existing[jid] = row
        if is_new:
            row["_new"] = True
    return seen


def write_xlsx(rows, today):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3A5F")
        c.alignment = Alignment(vertical="center")

    def sort_key(r):
        closed = r["Last Seen"] < (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d")
        return (closed, -(r.get("Score") or 0), r.get("Posted") or "")

    for r in sorted(rows.values(), key=sort_key):
        ws.append([r.get(c) for c in COLUMNS])
        cell = ws.cell(row=ws.max_row, column=COLUMNS.index("Link") + 1)
        if r.get("Link"):
            cell.hyperlink = r["Link"]
            cell.value = "Open posting"
            cell.font = Font(name="Arial", color="0563C1", underline="single")
    link_col = COLUMNS.index("Link") + 1
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.column != link_col:
                c.font = Font(name="Arial")

    widths = {"Score": 7, "Status": 12, "Title": 42, "Organization": 30, "Location": 24,
              "Remote": 8, "Salary": 18, "Posted": 11, "Deadline": 11, "Link": 14,
              "Source": 11, "First Seen": 11, "Last Seen": 11, "Notes": 40, "ID": 14}
    for i, name in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths[name]
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = ws.dimensions

    last = ws.max_row
    sc = COLUMNS.index("Status") + 1
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(sc)}2:{get_column_letter(sc)}2000")

    # Highlights: new (green), applied/interview (blue), closed/not seen 3+ days (grey text)
    rng = f"A2:{get_column_letter(len(COLUMNS))}2000"
    st, ls = get_column_letter(sc), get_column_letter(COLUMNS.index("Last Seen") + 1)
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR(${st}2="Pass",${st}2="Rejected")'], font=Font(color="999999", strike=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR(${st}2="Applied",${st}2="Interview",${st}2="Offer")'],
        fill=PatternFill("solid", bgColor="DDEBF7")))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${st}2="New"'], fill=PatternFill("solid", bgColor="E2F0D9")))

    # Legend sheet
    lg = wb.create_sheet("How to use")
    for line in [
        "Edit only: Status (dropdown), Deadline, Notes. Everything else refreshes automatically.",
        "Your Status / Deadline / Notes are preserved every time the script runs.",
        "Rows sorted by Score (title match, mission keywords, local, fresh). Likely-closed rows sink to the bottom.",
        "'Last Seen' older than ~3 days usually means the posting was taken down.",
        f"Last refreshed: {today}",
    ]:
        lg.append([line])
    lg.column_dimensions["A"].width = 110
    for r in lg.iter_rows():
        for c in r:
            c.font = Font(name="Arial")
    wb.save(OUTPUT_PATH)


# -------------------------------------------------------------------- alerts ----
def alert(new_rows, cfg):
    hits = [r for r in new_rows if (r.get("Score") or 0) >= cfg.get("alert_min_score", 60)]
    if not hits:
        return
    msg = "New job matches:\n" + "\n".join(
        f"• [{r['Score']}] {r['Title']} @ {r['Organization']} ({r['Location']})\n  {r['Link']}"
        for r in sorted(hits, key=lambda r: -r["Score"])[:15])
    print("\n" + msg)
    hook = os.getenv("SLACK_WEBHOOK_URL") or cfg.get("slack_webhook_url")
    if hook:
        requests.post(hook, json={"text": msg}, timeout=TIMEOUT)
    topic = os.getenv("NTFY_TOPIC") or cfg.get("ntfy_topic")  # free phone push: ntfy.sh
    if topic:
        requests.post(f"https://ntfy.sh/{topic}", data=msg.encode(), timeout=TIMEOUT)


# ---------------------------------------------------------------------- main ----
def main():
    cfg = load_config()
    today = datetime.now().strftime("%Y-%m-%d")
    raw = fetch_all(cfg)
    print(f"Fetched {len(raw)} raw postings")

    kept = []
    for j in raw:
        s = score(j, cfg)
        if s is not None and s >= cfg.get("min_score", 40):
            j["score"] = s
            kept.append(j)
    print(f"{len(kept)} passed title/location/score filters")

    existing = read_existing()
    first_run = not existing
    merge(existing, kept, today)
    new_rows = [r for r in existing.values() if r.pop("_new", False)]
    write_xlsx(existing, today)
    prev = []
    if (HERE / "jobs.json").exists():
        try:
            prev = [j["id"] for j in json.loads((HERE / "jobs.json").read_text())]
        except Exception:
            prev = []
    order = [i for i in prev if i in existing] + sorted(
        (i for i in existing if i not in prev), key=lambda i: -(existing[i].get("Score") or 0))
    export = [{"id": i, "title": existing[i]["Title"], "org": existing[i]["Organization"],
               "location": existing[i]["Location"], "remote": existing[i]["Remote"] == "Yes",
               "salary": existing[i].get("Salary") or "", "posted": existing[i].get("Posted") or "",
               "link": existing[i]["Link"], "source": existing[i]["Source"], "score": existing[i]["Score"],
               "firstSeen": existing[i]["First Seen"], "lastSeen": existing[i]["Last Seen"]} for i in order]
    (HERE / "jobs.json").write_text(json.dumps(export, indent=1))
    import csv  # rows only ever get appended, so a Google Sheet's extra columns stay lined up
    with open(HERE / "jobs.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(export[0].keys()) if export else ["id"])
        w.writeheader()
        w.writerows(export)
    print(f"Wrote {OUTPUT_PATH} ({len(existing)} total, {len(new_rows)} new)")
    if not first_run:
        alert(new_rows, cfg)


if __name__ == "__main__":
    main()
