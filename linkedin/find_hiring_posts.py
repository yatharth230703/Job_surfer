"""Collect LinkedIn posts from people who are HIRING for Soumya's kind of roles, into hiring_posts.xlsx.

    .venv/bin/python find_hiring_posts.py --searches 10            # run 10 searches, append new posts
    .venv/bin/python find_hiring_posts.py --searches 2 --dry-run   # print what it would keep, write nothing
    .venv/bin/python find_hiring_posts.py --selftest

Each search = one hiring phrase + one role/skill term, e.g.  hiring "content writer"  (Posts tab, Latest, Past week).
LinkedIn returns nothing for long Boolean queries (> ~3 quoted phrases), so terms are searched one at a time.
Terms searched longest ago go first (post_searches.csv), so repeated runs rotate through the whole bank.
A post is kept only if it's < 7 days old and its text scores as "hiring" (not "looking for a job").
The sheet is raw material for a later AI ranking step: every row carries its signals and score.
"""
import argparse
import csv
import hashlib
import random
import re
import time
import urllib.parse as up
from datetime import datetime, timedelta
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

import send_request as sr

HERE = Path(__file__).parent
POSTS_XLSX = HERE / "hiring_posts.xlsx"
SEARCH_LOG = HERE / "post_searches.csv"

# ---------------- keyword bank ----------------
# Short queries only: `hiring "business analyst"` works, 3+ quoted phrases with OR/AND returns zero results.
HIRING_PHRASES = ["hiring", "hiring", '"we are hiring"', '"looking to hire"', '"looking for"']  # plain 'hiring' twice: best yield

ROLE_CLUSTERS = {
    # content creation & management
    "content writing": ['"content writer"', '"content writing"', 'copywriter', '"blog writer"', '"SEO writer"'],
    "content strategy": ['"content strategist"', '"content creator"', '"content marketing"', '"content manager"', '"content associate"'],
    "ghost/script writing": ['ghostwriter', '"LinkedIn ghostwriter"', 'scriptwriter', '"newsletter writer"', '"UGC creator"'],
    "social media": ['"social media manager"', '"social media executive"', '"social media intern"', '"social media associate"', '"community manager"'],
    "medical writing": ['"medical writer"', '"medical content writer"', '"scientific writer"', '"healthcare content"', '"pharma content"'],
    "comms & PR": ['"communications associate"', '"PR executive"', '"public relations"', '"brand associate"', '"influencer marketing"'],
    # founder's office
    "founder's office": ['"founder\'s office"', '"founders office"', '"founder\'s associate"', '"office of the CEO"', '"CEO office"'],
    "chief of staff": ['"chief of staff"', '"strategy and operations"', '"business operations associate"', '"special projects"', '"BizOps"'],
    # growth & marketing
    "growth": ['"growth associate"', '"growth intern"', '"growth marketer"', '"growth analyst"', '"growth executive"'],
    "digital marketing": ['"digital marketing executive"', '"performance marketing"', '"marketing associate"', '"email marketing"', '"marketing intern"'],
    # business development / sales / CS
    "business development": ['"business development associate"', '"business development executive"', '"BD associate"', '"partnerships associate"', '"strategic partnerships"'],
    "customer success": ['"customer success associate"', '"customer success executive"', '"account executive"', '"key account"', '"inside sales"'],
    # pharma & healthcare
    "pharmacovigilance": ['pharmacovigilance', '"drug safety associate"', '"PV associate"', '"case processing"', '"safety associate"'],
    "regulatory & quality": ['"regulatory affairs"', '"QA executive"', '"QC executive"', '"quality assurance" pharma', '"quality control" pharma'],
    "clinical research": ['"clinical research associate"', '"clinical research coordinator"', '"clinical data"', '"clinical operations"', '"medical affairs"'],
    "healthcare ops": ['"healthcare operations"', '"hospital operations"', '"medical coding"', '"pharmacy operations"', 'healthtech'],
    "pharma sales": ['"medical representative"', '"pharma sales"', '"medical business development"', '"institutional sales"', '"provider relations"'],
    # consulting & analysis
    "healthcare consulting": ['"healthcare consulting"', '"healthcare analyst"', '"life sciences analyst"', '"pharma consulting"', '"healthcare strategy"'],
    "analyst": ['"business analyst"', '"research analyst"', '"market research"', '"strategy analyst"', '"operations analyst"'],
    # freelance / assistant / ops
    "virtual assistant": ['"virtual assistant"', '"executive assistant"', '"personal assistant"', '"admin assistant"', '"operations assistant"'],
    "research & docs": ['"research assistant"', '"content researcher"', '"documentation"', '"project coordinator"', '"program associate"'],
    "design support": ['"presentation designer"', '"Canva designer"', '"pitch deck"', '"graphic design intern"', '"social media designer"'],
    # skills, for posts that describe the work rather than the title
    "skills: writing": ['storytelling', 'copywriting', '"SEO writing"', '"LinkedIn content"', '"long-form content"'],
    "skills: tools": ['Canva', 'Notion', '"Google Sheets"', '"Excel"', '"AI tools"'],
    "skills: outreach": ['"cold outreach"', '"lead generation"', '"email outreach"', '"influencer outreach"', '"partnership outreach"'],
}


TERMS = [(c, t) for c, ts in ROLE_CLUSTERS.items() for t in ts]  # (cluster, term)


def search_url(query):
    q = {"keywords": query, "datePosted": '"past-week"', "sortBy": '"date_posted"'}
    return "https://www.linkedin.com/search/results/content/?" + up.urlencode(q)


# ---------------- is this person hiring? ----------------
_r = lambda *ps: [re.compile(p, re.I | re.M) for p in ps]
HIRE_SIGNALS = _r(
    r"\bwe(?:'re|’re| are)? (?:now )?hiring\b", r"#hiring\b", r"\bnow hiring\b", r"\bis hiring\b", r"\bhiring (?:for|a|an|now)\b",
    r"\blooking to hire\b", r"\b(?:we(?:'re|’re| are)|team is|company is|startup is) (?:looking|searching) for\b",
    r"^\W*(?:looking|searching) for (?!a job|job|work|opportunit|referral)", r"\bopen (?:positions?|roles?)\b",
    r"\bjob opening", r"\bvacanc(?:y|ies)\b", r"\bjoin (?:our|my) (?:team|startup|company)\b", r"\bapplications? (?:are )?open\b",
    r"\b(?:send|share|drop|email|mail) (?:me |us )?(?:your |some |a )?(?:cv|resume|portfolio|samples|work samples)\b",
    r"\bapply (?:here|now|via|at|through|using|on|by)\b", r"\binterested (?:candidates|applicants|folks|people)\b",
    r"\bto apply\b", r"\b(?:job )?description\s*[:\-]", r"\b(?:location|stipend|ctc|salary|experience required)\s*[:\-]",
    r"\bresponsibilities\b", r"\bdm (?:me|us)\b", r"\b(?:position|role|job title)\s*[:\-]", r"\bwe(?:'re|’re| are) (?:growing|expanding)\b",
)
SEEK_SIGNALS = _r(
    r"#?open ?to ?work\b", r"\bi(?:'m|’m| am) (?:actively |currently )?(?:looking|seeking|searching) for (?:a |an |new |my next |)"
    r"(?:job|role|position|opportunit|internship|full[- ]time)", r"\bseeking (?:a |new |a new |)(?:opportunit|role|position|job|internship)",
    r"\bactively (?:looking|seeking|job ?hunting|applying)\b", r"\bhire me\b", r"\bi(?:'m|’m| am) available for\b",
    r"\bplease (?:refer|help me find)\b", r"\blaid off\b", r"\bmy (?:resume|cv) is (?:attached|below)\b",
    r"\bopen to (?:new )?(?:opportunit|roles|work)", r"\breferral (?:request|needed)\b", r"\blooking for (?:a )?(?:job|referral)\b",
)
MIN_SCORE = 2  # ponytail: regex score; the AI step does the real judging, this only drops obvious non-hiring posts


def classify(text, open_to_work_badge=False):
    hit = lambda pats: [m.group(0).strip() for m in (p.search(text) for p in pats) if m]  # the matched words, readable
    hire, seek = hit(HIRE_SIGNALS), hit(SEEK_SIGNALS) + (["author has Open to Work badge"] if open_to_work_badge else [])
    return len(hire) - 2 * len(seek), hire, seek


AGE = re.compile(r"^(\d+)\s*(m|h|d|w|mo|yr)\b")
UNIT_MIN = {"m": 1, "h": 60, "d": 1440, "w": 10080, "mo": 43200, "yr": 525600}


def age_minutes(stamp):
    """'33m' -> 33, '3d • Edited' -> 4320; None if it isn't a timestamp."""
    m = AGE.match(stamp.strip())
    return int(m.group(1)) * UNIT_MIN[m.group(2)] if m else None


# ---------------- page reading ----------------
CARDS_JS = r"""() => [...document.querySelectorAll('[role=listitem][componentkey^="update-card-focus"]')].map(c => {
    const menu = c.querySelector('button[aria-label^="Open control menu for post by"]');
    const a = c.querySelector('a[href*="/in/"], a[href*="/company/"]');
    const box = c.querySelector('[data-testid="expandable-text-box"]');
    const ps = [...c.querySelectorAll('p')].map(p => p.innerText.trim()).filter(Boolean);
    const ti = ps.findIndex(t => /^\d+\s*(m|h|d|w|mo|yr)\b/.test(t));
    const head = ti > 0 ? ps.slice(0, ti).filter(t => !t.startsWith('•')) : [];
    return {
        key: c.getAttribute('componentkey'),
        author: menu ? menu.getAttribute('aria-label').replace('Open control menu for post by ', '').trim() : '',
        author_url: a ? a.href.split('?')[0] : '',
        headline: head.length > 1 ? head[head.length - 1] : '',
        stamp: ti >= 0 ? ps[ti] : '',
        otw: !!c.querySelector('[aria-label*="Open to work"]'),
        text: box ? box.innerText.replace(/\s*…\s*more\s*$/, '').trim() : '',
    };
})"""
EMAILS = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")


def post_id(p):
    return hashlib.sha1(f"{p['author']}|{p['text'][:300]}".encode()).hexdigest()[:16]


def copy_link(page, card_key):
    """Post's own URL via '…' -> 'Copy link to post' (the card has no permalink in the DOM)."""
    try:
        page.evaluate("navigator.clipboard.writeText('')")  # so a failed copy can't hand back the previous post's link
        card = page.locator(f'[role=listitem][componentkey="{card_key}"]').first
        sr.human_click(page, card.locator('button[aria-label^="Open control menu for post by"]').first)
        sr.pause(0.5, 1.2)
        sr.human_click(page, page.get_by_text("Copy link to post", exact=True).first)
        sr.pause(0.5, 1.0)
        return page.evaluate("navigator.clipboard.readText()").strip()
    except Exception as e:
        print(f"     (no link: {e.__class__.__name__})")
        page.keyboard.press("Escape")
        return ""


def skim_results(page, rounds):
    """Wheel down the results like a reader; LinkedIn lazy-loads ~10 posts per screenful."""
    vw, vh = page.evaluate("[innerWidth, innerHeight]")
    sr.move_to(page, vw * random.uniform(0.3, 0.45), vh * random.uniform(0.4, 0.6))
    for _ in range(rounds):
        for _ in range(random.randint(4, 9)):
            page.mouse.wheel(0, random.uniform(80, 180))
            time.sleep(random.uniform(0.02, 0.08))
        sr.pause(1.2, 4)
        if random.random() < 0.3:
            sr.wander(page)


# ---------------- storage ----------------
XLSX_COLS = [("Post ID", 18), ("Found on", 17), ("Posted (approx)", 17), ("Age", 8), ("Author", 24), ("Author headline", 40),
             ("Author URL", 36), ("Post link", 30), ("Post text", 110), ("Emails", 28), ("Hire score", 10),
             ("Hire signals", 40), ("Seeker signals", 30), ("Cluster", 20), ("Search query", 30), ("AI verdict", 14),
             ("AI notes", 40)]


def open_sheet():
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    if POSTS_XLSX.exists():
        wb = load_workbook(POSTS_XLSX)
        return wb, wb.active
    wb = Workbook(); ws = wb.active; ws.title = "Hiring posts"
    ws.append([c for c, _ in XLSX_COLS])
    for i, (_, w) in enumerate(XLSX_COLS, 1):
        ws.cell(1, i).font = Font(bold=True)
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"
    return wb, ws


def save_sheet(wb):
    try:
        wb.save(POSTS_XLSX)
    except Exception as e:  # file open in Excel -> keep going, rows are retried next save
        print(f"   (couldn't save {POSTS_XLSX.name}: {e})")


def searches_by_staleness():
    last = {}
    if SEARCH_LOG.exists():
        for r in csv.DictReader(open(SEARCH_LOG)):
            last[r["term"]] = r["time"]
    terms = TERMS[:]
    random.shuffle(terms)
    return sorted(terms, key=lambda ct: last.get(ct[1], ""))  # never-searched first, then oldest


def log_search(term, query, cards, kept):
    new = not SEARCH_LOG.exists()
    with open(SEARCH_LOG, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["term", "query", "time", "cards", "kept"])
        w.writerow([term, query, time.strftime("%Y-%m-%d %H:%M:%S"), cards, kept])


# ---------------- main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profile", default="Profile 26")
    ap.add_argument("--searches", type=int, default=10, help=f"searches this run ({len(TERMS)} terms in the bank)")
    ap.add_argument("--scrolls", type=int, default=6, help="scroll rounds per search (~10 posts each)")
    ap.add_argument("--min-delay", type=float, default=15)
    ap.add_argument("--max-delay", type=float, default=60)
    ap.add_argument("--dry-run", action="store_true", help="print kept posts, don't copy links or write anything")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        return selftest()

    wb, ws = open_sheet()
    known = {r[0] for r in ws.iter_rows(min_row=2, max_col=1, values_only=True) if r[0]}
    print(f"{len(known)} posts already in {POSTS_XLSX.name}")
    sr.prepare_profile(args.profile)
    added, next_break = 0, random.randint(4, 7)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(sr.LOCAL_DATA), channel="chrome", headless=False,
            args=[f"--profile-directory={args.profile}"], no_viewport=True,
            ignore_default_args=["--use-mock-keychain"],
        )
        ctx.grant_permissions(["clipboard-read", "clipboard-write"], origin="https://www.linkedin.com")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://www.linkedin.com/feed/")
        if any(k in page.url for k in ("/login", "/authwall", "/uas/", "/checkpoint/")):
            input("Log in to LinkedIn in the opened window, then press Enter...")
        sr.read_profile(page)

        try:
            for n, (cluster, term) in enumerate(searches_by_staleness()[: args.searches], 1):
                query = f"{random.choice(HIRING_PHRASES)} {term}"
                page.goto(search_url(query), wait_until="domcontentloaded")
                sr.check_session(page)
                try:
                    page.locator('[role=listitem][componentkey^="update-card-focus"]').first.wait_for(timeout=15000)
                except PWTimeout:
                    print(f"[{n}] {query}: no results")
                    if not args.dry_run:
                        log_search(term, query, 0, 0)
                    continue
                sr.pause(2, 4)
                skim_results(page, args.scrolls)
                cards = page.evaluate(CARDS_JS)
                kept = 0
                for c in cards:
                    age = age_minutes(c["stamp"])
                    if not c["text"] or age is None or age >= UNIT_MIN["w"]:  # strictly under 7 days
                        continue
                    pid = post_id(c)
                    if pid in known:
                        continue
                    known.add(pid)
                    score, hire, seek = classify(c["text"], c["otw"])
                    if score < MIN_SCORE:
                        continue
                    kept += 1
                    print(f"   + [{score}] {c['author']} ({c['stamp']}): {c['text'][:90]!r}")
                    if args.dry_run:
                        continue
                    link = copy_link(page, c["key"])
                    posted = (datetime.now() - timedelta(minutes=age)).strftime("%Y-%m-%d %H:%M")
                    ws.append([pid, time.strftime("%Y-%m-%d %H:%M"), posted, c["stamp"].split("•")[0].strip(), c["author"],
                               c["headline"], c["author_url"], link, c["text"][:32000], ", ".join(sorted(set(EMAILS.findall(c["text"])))),
                               score, " | ".join(hire), " | ".join(seek), cluster, query, "", ""])
                    added += 1
                print(f"[{n}] {query}: {len(cards)} cards, {kept} new hiring posts (total added {added})")
                if not args.dry_run:
                    save_sheet(wb)
                    log_search(term, query, len(cards), kept)
                if n == next_break:
                    secs = random.uniform(180, 420)
                    print(f"   taking a {secs / 60:.1f} min break")
                    sr.idle_wait(page, secs, secs)
                    next_break += random.randint(4, 7)
                else:
                    sr.idle_wait(page, args.min_delay, args.max_delay)
        except sr.SessionLost as e:
            print(f"Session lost ({e}); stopping.")
        except Exception as e:
            if not any(k in str(e).lower() for k in sr.SESSION_ERRORS):
                raise
            print(f"Network/session error, stopping: {str(e)[:120]}")
        finally:
            if not args.dry_run:
                save_sheet(wb)
            print(f"\nDone: {added} posts added; {POSTS_XLSX.name} now has {ws.max_row - 1} rows.")
            ctx.close()


def selftest():
    hiring = [
        "We're hiring a Content Writer! Location: Remote. Send your CV to hr@x.com #hiring",
        "Looking for experienced academic content writers. Send me some samples on a@b.com\n#hiring #subjectmatterexperts",
        "HIRING\nCopywriter | IFS Group\nLocation: Lagos\nResponsibilities ... Interested candidates should apply via the link",
        "Our startup is looking for a Founder's Office intern. DM me. Stipend: 20k",
    ]
    seeking = [
        "I'm actively looking for a content writer role. #OpenToWork Please refer me!",
        "Hi all, I was laid off last week and I am seeking new opportunities in growth marketing. Hire me!",
        "Most businesses hire a writer based on a portfolio alone. Here's what to check instead.",
        "What are your salary expectations? Medical writers, here's how I'd prepare.",
    ]
    for t in hiring:
        assert classify(t)[0] >= MIN_SCORE, (t, classify(t))
    for t in seeking:
        assert classify(t)[0] < MIN_SCORE, (t, classify(t))
    assert classify(hiring[0], open_to_work_badge=True)[0] < classify(hiring[0])[0]
    assert age_minutes("33m") == 33 and age_minutes("3d • Edited") == 4320 and age_minutes("1w") == 10080
    assert age_minutes("2mo") == 86400 and age_minutes("Follow") is None
    assert len({t for _, t in TERMS}) == len(TERMS), "duplicate term in ROLE_CLUSTERS"
    print("selftest ok")


if __name__ == "__main__":
    main()
