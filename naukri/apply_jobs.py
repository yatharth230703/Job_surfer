"""Naukri auto-apply bot for Soumya.

    ../linkedin/.venv/bin/python apply_jobs.py --limit 20            # real applications
    ../linkedin/.venv/bin/python apply_jobs.py --limit 3 --dry-run   # finds jobs, hovers Apply, never clicks
    ../linkedin/.venv/bin/python apply_jobs.py --selftest

Uses its own fresh copy of Chrome "Profile 26" (soumya) in naukri/chrome_data (holds the Naukri login),
plus the LinkedIn bot's keyword bank, title filter and answer rules. Each keyword is searched twice:
Delhi NCR and remote, 0-1 yrs experience, posted in the last --days days.

Verified on the live site (18 Nov 2026):
- cards: .srp-jobtuple-wrapper[data-job-id], title a.title, .comp-name, .locWdth
- job page: #apply-button "Apply" (Naukri apply) | #company-site-button (external) | #already-applied
- after Apply: either redirect to /myapply/saveApply ("You have successfully applied"), or a .chatbot_Drawer
  with recruiter questions. The application already counts once the drawer opens (job shows "Applied").
"""
import argparse
import csv
import random
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "linkedin"))
import send_request as sr  # noqa: E402  human mouse/scroll helpers, profile copy, session checks
import apply_jobs as li  # noqa: E402  SEARCH_QUERIES, title_matches, answer, NCR_PLACES

sr.LOCAL_DATA = HERE / "chrome_data"  # Naukri's own profile copy, separate from LinkedIn's
APPLIED_LOG = HERE / "applied_log.csv"
UNANSWERED_LOG = HERE / "unanswered_questions.csv"
RETRY_STATUSES = ("error", "timeout", "no_button", "unclear", "not_applied")
GENDER = li.env("GENDER")  # e.g. Female; unset -> gender questions are left unanswered
DAILY_LIMIT = re.compile(r"daily (apply )?(limit|quota)|reached (the|your) (daily )?(apply|application) limit", re.I)
CREDENTIAL = re.compile(r"certifi|licen[cs]|registered (with|pharmacist|nurse)|membership|member id", re.I)
SUCCESS = re.compile(r"successfully applied", re.I)
DRAWER = ".chatbot_Drawer"

CARDS_JS = r"""() => [...document.querySelectorAll('.srp-jobtuple-wrapper[data-job-id]')].map(c => ({
  id: c.dataset.jobId, title: (c.querySelector('a.title')?.innerText || '').trim(), url: c.querySelector('a.title')?.href,
  company: (c.querySelector('.comp-name')?.innerText || '').trim(), loc: (c.querySelector('.locWdth')?.innerText || '').trim()}))"""


def search_url(query, mode, days):
    slug = re.sub(r"[^a-z0-9]+", "-", query.lower()).strip("-")
    q = quote(query)
    if mode == "ncr":
        return f"https://www.naukri.com/{slug}-jobs-in-delhi-ncr?k={q}&l=delhi%20ncr&experience=1&jobAge={days}"
    return f"https://www.naukri.com/{slug}-jobs?k={q}&experience=1&jobAge={days}&wfhType=2"


def location_ok(loc, mode):
    loc = loc.lower()
    return "remote" in loc or (mode == "ncr" and any(p in loc for p in li.NCR_PLACES))


# ---------------- logs ----------------
def _append(path, header, row):
    new = not path.exists()
    with open(path, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(header)
        w.writerow(row)


def load_done():
    if not APPLIED_LOG.exists():
        return set()
    return {r["job_id"] for r in csv.DictReader(open(APPLIED_LOG)) if not r["status"].startswith(RETRY_STATUSES)}


def log_job(job, status, query):
    _append(APPLIED_LOG, ["job_id", "title", "company", "location", "url", "status", "query", "time"],
            [job["id"], job["title"], job["company"], job["loc"], job["url"], status, query,
             time.strftime("%Y-%m-%d %H:%M:%S")])


# ---------------- search ----------------
def collect(page, url, pages):
    page.goto(url, wait_until="domcontentloaded")
    sr.check_session(page)
    sr.pause(3, 6)
    cards = {}
    for n in range(pages):
        for c in page.evaluate(CARDS_JS):
            if c["url"]:
                cards.setdefault(c["id"], c)
        sr.read_profile(page)  # skim the results
        nxt = page.locator('nav[aria-label="Pagination"] a:has(span:text-is("Next"))')
        if n + 1 >= pages or not nxt.count() or nxt.first.get_attribute("aria-disabled") == "true":
            break
        first = page.evaluate(CARDS_JS)[:1]
        sr.human_click(page, nxt.first)  # click, don't goto: the href drops the filters
        for _ in range(20):
            sr.pause(0.5, 0.8)
            if page.evaluate(CARDS_JS)[:1] != first:
                break
        sr.pause(2, 4)
    return list(cards.values())


# ---------------- chatbot questions ----------------
def chat_options(page, kind):
    boxes = page.locator(f"{DRAWER} input[type={kind}]")
    opts = []
    for i in range(boxes.count()):
        bid = boxes.nth(i).get_attribute("id")
        lab = page.locator(f'{DRAWER} label[for="{bid}"]') if bid else None
        opts.append((lab.first.inner_text().strip() if lab and lab.count() else boxes.nth(i).get_attribute("value") or "",
                     lab.first if lab and lab.count() else boxes.nth(i)))
    return opts


def answer(q, kind, opts=()):
    if re.search(r"\bgender\b", q, re.I):
        if not GENDER:
            return None
        return next((o for o in opts if o.lower() == GENDER.lower()), None) if kind == "radio" else GENDER
    if kind == "radio" and CREDENTIAL.search(q):
        return None  # "Are you CPC certified?" -> never claim a credential she may not have
    return li.answer(q, kind, opts)


def answer_chat(page, job):
    """Answer the recruiter questions in the drawer. Returns None when done, else a reason."""
    last_q, repeats = None, 0
    for _ in range(15):
        sr.pause(2, 3.5)
        if not page.locator(DRAWER).count() or not page.locator(DRAWER).first.is_visible():
            return None
        q = page.locator(f"{DRAWER} li.botItem").last.inner_text().strip()
        repeats = repeats + 1 if q == last_q else 0
        if repeats >= 2:
            return f"stuck on {q[:60]!r}"
        last_q = q
        radios, checks = chat_options(page, "radio"), chat_options(page, "checkbox")
        is_text = False
        chips = page.locator(f"{DRAWER} .chatbot_Chip, {DRAWER} .chipItem")  # ponytail: chip markup not seen live yet
        text = page.locator(f"{DRAWER} .textArea[contenteditable]")
        if radios or checks:
            opts = radios or checks
            a = answer(q, "radio", [o for o, _ in opts])
            target = next((loc for o, loc in opts if o == a), None)
        elif chips.count():
            opts = [chips.nth(i).inner_text().strip() for i in range(chips.count())]
            a = answer(q, "radio", opts)
            target = chips.nth(opts.index(a)) if a in opts else None
        elif text.count() and text.first.is_visible():
            a = answer(q, "text")
            target, is_text = text.first, True
        else:
            return f"unknown input for {q[:60]!r}"
        if target is None or a is None:
            _append(UNANSWERED_LOG, ["question", "options", "job_url", "time"],
                    [q, " | ".join(o for o, _ in (radios or checks)), job["url"], time.strftime("%Y-%m-%d %H:%M:%S")])
            return f"can't answer {q[:60]!r}"
        print(f"     Q {q[:70]!r} -> {a!r}")
        sr.human_click(page, target)
        if is_text:
            sr.pause(0.4, 1)
            page.keyboard.type(a, delay=random.randint(70, 160))
        sr.pause(0.8, 2)
        save = page.locator(f"{DRAWER} .sendMsgbtn_container .send")  # sr.human_click's cover check trips on it
        if save.count() and save.first.is_visible():
            box = save.first.bounding_box()
            sr.move_to(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            sr.pause(0.1, 0.4)
            page.mouse.down()  # locator.click() on .send hangs; a real click at the point works
            time.sleep(random.uniform(0.05, 0.15))
            page.mouse.up()
        for _ in range(30):  # the next question can take 5+ s to appear
            time.sleep(0.5)
            if not page.locator(DRAWER).first.is_visible():
                break
            last = page.locator(f"{DRAWER} li.botItem").last.inner_text().strip()
            if last and last != q:  # a "typing" bubble is empty, so wait for real text
                break
    return "too many questions"


def close_drawer(page):
    x = page.locator(f"{DRAWER} .crossIcon")
    if x.count() and x.first.is_visible():
        x.first.click()
        sr.pause(1, 2)


# ---------------- apply ----------------
def apply_job(page, job, dry_run):
    page.goto(job["url"], wait_until="domcontentloaded")
    sr.check_session(page)
    sr.pause(3, 6)
    if page.locator("#already-applied").count():
        return "already_applied"
    if page.locator("#company-site-button").count():
        return "external"
    btn = sr.first_visible(page, '#apply-button:text-is("Apply")')
    if not btn:
        return "no_button"
    sr.read_profile(page)  # skim the description
    if dry_run:
        sr.wheel_into_view(page, btn)
        box = btn.bounding_box()
        sr.move_to(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        sr.pause(1, 2)
        return "dry_run_ok"
    sr.human_click(page, btn)
    for _ in range(20):
        sr.pause(0.8, 1.2)
        body = page.evaluate("document.body.innerText")
        if DAILY_LIMIT.search(body):
            raise DailyLimit()
        if "/myapply/" in page.url or SUCCESS.search(body):
            return "applied"
        if page.locator(DRAWER).count() and page.locator(DRAWER).first.is_visible():
            problem = answer_chat(page, job)
            if problem:
                print(f"     questions left: {problem}")
                close_drawer(page)
            sr.pause(2, 4)
            # truth is the job page's "Applied" badge (closing the drawer with X usually still applies)
            page.goto(job["url"], wait_until="domcontentloaded")
            page.locator("#already-applied, #apply-button, #company-site-button").first.wait_for(timeout=20000)
            if not page.locator("#already-applied").count():
                return "not_applied"
            return "applied_questions_pending" if problem else "applied"
    return "applied" if page.locator("#already-applied").count() else "unclear"


class DailyLimit(Exception):
    pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profile", default="Profile 26")
    ap.add_argument("--limit", type=int, default=20, help="max applications this run (Naukri caps ~50/day)")
    ap.add_argument("--pages", type=int, default=2, help="result pages per search (20 jobs each)")
    ap.add_argument("--days", type=int, default=7, help="only jobs posted in the last N days")
    ap.add_argument("--min-delay", type=float, default=30)
    ap.add_argument("--max-delay", type=float, default=120)
    ap.add_argument("--dry-run", action="store_true", help="find jobs and hover Apply, never click it")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        return selftest()

    done = load_done()
    searches = [(q, m) for q in li.SEARCH_QUERIES for m in ("ncr", "remote")]
    random.shuffle(searches)
    sr.prepare_profile(args.profile)
    applied, next_break = 0, random.randint(4, 7)
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(sr.LOCAL_DATA), channel="chrome", headless=False,
            args=[f"--profile-directory={args.profile}"], no_viewport=True,
            ignore_default_args=["--use-mock-keychain"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://www.naukri.com/mnjuser/homepage", wait_until="domcontentloaded")
        sr.pause(3, 5)
        if "login" in page.url:
            input("Log in to Naukri in the opened window, then press Enter...")
        try:
            for query, mode in searches:
                if applied >= args.limit:
                    break
                cards = collect(page, search_url(query, mode, args.days), args.pages)
                picks = [c for c in cards if c["id"] not in done and li.title_matches(c["title"])
                         and location_ok(c["loc"], mode)]
                print(f"[{query} / {mode}] {len(cards)} jobs, {len(picks)} to try")
                for job in picks:
                    print(f"  -> {job['title']} | {job['company']} | {job['loc']}")
                    try:
                        status = apply_job(page, job, args.dry_run)
                    except (sr.SessionLost, DailyLimit):
                        raise
                    except PWTimeout:
                        status = "timeout"
                    except Exception as e:
                        if any(k in str(e).lower() for k in sr.SESSION_ERRORS):
                            raise sr.SessionLost(str(e)[:120])
                        status = f"error: {e.__class__.__name__}: {str(e)[:80]}"
                    for extra in ctx.pages[1:]:
                        extra.close()
                    print(f"     {status}")
                    if not args.dry_run:
                        log_job(job, status, query)
                    done.add(job["id"])
                    if status.startswith("applied") or status == "dry_run_ok":
                        applied += 1
                        if applied >= args.limit:
                            break
                        if applied == next_break:
                            secs = random.uniform(180, 480)
                            print(f"   taking a {secs / 60:.1f} min break")
                            sr.idle_wait(page, secs, secs)
                            next_break += random.randint(4, 7)
                        else:
                            sr.idle_wait(page, args.min_delay, args.max_delay)
                    else:
                        sr.pause(3, 10)
                sr.pause(4, 12)
        except DailyLimit:
            print("Naukri's daily apply limit reached; stopping.")
        except sr.SessionLost as e:
            print(f"Session lost ({e}); stopping.")
        finally:
            print(f"\nDone: {applied} {'dry-run jobs' if args.dry_run else 'applications'} this run.")
            ctx.close()


def selftest():
    assert search_url("Founder's Office", "ncr", 7) == ("https://www.naukri.com/founder-s-office-jobs-in-delhi-ncr"
                                                        "?k=Founder%27s%20Office&l=delhi%20ncr&experience=1&jobAge=7")
    assert search_url("Content Writer", "remote", 3).endswith("content-writer-jobs?k=Content%20Writer&experience=1&jobAge=3&wfhType=2")
    assert location_ok("Noida(Sector 8)", "ncr") and location_ok("Remote", "remote")
    assert not location_ok("Bengaluru", "ncr") and not location_ok("Gurugram", "remote")
    assert answer("Are you CPC certified?", "radio", ["Yes", "No"]) is None
    assert answer("Do you hold a valid drug license?", "radio", ["Yes", "No"]) is None
    assert answer("Are you comfortable with field sales?", "radio", ["Yes", "No"]) == "Yes"
    assert answer("What is your gender identity?", "radio", ["Male", "Female"]) in (None, "Male", "Female")
    assert li.answer("How much experience do you have in Business Development?", "text") == "1"
    assert li.answer("Are you comfortable with field work?", "radio", ["Yes", "No"]) == "Yes"
    assert DAILY_LIMIT.search("You have reached your daily apply limit") and SUCCESS.search("You have successfully applied to")
    print("selftest ok")


if __name__ == "__main__":
    main()
