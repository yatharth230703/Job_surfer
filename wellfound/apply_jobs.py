"""Wellfound auto-apply bot for Soumya.

    ../linkedin/.venv/bin/python apply_jobs.py --limit 15             # real applications
    ../linkedin/.venv/bin/python apply_jobs.py --limit 3 --dry-run    # fills the form, never sends
    ../linkedin/.venv/bin/python apply_jobs.py --selftest

Reuses the LinkedIn bot's Chrome profile copy (linkedin/chrome_data, Profile 26), human mouse helpers (sr),
keyword bank, title filter and rule-based answers. Only free-text questions the rules can't answer
("What interests you about working for this company?") go to Gemini, with the resume PDF and job description.
Search: each keyword is typed into Wellfound's role box as a free-text term; location/remote filters are whatever
Soumya's saved Wellfound search already has. Off-platform jobs are hidden via Wellfound's own checkbox.
"""
import argparse
import base64
import csv
import json
import random
import re
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "linkedin"))
import send_request as sr  # noqa: E402  human mouse/scroll helpers, profile copy
import apply_jobs as li  # noqa: E402  env(), SEARCH_QUERIES, title_matches, answer(), resume_for

APPLIED_LOG = HERE / "applied_log.csv"
ANSWERS_LOG = HERE / "answers_log.csv"  # every question + the answer sent, so the AI notes can be reviewed
RETRY_STATUSES = ("error", "timeout", "unknown_question", "form_error", "no_button", "ai_error")
GEMINI_MODELS = ("gemini-2.5-flash", "gemini-flash-latest", "gemini-flash-lite-latest")
MODAL = '[data-test="JobApplication-Modal"]'
SUBMIT = '[data-test="JobApplicationModal--SubmitButton"]'

JOB_LINKS_JS = r"""() => [...document.querySelectorAll('a[href^="/jobs/"]')]
  .map(a => [a.getAttribute('href'), a.innerText.split('\n')[0].trim()])
  .filter(([h]) => /^\/jobs\/\d+/.test(h))"""

FIELDS_JS = r"""() => {
  const m = document.querySelector('[data-test="JobApplication-Modal"] form'); if (!m) return [];
  return [...m.querySelectorAll('textarea, input, select')].filter(c => c.type !== 'hidden' && c.offsetParent).map(c => {
    const lab = c.closest('label') || (c.id && m.querySelector(`label[for="${CSS.escape(c.id)}"]`)) || c.closest('fieldset') || c.parentElement;
    let q = (lab.innerText || '').replace(/\d+\s*\/\s*\d+\s*$/, '').trim();
    if (!q) {  // the default "userNote" box has no label: "Your hiring contact is X / Let them know why you are a good fit."
      const h = m.querySelector('h4');
      q = [h && h.parentElement.innerText.replace(/\s+/g, ' '), c.placeholder].filter(Boolean).join(' ').trim();
    }
    const kind = c.tagName === 'TEXTAREA' ? 'textarea' : c.tagName === 'SELECT' ? 'select' : c.type;
    const opts = c.tagName === 'SELECT' ? [...c.options].map(o => o.text.trim()) : [];
    return {name: c.name, id: c.id, q, kind, opts, value: c.value, required: c.required};
  });
}"""


def gemini(question, job, desc):
    """One short, honest first-person answer written from the resume + job description."""
    pdf = HERE.parent / li.resume_for(job["title"])
    prompt = (f"You are filling a job application on Wellfound for the candidate whose resume is attached.\n"
              f"Role: {job['title']} at {job['company']}\nJob description:\n{desc[:6000]}\n\n"
              f"Application question: {question}\n\n"
              "Write the candidate's answer: first person, 70-130 words, plain text, no greeting or sign-off, "
              "no placeholders or brackets. Be specific to this company and role, link 1-2 concrete things from the "
              "resume to what they need. Never invent experience, numbers or skills not in the resume; if the "
              "candidate lacks something, show eagerness to learn instead. Output only the answer.")
    body = {"contents": [{"parts": [{"inline_data": {"mime_type": "application/pdf",
                                                     "data": base64.b64encode(pdf.read_bytes()).decode()}},
                                    {"text": prompt}]}]}
    for attempt in range(2):
        for model in GEMINI_MODELS:  # 503 "high demand" is common; fall through to the next model
            req = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                data=json.dumps(body).encode(), headers={"Content-Type": "application/json",
                                                         "x-goog-api-key": li.env("GEMINI_API_KEY")})
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    parts = json.load(r)["candidates"][0]["content"]["parts"]
                text = "".join(p.get("text", "") for p in parts if not p.get("thought")).strip()
                if text:
                    return text
            except Exception as e:
                print(f"     gemini {model} failed: {e}")
        time.sleep(20)
    return None


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
    _append(APPLIED_LOG, ["job_id", "title", "company", "url", "status", "query", "time"],
            [job["id"], job["title"], job["company"], job["url"], status, query, time.strftime("%Y-%m-%d %H:%M:%S")])


# ---------------- page actions ----------------
def search(page, keyword):
    """Replace the free-text role terms with `keyword`; keeps Soumya's saved location/remote filters."""
    page.goto("https://wellfound.com/jobs", wait_until="domcontentloaded")
    sr.check_session(page)
    page.locator('[data-test="SearchBar"]').wait_for(timeout=20000)
    sr.pause(2, 4)
    off = page.locator("#showOffPlatformJobs")
    if off.count() and not off.is_checked():  # hide "apply on company website" jobs
        sr.human_click(page, page.locator('label[for="showOffPlatformJobs"]'))
        sr.pause(2, 4)
    sr.human_click(page, page.locator('[data-test="SearchBar-RoleSelect-FocusButton"]'))
    sr.pause(0.5, 1.2)
    for _ in range(8):  # remove existing role chips
        page.keyboard.press("Backspace")
        time.sleep(random.uniform(0.08, 0.2))
    page.keyboard.type(keyword, delay=random.randint(60, 140))
    sr.pause(0.8, 1.6)
    page.keyboard.press("Enter")
    page.keyboard.press("Escape")
    sr.pause(4, 7)


def collect_jobs(page, rounds):
    """Scroll the (infinite) results list and return {url: title}."""
    found = {}
    for _ in range(rounds):
        for href, title in page.evaluate(JOB_LINKS_JS):
            found.setdefault("https://wellfound.com" + href.split("?")[0], title)
        for _ in range(random.randint(4, 9)):
            page.mouse.wheel(0, random.uniform(150, 350))
            time.sleep(random.uniform(0.03, 0.1))
        sr.pause(1.5, 3.5)
    return found


def job_description(page):
    body = page.evaluate("document.body.innerText")
    return body.split("Similar jobs to")[0][-8000:]


def fill_form(page, job, desc, dry_run):
    """Answer every visible field. Returns None if all filled, else a status."""
    answers = []
    for f in page.evaluate(FIELDS_JS):
        if f["value"].strip() and f["kind"] != "select":
            continue  # prefilled by Wellfound
        sel = f'{MODAL} [name="{f["name"]}"]' if f["name"] else f'{MODAL} [id="{f["id"]}"]'
        loc = page.locator(sel).first
        kind = {"textarea": "text", "number": "text", "email": "text", "tel": "text", "url": "text"}.get(f["kind"], f["kind"])
        ans, src = li.answer(f["q"], kind, f["opts"]), "rule"
        if ans is None and f["kind"] in ("textarea", "text"):
            ans, src = gemini(f["q"], job, desc), "gemini"
            if ans is None:
                return "ai_error"
        if ans is None:
            if not f["required"] and f["kind"] in ("checkbox", "radio"):
                continue
            li.bank_record(f["q"], f["kind"], f["opts"], "Wellfound", job["url"])
            print(f"     can't answer: {f['q'][:80]!r} ({f['kind']})")
            return "unknown_question"
        print(f"     [{src}] {f['q'][:60]!r} -> {ans[:70]!r}{'...' if len(ans) > 70 else ''}")
        if f["kind"] == "select":
            loc.select_option(label=ans)
        else:
            sr.human_click(page, loc)
            sr.pause(0.5, 1.5)
            page.keyboard.type(ans, delay=random.randint(15, 40))
        sr.pause(1, 3)
        answers.append((f["q"], ans, src))
    for q, a, src in answers:
        _append(ANSWERS_LOG, ["job_url", "title", "company", "question", "answer", "source", "dry_run", "time"],
                [job["url"], job["title"], job["company"], q, a, src, dry_run, time.strftime("%Y-%m-%d %H:%M:%S")])
    return None


def close_modal(page):
    btn = page.locator(f'{MODAL} button:text-is("Cancel")')
    if btn.count():
        try:
            btn.first.click(timeout=3000)
        except Exception:
            page.keyboard.press("Escape")
    sr.pause(1, 2)


def apply_job(page, job, dry_run):
    page.goto(job["url"], wait_until="domcontentloaded")
    sr.check_session(page)
    sr.pause(3, 6)
    body = page.evaluate("document.body.innerText")
    job["company"] = (re.search(r"^(.+?) at (.+)$", next((l for l in body.splitlines() if l.startswith(job["title"] + " at ")), "")) or [None, "", ""])[2]
    desc = job_description(page)
    sr.read_profile(page)  # skim the description
    btn = sr.first_visible(page, ':is(button, a):text-is("Apply")') or sr.first_visible(page, ':is(button, a):text-is("Apply now")')
    if not btn:
        return "already_applied" if re.search(r"\bApplied\b", body) else "no_button"
    sr.human_click(page, btn)
    try:
        page.locator(MODAL).wait_for(timeout=15000)
    except PWTimeout:
        return "external" if len(page.context.pages) > 1 else "no_modal"
    sr.pause(2, 4)
    status = fill_form(page, job, desc, dry_run)
    if status:
        close_modal(page)
        return status
    send = page.locator(SUBMIT).first
    if dry_run:
        sr.wheel_into_view(page, send)
        box = send.bounding_box()
        sr.move_to(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        sr.pause(1, 2)
        close_modal(page)
        return "dry_run_ok"
    sr.human_click(page, send)
    for _ in range(20):  # after Send: modal closes, or it turns into the optional "AI interview" upsell
        sr.pause(1, 1.2)
        if not page.locator(MODAL).count():
            return "applied"
        txt = page.locator(MODAL).inner_text()
        if re.search(r"Start the Interview|After this application", txt):
            page.keyboard.press("Escape")  # skip the optional interview
            sr.pause(1, 2)
            return "applied"
        if not page.locator(SUBMIT).count():
            break
    # anything else: Send was clicked, so never retry automatically (a retry could apply twice)
    print("     unclear state after Send:", page.locator(MODAL).inner_text()[-300:].replace("\n", " | ") if page.locator(MODAL).count() else "")
    page.keyboard.press("Escape")
    return "sent_unverified"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profile", default="Profile 26")
    ap.add_argument("--limit", type=int, default=15, help="max applications this run")
    ap.add_argument("--min-delay", type=float, default=30)
    ap.add_argument("--max-delay", type=float, default=120)
    ap.add_argument("--scrolls", type=int, default=4, help="scroll rounds per search (more = more jobs)")
    ap.add_argument("--dry-run", action="store_true", help="fill the form, hover 'Send application', cancel")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        return selftest()
    if not li.env("GEMINI_API_KEY"):
        raise SystemExit("set GEMINI_API_KEY in the project-root .env")

    done, seen = load_done(), set()
    queries = li.SEARCH_QUERIES[:]
    random.shuffle(queries)
    sr.prepare_profile(args.profile)
    applied, next_break = 0, random.randint(4, 7)
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(sr.LOCAL_DATA), channel="chrome", headless=False,
            args=[f"--profile-directory={args.profile}"], no_viewport=True,
            ignore_default_args=["--use-mock-keychain"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://wellfound.com/jobs", wait_until="domcontentloaded")
        sr.pause(3, 5)
        if "/login" in page.url or page.locator('a:text-is("Log in")').count():
            input("Log in to Wellfound in the opened window, then press Enter...")
        try:
            for query in queries:
                if applied >= args.limit:
                    break
                search(page, query)
                jobs = collect_jobs(page, args.scrolls)
                picks = [(u, t) for u, t in jobs.items() if u not in seen and li.title_matches(t)
                         and re.search(r"/jobs/(\d+)", u).group(1) not in done]
                seen.update(jobs)
                print(f"[{query}] {len(jobs)} jobs, {len(picks)} to try")
                for url, title in picks:
                    if applied >= args.limit:
                        break
                    job = {"id": re.search(r"/jobs/(\d+)", url).group(1), "url": url, "title": title, "company": ""}
                    print(f"  -> {title}  {url}")
                    try:
                        status = apply_job(page, job, args.dry_run)
                    except sr.SessionLost:
                        raise
                    except PWTimeout:
                        status = "timeout"
                    except Exception as e:
                        if any(k in str(e).lower() for k in sr.SESSION_ERRORS):
                            raise sr.SessionLost(str(e)[:120])
                        status = f"error: {e.__class__.__name__}: {str(e)[:80]}"
                    for extra in ctx.pages[1:]:  # company-site tabs opened by "Apply"
                        extra.close()
                    print(f"     {status}")
                    if not args.dry_run:
                        log_job(job, status, query)
                    done.add(job["id"])
                    if status in ("applied", "dry_run_ok"):
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
        except sr.SessionLost as e:
            print(f"Session lost ({e}); stopping.")
        finally:
            print(f"\nDone: {applied} {'dry-run forms completed' if args.dry_run else 'applications'} this run.")
            ctx.close()


def selftest():
    assert li.title_matches("Content Writer") and not li.title_matches("Senior Software Engineer")
    assert li.answer("What interests you about working for this company?", "text") is None  # -> goes to Gemini
    assert li.answer("How many years of experience do you have?", "text") == "1"  # -> rule, no AI
    job = {"title": "Content Writer", "company": "Acme Health", "url": ""}
    a = gemini("What interests you about working for this company?", job,
               "Acme Health builds patient-education content for hospitals in India. You will write blogs and social posts.")
    assert a and 30 < len(a.split()) < 200 and "[" not in a, a
    print(a)
    print("selftest ok")


if __name__ == "__main__":
    main()
