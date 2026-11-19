"""LinkedIn Easy Apply bot for Soumya.

    .venv/bin/python apply_jobs.py --profile "Profile 26" --limit 20            # real applications
    .venv/bin/python apply_jobs.py --profile "Profile 26" --limit 5 --dry-run   # fills forms, never submits

Searches every query in SEARCH_QUERIES twice: Delhi NCR (any workplace type) and worldwide remote.
Only Easy Apply jobs posted in the past week whose title passes the keyword bank are opened.
Forms: years of experience -> 1, every Yes/No -> Yes, current CTC 0, notice 0, expected CTC 3.5 LPA.
Any question the rules can't handle -> application discarded, question added to ../question_bank.xlsx.
Fill in its answer there and every bot (LinkedIn, Wellfound, Naukri) uses it from then on.
"""
import argparse
import base64
import contextlib
import csv
import fcntl
import json
import os
import random
import re
import time
import urllib.parse as up
import urllib.request
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

import send_request as sr  # human mouse/scroll helpers, profile copy, session checks

HERE = Path(__file__).parent


def env(key):
    """KEY=value from the project-root .env (gitignored), falling back to the real environment."""
    f = HERE.parent / ".env"
    for line in (f.read_text().splitlines() if f.exists() else []):
        k, _, v = line.partition("=")
        if k.strip() == key:
            return v.strip().strip("'\"")
    return os.environ.get(key, "")


PHONE = env("PHONE")  # no phone set -> phone questions count as unknown and the application is skipped
APPLIED_LOG = HERE / "applied_log.csv"
QUESTION_BANK = HERE.parent / "question_bank.xlsx"  # shared by all bots; Soumya fills the answer column by hand
JOBS_XLSX = HERE / "applied_jobs.xlsx"  # one row per submitted application, with the full description
EXTERNAL_XLSX = HERE / "external_jobs.xlsx"  # matching jobs without Easy Apply, to apply to by hand
DAILY_LIMIT = re.compile(r"reached (the |today.s )?(daily )?(Easy Apply|application) limit", re.I)
RETRY_STATUSES = ("error", "timeout", "unknown_question", "form_error", "stuck", "no_button")

NCR_GEO, WORLD_GEO = "90009626", "92000000"
NCR_PLACES = ("delhi", "gurugram", "gurgaon", "noida", "faridabad", "ghaziabad", "gautam buddha nagar",
              "shahdara", "shahadara", "dwarka", "sonipat", "bahadurgarh", "ncr")

# ---------------- keyword bank ----------------
SEARCH_QUERIES = [
    # Pharma, hospital & core healthcare
    "Pharmacovigilance Associate", "Drug Safety Trainee", "Drug Safety Associate", "QA Trainee", "QA Executive pharma",
    "Quality Assurance Executive", "QC Trainee", "QC Executive", "Quality Control Executive", "Hospital Operations Executive",
    "Clinical Operations Assistant", "Medical Records Executive", "Pharmacy Operations Executive", "Regulatory Affairs Trainee",
    "Regulatory Affairs Associate", "Production Executive Trainee", "Production Trainee pharma", "Healthcare Operations Associate",
    "Pharmacist", "Clinical Research Associate", "Clinical Research Coordinator", "Medical Coding",
    # HealthTech & business development
    "Customer Success Associate HealthTech", "Customer Success Associate", "Business Development Associate HealthTech",
    "Business Development Associate", "Business Development Executive healthcare", "Clinical Data Management Associate",
    "Clinical Data Coordinator", "Implementation Associate healthcare software", "Healthcare Business Analyst",
    "Product Operations Associate", "Healthcare Market Research Associate", "Partnerships Associate HealthTech",
    "Clinical Informatics Associate", "Healthcare Partnerships Associate", "Hospital Partnerships Executive",
    "Institutional Sales Executive", "Key Account Executive Healthcare", "Strategic Partnerships Associate",
    "Healthcare Growth Associate", "Medical Business Development Executive", "Provider Relations Associate",
    "Corporate Relations Executive Healthcare", "Medical Representative", "Pharma Sales Executive",
    # Consulting
    "Healthcare Consulting Analyst", "Life Sciences Analyst", "Healthcare Research Associate", "Healthcare Strategy Analyst",
    "Pharma Consulting Analyst", "Healthcare Advisory Associate", "Healthcare Operations Analyst",
    "Hospital Management Consultant", "Process Improvement Analyst", "Healthcare Transformation Analyst",
    "Clinical Operations Consultant", "Digital Health Consultant", "Healthcare IT Business Analyst",
    "Healthcare Implementation Consultant", "Health Informatics Analyst", "Technology Consulting Analyst Healthcare",
    "Product Strategy Analyst", "Business Analyst fresher",
    # Freelance roles & services
    "Virtual Assistant", "Social Media Assistant", "Brand Outreach Assistant", "Content Researcher", "Canva Graphic Designer",
    "Email Marketing Assistant", "Copywriting Assistant", "Presentation Designer", "Excel Spreadsheet Specialist",
    "Research Report Assistant", "Documentation and Operations Assistant", "Recruitment Sourcing Assistant",
    "Podcast Research Assistant", "E-commerce Virtual Assistant", "Online Course Assistant", "Project Management Assistant",
    "Email Outreach Specialist", "Website Content Assistant", "Data Analyst Freelancer", "Data Entry Operator",
    "Operations Assistant", "Administrative Assistant",
    # Growth
    "Growth Associate", "Growth Executive", "Growth Analyst", "Growth Intern", "Growth Marketing Associate",
    "Growth Operations Associate", "Growth Hacker", "Performance Marketing Associate", "Digital Marketing Executive",
    "Digital Marketing Associate", "User Acquisition Associate", "Marketing Associate", "Growth Associate HealthTech",
    # Founder's office
    "Founder's Office", "Founders Office Associate", "Founder's Office Intern", "Founder's Office Analyst",
    "Chief of Staff", "Associate Chief of Staff", "Office of the CEO", "CEO Office Associate",
    "Strategy and Operations Associate", "Special Projects Associate", "Business Operations Associate",
    # Content creation & management
    "Content Writer", "Content Creator", "Content Strategist", "Content Marketing Associate", "Content Marketing Executive",
    "Content Executive", "Content Associate", "Content Manager", "Content Operations Associate", "Content Editor",
    "Copywriter", "Social Media Executive", "Social Media Associate", "Social Media Manager", "Community Associate",
    "Community Manager", "Influencer Marketing Associate", "UGC Content Creator", "Scriptwriter", "Blog Writer",
    "SEO Content Writer", "Medical Writer", "Medical Content Writer", "Scientific Writer", "Healthcare Content Writer",
    "Pharma Content Writer", "Communications Associate", "Public Relations Executive",
]

# A title must contain one of these (word-start match, case-insensitive)...
TITLE_KEYWORDS = [
    # domain
    "pharma", "pharmacy", "pharmacist", "pharmacovigilance", "drug safety", "drug", "clinical", "hospital", "health",
    "healthcare", "healthtech", "medical", "medic", "patient", "life science", "biotech", "regulatory", "gmp",
    "quality assurance", "quality control", "quality", "qa", "qc", "production", "manufacturing", "compliance",
    "informatics", "medical records", "medical coding", "diagnostic", "wellness", "nutrition",
    # business / ops / consulting
    "customer success", "client success", "customer experience", "client servic", "account executive", "key account",
    "business development", "bde", "bda", "partnership", "alliances", "institutional sales", "sales executive",
    "sales associate", "inside sales", "growth", "provider relations", "corporate relations", "relationship",
    "implementation", "onboarding", "business analyst", "operations", "ops", "product operations", "process",
    "market research", "research", "strategy", "strategic", "consult", "advisory", "analyst", "transformation",
    "digital health", "project coordinator", "project management", "program coordinator",
    # growth
    "growth", "performance marketing", "digital marketing", "user acquisition", "acquisition", "retention", "gtm",
    "go-to-market", "demand gen", "marketing associate", "marketing executive", "marketing coordinator",
    # founder's office
    "founder", "chief of staff", "office of the ceo", "ceo office", "ceo's office", "special projects", "strategy",
    "business operations",
    # content creation / management
    "content", "creator", "writer", "writing", "copywrit", "editor", "editorial", "blog", "seo", "ugc", "reels",
    "instagram", "youtube", "linkedin", "community", "influencer", "storytell", "script", "communications",
    "public relations", "medical writ", "scientific writ", "newsletter",
    # freelance / support
    "virtual assistant", "executive assistant", "administrative", "admin", "assistant", "social media", "brand",
    "outreach", "content", "canva", "graphic design", "designer", "email marketing", "marketing", "copywrit",
    "presentation", "excel", "spreadsheet", "documentation", "data entry", "data analyst", "data", "recruit",
    "sourcing", "talent acquisition", "hr", "podcast", "e-commerce", "ecommerce", "course", "website",
]  # no bare level words (associate/intern/executive): they let unrelated titles like "Legal Associate" through
# ...and none of these.
TITLE_EXCLUDE = [
    "senior", "sr", "lead", "manager", "head", "director", "principal", "vp", "vice president", "chief", "architect",
    "engineer", "developer", "software", "python", "java", "devops", "sde", "full stack", "frontend", "backend",
    "doctor", "physician", "mbbs", "md", "surgeon", "dentist", "nurse", "nursing", "physiotherap", "radiolog",
    "teacher", "tutor", "faculty", "professor", "driver", "security guard", "telecaller", "accountant", "chartered",
    "lawyer", "attorney", "legal", "law", "quantitative", "quant", "politic", "geopolitic", "scientist",
    "co-founder", "cofounder", "video editor", "motion", "animator", "3d",
    "10+", "8+", "5+",
]
# Junior titles that contain an otherwise-excluded word (manager / chief). Still rejected if senior/lead/head etc. remain.
TITLE_ALLOW = re.compile(r"chief of staff|(content|social media|community|content marketing) manager", re.I)
_kw = lambda words: re.compile(r"(?<![a-z0-9])(" + "|".join(re.escape(w) for w in words) + r")", re.I)
TITLE_OK, TITLE_BAD = _kw(TITLE_KEYWORDS), _kw(TITLE_EXCLUDE)

# Content-creation roles get resume_content.pdf; everything else gets resume.pdf (both already saved on LinkedIn).
CONTENT_TITLE = _kw(["content", "creator", "writer", "writing", "copywrit", "editor", "editorial", "blog", "seo",
                     "ugc", "reels", "instagram", "youtube", "social media", "community", "influencer", "storytell",
                     "script", "newsletter", "communications", "public relations"])
resume_for = lambda title: "resume_content.pdf" if CONTENT_TITLE.search(title) else "resume.pdf"


def title_matches(title):
    return bool(TITLE_OK.search(title)) and not TITLE_BAD.search(TITLE_ALLOW.sub(" ", title))


def location_ok(loc, mode):
    loc = loc.lower()
    if "(remote)" in loc:
        return True  # remote is accepted from either search
    return mode == "ncr" and any(p in loc for p in NCR_PLACES)


# ---------------- answer bank ----------------
# ---------------- question bank (question_bank.xlsx) ----------------
BANK_COLS = [("Question", 70), ("Answer (fill in)", 40), ("Type", 10), ("Options", 50), ("Site", 10),
             ("Times seen", 11), ("Last seen", 17), ("Example job", 50), ("Answered by", 12)]
AI_SHEET = "AI per-job answers"  # job-specific AI answers: logged for review, never reused
AI_COLS = [("Time", 17), ("Site", 10), ("Job", 40), ("Question", 60), ("Answer", 80), ("Job URL", 50)]
_bank = {"mtime": None, "answers": {}}


def norm_q(q):
    """Same question despite '*', '0/20' counters, punctuation, case or spacing."""
    q = re.sub(r"\d+\s*(/|of)\s*\d+(\s*characters)?\s*$", "", q.strip(), flags=re.I)
    return " ".join(re.sub(r"[^a-z0-9]+", " ", q.lower()).split())


def _bank_answers():
    """{normalised question: answer} for filled-in rows; re-read whenever the file changes (edit it mid-run)."""
    try:
        mtime = QUESTION_BANK.stat().st_mtime
    except FileNotFoundError:
        return {}
    if mtime != _bank["mtime"]:
        from openpyxl import load_workbook
        try:
            rows = load_workbook(QUESTION_BANK, read_only=True).worksheets[0].iter_rows(min_row=2, max_col=2, values_only=True)
            _bank["answers"] = {norm_q(str(q)): str(a).strip() for q, a in rows if q and a is not None and str(a).strip()}
            _bank["mtime"] = mtime
        except Exception as e:  # half-saved / locked file: keep the last good copy
            print(f"     (couldn't read {QUESTION_BANK.name}: {e})")
    return _bank["answers"]


def bank_answer(q, opts=()):
    """The hand-written answer for q, matched to one of opts when the field has options."""
    a = _bank_answers().get(norm_q(q))
    if a is None or not opts:
        return a
    low = [o.lower().strip() for o in opts]
    if a.lower() in low:
        return opts[low.index(a.lower())]
    starts = [o for o, l in zip(opts, low) if re.match(re.escape(a.lower()) + r"(?![a-z0-9])", l)]  # "1" -> "1 year", not "10"
    return starts[0] if len(starts) == 1 else None  # answer doesn't fit these options -> treated as unknown


@contextlib.contextmanager
def _bank_lock():
    """Bots run in parallel (LinkedIn + Naukri): serialise load-modify-save so no row is lost."""
    with open(QUESTION_BANK.with_suffix(".lock"), "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        yield


def _sheet(wb, title, cols):
    from openpyxl.styles import Font
    ws = wb[title] if title in wb.sheetnames else wb.create_sheet(title)
    ws.freeze_panes = "A2"
    for i, (c, w) in enumerate(cols, 1):  # also adds columns introduced later (e.g. "Answered by")
        if not ws.cell(1, i).value:
            ws.cell(1, i).value = c
        ws.cell(1, i).font = Font(bold=True)
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    return ws


def bank_record(q, kind, opts=(), site="", job_url="", answer=None, source=""):
    """Add a question to the bank (blank answer = for Soumya to fill), or bump its count if it's already there."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font
    opts = " | ".join(o for o in opts if o and not re.match(r"select( an)? option", o, re.I))
    now = time.strftime("%Y-%m-%d %H:%M")
    try:
        with _bank_lock():
            if QUESTION_BANK.exists():
                wb = load_workbook(QUESTION_BANK)
            else:
                wb = Workbook(); wb.active.title = "Questions"
            ws = _sheet(wb, wb.worksheets[0].title, BANK_COLS)
            key = norm_q(q)
            for row in ws.iter_rows(min_row=2):
                if row[0].value and norm_q(str(row[0].value)) == key:
                    row[5].value = (row[5].value or 0) + 1
                    row[6].value = now
                    if opts and not row[3].value:
                        row[3].value = opts
                    if answer and not row[1].value:
                        ws.cell(row[0].row, 2).value = answer
                        ws.cell(row[0].row, 9).value = source
                    break
            else:
                ws.append([q.strip(), answer, kind, opts, site, 1, now, job_url, source or None])
                for c in ws[ws.max_row]:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            wb.save(QUESTION_BANK)
    except Exception as e:  # e.g. file locked; never kill a run over the spreadsheet
        print(f"     (couldn't write {QUESTION_BANK.name}: {e})")


# ---------------- Gemini: first sight of a question the bank and rules can't answer ----------------
GEMINI_MODELS = ("gemini-2.5-flash", "gemini-flash-latest", "gemini-flash-lite-latest")
CREDENTIAL = re.compile(r"certifi|licen[cs]|registered (with|pharmacist|nurse)|membership|member id", re.I)
# Asking applicants to pay is a classic scam: never answered, so LinkedIn discards the form and it lands in the bank.
FEE = re.compile(r"\b(fees?|deposit|refundable|registration (charge|amount)|security (amount|money)|training charges?"
                 r"|pay (an?|the|any|some) (amount|fee|charge))\b", re.I)
_cooldown = {}  # model -> time it returned 429; skipped for 10 min instead of waiting on it every question
AI_PROMPT = """You are filling in a job application form for the candidate whose resume is attached.
Job: {title} at {company}
Job description:
{desc}

Form question: {q}
Field type: {kind}
{opts}
Rules:
- Facts (degrees, certifications, licences, tools, skills, employers, sales or revenue numbers, years of experience in a
  specific skill) come ONLY from the resume. If the resume doesn't show it, answer honestly negative ("No", "0", or the
  lowest option). Never invent experience, credentials or numbers. Years in a specific skill: "1" if the resume shows
  internship or project work in it, else "0".
- The job description describes the employer, never the candidate. The candidate has no sales history: questions about
  what she sells, her clients, revenue, targets, deals or accounts managed get "None", "0", the lowest option or a
  "not applicable / skip" option. Never copy such facts from the job description.
- You may infer obvious facts from the resume (e.g. the state from the city).
- Preferences and availability (relocation, shifts, travel, field work, working days, commute, joining date, staying on
  after an internship): the candidate is flexible -> "Yes" or the most flexible option.
- Personal details the resume doesn't give and you can't infer (gender, date of birth, ID or member numbers, links,
  references, reasons for gaps): answer null.
- With options, the answer must be exactly one of the option texts. A number question ("how many", "rate out of 10")
  gets digits only. Short factual questions get a few words. Open questions ("why", "describe", "tell us") get a first-
  person answer of 50-100 words (or the length the question asks for): plain text, no greeting, no placeholders.
- job_specific is true when the answer mentions or depends on this company or role, false when the same answer would
  fit any job.
Reply as JSON: {{"answer": string or null, "job_specific": true or false}}"""


def gemini_json(prompt, pdf):
    """Gemini with the resume PDF inline; returns the parsed JSON reply, or None if every model failed."""
    key = env("GEMINI_API_KEY")
    if not key:
        return None
    body = {"contents": [{"parts": [{"inline_data": {"mime_type": "application/pdf",
                                                     "data": base64.b64encode(pdf.read_bytes()).decode()}},
                                    {"text": prompt}]}],
            "generationConfig": {"responseMimeType": "application/json"}}
    for attempt in range(2):
        for model in GEMINI_MODELS:  # 503 "high demand" is common; fall through to the next model
            if time.time() - _cooldown.get(model, 0) < 600:
                continue
            req = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                data=json.dumps(body).encode(), headers={"Content-Type": "application/json", "x-goog-api-key": key})
            try:
                with urllib.request.urlopen(req, timeout=40) as r:
                    parts = json.load(r)["candidates"][0]["content"]["parts"]
                return json.loads("".join(p.get("text", "") for p in parts if not p.get("thought")))
            except Exception as e:
                if getattr(e, "code", None) == 429:
                    _cooldown[model] = time.time()
                print(f"     gemini {model} failed: {str(e)[:80]}")
        time.sleep(15)
    return None


def ai_answer(q, kind, opts, job):
    """Ask Gemini once; generic answers go into the bank (reused without AI), job-specific ones only to the AI sheet."""
    choices = [o for o in opts if o and not re.match(r"select( an)? option", o, re.I)]
    reply = gemini_json(AI_PROMPT.format(
        title=job.get("title", ""), company=job.get("company", ""), desc=(job.get("desc") or "")[:6000], q=q, kind=kind,
        opts=("Options: " + " | ".join(choices) + "\n") if choices else ""), HERE.parent / resume_for(job.get("title", "")))
    a = str((reply or {}).get("answer") or "").strip()
    if not a or a.lower() in ("null", "none"):
        return None
    if choices:
        low = [o.lower().strip() for o in choices]
        if a.lower() not in low:
            return None
        a = choices[low.index(a.lower())]
    if CREDENTIAL.search(q) and re.match(r"yes\b", a, re.I):
        return None  # second check: never claim a credential, even if the model does
    print(f"     [gemini{' per-job' if reply.get('job_specific') else ''}] {q[:60]!r} -> {a[:70]!r}")
    if reply.get("job_specific"):
        _ai_log(job, q, a)
    else:
        bank_record(q, kind, opts, job.get("site", ""), job.get("url", ""), answer=a, source="Gemini")
    return a


def _ai_log(job, q, a):
    from openpyxl import Workbook, load_workbook
    try:
        with _bank_lock():
            wb = load_workbook(QUESTION_BANK) if QUESTION_BANK.exists() else Workbook()
            if not QUESTION_BANK.exists():
                wb.active.title = "Questions"
                _sheet(wb, "Questions", BANK_COLS)
            _sheet(wb, AI_SHEET, AI_COLS).append([time.strftime("%Y-%m-%d %H:%M"), job.get("site", ""),
                                                 f"{job.get('title', '')} @ {job.get('company', '')}", q, a, job.get("url", "")])
            wb.save(QUESTION_BANK)
    except Exception as e:
        print(f"     (couldn't write {QUESTION_BANK.name}: {e})")


def smart_answer(q, kind, opts=(), job=None):
    """Bank -> fixed rules (salary, notice, phone, city, languages, total experience) -> Gemini on first sight.
    The blanket "every Yes/No -> Yes" and "1 year in any skill" rules are off here: Gemini answers those from the resume."""
    a = answer(q, kind, opts, blanket=False)
    if a is None and FEE.search(q):
        print(f"     !! asks the applicant to pay money, left unanswered: {q[:80]!r}")
        return None
    if a is None and job is not None:
        a = ai_answer(q, kind, opts, job)
    return a


def answer(q, kind, opts=(), blanket=True):
    """Return the answer for a form question, or None if we don't know it. Hand-written bank answers win.
    blanket=False drops the guesses (Yes to every Yes/No, 1 year in any skill) so the caller can ask Gemini instead."""
    banked = bank_answer(q, opts)
    if banked is not None:
        return banked
    ql = q.lower()
    opt_l = [o.lower() for o in opts]

    def pick(pattern):
        for o in opts:
            if re.search(pattern, o, re.I):
                return o
        return None

    if re.search(r"notice period|how soon can you|joining (date|time)|available to (join|start)|when can you (join|start)"
                 r"|availability to (join|start)|start date", ql):
        if kind not in ("select", "radio"):
            return "0"
        hit = pick(r"immediate|^0|\b0 ?(days?|months?)|15 days or less|less than|within")
        if hit:
            return hit  # otherwise fall through (e.g. a plain Yes/No -> Yes below)
    if re.search(r"current (ctc|salary|compensation|package|annual|fixed|pay)|present (ctc|salary)|last drawn", ql):
        return "0" if kind not in ("select", "radio") else None
    if re.search(r"expect\w*\s+(\w+\s+){0,2}(ctc|salary|compensation|package|pay)|salary expect|desired (salary|pay|compensation)", ql):
        if kind in ("select", "radio"):
            return None
        if re.search(r"lakh|lpa|lacs?\b", ql):
            return "3.5"
        if re.search(r"month", ql):
            return "29000"
        return "350000"
    if blanket and kind in ("select", "radio") and "yes" in opt_l:
        return opts[opt_l.index("yes")]  # every Yes/No -> Yes
    if re.search(r"proficiency in english|english (proficiency|level)|level of english", ql) and kind in ("select", "radio"):
        return pick(r"^professional")
    if re.search(r"proficiency in hindi|hindi (proficiency|level)|level of hindi", ql) and kind in ("select", "radio"):
        return pick(r"^native")
    if kind == "text" and re.search(r"\b(phone|mobile|contact number|whatsapp)\b", ql):
        return PHONE or None
    if re.search(r"how many years|years of (work )?experience|experience (do you have|in years)|(total|relevant|work) (work )?experience"
                 r"|experience\W*\(?in years", ql):
        if not blanket and re.search(r"experience\s+(do you have\s+)?(in|with|using|of|on)\s+(?!years?\b|total\b)\w", ql):
            return None  # "experience in Revit" is a skill question -> Gemini, not a flat 1
        if kind in ("select", "radio"):
            return pick(r"(?<![\d.])1(?![\d.])") or pick(r"0\s*-\s*[12]|less than 2|1\s*-\s*[23]")
        return "1"
    if kind in ("text", "typeahead") and re.search(r"\b(city|location|where are you (based|located)|current location)\b", ql):
        return "Faridabad"
    return None


# ---------------- logs ----------------
def load_done():
    if not APPLIED_LOG.exists():
        return set()
    return {r["job_id"] for r in csv.DictReader(open(APPLIED_LOG)) if not r["status"].startswith(RETRY_STATUSES)}


def _append(path, header, row):
    new = not path.exists()
    with open(path, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(header)
        w.writerow(row)


def log_job(job, status, query):
    _append(APPLIED_LOG, ["job_id", "title", "company", "location", "status", "query", "time"],
            [job["id"], job["title"], job["company"], job["location"], status, query, time.strftime("%Y-%m-%d %H:%M:%S")])


def log_question(job, field):
    bank_record(field["q"], field["kind"], field.get("opts") or [], "LinkedIn",
                f"https://www.linkedin.com/jobs/view/{job['id']}/")


# Skills live inside the description under headings like "Skills Required" / "Qualifications" / "Requirements".
SKILL_HEAD = re.compile(r"skill|qualification|requirement|what we.?re looking for|what you.?(ll )?(need|bring)|who you are|"
                        r"must.?have|good to have|nice to have|preferred|eligibility|competenc|you (have|bring|should)|"
                        r"ideal candidate|desired profile|candidate profile|key attributes|tools", re.I)
OTHER_HEAD = re.compile(r"responsibilit|what you.?ll do|role overview|about (us|the|our)|benefit|perks|what we offer|"
                        r"why (join|work)|compensation|salary|location|how to apply|job type|schedule|work mode|"
                        r"duties|day.to.day|the role|key result|kras?\b|interview|equal opportunity", re.I)


def skills_from(desc):
    """Pull the skills/qualifications sections out of a description (headings are short lines)."""
    out, take = [], False
    for line in (desc or "").splitlines():
        t = line.strip().strip("*•-–#").strip()
        if not t:
            continue
        strong = len(t) <= 60 and (t.endswith(":") or t.isupper())  # "SKILLS:" style heading
        weak = len(t.split()) <= 6                                    # short line; inside a section it may be a bullet
        if (strong or weak) and OTHER_HEAD.match(t):  # "Location: ... preferred" is not a skills heading
            take = False
        elif (strong or (weak and not take)) and SKILL_HEAD.search(t):
            take = True
            out.append(f"[{t.rstrip(':')}]")
        elif (strong or weak) and OTHER_HEAD.search(t):
            take = False
        elif take:
            out.append("• " + t)
    return "\n".join(out)[:8000]


XLSX_COLS = [("Applied on", 18), ("Title", 34), ("Company", 24), ("Location", 26), ("Search query", 30),
             ("Job URL", 40), ("Description", 120), ("Resume", 22), ("Skills", 70)]


EXTERNAL_COLS = [("Found on", 18), ("Title", 34), ("Company", 24), ("Location", 26), ("Search query", 30),
                 ("Job URL", 40), ("Resume", 22), ("Applied? (fill in)", 18)]


def _xlsx_append(path, sheet, cols, row):
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font
    try:
        if path.exists():
            wb = load_workbook(path); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active; ws.title = sheet
            ws.append([c for c, _ in cols])
            for i, (_, w) in enumerate(cols, 1):
                ws.cell(1, i).font = Font(bold=True)
                ws.column_dimensions[ws.cell(1, i).column_letter].width = w
            ws.freeze_panes = "A2"
        ws.append(row)
        for c in ws[ws.max_row]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        wb.save(path)
    except Exception as e:  # e.g. file open + locked; never kill a run over the spreadsheet
        print(f"     (couldn't write {path.name}: {e})")


def save_job_xlsx(job, query, desc, when=None):
    _xlsx_append(JOBS_XLSX, "Applied jobs", XLSX_COLS,
                 [when or time.strftime("%Y-%m-%d %H:%M"), job["title"], job["company"], job["location"], query,
                  f"https://www.linkedin.com/jobs/view/{job['id']}/", (desc or "")[:32000], resume_for(job["title"]), skills_from(desc)])  # Excel cell cap ~32k


def load_external():
    """Job ids already in external_jobs.xlsx, so reruns don't add duplicates."""
    if not EXTERNAL_XLSX.exists():
        return set()
    from openpyxl import load_workbook
    return {str(r[5]).rstrip("/").rsplit("/", 1)[-1] for r in load_workbook(EXTERNAL_XLSX).active.iter_rows(min_row=2, values_only=True) if r[5]}


def save_external(job, query):
    _xlsx_append(EXTERNAL_XLSX, "External jobs", EXTERNAL_COLS,
                 [time.strftime("%Y-%m-%d %H:%M"), job["title"], job["company"], job["location"], query,
                  f"https://www.linkedin.com/jobs/view/{job['id']}/", resume_for(job["title"]), ""])


# ---------------- page reading ----------------
CARDS_JS = r"""() => [...document.querySelectorAll('[role=button][componentkey^="job-card-component-ref-"]')].map(c => {
    const lines = c.innerText.split('\n').map(s => s.trim()).filter(Boolean);
    if (lines[0] && /\(Verified job\)$/.test(lines[0])) lines.shift();
    if (lines[1] === lines[0]) lines.splice(1, 1);  // unverified cards repeat the title
    const loc = lines.find(l => /\((Remote|On-site|Hybrid)\)$/.test(l)) || lines[2] || '';
    return {id: c.getAttribute('componentkey').replace('job-card-component-ref-', ''), title: lines[0] || '',
            company: lines[1] || '', location: loc, applied: lines.includes('Applied'), easy: /Easy Apply/.test(c.innerText)};
})"""

# Every unanswered-or-answered control in the open Easy Apply <dialog>, with its question text.
FIELDS_JS = r"""() => {
  const d = document.querySelector('dialog[open]'); if (!d) return null;
  const txt = e => (e && e.innerText || '').trim().replace(/\s+/g, ' ');
  // strip character counters like "0/20" and "0 of 20 characters"
  const clean = t => t.replace(/\b\d+\s*of\s*\d+\s*characters?\b/gi, ' ').replace(/\b\d+\s*\/\s*\d+\b/g, ' ').replace(/\s+/g, ' ').trim();
  const lab = el => el.id ? clean(txt(d.querySelector(`label[for="${CSS.escape(el.id)}"]`))) : '';
  // visible text of one radio option: its label, a wrapping label, or the nearest ancestor holding only this radio
  const optText = r => {
    // visible text first: on LinkedIn the radio's own aria-label is often the *question*, not the option
    let t = lab(r) || clean(txt(r.closest('label')));
    for (let a = r.parentElement, i = 0; !t && a && i < 4; a = a.parentElement, i++) {
      if (a.querySelectorAll('input[type=radio]').length > 1) break;
      t = clean(txt(a));
    }
    return t || (r.getAttribute('aria-label') || '').trim() || r.value;
  };
  const question = (el, exclude) => {
    const lb = el.getAttribute('aria-labelledby');
    if (lb) { const t = clean(lb.split(' ').map(i => txt(document.getElementById(i))).join(' ')); if (t.length > 3) return t; }
    const l = lab(el); if (l.length > 3) return l;
    const lg = el.closest('fieldset') && el.closest('fieldset').querySelector('legend');
    if (lg && clean(txt(lg)).length > 3) return clean(txt(lg));
    let a = el.closest('fieldset') || el;
    for (let i = 0; i < 6 && a.parentElement; i++) {
      a = a.parentElement; let t = txt(a);
      for (const x of exclude) if (x) t = t.split(x).join(' ');
      t = clean(t);
      if (t.length > 3) return t.slice(0, 300);
    }
    return '';
  };
  const out = [], seen = new Set(); let n = 0;
  for (const el of d.querySelectorAll('input, select, textarea')) {
    if (['hidden', 'file', 'submit'].includes(el.type)) continue;
    const req = el.required || el.getAttribute('aria-required') === 'true';
    if (el.type === 'radio') {
      if (seen.has(el.name)) continue; seen.add(el.name);
      const rs = [...d.querySelectorAll(`input[type=radio][name="${CSS.escape(el.name)}"]`)];
      const opts = rs.map(r => { r.dataset.bot = 'b' + (n++); return {id: r.dataset.bot, label: optText(r)}; });
      let q = question(rs[0], opts.map(o => o.label));
      const aria = clean(rs[0].getAttribute('aria-label') || '');
      if (aria.length > 3 && !opts.some(o => o.label === aria) && q.replace(/[*\s]|yes|no/gi, '').length < 4) q = aria;
      out.push({kind: 'radio', q: q, opts: opts.map(o => o.label), ids: opts.map(o => o.id),
                filled: rs.some(r => r.checked), required: true});
      continue;
    }
    if (!el.offsetParent && el.type !== 'checkbox') continue;
    el.dataset.bot = 'b' + (n++);
    if (el.tagName === 'SELECT') {
      const cur = el.options[el.selectedIndex];
      out.push({kind: 'select', id: el.dataset.bot, q: question(el, []), opts: [...el.options].map(o => o.text.trim()),
                filled: !!cur && el.selectedIndex > 0 && !/select an option/i.test(cur.text), required: req});
    } else if (el.type === 'checkbox') {
      out.push({kind: 'checkbox', id: el.dataset.bot, q: question(el, []), label: lab(el), filled: el.checked, required: req});
    } else {
      out.push({kind: el.getAttribute('role') === 'combobox' ? 'typeahead' : (el.tagName === 'TEXTAREA' ? 'textarea' : 'text'),
                id: el.dataset.bot, q: question(el, []), filled: !!el.value.trim(), required: req});
    }
  }
  return out;
}"""

ERRORS_JS = r"""() => { const d = document.querySelector('dialog[open]'); if (!d) return [];
  return [...d.querySelectorAll('[role=alert], [aria-invalid=true], .artdeco-inline-feedback--error')]
    .map(e => (e.innerText || e.getAttribute('aria-label') || 'invalid field').trim()).filter(Boolean).slice(0, 5); }"""


# ---------------- interaction ----------------
def modal_click(page, loc):
    """Curved mouse move onto the element, then Playwright's checked click at that same point."""
    loc.scroll_into_view_if_needed(timeout=5000)
    sr.pause(0.2, 0.6)
    box = loc.bounding_box()
    if not box:
        return loc.click(timeout=5000)
    ox, oy = box["width"] * random.uniform(0.3, 0.7), box["height"] * random.uniform(0.35, 0.65)
    sr.move_to(page, box["x"] + ox, box["y"] + oy)
    sr.pause(0.08, 0.35)
    loc.click(position={"x": ox, "y": oy}, timeout=5000)


def human_type(page, loc, text):
    modal_click(page, loc)
    loc.fill("")
    for ch in text:
        page.keyboard.type(ch)
        time.sleep(random.uniform(0.06, 0.18))
    sr.pause(0.3, 0.9)


def pick_resume(page, job):
    """On the resume step, select the right saved resume. False if the step is there but our file isn't."""
    radios = page.locator('dialog[open] input[type=radio][aria-label$=".pdf"], dialog[open] input[type=radio][aria-label$=".docx"]')
    if not radios.count():
        return True  # not the resume step
    want = resume_for(job["title"])
    radio = page.locator(f'dialog[open] input[type=radio][aria-label="{want}"]')
    if not radio.count():
        return False
    if not radio.first.is_checked():
        card = page.locator(f'dialog[open] [role=button]:has(span:text-is("{want}"))')
        if card.count() and card.first.is_visible():
            modal_click(page, card.first)
            sr.pause(0.5, 1)
        if not radio.first.is_checked():
            radio.first.check(force=True)
    return radio.first.is_checked()


def fill_page(page, job):
    """Fill the current Easy Apply step. Returns the fields we can't answer (all of them, so they reach the bank together)."""
    unknown = []
    for f in page.evaluate(FIELDS_JS) or []:
        el = page.locator(f'[data-bot="{f.get("id", "")}"]')
        if f["kind"] == "checkbox":
            if "follow" in f["label"].lower() and f["filled"]:
                modal_click(page, page.locator(f'label[for="{el.get_attribute("id")}"]'))  # don't follow the company
            elif f["required"] and not f["filled"]:
                modal_click(page, page.locator(f'label[for="{el.get_attribute("id")}"]'))  # consent / terms
            continue
        if f["filled"]:
            continue
        a = smart_answer(f["q"], f["kind"], f.get("opts") or [], job)
        if a is None:
            if f["required"] or f["kind"] == "radio":
                unknown.append(f)
            continue  # optional and unknown -> leave blank
        if f["kind"] == "radio":
            rid = f["ids"][f["opts"].index(a)]
            radio = page.locator(f'[data-bot="{rid}"]')
            label = page.locator(f'label[for="{radio.get_attribute("id")}"]')
            if label.count() and label.first.is_visible() and (label.first.bounding_box() or {}).get("width", 0) > 4:
                modal_click(page, label.first)
            else:
                box = radio.bounding_box()
                if box:
                    sr.move_to(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                radio.check(force=True)  # styled radios hide the real <input>
        elif f["kind"] == "select":
            modal_click(page, el)
            sr.pause(0.3, 0.8)
            el.select_option(label=a)
        elif f["kind"] == "typeahead":
            human_type(page, el, a)
            sr.pause(1, 2)
            page.keyboard.press("ArrowDown")
            page.keyboard.press("Enter")
        else:
            human_type(page, el, a)
        sr.pause(0.4, 1.2)
    return unknown


def discard(page):
    try:
        page.locator('dialog[open] button[aria-label="Dismiss"]').first.click(timeout=4000)
        sr.pause(0.6, 1.2)
        d = page.locator("dialog[open] button", has_text="Discard")
        if d.count():
            d.first.click(timeout=4000)
    except Exception:
        page.keyboard.press("Escape")
    sr.pause(0.8, 1.5)


def debug_shot(page, tag, job_id=""):
    """ponytail: temporary debug screenshots at failure points. Remove this function and every debug_shot( call once runs are clean."""
    name = f"{time.strftime('%m%d_%H%M%S')}_{tag}_{job_id}.png"
    try:
        (HERE / "debug").mkdir(exist_ok=True)
        page.screenshot(path=str(HERE / "debug" / name))
        print(f"      screenshot: debug/{name}")
    except Exception:
        pass


SENT_TEXT = re.compile(r"application was sent|application sent", re.I)


def close_sent_box(page):
    """Close the "Application sent" box; verify it's gone, since a click on Done has silently failed before."""
    for _ in range(3):
        if not page.get_by_text(SENT_TEXT).locator("visible=true").count():
            return True
        done = page.locator(':is(button, a, [role=button]):text-is("Done"):visible, [aria-label="Dismiss"]:visible')
        try:
            if done.count():
                sr.human_click(page, done.first)
            else:
                page.keyboard.press("Escape")
        except Exception:
            page.keyboard.press("Escape")
        sr.pause(0.8, 1.5)
    # ponytail: debug dump; remove with debug_shot once the box always closes
    print("      'Application sent' box won't close; Done candidates:",
          page.evaluate("""() => [...document.querySelectorAll('*')].filter(e => e.children.length < 3 && (e.innerText||'').trim() === 'Done')
                            .map(e => e.outerHTML.slice(0, 200))"""))
    return False


def close_leftovers(page):
    """Close a lingering "Application sent" box (or any open dialog) so it can't swallow the next click."""
    for _ in range(3):
        if not page.locator("dialog[open]:visible, [role=dialog]:visible").count():
            return
        btn = page.locator("dialog[open], [role=dialog]").get_by_role("button", name=re.compile(r"^(Done|Dismiss|Close)$", re.I))
        try:
            btn.locator("visible=true").first.click(timeout=2000)
        except Exception:
            page.keyboard.press("Escape")
        sr.pause(0.6, 1.2)


DESC_JS = r"""() => {
  const h = [...document.querySelectorAll('h1,h2,h3,span,div')]
    .find(e => /^\s*About the job\s*$/.test(e.innerText || '') && e.children.length < 3);
  if (!h) return '';
  let n = h; while (n.parentElement && n.innerText.length < 300) n = n.parentElement;
  return n.innerText.replace(/^\s*About the job\s*/, '').replace(/\s*… more\s*$/, '')
                    .replace(/\n{3,}/g, '\n\n').trim();
}"""


def job_description(page):
    try:
        return page.evaluate(DESC_JS)
    except Exception:
        return ""


def read_job(page):
    """Skim the description in the right-hand pane."""
    vw, vh = page.evaluate("[innerWidth, innerHeight]")
    sr.move_to(page, vw * random.uniform(0.5, 0.7), vh * random.uniform(0.4, 0.7))
    for _ in range(random.randint(2, 4)):
        for _ in range(random.randint(3, 7)):
            page.mouse.wheel(0, random.uniform(40, 110))
            time.sleep(random.uniform(0.02, 0.08))
        sr.pause(1, 3.5)


def apply_open_job(page, job, dry_run):
    """Job is selected in the search page. Returns a status string."""
    job.update(desc=job_description(page), site="LinkedIn", url=f"https://www.linkedin.com/jobs/view/{job['id']}/")  # Gemini context
    # the card said Easy Apply, so wait for the pane to load the button instead of checking once
    # LinkedIn serves this as a <button> or an <a>, depending on the build
    btn = page.locator(':is(button, a)[aria-label="Easy Apply to this job"]:visible, :is(button, a)[aria-label^="Continue applying"]:visible').first
    try:
        btn.wait_for(timeout=8000)
    except PWTimeout:
        debug_shot(page, "no_button", job["id"])
        return "no_button"  # retried next run
    sr.human_click(page, btn)  # wheels back up to the button after reading
    try:
        page.locator("dialog[open]").wait_for(timeout=10000)
    except PWTimeout:
        debug_shot(page, "no_dialog", job["id"])
        return "no_dialog"
    sr.pause(1, 2)
    if page.locator("dialog[open] h2", has_text=re.compile("safety reminder", re.I)).count():
        discard(page)
        return "safety_warning"

    last_sig = None
    for _ in range(15):
        if page.get_by_text(DAILY_LIMIT).count():
            discard(page)
            return "daily_limit"
        if not pick_resume(page, job):
            print(f"      {resume_for(job['title'])} isn't among the saved resumes; skipping")
            discard(page)
            return "resume_missing"
        unknown = fill_page(page, job)
        for u in unknown:
            log_question(job, u)
            print(f"      unknown question: {u['q'][:100]!r} ({u['kind']})")
        if unknown:
            discard(page)
            return "unknown_question"
        dlg = page.locator("dialog[open]")
        submit = dlg.locator("button", has_text=re.compile(r"^\s*Submit application\s*$"))
        if submit.count():
            if dry_run:
                print("      dry run: reached 'Submit application', not clicking")
                discard(page)
                return "dry_run_ok"
            sr.pause(1, 3)  # glance over the review page
            modal_click(page, submit.first)
            try:
                page.get_by_text(re.compile(r"application was sent|application sent|application submitted", re.I)).first.wait_for(timeout=10000)
            except PWTimeout:
                debug_shot(page, "submitted_unverified", job["id"])
                return "submitted_unverified"
            sr.pause(1, 2)
            close_sent_box(page)
            sr.pause(1, 1.5)
            debug_shot(page, "after_applied", job["id"])  # is the "Application sent" box really gone?
            return "applied"
        nxt = dlg.locator("button", has_text=re.compile(r"^\s*(Next|Continue|Review)\s*$"))
        if not nxt.count():
            debug_shot(page, "stuck_no_next", job["id"])
            discard(page)
            return "stuck"
        sig = page.evaluate("() => document.querySelector('dialog[open]').innerText.slice(0, 400)")
        if sig == last_sig:  # pressing Next didn't move us
            errs = page.evaluate(ERRORS_JS)
            print(f"      form error: {errs}")
            debug_shot(page, "form_error", job["id"])
            discard(page)
            return "form_error"
        last_sig = sig
        modal_click(page, nxt.first)
        sr.pause(1.2, 2.5)
    debug_shot(page, "stuck_15_steps", job["id"])
    discard(page)
    return "stuck"


CARD_SEL = '[role=button][componentkey^="job-card-component-ref-"]'


def next_page(page):
    """Click "Next" under the results list and wait for the new cards. False on the last page."""
    btn = page.locator('[data-testid="pagination-controls-next-button-visible"]')
    if not btn.count():
        return False
    sr.dismiss_popups(page)  # a leftover "Application sent" box swallows the click on Next
    first = page.locator(CARD_SEL).first
    key, box = first.get_attribute("componentkey"), first.bounding_box()
    if box:  # mouse over the list so the wheel scrolls the list down to the buttons
        sr.move_to(page, box["x"] + box["width"] * random.uniform(0.3, 0.7), page.evaluate("innerHeight") * random.uniform(0.4, 0.6))
    for attempt in (1, 2):
        sr.human_click(page, btn.first)
        try:
            page.wait_for_function("([sel, k]) => document.querySelector(sel)?.getAttribute('componentkey') !== k",
                                   arg=[CARD_SEL, key], timeout=15000)
            break
        except PWTimeout:
            if attempt == 2:
                print("   next page didn't load; moving to the next search")
                debug_shot(page, "next_page_failed")
                return False
            sr.dismiss_popups(page)
    sr.pause(2, 4)
    return True


def search_url(query, mode):
    q = {"keywords": query if mode == "ncr" else f"{query} remote", "f_TPR": "r604800",  # no f_AL: external jobs get saved too
         "geoId": NCR_GEO if mode == "ncr" else WORLD_GEO}
    return "https://www.linkedin.com/jobs/search-results/?" + up.urlencode(q)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profile", default="Profile 26")
    ap.add_argument("--limit", type=int, default=20, help="max applications this run")
    ap.add_argument("--min-delay", type=float, default=30)
    ap.add_argument("--max-delay", type=float, default=120)
    ap.add_argument("--pages", type=int, default=3, help="result pages to read per search (25 jobs each)")
    ap.add_argument("--dry-run", action="store_true", help="fill forms up to 'Submit application', then discard")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--backfill", action="store_true", help="add past applications missing from applied_jobs.xlsx, then exit")
    args = ap.parse_args()
    if args.selftest:
        return selftest()

    done, seen, external = load_done(), set(), load_external()
    searches = [(q, m) for q in SEARCH_QUERIES for m in ("ncr", "remote")]
    random.shuffle(searches)
    sr.prepare_profile(args.profile)
    applied = 0
    next_break = random.randint(4, 7)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(sr.LOCAL_DATA), channel="chrome", headless=False,
            args=[f"--profile-directory={args.profile}"], no_viewport=True,
            ignore_default_args=["--use-mock-keychain"],
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://www.linkedin.com/feed/")
        if any(k in page.url for k in ("/login", "/authwall", "/uas/", "/checkpoint/")):
            input("Log in to LinkedIn in the opened window, then press Enter...")
        if args.backfill:
            return backfill(page, ctx)
        sr.read_profile(page)

        try:
            for query, mode in searches:
                if applied >= args.limit:
                    break
                page.goto(search_url(query, mode), wait_until="domcontentloaded")
                sr.check_session(page)
                try:
                    page.locator('[componentkey^="job-card-component-ref-"]').first.wait_for(timeout=15000)
                except PWTimeout:
                    print(f"[{query} / {mode}] no results")
                    continue
                for page_no in range(1, args.pages + 1):
                    sr.pause(2, 4)
                    jobs = [j for j in page.evaluate(CARDS_JS) if j["id"] not in seen]
                    seen.update(j["id"] for j in jobs)
                    fits = [j for j in jobs if not j["applied"] and title_matches(j["title"]) and location_ok(j["location"], mode)]
                    for j in fits:
                        if not j["easy"] and j["id"] not in external:
                            save_external(j, f"{query} / {mode}")
                            external.add(j["id"])
                            print(f"  saved external: {j['title']} | {j['company']} | {j['location']}")
                    picks = [j for j in fits if j["easy"] and j["id"] not in done]
                    print(f"[{query} / {mode} p{page_no}] {len(jobs)} new cards, {len(picks)} easy apply to try")
                    for job in picks:
                        if applied >= args.limit:
                            break
                        print(f"  -> {job['title']} | {job['company']} | {job['location']}")
                        try:
                            close_leftovers(page)
                            card = page.locator(f'[role=button][componentkey="job-card-component-ref-{job["id"]}"]').first
                            box = card.bounding_box()
                            if box:  # put the mouse over the list so the wheel scrolls the list, not the detail pane
                                vh = page.evaluate("innerHeight")
                                sr.move_to(page, box["x"] + box["width"] * random.uniform(0.3, 0.7), vh * random.uniform(0.4, 0.6))
                            for attempt in (1, 2):  # first click after an application often just closes the "sent" box
                                sr.human_click(page, card)
                                try:
                                    page.wait_for_url(re.compile(f"currentJobId={job['id']}"), timeout=10000)
                                    break
                                except PWTimeout:
                                    if attempt == 2:
                                        raise
                            sr.pause(1.5, 3)
                            read_job(page)
                            desc = job_description(page)
                            status = apply_open_job(page, job, args.dry_run)
                        except sr.SessionLost:
                            raise
                        except PWTimeout as e:
                            status, desc = "timeout", ""
                            import traceback  # ponytail: debug aid for timeouts; drop with debug_shot
                            here = [f for f in traceback.extract_tb(e.__traceback__) if f.filename.endswith(("apply_jobs.py", "send_request.py"))]
                            print(f"      timed out at {here[-1].name}:{here[-1].lineno} -> {here[-1].line}")
                            debug_shot(page, "timeout", job["id"])
                            discard(page)
                        except Exception as e:
                            if any(k in str(e).lower() for k in sr.SESSION_ERRORS):
                                raise sr.SessionLost(str(e)[:120])
                            status = f"error: {e.__class__.__name__}"
                            discard(page)
                        print(f"     {status}")
                        if status == "daily_limit":
                            print("LinkedIn says the Easy Apply limit is reached; stopping.")
                            return
                        if not args.dry_run:
                            log_job(job, status, f"{query} / {mode}")
                        done.add(job["id"])
                        if status in ("applied", "submitted_unverified") and not args.dry_run:
                            save_job_xlsx(job, f"{query} / {mode}", desc)
                        if status in ("applied", "submitted_unverified", "dry_run_ok"):
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
                    if applied >= args.limit or page_no == args.pages or not next_page(page):
                        break
                sr.pause(4, 12)  # between searches
        except sr.SessionLost as e:
            print(f"Session lost ({e}); stopping.")
        finally:
            print(f"\nDone: {applied} {'dry-run forms completed' if args.dry_run else 'applications'} this run.")
            ctx.close()


def backfill(page, ctx):
    """Fetch descriptions for applications logged in applied_log.csv but not yet in the spreadsheet."""
    from openpyxl import load_workbook
    have = set()
    if JOBS_XLSX.exists():
        have = {str(r[5]).rstrip("/").rsplit("/", 1)[-1] for r in load_workbook(JOBS_XLSX).active.iter_rows(min_row=2, values_only=True)}
    rows = [r for r in csv.DictReader(open(APPLIED_LOG))
            if r["status"] in ("applied", "submitted_unverified") and r["job_id"] not in have]
    print(f"backfilling {len(rows)} jobs")
    try:
        for r in rows:
            page.goto(f"https://www.linkedin.com/jobs/view/{r['job_id']}/", wait_until="domcontentloaded")
            sr.check_session(page)
            sr.pause(4, 7)
            read_job(page)
            desc = job_description(page)
            print(f"  {r['title']} | {r['company']}: {len(desc)} chars")
            job = {"id": r["job_id"], "title": r["title"], "company": r["company"], "location": r["location"]}
            save_job_xlsx(job, r["query"], desc, r["time"][:16])
            sr.pause(5, 15)
    finally:
        ctx.close()


def selftest():
    d = ("About Us\nWe are a startup.\nResponsibilities:\nCold calling\nSkills Required:\nMS Excel\nCommunication skills\n"
         "Benefits\nHealth insurance\nPreferred Qualifications\nB.Pharm")
    assert skills_from(d) == "[Skills Required]\n• MS Excel\n• Communication skills\n[Preferred Qualifications]\n• B.Pharm", skills_from(d)
    assert skills_from("Location: Gurugram (Local candidate is preferred)\nWork days: 5") == ""
    assert skills_from("Just a paragraph with no headings at all.") == ""
    for t in ("Content Writer", "Social Media Executive", "Medical Content Writer", "Copywriter", "Community Associate",
              "PR & Communications Associate", "UGC Content Creator"):
        assert resume_for(t) == "resume_content.pdf", t
    for t in ("QA Executive", "Founder's Office Associate", "Business Development Associate", "Growth Associate",
              "Clinical Trial Coordinator"):
        assert resume_for(t) == "resume.pdf", t
    for t in ("Pharmacovigilance Associate", "QA Executive", "Quality Control Trainee", "Customer Success Executive",
              "Business Development Associate - HealthTech", "Virtual Assistant", "Canva Graphic Designer", "Data Analyst",
              "Founder's Office Associate", "Founders Office Intern", "Chief of Staff", "Associate - Office of the CEO",
              "Growth Associate", "Growth Marketing Intern", "Performance Marketing Executive", "Content Writer",
              "Content Manager", "Social Media Manager", "Community Manager", "UGC Content Creator", "Medical Writer",
              "Copywriter", "Content Strategist", "Scientific Writer - Pharma"):
        assert title_matches(t), t
    for t in ("Pharma/Clinical Operations QA/Test Engineer", "Senior QA Manager", "Staff Nurse", "Sr. Business Analyst",
              "Python Developer", "MBBS Doctor", "Legal Associate", "Quantitative Researcher - Early Career (INDIA)",
              "Intern (Onsite) :: North America- Geopolitics", "Associate Scientist, Assay Validation",
              "Senior Content Manager", "Head of Growth", "Growth Manager", "Co-Founder (Equity)", "Video Editor",
              "Lead Social Media Manager", "Chief Operating Officer"):
        assert not title_matches(t), t
    assert location_ok("India (Remote)", "remote") and location_ok("Gurugram (On-site)", "ncr")
    assert location_ok("Delhi, India (Hybrid)", "ncr") and not location_ok("Mumbai (On-site)", "ncr")
    assert not location_ok("New Delhi (On-site)", "remote")
    assert answer("How many years of experience do you have with Python?*", "text") == "1"
    assert answer("Do you have experience in healthcare compliance or auditing? *", "radio", ["Yes", "No"]) == "Yes"
    assert answer("Are you comfortable commuting?", "select", ["Select an option", "Yes", "No"]) == "Yes"
    assert answer("What is your current CTC?", "text") == "0"
    assert answer("Expected CTC (in LPA)", "text") == "3.5"
    assert answer("Expected salary per month", "text") == "29000"
    assert answer("What is your expected CTC?", "text") == "350000"
    assert answer("What is your notice period?", "text") == "0"
    assert answer("Notice period", "select", ["Select an option", "Immediate", "30 days"]) == "Immediate"
    assert answer("Years of experience in sales", "select", ["Select", "0", "1", "2-3"]) == "1"
    assert answer("Why do you want to join us?", "textarea") is None
    assert PHONE and answer("Phone*", "text") == PHONE, "set PHONE in the project-root .env"
    assert answer("What is your total work experience? (in years)*", "text") == "1"
    assert answer("What is your availability to start (in days) upon job offer?*", "text") == "0"
    assert answer("Please specify your expected monthly salary in INR.*", "text") == "29000"
    assert answer("Upon successful completion of the 3-month internship, we will be offering a full-time employment "
                  "opportunity. Please apply only if you are also interested and available to join us in a full-time role.*",
                  "radio", ["Yes", "No"]) == "Yes"
    assert answer("What is your level of proficiency in English?*", "select",
                  ["Select an option", "None", "Conversational", "Professional", "Native or bilingual"]) == "Professional"
    assert answer("What is your level of proficiency in Hindi?*", "select",
                  ["Select an option", "None", "Conversational", "Professional", "Native or bilingual"]) == "Native or bilingual"
    assert answer("What is your level of proficiency in Punjabi?*", "select",
                  ["Select an option", "None", "Conversational", "Professional", "Native or bilingual"]) is None
    assert answer("Can you join immediately?*", "radio", ["Yes", "No"]) == "Yes"
    assert DAILY_LIMIT.search("You reached today’s Easy Apply limit") and DAILY_LIMIT.search("reached the daily Easy Apply limit")
    print("selftest ok")


if __name__ == "__main__":
    main()
