"""LinkedIn Easy Apply bot for Soumya.

    .venv/bin/python apply_jobs.py --profile "Profile 26" --limit 20            # real applications
    .venv/bin/python apply_jobs.py --profile "Profile 26" --limit 5 --dry-run   # fills forms, never submits

Searches every query in SEARCH_QUERIES twice: Delhi NCR (any workplace type) and worldwide remote.
Only Easy Apply jobs posted in the past week whose title passes the keyword bank are opened.
Forms: years of experience -> 1, every Yes/No -> Yes, current CTC 0, notice 0, expected CTC 3.5 LPA.
Any question the answer bank can't handle -> application discarded, question logged to unanswered_questions.csv.
"""
import argparse
import csv
import os
import random
import re
import time
import urllib.parse as up
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
UNANSWERED_LOG = HERE / "unanswered_questions.csv"
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
def answer(q, kind, opts=()):
    """Return the answer for a form question, or None if we don't know it."""
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
    if kind in ("select", "radio") and "yes" in opt_l:
        return opts[opt_l.index("yes")]  # every Yes/No -> Yes
    if re.search(r"proficiency in english|english (proficiency|level)|level of english", ql) and kind in ("select", "radio"):
        return pick(r"^professional")
    if re.search(r"proficiency in hindi|hindi (proficiency|level)|level of hindi", ql) and kind in ("select", "radio"):
        return pick(r"^native")
    if kind == "text" and re.search(r"\b(phone|mobile|contact number|whatsapp)\b", ql):
        return PHONE or None
    if re.search(r"how many years|years of (work )?experience|experience (do you have|in years)|(total|relevant|work) (work )?experience"
                 r"|experience\W*\(?in years", ql):
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
    _append(UNANSWERED_LOG, ["question", "kind", "options", "job_id", "job_title", "time"],
            [field["q"], field["kind"], " | ".join(field.get("opts") or []), job["id"], job["title"], time.strftime("%Y-%m-%d %H:%M:%S")])


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
    """Fill the current Easy Apply step. Returns the first field we can't answer, or None."""
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
        a = answer(f["q"], f["kind"], f.get("opts") or [])
        if a is None:
            if not f["required"] and f["kind"] != "radio":
                continue  # optional and unknown -> leave blank
            return f
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
    return None


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
        if unknown:
            log_question(job, unknown)
            print(f"      unknown question: {unknown['q'][:100]!r} ({unknown['kind']})")
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
