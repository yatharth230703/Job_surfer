
import argparse
import csv
import random
import re
import shutil
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

HERE = Path(__file__).parent
DATA_DIR = HERE.parent / "profile_data"
SENT_LOG = HERE / "sent_log.csv"
LOCAL_DATA = HERE / "chrome_data"
CHROME_DIR = Path.home() / "Library/Application Support/Google/Chrome"
INVITE = re.compile(r"^Invite .+ to connect$")
RETRY_STATUSES = ("timeout", "error", "connect_failed")
SESSION_ERRORS = ("err_too_many_redirects", "err_tunnel_connection_failed",
                  "err_proxy_connection_failed", "err_connection_reset", "err_internet_disconnected")


class SessionLost(Exception):
    pass


def load_urls():
    urls = {}
    for f in sorted(DATA_DIR.glob("*.csv")):
        for row in csv.DictReader(open(f, newline="", encoding="utf-8")):
            m = re.search(r"linkedin\.com/in/([^/?#\s]+)", row.get("LinkedIn Profile URL", ""))
            if m:
                urls.setdefault(f"https://www.linkedin.com/in/{m.group(1)}/", row["Full Name"])
    return urls


def load_done():
    if not SENT_LOG.exists():
        return set()
    # timeouts/errors/failed clicks are retried next run
    return {r["url"] for r in csv.DictReader(open(SENT_LOG))
            if not r["status"].startswith(RETRY_STATUSES)}


def log(url, name, status):
    new = not SENT_LOG.exists()
    with open(SENT_LOG, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["url", "name", "status", "time"])
        w.writerow([url, name, status, time.strftime("%Y-%m-%d %H:%M:%S")])


def prepare_profile(profile):
    if (LOCAL_DATA / profile).exists():
        return
    src = CHROME_DIR / profile
    if not src.exists():
        raise SystemExit(f"No Chrome profile '{profile}' in {CHROME_DIR}")
    print(f"Copying Chrome profile '{profile}' -> {LOCAL_DATA} (one-time)...")
    LOCAL_DATA.mkdir(exist_ok=True)
    shutil.copy2(CHROME_DIR / "Local State", LOCAL_DATA / "Local State")
    shutil.copytree(src, LOCAL_DATA / profile, symlinks=True,
                    ignore=shutil.ignore_patterns("*Cache*", "Service Worker", "Singleton*", "*.lock"))


# ---------- human-like behaviour ----------
_pos = [random.randint(200, 600), random.randint(150, 400)]  # Playwright doesn't expose cursor position


def pause(lo, hi):
    """Wait lo..hi seconds, skewed toward short waits with an occasional long one (beta 2,5)."""
    time.sleep(lo + (hi - lo) * random.betavariate(2, 5))


def bezier_points(start, end, steps):
    """Cubic Bézier path from start to end with random control points, eased (slow-fast-slow)."""
    (x0, y0), (x3, y3) = start, end
    dist = max(((x3 - x0) ** 2 + (y3 - y0) ** 2) ** 0.5, 1)
    spread = min(dist * 0.4, 200)
    x1, y1 = x0 + (x3 - x0) * random.uniform(0.2, 0.4) + random.uniform(-spread, spread), y0 + (y3 - y0) * random.uniform(0.2, 0.4) + random.uniform(-spread, spread)
    x2, y2 = x0 + (x3 - x0) * random.uniform(0.6, 0.8) + random.uniform(-spread, spread), y0 + (y3 - y0) * random.uniform(0.6, 0.8) + random.uniform(-spread, spread)
    pts = []
    for i in range(1, steps + 1):
        t = i / steps
        t = t * t * (3 - 2 * t)  # smoothstep easing
        u = 1 - t
        pts.append((u**3 * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t**3 * x3,
                    u**3 * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t**3 * y3))
    return pts


def move_to(page, x, y):
    dist = ((x - _pos[0]) ** 2 + (y - _pos[1]) ** 2) ** 0.5
    steps = int(min(max(dist / random.uniform(8, 14), 12), 70))
    target = (x, y)
    if dist > 300 and random.random() < 0.3:  # overshoot a little, then correct
        target = (x + random.uniform(-12, 12), y + random.uniform(-10, 10))
    for px, py in bezier_points(_pos, target, steps):
        page.mouse.move(px + random.uniform(-0.6, 0.6), py + random.uniform(-0.6, 0.6))
        time.sleep(random.uniform(0.004, 0.018))
    if target != (x, y):
        time.sleep(random.uniform(0.05, 0.15))
        for px, py in bezier_points(target, (x, y), random.randint(5, 9)):
            page.mouse.move(px, py)
            time.sleep(random.uniform(0.01, 0.03))
    _pos[:] = [x, y]


def wheel_into_view(page, locator):
    """Scroll with the mouse wheel (in uneven flicks) until the element sits comfortably on screen."""
    vh = page.evaluate("innerHeight")
    last_top = None
    for _ in range(30):
        box = locator.bounding_box()
        if not box:
            break
        top = box["y"]
        if 90 <= top and top + box["height"] <= vh - 60:  # 90px clears LinkedIn's fixed top nav
            pause(0.4, 1.2)  # eyes land on it before the hand moves
            return
        if top == last_top:
            break  # wheel isn't moving this element -> fall back below
        last_top = top
        delta = top - vh * random.uniform(0.25, 0.45)  # aim for upper-middle of the screen
        flick = random.uniform(250, 600)
        flick = max(min(delta, flick), -flick)
        ticks = random.randint(3, 7)
        for _ in range(ticks):
            page.mouse.wheel(0, flick / ticks)
            time.sleep(random.uniform(0.02, 0.08))
        pause(0.3, 1.0)
    locator.scroll_into_view_if_needed(timeout=5000)
    pause(0.4, 1.2)


def human_click(page, locator):
    """Curved move + click, but only if the target is really what's under the cursor."""
    wheel_into_view(page, locator)
    box = locator.bounding_box()
    if box:
        x = box["x"] + box["width"] * random.uniform(0.3, 0.7)
        y = box["y"] + box["height"] * random.uniform(0.35, 0.65)
        move_to(page, x, y)
        pause(0.08, 0.4)  # hover before clicking
        on_target = locator.evaluate(
            "(el, [x, y]) => { const t = document.elementFromPoint(x, y); return !!t && (el === t || el.contains(t)); }",
            [x, y])
        if on_target:
            page.mouse.down()
            time.sleep(random.uniform(0.05, 0.14))
            page.mouse.up()
            return
        blocker = page.evaluate("""([x, y]) => { const t = document.elementFromPoint(x, y); if (!t) return 'nothing';
            const c = t.closest('[aria-label],[data-testid],[componentkey],button,a') || t;
            return `${t.tagName} in ${c.tagName} aria=${c.getAttribute('aria-label')} testid=${c.getAttribute('data-testid')} text=${(c.innerText||'').trim().slice(0, 40)}`; }""", [x, y])
        print(f"   something is covering the button ({blocker}), closing popups")  # ponytail: blocker detail is debug; trim once understood
        dismiss_popups(page)
    locator.click(timeout=5000)  # Playwright refuses if still covered -> raises, lead retried next run


def wander(page):
    """Idle mouse drift somewhere on screen."""
    vw, vh = page.evaluate("[innerWidth, innerHeight]")
    move_to(page, random.uniform(vw * 0.15, vw * 0.85), random.uniform(vh * 0.15, vh * 0.85))


def read_profile(page):
    """Skim the profile like a person: drift, scroll down in uneven chunks, pause, come back up."""
    pause(1.5, 4)
    wander(page)
    for _ in range(random.randint(2, 5)):
        for _ in range(random.randint(3, 8)):  # one 'flick' of the wheel = several small ticks
            page.mouse.wheel(0, random.uniform(40, 120))
            time.sleep(random.uniform(0.02, 0.08))
        pause(0.8, 4)
        if random.random() < 0.4:
            wander(page)
    if random.random() < 0.25:  # sometimes scroll up a bit to re-read
        page.mouse.wheel(0, -random.uniform(150, 400))
        pause(0.8, 2.5)
    # no jump back to top here: human_click wheels back up to whichever button it needs


def idle_wait(page, lo, hi):
    """Long wait between profiles, broken into chunks with small scrolls / mouse drift in between."""
    remaining = lo + (hi - lo) * random.betavariate(2, 5)
    while remaining > 0:
        chunk = min(random.uniform(10, 45), remaining)
        time.sleep(chunk)
        remaining -= chunk
        if remaining <= 0:
            break
        try:
            r = random.random()
            if r < 0.4:
                page.mouse.wheel(0, random.uniform(-150, 150))
            elif r < 0.7:
                wander(page)
        except Exception:
            pass


# ---------- LinkedIn ----------
def dismiss_popups(page):
    for sel in ('[role="dialog"] button[aria-label="Dismiss"]', '[role="dialog"] button[aria-label="Close"]',
                '[role="alertdialog"] button[aria-label="Dismiss"]', '[data-test-modal-close-btn]'):
        btn = page.locator(sel)
        for i in range(btn.count()):
            if btn.nth(i).is_visible():
                try:
                    btn.nth(i).click(timeout=2000)
                    pause(0.4, 1)
                except Exception:
                    pass
    page.keyboard.press("Escape")
    pause(0.3, 0.6)


def check_session(page):
    if any(k in page.url for k in ("/login", "/checkpoint/", "/authwall", "/uas/")):
        raise SessionLost(f"LinkedIn sent us to {page.url}")


def profile_name(page):
    """'(3) Kunal Banerjee | LinkedIn' -> 'Kunal Banerjee'"""
    title = re.sub(r"^\(\d+\+?\)\s*", "", page.title())
    name = title.rsplit("|", 1)[0].strip() if "|" in title else ""
    return "" if name.lower() in ("linkedin", "page not found", "sign in", "security verification") else name


def first_visible(page, selector):
    loc = page.locator(selector)
    for i in range(loc.count()):
        if loc.nth(i).is_visible():
            return loc.nth(i)
    return None


def connect(page, dry_run=False):
    """Returns a status string for the current profile page."""
    check_session(page)
    name = profile_name(page)
    if not name:
        return "not_a_profile"
    slug = page.url.split("/in/")[-1].split("/")[0].split("?")[0]
    q = name.replace('"', '\\"')
    # Scoped to THIS person (name in aria-label or slug in the invite link) so sidebar
    # "People also viewed" Connect buttons can never match.
    invite_sel = f'[aria-label="Invite {q} to connect"], [href*="custom-invite/?vanityName={slug}"]'
    pending_sel = f'[aria-label*="{q}"][aria-label*="Pending"], [aria-label*="{q}"][aria-label*="ithdraw"]'

    dismiss_popups(page)
    if first_visible(page, pending_sel):
        return "already_pending"
    btn = first_visible(page, invite_sel)
    if not btn:
        # top-card 'More' is a plain <button><span>More</span></button> (no aria-label); first in DOM = top card
        more = first_visible(page, 'main [aria-label="More actions"], main [aria-label="More"], '
                                   'main button:has(span:text-is("More"))')
        if not more:
            return "no_connect_option"
        print("   Connect is under the More menu")
        human_click(page, more)
        pause(0.6, 1.5)
        # dropdown item: usually has the same aria-label; newer layouts only show the text "Connect".
        # Text matches are limited to the open menu so sidebar cards can't match.
        btn = first_visible(page, invite_sel) or first_visible(
            page, '[role="menu"] :text-is("Connect"), [role="menuitem"]:has-text("Connect"), '
                  '.artdeco-dropdown__content--is-open :text-is("Connect")')
        if not btn:
            page.keyboard.press("Escape")
            return "no_connect_option"  # already connected / follow-only
    if dry_run:  # scroll to + hover the Connect button, never click it (some profiles send with no modal)
        wheel_into_view(page, btn)
        box = btn.bounding_box()
        if box:
            move_to(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        pause(1, 2)
        page.keyboard.press("Escape")  # closes the More menu if it was opened
        return "dry_run_ok"
    human_click(page, btn)

    send = page.get_by_role("button", name=re.compile(r"^Send( without a note| now)?$")).first
    try:
        send.wait_for(timeout=8000)
        pause(0.8, 2.5)  # 'read' the modal
        human_click(page, send)
    except PWTimeout:
        pass  # some profiles send immediately with no modal
    pause(1.5, 3)

    if page.get_by_text(re.compile("weekly invitation limit|reached the limit", re.I)).count():
        return "weekly_limit"
    try:  # confirm: this person's button must now say Pending
        page.locator(pending_sel).first.wait_for(state="visible", timeout=6000)
        return "sent"
    except PWTimeout:
        pass
    # After sending via the More menu, LinkedIn only shows Pending on the top card after a reload.
    pause(1, 3)
    page.reload(wait_until="domcontentloaded")
    check_session(page)
    try:
        page.locator(pending_sel).first.wait_for(state="visible", timeout=10000)
        return "sent"
    except PWTimeout:
        return "connect_failed" if first_visible(page, invite_sel) else "sent_unverified"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profile", default="Default", help='Chrome profile folder, e.g. "Profile 1"')
    ap.add_argument("--limit", type=int, default=20, help="max requests this run (LinkedIn caps ~100/week)")
    ap.add_argument("--min-delay", type=float, default=20)
    ap.add_argument("--max-delay", type=float, default=120)
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="do everything up to the invite modal, then cancel; logs nothing")
    args = ap.parse_args()
    if args.selftest:
        return selftest()

    done = load_done()
    todo = [(u, n) for u, n in load_urls().items() if u not in done]
    random.shuffle(todo)  # don't walk the CSVs in order
    todo = todo[: args.limit]
    next_break = random.randint(4, 8)
    print(f"{len(todo)} profiles to process")
    prepare_profile(args.profile)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(LOCAL_DATA), channel="chrome", headless=False,
            args=[f"--profile-directory={args.profile}"], no_viewport=True,
            ignore_default_args=["--use-mock-keychain"],  # real Keychain, so copied cookies decrypt
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto("https://www.linkedin.com/feed/")
        if any(k in page.url for k in ("/login", "/authwall", "/uas/", "/checkpoint/")):
            input("Log in to LinkedIn in the opened window, then press Enter...")
        read_profile(page)  # warm up on the feed first

        for i, (url, name) in enumerate(todo, 1):
            try:
                page.goto(url, wait_until="domcontentloaded")
                check_session(page)
                page.locator("main").first.wait_for(timeout=15000)
                read_profile(page)
                status = connect(page, args.dry_run)
            except SessionLost as e:
                print(f"[{i}/{len(todo)}] {name} -> session lost ({e}); stopping, nothing logged for this lead")
                break
            except PWTimeout:
                status = "timeout"
            except Exception as e:
                if any(k in str(e).lower() for k in SESSION_ERRORS):
                    print(f"[{i}/{len(todo)}] {name} -> network/session error, stopping: {str(e)[:120]}")
                    break
                status = f"error: {e.__class__.__name__}"
            print(f"[{i}/{len(todo)}] {name} -> {status}")
            if status == "weekly_limit":
                break
            if not args.dry_run:
                log(url, name, status)
            if i == len(todo):
                break
            if i == next_break:  # occasional longer break, like stepping away
                secs = random.uniform(180, 480)
                print(f"   taking a {secs / 60:.1f} min break")
                idle_wait(page, secs, secs)
                next_break += random.randint(4, 8)
            else:
                idle_wait(page, args.min_delay, args.max_delay)
        ctx.close()


def selftest():
    for _ in range(200):
        s, e = (random.uniform(0, 1000), random.uniform(0, 800)), (random.uniform(0, 1000), random.uniform(0, 800))
        pts = bezier_points(s, e, 30)
        assert len(pts) == 30 and abs(pts[-1][0] - e[0]) < 1e-6 and abs(pts[-1][1] - e[1]) < 1e-6
    t0 = time.time(); pause(0.01, 0.05); assert 0.01 <= time.time() - t0 < 0.2

    class FakePage:
        def __init__(self, t): self.t = t
        def title(self): return self.t
    assert profile_name(FakePage("(3) Kunal Banerjee | LinkedIn")) == "Kunal Banerjee"
    assert profile_name(FakePage("Page not found | LinkedIn")) == ""
    print("selftest ok")


if __name__ == "__main__":
    main()
