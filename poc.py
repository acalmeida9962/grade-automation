"""
Grade automation — fills grades into the (API-less) grading platform by driving a
real browser: click a cell -> a decimal picker pops up -> click the value.

Runs on macOS, Windows, and Linux (Python + Playwright). Login is manual: the
script opens a real (non-headless) Chromium; you log in and open the grade sheet,
then hand off. A persistent profile in ./user-data keeps you logged in.

Default mode is an interactive "session" — one browser boot, then type commands:

  * excel    read grades from an Excel sheet and fill the page (main feature)
  * dry      Milestone-0 test fill (dry-run: confirm values, select nothing)
  * fill     Milestone-0 test fill (live: hardcoded test grades)
  * probe    click one cell and dump the picker markup to ./artifacts/
  * inspect  dump the full DOM + screenshot to ./artifacts/

Nothing is ever saved automatically: the script never clicks "Guardar".
"""

from __future__ import annotations

import argparse
import pathlib
import re
import sys
import unicodedata

from playwright.sync_api import sync_playwright

import excel_loader


def _force_utf8_console() -> None:
    """Windows consoles often default to a legacy code page; force UTF-8 so
    accented names/columns (Autoevaluación, Ñ, ...) print without crashing."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001 — older Python / non-standard streams
            pass

HERE = pathlib.Path(__file__).resolve().parent
USER_DATA_DIR = HERE / "user-data"
ARTIFACTS = HERE / "artifacts"
EXCELS_DIR = HERE / "excels"

# ---- Milestone 0 hardcoded test target (from the screenshots) ---------------
# Column identified by its hover tooltip name.
TARGET_COLUMN = "UNA PRUEBA"
# Student full name (as shown in NOMBRES) -> grade to set (0.1 steps, 0.0-5.0).
TARGET_GRADES = {
    "ACOSTA ARRIETA YOVANNI ANDRES": 3.0,
    "ALVAREZ BLANCO MIGUEL ANGEL": 3.5,
    "AMADOR AMADOR MICHELLE ELIANA": 4.0,
}


def normalize(name: str) -> str:
    """Uppercase, strip accents, collapse whitespace — for robust name matching."""
    nfkd = unicodedata.normalize("NFKD", name)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(no_accents.upper().split())


def wait_for_handoff() -> None:
    print("\n" + "=" * 70)
    print("  Log in and open the grade sheet in the browser window.")
    print("  When the student table with grades is visible, come back here")
    print("  and press ENTER to continue.")
    print("=" * 70)
    input("  Press ENTER when ready... ")


def do_inspect(page) -> None:
    ARTIFACTS.mkdir(exist_ok=True)

    # 1. Full DOM.
    html = page.content()
    (ARTIFACTS / "dom.html").write_text(html, encoding="utf-8")

    # 2. Screenshot (full page).
    page.screenshot(path=str(ARTIFACTS / "page.png"), full_page=True)

    # 3. Heuristics to help build selectors.
    lines: list[str] = []

    def report(msg: str) -> None:
        lines.append(msg)
        print(msg)

    report(f"DOM bytes: {len(html)}")

    # Elements carrying a title attribute (candidate for tooltip column names).
    titled = page.locator("[title]")
    n_titled = titled.count()
    report(f"\nElements with a title attribute: {n_titled}")
    for i in range(min(n_titled, 40)):
        el = titled.nth(i)
        try:
            t = el.get_attribute("title")
            tag = el.evaluate("e => e.tagName.toLowerCase()")
            report(f"  [{i}] <{tag}> title={t!r}")
        except Exception as exc:  # noqa: BLE001
            report(f"  [{i}] <error reading: {exc}>")

    # aria-label candidates.
    aria = page.locator("[aria-label]")
    n_aria = aria.count()
    report(f"\nElements with aria-label: {n_aria}")
    for i in range(min(n_aria, 40)):
        try:
            report(f"  [{i}] aria-label={aria.nth(i).get_attribute('aria-label')!r}")
        except Exception:  # noqa: BLE001
            pass

    # Locate each target student and dump the ancestor HTML so we can see the
    # row/cell structure.
    for raw_name in TARGET_GRADES:
        loc = page.get_by_text(raw_name, exact=False)
        cnt = loc.count()
        report(f"\nName {raw_name!r}: {cnt} text match(es)")
        if cnt:
            try:
                # Dump a few ancestor levels of outerHTML for the first match.
                outer = loc.first.evaluate(
                    """el => {
                        let node = el;
                        for (let i = 0; i < 4 && node.parentElement; i++) node = node.parentElement;
                        return node.outerHTML;
                    }"""
                )
                snippet = outer[:4000]
                (ARTIFACTS / f"row_{normalize(raw_name).replace(' ', '_')}.html").write_text(
                    outer, encoding="utf-8"
                )
                report(f"  ancestor outerHTML (first 400 chars):\n  {snippet[:400]}")
            except Exception as exc:  # noqa: BLE001
                report(f"  <error dumping ancestor: {exc}>")

    (ARTIFACTS / "inspect_report.txt").write_text("\n".join(lines), encoding="utf-8")
    report(f"\nSaved artifacts to: {ARTIFACTS}")
    report("Files: dom.html, page.png, inspect_report.txt, row_*.html")


def do_probe(page) -> None:
    """Click one grade cell to capture the dynamically-built decimal picker."""
    ARTIFACTS.mkdir(exist_ok=True)
    lines: list[str] = []

    def report(msg: str) -> None:
        lines.append(msg)
        print(msg)

    # 1. Column header map: data-tooltip (name) -> pkActividad (from onclick).
    headers = page.evaluate(
        r"""() => {
            const out = [];
            document.querySelectorAll('button.button-radius[onclick*="Planilla.ver"]').forEach(b => {
                const m = (b.getAttribute('onclick')||'').match(/Planilla\.ver\((\d+)\)/);
                out.push({name: b.getAttribute('data-tooltip'), pk: m ? m[1] : null, label: (b.textContent||'').trim()});
            });
            return out;
        }"""
    )
    report("=== Columns in this view (name -> pkActividad) ===")
    target_pk = None
    for h in headers:
        mark = ""
        if h["name"] and normalize(h["name"]) == normalize(TARGET_COLUMN):
            target_pk = h["pk"]
            mark = "   <-- TARGET"
        report(f"  {h['label']:>3}  pk={h['pk']}  {h['name']!r}{mark}")
    if target_pk:
        report(f"\nTarget column {TARGET_COLUMN!r} -> pkActividad={target_pk}")
    else:
        report(f"\nWARNING: target column {TARGET_COLUMN!r} not found in this view.")

    # 2. Click an editable cell (prefer one in the target column so the captured
    #    picker matches the real scale we'll be filling).
    if target_pk:
        cell = page.locator(
            f"input.data[onclick*='Panel.calificar'][data-pkactividad='{target_pk}']"
        ).first
    else:
        cell = page.locator("input.data[onclick*='Panel.calificar']").first
    cell.wait_for(state="visible", timeout=10000)
    cell.scroll_into_view_if_needed()
    info = cell.evaluate(
        "e => ({pkm: e.getAttribute('data-pkmatricula'), pka: e.getAttribute('data-pkactividad'), "
        "ph: e.getAttribute('placeholder'), val: e.value, escala: e.getAttribute('data-escala')})"
    )
    report(f"\n=== Clicking cell pkm={info['pkm']} pka={info['pka']} "
           f"placeholder={info['ph']!r} escala={info['escala']} ===")
    cell.click()
    page.wait_for_timeout(700)

    # 3. Find the visible decimal option cells the picker just rendered.
    picker = page.evaluate(
        r"""() => {
            const isVisible = el => !!(el.offsetParent || el.getClientRects().length);
            const opts = [];
            document.querySelectorAll('*').forEach(el => {
                if (el.children.length) return;              // leaf nodes only
                const t = (el.textContent||'').trim();
                if (/^[0-5]\.[0-9]$/.test(t) && isVisible(el)) {
                    opts.push(el);
                }
            });
            if (!opts.length) return {found: 0};
            // Common ancestor of the option cells = the picker container.
            let anc = opts[0];
            const contains = (a, b) => a.contains(b);
            outer: for (let up = 0; up < 8 && anc.parentElement; up++) {
                if (opts.every(o => contains(anc, o))) break outer;
                anc = anc.parentElement;
            }
            const desc = el => ({
                tag: el.tagName.toLowerCase(),
                cls: el.getAttribute('class'),
                id: el.id || null,
                onclick: el.getAttribute('onclick'),
                html: el.outerHTML.slice(0, 200),
            });
            return {
                found: opts.length,
                sampleOptions: opts.slice(0, 6).map(desc),
                container: {
                    tag: anc.tagName.toLowerCase(), cls: anc.getAttribute('class'),
                    id: anc.id || null, html: anc.outerHTML.slice(0, 600),
                },
            };
        }"""
    )
    report(f"\n=== Picker: {picker.get('found')} decimal option cells found ===")
    if picker.get("found"):
        report(f"container: <{picker['container']['tag']}> id={picker['container']['id']!r} "
               f"class={picker['container']['cls']!r}")
        report(f"container html (600): {picker['container']['html']}")
        report("sample option cells:")
        for o in picker["sampleOptions"]:
            report(f"  <{o['tag']}> class={o['cls']!r} onclick={o['onclick']!r}")
            report(f"     {o['html']}")

    page.screenshot(path=str(ARTIFACTS / "modal.png"))
    (ARTIFACTS / "modal_dom.html").write_text(page.content(), encoding="utf-8")

    # 4. Close the picker WITHOUT selecting, via its red X (data-value=-1).
    dismiss_picker(page)
    report("\nClicked the X to close the picker without changing any grade.")

    (ARTIFACTS / "probe_report.txt").write_text("\n".join(lines), encoding="utf-8")
    report(f"\nSaved: modal.png, modal_dom.html, probe_report.txt in {ARTIFACTS}")


def resolve_columns(page) -> dict[str, str]:
    """Map normalized column name -> pkActividad, read from the header buttons."""
    headers = page.evaluate(
        r"""() => {
            const out = [];
            document.querySelectorAll('button.button-radius[onclick*="Planilla.ver"]').forEach(b => {
                const m = (b.getAttribute('onclick')||'').match(/Planilla\.ver\((\d+)\)/);
                out.push({name: b.getAttribute('data-tooltip'), pk: m ? m[1] : null});
            });
            return out;
        }"""
    )
    return {normalize(h["name"]): h["pk"] for h in headers if h["name"] and h["pk"]}


def resolve_target_pk(page) -> str | None:
    """Map TARGET_COLUMN (by normalized name) to its pkActividad."""
    return resolve_columns(page).get(normalize(TARGET_COLUMN))


def build_roster(page) -> dict[str, str]:
    """Map normalized student name -> pkMatricula, read from the rendered rows."""
    rows = page.evaluate(
        r"""() => {
            const out = [];
            document.querySelectorAll('tr').forEach(tr => {
                const nameTd = tr.querySelector('td.recover');
                const inp = tr.querySelector('input.data[data-pkmatricula]');
                if (nameTd && inp) {
                    out.push({name: nameTd.textContent.trim(),
                              pkm: inp.getAttribute('data-pkmatricula')});
                }
            });
            return out;
        }"""
    )
    return {normalize(r["name"]): r["pkm"] for r in rows if r["name"]}


def dismiss_picker(page) -> None:
    """Close the picker via its red X (data-value=-1). Safe no-op on empty cells."""
    close = page.locator("#panel a[data-value='-1']")
    if close.count():
        close.first.click()


def set_grade(page, pkm: str, pka: str, value_str: str, dry_run: bool) -> tuple[str, str]:
    """Set one cell via click -> picker -> click value. Returns (tag, detail).

    tag is one of: MISS, FAIL, OK (dry-run), SET, WARN.
    """
    panel = page.locator("#panel")
    cell = page.locator(
        f"input.data[data-pkmatricula='{pkm}'][data-pkactividad='{pka}']"
    )
    if cell.count() == 0:
        return ("MISS", "no cell for this student/column on the page")

    before = cell.first.evaluate("e => e.value || e.getAttribute('placeholder') || ''")
    cell.first.scroll_into_view_if_needed()
    cell.first.click()

    try:
        panel.wait_for(state="visible", timeout=5000)
    except Exception:  # noqa: BLE001
        return ("FAIL", "picker did not open")

    option = panel.locator("td.item-td:not(.disabled) a").filter(
        has_text=re.compile(rf"^{re.escape(value_str)}$")
    )
    if option.count() == 0:
        any_match = panel.locator("a").filter(
            has_text=re.compile(rf"^{re.escape(value_str)}$")
        )
        why = "disabled (below the column minimum)" if any_match.count() else "not in picker"
        dismiss_picker(page)
        _wait_hidden(panel)
        return ("FAIL", f"value {value_str} {why}")

    if dry_run:
        dismiss_picker(page)  # close via X (no-op on empty cells); never selects
        _wait_hidden(panel)
        return ("OK", f"would set {value_str} (was {before!r})")

    option.first.click()
    if not _wait_hidden(panel):
        dismiss_picker(page)
        _wait_hidden(panel)
    after = cell.first.evaluate("e => e.value || e.getAttribute('placeholder') || ''")
    if after.strip() == value_str:
        return ("SET", f"{value_str} (was {before!r})")
    return ("WARN", f"tried {value_str}, cell now {after!r} (was {before!r})")


def _wait_hidden(panel, timeout: int = 2500) -> bool:
    try:
        panel.wait_for(state="hidden", timeout=timeout)
        return True
    except Exception:  # noqa: BLE001
        return False


def do_fill(page, dry_run: bool) -> None:
    """Milestone 0: fill the hardcoded test grades into TARGET_COLUMN."""
    mode = "DRY-RUN (no grades will be set)" if dry_run else "LIVE (grades will be set; NOT saved)"
    print(f"\n=== FILL — {mode} ===")

    target_pk = resolve_target_pk(page)
    if not target_pk:
        print(f"ABORT: target column {TARGET_COLUMN!r} not found in this view.")
        return
    print(f"Target column {TARGET_COLUMN!r} -> pkActividad={target_pk}")

    roster = build_roster(page)
    print(f"Roster: {len(roster)} students found on the page.")

    for raw_name, grade in TARGET_GRADES.items():
        pkm = roster.get(normalize(raw_name))
        if not pkm:
            print(f"  MISS  {raw_name!r}: not found on page")
            continue
        tag, detail = set_grade(page, pkm, target_pk, f"{grade:.1f}", dry_run)
        print(f"  {tag:<4}  {raw_name!r}: {detail}")

    print("\nDone. No 'Guardar' was clicked — review the values in the browser.")


def _ask(prompt: str) -> str:
    return input(prompt).strip()


def _choose_excel_file() -> pathlib.Path | None:
    """Ask for the Excel file by absolute path or by filename in ./excels/."""
    src = _ask("Read the Excel by [p]ath or [f]ilename? ").lower()
    if src in ("p", "path"):
        raw = _ask("Absolute path to the .xlsx file: ").strip().strip('"').strip("'")
        path = pathlib.Path(raw).expanduser()
    elif src in ("f", "filename"):
        EXCELS_DIR.mkdir(exist_ok=True)
        excels = sorted(
            (p for p in EXCELS_DIR.iterdir()
             if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
             and not p.name.startswith("~$")),  # skip Excel lock files
            key=lambda p: p.name.lower(),
        )
        if not excels:
            print(f"\nNo Excel files found in: {EXCELS_DIR}")
            print("Put your .xlsx file in that folder, then run 'excel' again.")
            return None
        print(f"\nExcel files in {EXCELS_DIR}:")
        for i, p in enumerate(excels, 1):
            print(f"  {i}. {p.name}")
        pick = _ask("Choose a file (number or exact name): ")
        if pick.isdigit() and 1 <= int(pick) <= len(excels):
            path = excels[int(pick) - 1]
        else:
            match = next((p for p in excels if p.name.lower() == pick.lower()), None)
            if match is None:
                print(f"Invalid choice: {pick!r}")
                return None
            path = match
    else:
        print("Cancelled (choose 'p' or 'f').")
        return None
    if not path.is_file():
        print(f"File not found: {path}")
        return None
    return path


def _choose_sheet(path: pathlib.Path) -> str | None:
    sheets = excel_loader.list_sheets(str(path))
    if not sheets:
        print("The workbook has no sheets.")
        return None
    print("\nSheets in this workbook:")
    for i, s in enumerate(sheets, 1):
        print(f"  {i}. {s}")
    pick = _ask("Choose a sheet (number or exact name): ")
    if pick.isdigit() and 1 <= int(pick) <= len(sheets):
        return sheets[int(pick) - 1]
    match = next((s for s in sheets if s.lower() == pick.lower()), None)
    if match:
        return match
    print(f"Invalid choice: {pick!r}")
    return None


def do_excel(page) -> None:
    """Milestone 1: read grades from an Excel sheet and fill the page."""
    print("\n=== EXCEL ===")

    path = _choose_excel_file()
    if not path:
        return
    sheet = _choose_sheet(path)
    if not sheet:
        return

    try:
        parsed = excel_loader.parse_sheet(str(path), sheet)
    except Exception as exc:  # noqa: BLE001
        print(f"Could not parse sheet: {exc}")
        return

    print(f"\nSheet {parsed.sheet!r}: {len(parsed.students)} students, "
          f"name column {parsed.name_header!r}.")
    print(f"Grade columns in the Excel ({len(parsed.grade_headers)}):")
    for h in parsed.grade_headers:
        print(f"    - {h}")

    # --- Ask the user to navigate to the grade page, then hand off. ---
    print("\n" + "-" * 70)
    print("  Now navigate the browser to the grade sheet you want filled:")
    print("  the correct class AND evaluation period, with the student table")
    print("  and its grade columns visible.")
    print("  Then come back here and press ENTER.")
    print("-" * 70)
    input("  Press ENTER when the page is ready... ")

    # --- Match Excel columns/students to what's on the page. ---
    page_columns = resolve_columns(page)          # normalized name -> pkActividad
    roster = build_roster(page)                   # normalized name -> pkMatricula

    matched: list[tuple[str, str]] = []           # (excel header, pkActividad)
    unmatched_cols: list[str] = []
    for h in parsed.grade_headers:
        pka = page_columns.get(normalize(h))
        (matched.append((h, pka)) if pka else unmatched_cols.append(h))

    print(f"\nColumns matched to the page ({len(matched)}): "
          + (", ".join(h for h, _ in matched) or "none"))
    if unmatched_cols:
        print(f"Columns NOT on the page (skipped): {', '.join(unmatched_cols)}")
    if not matched:
        print("No Excel columns match this page. Are you on the right class/period?")
        return

    # --- Confirm before writing. ---
    choice = _ask(f"\n[f]ill for real, [d]ry-run, or [c]ancel? "
                  f"({len(parsed.students)} students × {len(matched)} columns) ").lower()
    if choice in ("c", "cancel", ""):
        print("Cancelled.")
        return
    dry_run = choice in ("d", "dry", "dry-run")

    # --- Fill. ---
    counts = {"SET": 0, "OK": 0, "skip": 0, "bad": 0, "FAIL": 0, "WARN": 0, "MISS": 0}
    problems: list[str] = []
    missing_students: list[str] = []

    for st in parsed.students:
        pkm = roster.get(normalize(st.raw_name))
        if not pkm:
            missing_students.append(st.raw_name)
            continue
        for header, pka in matched:
            kind, payload = excel_loader.to_grade(st.cells.get(header))
            if kind == "skip":
                counts["skip"] += 1
                continue
            if kind == "bad":
                counts["bad"] += 1
                problems.append(f"  BAD   {st.raw_name} / {header}: {payload}")
                continue
            value_str = excel_loader.format_grade(payload)
            tag, detail = set_grade(page, pkm, pka, value_str, dry_run)
            counts[tag] = counts.get(tag, 0) + 1
            if tag in ("FAIL", "WARN", "MISS"):
                problems.append(f"  {tag:<4} {st.raw_name} / {header}: {detail}")

    # --- Summary. ---
    print("\n=== SUMMARY " + ("(DRY-RUN)" if dry_run else "(LIVE)") + " ===")
    verb = "would set" if dry_run else "set"
    print(f"  {verb}: {counts['OK'] + counts['SET']}   "
          f"blank/skipped: {counts['skip']}   bad cells: {counts['bad']}   "
          f"failed: {counts['FAIL']}   warnings: {counts['WARN']}")
    if missing_students:
        print(f"  students in Excel not found on page ({len(missing_students)}): "
              + ", ".join(missing_students))
    if problems:
        print("  issues:")
        print("\n".join(problems))
    if not dry_run:
        print("\nDone. No 'Guardar' was clicked — review the values, then save manually.")


SESSION_HELP = """
Commands (run against the already-open page — no reboot):
  excel    read grades from an Excel sheet and fill the page (Milestone 1)
  dry      Milestone-0 test fill DRY-RUN: confirm the values, select nothing
  fill     Milestone-0 test fill LIVE: set the hardcoded test grades
  probe    click one target-column cell and dump the picker markup
  inspect  dump full DOM + screenshot to ./artifacts/
  help     show this help
  quit     close the browser and exit
"""


def run_session(page) -> None:
    """One boot: log in once, then run steps repeatedly against the live page."""
    wait_for_handoff()
    print(SESSION_HELP)
    actions = {
        "excel": lambda: do_excel(page),
        "dry": lambda: do_fill(page, dry_run=True),
        "fill": lambda: do_fill(page, dry_run=False),
        "probe": lambda: do_probe(page),
        "inspect": lambda: do_inspect(page),
    }
    while True:
        try:
            cmd = input("\ngrade> ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            break
        if cmd in ("quit", "q", "exit"):
            break
        if cmd in ("help", "h", "?", ""):
            print(SESSION_HELP)
            continue
        action = actions.get(cmd)
        if not action:
            print(f"Unknown command {cmd!r}. Type 'help'.")
            continue
        try:
            action()
        except Exception as exc:  # noqa: BLE001 — keep the session alive on errors
            print(f"[error] {type(exc).__name__}: {exc}")


def launch_context(p, browser: str):
    """Open the browser we drive.

    Prefers the real installed Google Chrome (channel="chrome"): it renders the
    heavy grade page much faster than Playwright's bundled Chromium, uses a browser
    the user already has (no 150 MB download needed), and avoids quirks that can
    make the page load slowly or stall after login. Falls back to bundled Chromium
    if Chrome isn't installed.

    We also drop the automation flags (--enable-automation / navigator.webdriver)
    that some sites react to by throttling or hanging.
    """
    opts = dict(
        user_data_dir=str(USER_DATA_DIR),
        headless=False,
        viewport=None,  # use the real window size
        ignore_default_args=["--enable-automation"],
        args=[
            "--disable-blink-features=AutomationControlled",
            "--start-maximized",
            # Stop Chrome from throttling/pausing the page when it thinks the
            # window is backgrounded or occluded. On Windows the native occlusion
            # check often mis-flags the automation window, which freezes page JS
            # and timers — so the CAPTCHA (and the page after login) never finish
            # loading until you open a second tab. These flags disable that.
            "--disable-features=CalculateNativeWinOcclusion",
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
            "--disable-background-timer-throttling",
            # Skip Chrome's first-run / "make default browser" dialogs that can
            # otherwise block the fresh profile on Windows.
            "--no-first-run",
            "--no-default-browser-check",
            # If the window was closed abruptly last time, don't show the
            # "Restore pages? / Chrome didn't shut down correctly" bubble.
            "--hide-crash-restore-bubble",
            # Don't touch the OS keyring/credential store (can prompt or hang).
            "--password-store=basic",
        ],
    )
    if browser in ("auto", "chrome"):
        try:
            ctx = p.chromium.launch_persistent_context(channel="chrome", **opts)
            print("[browser] using installed Google Chrome.")
            return ctx
        except Exception as exc:  # noqa: BLE001 — Chrome not installed / not found
            if browser == "chrome":
                raise
            print(f"[browser] Google Chrome not available ({type(exc).__name__}); "
                  "falling back to bundled Chromium.")
    ctx = p.chromium.launch_persistent_context(**opts)
    print("[browser] using bundled Chromium.")
    return ctx


def main() -> int:
    parser = argparse.ArgumentParser(description="Grade-entry POC (Milestone 0).")
    parser.add_argument(
        "--mode",
        choices=["session", "inspect", "probe", "fill"],
        default="session",
        help="'session' (default) boots the browser once and takes commands interactively.",
    )
    parser.add_argument(
        "--browser",
        choices=["auto", "chrome", "chromium"],
        default="auto",
        help="Which browser to drive: 'auto' (real Chrome if present, else bundled "
             "Chromium), 'chrome' (force real Chrome), or 'chromium' (force bundled).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="fill mode only: open each picker and confirm the value, but do NOT select it.",
    )
    parser.add_argument(
        "--url",
        default="about:blank",
        help="Optional page to open on launch. You can also just navigate manually.",
    )
    args = parser.parse_args()

    _force_utf8_console()
    USER_DATA_DIR.mkdir(exist_ok=True)
    EXCELS_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        context = launch_context(p, args.browser)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.bring_to_front()  # ensure the window is foregrounded/active
        except Exception:  # noqa: BLE001
            pass
        if args.url and args.url != "about:blank":
            page.goto(args.url)

        try:
            if args.mode == "session":
                run_session(page)
            else:
                # One-shot modes (kept for convenience / scripting).
                wait_for_handoff()
                if args.mode == "inspect":
                    do_inspect(page)
                elif args.mode == "probe":
                    do_probe(page)
                else:
                    do_fill(page, dry_run=args.dry_run)
                input("\nDone. Press ENTER to close the browser... ")
        finally:
            context.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
