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
import json
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
# Separate login profiles per browser: a profile created by bundled Chromium must
# never be reopened by Google Chrome (different build) — that can hang/corrupt it.
USER_DATA_CHROME = HERE / "user-data-chrome"
USER_DATA_CHROMIUM = HERE / "user-data-chromium"
USER_DATA_EDGE = HERE / "user-data-edge"
ARTIFACTS = HERE / "artifacts"
EXCELS_DIR = HERE / "excels"
SETTINGS_FILE = HERE / "settings.json"  # per-machine prefs (gitignored)


def _load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 — missing/corrupt file: start fresh
        return {}


def _save_settings(**kw) -> None:
    s = _load_settings()
    s.update(kw)
    try:
        SETTINGS_FILE.write_text(json.dumps(s, indent=2), encoding="utf-8")
    except Exception:  # noqa: BLE001 — never let prefs break the flow
        pass

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
    print("  Inicia sesión y abre la planilla de notas en el navegador.")
    print("  Cuando la tabla de estudiantes con sus notas esté visible,")
    print("  vuelve aquí y presiona ENTER para continuar.")
    print("=" * 70)
    input("  Presiona ENTER cuando estés listo... ")


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
        return ("MISS", "no hay celda para este estudiante/columna en la página")

    before = cell.first.evaluate("e => e.value || e.getAttribute('placeholder') || ''")
    cell.first.scroll_into_view_if_needed()
    cell.first.click()

    try:
        panel.wait_for(state="visible", timeout=5000)
    except Exception:  # noqa: BLE001
        return ("FAIL", "no se abrió el selector de notas")

    option = panel.locator("td.item-td:not(.disabled) a").filter(
        has_text=re.compile(rf"^{re.escape(value_str)}$")
    )
    if option.count() == 0:
        any_match = panel.locator("a").filter(
            has_text=re.compile(rf"^{re.escape(value_str)}$")
        )
        why = ("está deshabilitada (bajo el mínimo de la columna)"
               if any_match.count() else "no está en el selector")
        dismiss_picker(page)
        _wait_hidden(panel)
        return ("FAIL", f"la nota {value_str} {why}")

    if dry_run:
        dismiss_picker(page)  # close via X (no-op on empty cells); never selects
        _wait_hidden(panel)
        return ("OK", f"pondría {value_str} (antes: {before!r})")

    option.first.click()
    if not _wait_hidden(panel):
        dismiss_picker(page)
        _wait_hidden(panel)
    after = cell.first.evaluate("e => e.value || e.getAttribute('placeholder') || ''")
    if after.strip() == value_str:
        return ("SET", f"{value_str} (antes: {before!r})")
    return ("WARN", f"se intentó {value_str}, la celda quedó {after!r} (antes: {before!r})")


def _wait_hidden(panel, timeout: int = 2500) -> bool:
    try:
        panel.wait_for(state="hidden", timeout=timeout)
        return True
    except Exception:  # noqa: BLE001
        return False


def do_fill(page, dry_run: bool) -> None:
    """Milestone 0 (avanzado): fill the hardcoded test grades into TARGET_COLUMN."""
    mode = ("SIMULACRO (no se pondrá ninguna nota)" if dry_run
            else "REAL (se pondrán notas; NO se guarda)")
    print(f"\n=== PRUEBA — {mode} ===")

    target_pk = resolve_target_pk(page)
    if not target_pk:
        print(f"ABORTADO: la columna {TARGET_COLUMN!r} no está en esta vista.")
        return
    print(f"Columna objetivo {TARGET_COLUMN!r} -> pkActividad={target_pk}")

    roster = build_roster(page)
    print(f"Estudiantes en la página: {len(roster)}.")

    for raw_name, grade in TARGET_GRADES.items():
        pkm = roster.get(normalize(raw_name))
        if not pkm:
            print(f"  MISS  {raw_name!r}: no está en la página")
            continue
        tag, detail = set_grade(page, pkm, target_pk, f"{grade:.1f}", dry_run)
        print(f"  {tag:<4}  {raw_name!r}: {detail}")

    print("\nListo. NO se hizo clic en 'Guardar' — revisa los valores en el navegador.")


def _ask(prompt: str) -> str:
    return input(prompt).strip()


def _choose_excel_file() -> pathlib.Path | None:
    """Ask for the Excel file: last used (ENTER), ./excels/ listing, or a path."""
    last = _load_settings().get("last_excel")
    last_path = pathlib.Path(last) if last else None
    if last_path and last_path.is_file():
        src = _ask(f"¿[ENTER] usar el último ({last_path.name}), "
                   "[f] elegir de la carpeta 'excels', o [p] escribir una ruta? ").lower()
        if src == "":
            return last_path
    else:
        src = _ask("¿Buscar el Excel por [f] archivo de la carpeta 'excels' "
                   "o [p] ruta completa? ").lower()

    if src in ("p", "path", "r", "ruta"):
        raw = _ask("Ruta completa al archivo .xlsx: ").strip().strip('"').strip("'")
        path = pathlib.Path(raw).expanduser()
    elif src in ("f", "filename", "a", "archivo"):
        EXCELS_DIR.mkdir(exist_ok=True)
        excels = sorted(
            (p for p in EXCELS_DIR.iterdir()
             if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
             and not p.name.startswith("~$")),  # skip Excel lock files
            key=lambda p: p.name.lower(),
        )
        if not excels:
            print(f"\nNo hay archivos de Excel en: {EXCELS_DIR}")
            print("Pon tu archivo .xlsx en esa carpeta y vuelve a intentar.")
            return None
        print(f"\nArchivos de Excel en {EXCELS_DIR}:")
        for i, p in enumerate(excels, 1):
            print(f"  {i}. {p.name}")
        pick = _ask("Elige un archivo (número o nombre exacto): ")
        if pick.isdigit() and 1 <= int(pick) <= len(excels):
            path = excels[int(pick) - 1]
        else:
            match = next((p for p in excels if p.name.lower() == pick.lower()), None)
            if match is None:
                print(f"Opción no válida: {pick!r}")
                return None
            path = match
    else:
        print("Cancelado (elige 'f' o 'p').")
        return None
    if not path.is_file():
        print(f"Archivo no encontrado: {path}")
        return None
    return path


def _choose_sheet(path: pathlib.Path) -> str | None:
    sheets = excel_loader.list_sheets(str(path))
    if not sheets:
        print("El archivo no tiene hojas.")
        return None
    last = _load_settings().get("last_sheet")
    default = last if last in sheets else None
    print("\nHojas del archivo:")
    for i, s in enumerate(sheets, 1):
        mark = "   <- última usada" if s == default else ""
        print(f"  {i}. {s}{mark}")
    hint = f" [ENTER = {default}]" if default else ""
    pick = _ask(f"Elige la hoja (número o nombre exacto){hint}: ")
    if pick == "" and default:
        return default
    if pick.isdigit() and 1 <= int(pick) <= len(sheets):
        return sheets[int(pick) - 1]
    match = next((s for s in sheets if s.lower() == pick.lower()), None)
    if match:
        return match
    print(f"Opción no válida: {pick!r}")
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
        print(f"No se pudo leer la hoja: {exc}")
        return

    # Recordar este archivo y hoja para la próxima vez.
    _save_settings(last_excel=str(path), last_sheet=parsed.sheet)

    print(f"\nHoja {parsed.sheet!r}: {len(parsed.students)} estudiantes, "
          f"columna de nombres {parsed.name_header!r}.")
    print(f"Columnas de notas en el Excel ({len(parsed.grade_headers)}):")
    for h in parsed.grade_headers:
        print(f"    - {h}")

    # --- Pedir al usuario que navegue a la planilla, luego continuar. ---
    print("\n" + "-" * 70)
    print("  Ahora ve en el navegador a la planilla que quieres llenar:")
    print("  la clase Y el período de evaluación correctos, con la tabla de")
    print("  estudiantes y sus columnas de notas visibles.")
    print("  Luego vuelve aquí y presiona ENTER.")
    print("-" * 70)
    input("  Presiona ENTER cuando la página esté lista... ")

    # --- Emparejar columnas/estudiantes del Excel con lo que hay en la página. ---
    page_columns = resolve_columns(page)          # normalized name -> pkActividad
    roster = build_roster(page)                   # normalized name -> pkMatricula

    matched: list[tuple[str, str]] = []           # (excel header, pkActividad)
    unmatched_cols: list[str] = []
    for h in parsed.grade_headers:
        pka = page_columns.get(normalize(h))
        (matched.append((h, pka)) if pka else unmatched_cols.append(h))

    print(f"\nColumnas emparejadas con la página ({len(matched)}): "
          + (", ".join(h for h, _ in matched) or "ninguna"))
    if unmatched_cols:
        print(f"Columnas que NO están en la página (se omiten): {', '.join(unmatched_cols)}")
    if not matched:
        print("Ninguna columna del Excel coincide con esta página. "
              "¿Estás en la clase/período correcto?")
        return

    # --- Confirmar antes de escribir. ---
    choice = _ask(f"\n¿[f] llenar de verdad, [d] simulacro, o [c] cancelar? "
                  f"({len(parsed.students)} estudiantes × {len(matched)} columnas) ").lower()
    if choice in ("c", "cancel", "cancelar", ""):
        print("Cancelado.")
        return
    dry_run = choice in ("d", "dry", "simulacro")

    # --- Llenar. ---
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
                problems.append(f"  MALA  {st.raw_name} / {header}: {payload}")
                continue
            value_str = excel_loader.format_grade(payload)
            tag, detail = set_grade(page, pkm, pka, value_str, dry_run)
            counts[tag] = counts.get(tag, 0) + 1
            if tag in ("FAIL", "WARN", "MISS"):
                problems.append(f"  {tag:<4} {st.raw_name} / {header}: {detail}")

    # --- Resumen. ---
    print("\n=== RESUMEN " + ("(SIMULACRO)" if dry_run else "(REAL)") + " ===")
    verb = "se pondrían" if dry_run else "puestas"
    print(f"  notas {verb}: {counts['OK'] + counts['SET']}   "
          f"vacías/omitidas: {counts['skip']}   celdas malas: {counts['bad']}   "
          f"fallidas: {counts['FAIL']}   advertencias: {counts['WARN']}")
    if missing_students:
        print(f"  estudiantes del Excel que no están en la página ({len(missing_students)}): "
              + ", ".join(missing_students))
    if problems:
        print("  problemas:")
        print("\n".join(problems))
    if not dry_run:
        print("\nListo. NO se hizo clic en 'Guardar' — revisa los valores y guarda tú.")


DIAG_TARGETS = [
    ("basic web", "https://example.com"),
    ("google / TLS", "https://www.google.com"),
    ("reCAPTCHA script", "https://www.google.com/recaptcha/api.js"),
    ("grading platform", "https://accesos.colombiaevaluadora.co/"),
]


def do_diag(page) -> None:
    """Connectivity/rendering diagnostics.

    Navigates the current tab through reference pages and records timings,
    network failures (net::ERR_* codes) and console errors. The resulting
    artifacts/diag_report.txt + diag_*.png screenshots pinpoint WHERE loading
    breaks: cert errors -> antivirus TLS interception; connection reset ->
    firewall/AV blocking; name-not-resolved -> DNS/proxy; nothing at all (even
    example.com) -> the automated browser itself is broken on this machine.
    """
    import platform as _platform
    import time as _time

    ARTIFACTS.mkdir(exist_ok=True)
    lines: list[str] = []

    def report(msg: str = "") -> None:
        lines.append(msg)
        print(msg)

    try:
        from importlib.metadata import version as _pkg_version
        pw_ver = _pkg_version("playwright")
    except Exception:  # noqa: BLE001
        pw_ver = "unknown"

    report("=== DIAGNÓSTICO ===")
    report(f"os: {_platform.platform()}")
    report(f"python: {sys.version.split()[0]}   playwright: {pw_ver}")
    try:
        b = page.context.browser
        report(f"versión del navegador: {b.version if b else 'desconocida (contexto persistente)'}")
    except Exception as exc:  # noqa: BLE001
        report(f"versión del navegador: <error: {exc}>")
    try:
        report(f"userAgent: {page.evaluate('navigator.userAgent')}")
        report(f"navigator.webdriver: {page.evaluate('navigator.webdriver')}")
    except Exception as exc:  # noqa: BLE001
        report(f"FALLÓ la ejecución de JS en la pestaña actual: {type(exc).__name__}: {exc}")
        report("  ^ si esto falla, la pestaña/navegador está muerto — tampoco "
               "funcionará nada de red.")

    failures: list[str] = []
    console_errors: list[str] = []

    def on_reqfail(req) -> None:
        failures.append(f"{req.failure or '?'}  {req.url[:120]}")

    def on_console(msg) -> None:
        if msg.type == "error":
            console_errors.append(msg.text[:200])

    page.on("requestfailed", on_reqfail)
    page.on("console", on_console)
    report("\nEsto abre unas páginas de prueba en la pestaña actual (~1-2 min).")

    for i, (label, url) in enumerate(DIAG_TARGETS, 1):
        failures.clear()
        console_errors.clear()
        report(f"\n--- [{i}] {label}: {url}")
        t0 = _time.time()
        try:
            resp = page.goto(url, timeout=20000, wait_until="load")
            dt = _time.time() - t0
            status = resp.status if resp else "?"
            report(f"    cargó en {dt:.1f}s  status={status}  título={page.title()[:60]!r}")
        except Exception as exc:  # noqa: BLE001
            dt = _time.time() - t0
            first = str(exc).splitlines()[0][:160]
            report(f"    FALLÓ tras {dt:.1f}s: {type(exc).__name__}: {first}")
        page.wait_for_timeout(2500)  # let stragglers fail and get recorded
        for f in failures[:8]:
            report(f"    fallo-red: {f}")
        if len(failures) > 8:
            report(f"    ... y {len(failures) - 8} fallos de red más")
        for c in console_errors[:5]:
            report(f"    error-consola: {c}")
        try:
            page.screenshot(path=str(ARTIFACTS / f"diag_{i}.png"))
        except Exception as exc:  # noqa: BLE001
            report(f"    falló la captura: {type(exc).__name__}")

    page.remove_listener("requestfailed", on_reqfail)
    page.remove_listener("console", on_console)

    # ¿Este navegador está gestionado por una organización (políticas)?
    try:
        page.goto("chrome://policy", timeout=10000)
        page.wait_for_timeout(1500)
        page.screenshot(path=str(ARTIFACTS / "diag_policy.png"))
        report("\nGuardada captura de chrome://policy (diag_policy.png) — muestra "
               "si el navegador está gestionado por una organización.")
    except Exception as exc:  # noqa: BLE001
        report(f"\nno se pudo capturar chrome://policy: {type(exc).__name__}")

    (ARTIFACTS / "diag_report.txt").write_text("\n".join(lines), encoding="utf-8")
    report(f"\nReporte guardado en: {ARTIFACTS / 'diag_report.txt'}")
    report("Envía diag_report.txt y las capturas diag_*.png para analizarlo.")


def do_record(page) -> None:
    """Flight recorder: passively log network/console/frame activity while the
    USER reproduces the problem in the browser, then dump everything.

    Unlike `diag` (which loads generic test pages), this captures the actual
    failing flow — e.g. the login page whose reCAPTCHA widget never renders."""
    import time as _time

    ARTIFACTS.mkdir(exist_ok=True)
    ctx = page.context
    events: list[str] = []
    t0 = _time.time()
    MAX_EVENTS = 3000

    def ts() -> str:
        return f"{_time.time() - t0:7.1f}s"

    def add(line: str) -> None:
        if len(events) < MAX_EVENTS:
            events.append(line)

    def on_req(req) -> None:
        # Documents, scripts, xhr and iframes tell the story; skip image noise.
        if req.resource_type in ("document", "script", "xhr", "fetch", "websocket"):
            add(f"{ts()} ->  {req.resource_type:9} {req.method:4} {req.url[:150]}")

    def on_resp(resp) -> None:
        if resp.status >= 400:
            add(f"{ts()} <-  HTTP {resp.status}  {resp.url[:150]}")

    def on_fail(req) -> None:
        add(f"{ts()} XX  {req.failure or 'failed'}  {req.url[:150]}")

    def hook_page(pg) -> None:
        pg.on("console", lambda m: add(f"{ts()} console-{m.type}: {m.text[:180]}")
              if m.type in ("error", "warning") else None)
        pg.on("pageerror", lambda e: add(f"{ts()} pageerror: {str(e)[:180]}"))
        pg.on("frameattached", lambda fr: add(f"{ts()} frame+  {fr.url[:150]}"))
        pg.on("framenavigated", lambda fr: add(f"{ts()} frame~  {fr.url[:150]}"))

    ctx.on("request", on_req)
    ctx.on("response", on_resp)
    ctx.on("requestfailed", on_fail)
    ctx.on("page", hook_page)
    for pg in ctx.pages:
        hook_page(pg)

    print("\n=== GRABANDO ===")
    print("  Ahora ve al navegador y reproduce el problema tal como pasa")
    print("  (por ejemplo: abre el login, espera el captcha que no aparece,")
    print("  haz clic en lo que se cuelga...). Tómate tu tiempo.")
    input("  Cuando termines, vuelve aquí y presiona ENTER para detener... ")

    ctx.remove_listener("request", on_req)
    ctx.remove_listener("response", on_resp)
    ctx.remove_listener("requestfailed", on_fail)
    ctx.remove_listener("page", hook_page)

    lines: list[str] = ["=== RECORDING REPORT ===", f"duration: {_time.time()-t0:.0f}s",
                        f"events captured: {len(events)}", ""]
    lines += events

    # Final page state: URL, frames (is the reCAPTCHA iframe even there?), JS.
    lines += ["", "=== FINAL PAGE STATE ==="]
    try:
        lines.append(f"url: {page.url}")
        lines.append(f"title: {page.title()[:80]!r}")
        for fr in page.frames:
            lines.append(f"frame: {fr.url[:150]}")
        has_recaptcha_frame = any("recaptcha" in fr.url.lower() for fr in page.frames)
        lines.append(f"reCAPTCHA iframe present: {has_recaptcha_frame}")
        lines.append(f"typeof grecaptcha: {page.evaluate('typeof grecaptcha')}")
        lines.append(f"document.readyState: {page.evaluate('document.readyState')}")
    except Exception as exc:  # noqa: BLE001
        lines.append(f"<final-state error: {type(exc).__name__}: {exc}>")
    try:
        page.screenshot(path=str(ARTIFACTS / "record.png"))
        lines.append("screenshot: record.png")
    except Exception as exc:  # noqa: BLE001
        lines.append(f"screenshot failed: {type(exc).__name__}")

    (ARTIFACTS / "record_report.txt").write_text("\n".join(lines), encoding="utf-8")
    print(f"\nGuardados {len(events)} eventos en {ARTIFACTS / 'record_report.txt'}")
    print("Últimos 25 eventos:")
    for line in events[-25:]:
        print("  " + line)
    print("\nEnvía record_report.txt y record.png para analizarlo.")


SESSION_HELP = """
Comandos:
  excel    leer las notas de un Excel y llenar la página  (o solo presiona ENTER)
  ayuda    mostrar esta ayuda
  salir    cerrar y terminar
"""

# Comandos avanzados (ocultos del menú). 'diag' y 'record' sirvieron para
# diagnosticar problemas del navegador; se mantienen por si hacen falta.
ADVANCED_HELP = """
Comandos avanzados (ocultos):
  diag     diagnóstico de conexión; guarda un reporte en ./artifacts/
  record   graba la actividad de red mientras reproduces un problema a mano
  dry      prueba interna (simulacro) sobre una columna fija
  fill     prueba interna (real) sobre una columna fija
  probe    volcar el selector de notas a ./artifacts/
  inspect  volcar el DOM + captura a ./artifacts/
"""


def run_session(page) -> None:
    """One boot: log in once, then run steps repeatedly against the live page."""
    wait_for_handoff()
    print(SESSION_HELP)
    actions = {
        "excel": lambda: do_excel(page),
        # avanzados / ocultos:
        "dry": lambda: do_fill(page, dry_run=True),
        "fill": lambda: do_fill(page, dry_run=False),
        "probe": lambda: do_probe(page),
        "inspect": lambda: do_inspect(page),
        "diag": lambda: do_diag(page),
        "record": lambda: do_record(page),
    }
    while True:
        try:
            cmd = input("\nnotas> ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            break
        if cmd in ("salir", "quit", "q", "exit"):
            break
        if cmd in ("ayuda", "help", "h", "?"):
            print(SESSION_HELP)
            continue
        if cmd == "avanzado":
            print(ADVANCED_HELP)
            continue
        # ENTER vacío = ejecutar 'excel' (lo único que usa el día a día).
        if cmd == "":
            cmd = "excel"
        action = actions.get(cmd)
        if not action:
            print(f"Comando desconocido: {cmd!r}. Escribe 'ayuda' "
                  "(o presiona ENTER para 'excel').")
            continue
        try:
            action()
        except Exception as exc:  # noqa: BLE001 — keep the session alive on errors
            print(f"[error] {type(exc).__name__}: {exc}")


def launch_context(p, browser: str, no_gpu: bool = False):
    """Open the browser we drive.

    Prefers the real installed Google Chrome (channel="chrome"): it renders the
    heavy grade page much faster than Playwright's bundled Chromium, uses a browser
    the user already has (no 150 MB download needed), and avoids quirks that can
    make the page load slowly or stall after login. Falls back to bundled Chromium
    if Chrome isn't installed.

    Each browser gets its own login-profile folder (a Chromium profile reopened by
    Chrome can hang the whole browser). `no_gpu=True` adds --disable-gpu, the
    classic fix for a headed Chrome that hangs with nothing rendering.
    """
    def build_opts(user_data_dir: pathlib.Path) -> dict:
        args = [
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
        ]
        if no_gpu:
            args += ["--disable-gpu", "--disable-software-rasterizer"]
        user_data_dir.mkdir(exist_ok=True)
        return dict(
            user_data_dir=str(user_data_dir),
            headless=False,
            viewport=None,  # use the real window size
            ignore_default_args=["--enable-automation"],
            args=args,
        )

    if browser in ("auto", "chrome", "edge"):
        if browser == "edge":
            channel, udir, name = "msedge", USER_DATA_EDGE, "Microsoft Edge"
        else:
            channel, udir, name = "chrome", USER_DATA_CHROME, "Google Chrome"
        try:
            ctx = p.chromium.launch_persistent_context(
                channel=channel, **build_opts(udir)
            )
            print(f"[navegador] usando {name} instalado"
                  + (" (GPU desactivada)." if no_gpu else "."))
            return ctx
        except Exception as exc:  # noqa: BLE001 — browser not installed / not found
            if browser in ("chrome", "edge"):
                raise
            print(f"[navegador] Google Chrome no disponible ({type(exc).__name__}); "
                  "usando el Chromium interno.")
    ctx = p.chromium.launch_persistent_context(**build_opts(USER_DATA_CHROMIUM))
    print("[navegador] usando el Chromium interno" + (" (GPU desactivada)." if no_gpu else "."))
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
        choices=["auto", "chrome", "edge", "chromium"],
        default="auto",
        help="Which browser to drive: 'auto' (real Chrome if present, else bundled "
             "Chromium), 'chrome' (force real Chrome), 'edge' (installed Microsoft "
             "Edge — every Windows PC has it), or 'chromium' (force bundled).",
    )
    parser.add_argument(
        "--no-gpu",
        action="store_true",
        help="Add --disable-gpu. Try this if the browser opens but hangs and nothing "
             "loads (a GPU/renderer deadlock on some Windows setups).",
    )
    parser.add_argument(
        "--attach",
        nargs="?",
        const="9222",
        metavar="PORT",
        help="Do not launch a browser; attach to one YOU started with "
             "--remote-debugging-port=PORT (default 9222). The browser behaves "
             "exactly like normal browsing, which sidesteps launch-mode problems. "
             "See the README for the exact command to start Edge/Chrome this way.",
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
    EXCELS_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        if args.attach:
            endpoint = f"http://127.0.0.1:{args.attach}"
            try:
                browser = p.chromium.connect_over_cdp(endpoint)
            except Exception as exc:  # noqa: BLE001
                print(f"No se pudo conectar a un navegador en el puerto {args.attach} "
                      f"({type(exc).__name__}).")
                print("Primero abre Edge con un puerto de depuración (en PowerShell):\n")
                print('  Start-Process msedge -ArgumentList '
                      '"--remote-debugging-port=9222",'
                      '"--user-data-dir=$env:LOCALAPPDATA\\edge-grades"\n')
                print("(o lo mismo con 'chrome' en vez de 'msedge'), "
                      "y vuelve a ejecutar con --attach.")
                return 1
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            context = None  # we did not launch it; never close the user's browser
            print(f"[navegador] conectado a tu navegador en el puerto {args.attach}.")
        else:
            context = launch_context(p, args.browser, no_gpu=args.no_gpu)
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
                input("\nListo. Presiona ENTER para cerrar el navegador... ")
        finally:
            if context is not None:
                context.close()
            else:
                print("Desconectado; tu navegador queda abierto.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
