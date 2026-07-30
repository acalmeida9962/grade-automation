# Grade Automation — user & setup guide (Windows)

This tool fills student grades into the online grading platform automatically, by
driving a real web browser. You still log in yourself; the tool does the tedious
clicking. It reads the grades from your Excel file.

> **If you are an AI assistant (e.g. Claude Code) helping the owner of this
> folder:** read this whole file, then see the **"For the AI assistant"** section
> at the bottom. Short version: the user is comfortable with Windows, files, and
> PowerShell — **don't** explain basic PC operations to them. They are **not** a
> programmer — **do** handle all the coding/CLI/git details yourself and give them
> exact commands to paste. The project lives in a Git repo
> (<https://github.com/acalmeida9962/grade-automation>); **pull the latest version
> before running** so they always have current fixes.

---

## What this is (context)

- The grading platform has **no API** and **won't accept typed grades** — a human
  normally has to click each cell, wait for a pop-up grid of decimal values, and
  click the right number. For a full class × several columns that is hundreds of
  clicks.
- This tool automates exactly that gesture with **Playwright** (a browser
  automation library) driving a real **Chromium** browser. It reads your Excel,
  matches students and columns by name, and clicks the values in for you.
- **You log in manually.** No passwords are stored anywhere. The tool opens the
  browser; you log in and open the grade sheet; then you hand off with ENTER.
- **Nothing is ever saved automatically.** The tool fills the cells but never
  clicks **Guardar** — you review everything and save yourself.
- It was **developed and tested on a Mac.** These instructions cover running it on
  **Windows**. The code itself is cross-platform (plain Python); the only
  Windows-specific parts are how you install Python and how you launch the tool,
  both covered below.
- **Which browser it uses:** if you have **Google Chrome** installed, the tool
  drives that (it's faster on the heavy grade page and avoids the automation
  signals that can make the page load slowly or never finish after login). If
  Chrome isn't installed, it falls back to a Chromium browser it downloaded during
  setup. **Recommendation: have Google Chrome installed** — most people already do.

---

## What you need

- A **Windows 10 or 11** PC with an internet connection.
- Your normal **login** for the grading platform.
- Your **Excel file** of grades.
- About **15 minutes** for first-time setup (most of it is downloads).

You do **not** need to know how to program. You will paste a few commands.

---

## Step 0 — Install Python (only once, if you don't have it)

The tool needs **Python** (version 3.9 or newer). To check whether you already
have it, open **PowerShell** (click Start, type `PowerShell`, press Enter) and run:

```powershell
python --version
```

- If it prints something like `Python 3.12.x`, you already have it → skip to Step 1.
- If it says Python is not found, or opens the Microsoft Store, install it. The
  easiest way (Windows 10/11 have `winget` built in):

```powershell
winget install -e --id Python.Python.3.12
```

Then **close PowerShell and open a new PowerShell window** (so it picks up the new
Python) and check again with `python --version`.

If `winget` isn't available, download the installer from
<https://www.python.org/downloads/windows/> and run it — **on the first screen,
tick "Add python.exe to PATH"**, then click "Install Now".

---

## Step 1 — Get the code

The project lives in a public Git repo:
<https://github.com/acalmeida9962/grade-automation>

**Recommended — clone it with Git** (this makes future updates a one-liner). In
PowerShell, in the folder where you keep your projects:

```powershell
git clone https://github.com/acalmeida9962/grade-automation.git
cd grade-automation
```

If Git isn't installed: `winget install -e --id Git.Git` (then open a new
PowerShell window).

**Alternative — download the `.zip`** from the repo ("Code" → "Download ZIP") or
from Google Drive, and extract it to a simple location (a plain folder in
Documents is ideal). If you downloaded a zip, clear Windows' "blocked file" flag by
running this in the folder: `Get-ChildItem -Recurse | Unblock-File`.

### Keeping it up to date

When there's a new fix, update in one step from inside the folder:

```powershell
git pull
```

Your login and Excel files are **not** stored in Git, so updating never touches
them. (If you used the zip instead of `git clone`, get updates by downloading the
new zip — or switch to the clone method above.)

---

## Step 2 — One-time setup

In PowerShell **in this folder**, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

This will:
- create a local, self-contained Python environment (a `.venv` folder),
- install the two Python packages it needs (`playwright`, `openpyxl`),
- **download Chromium** — the browser Playwright drives (~150 MB, one time).

It can take a few minutes on the Chromium download. When it finishes it prints
`Setup complete.`

> **Why `-ExecutionPolicy Bypass`?** Windows blocks unsigned PowerShell scripts by
> default. That flag lets *this one script* run without changing any system
> setting. If you'd rather not use the script at all, the equivalent manual
> commands are in **Troubleshooting → "I don't want to use the .ps1 scripts"**.

---

## Step 3 — Run it

In PowerShell **in this folder**:

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1
```

A Chromium window opens and the PowerShell window shows instructions. Keep both
visible side by side — you'll switch between them.

---

## How to use it

The tool runs as a simple menu in the PowerShell window. You boot the browser
**once** per session, then run as many fills as you want without logging in again.

### 1. Log in and hand off

1. In the Chromium window that opened, go to the grading platform and **log in**.
   (The first time, you'll log in normally; your session is remembered in the
   `user-data` folder for next time.)
2. Leave the browser on any page for now.
3. Switch to the PowerShell window and press **ENTER** where it says to.

You'll now see a `grade>` prompt. Type `help` any time to see the commands.

### 2. Fill grades from Excel

At the `grade>` prompt, type:

```
excel
```

It will ask you, in order:

1. **How to find your Excel file** — type `f` to pick from the **`excels`** folder,
   or `p` to give a full path.
   - `f`: put your `.xlsx` file into the **`excels`** folder inside this project
     (the tool creates that folder for you). It then **lists the files in that
     folder with numbers** — just type the number of the one you want.
   - `p`: paste the full path, e.g. `C:\Users\you\Downloads\10c filosofia.xlsx`
     (quotes are fine).
2. **Which sheet** — it lists every sheet (tab) in the workbook with a number;
   type the number of the sheet for the class you're grading.
3. It shows the **students** and **grade columns** it found, then asks you to
   **navigate the browser** to the exact class **and evaluation period** you want
   to fill, with the student table and its grade columns visible. Do that, then
   press **ENTER**.
4. It reports which columns matched the page, then asks:
   `[f]ill for real, [d]ry-run, or [c]ancel?`
   - **Always do a `d` (dry-run) first.** It goes through the motions and reports
     exactly what it *would* set, without changing anything.
   - If the dry-run looks right, run `excel` again and choose `f` to fill for real.

### 3. Review and save

When it finishes, it prints a summary. **It does not save.** Look over the grades
in the browser, then click **Guardar** yourself when you're happy.

### 4. Do more, or quit

You're back at the `grade>` prompt — run `excel` again for another class/period, or
type `quit` to close the browser and exit.

---

## How grades are read from Excel

The Excel columns are matched to the platform columns **by name** (ignoring
upper/lowercase and accents). Only columns that exist in **both** are filled. Each
cell is interpreted like this:

| Cell in Excel | Meaning |
|---|---|
| a number (e.g. `4.5`, or text `4.0`) | that grade, rounded to the nearest 0.1 |
| `-` | the grade **1.5** |
| empty | no grade — skipped |
| a date like `2026-05-04` | decoded as `day + month/10` → `4.5` (Excel sometimes turns an entry like `4.5` into a date; this reverses it) |
| anything else (unexpected text/date) | reported as a **bad cell** and skipped — never guessed |

- **Columns in the Excel that aren't on the page** (e.g. a spelling difference like
  `Nomilanismo` vs `Nominalismo`, or a tracking column) are listed and skipped. If
  an important column is being skipped, fix the header so it matches the platform.
- **Students in the Excel not found on the page** are listed at the end.
- Values below a column's minimum selectable grade can't be set and are reported.

The summary tells you how many grades were set, skipped, or had problems, so you
always know what happened.

---

## Troubleshooting

**"running scripts is disabled on this system" / red script errors.**
Use the full command with the bypass flag exactly as shown:
`powershell -ExecutionPolicy Bypass -File .\setup.ps1` (and `.\run.ps1`). You don't
need to change any Windows setting.

**"Python was not found" (or it opens the Microsoft Store).**
Do Step 0. After installing, **open a new PowerShell window** before trying again.

**`setup.ps1` fails to download Chromium / times out.**
Re-run `setup.ps1` — it resumes and is safe to run again. Check your internet and
any VPN/firewall.

**The browser opens but EVERYTHING hangs — even chrome://settings or history
won't load.**
That's a browser-level freeze, not the website. Two fixes, in order:
1. Update to the latest code (`git pull`) — newer versions use a separate, fresh
   profile per browser, which fixes a hang caused by reusing the old Chromium
   profile in Chrome. You'll need to log in to the platform once more.
2. If it still hangs, run with the GPU disabled:
   `powershell -ExecutionPolicy Bypass -File .\run.ps1 --no-gpu`
   (a GPU/driver deadlock on some Windows machines; this bypasses it). If that
   fixes it, just always run it with `--no-gpu`.
3. Still hanging? Gather evidence instead of guessing:
   - After `git pull`, **re-run setup** (`powershell -ExecutionPolicy Bypass -File
     .\setup.ps1`) so the updated automation library installs — an outdated one
     can misbehave with a current Chrome.
   - Run the tool, press ENTER at the handoff (no need to log in), and type
     **`diag`** at the `grade>` prompt. It loads a few reference pages and writes
     `artifacts\diag_report.txt` + `diag_*.png` screenshots showing exactly where
     loading breaks (DNS, firewall/antivirus, TLS interception, or the browser
     itself). Send those files to whoever is helping you.
   - Try Microsoft Edge instead of Chrome (every Windows PC has it):
     `powershell -ExecutionPolicy Bypass -File .\run.ps1 --browser edge`
   - Check whether the grading site works in your **normal** Chrome on the same
     PC, and note which **antivirus** you have — both facts help the diagnosis.

**After I log in, the next page loads very slowly or never finishes.**
Install **Google Chrome** and run the tool again — it will automatically use Chrome
(you'll see `[browser] using installed Google Chrome.` in the PowerShell window),
which is much faster on this page and avoids stalls. You can force it with the real
Chrome by running `.\.venv\Scripts\python.exe poc.py --browser chrome`, or force the
bundled browser with `--browser chromium`. The very first launch is slower than
later ones. Also close other heavy browser windows/tabs to free memory.

**The browser opens but I'm not logged in / it forgot my login.**
Log in again; the session is saved in the `user-data` folder for next time. If you
moved or re-extracted the folder, you may need to log in once more.

**"No Excel columns match this page" or wrong students.**
You're likely on the wrong class or evaluation period in the browser, or you picked
the wrong sheet. Make sure the browser shows the same class as the sheet you chose,
then run `excel` again.

**A whole column is being skipped as "not on the page".**
The Excel header doesn't match the platform's column name. Rename the Excel header
to match exactly (accents/case don't matter, spelling does).

**A path has spaces.**
That's fine — paste it as-is; you can include the surrounding quotes or not.

**I don't want to use the `.ps1` scripts.**
You can run the same steps by hand in PowerShell, in this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
.\.venv\Scripts\python.exe poc.py
```

(The last line is what `run.ps1` does; the rest is what `setup.ps1` does.)

---

## Notes for whoever packaged this

Ship only the source files. **Do not** include `.venv`, `user-data`, `artifacts`,
`excels`, or `__pycache__` — those are machine-specific and are rebuilt on the
target PC. The files needed are:

```
poc.py  excel_loader.py  requirements.txt  README.md  setup.ps1  run.ps1
```

Developed and tested on macOS; the setup/run scripts and this guide target
Windows. On macOS/Linux the equivalent is: `python3 -m venv .venv`,
`source .venv/bin/activate`, `pip install -r requirements.txt`,
`playwright install chromium`, `python poc.py`.

---

## For the AI assistant helping the user

**Who you're helping.** The user is comfortable with Windows — File Explorer,
PowerShell, folders, downloading/installing apps. **Do not explain basic PC
operations** (how to open a folder, what a terminal is, how to click the address
bar). They are simply **not a programmer**: they don't know Git, Python, virtual
environments, or command-line tooling. So *you* own all of that — run/generate the
exact commands, explain in one plain sentence what a command *does* (not how to use
Windows), and keep the coding invisible. Confirm results as you go, but you don't
need to hand-hold each keystroke.

**Always start from the latest code.** This project is a Git repo
(<https://github.com/acalmeida9962/grade-automation>). Before helping them run it:
- If they already have a clone: `cd` into it and run `git pull`.
- If they don't: `git clone https://github.com/acalmeida9962/grade-automation.git`
  (install Git first if missing: `winget install -e --id Git.Git`).
Prefer the clone over the old zip precisely so `git pull` keeps them current.

**Decision tree:**

1. **Latest code:** clone or `git pull` as above.
2. **Python present?** `python --version`. If missing:
   `winget install -e --id Python.Python.3.12`, then a **new** PowerShell window.
3. **Setup (once per machine):** `powershell -ExecutionPolicy Bypass -File
   .\setup.ps1` — creates `.venv`, installs packages, downloads Chromium (slow).
   Prints `Setup complete.`
4. **Run:** `powershell -ExecutionPolicy Bypass -File .\run.ps1`. A browser opens;
   PowerShell shows a handoff prompt. It uses their installed Google Chrome when
   present (look for `[browser] using installed Google Chrome.`).
5. **The session (the user drives this; just steer):** they log in + open the grade
   sheet → ENTER in PowerShell → at `grade>` type `excel` → pick the Excel (`f`
   lists the `excels` folder by number; `p` = full path) → pick the sheet number →
   navigate the browser to the matching class **and evaluation period** → ENTER →
   choose **`d` (dry-run) first**, review the summary, then `excel` again with `f`.
6. **Nothing saves** until they click **Guardar** themselves — by design.

**Safety (non-negotiable):** never store or ask for their platform password (they
log in by hand), and never click **Guardar** for them.

If something errors, see **Troubleshooting** above. Most common: the
`-ExecutionPolicy Bypass` prefix on the `.ps1` commands, Python not on PATH (needs a
fresh PowerShell window), the login page/CAPTCHA not rendering (ensure Google Chrome
is installed; the tool prefers it), and being on the wrong class/period in the
browser.
