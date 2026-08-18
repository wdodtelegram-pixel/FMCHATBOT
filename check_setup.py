"""
check_setup.py
==============
PURPOSE
    Run this whenever something will not start. It checks every prerequisite in
    order and tells you which one failed and how to fix it -- instead of you
    reading a traceback to work out that a file is missing.

    Checks, in order:
      1. Python version and architecture (32-bit cannot run pandas)
      2. Required packages are importable
      3. .env file exists and is readable
      4. Bot token is present and correctly shaped
      5. Excel workbook exists, opens, and has the right columns
      6. Every category in config.py actually matches rows in the sheet

    It makes NO network calls, so it works offline and cannot be rate-limited.

RUN WITH
    python check_setup.py
"""

import sys
from pathlib import Path

import console  # noqa: F401 -- imported for its side effect: UTF-8 console output

PASS = "[ OK ]"
FAIL = "[FAIL]"
WARN = "[WARN]"

failures = []
warnings = []


def report(ok: bool, label: str, detail: str = "", fix: str = "", warn_only: bool = False):
    """PURPOSE: print one aligned result line, and remember failures for the summary."""
    mark = PASS if ok else (WARN if warn_only else FAIL)

    print(f"{mark}  {label}")
    if detail:
        print(f"        {detail}")
    if not ok:
        if fix:
            print(f"        FIX: {fix}")
        (warnings if warn_only else failures).append(label)
    print()


print("=" * 70)
print("  FM CHATBOT — SETUP CHECK")
print("=" * 70)
print()

# ---------------------------------------------------------------------------
# 1. Python version and architecture
# ---------------------------------------------------------------------------
# PURPOSE: pandas has published no 32-bit Windows wheels since Python 3.12,
# so a 32-bit interpreter cannot run this project at all.
version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
is_64bit = sys.maxsize > 2**32

report(
    is_64bit,
    "Python is 64-bit",
    f"Python {version}, {'64-bit' if is_64bit else '32-BIT'} at {sys.executable}",
    "Install 64-bit Python from python.org, then re-select the interpreter "
    "in VS Code (Ctrl+Shift+P -> Python: Select Interpreter).",
)

report(
    sys.version_info >= (3, 9),
    "Python version is supported",
    f"Found {version}, need 3.9 or newer",
    "Install a newer Python from python.org.",
)

# ---------------------------------------------------------------------------
# 2. Required packages
# ---------------------------------------------------------------------------
# PURPOSE: an ImportError halfway through bot.py is confusing. Check them all
# up front and report every missing one at once, not just the first.
for module, package in [
    ("telegram", "python-telegram-bot[job-queue]"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("dotenv", "python-dotenv"),
]:
    try:
        imported = __import__(module)
        ver = getattr(imported, "__version__", "unknown")
        report(True, f"Package '{package}' installed", f"version {ver}")
    except ImportError:
        report(
            False,
            f"Package '{package}' installed",
            "not found",
            f"pip install {package}",
        )

# APScheduler is what powers the background refresh. It arrives via the
# [job-queue] extra, which people often miss by installing plain
# python-telegram-bot.
try:
    __import__("apscheduler")
    report(True, "Background scheduler available", "APScheduler present")
except ImportError:
    report(
        False,
        "Background scheduler available",
        "APScheduler missing — the 5-minute auto-refresh will be skipped",
        'pip install "python-telegram-bot[job-queue]"',
        warn_only=True,
    )

# ---------------------------------------------------------------------------
# 3, 4. Config and token
# ---------------------------------------------------------------------------
try:
    import config

    report(
        config.ENV_PATH.exists(),
        ".env file found",
        f"Looked in: {config.ENV_PATH}",
        "Create a file named exactly '.env' (leading dot, no extension) "
        "in the same folder as bot.py.",
    )

    token_ok, token_message = config.validate_token()
    if token_ok:
        # Show only the first few characters -- never print a full token to a
        # terminal that might end up in a screenshot or a shared log.
        masked = config.BOT_TOKEN[:8] + "..." + config.BOT_TOKEN[-4:]
        report(True, "Bot token looks valid", f"Token: {masked}")
    else:
        report(False, "Bot token looks valid", token_message.replace("\n", "\n        "))

except Exception as exc:
    report(False, "config.py imports cleanly", str(exc), "Check config.py for syntax errors.")
    config = None

# ---------------------------------------------------------------------------
# 5, 6. Data source (Excel workbook OR published Google CSV)
# ---------------------------------------------------------------------------
if config is not None:
    source = getattr(config, "DATA_SOURCE", "excel")
    print(f"[ -- ]  Data source: {source}\n")

    if source == "google_csv":
        # ---- Option C: published Google Form CSV ----
        url = getattr(config, "GOOGLE_CSV_URL", "")
        report(
            bool(url),
            "GOOGLE_CSV_URL is set",
            "Published-to-web CSV link found" if url else "not set",
            "Add GOOGLE_CSV_URL to your .env. See GOOGLE_FORM_SETUP.md.",
        )
        report(
            url.endswith("output=csv") if url else False,
            "CSV URL looks correct",
            "ends with output=csv" if url.endswith("output=csv") else
            "should end with 'output=csv'",
            "Re-copy the link from File -> Share -> Publish to web -> CSV.",
            warn_only=True,
        )
        proceed = bool(url)
    else:
        # ---- Excel workbook ----
        path: Path = config.EXCEL_PATH
        if not path.exists():
            report(
                False,
                "Excel workbook found",
                f"Not found at: {path}",
                "Run 'python create_workbook.py' to generate a starter file, "
                "or set EXCEL_PATH in .env to your OneDrive copy.",
            )
            proceed = False
        else:
            size = path.stat().st_size
            report(True, "Excel workbook found", f"{path}  ({size:,} bytes)")
            if size < 1000:
                report(
                    False,
                    "Workbook has real content",
                    f"Only {size} bytes — this looks like a OneDrive placeholder",
                    "Right-click the file in Explorer -> 'Always keep on this device'.",
                    warn_only=True,
                )
            proceed = True

    if proceed:
        try:
            import excel_loader

            df = excel_loader.load_jobs(force=True)
            noun = "responses" if source == "google_csv" else "job row(s)"
            report(True, "Data loads and parses", f"{len(df)} {noun} loaded")

            missing = [c for c in config.REQUIRED_COLUMNS if c not in df.columns]
            hint = (
                "Name your Google Form questions exactly like the columns, "
                "or fill in FORM_COLUMN_MAP in config.py."
                if source == "google_csv"
                else "Check the header row spelling in the 'Jobs' sheet against config.py."
            )
            report(
                not missing,
                "All required columns present",
                f"Missing: {missing}" if missing else f"{len(df.columns)} columns found",
                hint,
            )

            # PURPOSE: catch the silent failure where a category renders an
            # empty list because the Category text does not match config exactly.
            print("        Category match check:")
            any_empty = False
            for key, meta in config.CATEGORIES.items():
                rows = excel_loader.get_jobs_for_category(key)
                status = "ok" if len(rows) else "NO MATCHING ROWS"
                if not len(rows):
                    any_empty = True
                print(f"          /{key:<12} '{meta['excel']}'  ->  {len(rows)} row(s)  {status}")
            print()

            if any_empty:
                warnings.append("Some categories matched no rows")
                print("        NOTE: a category with no rows is fine if you have not")
                print("        added those jobs yet. If you HAVE, the Category value")
                print("        must match config.CATEGORIES['...']['excel'] exactly.")
                print()

        except Exception as exc:
            hint = (
                "Check the sheet is still Published to web and GOOGLE_CSV_URL is correct."
                if source == "google_csv"
                else "Make sure the sheet tab is named 'Jobs' and the file is not "
                "currently open in Excel with unsaved changes."
            )
            report(False, "Data loads and parses", f"{type(exc).__name__}: {exc}", hint)

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
print("=" * 70)
if failures:
    print(f"  {len(failures)} CHECK(S) FAILED — fix these before running bot.py")
    for item in failures:
        print(f"    - {item}")
elif warnings:
    print(f"  ALL CHECKS PASSED, with {len(warnings)} warning(s)")
    for item in warnings:
        print(f"    - {item}")
    print("\n  You can run:  python bot.py")
else:
    print("  ALL CHECKS PASSED")
    print("\n  You can run:  python bot.py")
print("=" * 70)

sys.exit(1 if failures else 0)