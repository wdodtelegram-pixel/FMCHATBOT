import sys
from pathlib import Path

PASS = "[ OK ]"
FAIL = "[FAIL]"
WARN = "[WARN]"

failures = []
warnings = []


def report(ok: bool, label: str, detail: str = "", fix: str = "", warn_only: bool = False):
    """PURPOSE: print one aligned result line, and remember failures for the summary."""
    if ok:
        mark = PASS
    else:
        mark = WARN if warn_only else FAIL

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
# 1. Python version
# ---------------------------------------------------------------------------
version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"

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
for module, package, warn_only in [
    ("telegram", "python-telegram-bot[job-queue]", False),
    ("openpyxl", "openpyxl", False),
    ("dotenv", "python-dotenv", False),
    ("pandas", "pandas", True),
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
            warn_only=warn_only,
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
# 5, 6. Workbook
# ---------------------------------------------------------------------------
if config is not None:
    path: Path = config.EXCEL_PATH

    if not path.exists():
        report(
            False,
            "Excel workbook found",
            f"Not found at: {path}",
            "Run 'python create_workbook.py' to generate a starter file, "
            "or set EXCEL_PATH in .env to your OneDrive copy.",
        )
    else:
        size = path.stat().st_size
        report(True, "Excel workbook found", f"{path}  ({size:,} bytes)")

        # A OneDrive cloud-only placeholder has a real path but almost no data.
        # This is a very common and very confusing failure on Windows.
        if size < 1000:
            report(
                False,
                "Workbook has real content",
                f"Only {size} bytes — this looks like a OneDrive placeholder",
                "Right-click the file in Explorer -> 'Always keep on this device'.",
                warn_only=True,
            )

        try:
            import excel_loader

            df = excel_loader.load_jobs(force=True)
            report(True, "Workbook opens and parses", f"{len(df)} job row(s) loaded")

            missing = [c for c in config.REQUIRED_COLUMNS if c not in df.columns]
            report(
                not missing,
                "All required columns present",
                f"Missing: {missing}" if missing else f"{len(df.columns)} columns found",
                "Check the header row spelling in the 'Jobs' sheet against config.py.",
            )

            # PURPOSE: catch the silent failure where a category renders an
            # empty list because the Excel text does not match config exactly.
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
                print("        added those jobs yet. If you HAVE, the 'excel' value in")
                print("        config.CATEGORIES must match the Category column exactly.")
                print()

        except Exception as exc:
            report(
                False,
                "Workbook opens and parses",
                f"{type(exc).__name__}: {exc}",
                "Make sure the sheet tab is named 'Jobs' and the file is not "
                "currently open in Excel with unsaved changes.",
            )

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
