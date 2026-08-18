"""
config.py
=========
PURPOSE
    One single place to change anything that is "settings" rather than "logic".
    Nothing in this file does work -- it only declares constants that the rest
    of the bot imports. If you rename an Excel column, add a 6th job category,
    or move the workbook, this is the ONLY file you should need to edit.
"""

import os
import re
from pathlib import Path

from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# 1. Secrets / environment
# ---------------------------------------------------------------------------
# PURPOSE: read the bot token from a .env file so the token never gets
# hard-coded into source control. Create a file named ".env" next to bot.py
# containing:  TELEGRAM_BOT_TOKEN=123456789:AAF-real-token-from-BotFather
#
# override=True matters: without it, a stale TELEGRAM_BOT_TOKEN left in your
# Windows environment variables would silently win over the .env file, and you
# would be debugging a token you cannot see anywhere in the project.
ENV_PATH = Path(__file__).parent / ".env"
load_dotenv(ENV_PATH, override=True)


def _clean_token(raw: str) -> str:
    """
    PURPOSE: strip the three things people accidentally include around a token.

      1. Surrounding quotes -- "123:ABC" or '123:ABC'. python-dotenv keeps
         these as part of the value in some cases, and Telegram then rejects
         a token that looks correct to the eye.
      2. Whitespace, including a trailing space before the newline.
      3. A leftover "TELEGRAM_BOT_TOKEN=" prefix, which happens when someone
         pastes the whole line into the value field.
    """
    token = (raw or "").strip()

    if token.upper().startswith("TELEGRAM_BOT_TOKEN="):
        token = token.split("=", 1)[1].strip()

    if len(token) >= 2 and token[0] == token[-1] and token[0] in "\"'":
        token = token[1:-1].strip()

    return token


def _env(name: str, default: str = "") -> str:
    """
    PURPOSE: read an environment variable, treating a BLANK value as absent.

    This fixes a subtle and very common .env bug. os.getenv("EXCEL_PATH", fallback)
    returns the fallback only when the variable is completely UNSET. If .env
    contains the line

        EXCEL_PATH=

    then the variable IS set -- to an empty string -- and the fallback is
    silently ignored. Path("") then evaluates to "." (the current directory),
    and pandas fails later with a confusing IsADirectoryError.

    Treating blank as absent means a commented-out or emptied setting behaves
    the way anyone would expect.
    """
    value = os.getenv(name)
    if value is None or not value.strip():
        return default
    return value.strip()


BOT_TOKEN = _clean_token(os.getenv("TELEGRAM_BOT_TOKEN", ""))

# A real BotFather token looks like:  8123456789:AAHk3jd8sLq...
# That is: digits, a colon, then ~35 characters of letters/digits/-/_
TOKEN_PATTERN = re.compile(r"^\d{6,16}:[A-Za-z0-9_-]{30,}$")

# PURPOSE: catch the single most common setup mistake -- copying .env.example
# to .env and never replacing the placeholder text. Without this check the
# error only surfaces as a 100-line async traceback from inside the Telegram
# library, which tells a beginner nothing useful.
#
# IMPORTANT: these are only ever tested against a token that has ALREADY failed
# the format check. Testing them against every token would cause false
# positives -- the characters after the colon are random, so a genuine token
# can contain a short sequence like "here" purely by chance, and blocking a
# working token is a far worse failure than missing a placeholder.
_PLACEHOLDER_MARKERS = (
    "paste",
    "your-token",
    "your_token",
    "yourtoken",
    "replace",
    "example",
    "xxxx",
    "<",
    "here",
    "botfather",
)


def validate_token() -> tuple[bool, str]:
    """
    PURPOSE: check the token BEFORE any network call, and return a message that
    says exactly what to do about it.

    Order matters. The format check runs FIRST: anything matching the real
    token pattern is accepted immediately and no further tests can reject it.
    Only once a token has failed that pattern do we try to work out WHY, so the
    diagnostic messages can never block a valid token.

    Returns (is_valid, message_to_show). bot.py prints the message and exits
    rather than letting the Telegram library raise InvalidToken.
    """
    # --- 1. Nothing supplied at all ----------------------------------------
    if not BOT_TOKEN:
        return False, (
            "TELEGRAM_BOT_TOKEN is empty or missing.\n\n"
            f"Expected a .env file at:\n  {ENV_PATH}\n\n"
            f"That file currently {'EXISTS' if ENV_PATH.exists() else 'DOES NOT EXIST'}.\n\n"
            "Create it with a single line:\n"
            "  TELEGRAM_BOT_TOKEN=8123456789:AAHk...your-real-token\n\n"
            "Get the token by messaging @BotFather on Telegram and sending /newbot."
        )

    # --- 2. Correct shape wins outright ------------------------------------
    if TOKEN_PATTERN.match(BOT_TOKEN):
        return True, "Token format looks valid."

    # --- 3. It failed. Work out the most useful thing to say. --------------
    lowered = BOT_TOKEN.lower()
    if any(marker in lowered for marker in _PLACEHOLDER_MARKERS):
        return False, (
            "Your .env still contains PLACEHOLDER text, not a real token.\n\n"
            f"Found:  {BOT_TOKEN}\n"
            f"In:     {ENV_PATH}\n\n"
            "Replace that whole value with the token @BotFather gave you.\n"
            "It should look like:  8123456789:AAHk3jd8sLq_realCharacters\n\n"
            "How to get one:\n"
            "  1. Open Telegram and search for  @BotFather\n"
            "  2. Send  /newbot\n"
            "  3. Choose a display name, then a username ending in 'bot'\n"
            "  4. Copy the long token it replies with into .env"
        )

    if ":" not in BOT_TOKEN:
        return False, (
            "The token is missing its colon.\n\n"
            f"Found:  {BOT_TOKEN[:20]}...\n\n"
            "A real token has the form  <digits>:<letters and digits>\n"
            "Make sure you copied the ENTIRE string from @BotFather."
        )

    prefix, _, suffix = BOT_TOKEN.partition(":")
    problems = []
    if not prefix.isdigit():
        problems.append(f"the part before the colon ('{prefix}') should be digits only")
    elif not 6 <= len(prefix) <= 16:
        problems.append(f"the part before the colon is {len(prefix)} digits, expected 6-16")
    if len(suffix) < 30:
        problems.append(f"the part after the colon is only {len(suffix)} chars, expected 30+")

    return False, (
        "The token is the wrong shape and will be rejected by Telegram.\n\n"
        + "\n".join(f"  - {p}" for p in problems)
        + "\n\nCommon causes:\n"
        "  - the token was cut short when copying\n"
        "  - a space or line break crept into the middle\n"
        "  - you copied the bot's username instead of its token"
    )


# ---------------------------------------------------------------------------
# 2. Where the maintenance workbook lives
# ---------------------------------------------------------------------------
# PURPOSE: per your architecture diagram, the FM Officer edits the workbook in
# Excel Online. The OneDrive *desktop* client mirrors that cloud file down to a
# local folder automatically. The bot only ever reads the LOCAL copy -- that is
# why no Microsoft login or Graph API credential is needed.
#
# Change EXCEL_PATH to your real synced location, e.g.
#   C:/Users/<you>/OneDrive - MyCompany/FM/maintenance_jobs.xlsx
BASE_DIR = Path(__file__).parent
DEFAULT_EXCEL_PATH = BASE_DIR / "data" / "maintenance_jobs.xlsx"

# expanduser() lets you write "~/OneDrive/..." instead of a full path.
# _env() means a blank "EXCEL_PATH=" line correctly falls back to the default
# rather than resolving to the current directory.
_excel_setting = _env("EXCEL_PATH")
EXCEL_PATH = Path(_excel_setting).expanduser() if _excel_setting else DEFAULT_EXCEL_PATH

# Name of the worksheet (tab) inside the workbook that holds the job rows.
SHEET_NAME = _env("SHEET_NAME", "Jobs")

# OPTIONAL fallback: if the machine running the bot does NOT have OneDrive
# desktop sync (e.g. a cloud VM), put a OneDrive "anyone with the link can
# view" share URL here and the loader will download the file over HTTP instead.
# Leave blank to use the local synced file only.
ONEDRIVE_SHARE_LINK = _env("ONEDRIVE_SHARE_LINK")

# ---------------------------------------------------------------------------
# 2b. Data source selector  (Option C: Google Form -> published CSV)
# ---------------------------------------------------------------------------
# PURPOSE: choose WHERE job data comes from without touching any other file.
#
#   "excel"       (default) -> read the local .xlsx workbook, as before.
#   "google_csv"  -> fetch a Google Sheet that has been "Published to web"
#                    as CSV. This is how Option C works with NO Google Cloud
#                    account: the sheet is served as a plain public URL and the
#                    bot just downloads it over HTTP.
DATA_SOURCE = _env("DATA_SOURCE", "excel").lower()

# The "Publish to web" CSV link for your Google Form's response sheet.
# Looks like:
#   https://docs.google.com/spreadsheets/d/e/2PACX-.../pub?gid=0&single=true&output=csv
# Only used when DATA_SOURCE = "google_csv".
GOOGLE_CSV_URL = _env("GOOGLE_CSV_URL")

# Google Forms is APPEND-ONLY: every submission adds a new row, it never edits
# an old one. So to "update" a job the officer resubmits the form with the same
# Job ID, and the bot keeps only the NEWEST row per Job ID. Set to false to show
# every raw submission instead.
DEDUPE_BY_JOB_ID = _env("DEDUPE_BY_JOB_ID", "true").lower() in ("1", "true", "yes", "on")

# The automatic column Google Forms adds to every response sheet. Used to work
# out which submission is newest. In non-English locales this header may differ;
# change it here if your sheet calls it something else.
FORM_TIMESTAMP_COLUMN = _env("FORM_TIMESTAMP_COLUMN", "Timestamp")

# How the Form timestamp is written. Singapore/UK forms use day-first
# (DD/MM/YYYY); US forms use month-first. Only affects tie-breaking of updates.
FORM_TIMESTAMP_DAYFIRST = _env("FORM_TIMESTAMP_DAYFIRST", "true").lower() in ("1", "true", "yes", "on")

# FORM_COLUMN_MAP is defined further down, AFTER the column-name constants it
# refers to (see section 3b).

# How often (seconds) the background refresh job re-reads the workbook.
# The loader ALSO checks the file timestamp on every command, so this is really
# just a safety net that keeps the cache warm.
# Wrapped in try/except so a typo like "5 minutes" falls back to the default
# instead of crashing the bot at import time with a ValueError.
try:
    REFRESH_INTERVAL_SECONDS = int(_env("REFRESH_INTERVAL_SECONDS", "300"))
except ValueError:
    REFRESH_INTERVAL_SECONDS = 300

# ---------------------------------------------------------------------------
# 3. Excel column names
# ---------------------------------------------------------------------------
# PURPOSE: the bot refers to columns through these constants instead of raw
# strings. If a colleague renames "Est Completion" to "Target Date" in the
# spreadsheet, you fix it here once and every module follows.
COL_JOB_ID = "Job ID"
COL_CATEGORY = "Category"
COL_DESCRIPTION = "Job Description"
COL_LOCATION = "Location"
COL_TYPE = "Maintenance Type"      # Preventive / Corrective
COL_STATUS = "Status"              # Yet to Start / In Progress / Completed
COL_EST_COMPLETION = "Estimated Completion"
COL_NEXT_JOB = "Next Job Date"
COL_UPDATED = "Last Updated"

REQUIRED_COLUMNS = [
    COL_JOB_ID,
    COL_CATEGORY,
    COL_DESCRIPTION,
    COL_LOCATION,
    COL_TYPE,
    COL_STATUS,
    COL_EST_COMPLETION,
    COL_NEXT_JOB,
]

# ---------------------------------------------------------------------------
# 3b. Google Form question -> column mapping  (Option C only)
# ---------------------------------------------------------------------------
# PURPOSE: map your Google Form QUESTION TITLES to the column names above.
# Defined here, after the COL_* constants, so you can reference them.
#
# Leave this EMPTY if you name the form questions exactly like the columns
# (recommended -- the setup guide tells you to do this). Only fill it in if you
# want friendlier question wording on the form. Example:
#     FORM_COLUMN_MAP = {
#         "Which job ID?": COL_JOB_ID,
#         "What type of maintenance?": COL_TYPE,
#     }
FORM_COLUMN_MAP: dict = {}

# ---------------------------------------------------------------------------
# 4. The job categories
# ---------------------------------------------------------------------------
# PURPOSE: this dictionary is the single source of truth for the whole bot.
# It simultaneously drives (a) which /commands exist, (b) the text of the
# /start help menu, (c) the inline keyboard buttons, and (d) the value the bot
# matches against the "Category" column in Excel.
#
# TO ADD A 6TH CATEGORY: add one entry here. No other file needs changing --
# the command handler, the menu and the buttons are all generated from this.
#
#   key       -> the Telegram command, i.e. /acmv
#   "name"    -> human-readable heading used in message titles
#   "excel"   -> the exact string written in the Excel "Category" column
#   "blurb"   -> short description shown in the /start menu
#   "emoji"   -> small visual marker to keep messages scannable
CATEGORIES = {
    "acmv": {
        "name": "ACMV",
        "excel": "ACMV",
        "blurb": "ACMV servicing",
        "emoji": "❄️",
    },
    "lift": {
        "name": "Lift & Escalator",
        "excel": "Lift & Escalator",
        "blurb": "Lift and escalator servicing",
        "emoji": "🛗",
    },
    "electrical": {
        "name": "Electrical",
        "excel": "Electrical",
        "blurb": "Electrical maintenance",
        "emoji": "⚡",
    },
    "fire": {
        "name": "Fire Protection",
        "excel": "Fire Protection",
        "blurb": "Fire protection services",
        "emoji": "🧯",
    },
    "plumbing": {
        "name": "Plumbing & Sanitary",
        "excel": "Plumbing & Sanitary",
        "blurb": "Plumbing and sanitary services",
        "emoji": "🚰",
    },
}

# ---------------------------------------------------------------------------
# 5. Status display
# ---------------------------------------------------------------------------
# PURPOSE: map the three allowed status values to an icon so a user can skim a
# long dashboard and immediately see what is done and what has not started.
# The keys are lower-cased before lookup, so "In Progress" and "in progress"
# both work -- staff typing into Excel are inconsistent and that is fine.
STATUS_ICONS = {
    "yet to start": "⚪",
    "in progress": "🟡",
    "completed": "🟢",
}

# Canonical spellings, used to normalise whatever staff typed in the sheet.
STATUS_CANONICAL = {
    "yet to start": "Yet to Start",
    "not started": "Yet to Start",
    "in progress": "In Progress",
    "ongoing": "In Progress",
    "completed": "Completed",
    "complete": "Completed",
    "done": "Completed",
}

TYPE_CANONICAL = {
    "preventive": "Preventive",
    "preventative": "Preventive",
    "pm": "Preventive",
    "corrective": "Corrective",
    "cm": "Corrective",
}