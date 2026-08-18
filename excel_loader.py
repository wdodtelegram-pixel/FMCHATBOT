"""
excel_loader.py
===============
PURPOSE
    This is the "Load Latest Excel Workbook -> Convert Excel into Pandas
    DataFrame" stage of your processing workflow. It is deliberately separated
    from the Telegram code so that the bot layer never touches pandas directly.

    Three jobs:
      1. Read maintenance_jobs.xlsx into a DataFrame.
      2. Cache it, and automatically re-read whenever the file changes on disk
         (which is what happens seconds after the FM Officer saves in Excel
         Online and OneDrive syncs the change down).
      3. Clean the data -- strip stray spaces, normalise Status/Type spellings,
         drop empty rows -- so the rest of the bot can trust what it gets.
"""

from __future__ import annotations

import logging
import threading
from pathlib import Path

import pandas as pd  # type: ignore

import config

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Module-level cache
# ---------------------------------------------------------------------------
# PURPOSE: re-reading a workbook on every single Telegram message would be slow
# and would hammer the disk. Instead we keep the last DataFrame in memory
# together with the file's modification time. A lock is used because
# python-telegram-bot can process several updates concurrently, and two threads
# rebuilding the cache at once would be wasteful.
_cache_df: pd.DataFrame | None = None
_cache_mtime: float | None = None
_lock = threading.Lock()


# ---------------------------------------------------------------------------
def _download_from_onedrive(destination: Path) -> bool:
    """
    PURPOSE: optional fallback for when the bot runs somewhere without the
    OneDrive desktop client (a VPS, Raspberry Pi, cloud container).

    A OneDrive "anyone with the link" URL can be turned into a direct-download
    URL by appending "&download=1". We fetch the bytes and write them to the
    same local path the rest of the code already expects, so nothing else in
    the bot has to know where the file came from.

    Returns True on success, False on any failure (the caller then falls back
    to whatever local copy already exists).
    """
    if not config.ONEDRIVE_SHARE_LINK:
        return False

    try:
        import requests

        url = config.ONEDRIVE_SHARE_LINK
        if "download=1" not in url:
            url += ("&" if "?" in url else "?") + "download=1"

        response = requests.get(url, timeout=30)
        response.raise_for_status()

        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(response.content)
        logger.info("Workbook downloaded from OneDrive share link.")
        return True
    except Exception as exc:
        logger.warning("OneDrive download failed (%s). Using local copy.", exc)
        return False


# ---------------------------------------------------------------------------
def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """
    PURPOSE: turn a messy human-maintained spreadsheet into predictable data.

    Real spreadsheets contain trailing spaces, blank filler rows, and status
    values typed as "in progress", "In Progress " and "ONGOING". Fixing that
    here means every downstream function can do a plain equality check.
    """
    # Strip whitespace from the header row -- "Status " is a very common typo
    # and would otherwise make the column impossible to find.
    df.columns = [str(c).strip() for c in df.columns]

    # Warn loudly if the sheet is missing something the bot needs, rather than
    # crashing later with an obscure KeyError.
    missing = [c for c in config.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        logger.warning("Workbook is missing expected column(s): %s", missing)
        for col in missing:
            df[col] = ""

    # Drop rows where the Category cell is blank -- these are usually spacing
    # rows or notes the FM Officer left at the bottom of the sheet.
    df = df[df[config.COL_CATEGORY].notna()].copy()

    # Trim whitespace on every text cell.
    # NOTE: do NOT test `df[col].dtype == object` here. In pandas 2.x text
    # columns are object dtype, but in pandas 3.0 they load as the new "str"
    # dtype, so that test silently becomes False and no stripping happens.
    # Checking isinstance(v, str) per value works identically on both versions.
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_datetime64_any_dtype(df[col]):
            continue
        df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)

    # Normalise Status to one of the three canonical values.
    df[config.COL_STATUS] = df[config.COL_STATUS].apply(
        lambda v: config.STATUS_CANONICAL.get(str(v).strip().lower(), str(v).strip())
        if pd.notna(v) else "Yet to Start"
    )

    # Normalise Maintenance Type to Preventive / Corrective.
    df[config.COL_TYPE] = df[config.COL_TYPE].apply(
        lambda v: config.TYPE_CANONICAL.get(str(v).strip().lower(), str(v).strip())
        if pd.notna(v) else "-"
    )

    return df


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# CSV cache (separate from the Excel mtime cache)
# ---------------------------------------------------------------------------
# PURPOSE: a published CSV URL has no file modification time we can stat, so we
# cannot use the mtime trick. Instead we time-box the cache: re-fetch only when
# the cached copy is older than REFRESH_INTERVAL_SECONDS, or when force=True.
_csv_cache_df: pd.DataFrame | None = None
_csv_cache_time: float = 0.0


def _rename_form_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    PURPOSE: apply config.FORM_COLUMN_MAP so friendly Google Form question
    titles become the column names the rest of the bot expects.

    Harmless when the map is empty (the recommended setup), in which case the
    form questions are already named exactly like the columns.
    """
    if config.FORM_COLUMN_MAP:
        df = df.rename(columns=config.FORM_COLUMN_MAP)
    return df


def _dedupe_newest(df: pd.DataFrame) -> pd.DataFrame:
    """
    PURPOSE: collapse an append-only Form response sheet down to one row per
    Job ID -- the newest submission wins.

    Google Forms never edits a row, so updating a job means submitting again
    with the same Job ID. Without this step the bot would show both the old
    and new state of every job that was ever updated.

    Newest is decided by the Form's Timestamp column. If that column is missing
    or unparseable, we fall back to row order (later row = newer), which is how
    Forms appends anyway.
    """
    if not config.DEDUPE_BY_JOB_ID or config.COL_JOB_ID not in df.columns:
        return df

    df = df.copy()

    # Separate rows that actually have a Job ID from any blank ones, so blanks
    # are never merged together by drop_duplicates.
    has_id = df[config.COL_JOB_ID].astype(str).str.strip().ne("")
    with_id = df[has_id]
    without_id = df[~has_id]

    ts_col = config.FORM_TIMESTAMP_COLUMN
    if ts_col in with_id.columns:
        parsed = pd.to_datetime(
            with_id[ts_col], dayfirst=config.FORM_TIMESTAMP_DAYFIRST, errors="coerce"
        )
        with_id = with_id.assign(_ts=parsed)
        # na_position="last" keeps rows with a real timestamp ahead of any that
        # failed to parse, so a valid submission always beats an unparseable one.
        with_id = with_id.sort_values("_ts", ascending=False, na_position="last")
        with_id = with_id.drop_duplicates(subset=[config.COL_JOB_ID], keep="first")
        with_id = with_id.drop(columns="_ts")
    else:
        # No timestamp column -> trust append order, last occurrence is newest.
        with_id = with_id.drop_duplicates(subset=[config.COL_JOB_ID], keep="last")

    return pd.concat([with_id, without_id], ignore_index=True)


def _load_from_csv(force: bool = False) -> pd.DataFrame:
    """
    PURPOSE: Option C. Fetch the published-to-web Google Sheet as CSV over plain
    HTTP, dedupe to newest-per-Job-ID, then run the same cleaning as Excel.

    No Google Cloud, no credentials -- the sheet is a public URL. See
    GOOGLE_FORM_SETUP.md for how to produce that URL.
    """
    global _csv_cache_df, _csv_cache_time

    import io
    import time

    import requests

    if not config.GOOGLE_CSV_URL:
        raise ValueError(
            "DATA_SOURCE is 'google_csv' but GOOGLE_CSV_URL is not set.\n"
            "Add the published CSV link to your .env file. "
            "See GOOGLE_FORM_SETUP.md."
        )

    now = time.monotonic()

    with _lock:
        # Time-boxed cache: reuse if fetched recently and not forced.
        if (
            not force
            and _csv_cache_df is not None
            and (now - _csv_cache_time) < config.REFRESH_INTERVAL_SECONDS
        ):
            return _csv_cache_df

        logger.info("Fetching CSV from published Google Sheet...")
        try:
            resp = requests.get(config.GOOGLE_CSV_URL, timeout=20)
            resp.raise_for_status()
        except requests.RequestException as exc:
            # If we have a stale cache, prefer serving it over crashing.
            if _csv_cache_df is not None:
                logger.warning("CSV fetch failed (%s); serving cached copy.", exc)
                return _csv_cache_df
            raise ConnectionError(
                f"Could not fetch the Google CSV: {exc}\n"
                "Check GOOGLE_CSV_URL and that the sheet is still published."
            ) from exc

        # Google serves the CSV as UTF-8. Guard against an HTML error page being
        # returned instead of CSV (happens if the publish link was revoked).
        text = resp.text
        if text.lstrip().lower().startswith("<!doctype html") or "<html" in text[:200].lower():
            raise ValueError(
                "The URL returned a web page, not CSV data.\n"
                "Re-check that the sheet is Published to web as CSV, and that "
                "GOOGLE_CSV_URL ends with 'output=csv'."
            )

        df = pd.read_csv(io.StringIO(text))
        df = _rename_form_columns(df)
        df = _dedupe_newest(df)
        df = _clean(df)

        _csv_cache_df = df
        _csv_cache_time = now
        logger.info("CSV loaded: %d job row(s) after dedupe.", len(df))
        return df


def _load_from_excel(force: bool = False) -> pd.DataFrame:
    """PURPOSE: the original Excel path, unchanged, now behind the source switch."""
    global _cache_df, _cache_mtime

    path = config.EXCEL_PATH

    if config.ONEDRIVE_SHARE_LINK:
        _download_from_onedrive(path)

    if not path.exists():
        raise FileNotFoundError(
            f"Maintenance workbook not found at: {path}\n"
            "Check EXCEL_PATH in config.py or your .env file."
        )

    mtime = path.stat().st_mtime

    with _lock:
        # Fast path -- nothing changed since last time, reuse the cache.
        if not force and _cache_df is not None and _cache_mtime == mtime:
            return _cache_df

        logger.info("Reading workbook from disk: %s", path)
        df = pd.read_excel(path, sheet_name=config.SHEET_NAME, engine="openpyxl")
        df = _clean(df)

        _cache_df = df
        _cache_mtime = mtime
        return df


def load_jobs(force: bool = False) -> pd.DataFrame:
    """
    PURPOSE: the main entry point used by every command handler.

    Routes to the configured data source:
      * "excel"      -> read the local .xlsx workbook (default).
      * "google_csv" -> fetch a published Google Form response sheet as CSV.

    Every other function (get_jobs_for_category, get_summary_counts) calls this,
    so switching DATA_SOURCE changes the whole bot with no other code edits.
    """
    if config.DATA_SOURCE == "google_csv":
        return _load_from_csv(force)
    return _load_from_excel(force)


# ---------------------------------------------------------------------------
def get_jobs_for_category(category_key: str) -> pd.DataFrame:
    """
    PURPOSE: the "Search DataFrame" stage of your workflow.

    Given a command key such as "acmv", look up the matching Excel category
    string from config.CATEGORIES and return only those rows. Matching is
    case-insensitive so "acmv", "ACMV" and "Acmv" in the sheet all match.
    """
    df = load_jobs()
    excel_name = config.CATEGORIES[category_key]["excel"]

    mask = df[config.COL_CATEGORY].astype(str).str.strip().str.lower() == excel_name.lower()
    return df[mask]


# ---------------------------------------------------------------------------
def get_summary_counts() -> dict:
    """
    PURPOSE: feed the /service dashboard.

    Returns a dictionary shaped like:
        { "acmv": {"total": 4, "Yet to Start": 1, "In Progress": 2, "Completed": 1}, ... }

    Doing the counting here keeps the message-building code in formatters.py
    purely about presentation.
    """
    df = load_jobs()
    result = {}

    for key, meta in config.CATEGORIES.items():
        rows = df[
            df[config.COL_CATEGORY].astype(str).str.strip().str.lower()
            == meta["excel"].lower()
        ]
        counts = rows[config.COL_STATUS].value_counts().to_dict()
        result[key] = {
            "total": len(rows),
            "Yet to Start": counts.get("Yet to Start", 0),
            "In Progress": counts.get("In Progress", 0),
            "Completed": counts.get("Completed", 0),
        }

    return result