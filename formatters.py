

from __future__ import annotations

import html
import math
from typing import Any

import config
from date_utils import to_week_format


def _is_missing(value: Any) -> bool:
    return value is None or (isinstance(value, float) and math.isnan(value))


def _esc(value: Any) -> str:
    """
    PURPOSE: escape user/spreadsheet text before it goes into an HTML message.

    If someone types "AHU <3" or "Level 5 & 6" into Excel, an unescaped "<" or
    "&" makes Telegram reject the whole message with a parse error. This tiny
    helper prevents that entire class of bug.
    """
    if _is_missing(value):
        return "-"
    text = str(value).strip()
    return html.escape(text) if text else "-"


def _status_icon(status: str) -> str:
    """PURPOSE: look up the coloured dot for a status, defaulting to a blank."""
    return config.STATUS_ICONS.get(str(status).strip().lower(), "⚫")


# ---------------------------------------------------------------------------
def welcome_message() -> str:
    """
    PURPOSE: the reply to /start -- the menu shown in your first mock-up.

    The command list is generated from config.CATEGORIES rather than typed out,
    so adding a 6th category automatically adds a 6th line here with no edit.
    """
    lines = [
        "👷 <b>Welcome to the Facilities Management Chatbot</b>",
        "",
        "Real-time enquiry system for routine maintenance jobs.",
        "",
        "<b>Here are the commands:</b>",
        "/service — shows Routine Maintenance Dashboard",
        "",
        "Further look into each job category by typing e.g. /acmv to look into ACMV jobs:",
    ]

    for key, meta in config.CATEGORIES.items():
        lines.append(f"/{key} — {meta['emoji']} {meta['blurb']}")

    lines += [
        "",
        "<b>Other commands</b>",
        "/pending — all jobs not yet completed",
        "/refresh — pull the latest data from the spreadsheet",
        "/help — show this menu again",
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
def dashboard_message(summary: dict) -> str:
    """
    PURPOSE: the reply to /service -- a one-screen overview of all categories.

    Shows a per-category count of the three statuses so an FM manager can see
    the whole building's position at a glance, then drill in with /acmv etc.
    """
    lines = [
        "📋 <b>Routine Maintenance Dashboard</b>",
        "",
    ]

    grand_total = 0
    grand_done = 0

    for key, meta in config.CATEGORIES.items():
        stats = summary[key]
        grand_total += stats["total"]
        grand_done += stats["Completed"]

        lines.append(f"{meta['emoji']} <b>{meta['name']}</b>  (/{key})")

        if stats["total"] == 0:
            lines.append("   No jobs recorded")
        else:
            lines.append(
                f"   ⚪ {stats['Yet to Start']}   "
                f"🟡 {stats['In Progress']}   "
                f"🟢 {stats['Completed']}   "
                f"— {stats['total']} job(s)"
            )
        lines.append("")

    lines.append("──────────────")
    lines.append(f"<b>Total:</b> {grand_total} job(s) · {grand_done} completed")
    lines.append("")
    lines.append("<i>⚪ Yet to start  🟡 In progress  🟢 Completed</i>")
    lines.append("Tap a category command above for full details.")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
def _job_block(row) -> str:
    """
    PURPOSE: render ONE job row into the compact five-field block you specified.

    Keeping this as its own function means the category view, the /pending view
    and any future view all display a job identically.
    """
    icon = _status_icon(row.get(config.COL_STATUS, ""))

    return "\n".join([
        f"{icon} <b>{_esc(row.get(config.COL_DESCRIPTION))}</b>",
        f"   📍 Location: {_esc(row.get(config.COL_LOCATION))}",
        f"   🔧 Type: {_esc(row.get(config.COL_TYPE))} maintenance",
        f"   📊 Status: {_esc(row.get(config.COL_STATUS))}",
        f"   🗓 Est. completion: {to_week_format(row.get(config.COL_EST_COMPLETION))}",
        f"   🔁 Next job: {to_week_format(row.get(config.COL_NEXT_JOB))}",
    ])


# ---------------------------------------------------------------------------
def category_message(category_key: str, rows) -> str:
    """
    PURPOSE: the reply to /acmv, /lift, /electrical, /fire, /plumbing --
    the second mock-up in your document.

    Handles the empty case explicitly so the user gets a clear sentence rather
    than a bare heading with nothing under it.
    """
    meta = config.CATEGORIES[category_key]
    header = f"{meta['emoji']} <b>{meta['name'].upper()} — Routine Maintenance Dashboard</b>"

    if rows.empty:
        return (
            f"{header}\n\n"
            "No jobs are currently recorded for this category.\n"
            "Use /service to see all categories."
        )

    # Sort so unfinished work appears first -- that is what people open the bot
    # to check. Completed jobs sink to the bottom.
    order = {"In Progress": 0, "Yet to Start": 1, "Completed": 2}
    rows = rows.copy()
    rows["_sort"] = rows[config.COL_STATUS].map(lambda s: order.get(s, 3))
    rows = rows.sort_values("_sort")

    blocks = [_job_block(row) for _, row in rows.iterrows()]

    footer = (
        "──────────────\n"
        f"<i>{len(rows)} job(s) in this category · /service for the full dashboard</i>"
    )

    return f"{header}\n\n" + "\n\n".join(blocks) + f"\n\n{footer}"


# ---------------------------------------------------------------------------
def pending_message(df) -> str:
    """
    PURPOSE: the reply to /pending -- every job across all categories that is
    not yet Completed, grouped by category.

    This is the view a duty officer wants at the start of a shift.
    """
    outstanding = df[df[config.COL_STATUS] != "Completed"]

    if outstanding.empty:
        return "✅ <b>No outstanding jobs.</b>\nEverything in the workbook is marked completed."

    lines = ["⏳ <b>Outstanding Jobs (all categories)</b>", ""]

    for key, meta in config.CATEGORIES.items():
        rows = outstanding[
            outstanding[config.COL_CATEGORY].astype(str).str.strip().str.lower()
            == meta["excel"].lower()
        ]
        if rows.empty:
            continue

        lines.append(f"{meta['emoji']} <b>{meta['name']}</b>")
        for _, row in rows.iterrows():
            lines.append(
                f"   {_status_icon(row[config.COL_STATUS])} "
                f"{_esc(row.get(config.COL_DESCRIPTION))} — "
                f"{_esc(row.get(config.COL_LOCATION))} — "
                f"due {to_week_format(row.get(config.COL_EST_COMPLETION))}"
            )
        lines.append("")

    lines.append(f"<i>{len(outstanding)} outstanding job(s) in total.</i>")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
def error_message(exc: Exception) -> str:
    """
    PURPOSE: turn an internal exception into something a facilities officer can
    actually act on, instead of showing a Python traceback in a chat window.
    """
    if isinstance(exc, FileNotFoundError):
        return (
            "⚠️ <b>Cannot reach the maintenance spreadsheet.</b>\n\n"
            "The workbook was not found at the configured path. "
            "Check that OneDrive is running and synced, then try /refresh."
        )
    return (
        "⚠️ <b>Something went wrong reading the maintenance data.</b>\n\n"
        f"<code>{html.escape(str(exc))[:300]}</code>\n\n"
        "The spreadsheet may be mid-save. Please try again in a moment."
    )
