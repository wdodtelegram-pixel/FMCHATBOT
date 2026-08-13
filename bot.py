
from __future__ import annotations

import logging
from functools import partial

from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.error import InvalidToken
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

import config
import excel_loader
import formatters

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
# PURPOSE: so that when the bot misbehaves at 2am you can see which command was
# called and what the spreadsheet did. httpx is silenced because it logs every
# single Telegram poll, which floods the log file.
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Shared UI
# ---------------------------------------------------------------------------
def category_keyboard() -> InlineKeyboardMarkup:
    """
    PURPOSE: build the tappable button grid under the dashboard.

    Typing /acmv works, but on a phone tapping a button is far quicker. The
    buttons carry a callback_data string like "cat:acmv", which the callback
    handler below decodes. Buttons are laid out two per row so the labels stay
    readable on a narrow screen.
    """
    buttons = [
        InlineKeyboardButton(f"{meta['emoji']} {meta['name']}", callback_data=f"cat:{key}")
        for key, meta in config.CATEGORIES.items()
    ]

    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    rows.append([InlineKeyboardButton("🔄 Refresh data", callback_data="refresh")])

    return InlineKeyboardMarkup(rows)


MAX_MESSAGE_LENGTH = 4096  # Telegram's hard limit


def _split_message(text: str, limit: int = MAX_MESSAGE_LENGTH - 100) -> list[str]:
    """
    PURPOSE: Telegram rejects any message longer than 4096 characters.

    A category with 40+ jobs would silently fail to send. This splits the text
    on blank lines (the gap between job blocks) so a single job is never cut in
    half, which would also break an unclosed <b> tag and cause a parse error.
    """
    if len(text) <= limit:
        return [text]

    chunks, current = [], ""
    for block in text.split("\n\n"):
        # +2 accounts for the "\n\n" we will re-insert.
        if len(current) + len(block) + 2 > limit and current:
            chunks.append(current)
            current = block
        else:
            current = f"{current}\n\n{block}" if current else block

    if current:
        chunks.append(current)
    return chunks


async def _reply(update: Update, text: str, keyboard: InlineKeyboardMarkup | None = None):
    """
    PURPOSE: one helper that can answer BOTH a typed command and a tapped
    button, so every handler below stays two lines long.

    A typed command arrives with update.message set; a button press arrives
    with update.callback_query set instead. Without this helper each handler
    would need that if/else duplicated.

    The keyboard is attached only to the LAST chunk, so the buttons always sit
    at the bottom of the conversation where the user expects them.
    """
    chunks = _split_message(text)

    for index, chunk in enumerate(chunks):
        is_last = index == len(chunks) - 1
        markup = keyboard if is_last else None

        # A button press edits the message in place, but only for the first
        # chunk -- any overflow has to be sent as a new message.
        if update.callback_query and index == 0:
            await update.callback_query.edit_message_text(
                chunk, parse_mode=ParseMode.HTML, reply_markup=markup
            )
        else:
            target = update.effective_chat
            await target.send_message(
                chunk, parse_mode=ParseMode.HTML, reply_markup=markup
            )


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: handles /start and /help -- the welcome menu from your mock-up.

    Also attaches the category keyboard so a new user can start browsing
    immediately without memorising any commands.
    """
    await _reply(update, formatters.welcome_message(), category_keyboard())


async def service_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: handles /service -- the Routine Maintenance Dashboard.

    Wrapped in try/except because this is the first command that touches the
    spreadsheet: if OneDrive is offline or the file is locked mid-save, the
    user gets a readable explanation instead of silence.
    """
    try:
        summary = excel_loader.get_summary_counts()
        await _reply(update, formatters.dashboard_message(summary), category_keyboard())
    except Exception as exc:
        logger.exception("/service failed")
        await _reply(update, formatters.error_message(exc))


async def category_command(update: Update, context: ContextTypes.DEFAULT_TYPE, key: str):
    """
    PURPOSE: handles /acmv, /lift, /electrical, /fire and /plumbing.

    ONE function serves all of them. The `key` argument is bound per-command
    using functools.partial in main(), which is why adding a 6th category to
    config.py needs no new code here.
    """
    try:
        rows = excel_loader.get_jobs_for_category(key)
        await _reply(update, formatters.category_message(key, rows), category_keyboard())
    except Exception as exc:
        logger.exception("/%s failed", key)
        await _reply(update, formatters.error_message(exc))


async def pending_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """PURPOSE: handles /pending -- every job across all categories not yet completed."""
    try:
        df = excel_loader.load_jobs()
        await _reply(update, formatters.pending_message(df), category_keyboard())
    except Exception as exc:
        logger.exception("/pending failed")
        await _reply(update, formatters.error_message(exc))


async def refresh_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: handles /refresh -- forces an immediate re-read of the workbook,
    bypassing the cache.

    The bot already detects file changes automatically, so this exists mainly
    to reassure a user who has just edited the sheet and wants confirmation
    that the bot has picked the change up.
    """
    try:
        df = excel_loader.load_jobs(force=True)
        await _reply(
            update,
            f"🔄 <b>Data refreshed.</b>\n{len(df)} job record(s) loaded from the spreadsheet.",
            category_keyboard(),
        )
    except Exception as exc:
        logger.exception("/refresh failed")
        await _reply(update, formatters.error_message(exc))


async def unknown_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: catches any /command that does not exist, plus plain text.

    Without this, a typo like /acvm produces total silence and the user assumes
    the bot is broken. Pointing them back at /start is a better dead end.
    """
    await update.message.reply_text(
        "🤔 I didn't recognise that. Send /start to see the available commands.",
        parse_mode=ParseMode.HTML,
    )


# ---------------------------------------------------------------------------
# Button (callback query) handler
# ---------------------------------------------------------------------------
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: decodes taps on the inline keyboard.

    answer() must be called first -- it stops the little loading spinner on the
    user's phone. Telegram shows the button as stuck-loading for ~30s if you
    forget, which looks like a hang.
    """
    query = update.callback_query
    await query.answer()

    data = query.data

    if data == "refresh":
        await refresh_command(update, context)
    elif data.startswith("cat:"):
        await category_command(update, context, key=data.split(":", 1)[1])


# ---------------------------------------------------------------------------
# Background refresh job
# ---------------------------------------------------------------------------
async def scheduled_refresh(context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: keeps the in-memory cache warm.

    Every REFRESH_INTERVAL_SECONDS this quietly re-reads the workbook so the
    first user of the morning does not pay the file-read delay. It also surfaces
    a broken OneDrive sync in the logs before a user ever notices.
    """
    try:
        excel_loader.load_jobs(force=True)
        logger.info("Scheduled workbook refresh completed.")
    except Exception as exc:
        logger.warning("Scheduled refresh failed: %s", exc)


# ---------------------------------------------------------------------------
# Error handler
# ---------------------------------------------------------------------------
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """
    PURPOSE: last line of defence. Any exception not caught inside a handler
    lands here and is logged instead of crashing the polling loop, so one bad
    spreadsheet row can never take the whole bot down.
    """
    logger.error("Unhandled exception", exc_info=context.error)


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------
async def post_init(application: Application):
    """
    PURPOSE: registers the command list with Telegram itself.

    This is what populates the blue "Menu" button next to the message box, so
    users see /service, /acmv and the rest as a native dropdown. Generated from
    config.CATEGORIES for the same reason as everything else.
    """
    commands = [
        BotCommand("start", "Show the welcome menu"),
        BotCommand("service", "Routine maintenance dashboard"),
    ]
    commands += [
        BotCommand(key, meta["blurb"]) for key, meta in config.CATEGORIES.items()
    ]
    commands += [
        BotCommand("pending", "All outstanding jobs"),
        BotCommand("refresh", "Reload the latest spreadsheet data"),
        BotCommand("help", "Show the command list"),
    ]

    await application.bot.set_my_commands(commands)
    logger.info("Bot commands registered with Telegram.")


def main():
    """
    PURPOSE: wires everything together and starts the bot.

    Order matters: the catch-all unknown_command handler is registered LAST,
    because python-telegram-bot uses the first matching handler in each group.
    Registered earlier, it would swallow every real command.
    """
    # -----------------------------------------------------------------------
    # Pre-flight token check
    # -----------------------------------------------------------------------
    # PURPOSE: validate the token BEFORE building the Application, so a setup
    # mistake produces four readable lines instead of a nested async traceback
    # from deep inside the Telegram library.
    is_valid, message = config.validate_token()
    if not is_valid:
        print("\n" + "=" * 68)
        print("  CANNOT START — TOKEN PROBLEM")
        print("=" * 68)
        print(message)
        print("=" * 68 + "\n")
        raise SystemExit(1)

    application = ApplicationBuilder().token(config.BOT_TOKEN).post_init(post_init).build()

    # --- core commands -----------------------------------------------------
    application.add_handler(CommandHandler(["start", "help"], start_command))
    application.add_handler(CommandHandler("service", service_command))
    application.add_handler(CommandHandler("pending", pending_command))
    application.add_handler(CommandHandler("refresh", refresh_command))

    # --- one handler per category, generated from config -------------------
    # partial() freezes the `key` argument so /acmv always calls
    # category_command(..., key="acmv"). A plain loop variable would not work
    # here: every handler would end up sharing the last value of `key`.
    for key in config.CATEGORIES:
        application.add_handler(
            CommandHandler(key, partial(category_command, key=key))
        )

    # --- inline buttons ----------------------------------------------------
    application.add_handler(CallbackQueryHandler(button_handler))

    # --- catch-all (must be last) ------------------------------------------
    application.add_handler(MessageHandler(filters.COMMAND, unknown_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, unknown_command))

    # --- error handler and background job ----------------------------------
    application.add_error_handler(error_handler)

    if application.job_queue:
        application.job_queue.run_repeating(
            scheduled_refresh,
            interval=config.REFRESH_INTERVAL_SECONDS,
            first=config.REFRESH_INTERVAL_SECONDS,
            name="workbook_refresh",
        )

    logger.info("FM bot starting (polling)...")

    # run_polling blocks forever, asking Telegram for new messages. It handles
    # its own event loop and shuts down cleanly on Ctrl-C.
    #
    # The try/except catches the case the format check CANNOT catch: a token
    # that is shaped correctly but has been revoked, belongs to a deleted bot,
    # or was mistyped by one character. Only Telegram's server knows that.
    try:
        application.run_polling(drop_pending_updates=True)
    except InvalidToken:
        print("\n" + "=" * 68)
        print("  TELEGRAM REJECTED YOUR TOKEN")
        print("=" * 68)
        print(
            "The token is correctly formatted, but Telegram does not recognise it.\n\n"
            "Likely causes:\n"
            "  - the token was revoked (/revoke in @BotFather issues a new one)\n"
            "  - the bot was deleted with /deletebot\n"
            "  - one character is wrong — re-copy the whole string\n\n"
            "Send /mybots to @BotFather, pick your bot, then 'API Token' to see\n"
            "the current valid token."
        )
        print("=" * 68 + "\n")
        raise SystemExit(1)
    except KeyboardInterrupt:
        # PURPOSE: Ctrl+C is a normal way to stop the bot, not a crash.
        # Without this, stopping produces an ugly traceback every time.
        logger.info("Bot stopped by user.")


if __name__ == "__main__":
    main()
