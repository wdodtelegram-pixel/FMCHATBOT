FM Chatbot
===========

A Telegram bot that reads a maintenance workbook and exposes maintenance jobs through chat commands.

Requirements
------------

- Python 3.10 or newer
- The bot works with the built-in `openpyxl` workbook reader when `pandas` is not installed

Installation
------------

1. Install the required Python packages using your Python interpreter:

   py -3 -m pip install --user -r requirements.txt

2. Create a `.env` file next to `bot.py` with your Telegram bot token:

   TELEGRAM_BOT_TOKEN=your-token-from-BotFather

3. Set `EXCEL_PATH` in `.env` if your workbook is not located at the default `data/maintenance_jobs.xlsx`.
   If you do not want to override the default, leave `EXCEL_PATH` unset rather than setting it to a blank value.

Run
---

py -3 bot.py
