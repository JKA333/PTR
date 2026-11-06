# PTR paper bot

The implementation is [`code/bot.py`](../code/bot.py). Read the [project README](../../README.md) for the architecture, setup, and current limitations.

The bot loads deduplicated transactions from `data/PTR.sqlite3` and falls back to `data/PTR_Transactions.xlsx`. It uses Alpaca paper credentials from `../.env`. Its default dry run checks the account and records `dry_run` results in SQLite. `--execute` submits paper orders.

**Execution-state issue:** the current `already_executed` check treats a dry run, skip, or error as completed. A previous run can therefore prevent a later paper order for the same trade UID. Review or fix this behavior before relying on retries or paper execution. The separate `risk_engine.py` prototype is not connected to the bot.
