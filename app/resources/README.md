# PTR resources

The [project README](../../README.md) is the canonical guide to PTR's architecture, workflows, setup, risks, and development priorities.

- [Pinned Python dependencies](requirements.txt)
- [Paper bot notes](bot.py.md)
- [Main scraper source](../code/ptr.py)
- [Alternative scraper source](../code/ptr_extended.py)

Historical setup instructions were replaced because they referred to old machine-specific paths and did not match the current code. Credentials belong in the private `../.env` file under the names `ALPACA_API_KEY` and `ALPACA_SECRET_KEY`, never in documentation.
