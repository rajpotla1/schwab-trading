"""
Shared location of the schwabdev tokens database.

schwabdev defaults to ~/.schwabdev/tokens.db, which differs per Windows user
(C:\\Users\\potlasr vs C:\\Users\\rajpo), so each computer needed its own login.
Keeping tokens.db in this project's Config folder lets OneDrive sync one token
to every computer. Set SCHWAB_TOKENS_DB in .env to override the location.
"""
import os

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

TOKENS_DB = os.path.expanduser(os.getenv("SCHWAB_TOKENS_DB") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "Config", "tokens.db"))
