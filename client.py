
import os
import schwabdev
from dotenv import load_dotenv

# 1. Load the hidden variables from the .env file into the system memory
load_dotenv()

# 2. Retrieve the variables securely using os.getenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")
account_id = os.getenv("SCHWAB_ACCOUNT_ID")
# 2. Initialize the client (This triggers the initial login flow)

client = schwabdev.Client(client_id, client_secret, callback_url)

client = schwabdev.Client(
    client_id,
    client_secret,
    callback_url,
    tokens_db="~/.schwabdev/tokens.db",
    encryption=None,
    timeout=10,
    call_on_auth=None,
    open_browser_for_auth=True,
    validate_params=True
)

# for browser call so you can update accounts to be linked

#client.update_tokens(force_refresh_token=True)
