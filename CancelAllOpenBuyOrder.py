import os
import schwabdev
from schwab_tokens import TOKENS_DB
from dotenv import load_dotenv
from datetime import datetime, timedelta, timezone

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 2. Initialize the Schwab client
client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)

# 3. Retrieve linked accounts and extract the account hash
linked_accounts = client.linked_accounts().json()
if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
    account_hash = linked_accounts[0].get('hashValue')
elif isinstance(linked_accounts, dict):
    account_securities = linked_accounts.get('securitiesAccountList', [])
    account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

# 4. Fetch all active/working orders from Schwab
print("Fetching active working orders from Schwab...")
now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

try:
    orders_resp = client.account_orders(
        accountHash=account_hash,
        fromEnteredTime=from_time,
        toEnteredTime=to_time,
        status="WORKING"
    )
    open_orders = []
    if orders_resp.status_code == 200:
        open_orders = orders_resp.json()
        print(f"Successfully retrieved {len(open_orders)} working orders from server.")
    else:
        print(f"⚠️ Warning: Could not fetch open orders from server. Status: {orders_resp.status_code}")
        print(f"Server response message: {orders_resp.text}")
        open_orders = []
except Exception as e:
    print(f"💥 Failed to query order endpoints: {e}")
    open_orders = []

# 5. Filter and isolate ALL open BUY orders globally
matched_order_ids = []

for order in open_orders:
    legs = order.get('orderLegCollection', [])
    if legs:
        # Check the first leg of the order execution
        leg = legs[0]
        leg_instruction = leg.get('instruction', '').upper()
        leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
        order_id = str(order.get('orderId'))
        
        # Match purely on whether it is a BUY instruction
        if leg_instruction == "BUY":
            matched_order_ids.append((order_id, leg_symbol))

# 6. Execute global cancellation loop
if not matched_order_ids:
    print("\n✅ No active working BUY orders found on the server to cancel.")
else:
    print(f"\n🔍 Found {len(matched_order_ids)} total open BUY order(s) across all symbols. Processing cancellations...")
    print("-" * 50)
    
    for order_id, symbol in matched_order_ids:
        try:
            resp = client.cancel_order(account_hash, order_id)
            # Schwab returns 200, 201, or 202 for successful cancel requests
            if resp.status_code in (200, 201, 202):
                print(f" ✅ Success! Cancelled Order {order_id} [Symbol: {symbol}]")
            else:
                print(f" ❌ Failed to cancel order {order_id} ({symbol}). Status: {resp.status_code}")
                print(f"    Response: {resp.text}")
        except Exception as e:
            print(f" 💥 Critical error executing cancellation for Order {order_id}: {e}")
            
    print("-" * 50)
    print("Global buy order cancellations complete.")
