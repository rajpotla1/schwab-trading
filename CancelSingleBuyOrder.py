import os
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
import schwabdev
from schwab_tokens import TOKENS_DB

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

# 4. Request the specific Stock Ticker from the user
print("\n" + "="*40)
target_symbol = input("Enter ticker symbol to cancel (e.g. AAPL): ").strip().upper()
print("="*40)

if not target_symbol:
    print("❌ Error: No symbol entered.")
    raise SystemExit()

# 5. Fetch live working orders
print(f"Checking for live working orders on {target_symbol}...")
now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

try:
    orders_resp = client.account_orders(
        accountHash=account_hash,
        status="WORKING",
        fromEnteredTime=from_time,
        toEnteredTime=to_time
    )
    open_orders = orders_resp.json() if orders_resp.status_code == 200 else []
except Exception as e:
    print(f"💥 Failed to query order endpoints: {e}")
    open_orders = []

# 6. Find and isolate ALL buy orders matching the symbol
matching_order_ids = []
for order in open_orders:
    legs = order.get('orderLegCollection', [])
    if legs:
        leg = legs[0]
        leg_instruction = leg.get('instruction', '').upper()
        leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
        
        # Collect all open BUY orders for the specified target symbol
        if leg_symbol == target_symbol and leg_instruction == "BUY":
            order_id = order.get('orderId')
            if order_id:
                matching_order_ids.append(str(order_id))

# 7. Execute cancellation requests for all found orders
if matching_order_ids:
    print(f"🎯 Found {len(matching_order_ids)} live BUY order(s) for {target_symbol}. Requesting cancellations...")
    
    success_count = 0
    for target_order_id in matching_order_ids:
        try:
            resp = client.cancel_order(account_hash, target_order_id)
            if resp.status_code in (200, 201, 202):
                print(f"✅ Success! Cancellation request sent for Order {target_order_id}.")
                success_count += 1
            else:
                print(f"❌ Failed for Order {target_order_id}. Status Code: {resp.status_code} | Response: {resp.text}")
        except Exception as e:
            print(f"💥 Critical error executing cancellation for Order {target_order_id}: {e}")
            
    print(f"\n📊 Summary: Successfully requested cancellation for {success_count}/{len(matching_order_ids)} orders.")
else:
    # Explicit message printed when no active buy order matches the symbol
    print("no present")
