import os
import pandas as pd
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

# 5. Load and iterate through your Excel file
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "CancelBuyOrder.xlsx")
try:
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} lines from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# 6. Loop through rows, auto-locate ALL matching active BUY orders, and CANCEL them
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper() # "BUY" or "SELL"
    
    # Strictly isolate BUY orders as requested
    if instruction != "BUY":
        print(f"Skipping row {index + 1}: Instruction is '{instruction}', not 'BUY'.")
        print("-" * 40)
        continue
        
    print(f"Processing row {index + 1}: Searching for open BUY orders on {symbol}...")
    
    # --- MULTI-LOOKUP LOGIC ---
    # Collect all matching order IDs for this symbol
    matched_order_ids = []
    for order in open_orders:
        legs = order.get('orderLegCollection', [])
        if legs:
            leg = legs[0]
            leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
            leg_instruction = leg.get('instruction', '').upper()
            
            # Match symbol and explicitly confirm it is an active BUY order
            if leg_symbol == symbol and leg_instruction == "BUY":
                matched_order_ids.append(str(order.get('orderId')))
    
    if not matched_order_ids:
        print(f"❌ Skipped row {index + 1}: No active working BUY orders found for {symbol} to cancel.")
        print("-" * 40)
        continue
        
    print(f"🔍 Found {len(matched_order_ids)} matching open BUY order(s) on Schwab server. Requesting cancellations...")
    
    # Route the Cancel Request for each matched ID
    for order_id in matched_order_ids:
        try:
            resp = client.cancel_order(account_hash, order_id)
            # Schwab usually returns 200 OK or 202 Accepted for a successful cancellation request
            if resp.status_code in (200, 201, 202):
                print(f"   ✅ Success! Cancellation request sent for Order {order_id} ({symbol}).")
            else:
                print(f"   ❌ Failed to cancel order {order_id}. Status Code: {resp.status_code}")
                print(f"      Response: {resp.text}")
        except Exception as e:
            print(f"   💥 Critical error executing cancellation for Order {order_id}: {e}")
            
    print("-" * 40)

print("Batch dynamic cancellations complete.")
