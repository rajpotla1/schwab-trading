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

# 4. Fetch all active/working orders from Schwab to find IDs automatically
print("Fetching active working orders from Schwab...")

# Define time window: from 30 days ago until right now (UTC timezone standard)
now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")



try:
    # Pass the required time windows explicitly alongside your status filter
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
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "ExistTdySellOrder.xlsx")
try:
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} lines from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# 6. Loop through rows, auto-locate matching active order, and replace
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper()  # "BUY" or "SELL"
    quantity = int(row['quantity'])
    order_type = str(row['orderType']).strip().upper()
    duration = str(row['duration']).strip().upper()
    session_type = str(row.get('session', 'NORMAL')).strip().upper()

    if session_type == "SEEMLESS":
        session_type = "SEAMLESS"

    print(f"Processing row {index + 1}: Searching for open {instruction} order on {symbol}...")

    # --- AUTO-LOOKUP LOGIC ---
    matched_order_id = None
    for order in open_orders:
        # Navigate Schwab's standard JSON order leg structure
        legs = order.get('orderLegCollection', [])
        if legs:
            leg = legs[0]
            leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
            leg_instruction = leg.get('instruction', '').upper()
            
            # Check if this active order matches our Excel symbol and buy/sell intent
            if leg_symbol == symbol and leg_instruction == instruction:
                matched_order_id = str(order.get('orderId'))
                break # Found it!

    if not matched_order_id:
        print(f"❌ Skipped row {index + 1}: No active working {instruction} order found for {symbol} to replace.")
        print("-" * 40)
        continue

    print(f"🔍 Found matching Order ID {matched_order_id} on Schwab server. Formatting replacement...")

    # Safety Check for Market Orders in Extended Hours
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping row {index + 1}: Extended session '{session_type}' does not allow MARKET orders.")
        print("-" * 40)
        continue

    # Build the modification payload format
    order_payload = {
        "orderType": order_type,
        "session": session_type,        
        "duration": duration,
        "orderStrategyType": "SINGLE",
        "orderLegCollection": [
            {
                "instruction": instruction,
                "quantity": quantity,
                "instrument": {
                    "symbol": symbol,
                    "assetType": "EQUITY"
                }
            }
        ]
    }

    if order_type == "LIMIT":
        order_payload["price"] = f"{float(row['price']):.2f}"

    # Route the dynamic Replace Request
    try:
        resp = client.replace_order(account_hash, matched_order_id, order_payload)
        
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            new_order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! Order {matched_order_id} replaced. New Order ID: {new_order_id}")
        else:
            print(f"❌ Failed to replace order {matched_order_id}. Status Code: {resp.status_code}")
            print(f"Response: {resp.text}")
            
    except Exception as e:
        print(f"💥 Critical error executing lookup-replacement: {e}")
        
    print("-" * 40)

print("Batch dynamic modifications complete.")
