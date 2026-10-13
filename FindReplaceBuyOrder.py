import os
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
import schwabdev

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 2. Initialize the Schwab client
client = schwabdev.Client(client_id, client_secret, callback_url)

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

# 4. Request the Stock Ticker and the NEW Price from the user
print("\n" + "="*40)
target_symbol = input("Enter ticker symbol to replace (e.g. AAPL): ").strip().upper()
try:
    new_price_input = input("Enter the NEW Limit Price: ").strip()
    new_price = f"{float(new_price_input):.2f}"
except ValueError:
    print("❌ Error: Invalid price entered.")
    raise SystemExit()
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

# 6. Find and isolate exactly ONE single buy order matching the symbol
target_order = None
for order in open_orders:
    legs = order.get('orderLegCollection', [])
    if legs:
        leg = legs[0]
        leg_instruction = leg.get('instruction', '').upper()
        leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
        
        # Isolate the first working BUY limit order matching your exact ticker
        if leg_symbol == target_symbol and leg_instruction == "BUY":
            target_order = order
            break 

# 7. Execute Cancel & Replace request
if target_order:
    target_order_id = str(target_order.get('orderId'))
    print(f"🎯 Found live BUY order for {target_symbol} (ID: {target_order_id}). Rebuilding payload...")
    
    # Rebuild the Schwab replacement payload based on the original order details
    original_leg = target_order.get('orderLegCollection', [])[0]
    
    replacement_payload = {
        "orderType": target_order.get("orderType", "LIMIT").upper(),
        "session": target_order.get("session", "NORMAL").upper(),
        "duration": target_order.get("duration", "DAY").upper(),
        "orderStrategyType": "SINGLE",
        "price": new_price,  # Injecting your new user-defined price
        "orderLegCollection": [
            {
                "instruction": "BUY",
                "quantity": int(original_leg.get("quantity")),
                "instrument": {
                    "symbol": target_symbol,
                    "assetType": "EQUITY"
                }
            }
        ]
    }
    
    print(f"🔄 Sending Cancel/Replace request for Order {target_order_id} with new price ${new_price}...")
    try:
        # Schwab uses replace_order(accountHash, orderId, payload) for modifying open orders
        resp = client.replace_order(account_hash, target_order_id, replacement_payload)
        
        if resp.status_code in (200, 201, 202):
            location_header = resp.headers.get('location', '')
            new_order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! Original order replaced. New Order ID: {new_order_id}")
        else:
            print(f"❌ Failed to replace order. Status Code: {resp.status_code} | Response: {resp.text}")
            
    except Exception as e:
        print(f"💥 Critical error executing replacement: {e}")
else:
    print("no present")
