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

# 4. Request the Stock Ticker, NEW Price, and NEW Quantity from the user
print("\n" + "="*40)
target_symbol = input("Enter ticker symbol to look up (e.g. AAPL): ").strip().upper()
if not target_symbol:
    print("❌ Error: No symbol entered.")
    raise SystemExit()

try:
    new_price_input = input("Enter the NEW Sell Limit Price: ").strip()
    new_price = f"{float(new_price_input):.2f}"
    
    new_qty_input = input("Enter the NEW Sell Quantity: ").strip()
    new_quantity = int(new_qty_input)
    if new_quantity <= 0:
        raise ValueError("Quantity must be greater than zero.")
except ValueError as e:
    print(f"❌ Error: Invalid price or quantity entered. ({e})")
    raise SystemExit()
print("="*40)

# 5. Fetch live working orders
print(f"Checking for live working SELL orders on {target_symbol}...")
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

# 6. Find and isolate exactly ONE single SELL order matching the symbol
target_order = None
for order in open_orders:
    legs = order.get('orderLegCollection', [])
    if legs:
        leg = legs[0]  # ✅ Fixed array indexing extraction bug here
        leg_instruction = leg.get('instruction', '').upper()
        leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
        
        # Isolate the first working SELL limit order matching your exact ticker
        if leg_symbol == target_symbol and leg_instruction == "SELL":
            target_order = order
            break 

# 7. Execute Cancel & Place request
if target_order:
    target_order_id = str(target_order.get('orderId'))
    print(f"🎯 Found live SELL order for {target_symbol} (ID: {target_order_id}). Proceeding to cancel...")
    
    try:
        # First, cancel the old order entirely
        cancel_resp = client.cancel_order(account_hash, target_order_id)
        
        if cancel_resp.status_code in (200, 201, 202):
            print(f"✅ Order {target_order_id} cancellation request submitted successfully.")
            
            # ✅ Indentation corrected for the block below
            new_order_payload = {
                "orderType": "LIMIT",
                "session": "SEAMLESS",          # Enables PM + Regular + AM sessions
                "duration": "GOOD_TILL_CANCEL",  # Correct Schwab GTC string
                "orderStrategyType": "SINGLE",
                "price": new_price,
                "orderLegCollection": [
                    {
                        "instruction": "SELL",
                        "quantity": new_quantity,
                        "instrument": {
                            "symbol": target_symbol,
                            "assetType": "EQUITY"
                        }
                    }
                ]
            }
          
            print(f"🚀 Placing new SELL order: {new_quantity} shares of {target_symbol} at ${new_price}...")
            place_resp = client.place_order(account_hash, new_order_payload)
            
            if place_resp.status_code in (200, 201, 202):
                location_header = place_resp.headers.get('location', '')
                new_order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
                print(f"🏆 Success! New SELL order placed. Order ID: {new_order_id}")
            else:
                print(f"❌ Failed to place new order. Status: {place_resp.status_code} | Response: {place_resp.text}")
        else:
            print(f"❌ Failed to cancel original order. Status: {cancel_resp.status_code} | Response: {cancel_resp.text}")
            
    except Exception as e:
        print(f"💥 Critical error executing cancel and place process: {e}")
else:
    print("no present")
