import os
import pandas as pd
import schwabdev
from dotenv import load_dotenv

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

# 2. Check environment variables BEFORE initializing the client
if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 3. Initialize the Schwab client
client = schwabdev.Client(client_id, client_secret, callback_url)

# 4. Retrieve linked accounts and extract the account hash
linked_accounts = client.linked_accounts().json()

# Handle variations in Schwab's account structure response
if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
    account_hash = linked_accounts[0].get('hashValue')
elif isinstance(linked_accounts, dict):
    account_securities = linked_accounts.get('securitiesAccountList', [])
    account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")
# ==============================================================================
# NEW STEP: Fetch Open/Working Orders with Price Tracking
# ==============================================================================
import datetime

print("Fetching current open orders from Schwab...")
open_orders_set = set()
try:
    now = datetime.datetime.now(datetime.timezone.utc)
    from_time = (now - datetime.timedelta(days=1)).isoformat() 
    to_time = now.isoformat()

    orders_resp = client.account_orders(account_hash, from_time, to_time, status="WORKING")
    
    if orders_resp.status_code == 200:
        orders_list = orders_resp.json()
        
        if isinstance(orders_list, list):
            for order in orders_list:
                legs = order.get("orderLegCollection", [])
                if legs and len(legs) > 0:
                    leg = legs[0]  
                    o_symbol = leg.get("instrument", {}).get("symbol", "").strip().upper()
                    o_instruction = leg.get("instruction", "").strip().upper()
                    o_qty = int(leg.get("quantity", 0))
                    
                    # 💵 Extract order price type to differentiate duplicates
                    o_type = order.get("orderType", "").strip().upper()
                    if o_type == "MARKET":
                        o_price = "MARKET"
                    else:
                        # Normalize price formatting to 2 decimal places to match Excel handling
                        price_val = order.get("price")
                        o_price = f"{float(price_val):.2f}" if price_val is not None else "0.00"
                    
                    if o_symbol and o_instruction and o_qty > 0:
                        # The tracking footprint now includes price
                        open_orders_set.add((o_symbol, o_instruction, o_qty, o_price))
                        
            print(f"Successfully cached {len(open_orders_set)} active open orders.\n")
        else:
            print("No active open orders found on account profile.")
    else:
        print(f"⚠️ Warning: Could not fetch open orders (Status {orders_resp.status_code}). Guard is inactive.")
except Exception as e:
    print(f"⚠️ Warning: Error parsing open orders: {e}. Duplication guard is inactive.")



# 5. Load and iterate through the Excel file
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "NewBuyTrades.xlsx")
try:
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} orders from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# ==============================================================================
# 6. Loop through each row and place the order (With Symbol, Qty, and Price Guard)
# ==============================================================================
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper()  
    quantity = int(row['quantity'])
    order_type = str(row['orderType']).strip().upper()      
    duration = str(row['duration']).strip().upper()        
    session_type = str(row.get('session', 'NORMAL')).strip().upper()

    # 💵 Determine price key matching structure
    if order_type == "MARKET":
        current_row_price = "MARKET"
    else:
        current_row_price = f"{float(row['price']):.2f}"

    print(f"Processing row {index + 1}: {instruction} {quantity} {symbol} @ {current_row_price}...")

    # 🛑 Strict Anti-Duplication Check: Symbol + Instruction + Quantity + Price
    if (symbol, instruction, quantity, current_row_price) in open_orders_set:
        print(f"⏭ Skipping row {index + 1}: An open order for {instruction} {quantity} {symbol} at {current_row_price} already exists.")
        print("-" * 40)
        continue

    # ⚠️ Safety Check: Extended hours sessions (SEAMLESS/SEPARATE) DO NOT support MARKET orders.
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping row {index + 1}: Extended session '{session_type}' does not allow MARKET orders.")
        print("-" * 40)
        continue

    # Dynamically build the Schwab order payload format
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

    # Only append price payload configuration if it's a LIMIT order
    if order_type == "LIMIT":
        order_payload["price"] = current_row_price

    # Place the trade via API
    try:
        resp = client.place_order(account_hash, order_payload)
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! Order ID: {order_id} (Status: {resp.status_code})")
            
            # ⚡ Instantly add to guard ledger to catch duplicate rows later in the same sheet
            open_orders_set.add((symbol, instruction, quantity, current_row_price))
        else:
            print(f"❌ Failed to place order for {symbol}. Status Code: {resp.status_code}")
            print(f"Response text: {resp.text}")
    except Exception as e:
        print(f"💥 Critical error placing order for {symbol}: {e}")
        
    print("-" * 40)

print("Batch order processing complete.")
