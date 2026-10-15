import os
import datetime
import pandas as pd
import schwabdev
from schwab_tokens import TOKENS_DB
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
client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)

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
# FETCH OPEN/WORKING ORDERS (Dictionary Tracking Structure)
# ==============================================================================
print("Fetching current open orders from Schwab...")
# Structure: {(SYMBOL, INSTRUCTION): {"order_id": str, "qty": int, "price": str}}
open_orders_dict = {}

try:
    now = datetime.datetime.now(datetime.timezone.utc)
    from_time = (now - datetime.timedelta(days=1)).isoformat()
    to_time = now.isoformat()
    
    orders_resp = client.account_orders(account_hash, from_time, to_time, status="WORKING")
    
    if orders_resp.status_code == 200:
        orders_list = orders_resp.json()
        if isinstance(orders_list, list):
            for order in orders_list:
                order_id = str(order.get("orderId"))
                legs = order.get("orderLegCollection", [])
                
                if legs and len(legs) > 0 and order_id:
                    leg = legs[0]
                    o_symbol = leg.get("instrument", {}).get("symbol", "").strip().upper()
                    o_instruction = leg.get("instruction", "").strip().upper()
                    o_qty = int(leg.get("quantity", 0))
                    
                    o_type = order.get("orderType", "").strip().upper()
                    if o_type == "MARKET":
                        o_price = "MARKET"
                    else:
                        price_val = order.get("price")
                        o_price = f"{float(price_val):.2f}" if price_val is not None else "0.00"
                    
                    if o_symbol and o_instruction and o_qty > 0:
                        key = (o_symbol, o_instruction)
                        open_orders_dict[key] = {
                            "order_id": order_id,
                            "qty": o_qty,
                            "price": o_price
                        }
            print(f"Successfully cached {len(open_orders_dict)} active open order targets.\n")
        else:
            print("No active open orders found on account profile.")
    else:
        print(f"⚠️ Warning: Could not fetch open orders (Status {orders_resp.status_code}). Engine will default to fresh placements.")
except Exception as e:
    print(f"⚠️ Warning: Error parsing open orders: {e}. Engine will default to fresh placements.")

# 5. Load and iterate through the Excel file
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "NewTdyExistBuyOrder.xlsx")
try:
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} orders from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# ==============================================================================
# 6. Loop through each row and place or REPLACE the order configuration
# ==============================================================================
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper()
    quantity = int(row['quantity'])
    order_type = str(row['orderType']).strip().upper()
    duration = str(row['duration']).strip().upper()
    session_type = str(row.get('session', 'NORMAL')).strip().upper()
    
    # Filter Constraint: Explicitly handle BUY statuses only
    if "BUY" not in instruction:
        print(f"⏭ Skipping row {index + 1}: Instruction '{instruction}' is not a BUY status instruction.")
        print("-" * 40)
        continue

    # Determine prospective target price matching profile
    if order_type == "MARKET":
        current_row_price = "MARKET"
    else:
        current_row_price = f"{float(row['price']):.2f}"

    print(f"Processing row {index + 1}: {instruction} {quantity} {symbol} @ {current_row_price}...")

    # Extended Hours Session Guard Rule Check
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping row {index + 1}: Extended session '{session_type}' does not allow MARKET orders.")
        print("-" * 40)
        continue

    # Dynamically build standard underlying Schwab payload details
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
        order_payload["price"] = current_row_price

    # Check for an active tracking instance to determine path logic
    lookup_key = (symbol, instruction)
    
    if lookup_key in open_orders_dict:
        existing_order = open_orders_dict[lookup_key]
        existing_id = existing_order["order_id"]
        existing_price = existing_order["price"]
        existing_qty = existing_order["qty"]
        
        # Scenario A: The identical configuration is already working (No changes needed)
        if existing_price == current_row_price and existing_qty == quantity:
            print(f"⏭ Skipping row {index + 1}: Match found. Order ID {existing_id} is already working at identical values.")
            print("-" * 40)
            continue
            
        # Scenario B: Target found, but Price AND/OR Quantity needs modernizing
        # (Replace Route)
        else:
            print(f"🔄 Replacing order ID {existing_id} for {symbol}...")
            print(f"   Old Layout -> Qty: {existing_qty}, Price: {existing_price}")
            print(f"   New Layout -> Qty: {quantity}, Price: {current_row_price}")

            try:
                resp = client.replace_order(account_hash, existing_id, order_payload)
                if resp.status_code in (200, 201):
                    print(f"✅ Replacement order successfully submitted! (Status Code: {resp.status_code})")
                    # Update cache in case duplicate rows for this symbol follow in the sheet
                    open_orders_dict[lookup_key] = {
                        "order_id": existing_id,
                        "qty": quantity,
                        "price": current_row_price
                    }
                else:
                    print(f"❌ Failed to replace order ID {existing_id}. Status: {resp.status_code}")
                    print(f"Response details: {resp.text}")
            except Exception as e:
                print(f"💥 Error processing replace_order sequence: {e}")
            print("-" * 40)
            continue

    # Scenario C: Brand-new positioning target found (Standard Placement Route)
    try:
        resp = client.place_order(account_hash, order_payload)
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! New Order ID: {order_id} (Status: {resp.status_code})")
            
            # Map tracking references to global state memory matrix
            open_orders_dict[lookup_key] = {
                "order_id": order_id,
                "qty": quantity,
                "price": current_row_price
            }
        else:
            print(f"❌ Failed to place new order for {symbol}. Status Code: {resp.status_code}")
            print(f"Response text: {resp.text}")
    except Exception as e:
        print(f"💥 Critical error placing order for {symbol}: {e}")
    print("-" * 40)

print("Batch order processing complete.")
