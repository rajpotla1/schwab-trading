import os
import datetime
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

# Handle variations in Schwab's account structure response safely
if isinstance(linked_accounts, list):
    if len(linked_accounts) > 0:
        # If it's a list of accounts, grab the first one
        account_hash = linked_accounts[0].get('hashValue')
    else:
        account_hash = None
elif isinstance(linked_accounts, dict):
    account_securities = linked_accounts.get('securitiesAccountList', [])
    account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")


# ==============================================================================
# Fetch Open/Working Orders with Price and Quantity Tracking
# ==============================================================================
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
                # Ensure legs is a non-empty list before targeting index 0
                if isinstance(legs, list) and len(legs) > 0:
                    leg = legs[0]  # Extract the dictionary from the list
                    o_symbol = leg.get("instrument", {}).get("symbol", "").strip().upper()
                    o_instruction = leg.get("instruction", "").strip().upper()
                    o_qty = int(leg.get("quantity", 0))
                    
                    o_type = order.get("orderType", "").strip().upper()
                    if o_type == "MARKET":
                        o_price = "MARKET"
                    else:
                        price_val = order.get("price")
                        o_price = f"{float(price_val):.2f}" if price_val is not None else "0.00"
                    
                    # Guard tracks symbol, instruction, price, AND quantity
                    if o_symbol and o_instruction and o_qty > 0:
                        open_orders_set.add((o_symbol, o_instruction, o_price, o_qty))
                        
            print(f"Successfully cached {len(open_orders_set)} active open orders.\n")
        else:
            print("No active open orders found on account profile.")
    else:
        print(f"⚠️ Warning: Could not fetch open orders (Status {orders_resp.status_code}). Guard is inactive.")
except Exception as e:
    print(f"⚠️ Warning: Error parsing open orders: {e}. Duplication guard is inactive.")


# ==============================================================================
# Parameterized Order Placement Engine
# ==============================================================================
def place_single_order(ticker, price, quantity=1, order_type="LIMIT", duration="DAY", instruction="BUY", session_type="SEAMLESS"):
    """
    Places an order for a dynamic quantity of shares (defaults to 1). Returns True if order sent successfully, False otherwise.
    """
    symbol = str(ticker).strip().upper()
    order_type = str(order_type).strip().upper()
    duration = str(duration).strip().upper()
    instruction = str(instruction).strip().upper()
    session_type = str(session_type).strip().upper()
    qty = int(quantity)

    # Determine price key matching structure
    if order_type == "MARKET":
        current_row_price = "MARKET"
    else:
        current_row_price = f"{float(price):.2f}"

    print(f"\nProcessing Order Request: {instruction} {qty} {symbol} @ {current_row_price}...")

    # 🛑 Strict Anti-Duplication Check: Symbol + Instruction + Price + Quantity
    if (symbol, instruction, current_row_price, qty) in open_orders_set:
        print(f"⏭ Skipping: An open order for {instruction} {qty} {symbol} at {current_row_price} already exists.")
        return False

    # ⚠️ Safety Check: Extended hours sessions (SEAMLESS/SEPARATE) DO NOT support MARKET orders.
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping: Extended session '{session_type}' does not allow MARKET orders.")
        return False

    # Dynamically build the Schwab order payload format
    order_payload = {
        "orderType": order_type,
        "session": session_type,
        "duration": duration,
        "orderStrategyType": "SINGLE",
        "orderLegCollection": [
            {
                "instruction": instruction,
                "quantity": qty,
                "instrument": {
                    "symbol": symbol,
                    "assetType": "EQUITY"
                }
            }
        ]
    }

    if order_type == "LIMIT":
        order_payload["price"] = current_row_price

    # Place the trade via API
    try:
        resp = client.place_order(account_hash, order_payload)
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! Order ID: {order_id} (Status: {resp.status_code})")
            return True
        else:
            print(f"❌ Failed to place order for {symbol}. Status Code: {resp.status_code}")
            print(f"Response text: {resp.text}")
            return False
    except Exception as e:
        print(f"💥 Critical error placing order for {symbol}: {e}")
        return False


# ==============================================================================
# Interactive Input Prompt (Enforces exactly one successful execution)
# ==============================================================================
if __name__ == "__main__":
    print("=== Schwab Interactive Order Entry (LIMIT, DAY, BUY) ===")
    print("Application will terminate automatically after exactly ONE successful order execution.\n")
    
    while True:
        try:
            # 1. Prompt for Ticker
            user_ticker = input("Enter Ticker Symbol (or type 'exit'): ").strip()
            if user_ticker.lower() == 'exit':
                print("Exiting order application. Goodbye!")
                break
            if not user_ticker:
                print("❌ Ticker cannot be empty.")
                continue

            symbol_upper = user_ticker.upper()
            suggested_price = None
            bid_str = "N/A"
            ask_str = "N/A"

            # --- Dynamic Live Market Quote Lookup ---
            try:
                quote_resp = client.quote(symbol_upper)
                if quote_resp.status_code == 200:
                    quote_data = quote_resp.json()
                    ticker_data = quote_data.get(symbol_upper, {})
                    quote_details = ticker_data.get("quote", {})
                    
                    last_price = quote_details.get("lastPrice")
                    bid_price = quote_details.get("bidPrice")
                    ask_price = quote_details.get("askPrice")
                    
                    if last_price is not None:
                        bid_str = f"{float(last_price):.2f}"
                    if bid_price is not None:
                        # FIX: Remove the "$" sign from the raw string value
                        suggested_price  = f"{float(bid_price):.2f}"
                    if ask_price is not None:
                        ask_str = f"{float(ask_price):.2f}"
            except Exception:
                pass

            # 2. Prompt for Quantity (Defaults explicitly to 1)
            user_qty_str = input(f"Enter Quantity for {symbol_upper} [Press Enter for 1]: ").strip()
            if not user_qty_str:
                user_qty = 1
            else:
                try:
                    user_qty = int(user_qty_str)
                    if user_qty <= 0:
                        raise ValueError
                except ValueError:
                    print("❌ Invalid quantity. Must be a whole number greater than 0.")
                    continue

            # 3. Prompt for Limit Price containing Quote Data Reference
            # FIX: Explicitly add "$" into the presentation text manually
            if suggested_price:
                print(f"📊 Market Quote -> Bid: ${suggested_price} | Ask: ${ask_str} | Last: ${bid_str}")
                prompt_text = f"Enter Limit Price for {symbol_upper} [${suggested_price}]: "
            else:
                prompt_text = f"Enter Limit Price for {symbol_upper} $: "

            user_price_str = input(prompt_text).strip()
            
            # Apply suggested default option if user skips input via Enter key
            if not user_price_str and suggested_price:
                user_price_str = suggested_price

            try:
                user_price = float(user_price_str)
                if user_price <= 0:
                    raise ValueError
            except ValueError:
                print("❌ Invalid price. Must be a valid numerical value greater than 0.")
                continue

            # 4. Route to engine and check execution status
            order_placed = place_single_order(ticker=symbol_upper, price=user_price, quantity=user_qty)
            
            #  Break out and stop completely if order was fired successfully
            if order_placed:
                print("\n🎯 Execution cap reached (1 order placed). Exiting script safely.")
                break
                
            print("\n" + "="*50 + "\n")

        except KeyboardInterrupt:
            print("\nProgram interrupted. Exiting safely.")
            break
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
