import os
import pandas as pd
import schwabdev
from schwab_tokens import TOKENS_DB
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 2. Initialize the Schwab client
try:
    client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)
except Exception as e:
    raise SystemExit(f"💥 Critical Error: Failed to initialize Schwab Client. Details: {e}")

# 3. Retrieve linked accounts and extract the account hash
try:
    linked_accounts = client.linked_accounts().json()
    if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
        account_hash = linked_accounts[0].get('hashValue')
    elif isinstance(linked_accounts, dict):
        account_securities = linked_accounts.get('securitiesAccountList', [])
        account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
    else:
        account_hash = None
except Exception as e:
    raise SystemExit(f"💥 Critical Error: Could not connect to your Schwab account to get linked accounts. Details: {e}")

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

# Prompt user for the target ticker symbol to check (auto-capitalized)
search_symbol = input("Enter ticker symbol to check portfolio position and open LIMIT buy orders: ").strip().upper()
if not search_symbol:
    raise SystemExit("Error: No ticker symbol provided.")

# =====================================================================
# NEW STEP: Fetch current portfolio positions for the symbol
# =====================================================================
print(f"\nChecking current portfolio positions for {search_symbol}...")
try:
    # Query account details with fields='positions' to expose current holdings
    acct_resp = client.account_details(accountHash=account_hash, fields="positions")
    if acct_resp.status_code == 200:
        acct_data = acct_resp.json()
        
        # Dig into the response payload structure
        securities_account = acct_data.get("securitiesAccount", {})
        positions = securities_account.get("positions", [])
        
        found_position = None
        for pos in positions:
            pos_symbol = pos.get("instrument", {}).get("symbol", "").upper()
            if pos_symbol == search_symbol:
                found_position = pos
                break
        
        if found_position:
            long_qty = found_position.get("longQuantity", 0)
            short_qty = found_position.get("shortQuantity", 0)
            net_qty = long_qty - short_qty
            market_value = found_position.get("marketValue", 0)
            avg_price = found_position.get("averagePrice", 0)
            
            print(f"📊 CURRENT POSITION FOUND:")
            print(f"   • Total Shares: {net_qty:,.2f}")
            print(f"   • Avg Cost Basis: ${avg_price:,.2f}")
            print(f"   • Market Value: ${market_value:,.2f}")
        else:
            print(f"📊 CURRENT POSITION: You do not currently hold any shares of {search_symbol}.")
    else:
        print(f"⚠️ Warning: Could not fetch position details. Status code {acct_resp.status_code}")
except Exception as e:
    print(f"💥 Error fetching position details: {e}")


# 4. Auto-calculate the boundaries for the time window (Look back 30 days)
now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

print(f"\nFetching orders to check if {search_symbol} is in a working limit buy order...")

# 5. Query WORKING statuses and filter for LIMIT BUY orders matching the symbol
try:
    resp = client.account_orders(
        accountHash=account_hash,
        fromEnteredTime=from_time,
        toEnteredTime=to_time,
        status="WORKING"
    )
    if resp.status_code == 200:
        orders = resp.json()
        matching_limit_buys = []
        
        for o in orders:
            # Filter specifically for LIMIT orders
            order_type = str(o.get("orderType", "")).upper()
            if "LIMIT" not in order_type:
                continue
                
            legs = o.get('orderLegCollection', [])
            leg = legs[0] if legs else {}
            instruction = str(leg.get("instruction", "")).upper()
            
            # Filter for active "BUY" configurations only
            if "BUY" not in instruction:
                continue
                
            symbol = leg.get("instrument", {}).get("symbol", "").upper()
            
            # Check if this order matches the symbol the user prompted for
            if symbol == search_symbol:
                quantity = o.get("quantity")
                price = o.get("price")  # Extract the limit price
                order_id = o.get("orderId")
                matching_limit_buys.append({
                    "OrderID": order_id,
                    "Quantity": quantity,
                    "Price": f"${price:,.2f}" if price is not None else "N/A"
                })

        # 6. Display the filtered results (Removed the "NOT FOUND" block)
        if matching_limit_buys:
            print(f"\n🎯 MATCH: Found {len(matching_limit_buys)} open LIMIT BUY order(s) for {search_symbol}!\n")
            print(f"{'SYMBOL':<10} | {'QUANTITY':<10} | {'LIMIT PRICE':<12} | {'ORDER ID':<12}")
            print("-" * 52)
            for order in matching_limit_buys:
                print(f"{search_symbol:<10} | {order['Quantity']:<10} | {order['Price']:<12} | {order['OrderID']:<12}")
    else:
        print(f"⚠️ Warning: Server responded with status code {resp.status_code}")
        print(f"Response details: {resp.text}")
except Exception as e:
    print(f"💥 Critical error fetching open orders: {e}")
