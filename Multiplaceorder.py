import os
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
    # If the response is a dictionary with a list inside it
    account_securities = linked_accounts.get('securitiesAccountList', [])
    account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

# 5. Load and iterate through the Excel file
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "trades.xlsx") 

try:
    # Read the excel sheet into a pandas DataFrame
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} orders from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# ==============================================================================
# REPLACE YOUR EXISTING SECTION 6 LOOP WITH THIS IMPLEMENTATION
# ==============================================================================

# 6. Loop through each row and place the order
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper() # "BUY" or "SELL"
    quantity = int(row['quantity'])
    order_type = str(row['orderType']).strip().upper()   # "LIMIT" or "MARKET"
    duration = str(row['duration']).strip().upper()       # "DAY" or "GOOD_TILL_CANCEL"
    
    # 🛑 Fix: Dynamically grab session from Excel row, fallback to 'NORMAL' if empty
    session_type = str(row.get('session', 'NORMAL')).strip().upper()

    print(f"Processing row {index + 1}: {instruction} {quantity} {symbol}...")

    # ⚠️ Safety Check: Extended hours sessions (SEAMLESS/SEPARATE) DO NOT support MARKET orders.
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping row {index + 1}: Extended session '{session_type}' does not allow MARKET orders.")
        print("-" * 40)
        continue

    # Dynamically build the Schwab order payload format
    order_payload = {
        "orderType": order_type,
        "session": session_type,        # ⚡ Linked dynamically now!
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

    # Only append price if it's a LIMIT order
    if order_type == "LIMIT":
        # Schwab API prefers string formats for decimals to avoid float rounding bugs
        order_payload["price"] = f"{float(row['price']):.2f}"

    # Place the trade via API
    try:
        resp = client.place_order(account_hash, order_payload)
        
        # 201 is the standard HTTP status code for a successfully created order
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"✅ Success! Order ID: {order_id} (Status: {resp.status_code})")
        else:
            print(f"❌ Failed to place order for {symbol}. Status Code: {resp.status_code}")
            print(f"Response text: {resp.text}")
            
    except Exception as e:
        print(f"💥 Critical error placing order for {symbol}: {e}")
        
    print("-" * 40)

print("Batch order processing complete.")
