import os
import pandas as pd
import schwabdev
from schwab_tokens import TOKENS_DB
from dotenv import load_dotenv
from datetime import datetime, timezone

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

# 4. Auto-calculate boundaries for today (midnight to now UTC)
now_utc = datetime.now(timezone.utc)
start_of_today = datetime(now_utc.year, now_utc.month, now_utc.day, tzinfo=timezone.utc)

from_time = start_of_today.strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

print("Fetching today's executed orders from Schwab...")

# 5. Query orders from the API
try:
    resp = client.account_orders(
        accountHash=account_hash,
        fromEnteredTime=from_time,
        toEnteredTime=to_time
    )
    
    if resp.status_code == 200:
        orders = resp.json()
        raw_data = []
        
        for o in orders:
            legs = o.get('orderLegCollection', [])
            leg = legs[0] if legs else {}
            
            instruction = str(leg.get("instruction", "")).upper()
            status = str(o.get("status", "")).upper()
            
            # Strict Filters: BUY instructions + executed statuses only
            if "BUY" not in instruction:
                continue
            if status not in ["FILLED", "PARTIALLY_FILLED"]:
                continue
                
            # Quantity and Pricing Logic
            price = o.get("price") or leg.get("instrument", {}).get("lastPrice", 0)
            
            # Value calculation based strictly on filled assets
            if status == "FILLED":
                shares_filled = o.get("quantity", 0)
            elif status == "PARTIALLY_FILLED":
                shares_filled = o.get("filledQuantity", 0)
            else:
                shares_filled = 0
                
            executed_value = shares_filled * price
                
            raw_data.append({
                "Symbol": leg.get("instrument", {}).get("symbol", "UNKNOWN"),
                "Shares Filled": shares_filled,
                "Executed Value": executed_value
            })
            
        if raw_data:
            # --- AGGREGATION & GROUPING BY TICKER ---
            df_raw = pd.DataFrame(raw_data)
            
            # Group by symbol to calculate total filled shares and executed cash value
            df = df_raw.groupby("Symbol", as_index=False).agg({
                "Shares Filled": "sum",
                "Executed Value": "sum"
            })
            
            # Dynamically calculate the Volume Weighted Average Price (VWAP) for the grouped tickers
            df["Avg Price"] = df["Executed Value"] / df["Shares Filled"]
            
            # Reorder columns for natural reading flow
            df = df[["Symbol", "Shares Filled", "Avg Price", "Executed Value"]]
            grand_total = df['Executed Value'].sum()
            
            # --- COMMAND LINE DISPLAY ---
            print("\n✅ Today's Grouped Buy Orders (Filled & Partially Filled):\n")
            print(df.to_string(index=False, formatters={
                'Shares Filled': '{:,}'.format,
                'Avg Price': '${:,.2f}'.format,
                'Executed Value': '${:,.2f}'.format
            }))
            print("-" * 65)
            print(f"{'':<40}GRAND TOTAL: ${grand_total:,.2f}\n")
            
            # --- ADVANCED EXCEL EXPORT WITH FORMATTING ---
            filename = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "schwab_buy_orders_today.xlsx")
            
            # Write out file with native Excel formats instead of string conversion
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name="Today's Buys", index=False)
                
                workbook  = writer.book
                worksheet = writer.sheets["Today's Buys"]
                
                # Format mappings for currency strings directly inside spreadsheet cells
                currency_format = workbook.add_format({'num_format': '$#,##0.00'})
                integer_format = workbook.add_format({'num_format': '#,##0'})
                bold_format = workbook.add_format({'bold': True})
                bold_currency_format = workbook.add_format({'bold': True, 'num_format': '$#,##0.00'})
                
                # Set dynamic column formatting and clean column widths
                worksheet.set_column('A:A', 14) # Symbol
                worksheet.set_column('B:B', 16, integer_format)  # Shares Filled
                worksheet.set_column('C:C', 16, currency_format) # Avg Price
                worksheet.set_column('D:D', 18, currency_format) # Executed Value
                
                # Append Summary Row in Excel
                last_row = len(df) + 1
                worksheet.write(last_row, 0, "GRAND TOTAL", bold_format)
                worksheet.write(last_row, 3, f"=SUM(D2:D{last_row})", bold_currency_format)
                
            print(f"📁 Successfully exported formatted summary data to Excel: {filename}")
            
        else:
            print("ℹ️ No filled or partially filled BUY orders found for today.")
    else:
        print(f"⚠️ Warning: Server responded with status code {resp.status_code}")
        print(f"Response details: {resp.text}")
        
except Exception as e:
    print(f"💥 Critical error processing orders: {e}")
