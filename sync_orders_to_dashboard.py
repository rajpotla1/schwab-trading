import os
import sys
from datetime import datetime, timezone
import pandas as pd
from dotenv import load_dotenv
import schwabdev
from schwab_tokens import TOKENS_DB
import openpyxl

# --- SETUP CONFIGURATIONS ---
# Reverted back to the original standard workbook file name
MASTER_FILENAME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "Trading_Master_Dashboard.xlsm")

# 1. Load system environment profiles
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: Sensitive credential keys failed to load from your .env file.")

# 2. Initialize Schwab connection
try:
    client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)
except Exception as e:
    raise SystemExit(f"💥 Connection Failure: Schwab Client rejected authentication: {e}")


def get_todays_executed_trades():
    """
    Fetches all executed trades for today across all linked accounts.
    Returns a DataFrame with trade details.
    """
    try:
        linked_accounts_raw = client.linked_accounts()
        linked_accounts = linked_accounts_raw.json() if hasattr(linked_accounts_raw, "json") else linked_accounts_raw
        
        if isinstance(linked_accounts, list):
            accounts_list = linked_accounts
        elif isinstance(linked_accounts, dict):
            accounts_list = linked_accounts.get("securitiesAccountList", [linked_accounts])
        else:
            accounts_list = []
        
        all_trades = []
        now_utc = datetime.now(timezone.utc)
        start_of_today = datetime(now_utc.year, now_utc.month, now_utc.day, tzinfo=timezone.utc)
        from_time = start_of_today.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        
        for account_item in accounts_list:
            if not isinstance(account_item, dict):
                continue
            
            sec_account = account_item.get("securitiesAccount", account_item)
            account_hash = sec_account.get("hashValue") or sec_account.get("accountNumber")
            
            if not account_hash:
                continue
            
            account_disp = f"[...{str(account_hash)[-8:]}]"
            
            try:
                orders_raw = client.account_orders(
                    accountHash=account_hash,
                    fromEnteredTime=from_time,
                    toEnteredTime=to_time,
                    status="FILLED"
                )
                orders_data = orders_raw.json() if hasattr(orders_raw, "json") else orders_raw
                
                if not isinstance(orders_data, list):
                    orders_data = [orders_data] if orders_data else []
                
                for order in orders_data:
                    if not isinstance(order, dict):
                        continue
                    
                    order_leg_collection = order.get("orderLegCollection", [])
                    for leg in order_leg_collection:
                        if not isinstance(leg, dict):
                            continue
                        
                        symbol = leg.get("instrument", {}).get("symbol", "N/A")
                        quantity = float(leg.get("quantity", 0) or 0)
                        instruction = leg.get("instruction", "").upper()
                        
                        all_trades.append({
                            "Account": account_disp,
                            "Symbol": symbol,
                            "Quantity": quantity,
                            "Type": instruction,
                            "Order ID": order.get("orderId", "N/A"),
                            "Order Time": order.get("enteredTime", "N/A"),
                            "Status": order.get("status", "N/A")
                        })
            except Exception as order_err:
                print(f"⚠ Could not fetch orders for account {account_disp}: {order_err}")
                continue
        
        if all_trades:
            df_trades = pd.DataFrame(all_trades)
            return df_trades
        else:
            return pd.DataFrame(columns=["Account", "Symbol", "Quantity", "Type", "Order ID", "Order Time", "Status"])
    
    except Exception as e:
        print(f"⚠ Error fetching today's trades: {e}")
        return pd.DataFrame(columns=["Account", "Symbol", "Quantity", "Type", "Order ID", "Order Time", "Status"])


def update_excel_sheet_in_place(ws, df, match_on_account=True):
    """
    Updates active rows, appends new rows, and deletes rows for closed positions.
    Guarantees Row 1 is reserved exclusively for column headers.
    """
    # 1. Ensure the sheet has proper labels in Row 1 if it's completely empty
    if ws.max_row <= 1 or ws.cell(row=1, column=1).value is None:
        ws.delete_rows(1, ws.max_row + 10)
        ws.append(list(df.columns))

    # 2. Map existing sheet data entries starting strictly from Row 2
    row_map = {}
    for r in range(2, ws.max_row + 1):
        if match_on_account:
            acc_val = str(ws.cell(row=r, column=1).value or "").strip()
            sym_val = str(ws.cell(row=r, column=2).value or "").strip()
            if acc_val and sym_val:
                row_map[(acc_val, sym_val)] = r
        else:
            dash_sym = str(ws.cell(row=r, column=1).value or "").strip()
            if dash_sym:
                row_map[(dash_sym,)] = r

    # Keep track of keys seen in this update cycle to identify active vs closed assets
    seen_keys = set()
    next_empty_row = ws.max_row + 1  # Track the next row to write to

    # 3. Update matching rows or append completely new ones to the bottom
    for _, row_data in df.iterrows():
        if match_on_account:
            match_key = (str(row_data["Account"]).strip(), str(row_data["Symbol"]).strip())
        else:
            match_key = (str(row_data["Symbol"]).strip(),)

        seen_keys.add(match_key)

        if match_key in row_map:
            target_row = row_map[match_key]
        else:
            # Write to the next empty row without appending (to work around openpyxl max_row issue)
            target_row = next_empty_row
            next_empty_row += 1

        # Write data column by column to prevent clearing custom data on the far right
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=target_row, column=col_idx)
            cell.value = row_data[col_name]
            
            # Re-apply exact formatting rules based on keywords
            if "Quantity" in col_name:
                cell.number_format = '#,##0.00'
            elif "Price" in col_name or "Basis" in col_name or "Value" in col_name or "Cost" in col_name or "High" in col_name or "Low" in col_name or "Close" in col_name or "$" in col_name:
                cell.number_format = '$#,##0.00'
            elif "%" in col_name or "Change" in col_name:
                cell.number_format = '0.00%'

    # 4. Clean up closed positions: Delete rows that are in Excel but absent from active payload
    for r in range(next_empty_row - 1, 1, -1):
        if match_on_account:
            acc_val = str(ws.cell(row=r, column=1).value or "").strip()
            sym_val = str(ws.cell(row=r, column=2).value or "").strip()
            check_key = (acc_val, sym_val)
        else:
            dash_sym = str(ws.cell(row=r, column=1).value or "").strip()
            check_key = (dash_sym,)

        # Skip rows that don't have valid keys or rows we just successfully updated
        if not any(check_key) or check_key in seen_keys:
            continue

        print(f"🗑 Closed Position Detected: Removing row {r} for {check_key}")
        ws.delete_rows(r, 1)


def sync_trading_engine_to_excel():
    if not os.path.exists(MASTER_FILENAME):
        print(f"❌ Error: Could not find '{MASTER_FILENAME}'. Please create this file manually first!")
        return

    # A. FETCH INVENTORY DATA FROM SCHWAB ACCOUNTS
    try:
        print("🔍 Querying Schwab API for linked profiles and live balances...")
        accounts_raw = client.account_details_all(fields="positions")
        accounts_data = accounts_raw.json() if hasattr(accounts_raw, "json") else accounts_raw
    except Exception as e:
        print(f"⚠ API Error: Could not reach Schwab server: {e}")
        return

    if isinstance(accounts_data, list):
        accounts_list = accounts_data
    elif isinstance(accounts_data, dict):
        accounts_list = accounts_data.get("securitiesAccountList", [accounts_data])
    else:
        accounts_list = []

    print(f"ℹ Found {len(accounts_list)} account container(s) to scan.")
    
    # Gather all unique symbols across accounts to batch query quotes efficiently
    unique_symbols = set()
    raw_positions_cache = []

    for account_item in accounts_list:
        if not isinstance(account_item, dict):
            continue
        sec_account = account_item.get("securitiesAccount", account_item)
        account_identifier = sec_account.get("hashValue") or sec_account.get("accountNumber")
        if not account_identifier:
            continue
        
        account_disp = f"[...{str(account_identifier)[-8:]}]"
        positions_list = sec_account.get("positions", [])
        
        for pos in positions_list:
            if not isinstance(pos, dict):
                continue
            symbol = pos.get("instrument", {}).get("symbol", "N/A")
            if symbol != "N/A":
                unique_symbols.add(symbol)
            raw_positions_cache.append((account_disp, pos))

    # B. SECONDARY FETCH: LOOKUP LIVE MARKET DATA QUOTES
    quotes_lookup = {}
    if unique_symbols:
        try:
            print(f"⚡ Batch querying real-time market data quotes for {len(unique_symbols)} asset(s)...")
            symbols_list = list(unique_symbols)
            quotes_raw = client.quotes(symbols=symbols_list)
            quotes_data = quotes_raw.json() if hasattr(quotes_raw, "json") else quotes_raw
            
            if isinstance(quotes_data, dict):
                quotes_lookup = quotes_data
        except Exception as q_err:
            print(f"⚠ Market Data Service Unavailable: Falling back to estimated values. Error: {q_err}")

    # C. COMBINE ENDPOINT PAYLOADS
    all_positions_data = []

    for account_disp, pos in raw_positions_cache:
        qty = float(pos.get("longQuantity", 0) or pos.get("shortQuantity", 0) or 0)
        avg_price = float(pos.get("averagePrice", 0.0) or pos.get("averageCost", 0.0) or 0.0)
        market_val = float(pos.get("marketValue", 0.0) or 0.0)
        symbol = pos.get("instrument", {}).get("symbol", "N/A")
        
        # Pull from our live Quotes dictionary mapping
        symbol_quote = quotes_lookup.get(symbol, {})
        quote_metrics = symbol_quote.get("quote", {}) if isinstance(symbol_quote, dict) else {}
        
        # Extract live values if quotes exist; otherwise use the backup calculation
        if quote_metrics:
            current_price = float(quote_metrics.get("lastPrice", 0.0) or quote_metrics.get("mark", 0.0) or 0.0)
            day_high = float(quote_metrics.get("highPrice", 0.0) or quote_metrics.get("dayHigh", 0.0) or 0.0)
            day_low = float(quote_metrics.get("lowPrice", 0.0) or quote_metrics.get("dayLow", 0.0) or 0.0)
            prev_close = float(quote_metrics.get("closePrice", 0.0) or quote_metrics.get("previousClose", 0.0) or 0.0)
            pct_change = float(quote_metrics.get("netChangePercent", 0.0) or 0.0) / 100.0
        else:
            # Automatic estimated price fallback logic
            current_price = float(pos.get("marketPrice", 0.0))
            if current_price == 0.0 and qty != 0:
                current_price = market_val / qty
            day_high = current_price
            day_low = current_price
            prev_close = current_price
            pct_change = float(pos.get("netChangePercent", 0.0) or 0.0) / 100.0
            
        cost_basis = qty * avg_price
        gain_loss_dlr = market_val - cost_basis
        gain_loss_pct = (gain_loss_dlr / cost_basis) if cost_basis != 0 else 0.0
        
        all_positions_data.append({
            "Account": account_disp,
            "Symbol": symbol,
            "Quantity": qty,
            "Cost/Share": avg_price,
            "Average Price": avg_price,
            "Cost Basis": cost_basis,
            "Market Value": market_val,
            "Current Price": current_price,
            "Day High": day_high,
            "Day Low": day_low,
            "Prev Close": prev_close,
            "% Change": pct_change,
            "Gain/Loss $": gain_loss_dlr,
            "Gain/Loss %": gain_loss_pct
        })

        # Convert to standard DataFrame for unified sorting (Largest Gain/Loss $ at the top)
    df_positions = pd.DataFrame(all_positions_data)
    if not df_positions.empty:
        df_positions = df_positions.sort_values(by="Gain/Loss $", ascending=False)
    else:
        df_positions = pd.DataFrame(columns=["Account", "Symbol", "Quantity", "Cost/Share", "Average Price", "Cost Basis", "Market Value", "Current Price", "Day High", "Day Low", "Prev Close", "% Change", "Gain/Loss $", "Gain/Loss %"])

    # C2. FETCH TODAY'S EXECUTED TRADES
    print("\n📈 Fetching today's executed trades...")
    df_todays_trades = get_todays_executed_trades()

    print(f"📊 Total open positions to sync: {len(df_positions)}")
    print(f"📊 Total today's trades to sync: {len(df_todays_trades)}")

    # D. IN-PLACE EXCEL UPDATE BY INDEX POSITION
    try:
        print(f"\n💾 Opening '{MASTER_FILENAME}' to dynamically sync sheets...")
        workbook = openpyxl.load_workbook(MASTER_FILENAME, keep_vba=True)

        # Ensure the workbook has at least three worksheets available to update
        while len(workbook.worksheets) < 3:
            workbook.create_sheet()

        # --- 1. Update Raw_Positions Sheet (with all open positions) ---
        ws_raw = None
        for ws in workbook.worksheets:
            if ws.title == "Raw_Positions":
                ws_raw = ws
                break
        
        if ws_raw is None:
            ws_raw = workbook.worksheets[0]
        
        print(f"📊 Updating '{ws_raw.title}' sheet with all open positions...")
        update_excel_sheet_in_place(ws_raw, df_positions, match_on_account=True)

        # --- 2. Update Dashboard Sheet (with positions) ---
        ws_dashboard = None
        for ws in workbook.worksheets:
            if ws.title == "Dashboard":
                ws_dashboard = ws
                break
        
        if ws_dashboard is None and len(workbook.worksheets) > 1:
            ws_dashboard = workbook.worksheets[1]
        
        if ws_dashboard:
            print(f"📊 Updating '{ws_dashboard.title}' sheet with positions...")
            update_excel_sheet_in_place(ws_dashboard, df_positions, match_on_account=True)

        # --- 3. Update Today's Trades Sheet ---
        ws_trades = None
        for ws in workbook.worksheets:
            if ws.title == "Today Trades":
                ws_trades = ws
                break
        
        if ws_trades is None:
            ws_trades = workbook.create_sheet("Today Trades")
        
        print(f"📊 Updating '{ws_trades.title}' sheet with today's executed trades...")
        if not df_todays_trades.empty:
            try:
                update_excel_sheet_in_place(ws_trades, df_todays_trades, match_on_account=True)
            except Exception as trade_err:
                print(f"⚠ Could not update trades sheet: {trade_err}")
        else:
            print("ℹ No trades executed today")
        
    # Fixed code:
        workbook.save(MASTER_FILENAME)
        print(f"🔄 Success: First and second sheets updated perfectly at {datetime.now().strftime('%H:%M:%S')}!")

        # 🚀 ADD THIS LINE TO AUTOMATICALLY OPEN THE EXCEL FILE:
        os.startfile(MASTER_FILENAME)


    except PermissionError:
        print(f"❌ Excel Sync Failed: The file '{MASTER_FILENAME}' is open in Excel. Close it and run the script again!")

    except Exception as ex:
        print(f"❌ Workbook Manipulation Error: {ex}")


if __name__ == "__main__":
    print("🚀 Launching Self-Hosted Non-Destructive Dashboard Engine...")
    sync_trading_engine_to_excel()
