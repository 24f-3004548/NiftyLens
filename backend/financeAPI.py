import time
from sqlalchemy import create_engine, text
import yfinance as yf
import pandas as pd
import json
import os
from datetime import datetime

# --- UPDATE YOUR PASSWORD HERE ---
engine = create_engine('mysql+pymysql://root:boopalan@localhost/PowerBITest')
# -------------------------------

# -------------------------
# 1. STATIC DATA SETUP
# -------------------------
nifty50 = [
    "RELIANCE.NS","TCS.NS","HDFCBANK.NS","INFY.NS","ICICIBANK.NS",
    "HINDUNILVR.NS","ITC.NS","SBIN.NS","BHARTIARTL.NS","KOTAKBANK.NS",
    "LT.NS","AXISBANK.NS","ASIANPAINT.NS","MARUTI.NS","SUNPHARMA.NS",
    "ULTRACEMCO.NS","TITAN.NS","NESTLEIND.NS","BAJFINANCE.NS","WIPRO.NS",
    "NTPC.NS","POWERGRID.NS","ONGC.NS","M&M.NS","TMPV.NS",
    "HCLTECH.NS","TECHM.NS","JSWSTEEL.NS","TATASTEEL.NS","INDUSINDBK.NS",
    "ADANIENT.NS","ADANIPORTS.NS","BAJAJFINSV.NS","DRREDDY.NS","CIPLA.NS",
    "GRASIM.NS","BRITANNIA.NS","DIVISLAB.NS","HEROMOTOCO.NS","EICHERMOT.NS",
    "COALINDIA.NS","APOLLOHOSP.NS","HDFCLIFE.NS","SBILIFE.NS","UPL.NS",
    "BAJAJ-AUTO.NS","TATACONSUM.NS","PIDILITIND.NS","HINDALCO.NS","SHREECEM.NS"
]

SECTOR_FILE = "sector_map.json"

def get_sector(stock):
    try:
        ticker = yf.Ticker(stock)
        return ticker.info.get("sector", "Unknown")
    except:
        return "Unknown"

# Load or fetch sector mapping once
if os.path.exists(SECTOR_FILE):
    with open(SECTOR_FILE, "r") as f:
        sector_map = json.load(f)
else:
    sector_map = {}
    print("Fetching sector data for the first time...")
    for stock in nifty50:
        sector_map[stock] = get_sector(stock)
    with open(SECTOR_FILE, "w") as f:
        json.dump(sector_map, f)

# -------------------------
# 2. CORE FETCH FUNCTION
# -------------------------
def fetch_and_process_data(period_str):
    """Fetches data for Nifty 50 and the Index, returning a clean DataFrame."""
    print(f"Downloading data for period: {period_str}...")
    
    data = yf.download(nifty50, period=period_str, interval="1d", group_by="ticker", threads=True)
    records = []

    for stock in nifty50:
        try:
            # Check if stock data exists (sometimes yfinance fails on single tickers)
            if stock not in data or data[stock].empty:
                continue
                
            df = data[stock].copy().reset_index()
            
            # --- THE FIX: Flatten the weird tuple column names yfinance creates ---
            df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]
            # ----------------------------------------------------------------------
            
            df["symbol"] = stock
            df["sector"] = sector_map.get(stock, "Unknown")

            df = df.rename(columns={
                "Date": "date", "Datetime": "date", # Handles both 1d and intraday index names
                "Open": "open", "High": "high", "Low": "low", 
                "Close": "close", "Volume": "volume"
            })
            
            # Strip timezone info from dates so MySQL doesn't throw an error
            if df['date'].dt.tz is not None:
                df['date'] = df['date'].dt.tz_localize(None)
                
            df = df[["date", "symbol", "sector", "open", "high", "low", "close", "volume"]]
            records.append(df)
        except Exception as e:
            print(f"Error processing {stock}: {e}")

    final_df = pd.concat(records)

    # Fetch Nifty Index
    nifty = yf.download("^NSEI", period=period_str, interval="1d").reset_index()
    
    # --- THE FIX (Applied to the Index too) ---
    nifty.columns = [c[0] if isinstance(c, tuple) else c for c in nifty.columns]
    # ------------------------------------------
    
    nifty = nifty.rename(columns={
        "Date": "date", "Datetime": "date",
        "Open": "open", "High": "high", "Low": "low", 
        "Close": "close", "Volume": "volume"
    })
    
    if nifty['date'].dt.tz is not None:
        nifty['date'] = nifty['date'].dt.tz_localize(None)
        
    nifty["symbol"] = "NIFTY"
    nifty["sector"] = "Index"
    
    # Keep only matching columns for concat to avoid errors
    nifty = nifty[["date", "symbol", "sector", "open", "high", "low", "close", "volume"]]

    final_df = pd.concat([final_df, nifty], ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["date", "symbol"])
    

    return final_df

# -------------------------
# 3. INITIAL HISTORICAL LOAD
# -------------------------
print("Checking database for existing historical data...")
try:
    with engine.connect() as conn:
        result = conn.execute(text("SELECT COUNT(*) FROM powerbitest")).scalar()
        has_data = result > 0
except Exception:
    has_data = False

if not has_data:
    print("No data found. Initiating 3-year historical load...")
    historical_df = fetch_and_process_data("3y")
    # historical_df.to_csv("nifty_raw_data.csv", index=False)
    historical_df.to_sql("powerbitest", engine, if_exists="replace", index=False)
    print("Historical data successfully loaded into SQL.")
else:
    print("Historical data already exists. Skipping initial load.")

# -------------------------
# 4. REAL-TIME UPDATE LOOP
# -------------------------
print("Starting real-time data push. Press Ctrl+C to stop.")

while True:
    # Fetch only the latest day's data
    live_df = fetch_and_process_data("1d")
    
    # Get the unique dates from the live data (usually just today's date)
    latest_dates = live_df['date'].dt.strftime('%Y-%m-%d').unique()
    date_strings = "', '".join(latest_dates)
    
    # SQL UPSERT TRICK: Delete today's rows from the DB, then append the new live values
    with engine.begin() as conn:
        delete_query = f"DELETE FROM powerbitest WHERE DATE(date) IN ('{date_strings}')"
        conn.execute(text(delete_query))
        
    # Append the fresh data
    live_df.to_sql("powerbitest", engine, if_exists="append", index=False)
    
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Live data updated successfully.")
    
    # Wait 20 minutes (1200 seconds)
    time.sleep(1200)
