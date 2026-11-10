import os
import time
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from tvDatafeed import TvDatafeed, Interval
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# =====================================================================
# CONFIGURATION PARAMETERS (STABILIZATION & STRATEGY)
# =====================================================================
BOX_PCT = 0.02          # 2.0% Box Size (6% move required for 3-box reversal)
REVERSAL = 3            # 3-Box Reversal
EMA_SPAN = 3            # 3-Day EMA smoothing to dampen single-day clustering swings
OVERSOLD_ENTRY = 30     # BPI Mean-Reversion Entry Trigger
OVERBOUGHT_EXIT = 70    # BPI Overbought Take-Profit Trigger
SMA_PERIOD = 200        # Macro Trend Regime Filter

# =====================================================================
# 1. HIGH-STABILITY TRADINGVIEW CLIENT
# =====================================================================
class TVClient:
    def __init__(self):
        self.tv_user = os.environ.get("TV_USERNAME", None)
        self.tv_pass = os.environ.get("TV_PASSWORD", None)
        self.connect()

    def connect(self):
        try:
            time.sleep(1.0)
            if self.tv_user and self.tv_pass:
                self.tv = TvDatafeed(username=self.tv_user, password=self.tv_pass)
            else:
                self.tv = TvDatafeed()
        except Exception as e:
            print(f"  ⚠️ Reconnect warning: {e}")
            self.tv = None

    def get_hist_safe(self, symbol, exchange='NSE', interval=Interval.in_daily, n_bars=1250, retries=3):
        for _ in range(retries):
            try:
                time.sleep(1.2)  # Per-security rate-limit throttle
                df = self.tv.get_hist(symbol=symbol, exchange=exchange, interval=interval, n_bars=n_bars)
                if df is not None and not df.empty:
                    df.index = df.index.normalize()
                    return df
            except Exception:
                pass
            time.sleep(4.0)  # Cooldown before retry to clear socket throttling
            self.connect()
        return None

# =====================================================================
# 2. POINT & FIGURE ENGINE (2% BOX SIZE, 3-BOX REVERSAL)
# =====================================================================
def calculate_historical_pnf_details(prices, box_pct=BOX_PCT, reversal=REVERSAL):
    states = pd.Series(index=prices.index, dtype=float)
    extremes = pd.Series(index=prices.index, dtype=float)
    valid_prices = prices.dropna()
    if len(valid_prices) == 0:
        return states, extremes

    trend = 1
    extreme_price = valid_prices.iloc[0]

    for date, price in prices.items():
        if pd.isna(price):
            states[date] = states.shift(1).loc[date] if date != prices.index[0] else np.nan
            extremes[date] = extremes.shift(1).loc[date] if date != prices.index[0] else np.nan
            continue

        if trend == 1:
            if price > extreme_price:
                extreme_price = price
            elif price <= extreme_price * (1 - (box_pct * reversal)):
                trend = 0
                extreme_price = price
        else:
            if price < extreme_price:
                extreme_price = price
            elif price >= extreme_price * (1 + (box_pct * reversal)):
                trend = 1
                extreme_price = price

        states[date] = trend
        extremes[date] = extreme_price

    return states, extremes

# =====================================================================
# 3. DAILY TRANSITIONS BUILDER
# =====================================================================
def build_daily_transitions(pnf_states, pnf_extremes, data, final_dates):
    state_changes = pnf_states.diff()
    daily_returns = data.pct_change() * 100
    shifted_extremes = pnf_extremes.shift(1)
    
    transition_records = []
    for date in final_dates:
        day_data = state_changes.loc[date]

        def format_tickers(ticker_list, flip_dir):
            if not ticker_list:
                return "-"
            items = []
            for t in ticker_list:
                day_ret = daily_returns.loc[date, t]
                if pd.isna(day_ret):
                    items.append(f"{t} (N/A)")
                    continue
                day_sign = "+" if day_ret > 0 else ""
                
                if flip_dir == 0:
                    items.append(f"{t} ({day_sign}{day_ret:.2f}%)")
                else:
                    prev_ext = shifted_extremes.loc[date, t]
                    curr_price = data.loc[date, t]
                    if pd.isna(prev_ext) or prev_ext == 0:
                        items.append(f"{t} ({day_sign}{day_ret:.2f}%)")
                    else:
                        cum_ret = ((curr_price - prev_ext) / prev_ext) * 100
                        cum_sign = "+" if cum_ret > 0 else ""
                        label = "Peak" if flip_dir == -1 else "Trough"
                        items.append(f"{t} (Day: {day_sign}{day_ret:.2f}% | {label}: {prev_ext:.2f}, Total Move: {cum_sign}{cum_ret:.2f}%)")
            return " || ".join(items)

        flipped_up = day_data[day_data == 1].index.tolist()
        flipped_down = day_data[day_data == -1].index.tolist()
        stayed_same = day_data[day_data == 0].index.tolist()

        transition_records.append({
            "Date": date.strftime('%Y-%m-%d'),
            "Flipped to Buy (Green)": format_tickers(flipped_up, 1),
            "Flipped to Sell (Red)": format_tickers(flipped_down, -1),
            "Stayed Same": format_tickers(stayed_same, 0)
        })

    return pd.DataFrame(transition_records)

# =====================================================================
# 4. STRATEGY BACKTEST ENGINE
# =====================================================================
def run_sector_backtest(df, oversold=OVERSOLD_ENTRY, overbought=OVERBOUGHT_EXIT):
    in_position = False
    entry_price, entry_date = 0, None
    trades = []

    for i in range(1, len(df)):
        date = df.index[i]
        today_bpi = df['Smoothed_BPI_Pct'].iloc[i]
        yest_bpi = df['Smoothed_BPI_Pct'].iloc[i-1]
        close_price = df['Index_Close'].iloc[i]
        sma_200 = df['SMA_200'].iloc[i]

        if pd.isna(sma_200):
            continue

        if not in_position:
            if yest_bpi < oversold and today_bpi >= oversold and close_price > sma_200:
                in_position = True
                entry_price = close_price
                entry_date = date
        elif in_position:
            if today_bpi >= overbought or close_price < sma_200:
                profit_pct = ((close_price - entry_price) / entry_price) * 100
                trades.append({
                    "Entry Date": entry_date.strftime('%Y-%m-%d'),
                    "Entry Price": round(entry_price, 2),
                    "Exit Date": date.strftime('%Y-%m-%d'),
                    "Exit Price": round(close_price, 2),
                    "Return (%)": round(profit_pct, 2)
                })
                in_position = False

    return pd.DataFrame(trades)

# =====================================================================
# 5. CSV LOADER HELPER
# =====================================================================
def load_symbols_from_csv(csv_filename):
    if not os.path.exists(csv_filename):
        return []
    df = pd.read_csv(csv_filename)
    col = [c for c in df.columns if 'symbol' in c.lower()][0]
    raw_tickers = df[col].dropna().unique().tolist()
    clean = []
    for t in raw_tickers:
        s = str(t).strip().replace('-', '_')
        if s == 'GVT&D':
            s = 'GVT_D'
        elif s == 'GMRP&UI':
            s = 'GMRPOWER'
        clean.append(s)
    return clean

# =====================================================================
# 6. MASTER UNIVERSE (6 BROAD INDICES + 13 SECTORAL INDICES)
# =====================================================================
SECTOR_CONFIG = {
    # --- BROAD MARKET INDICES ---
    "Nifty 50": {
        "category": "Broad",
        "tv_symbols": ["NIFTY", "NIFTY_50", "CNXNIFTY"],
        "sheet_label": "Nifty50",
        "csv_file": "ind_nifty50list (1).csv",
        "fallback_tickers": []
    },
    "Nifty 100": {
        "category": "Broad",
        "tv_symbols": ["CNX100", "NIFTY100", "NIFTY_100"],
        "sheet_label": "Nifty100",
        "csv_file": "ind_nifty100list.csv",
        "fallback_tickers": []
    },
    "Nifty Midcap 150": {
        "category": "Broad",
        "tv_symbols": ["NIFTYMIDCAP150", "NIFTY_MIDCAP_150", "CNXMIDCAP"],
        "sheet_label": "Midcap150",
        "csv_file": "ind_niftymidcap150list.csv",
        "fallback_tickers": []
    },
    "Nifty Smallcap 250": {
        "category": "Broad",
        "tv_symbols": ["NIFTYSMALLCAP250", "NIFTY_SMALLCAP_250", "CNXSMALLCAP"],
        "sheet_label": "Smallcap250",
        "csv_file": "ind_niftysmallcap250list.csv",
        "fallback_tickers": []
    },
    "Nifty 500": {
        "category": "Broad",
        "tv_symbols": ["CNX500", "NIFTY500", "NIFTY_500"],
        "sheet_label": "Nifty500",
        "csv_file": "ind_nifty500list.csv",
        "fallback_tickers": []
    },
    "Nifty Total Market": {
        "category": "Broad",
        "tv_symbols": ["NIFTYTOTALMARKET", "NIFTY_TOTAL_MARKET"],
        "sheet_label": "TotalMarket",
        "csv_file": "ind_niftytotalmarket_list.csv",
        "fallback_tickers": []
    },

    # --- 13 SECTORAL INDICES ---
    "Nifty Auto": {
        "category": "Sector",
        "tv_symbols": ["CNXAUTO", "NIFTY_AUTO"],
        "sheet_label": "Auto",
        "csv_file": "ind_niftyautolist.csv",
        "fallback_tickers": ['ASHOKLEY', 'BAJAJ_AUTO', 'BHARATFORG', 'BOSCHLTD', 'EICHERMOT', 'EXIDEIND', 'HEROMOTOCO', 'M&M', 'MARUTI', 'MOTHERSON', 'SONACOMS', 'TVSMOTOR', 'TMPV', 'TIINDIA', 'UNOMINDA']
    },
    "Nifty Bank": {
        "category": "Sector",
        "tv_symbols": ["BANKNIFTY", "NIFTY_BANK"],
        "sheet_label": "Bank",
        "csv_file": "ind_niftybanklist.csv",
        "fallback_tickers": ['AUBANK', 'AXISBANK', 'BANKBARODA', 'CANBK', 'FEDERALBNK', 'HDFCBANK', 'ICICIBANK', 'IDFCFIRSTB', 'INDUSINDBK', 'KOTAKBANK', 'PNB', 'SBIN', 'UNIONBANK', 'YESBANK']
    },
    "Nifty Capital Goods": {
        "category": "Sector",
        "tv_symbols": ["NIFTY_CAPITAL_GOODS", "CNX_CAPITAL_GOODS", "CNXCAPITALGOODS"],
        "sheet_label": "CapGoods",
        "csv_file": "ind_niftyCapitalGoods_list.csv",
        "fallback_tickers": ['ABB', 'AIAENG', 'APLAPOLLO', 'APARINDS', 'ASHOKLEY', 'ASTRAMICRO', 'ASTRAL', 'BEML', 'BDL', 'BEL', 'BHEL', 'CGPOWER', 'CARBORUNIV', 'COCHINSHIP', 'CUMMINSIND', 'DATAPATTNS', 'ELGIEQUIP', 'ESCORTS', 'FINCABLES', 'FINPIPE', 'GVT_D', 'GRSE', 'HBLENGINE', 'HAL', 'POWERINDIA', 'HONAUT', 'INOXWIND', 'JYOTICNC', 'KEI', 'KAYNES', 'KIRLOSENG', 'MAZDOCK', 'PTCIL', 'POLYCAB', 'PRAJIND', 'PREMIERENE', 'SIEMENS', 'SUPREMEIND', 'SUZLON', 'TDPOWERSYS', 'THERMAX', 'TIMKEN', 'TITAGARH', 'TRITURBINE', 'USHAMART', 'VOLTAMP', 'WAAREEENER', 'WELCORP']
    },
    "Nifty Chemicals": {
        "category": "Sector",
        "tv_symbols": ["CNXCOMMODITIES", "NIFTYCOMMODITIES"],
        "sheet_label": "Chemicals",
        "csv_file": "ind_niftyChemicals_list (1).csv",
        "fallback_tickers": ['AARTIIND', 'ATUL', 'BAYERCROP', 'CHAMBLFERT', 'COROMANDEL', 'DEEPAKNTR', 'EIDPARRY', 'FLUOROCHEM', 'GNFC', 'HSCL', 'LINDEINDIA', 'NAVINFLUOR', 'PCBL', 'PIIND', 'PIDILITIND', 'SRF', 'SOLARINDS', 'SUMICHEM', 'TATACHEM', 'UPL']
    },
    "Nifty FMCG": {
        "category": "Sector",
        "tv_symbols": ["CNXFMCG", "NIFTY_FMCG"],
        "sheet_label": "FMCG",
        "csv_file": "ind_niftyfmcglist.csv",
        "fallback_tickers": ['BRITANNIA', 'COLPAL', 'DABUR', 'EMAMILTD', 'GODREJCP', 'HINDUNILVR', 'ITC', 'MARICO', 'NESTLEIND', 'PATANJALI', 'RADICO', 'TATACONSUM', 'UBL', 'UNITDSPR', 'VBL']
    },
    "Nifty IT": {
        "category": "Sector",
        "tv_symbols": ["CNXIT", "NIFTY_IT"],
        "sheet_label": "IT",
        "csv_file": "ind_niftyitlist.csv",
        "fallback_tickers": ['COFORGE', 'HCLTECH', 'INFY', 'LTIM', 'MPHASIS', 'OFSS', 'PERSISTENT', 'TCS', 'TECHM', 'WIPRO']
    },
    "Nifty Metal": {
        "category": "Sector",
        "tv_symbols": ["CNXMETAL", "NIFTY_METAL"],
        "sheet_label": "Metal",
        "csv_file": "ind_niftymetallist.csv",
        "fallback_tickers": ['APLAPOLLO', 'ADANIENT', 'HINDALCO', 'HINDCOPPER', 'HINDZINC', 'JSWSTEEL', 'JSL', 'JINDALSTEL', 'LLOYDSME', 'NMDC', 'NATIONALUM', 'SAIL', 'TATASTEEL', 'VEDL', 'WELCORP']
    },
    "Nifty Pharma": {
        "category": "Sector",
        "tv_symbols": ["CNXPHARMA", "NIFTY_PHARMA"],
        "sheet_label": "Pharma",
        "csv_file": "ind_niftypharmalist.csv",
        "fallback_tickers": ['ABBOTINDIA', 'AJANTPHARM', 'ALKEM', 'AUROPHARMA', 'BIOCON', 'CIPLA', 'DIVISLAB', 'DRREDDY', 'GLAND', 'GLENMARK', 'IPCALAB', 'LAURUSLABS', 'LUPIN', 'MANKIND', 'PPLPHARMA', 'SAILIFE', 'SUNPHARMA', 'TORNTPHARM', 'WOCKPHARMA', 'ZYDUSLIFE']
    },
    "Nifty Power": {
        "category": "Sector",
        "tv_symbols": ["CNXENERGY", "NIFTYENERGY"],
        "sheet_label": "Power",
        "csv_file": "ind_niftyPower_list.csv",
        "fallback_tickers": ['ACMESOLAR', 'ADANIENSOL', 'ADANIGREEN', 'ADANIPOWER', 'CESC', 'JSWENERGY', 'JPPOWER', 'KPIGREEN', 'NHPC', 'NLCINDIA', 'NTPCGREEN', 'NTPC', 'NAVA', 'PTC', 'POWERGRID', 'RTNPOWER', 'RPOWER', 'SJVN', 'TATAPOWER', 'TORNTPOWER']
    },
    "Nifty Private Bank": {
        "category": "Sector",
        "tv_symbols": ["NIFTY_PVT_BANK", "NIFTYPVTBANK"],
        "sheet_label": "PvtBank",
        "csv_file": "ind_nifty_privatebanklist.csv",
        "fallback_tickers": ['AXISBANK', 'BANDHANBNK', 'FEDERALBNK', 'HDFCBANK', 'ICICIBANK', 'IDFCFIRSTB', 'INDUSINDBK', 'KOTAKBANK', 'RBLBANK', 'YESBANK']
    },
    "Nifty PSU Bank": {
        "category": "Sector",
        "tv_symbols": ["CNXPSUBANK", "NIFTY_PSU_BANK"],
        "sheet_label": "PSUBank",
        "csv_file": "ind_niftypsubanklist.csv",
        "fallback_tickers": ['BANKBARODA', 'BANKINDIA', 'MAHABANK', 'CANBK', 'CENTRALBK', 'INDIANB', 'IOB', 'PSB', 'PNB', 'SBIN', 'UCOBANK', 'UNIONBANK']
    },
    "Nifty Realty": {
        "category": "Sector",
        "tv_symbols": ["CNXREALTY", "NIFTY_REALTY"],
        "sheet_label": "Realty",
        "csv_file": "ind_niftyrealtylist.csv",
        "fallback_tickers": ['ABREL', 'ANANTRAJ', 'BRIGADE', 'DLF', 'GODREJPROP', 'LODHA', 'OBEROIRLTY', 'PHOENIXLTD', 'PRESTIGE', 'SOBHA']
    },
    "Nifty Oil and Gas": {
        "category": "Sector",
        "tv_symbols": ["NIFTY_OIL_AND_GAS", "NIFTYOILGAS", "CNXOILGAS"],
        "sheet_label": "OilGas",
        "csv_file": "ind_niftyoilgaslist.csv",
        "fallback_tickers": ['ATGL', 'AEGISLOG', 'BPCL', 'CASTROLIND', 'CHENNPETRO', 'GAIL', 'HINDPETRO', 'IOC', 'IGL', 'MGL', 'ONGC', 'OIL', 'PETRONET', 'RELIANCE']
    }
}

# =====================================================================
# 7. EXECUTION PIPELINE
# =====================================================================
print("Connecting to TradingView...")
client = TVClient()

wb = Workbook()
ws_dash = wb.active
ws_dash.title = "Sector Dashboard"

dashboard_records = []
analysis_matrix_records = {}  # Sector Name -> Series of daily smoothed BPI
regime_dict = {}

GLOBAL_PRICE_CACHE = {}
GLOBAL_PNF_CACHE = {}

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
border_thin = Border(left=Side(style='thin', color='D9D9D9'),
                     right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'),
                     bottom=Side(style='thin', color='D9D9D9'))
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")

def format_excel_sheet(ws, is_transition=False):
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border_thin

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border_thin
            cell.alignment = align_left if is_transition and cell.column > 1 else align_center
            if cell.row % 2 == 0:
                cell.fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    ws.freeze_panes = "A2"

for sector_name, meta in SECTOR_CONFIG.items():
    symbols_to_try = meta["tv_symbols"]
    sheet_label = meta["sheet_label"]
    csv_file = meta["csv_file"]
    
    tickers = load_symbols_from_csv(csv_file)
    if not tickers:
        tickers = meta["fallback_tickers"]

    print(f"\n{'='*60}")
    print(f"🚀 Processing: {sector_name} ({len(tickers)} stocks) [Stabilized 2% Box / 3-Day EMA]")
    print(f"{'='*60}")

    sector_price_dict = {}
    sector_pnf_states = {}
    sector_pnf_extremes = {}

    fetch_counter = 0
    for idx, ticker in enumerate(tickers, start=1):
        if ticker in GLOBAL_PRICE_CACHE:
            sector_price_dict[ticker] = GLOBAL_PRICE_CACHE[ticker]
            s, e = GLOBAL_PNF_CACHE[ticker]
            sector_pnf_states[ticker] = s
            sector_pnf_extremes[ticker] = e
        else:
            df = client.get_hist_safe(symbol=ticker, exchange='NSE', interval=Interval.in_daily, n_bars=1250)
            if df is not None and not df.empty:
                series = df['close']
                GLOBAL_PRICE_CACHE[ticker] = series
                sector_price_dict[ticker] = series

                s, e = calculate_historical_pnf_details(series, box_pct=BOX_PCT, reversal=REVERSAL)
                GLOBAL_PNF_CACHE[ticker] = (s, e)
                sector_pnf_states[ticker] = s
                sector_pnf_extremes[ticker] = e

                print(f"  [{idx:03d}/{len(tickers):03d}] {ticker:<12} -> OK")
                fetch_counter += 1
            else:
                print(f"  [{idx:03d}/{len(tickers):03d}] {ticker:<12} -> ⚠️ Unavailable")

        if fetch_counter > 0 and fetch_counter % 12 == 0:
            print("  ⏳ Mini-batch pause (5s)...")
            time.sleep(5.0)
            fetch_counter += 1

    data = pd.DataFrame(sector_price_dict)
    if data.empty:
        print(f"❌ No constituent data retrieved for {sector_name}. Skipping.")
        continue

    pnf_states = pd.DataFrame(sector_pnf_states, index=data.index)
    pnf_extremes = pd.DataFrame(sector_pnf_extremes, index=data.index)

    print("  🔄 Checking Index ticker connection...")
    client.connect()
    time.sleep(2.0)

    index_close = None
    for sym in symbols_to_try:
        idf = client.get_hist_safe(symbol=sym, exchange='NSE', interval=Interval.in_daily, n_bars=1250)
        if idf is not None and not idf.empty:
            index_close = idf['close'].squeeze()
            print(f"  ✅ Connected to Index ticker: {sym}")
            break

    if index_close is None:
        print(f"  ℹ️ Official index unavailable. Synthesizing equal-weight sector composite.")
        index_close = data.mean(axis=1)

    daily_bpi_raw = (pnf_states.mean(axis=1) * 100).round(2)
    daily_bpi_smoothed = daily_bpi_raw.ewm(span=EMA_SPAN, adjust=False).mean().round(2)
    sma_200 = index_close.rolling(window=SMA_PERIOD).mean().round(2)

    final_df = pd.DataFrame({
        "Index_Close": index_close.round(2),
        "SMA_200": sma_200,
        "Raw_BPI_Pct": daily_bpi_raw,
        "Smoothed_BPI_Pct": daily_bpi_smoothed
    }).dropna()

    if final_df.empty:
        continue

    # Store for Analysis.xlsx matrix
    analysis_matrix_records[sector_name] = final_df['Smoothed_BPI_Pct']

    trades_df = run_sector_backtest(final_df, oversold=OVERSOLD_ENTRY, overbought=OVERBOUGHT_EXIT)

    latest_close = final_df['Index_Close'].iloc[-1]
    latest_sma = final_df['SMA_200'].iloc[-1]
    latest_smoothed_bpi = final_df['Smoothed_BPI_Pct'].iloc[-1]
    latest_raw_bpi = final_df['Raw_BPI_Pct'].iloc[-1]
    regime = "Bullish (>200 SMA)" if latest_close > latest_sma else "Bearish (<200 SMA)"
    regime_dict[sector_name] = regime

    if not trades_df.empty:
        tot_trades = len(trades_df)
        wins = len(trades_df[trades_df['Return (%)'] > 0])
        win_rate = f"{(wins / tot_trades) * 100:.1f}%"
        cum_ret = f"{trades_df['Return (%)'].sum():.1f}%"
        avg_ret = f"{trades_df['Return (%)'].mean():.1f}%"
    else:
        tot_trades, win_rate, cum_ret, avg_ret = 0, "0.0%", "0.0%", "0.0%"

    dashboard_records.append({
        "Sector": sector_name,
        "Current BPI (%)": latest_smoothed_bpi,
        "Raw BPI (%)": latest_raw_bpi,
        "Market Regime": regime,
        "Backtest Trades": tot_trades,
        "Win Rate": win_rate,
        "Cumulative Return": cum_ret,
        "Avg Yield": avg_ret
    })

    transitions_df = build_daily_transitions(pnf_states, pnf_extremes, data, final_df.index)

    # Sheet A: Trades
    ws_trades = wb.create_sheet(title=f"{sheet_label}_Trades")
    if not trades_df.empty:
        for r in dataframe_to_rows(trades_df, index=False, header=True):
            ws_trades.append(r)
        format_excel_sheet(ws_trades, is_transition=False)
        for row in ws_trades.iter_rows(min_row=2, max_row=ws_trades.max_row, min_col=5, max_col=5):
            for cell in row:
                val = float(cell.value)
                cell.font = Font(color="27AE60" if val > 0 else "C0392B", bold=True)
    else:
        ws_trades.append(["No historical trades triggered under current rules."])

    # Sheet B: Momentum
    ws_mom = wb.create_sheet(title=f"{sheet_label}_Momentum")
    mom_df_out = final_df.reset_index()
    mom_df_out.columns = ["Date", "Index_Close", "SMA_200", "Raw_BPI_Pct", "Breadth_Momentum_Pct"]
    mom_df_out["Date"] = mom_df_out["Date"].dt.strftime('%Y-%m-%d')
    for r in dataframe_to_rows(mom_df_out, index=False, header=True):
        ws_mom.append(r)
    format_excel_sheet(ws_mom, is_transition=False)

    # Sheet C: Transitions
    ws_trans = wb.create_sheet(title=f"{sheet_label}_Transitions")
    for r in dataframe_to_rows(transitions_df, index=False, header=True):
        ws_trans.append(r)
    format_excel_sheet(ws_trans, is_transition=True)

    time.sleep(5.0)

# =====================================================================
# 8. WRITE MASTER WORKBOOK (Nifty_Multi_Sector_Momentum_Report.xlsx)
# =====================================================================
dash_df = pd.DataFrame(dashboard_records)
if not dash_df.empty:
    dash_df = dash_df.sort_values(by="Current BPI (%)", ascending=True)
    for r in dataframe_to_rows(dash_df, index=False, header=True):
        ws_dash.append(r)

    for cell in ws_dash[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border_thin

    for row in ws_dash.iter_rows(min_row=2, max_row=ws_dash.max_row, min_col=1, max_col=ws_dash.max_column):
        for cell in row:
            cell.alignment = align_center
            cell.border = border_thin
            if cell.row % 2 == 0:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

for sheet in wb.worksheets:
    for col in sheet.columns:
        col_letter = col[0].column_letter
        if "Transitions" in sheet.title:
            sheet.column_dimensions[col_letter].width = 14 if col_letter == 'A' else 50
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            sheet.column_dimensions[col_letter].width = min(max_len + 3, 32)

master_output_filename = "Nifty_Multi_Sector_Momentum_Report.xlsx"
wb.save(master_output_filename)
print(f"\n✅ Master report saved: {master_output_filename}")

# =====================================================================
# 9. GENERATE EXACT ANALYSIS.XLSX MATRIX (BROAD TOP, SECTORS BOTTOM)
# =====================================================================
print("Generating Analysis.xlsx day-on-day breadth matrix...")
matrix_df = pd.DataFrame(analysis_matrix_records)
matrix_df.index = matrix_df.index.strftime('%Y-%m-%d')
matrix_df = matrix_df.sort_index(ascending=False)  # Recent dates first (left to right)
matrix_transposed = matrix_df.T

# Separate Broad indices vs Sector indices
broad_order = ["Nifty 50", "Nifty 100", "Nifty Midcap 150", "Nifty Smallcap 250", "Nifty 500", "Nifty Total Market"]
sector_order = [s for s in SECTOR_CONFIG.keys() if s not in broad_order]

rows_broad = []
for name in broad_order:
    if name in matrix_transposed.index:
        regime = regime_dict.get(name, "N/A")
        vals = matrix_transposed.loc[name].tolist()
        rows_broad.append([name, regime] + vals)

rows_sector = []
for name in sector_order:
    if name in matrix_transposed.index:
        regime = regime_dict.get(name, "N/A")
        vals = matrix_transposed.loc[name].tolist()
        rows_sector.append([name, regime] + vals)

date_cols = list(matrix_transposed.columns)
headers = ["Sector", "Market Regime"] + date_cols

wb_analysis = Workbook()
ws_an = wb_analysis.active
ws_an.title = "Sheet1"

ws_an.append(headers)
for r in rows_broad:
    ws_an.append(r)

# Insert empty separator row between Broad and Sectoral indices
ws_an.append([None] * len(headers))

for r in rows_sector:
    ws_an.append(r)

# Format Analysis.xlsx
for cell in ws_an[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = align_center
    cell.border = border_thin

for row in ws_an.iter_rows(min_row=2, max_row=ws_an.max_row, min_col=1, max_col=ws_an.max_column):
    for cell in row:
        cell.border = border_thin
        if cell.column == 1:
            cell.alignment = align_left
            cell.font = Font(bold=True)
        elif cell.column == 2:
            cell.alignment = align_center
            if cell.value and "Bullish" in str(cell.value):
                cell.font = Font(color="27AE60", bold=True)
            elif cell.value and "Bearish" in str(cell.value):
                cell.font = Font(color="C0392B", bold=True)
        else:
            cell.alignment = align_center

for col in ws_an.columns:
    col_letter = col[0].column_letter
    if col_letter in ['A', 'B']:
        ws_an.column_dimensions[col_letter].width = 24
    else:
        ws_an.column_dimensions[col_letter].width = 14

analysis_filename = "Analysis.xlsx"
wb_analysis.save(analysis_filename)
print(f"✅ Day-on-day breadth file saved: {analysis_filename}")

# =====================================================================
# 10. DISPATCH EMAIL WITH ATTACHED ANALYSIS.XLSX
# =====================================================================
def send_email_report(attachment_path):
    sender = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    recipient = os.environ.get("RECIPIENT_EMAIL")

    if not sender or not password or not recipient:
        print("ℹ️ Skipping email dispatch: SENDER_EMAIL, SENDER_APP_PASSWORD, or RECIPIENT_EMAIL not set.")
        return

    latest_date = date_cols[0] if date_cols else "Today"
    msg = EmailMessage()
    msg['Subject'] = f"Daily Multi-Sector & Broad Market Breadth Report - {latest_date}"
    msg['From'] = sender
    msg['To'] = recipient

    body = f"""Hi,

Attached is today's daily Market Breadth & Momentum Report ({latest_date}).

Key Highlights:
- Broad Market Universe: Nifty 50, 100, 500, Midcap 150, Smallcap 250, Total Market.
- Sectoral Universe: 13 NSE Sectoral Baskets.
- Stabilization Model: 2.0% Box / 3-Box Reversal with 3-Day EMA Breadth Smoothing.

The full 40-tab master workbook (Nifty_Multi_Sector_Momentum_Report.xlsx) with individual trade logs and daily transitions has been saved to your repository/artifacts.

Regards,
Automated Market Breadth Engine
"""
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, password)
            server.send_message(msg)
        print(f"📧 Successfully dispatched {file_name} to {recipient}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

send_email_report(analysis_filename)
print("\n🏁 Daily run completed successfully.")
