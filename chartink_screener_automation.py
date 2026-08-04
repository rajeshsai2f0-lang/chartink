"""
╔══════════════════════════════════════════════════════════════╗
║   CHARTINK SCREENER AUTOMATION                               ║
║   Runs multiple screeners → saves to Excel                   ║
╠══════════════════════════════════════════════════════════════╣
║  INSTALL ONCE (run in Command Prompt):                       ║
║   pip install requests beautifulsoup4 openpyxl pandas        ║
╚══════════════════════════════════════════════════════════════╝
"""

import sys, os, time, datetime

missing = []
for pkg in ['requests', 'bs4', 'openpyxl', 'pandas']:
    try: __import__(pkg)
    except ImportError: missing.append(pkg)
if missing:
    print(f"\n❌  Missing: {', '.join(missing)}")
    print(f"    Run:  pip install {' '.join(missing)}")
    if not os.environ.get("GITHUB_ACTIONS"):
        input("\nPress Enter to exit...")
    sys.exit(1)

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
#  ✏️  YOUR SCREENERS
#  Format: ("Display Name", "clause", "your chartink scan clause here")
# ══════════════════════════════════════════════════════════════════════════════
SCREENERS = [
    (
        "IPO Screener",
        "clause",
        "( {cash} not( 18 months ago close > 0 ) )"
    ),
    (
        "Combined Winners",
        "clause",
        "( {cash} ( ( {cash} ( ( {cash} (  daily close /  22 days ago close >  1.2 and  market cap >  1 and  daily close *  daily sma(  daily volume , 20 ) >  30000000 and  daily close >  daily sma(  daily close , 200 ) ) ) or( {cash} (  daily close /  66 days ago close >=  1.3 and  market cap >  0 and  daily close >=  1 and  daily close *  daily sma(  daily volume , 20 ) >  30000000 and  daily close >  daily sma(  daily close , 200 ) ) ) ) ) or( {cash} (  market cap >=  1000 and  daily close >  1 day ago max( 252 ,  daily high ) *  0.75 and  daily close >  daily sma(  daily close , 50 ) and  daily close >  daily sma(  daily close , 200 ) and  daily close *  daily sma(  daily volume , 20 ) >  30000000 ) ) ) )"
    ),
    (
        "25 Percent 52 Week",
        "clause",
        "( {cash} (  daily close >=  30 and  daily close >=  daily ema(  daily close , 50 ) and  market cap >=  1000 ) )"
    ),
    
    (
        "CopyATR",
        "clause",
        "( {cash} (  daily avg true range( 14 ) <  10 days ago avg true range( 14 ) and  daily avg true range( 14 ) /  daily close <  0.08 and  daily close >  (  weekly max( 52 ,  weekly close ) *  0.75 ) and  daily ema(  daily close , 50 ) >  daily ema(  daily close , 150 ) and  daily ema(  daily close , 150 ) >  daily ema(  daily close , 200 ) and  daily close >  daily ema(  daily close , 50 ) and  daily close >  10 and  daily close *  daily volume >  1000000 ) )"
    ),
    # ── NEW SCREENERS ────────────────────────────────────────────────────────
    
    (
        "Volume Shockers",
        "clause",
        "( {57960} (  daily volume >  daily sma( volume,10 ) *  2 and( {cash} (  daily close >  1 day ago close *  1.05 or  daily close <  1 day ago close *  0.95 ) ) ) )"
    ),
    
    (
        "RSI Strong",
        "clause",
        "( {cash} (  daily volume >  daily sma(  daily volume , 20 ) and  daily rsi( 14 ) >  60 and  weekly rsi( 14 ) >  60 and  monthly rsi( 14 ) >  60 and  daily close >  300 and  market cap >  1000 ) )"
    ),
    
    (
        "BIG Breakout Scan",
        "clause",
        "( {cash} ( ( {cash} ( ( {cash} ( ( {cash} ( ( {cash} (  daily close /  50 days ago open <  1.58 and  daily close /  60 days ago open <  1.65 and  daily close /  40 days ago open <  1.4 and  daily close /  30 days ago open <  1.4 and  daily high /  20 days ago open <  1.4 and  daily high /  21 days ago low <  1.4 and  daily high /  22 days ago low <  1.4 and  daily high /  23 days ago low <  1.4 and  daily high /  24 days ago low <  1.4 and  daily close /  25 days ago low <  1.4 and  daily high /  26 days ago low <  1.4 and  daily high /  27 days ago low <  1.4 and  daily high /  28 days ago low <  1.4 and  daily high /  29 days ago low <  1.4 and  daily high /  30 days ago low <  1.4 and  daily close /  20 days ago open <  1.4 and  daily close /  10 days ago open <  1.4 and  daily close /  10 days ago close <  1.35 and  daily high /  13 days ago open <  40 and  daily high /  14 days ago open <  40 and  daily high /  15 days ago open <  40 and  daily high /  16 days ago open <  40 and  daily close /  4 days ago low <  1.21 and  daily high /  4 days ago open <  1.21 and  daily high /  4 days ago close <  1.21 and  daily high /  5 days ago open <  1.23 and  daily high /  5 days ago close <  1.23 and  daily high /  2 days ago close <  1.25 and  daily high /  1 day ago close <  1.20 and  daily high /  1 day ago low <  1.20 and  daily close >  1 day ago high *  0.997 and  daily close <  7000 and( {cash} (  daily volume >=  daily sma( volume,20 ) *  0.95 and( {cash} (  daily volume >=  daily sma( volume,20 ) *  2.8 or  1 day ago volume >=  1 day ago sma( volume,20 ) *  0.8 or  2 days ago volume >=  2 days ago sma( volume,20 ) *  0.8 or  3 days ago volume >=  3 days ago sma( volume,20 ) *  0.8 or  4 days ago volume >=  4 days ago sma( volume,20 ) *  1 ) ) ) ) and  daily close >  daily ema(  daily close , 10 ) and  daily close >  weekly ema(  weekly close , 40 ) and  daily close >  daily ema(  daily close , 450 ) and  daily close >  daily sma(  daily close , 50 ) and  daily high <  daily sma(  daily close , 50 ) *  1.27 and  market cap >  400 and  market cap <  300000 and( {cash} (  1 day ago high <  1 day ago sma(  daily close , 50 ) *  1.15 or  2 days ago high <  2 days ago sma(  daily close , 50 ) *  1.15 or  3 days ago high <  3 days ago sma(  daily close , 50 ) *  1.15 or  4 days ago high <  4 days ago sma(  daily close , 50 ) *  1.15 or  5 days ago high <  5 days ago sma(  daily close , 50 ) *  1.15 or  6 days ago high <  6 days ago sma(  daily close , 50 ) *  1.15 or  7 days ago high <  7 days ago sma(  daily close , 50 ) *  1.15 or  8 days ago high <  8 days ago sma(  daily close , 50 ) *  1.15 or  9 days ago high <  9 days ago sma(  daily close , 50 ) *  1.15 or  10 days ago high <  10 days ago sma(  daily close , 50 ) *  1.15 or  11 days ago high <  11 days ago sma(  daily close , 50 ) *  1.15 or  12 days ago high <  12 days ago sma(  daily close , 50 ) *  1.15 or  13 days ago high <  13 days ago sma(  daily close , 50 ) *  1.15 or  14 days ago high <  14 days ago sma(  daily close , 50 ) *  1.15 or  15 days ago high <  14 days ago sma(  daily close , 50 ) *  1.15 or  16 days ago high <  14 days ago sma(  daily close , 50 ) *  1.15 ) ) ) ) and( {cash} (  daily close >=  (  daily high -  daily open *  0.52 ) +  daily open or  daily close >=  (  daily high -  daily low *  0.52 ) +  daily low ) ) and  daily \"close - 1 candle ago close / 1 candle ago close * 100\" >=  1 and  weekly ema(  weekly close , 10 ) >  weekly ema(  weekly close , 40 ) and  daily \"close - 1 candle ago close / 1 candle ago close * 100\" <  12 and  1 day ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  4.5 and  1 day ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -2 and  2 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  7 and  2 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -4 and  3 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  7 and  3 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -7 and  4 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  10 and  5 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  10 and  6 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  10 and  7 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" <  12 and  1 day ago low >=  2 days ago low *  .95 and  1 day ago close >=  2 days ago low *  .968 and  daily high /  daily low <  1.15 and  2 days ago high /  2 days ago low <  1.1 and  1 day ago high /  1 day ago low <  1.1 and( {cash} (  daily low <  daily ema(  daily close , 10 ) or  1 day ago low <  1 day ago ema(  daily close , 10 ) or  2 days ago low <  2 days ago ema(  daily close , 10 ) or  3 days ago low <  3 days ago ema(  daily close , 10 ) or  4 days ago low <  4 days ago ema(  daily close , 10 ) or  5 days ago low <  5 days ago ema(  daily close , 10 ) or  6 days ago low <  6 days ago ema(  daily close , 10 ) or  6 days ago low <  6 days ago sma(  daily close , 50 ) or  7 days ago low <  7 days ago sma(  daily close , 50 ) or  7 days ago low <  7 days ago ema(  daily close , 10 ) ) ) and( {cash} not(  2 days ago close <  2 days ago high *  .96 and  daily close <  2 days ago high ) ) ) ) and( {cash} not(  1 day ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  2.5 and  2 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  2.5 and  1 day ago high >  2 days ago high and  2 days ago low <  1 day ago low ) ) and( {cash} not(  3 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -5 and  daily close <  3 days ago high ) ) ) ) and( {cash} ( ( {cash} (  15 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  14 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  13 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  12 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  11 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  10 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  9 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  8 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  7 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  6 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -10 and  5 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -7 and  2 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -7 and  4 days ago \"close - 1 candle ago close / 1 candle ago close * 100\" >  -7 ) ) or( {cash} (  daily high >  1 day ago max( 10 ,  daily high ) *  0.96 or  daily high >  1 day ago max( 15 ,  daily high ) *  0.96 ) ) ) ) and( {cash} not(  daily high /  daily low >  1.07 and  daily close >  daily sma(  daily close , 50 ) and  1 day ago  close <=  1 day ago  sma(  daily close , 50 ) ) ) and  daily close *  daily volume >  10000000 and  1 day ago close <  1 day ago ema(  daily close , 10 ) *  1.16 and( {cash} not(  1 day ago close >  1 day ago ema(  daily close , 10 ) and  2 days ago close >  2 days ago ema(  daily close , 10 ) and  3 days ago close >  3 days ago ema(  daily close , 10 ) and  4 days ago close >  4 days ago ema(  daily close , 10 ) and  5 days ago close >  5 days ago ema(  daily close , 10 ) and  6 days ago close >  6 days ago ema(  daily close , 10 ) and  7 days ago close >  7 days ago ema(  daily close , 10 ) and  8 days ago close >  8 days ago ema(  daily close , 10 ) and  9 days ago close >  9 days ago ema(  daily close , 10 ) and  10 days ago close >  10 days ago ema(  daily close , 10 ) and  11 days ago close >  11 days ago ema(  daily close , 10 ) and  12 days ago close >  12 days ago ema(  daily close , 10 ) and  13 days ago close >  13 days ago ema(  daily close , 10 ) and  14 days ago close >  14 days ago ema(  daily close , 10 ) and  15 days ago close >  15 days ago ema(  daily close , 10 ) and  16 days ago close >  16 days ago ema(  daily close , 10 ) ) ) ) ) ) )"
    ),
    
    (
        "Darvas Breakout",
        "clause",
        "( {cash} (  weekly high =  weekly max( 12 ,  weekly high ) and  daily close >  1 day ago high and  1 day ago close >  2 days ago close and  2 days ago close >=  3 days ago close and  daily volume >  daily min( 3 ,  daily volume ) and  daily volume >=  100000 and  daily close >  daily ema(  daily close , 200 ) ) )"
    ),
    (
        "20 Days 100Cr Turnover",
        "clause",
        "( {33489} (  daily close *  daily volume >=  1000000000 and  1 day ago volume *  1 day ago volume >=  1000000000 and  2 days ago close *  2 days ago volume >=  1000000000 and  3 days ago close *  3 days ago volume >=  1000000000 and  5 days ago close *  5 days ago volume >=  1000000000 and  7 days ago volume *  7 days ago close >=  1000000000 and  8 days ago close *  8 days ago volume >=  1000000000 and  9 days ago close *  9 days ago volume >=  1000000000 and  10 days ago close *  10 days ago volume >=  1000000000 and  11 days ago close *  11 days ago volume >=  1000000000 and  12 days ago close *  12 days ago volume >=  1000000000 and  13 days ago close *  13 days ago volume >=  1000000000 and  14 days ago close *  14 days ago volume >=  1000000000 and  15 days ago close *  15 days ago volume >=  1000000000 and  16 days ago close *  16 days ago volume >=  1000000000 and  18 days ago close *  18 days ago volume >=  1000000000 and  19 days ago close *  19 days ago volume >=  1000000000 and  20 days ago close *  20 days ago volume >=  1000000000 ) )"
    ),
    
    
    
    (
        "3 Week Tight",
        "clause",
        "( {cash} ( ( {cash} (  daily close >  20 and  daily sma(  daily volume , 50 ) *  daily close >=  2000000 and  daily close >  daily ema(  daily close , 50 ) and  abs(  (  weekly max( 3 ,  weekly close ) /  weekly min( 3 ,  weekly close ) -  1 ) *  100 ) <=  3 and(  3 weeks ago max( 12 ,  weekly close ) /  3 weeks ago min( 12 ,  weekly close ) -  1 ) *  100 >=  30 and  market cap <=  40000 ) ) ) )"
    ),
    (
        "Master Candle Breakout",
        "clause",
        "( {cash} ( ( {cash} (  daily high <=  5 days ago high and  1 day ago high <=  5 days ago high and  2 days ago high <=  5 days ago high and  3 days ago high <=  5 days ago high and  4 days ago high <=  5 days ago high and  5 days ago high >=  5 days ago low *  1.04 and  5 days ago close >=  5 days ago open and  daily low >=  5 days ago low and  1 day ago low >=  5 days ago low and  2 days ago low >=  5 days ago low and  3 days ago low >=  5 days ago low and  4 days ago low >=  5 days ago low ) ) and  market cap >=  1 and  daily close >=  daily ema(  daily close , 50 ) ) )"
    ),
    (
        "Multi Bagger StockExploder",
        "clause",
        "( {cash} (  monthly \"close - 1 candle ago close / 1 candle ago close * 100\" >=  20 and  monthly rsi( 14 ) >=  50 and  monthly ema(  monthly close , 10 ) >=  monthly ema(  monthly close , 20 ) and  daily ema(  daily volume , 30 ) >=  50000 and  daily close >=  20 and( {cash} (  monthly count( 20, 1 where  monthly ema(  monthly close , 10 ) >  monthly ema(  monthly close , 20 ) and  1 month ago  ema(  monthly close , 10 )<=  1 month ago  ema(  monthly close , 20 ) ) >=  1 or  monthly close >  monthly ema(  monthly close , 10 ) and  1 month ago  close <=  1 month ago  ema(  monthly close , 10 ) ) ) ) )"
    ),
    (
        "RB StockExploder",
        "clause",
        "( {cash} (  daily wma( close,1 ) >  monthly wma( close,2 ) +  1 and  monthly wma( close,2 ) >  monthly wma( close,4 ) +  2 and  daily wma( close,1 ) >  weekly wma( close,6 ) +  2 and  weekly wma( close,6 ) >  weekly wma( close,12 ) +  2 and  daily wma( close,1 ) >  4 days ago wma( close,12 ) +  2 and  daily wma( close,1 ) >  2 days ago wma( close,20 ) +  2 and  daily close >  25 and  daily close <=  500 and  weekly volume >  85000 ) )"
    ),
    (
        "ATR Tight",
        "clause",
        "( {cash} (  daily avg true range( 14 ) <  10 days ago avg true range( 14 ) and  daily avg true range( 14 ) /  daily close <  0.08 and  daily close >  (  weekly max( 52 ,  weekly close ) *  0.75 ) and  daily ema(  daily close , 50 ) >  daily ema(  daily close , 150 ) and  daily ema(  daily close , 150 ) >  daily ema(  daily close , 200 ) and  daily close >  daily ema(  daily close , 50 ) and  daily close >  10 and  daily close *  daily volume >  1000000 ) )"
    ),
    
     
    (
        "Institutional Buying",
        "clause",
        "( {cash} (  daily close >  1 day ago min( 20 ,  daily close ) *  1.2 and  daily count( 20, 1 where  daily volume /  20 days ago sma(  daily volume , 50 ) >  2 ) >  5 and( {cash} (  daily max( 20 ,  daily volume ) >=  20 days ago max( 260 ,  21 days ago volume ) or  daily max( 20 ,  daily volume ) >=  20 days ago max( 66 ,  21 days ago volume ) or  daily max( 20 ,  daily volume ) >=  20 days ago max( 22 ,  21 days ago volume ) or  daily max( 20 ,  daily \"close - 1 candle ago close / 1 candle ago close * 100\" ) >=  8 ) ) and  market cap >  500 ) )"
    ),
    (
        "52 Week Highest Close",
        "clause",
        "( {cash} (  daily close >=  50 and  daily ema( close,5 ) >  daily ema( close,26 ) and  daily ema( close,13 ) >  daily ema( close,26 ) and  daily close >  1 day ago close *  1.03 and  daily volume >  daily sma( volume,20 ) *  1.0 and  daily ema( close,5 ) >  daily ema( close,13 ) and  daily high =  daily max( 250 ,  daily high ) *  1 and  1 day ago close >  2 days ago close *  0.98 ) )"
    ),
    (
        "Weekly Pennant",
        "clause",
        "( {cash} (  daily high <  1 week ago high and  1 week ago high <  2 weeks ago high and  daily low >  1 week ago low and  1 week ago low >  2 weeks ago low and  daily open <  2 weeks ago open and  daily close >  1 week ago close and  3 weeks ago close >  4 weeks ago close and  4 weeks ago close >  5 weeks ago close ) )"
    ),
    (
        "Multi Year Breakout SK",
        "clause",
        "( {cash} (  monthly macd line( 26,12,9 ) >  0 and  monthly rsi( 14 ) >  69 and  daily close >  20 and  daily volume >  50000 and  market cap >  250 and( {cash} (  monthly macd line( 5,8,3 ) >  1 month ago max( 35 ,  monthly macd line( 5,8,3 ) ) or  monthly macd line( 13,21,8 ) >  1 month ago max( 35 ,  monthly macd line( 13,21,8 ) ) or  monthly macd line( 26,12,9 ) >  1 month ago max( 35 ,  monthly macd line( 26,12,9 ) ) ) ) ) )"
    ),

(
        "Gap And Go Scan",
        "clause",
        "( {cash} (  daily open >  1 day ago close *  1.03 and  market cap >  2000 and  daily volume *  daily close >  50000000 ) )"
    ),



]

PAUSE_BETWEEN = 5   # seconds between requests (keep ≥ 4 to avoid 419)

OUTPUT_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    f"Chartink_Screener_{datetime.date.today().strftime('%Y-%m-%d')}.xlsx"
)

PROCESS_URL = "https://chartink.com/screener/process"

# ══════════════════════════════════════════════════════════════════════════════
#  🚫  INDEX & ETF EXCLUSION LIST
#  Tickers matching these exact names OR containing these substrings are dropped
# ══════════════════════════════════════════════════════════════════════════════
_EXCLUDE_EXACT = {
    # ── NSE Indices ────────────────────────────────────────────────────────────
    "NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "NIFTYNXT50",
    "NIFTY50", "NIFTY100", "NIFTY200", "NIFTY500", "NIFTYMIDCAP50",
    "NIFTYMIDCAP100", "NIFTYMIDCAP150", "NIFTYMIDCAP400",
    "NIFTYSMALLCAP50", "NIFTYSMALLCAP100", "NIFTYSMALLCAP250",
    "NIFTYMICROCAP250", "NIFTYLARGECAP", "SENSEX", "BSE500",
    "NIFTYAUTO", "NIFTYBANK", "NIFTYFIN", "NIFTYFINSERVICE",
    "NIFTYFMCG", "NIFTYIT", "NIFTYMEDIA", "NIFTYMETAL",
    "NIFTYPHARMA", "NIFTYREALTY", "NIFTYPSUBANK", "NIFTYINFRA",
    "NIFTYCPSE", "NIFTYMHC", "NIFTYENERGY", "NIFTYDIVOPPS50",
    "NIFTYALPHA50", "NIFTYQUALITY30", "NIFTYLOWVOL30",
    "NIFTY100LOWVOL30", "NIFTY50DIVPNTS", "NIFTYINDIALVOL",
    "INDIA VIX", "INDIAVIX",
    # ── Common ETFs on NSE ────────────────────────────────────────────────────
    "NIFTYBEES", "BANKBEES", "JUNIORBEES", "PSUBNKBEES", "ITBEES",
    "GOLDBEES", "SILVERBEES", "LIQUIDBEES", "LIQUIDCASE", "LIQUIDETF",
    "SETFNIF50", "SETFNN50", "SETFNIFIT", "SETFBSE100", "SETFNIFBK",
    "ICICIB22", "ICICITECH", "MOM30IETF", "MAFANG", "MOMOMENTUM",
    "MOSMALL250", "MONIFTY500", "MOVALUE", "MOQUALITY", "MOHEALTH",
    "NIFTIETF", "HDFCNIFTY", "HDFCMID150", "HDFCSML250", "HDFCSMALL",
    "HDFCLIQUID", "HDFCGOLD", "HDFCSILVER", "HDFCNIFIT", "HDFCNIFBAN",
    "SBIETFQLTY", "SBIETFPB", "SBIETFIT", "SBIETFCONS", "SBIETFPHARMA",
    "SBIETFMID150", "SBIETF1000",
    "KOTAKNIFTY", "KOTAKBANK", "KOTAKGOLD", "KOTAKSILVER",
    "KOTAKNIFBK", "KOTAKMID50", "KOTAKPSUBK", "KOTAKBANKSO",
    "AXISNIFTY", "AXISSMALL", "AXISNIFTYBK",
    "UTINIFTETF", "UTISENSETF", "UTIBANKETF", "UTINEXT50",
    "ABSLLIQUID", "ABSLPSE",
    "MIRAERASETF", "MIRAEASETBK", "MIRAEASETGD",
    "BSLGOLDETF", "BSLNIFTY",
    "TATANIFTY50", "TATADIGITAL", "TATASMCAP",
    "ITETF", "AUTIETF", "PHARMABEES", "INFRABEES",
    "CPSEETF", "CONSUMBEES", "PVTBANKETF",
    "NIFTYIETF", "NIF100IETF", "NIF100BEES", "NIFTYMID",
    "LOWVOLIETF", "ALPHAETF", "QUAL30IETF",
    "PSUBANK", "SHARIABEES",
}

_EXCLUDE_SUBSTRINGS = (
    "BEES", "ETF", "LIQUIDCASE", "LIQETF", "INDEX", "NIFTYBK",
    "IETF", "SETF",
)

def is_index_or_etf(ticker: str) -> bool:
    """Return True if the ticker looks like an index or ETF (name-based check)."""
    t = ticker.strip().upper()
    if t in _EXCLUDE_EXACT:
        return True
    for sub in _EXCLUDE_SUBSTRINGS:
        if sub in t:
            return True
    return False

def filter_stocks_only(df):
    """
    Drop rows that are indices or ETFs using two layers:
      1. Market Cap == 0 or null  ->  not a real stock (primary, most reliable)
      2. Name-pattern check       ->  catches anything with mcap missing from API
    Returns a clean df with only real equity stocks.
    """
    if df.empty:
        return df

    ticker_col = "Ticker" if "Ticker" in df.columns else df.columns[0]

    # Layer 1: market cap filter (ETFs/indices have 0 or null mcap in Chartink)
    if "Market Cap" in df.columns:
        def _mcap_ok(val):
            try:
                return float(str(val).replace(",", "").strip() or "0") > 0
            except Exception:
                return True   # if unparseable, keep and let layer 2 decide
        mcap_mask = df["Market Cap"].apply(_mcap_ok)
    else:
        mcap_mask = pd.Series([True] * len(df), index=df.index)

    # Layer 2: name pattern filter
    name_mask = ~df[ticker_col].astype(str).str.strip().apply(is_index_or_etf)

    before = len(df)
    df = df[mcap_mask & name_mask].reset_index(drop=True)
    removed = before - len(df)
    if removed:
        print(f"   🚫  Removed {removed} index/ETF row(s) — {len(df)} stocks remain")
    return df

# ══════════════════════════════════════════════════════════════════════════════
#  FETCH DATA FROM CHARTINK
# ══════════════════════════════════════════════════════════════════════════════
def fetch_chartink(session, scan_clause):
    try:
        page_headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        page = session.get(
            "https://chartink.com/screener/",
            headers=page_headers,
            timeout=20,
        )
        time.sleep(3)

        if page.status_code != 200:
            print(f"   ❌  Could not load chartink.com (HTTP {page.status_code})")
            return pd.DataFrame()

        soup = BeautifulSoup(page.text, "html.parser")
        meta = soup.find("meta", {"name": "csrf-token"})

        if meta and meta.get("content"):
            csrf_token = meta["content"]
            print(f"   🔑  CSRF from HTML meta tag: {csrf_token[:16]}...")
        else:
            raw = session.cookies.get("XSRF-TOKEN", "")
            csrf_token = requests.utils.unquote(raw)
            if csrf_token:
                print(f"   🔑  CSRF from cookie fallback: {csrf_token[:16]}...")
            else:
                print("   ❌  Could not find CSRF token — Chartink may have changed its page structure")
                return pd.DataFrame()

        post_headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            "Referer":          "https://chartink.com/screener/",
            "Origin":           "https://chartink.com",
            "X-Requested-With": "XMLHttpRequest",
            "X-CSRF-TOKEN":     csrf_token,
            "Content-Type":     "application/x-www-form-urlencoded; charset=UTF-8",
            "Accept":           "application/json, text/javascript, */*; q=0.01",
            "Accept-Language":  "en-US,en;q=0.9",
        }

        resp = session.post(
            PROCESS_URL,
            data={"scan_clause": scan_clause, "_token": csrf_token},
            headers=post_headers,
            timeout=25,
        )

        if resp.status_code == 419:
            print("   ❌  HTTP 419 — CSRF token rejected by Chartink.")
            print("       Possible fixes:")
            print("       1. Increase PAUSE_BETWEEN at the top of this script (try 8+)")
            print("       2. Open chartink.com in your browser and log in, then run again")
            print("       3. Wait a few minutes and try again")
            return pd.DataFrame()

        if resp.status_code != 200:
            print(f"   ❌  HTTP {resp.status_code}")
            return pd.DataFrame()

        try:
            data = resp.json()
        except Exception:
            if "<html" in resp.text.lower() or "<!doctype" in resp.text.lower():
                print("   ❌  Chartink returned an HTML page — you may need to log in via browser")
                print(f"       Response snippet: {resp.text[:200]}")
            else:
                print(f"   ❌  Could not parse response. Snippet: {resp.text[:300]}")
            return pd.DataFrame()

        if "data" not in data:
            print(f"   ⚠️  Unexpected response keys: {list(data.keys())}")
            print(f"       Full response: {str(data)[:300]}")
            return pd.DataFrame()

        raw_data = data["data"]
        print(f"   📊  Records received: {len(raw_data)}")
        if not raw_data:
            print("   ⚠️  API returned 0 rows")
            return pd.DataFrame()

        # ── ONE-TIME DEBUG: print all raw fields from Chartink ────────────────
        if not getattr(fetch_chartink, "_fields_printed", False):
            fetch_chartink._fields_printed = True
            sample = raw_data[0]
            print(f"\n   🔍  RAW CHARTINK FIELDS (all keys in API response):")
            for k, v in sample.items():
                print(f"       {k!r:30s} = {str(v)[:60]!r}")
            print()

        df = pd.DataFrame(raw_data)

        rename_map = {
            "nsecode":  "Ticker",
            "bsecode":  "BSE Code",
            "per_chg":  "% Change",
            "close":    "Close",
            "volume":   "Volume",
            "turnover": "Turnover (Cr)",
            "mcap":     "Market Cap",
            "sr_no":    "Sr No",
        }
        df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

        if "Ticker" in df.columns:
            cols = ["Ticker"] + [c for c in df.columns if c != "Ticker"]
            df   = df[cols]

        return df

    except Exception as e:
        print(f"   ❌  Unexpected error: {e}")
        import traceback; traceback.print_exc()
        return pd.DataFrame()

# ══════════════════════════════════════════════════════════════════════════════
#  SECTOR FETCH  (yfinance — NSE suffix)
# ══════════════════════════════════════════════════════════════════════════════
def fetch_sectors(tickers: list) -> dict:
    """
    Returns {TICKER: {"sector": ..., "industry": ...}} for each ticker.
    Uses yfinance with .NS suffix. Falls back to 'Unknown' on any error.
    Industry is the fine-grained grouping (e.g. "Software—Application"),
    sector is the broad bucket (e.g. "Technology").
    """
    try:
        import yfinance as yf
    except ImportError:
        print("\n⚠️  yfinance not installed — sector tab will show 'Unknown'.")
        print("    Run:  pip install yfinance   then re-run the script.")
        return {t.upper(): {"sector": "Unknown", "industry": "Unknown"} for t in tickers}

    sector_map = {}
    total = len(tickers)
    print(f"\n🏭  Fetching sector + industry for {total} tickers via yfinance...")

    for i, ticker in enumerate(tickers, 1):
        symbol = ticker.upper() + ".NS"
        try:
            info     = yf.Ticker(symbol).info
            sector   = info.get("sector")   or "Unknown"
            industry = info.get("industry") or info.get("industryDisp") or sector
        except Exception:
            sector = industry = "Unknown"
        sector_map[ticker.upper()] = {"sector": sector, "industry": industry}

        if i % 10 == 0 or i == total:
            pct = int(i / total * 30)
            bar = "█" * pct + "░" * (30 - pct)
            print(f"   [{bar}] {i}/{total}", end="\r")

    print(f"\n   ✅  Sector/industry data ready for {total} tickers")
    return sector_map


# ══════════════════════════════════════════════════════════════════════════════
#  BUILD SECTOR TAB
# ══════════════════════════════════════════════════════════════════════════════
def build_sector_tab(wb, unique: list, counts: dict, sector_map: dict,
                     now_str: str, fill_fn, thin_fn, medium_fn):
    """
    Adds a 'By Sector' sheet with a sortable table:
      # | Ticker | Sector | Industry | In # Screeners
    Sorted by Industry (fine-grained) then Ticker.
    Each unique Industry gets its own background colour band so
    peers like TCS/HCL/Infosys appear visually grouped.
    """
    C_DARK  = "1A3A5C"; C_HDR  = "2E5FA3"; C_WHITE = "FFFFFF"
    C_GREEN = "00763D"; C_GOLD = "FFD700"

    INDUSTRY_COLORS = [
        "E3F2FD", "E8F5E9", "FFF8E1", "FCE4EC", "F3E5F5",
        "E0F7FA", "FBE9E7", "F1F8E9", "EDE7F6", "E0F2F1",
        "FFF3E0", "E8EAF6", "F9FBE7", "FFEBEE", "E1F5FE",
        "F0F4C3", "FFE0B2", "D7CCC8", "CFD8DC", "DCEDC8",
    ]

    ws = wb.create_sheet(title="By Sector")
    ws.sheet_properties.tabColor = "70AD47"

    ws.row_dimensions[1].height = 5
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 14

    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value     = f"  BY SECTOR & INDUSTRY — {len(unique)} stocks   |   {now_str}"
    c.font      = Font(name="Arial", bold=True, size=13, color="000000")
    c.fill      = fill_fn(C_GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A3:E3")
    c = ws["A3"]
    all_info        = sector_map.values()
    industries_cnt  = len({v["industry"] for v in all_info if v["industry"] != "Unknown"})
    sectors_cnt     = len({v["sector"]   for v in sector_map.values() if v["sector"] != "Unknown"})
    c.value     = (f"  {sectors_cnt} sectors  |  {industries_cnt} industries  "
                   f"|  grouped by industry  |  sortable via Excel filter")
    c.font      = Font(name="Arial", italic=True, size=9, color="555555")
    c.fill      = fill_fn(C_GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center")

    # Column headers row 4
    ws.row_dimensions[4].height = 22
    headers = [
        ("#",              "A",  5),
        ("Ticker",         "B", 14),
        ("Sector",         "C", 22),
        ("Industry",       "D", 32),
        ("In # Screeners", "E", 16),
    ]
    for hdr, col, width in headers:
        c = ws[f"{col}4"]
        c.value     = hdr
        c.font      = Font(name="Arial", bold=True, size=9, color=C_WHITE)
        c.fill      = fill_fn(C_HDR)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = medium_fn()
        ws.column_dimensions[col].width = width

    # Build rows sorted by Industry → Ticker
    rows = []
    for t in unique:
        u    = t.upper()
        info = sector_map.get(u, {"sector": "Unknown", "industry": "Unknown"})
        rows.append((
            info["sector"],
            info["industry"],
            u,
            counts.get(u, 1),
        ))
    rows.sort(key=lambda x: (x[1].lower(), x[2]))   # sort by industry, then ticker

    # Assign colour per industry
    industry_color = {}
    for _, industry, _, _ in rows:
        if industry not in industry_color:
            idx = len(industry_color) % len(INDUSTRY_COLORS)
            industry_color[industry] = INDUSTRY_COLORS[idx]

    # Write data rows
    for i, (sector, industry, ticker, cnt) in enumerate(rows):
        r  = 5 + i
        bg = industry_color.get(industry, "FFFFFF")
        ws.row_dimensions[r].height = 15

        for col, val, bold, align, color in [
            ("A", i + 1,    False, "center", "777777"),
            ("B", ticker,   True,  "center", C_DARK),
            ("C", sector,   False, "left",   "000000"),
            ("D", industry, False, "left",   "000000"),
            ("E", cnt,      True,  "center", C_GREEN if cnt > 1 else "000000"),
        ]:
            c = ws[f"{col}{r}"]
            c.value     = val
            c.font      = Font(name="Arial", bold=bold, size=9, color=color)
            c.fill      = fill_fn(bg)
            c.alignment = Alignment(horizontal=align, vertical="center")
            c.border    = thin_fn()

    ws.auto_filter.ref = f"A4:E{4 + len(rows)}"
    ws.freeze_panes   = "A5"
    return ws


# ══════════════════════════════════════════════════════════════════════════════
#  BUILD EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def build_excel(screener_results, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    def fill(h):
        return PatternFill("solid", fgColor=h)

    def thin(color="CCCCCC"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    def medium(color="666666"):
        s = Side(style="medium", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    C_DARK  = "1A3A5C"; C_HDR   = "2E5FA3"; C_GOLD  = "FFD700"; C_WHITE = "FFFFFF"
    C_S1    = "FFFFFF"; C_S2    = "EBF3FB"; C_TICK  = "DDEEFF"
    C_POSB  = "E6F4EA"; C_POSF  = "137333"
    C_NEGB  = "FCE8E6"; C_NEGF  = "C5221F"
    C_NOTE  = "FFFBEA"; C_GREEN = "00763D"; C_RED   = "C00000"

    now_str = datetime.datetime.now().strftime("%d %b %Y  %H:%M")
    all_tickers_flat = []
    tab_colors = ["2E75B6","70AD47","C00000","FF8C00","7030A0",
                  "00B0F0","BF9000","375623","4472C4","ED7D31",
                  "A9D18E","0070C0","7F7F7F","FFC000","FF0000"]

    for idx, (name, df) in enumerate(screener_results):
        safe = (name[:31].replace("/","_").replace("\\","_")
                .replace("?","").replace("*","")
                .replace("[","").replace("]","").replace(":",""))
        ws = wb.create_sheet(title=safe)
        ws.sheet_properties.tabColor = tab_colors[idx % len(tab_colors)]

        n_cols   = max(len(df.columns), 4) if not df.empty else 4
        last_col = get_column_letter(n_cols)

        ws.row_dimensions[1].height = 5
        ws.row_dimensions[2].height = 28
        ws.row_dimensions[3].height = 14

        ws.merge_cells(f"A2:{last_col}2")
        c = ws["A2"]
        c.value     = f"  {name.upper()}   |   {len(df)} stocks   |   {now_str}"
        c.font      = Font(name="Arial", bold=True, size=12, color=C_WHITE)
        c.fill      = fill(C_DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(f"A3:{last_col}3")
        c = ws["A3"]
        c.value     = f"  Source: chartink.com   |   Run: {now_str}"
        c.font      = Font(name="Arial", italic=True, size=8, color="AAAAAA")
        c.fill      = fill(C_DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")

        if df.empty:
            ws["A5"].value = "⚠️  No data retrieved for this screener"
            ws["A5"].font  = Font(name="Arial", size=11, color=C_RED)
            continue

        # ── Drop indices / ETFs — keep real stocks only ───────────────────────
        df = filter_stocks_only(df)
        ws["A2"].value = f"  {name.upper()}   |   {len(df)} stocks   |   {now_str}"
        if df.empty:
            ws["A5"].value = "⚠️  No stocks after excluding indices/ETFs"
            ws["A5"].font  = Font(name="Arial", size=11, color=C_RED)
            continue

        pct_cols = {col for col in df.columns
                    if col == "% Change"
                    or df[col].dropna().astype(str).head(20).str.contains(r"%").mean() > 0.3}

        num_cols = set()
        for col in df.columns:
            if col in pct_cols: continue
            try:
                pd.to_numeric(df[col].dropna().head(20), errors="raise")
                num_cols.add(col)
            except Exception:
                pass

        ticker_col  = "Ticker" if "Ticker" in df.columns else df.columns[0]
        ticker_cidx = list(df.columns).index(ticker_col) + 1

        ws.row_dimensions[4].height = 22
        for ci, col_name in enumerate(df.columns, 1):
            c = ws.cell(row=4, column=ci)
            c.value     = col_name
            c.font      = Font(name="Arial", bold=True, size=9, color=C_WHITE)
            c.fill      = fill(C_HDR)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border    = medium()
            ws.column_dimensions[get_column_letter(ci)].width = max(len(str(col_name)) + 3, 10)

        for ri, (_, row) in enumerate(df.iterrows()):
            er = ri + 5
            ws.row_dimensions[er].height = 15
            bg = C_S1 if ri % 2 == 0 else C_S2

            for ci, (col_name, val) in enumerate(zip(df.columns, row), 1):
                c   = ws.cell(row=er, column=ci)
                txt = str(val).strip() if pd.notna(val) else ""
                c.value = txt

                if ci == ticker_cidx:
                    c.font      = Font(name="Arial", bold=True, size=9, color=C_DARK)
                    c.fill      = fill(C_TICK)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border    = thin("AACCEE")
                elif col_name in pct_cols and txt:
                    try:
                        neg = float(txt.replace("%","").replace(",","")) < 0
                    except Exception:
                        neg = txt.startswith("-")
                    c.font      = Font(name="Arial", size=9, color=C_NEGF if neg else C_POSF)
                    c.fill      = fill(C_NEGB if neg else C_POSB)
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    c.border    = thin()
                elif col_name in num_cols and txt:
                    c.font      = Font(name="Arial", size=9)
                    c.fill      = fill(bg)
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    c.border    = thin()
                else:
                    c.font      = Font(name="Arial", size=9)
                    c.fill      = fill(bg)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border    = thin()

                cur_w = ws.column_dimensions[get_column_letter(ci)].width
                ws.column_dimensions[get_column_letter(ci)].width = min(
                    max(cur_w, len(txt) + 2), 30
                )

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:{last_col}{4 + len(df)}"

        tickers = df[ticker_col].dropna().astype(str).str.strip().tolist()
        tickers = [t for t in tickers if t not in ("", "nan", "NaN")]
        tickers = [t for t in tickers if not is_index_or_etf(t)]
        all_tickers_flat.extend(tickers)

    # ── Summary tab ───────────────────────────────────────────────────────────
    from collections import Counter
    counts = Counter(t.upper() for t in all_tickers_flat)
    seen, unique = set(), []
    for t in all_tickers_flat:
        u = t.upper()
        if u not in seen and not is_index_or_etf(u):
            seen.add(u); unique.append(u)
    unique.sort()

    ws_s = wb.create_sheet(title="All Tickers")
    ws_s.sheet_properties.tabColor = "FFD700"
    ws_s.row_dimensions[1].height  = 5
    ws_s.row_dimensions[2].height  = 30
    ws_s.row_dimensions[3].height  = 14

    ws_s.merge_cells("A2:G2")
    c = ws_s["A2"]
    c.value     = f"  ALL TICKERS — DEDUPLICATED   |   {len(unique)} unique   |   {now_str}"
    c.font      = Font(name="Arial", bold=True, size=13, color="000000")
    c.fill      = fill(C_GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws_s.merge_cells("A3:G3")
    c = ws_s["A3"]
    c.value     = f"  Total raw: {len(all_tickers_flat)}   |   After dedup: {len(unique)}"
    c.font      = Font(name="Arial", italic=True, size=9, color="555555")
    c.fill      = fill(C_GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws_s.row_dimensions[5].height = 18
    ws_s.merge_cells("A5:G5")
    c = ws_s["A5"]
    c.value     = "  ✂  COPY-PASTE READY — Comma-Separated Ticker List"
    c.font      = Font(name="Arial", bold=True, size=10, color=C_WHITE)
    c.fill      = fill(C_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws_s.row_dimensions[6].height = 80
    ws_s.merge_cells("A6:G6")
    c = ws_s["A6"]
    c.value     = ", ".join(unique)
    c.font      = Font(name="Courier New", size=9)
    c.fill      = fill(C_NOTE)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    bd = Side(style="medium", color=C_HDR)
    c.border    = Border(left=bd, right=bd, top=bd, bottom=bd)

    for cl, w in [("A",6),("B",14),("C",20),("D",12),("E",12),("F",12),("G",12)]:
        ws_s.column_dimensions[cl].width = w

    ws_s.row_dimensions[8].height = 20
    for cl, hdr in [("A","#"),("B","Ticker"),("C","In # Screeners"),("D",""),("E",""),("F",""),("G","")]:
        c = ws_s[f"{cl}8"]
        c.value     = hdr
        c.font      = Font(name="Arial", bold=True, size=9, color=C_WHITE)
        c.fill      = fill(C_HDR)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = medium()

    for i, ticker in enumerate(unique):
        r   = 9 + i
        bg  = C_S1 if i % 2 == 0 else C_S2
        cnt = counts.get(ticker, 1)
        ws_s.row_dimensions[r].height = 15
        for cl, val in [("A", i + 1), ("B", ticker), ("C", cnt)]:
            c = ws_s[f"{cl}{r}"]
            c.value     = val
            c.font      = Font(name="Arial", size=9, bold=(cl == "B"),
                               color=C_GREEN if cnt > 1 else "000000")
            c.fill      = fill(C_S2 if cnt > 1 else bg)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = thin()

    ws_s.freeze_panes = "A9"
    wb.move_sheet("All Tickers", offset=-len(wb.sheetnames) + 1)

    # ── Sector tab ────────────────────────────────────────────────────────────
    sector_map = fetch_sectors(unique)
    build_sector_tab(wb, unique, counts, sector_map, now_str, fill, thin, medium)
    wb.move_sheet("By Sector", offset=-len(wb.sheetnames) + 1)

    wb.save(output_path)
    print(f"\n✅  Excel saved: {output_path}")
    return unique

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  CHARTINK SCREENER AUTOMATION")
    print(f"  Total screeners: {len(SCREENERS)}")
    print("=" * 60)

    session = requests.Session()
    all_results = []

    for name, mode, value in SCREENERS:
        print(f"\n📡  Running: {name}")
        if mode.lower() != "clause":
            print("   ⚠️  Only 'clause' mode supported — skipping")
            continue

        df = fetch_chartink(session, value)
        if df.empty:
            print("   ⚠️  No results")
            all_results.append((name, df))
        else:
            print(f"   ✅  {len(df)} stocks found")
            all_results.append((name, df))

        time.sleep(PAUSE_BETWEEN)

    if not all_results:
        print("\n❌  No data retrieved.")
        if not os.environ.get("GITHUB_ACTIONS"):
            input("\nPress Enter to exit...")
        sys.exit(1)

    print(f"\n\n📊  Building Excel...")
    unique = build_excel(all_results, OUTPUT_FILE)

    print(f"\n🎯  {len(unique)} unique tickers")
    print(f"    Preview: {', '.join(unique[:12])}{'...' if len(unique) > 12 else ''}")

    if not os.environ.get("GITHUB_ACTIONS"):
        try:
            os.startfile(OUTPUT_FILE)
        except Exception:
            import subprocess
            subprocess.Popen(["start", OUTPUT_FILE], shell=True)

        input("\nPress Enter to close...")
    else:
        # Running in CI — write the output path so the workflow can pick it up
        print(f"::notice::Excel file ready at {OUTPUT_FILE}")
