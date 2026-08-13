#!/usr/bin/env python3
"""
NSE Inside Bar Scanner — local Python version, writes results into Google Sheets.

WHAT IT DOES
    Scans NSE symbols priced at or below --max-price (default Rs 500), while
    skipping low-liquidity names using --min-avg-volume and
    --min-turnover-lakhs, for the "inside bar" pattern on either a daily or
    hourly timeframe, and writes color-coded signals into a worksheet tab of
    your Google Sheet:
        WATCH              -> latest candle is an inside bar (amber). No
                               trade yet, a name to watch for the next
                               candle's breakout.
        BULLISH (green)    -> previous candle was an inside bar, and the
        BEARISH (red)         LATEST candle closed beyond the mother bar's
                               high/low (close-based confirmation).
    Confirmed breakouts get scored 0-3 against false-breakout filters:
        trend  - close on the correct side of the 50-period EMA
        volume - breakout candle volume >= 1.5x the prior 20-period average
        RSI    - RSI(14) in a sane momentum band, not already at an extreme
    Entry / Stop Loss / Target are computed off the mother bar's range:
        BULLISH: entry = mother high, stop = mother low,
                 target = mother high + (mother high - mother low)
        BEARISH: entry = mother low,  stop = mother high,
                 target = mother low  - (mother high - mother low)
        WATCH:   both potential trigger levels are shown; direction isn't
                 known yet so no stop/target until it actually breaks out.

PATTERN LOGIC (3-candle window, newest = candle[-1]):
    candle[-3] = mother bar candidate (for confirmed breakout check)
    candle[-2] = inside bar candidate (must be inside candle[-3] high/low)
    candle[-1] = breakout/signal candle:
        - If candle[-2] was inside candle[-3] AND candle[-1] closes
          beyond candle[-3] high/low => BULLISH / BEARISH
        - If candle[-1] is inside candle[-2]  => WATCH (fresh inside bar)

TWO-STAGE SCAN (keeps this fast even across the whole NSE list)
    Stage 1: a quick 5-day price check on every symbol, to drop anything
             priced above --max-price before doing any real work.
    Stage 2: full historical fetch + indicator calc, only on what's left.

DATA SOURCE
    Yahoo Finance via the `yfinance` package (SYMBOL.NS). Free, no key.
    60-minute intraday data is typically ~15 min delayed and can be patchy
    for illiquid names — treat it as "near-live" for spotting setups, not
    as a tick-accurate execution feed.

USAGE
    python nse_scanner.py --timeframe daily
    python nse_scanner.py --timeframe hourly
    python nse_scanner.py --timeframe daily --max-price 300
    python nse_scanner.py --timeframe daily --min-avg-volume 200000 --min-turnover-lakhs 100
    python nse_scanner.py --timeframe daily --symbols-file my_symbols.txt

See SETUP.md for one-time Google Sheets credential setup and how to
schedule this with cron so it runs automatically, hourly and daily.
"""

import argparse
import csv
import sys
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import yfinance as yf

# ============================== CONFIG ======================================
SPREADSHEET_ID = "1YDfmA6wa8t8uqPsavOfsvPzM8SpOTFgxql9wp-07uKk"  # from your sheet URL

CONFIG = {
    "vol_ratio_min": 1.5,
    "min_avg_volume": 100_000,
    "min_turnover_lakhs": 50.0,
    "rsi_bull_min": 40, "rsi_bull_max": 80,
    "rsi_bear_min": 20, "rsi_bear_max": 60,
    "default_max_price": 500.0,
}

TIMEFRAMES = {
    "daily": {
        "interval": "1d", "period": "6mo", "min_candles": 60,
        "sheet_name": "Scan Results - Daily",
        "ema_trend": 50, "rsi_period": 14, "vol_avg_period": 20,
    },
    "hourly": {
        "interval": "60m", "period": "3mo", "min_candles": 60,
        "sheet_name": "Scan Results - Hourly",
        "ema_trend": 50, "rsi_period": 14, "vol_avg_period": 20,
    },
}

NIFTY50_FALLBACK = [
    "RELIANCE", "TCS", "HDFCBANK", "ICICIBANK", "INFY", "HINDUNILVR", "ITC", "SBIN",
    "BHARTIARTL", "KOTAKBANK", "LT", "AXISBANK", "BAJFINANCE", "ASIANPAINT", "MARUTI",
    "SUNPHARMA", "TITAN", "ULTRACEMCO", "WIPRO", "ONGC", "NTPC", "POWERGRID", "NESTLEIND",
    "HCLTECH", "TATAMOTORS", "TATASTEEL", "ADANIENT", "ADANIPORTS", "JSWSTEEL",
    "COALINDIA", "BAJAJFINSV", "TECHM", "INDUSINDBK", "DRREDDY", "GRASIM", "CIPLA",
    "EICHERMOT", "BRITANNIA", "DIVISLAB", "HEROMOTOCO", "BPCL", "HDFCLIFE", "SBILIFE",
    "APOLLOHOSP", "TATACONSUM", "UPL", "BAJAJ-AUTO",
]  # Approximate starter list only — pass --symbols-file for full NSE coverage.

SIGNAL_RANK = {"BULLISH": 0, "BEARISH": 1, "WATCH": 2}
SIGNAL_COLOR = {
    "BULLISH": {"red": 0.80, "green": 0.94, "blue": 0.80},
    "BEARISH": {"red": 0.97, "green": 0.80, "blue": 0.80},
    "WATCH": {"red": 1.00, "green": 0.95, "blue": 0.75},
}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}

HEADER = [
    "Symbol", "Signal", "Pattern", "Mother High", "Mother Low",
    "Last Close", "Avg Volume", "Turnover (L)", "Vol Ratio",
    "Trend vs EMA", "RSI", "Score", "Entry", "Stop Loss", "Target", "Updated At",
]


# ============================ MARKET HOURS ==================================
def is_market_hours() -> bool:
    now = datetime.now(ZoneInfo("Asia/Kolkata"))
    if now.weekday() >= 5:  # Sat/Sun
        return False
    hm = now.strftime("%H%M")
    return "0915" <= hm <= "1530"


# ============================ SYMBOL LIST ===================================
def load_symbols(symbols_file: str | None) -> list[str]:
    if symbols_file:
        path = Path(symbols_file)
        with path.open() as f:
            return [line.strip() for line in f if line.strip()]

    try:
        url = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if resp.status_code == 200:
            reader = csv.reader(resp.text.splitlines())
            next(reader)  # header
            syms = [row[0].strip() for row in reader if row and row[0].strip()]
            if syms:
                return syms
    except Exception:
        pass

    print(
        "Could not fetch the live NSE symbol list (NSE often blocks non-browser "
        "requests). Using a 47-stock starter list instead.\n"
        "For full NSE coverage: download EQUITY_L.csv from nseindia.com yourself "
        "and pass it with --symbols-file (one symbol per line).",
        file=sys.stderr,
    )
    return NIFTY50_FALLBACK


# ============================ DATA FETCH ====================================
def chunked(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i : i + n]


def fetch_batch(symbols: list[str], interval: str, period: str) -> pd.DataFrame:
    tickers = [s + ".NS" for s in symbols]
    return yf.download(
        tickers=tickers,
        period=period,
        interval=interval,
        group_by="ticker",
        threads=True,
        progress=False,
        auto_adjust=False,
    )


def extract_symbol_df(data: pd.DataFrame, symbol: str, chunk_len: int) -> pd.DataFrame | None:
    """
    Extract per-symbol OHLCV DataFrame from a yfinance batch download result.

    yfinance >= 0.2 returns a MultiIndex with (field, ticker) at the column level,
    i.e. level-0 = field ('Open','High','Low','Close','Volume','Adj Close'),
         level-1 = ticker ('RELIANCE.NS', ...).
    Older versions used (ticker, field) OR flat columns when only 1 ticker was fetched.
    We handle all three cases explicitly to avoid the most common source of false signals.
    """
    ticker = symbol + ".NS"

    if isinstance(data.columns, pd.MultiIndex):
        level0_vals = data.columns.get_level_values(0).unique().tolist()
        level1_vals = data.columns.get_level_values(1).unique().tolist()

        # Modern yfinance: level-0 = field, level-1 = ticker  e.g. ('Close', 'RELIANCE.NS')
        if ticker in level1_vals:
            df = data.xs(ticker, level=1, axis=1).copy()
        # Legacy yfinance: level-0 = ticker, level-1 = field  e.g. ('RELIANCE.NS', 'Close')
        elif ticker in level0_vals:
            df = data[ticker].copy()
        else:
            return None
    else:
        # Single-ticker download — yfinance returns flat columns directly
        df = data.copy()

    if df is None or len(df) == 0:
        return None

    # If there are still nested column levels (edge-case), flatten them
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    try:
        df = df.dropna(subset=["Open", "High", "Low", "Close"])
    except KeyError:
        return None

    return df if len(df) > 0 else None


def prefilter_by_price(symbols: list[str], max_price: float, chunk_size: int) -> list[str]:
    """Stage 1: quick price check to drop anything above max_price before the
    expensive historical fetch. Uses a light 5-day/1-day pull, not the full
    lookback needed for indicators."""
    keep = []
    chunks = list(chunked(symbols, chunk_size))
    for idx, chunk in enumerate(chunks, 1):
        try:
            data = fetch_batch(chunk, "1d", "5d")
        except Exception as e:
            print(f"  price-check batch {idx}/{len(chunks)} failed: {e}", file=sys.stderr)
            continue
        for sym in chunk:
            try:
                df = extract_symbol_df(data, sym, len(chunk))
                if df is None or len(df) == 0:
                    continue
                last_close = float(df["Close"].iloc[-1])
                if pd.isna(last_close):
                    continue
                if last_close <= max_price:
                    keep.append(sym)
            except Exception:
                continue
        print(f"  price-check batch {idx}/{len(chunks)} done — {len(keep)} under Rs {max_price} so far")
        time.sleep(0.5)
    return keep


# ============================ INDICATORS ====================================
def calc_rsi(close: pd.Series, period: int) -> pd.Series:
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))


# ============================ CORE LOGIC ====================================
def analyze_symbol(
    symbol: str,
    df: pd.DataFrame,
    tf: dict,
    min_avg_volume: int,
    min_turnover_lakhs: float,
) -> dict | None:
    """
    Evaluate the last 3 candles for inside-bar breakout signals.

    Candle window (newest last):
        df.iloc[-3]  = candle_3  (potential mother bar for confirmed breakout)
        df.iloc[-2]  = candle_2  (potential inside bar / mother bar for WATCH)
        df.iloc[-1]  = candle_1  (the current / most-recent candle)

    CONFIRMED BREAKOUT  (BULLISH / BEARISH):
        Condition:  candle_2 is fully inside candle_3
                    AND candle_1 closes ABOVE candle_3 high  → BULLISH
                    OR  candle_1 closes BELOW candle_3 low   → BEARISH
        Rationale:  We require the close to decisively exceed the MOTHER bar's
                    range, not just the inside bar's range, to avoid weak
                    breakouts that stall at the inside bar's edges.

    WATCH:
        Condition:  candle_1 is fully inside candle_2 (fresh inside bar).
        Rationale:  Direction is unknown; show both trigger levels so the
                    trader can set a bracket order for the next session.

    Scoring (0-3, applied only to confirmed breakouts):
        +1  trend  — close is on the correct side of the 50-period EMA
        +1  volume — breakout-candle volume >= 1.5× the 20-period average
        +1  RSI    — RSI(14) is in a momentum-sane band (not overbought/oversold)

    Liquidity gate (applied before any pattern checks):
        Skips stocks where the 20-period average volume < min_avg_volume OR
        average rupee turnover < min_turnover_lakhs. This prevents signals
        on stocks that are technically clean but impossible/expensive to trade.
    """
    if df is None or len(df) < tf["min_candles"]:
        return None

    n = len(df)
    closes = df["Close"]
    ema = closes.ewm(span=tf["ema_trend"], adjust=False, min_periods=tf["ema_trend"]).mean()
    rsi = calc_rsi(closes, tf["rsi_period"])
    vol = df["Volume"]

    ema_last = ema.iloc[-1]
    rsi_last = rsi.iloc[-1]
    if pd.isna(ema_last) or pd.isna(rsi_last):
        return None

    # Volume average over the 20 candles BEFORE the current one (exclude current)
    avg_vol = vol.iloc[max(0, n - 1 - tf["vol_avg_period"]) : n - 1].mean()

    # ----- 3-candle window -----
    # candle_1 = most recent (the signal / breakout candle)
    # candle_2 = one before that (inside bar candidate or mother bar for WATCH)
    # candle_3 = two before that (mother bar candidate for confirmed breakout)
    candle_1 = df.iloc[-1]
    candle_2 = df.iloc[-2]
    candle_3 = df.iloc[-3]

    # Guard: ensure OHLC values are usable floats
    try:
        c1_high  = float(candle_1["High"])
        c1_low   = float(candle_1["Low"])
        c1_close = float(candle_1["Close"])
        c1_vol   = float(candle_1["Volume"])
        c2_high  = float(candle_2["High"])
        c2_low   = float(candle_2["Low"])
        c3_high  = float(candle_3["High"])
        c3_low   = float(candle_3["Low"])
    except (TypeError, ValueError):
        return None

    if any(pd.isna(v) for v in [c1_high, c1_low, c1_close, c2_high, c2_low, c3_high, c3_low]):
        return None

    # ── Liquidity gate ──────────────────────────────────────────────────────
    # Compute daily turnover in lakhs using last close × avg volume
    avg_vol_safe = float(avg_vol) if (avg_vol is not None and not pd.isna(avg_vol)) else 0.0
    turnover_lakhs = (c1_close * avg_vol_safe) / 100_000 if avg_vol_safe > 0 else 0.0

    if avg_vol_safe < min_avg_volume or turnover_lakhs < min_turnover_lakhs:
        return None
    # ────────────────────────────────────────────────────────────────────────

    result = {
        "symbol": symbol, "pattern": "", "signal": "NONE",
        "mother_high": None, "mother_low": None,
        "last_close": round(c1_close, 2),
        "avg_volume": int(round(avg_vol_safe)),
        "turnover_lakhs": round(turnover_lakhs, 2),
        "vol_ratio": None, "trend": "", "rsi": round(float(rsi_last), 2), "score": 0,
        "entry": "", "stop_loss": "", "target": "",
        "updated_at": datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M"),
    }

    # ------------------------------------------------------------------ #
    #  Case 1: CONFIRMED BREAKOUT                                         #
    #    candle_2 is an inside bar relative to candle_3 (mother bar),     #
    #    AND candle_1 (current) closes beyond the mother bar's range.     #
    # ------------------------------------------------------------------ #
    candle_2_is_inside_candle_3 = (c2_high <= c3_high) and (c2_low >= c3_low)

    if candle_2_is_inside_candle_3:
        mother_high = round(c3_high, 2)
        mother_low  = round(c3_low,  2)
        rng = round(mother_high - mother_low, 2)

        # Vol ratio: use candle_1 volume vs prior 20-candle average
        c1_vol_safe = c1_vol if not pd.isna(c1_vol) else 0.0
        vol_ratio = (c1_vol_safe / avg_vol_safe) if avg_vol_safe > 0 else None
        safe_vol_ratio = round(float(vol_ratio), 2) if (vol_ratio is not None and not pd.isna(vol_ratio)) else None

        if c1_close > c3_high:          # BULLISH breakout
            result.update(
                pattern="Breakout confirmed",
                signal="BULLISH",
                mother_high=mother_high,
                mother_low=mother_low,
                entry=mother_high,
                stop_loss=mother_low,
                target=round(mother_high + rng, 2),
                vol_ratio=safe_vol_ratio,
            )
            result["trend"] = (
                f"above EMA{tf['ema_trend']}" if c1_close > float(ema_last)
                else f"below EMA{tf['ema_trend']}"
            )
            if c1_close > float(ema_last):
                result["score"] += 1
            if safe_vol_ratio is not None and safe_vol_ratio >= CONFIG["vol_ratio_min"]:
                result["score"] += 1
            if CONFIG["rsi_bull_min"] <= float(rsi_last) <= CONFIG["rsi_bull_max"]:
                result["score"] += 1

        elif c1_close < c3_low:         # BEARISH breakout
            result.update(
                pattern="Breakout confirmed",
                signal="BEARISH",
                mother_high=mother_high,
                mother_low=mother_low,
                entry=mother_low,
                stop_loss=mother_high,
                target=round(mother_low - rng, 2),
                vol_ratio=safe_vol_ratio,
            )
            result["trend"] = (
                f"below EMA{tf['ema_trend']}" if c1_close < float(ema_last)
                else f"above EMA{tf['ema_trend']}"
            )
            if c1_close < float(ema_last):
                result["score"] += 1
            if safe_vol_ratio is not None and safe_vol_ratio >= CONFIG["vol_ratio_min"]:
                result["score"] += 1
            if CONFIG["rsi_bear_min"] <= float(rsi_last) <= CONFIG["rsi_bear_max"]:
                result["score"] += 1

    # ------------------------------------------------------------------ #
    #  Case 2: FRESH INSIDE BAR → WATCH                                   #
    #    candle_1 (the just-closed candle) is inside candle_2.            #
    #    No confirmed breakout yet — alert the trader to watch.           #
    # ------------------------------------------------------------------ #
    if result["signal"] == "NONE":
        candle_1_is_inside_candle_2 = (c1_high <= c2_high) and (c1_low >= c2_low)
        if candle_1_is_inside_candle_2:
            mother_high = round(c2_high, 2)
            mother_low  = round(c2_low,  2)
            result.update(
                pattern="Inside bar (watchlist)",
                signal="WATCH",
                mother_high=mother_high,
                mother_low=mother_low,
                entry=f"Buy>{mother_high} / Sell<{mother_low}",
            )
            result["trend"] = (
                f"above EMA{tf['ema_trend']}" if c1_close > float(ema_last)
                else f"below EMA{tf['ema_trend']}"
            )

    return result if result["signal"] != "NONE" else None


# ============================ GOOGLE SHEETS =================================
def get_gspread_client(credentials_path: str):
    import gspread
    from google.oauth2.service_account import Credentials

    if not Path(credentials_path).exists():
        print(
            f"\nCredentials file not found: {credentials_path}\n"
            "See SETUP.md for how to create a Google service account key.",
            file=sys.stderr,
        )
        sys.exit(1)

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return gspread.authorize(creds)


def result_sort_key(result: dict) -> tuple:
    """Rank actionable, liquid, high-confirmation setups first."""
    return (
        SIGNAL_RANK.get(result["signal"], 9),
        -result["score"],
        -float(result.get("vol_ratio") or 0),
        -float(result.get("turnover_lakhs") or 0),
        result["symbol"],
    )


def write_results(gc, spreadsheet_id: str, sheet_name: str, results: list[dict]):
    import gspread

    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=2000, cols=len(HEADER) + 2)

    results_sorted = sorted(results, key=result_sort_key)
    rows = [
        [r["symbol"], r["signal"], r["pattern"], r["mother_high"], r["mother_low"],
         r["last_close"], r["avg_volume"], r["turnover_lakhs"], r["vol_ratio"],
         r["trend"], r["rsi"], r["score"], r["entry"], r["stop_loss"],
         r["target"], r["updated_at"]]
        for r in results_sorted
    ]

    ws.clear()
    ws.update([HEADER] + rows, "A1")

    # Reset formatting, then color-code contiguous blocks of the same signal.
    num_cols = len(HEADER)
    requests = [{
        "repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": max(len(rows) + 1, 2),
                       "startColumnIndex": 0, "endColumnIndex": num_cols},
            "cell": {"userEnteredFormat": {"backgroundColor": WHITE}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    }]

    if rows:
        block_start = 0
        for i in range(1, len(results_sorted) + 1):
            changed = (
                i == len(results_sorted)
                or results_sorted[i]["signal"] != results_sorted[block_start]["signal"]
            )
            if changed:
                signal = results_sorted[block_start]["signal"]
                color = SIGNAL_COLOR.get(signal)
                if color:
                    requests.append({
                        "repeatCell": {
                            "range": {"sheetId": ws.id,
                                       "startRowIndex": block_start + 1, "endRowIndex": i + 1,
                                       "startColumnIndex": 0, "endColumnIndex": num_cols},
                            "cell": {"userEnteredFormat": {"backgroundColor": color}},
                            "fields": "userEnteredFormat.backgroundColor",
                        }
                    })
                block_start = i

    sh.batch_update({"requests": requests})


# ============================== MAIN ========================================
def main():
    parser = argparse.ArgumentParser(description="NSE inside bar scanner")
    parser.add_argument("--timeframe", choices=["daily", "hourly"], required=True)
    parser.add_argument("--symbols-file", default=None,
                         help="Optional text file, one NSE symbol per line. "
                              "Default: fetch NSE's list, or fall back to Nifty 50.")
    parser.add_argument("--credentials", default="credentials.json",
                         help="Path to the Google service account JSON key.")
    parser.add_argument("--spreadsheet-id", default=SPREADSHEET_ID)
    parser.add_argument("--chunk-size", type=int, default=50,
                         help="Symbols per Yahoo Finance batch request.")
    parser.add_argument("--max-price", type=float, default=CONFIG["default_max_price"],
                         help="Only report stocks priced at or below this (Rs). Default 500.")
    parser.add_argument("--min-avg-volume", type=int, default=CONFIG["min_avg_volume"],
                         help="Only report stocks whose prior average volume meets this threshold. "
                              "Default 100000 shares.")
    parser.add_argument("--min-turnover-lakhs", type=float, default=CONFIG["min_turnover_lakhs"],
                         help="Only report stocks whose average rupee turnover meets this threshold, "
                              "in lakhs. Default 50.")
    parser.add_argument("--force", action="store_true",
                         help="Run an hourly scan even outside market hours (for testing).")
    args = parser.parse_args()

    if args.timeframe == "hourly" and not args.force and not is_market_hours():
        print("Outside NSE market hours (9:15 AM-3:30 PM IST, Mon-Fri) — skipping. "
              "Use --force to run anyway.")
        return

    tf = TIMEFRAMES[args.timeframe]
    all_symbols = load_symbols(args.symbols_file)
    print(f"Loaded {len(all_symbols)} symbols. Filtering to price <= Rs {args.max_price}...")

    symbols = prefilter_by_price(all_symbols, args.max_price, args.chunk_size)
    print(
        f"{len(symbols)} symbols at or below Rs {args.max_price}. "
        f"Skipping stocks below {args.min_avg_volume:,} avg volume or "
        f"Rs {args.min_turnover_lakhs:g}L average turnover. "
        f"Running {args.timeframe} scan..."
    )

    if not symbols:
        print("No symbols left after the price filter — nothing to scan.")
        return

    gc = get_gspread_client(args.credentials)

    results = []
    chunks = list(chunked(symbols, args.chunk_size))
    for idx, chunk in enumerate(chunks, 1):
        try:
            data = fetch_batch(chunk, tf["interval"], tf["period"])
        except Exception as e:
            print(f"  batch {idx}/{len(chunks)} failed to fetch: {e}", file=sys.stderr)
            continue

        for sym in chunk:
            try:
                df = extract_symbol_df(data, sym, len(chunk))
                r = analyze_symbol(sym, df, tf, args.min_avg_volume, args.min_turnover_lakhs)
                if r:
                    results.append(r)
            except Exception:
                continue

        print(f"  batch {idx}/{len(chunks)} done — {len(results)} signals so far")
        time.sleep(1)  # polite pacing between batches

    write_results(gc, args.spreadsheet_id, tf["sheet_name"], results)
    print(f"Done. {len(results)} signals written to '{tf['sheet_name']}'.")


if __name__ == "__main__":
    main()