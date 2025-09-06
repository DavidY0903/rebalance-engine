# -*- coding: utf-8 -*-
"""
Rebalance Engine v1.4 — Revised — Full Version
Enhancements over v1.0:
  • Sheet2 (rebalance recommendation) adds 4 columns:
      - Action, Order Type, Limit Price, Limit Justification
  • All original logic and structure preserved from v1.0
  • Input/output interface and Excel formatting kept unchanged
"""

import os
import re
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

import time
import random

import requests
import yfinance.shared as yfs
import yfinance


try:
    import yfinance as yf
except ImportError:
    raise ImportError("Please install yfinance first:  pip install yfinance")

# ---------- Optional knobs ----------
MIN_CASH_BUFFER = 1.00      # try to deploy cash until remainder <= $1 (only used if partial shares)
PRICE_BIAS = 0.0            # 0.0 = neutral; >0 favors higher-priced names, <0 favors cheaper

# --- Setup a browser-like session for yfinance ---
_browser_session = requests.Session()
_browser_session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/114.0.0.0 Safari/537.36"
    )
})

# --- Force all yfinance internals to reuse our session ---
yfs._requests = _browser_session
yfinance.utils.get_yf_session = lambda: _browser_session


# ---------------------------
# Helpers
# ---------------------------



def _safe_get(t, retries=3):
    for attempt in range(retries):
        try:
            info = t.get_info()
            price = info.get("regularMarketPrice")
            if price is not None and np.isfinite(price):
                return float(price)
        except Exception as e:
            if "Too Many Requests" in str(e) and attempt < retries - 1:
                wait = (2 ** attempt) + random.random()
                print(f"Rate limited, retrying in {wait:.1f}s...")
                time.sleep(wait)
                continue
        break
    return np.nan

def _norm_key(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(s).lower())


def _pick_input_filename() -> str:
    try:
        return input_filename  # type: ignore
    except NameError:
        pass

    from tkinter import Tk, filedialog
    root = Tk()
    root.withdraw()
    fn = filedialog.askopenfilename(
        title="Select Input Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.update()
    if not fn:
        raise FileNotFoundError("No input file selected.")
    return fn


def _extract_config(raw_df: pd.DataFrame) -> dict:
    cfg = raw_df.iloc[2:7, [1, 2]].copy()
    cfg.iloc[:, 0] = cfg.iloc[:, 0].astype(str).str.strip().str.lower()
    kv = { _norm_key(k): v for k, v in zip(cfg.iloc[:, 0], cfg.iloc[:, 1]) }

    def getv(name, pct=False, yesno=False):
        k = _norm_key(name)
        if k not in kv:
            raise ValueError(f"Missing config: '{name}'")
        v = kv[k]
        if yesno:
            return str(v).strip().lower() in ("yes", "y", "true", "1")
        if pct:
            return float(str(v).replace("%", "")) / 100.0
        return float(str(v).replace(",", ""))

    return {
        "Cash Contribution": getv("cash contribution"),
        "Upper Bound": getv("upper bound", pct=True),
        "Lower Bound": getv("lower bound", pct=True),
        "Relax Limit": getv("relax limit", pct=True),
        "Allow Partial Shares": getv("allow partial shares", yesno=True),
    }


def _extract_section(raw_df: pd.DataFrame, title: str, required_cols: list[str]) -> pd.DataFrame:
    m = raw_df[raw_df.iloc[:, 0].astype(str).str.contains(title, na=False)]
    if m.empty:
        raise ValueError(f"❌ Section not found: {title}")
    start = m.index[0]
    header = raw_df.iloc[start + 1].tolist()
    for col in required_cols:
        if col not in header:
            raise ValueError(f"❌ Missing column '{col}' in section '{title}'.")
    data = []
    for i in range(start + 2, len(raw_df)):
        row = raw_df.iloc[i].tolist()
        if pd.isna(row[0]) or ("Portfolio" in str(row[0]) and i > start + 2):
            break
        data.append(row[:len(header)])
    return pd.DataFrame(data, columns=header).dropna(how="all").reset_index(drop=True)


# def _fetch_price(ticker: str) -> float:
#     try:
#         t = yf.Ticker(ticker)
#         p = t.info.get("regularMarketPrice", None)
#         if p is not None and np.isfinite(p):
#             return float(p)
#         hist = t.history(period="5d")
#         if not hist.empty:
#             return float(hist["Close"].iloc[-1])
#     except Exception:
#         pass
#     return np.nan

def _fetch_prices(tickers: list[str], max_retries: int = 3, backoff: float = 1.5) -> dict[str, float]:
    """
    Fetch live prices for multiple tickers at once with retries and backoff.
    Uses a patched requests session with a browser-like User-Agent
    to avoid empty JSON issues inside Docker.
    """
    prices = {t: np.nan for t in tickers}

    # Step 1: try fast_info individually
    for t in tickers:
        for attempt in range(max_retries):
            try:
                p = yf.Ticker(t).fast_info.last_price
                if p and np.isfinite(p):
                    prices[t] = float(p)
                    break
            except Exception as e:
                print(f"Fast info error for {t} (attempt {attempt+1}):", e)
            time.sleep(backoff ** attempt)  # exponential backoff

    # Step 2: batch fetch missing tickers via yf.download
    missing = [t for t, v in prices.items() if np.isnan(v)]
    if missing:
        for attempt in range(max_retries):
            try:
                data = yf.download(
                    missing,
                    period="5d",
                    interval="1d",
                    progress=False,
                    group_by="ticker",
                    threads=False,  # safer inside Docker
                )
                if isinstance(data.columns, pd.MultiIndex):
                    for t in missing:
                        try:
                            hist = yf.Ticker(t).history(period="5d", interval="1d")
                            if not hist.empty:
                                prices[t] = float(hist["Close"].iloc[-1])
                        except Exception as e:
                            print(f"History fallback failed for {t}:", e)
                else:
                    try:
                        prices[missing[0]] = float(data["Close"].iloc[-1])
                    except Exception:
                        pass
                break  # success → break retry loop
            except Exception as e:
                print(f"Batch fetch error (attempt {attempt+1}):", e)
                time.sleep(backoff ** attempt)

    # Step 3: fallback to history() for any remaining
    still_missing = [t for t, v in prices.items() if np.isnan(v)]
    for t in still_missing:
        try:
            hist = yf.Ticker(t).history(period="5d", interval="1d")
            if not hist.empty:
                prices[t] = float(hist["Close"].iloc[-1])
        except Exception as e:
            print(f"History fallback failed for {t}:", e)

    print("[fetch_prices] Final results:", {t: prices[t] for t in tickers})
    return prices


def _fetch_price(ticker: str) -> float:
    """
    Wrapper for single ticker — just calls _fetch_prices with one item.
    """
    return _fetch_prices([ticker]).get(ticker, np.nan)

def run_engine(input_file: str, output_file: str) -> str:
    """
    Public entrypoint for HTTP service.
    Always writes the Excel output to the given output_file.
    """
    global input_filename
    input_filename = input_file

    main(input_file, output_file)
    return output_file


def main(input_path: str = None, output_path: str = None):

    """
    Core rebalance logic.
    - input_path: Excel input file path
    - output_path: Excel output file path
    If not provided, falls back to GUI prompt and dynamic filename (for local runs).
    """
    if input_path is None:
        input_path = _pick_input_filename()

    # If API gives us output_path, we use it directly.
    # Otherwise, fall back to your old timestamp-based filename.
    if output_path is None:
        hk_now = datetime.utcnow() + timedelta(hours=8)
        timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")
        try:
            user_df__tmp = pd.read_excel(input_path, sheet_name="Input File", header=None)
            _a1 = user_df__tmp.iloc[0, 0] if user_df__tmp.shape[0] > 0 and user_df__tmp.shape[1] > 0 else ""
            if isinstance(_a1, str) and _a1.strip():
                user_name = _a1.strip()
            else:
                _m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
                user_name = _m.group(1) if _m else "User"
        except Exception:
            _m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
            user_name = _m.group(1) if _m else "User"
        output_filename = f"rebalance recommendation ({user_name})({timestamp_str}).xlsx"
        output_path = os.path.join(os.path.dirname(os.path.abspath(input_path)), output_filename)

    xl = pd.ExcelFile(input_path)
    sheet_name = next((s for s in xl.sheet_names if s.strip().lower() == "current profile"), xl.sheet_names[0])
    raw_input_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    for section in ["Target Portfolio & Weight", "Current Portfolio & Quantity"]:
        if not raw_input_df.iloc[:, 0].astype(str).str.contains(section, na=False).any():
            raise ValueError(f"❌ Missing section: {section}")

    # Metadata
    hk_now = datetime.utcnow() + timedelta(hours=8)
    timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")
    # FIX #1: extract user name from "Input File" sheet (fallback to filename)
    try:
        user_df__tmp = pd.read_excel(input_path, sheet_name="Input File", header=None)
        _a1 = user_df__tmp.iloc[0, 0] if user_df__tmp.shape[0] > 0 and user_df__tmp.shape[1] > 0 else ""
        _name = _a1.strip() if isinstance(_a1, str) and _a1.strip() else ""
        if not _name:
            for _r in range(min(10, user_df__tmp.shape[0])):
                for _c in range(min(5, user_df__tmp.shape[1])):
                    _v = user_df__tmp.iloc[_r, _c]
                    if isinstance(_v, str) and _v.strip():
                        _name = _v.strip()
                        break
                if _name:
                    break
        if not _name:
            _m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
            _name = _m.group(1) if _m else "User"
        user_name = _name
    except Exception:
        _m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
        user_name = _m.group(1) if _m else "User"
    output_filename = f"rebalance recommendation ({user_name})({timestamp_str}).xlsx"
    # output_path = os.path.join(os.path.dirname(os.path.abspath(input_path)), output_filename)

    # Config and section extraction
    config = _extract_config(raw_input_df)
    target_df = _extract_section(raw_input_df, "Target Portfolio & Weight", ["Ticker", "Weight"])[["Ticker", "Weight"]]
    portfolio_df = _extract_section(raw_input_df, "Current Portfolio & Quantity", ["Ticker", "Quantity"])[["Ticker", "Quantity"]]

    target_df.columns = ["Ticker", "Target Weight (%)"]
    target_df["Target Weight (%)"] = (
        target_df["Target Weight (%)"]
        .astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False).astype(float)
    )
    if target_df["Target Weight (%)"].sum() > 1.5:
        target_df["Target Weight (%)"] /= 100.0

    portfolio_df.columns = ["Ticker", "Current Shares"]
    portfolio_df["Current Shares"] = portfolio_df["Current Shares"].astype(str).str.replace(",", "", regex=False).astype(float)

    # Merge + Prices
    df = pd.merge(target_df, portfolio_df, on="Ticker", how="outer").fillna({"Target Weight (%)": 0.0, "Current Shares": 0.0})
    # df["Price per Share"] = df["Ticker"].apply(_fetch_price)
    all_tickers = df["Ticker"].dropna().astype(str).tolist()
    price_map = _fetch_prices(all_tickers)
    df["Price per Share"] = df["Ticker"].map(price_map)

    if df["Price per Share"].isnull().any():
        missing = df[df["Price per Share"].isnull()]["Ticker"].tolist()
        raise ValueError(f"❌ Missing live prices for: {missing}")

    df["Market Value"] = df["Current Shares"] * df["Price per Share"]
    initial_total_assets = df["Market Value"].sum() + float(config["Cash Contribution"])
    df["Current Weight"] = np.where(initial_total_assets > 0, df["Market Value"] / initial_total_assets, 0.0)

    df["Lower Bound"] = df["Target Weight (%)"] * config["Lower Bound"]
    df["Upper Bound"] = df["Target Weight (%)"] * config["Upper Bound"]

    def _classify(row):
        if row["Target Weight (%)"] == 0:
            return "Sell All"
        if row["Current Weight"] > row["Upper Bound"]:
            return "Overweight"
        if row["Current Weight"] < row["Lower Bound"]:
            return "Underweight"
        return "Within Range"

    df["Status"] = df.apply(_classify, axis=1)
    df["_LiveWeight"] = df["Current Weight"].copy()

    # --- Sell Overweight and Sell All ---
    cash = float(config["Cash Contribution"])
    adjustments: list[tuple[str, float, float]] = []

    ow = df[df["Status"] == "Overweight"].copy()
    for _, row in ow.iterrows():
        target_val = row["Upper Bound"] * initial_total_assets
        sell_amt = max(row["Market Value"] - target_val, 0.0)
        if sell_amt <= 0 or row["Price per Share"] <= 0:
            continue
        shares_to_sell = (sell_amt / row["Price per Share"]) if config["Allow Partial Shares"] else int(sell_amt // row["Price per Share"])
        if shares_to_sell > 0:
            value = shares_to_sell * row["Price per Share"]
            adjustments.append((row["Ticker"], -shares_to_sell, -value))
            cash += value

    sa = df[df["Status"] == "Sell All"].copy()
    for _, row in sa.iterrows():
        if row["Current Shares"] > 0 and row["Price per Share"] > 0:
            shares_to_sell = row["Current Shares"]
            value = shares_to_sell * row["Price per Share"]
            adjustments.append((row["Ticker"], -shares_to_sell, -value))
            cash += value

    # --- Buy Underweight (future-safe cap) ---
    df["Market Value"] = df["Current Shares"] * df["Price per Share"]
    order = (
        (np.minimum(df["Upper Bound"], df["Target Weight (%)"] * config["Relax Limit"]) - df["_LiveWeight"])
        .clip(lower=0)
        .sort_values(ascending=False)
        .index
    )

    for i in order:
        if cash <= 0:
            break
        price = float(df.at[i, "Price per Share"])
        if price <= 0 or df.at[i, "Status"] != "Underweight":
            continue

        tgt = float(df.at[i, "Target Weight (%)"])
        MV = float(df.at[i, "Market Value"])
        TA = df["Market Value"].sum() + cash
        cap_w = min(df.at[i, "Upper Bound"], tgt * config["Relax Limit"])
        spend_max = max((cap_w * TA - MV) / (1.0 + cap_w), 0.0)
        alloc = min(cash, spend_max)
        sh = (alloc / price) if config["Allow Partial Shares"] else int(alloc // price)
        if sh <= 0:
            continue

        spend = sh * price
        df.at[i, "Current Shares"] += sh
        df.at[i, "Market Value"] = df.at[i, "Current Shares"] * price
        cash -= spend
        adjustments.append((df.at[i, "Ticker"], sh, spend))

        TA_new = df["Market Value"].sum() + cash
        df["_LiveWeight"] = np.where(TA_new > 0, df["Market Value"] / TA_new, 0.0)

    # --- Final Buy (gap/price with optional price bias) ---
    if cash > 0:
        cap_w_series = np.minimum(df["Upper Bound"], df["Target Weight (%)"] * config["Relax Limit"])
        gap_w = (cap_w_series - df["_LiveWeight"]).clip(lower=0)
        prices = df["Price per Share"].replace(0, np.nan)
        median_price = float(prices.median(skipna=True)) if prices.notna().any() else 1.0
        bias_factor = (prices / median_price).pow(PRICE_BIAS).fillna(1.0)
        score = np.where(df["Price per Share"] > 0, gap_w / df["Price per Share"], 0.0) * bias_factor
        order2 = pd.Series(score).sort_values(ascending=False).index

        for i in order2:
            if cash <= 0:
                break
            price = float(df.at[i, "Price per Share"])
            if price <= 0 or float(df.at[i, "Target Weight (%)"]) <= 0:
                continue

            MV = float(df.at[i, "Market Value"])
            TA = df["Market Value"].sum() + cash
            cap_w = min(float(df.at[i, "Upper Bound"]), float(df.at[i, "Target Weight (%)"]) * config["Relax Limit"])
            spend_max = max((cap_w * TA - MV) / (1.0 + cap_w), 0.0)
            alloc = min(cash, spend_max)
            sh = (alloc / price) if config["Allow Partial Shares"] else int(alloc // price)
            if sh <= 0:
                continue

            spend = sh * price
            df.at[i, "Current Shares"] += sh
            df.at[i, "Market Value"] = df.at[i, "Current Shares"] * price
            cash -= spend
            adjustments.append((df.at[i, "Ticker"], sh, spend))

            TA_new = df["Market Value"].sum() + cash
            df["_LiveWeight"] = np.where(TA_new > 0, df["Market Value"] / TA_new, 0.0)

    # --- Merge Adjustments ---
    action_df = pd.DataFrame(adjustments, columns=["Ticker", "Shares to Buy/Sell", "Actual Trade $"])
    if not action_df.empty:
        action_df = action_df.groupby("Ticker").sum(numeric_only=True).reset_index()
    else:
        action_df = pd.DataFrame(columns=["Ticker", "Shares to Buy/Sell", "Actual Trade $"])

    df = df.merge(action_df, on="Ticker", how="left").fillna({"Shares to Buy/Sell": 0, "Actual Trade $": 0})
    df["Final Shares"] = df["Current Shares"] + df["Shares to Buy/Sell"]
    df["Final Value"] = df["Final Shares"] * df["Price per Share"]
    final_total_assets = df["Final Value"].sum() + cash
    df["Final Weight (Raw)"] = np.where(final_total_assets > 0, df["Final Value"] / final_total_assets, 0.0)
    df["Final Weight"] = df["Final Weight (Raw)"].map("{:.2%}".format)

    def action_label(r):
        if r["Target Weight (%)"] == 0:
            return "Sell All"
        elif r["Shares to Buy/Sell"] > 0:
            return "Buy"
        elif r["Shares to Buy/Sell"] < 0:
            return "Sell"
        return "Hold"

    df["Action"] = df.apply(action_label, axis=1)
    df["% Deviation vs Target"] = ((df["Final Value"] / final_total_assets - df["Target Weight (%)"]) * 100).round(2)
    df["Balance ($)"] = ""

    # --- v1.3 additions ---
    df["Order Type"] = np.where(df["Action"].isin(["Buy", "Sell"]), "Limit", "")
    df["Limit Price"] = np.where(df["Action"] == "Buy",
                                 df["Price per Share"] + 0.01,
                                 np.where(df["Action"] == "Sell",
                                          df["Price per Share"] - 0.01, ""))
    df["Limit Price"] = df["Limit Price"].apply(lambda x: round(x, 2) if isinstance(x, float) else x)
    df["Limit Justification"] = np.where(df["Action"] == "Buy", "Ask + $0.01",
                                         np.where(df["Action"] == "Sell", "Bid - $0.01", ""))

    # --- Export to Excel ---
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        pd.read_excel(input_path, sheet_name=sheet_name, header=None).to_excel(
            writer, index=False, header=False, sheet_name="current profile"
        )
        # FIX #2: insert row index column as first column

        df.insert(0, "#", range(1, len(df) + 1))

        

        # FIX #3: append Total row at end of table

        total_row = {

            "#": "Total",

            "Target Weight (%)": "100.00%",

            "Actual Trade $": round(pd.to_numeric(df.get("Actual Trade $", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(), 2),

            "Balance ($)": round(float(cash), 2),

            "Final Weight": "100.00%",

            "Market Value": round(pd.to_numeric(df.get("Market Value", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(), 2),

            "Final Value": round(pd.to_numeric(df.get("Final Value", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(), 2),

        }

        df.loc[len(df)] = total_row
        # === Bugfix/Improvements for Sheet 2 ===
        # 1) Fix Current Weight = Market Value / sum(Market Value)
        try:
            _mv_sum = pd.to_numeric(df.get("Market Value", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
            if _mv_sum > 0:
                df["Current Weight"] = pd.to_numeric(df.get("Market Value"), errors="coerce").fillna(0) / _mv_sum
            else:
                df["Current Weight"] = 0.0
        except Exception:
            df["Current Weight"] = 0.0

        # Ensure 'Sell All' shows under Action column (not outside)
        try:
            _status_str = df.get("Status", pd.Series([""]*len(df))).astype(str).str.lower()
            df.loc[_status_str.str.contains("sell all"), "Action"] = "Sell All"
        except Exception:
            pass

        # --- Keep Total row out of trading logic ---
        total_mask = df.get("#", pd.Series([""]*len(df))).astype(str).str.lower().eq("total")
        main_df = df.loc[~total_mask].copy()
        total_df = df.loc[total_mask].copy()
        df = main_df

        # === v1.4: Market vs Limit policy + auto Limit Price and Order Instruction ===
        px = pd.to_numeric(df.get("Price per Share", pd.Series(dtype=float)), errors="coerce").fillna(0.0)

        tier_a = set(["SPY","QQQ","IVV","VTI","VOO","IWM","EFA","EEM"])
        tier_b = set(["TLT","LQD","HYG","GLD","XLK","XLF","XLE","USMV","BND"])

        def pick_epsilon_bps(t):
            tt = str(t).upper()
            if tt in tier_a: return 2
            if tt in tier_b: return 5
            return 15

        eps_bps = df.get("Ticker", pd.Series([""]*len(df))).apply(pick_epsilon_bps).astype(int)
        eps = eps_bps / 10000.0

        action_lc = df.get("Action", pd.Series([""]*len(df))).astype(str).str.lower()
        is_buy  = action_lc.eq("buy") | action_lc.eq("buy all")
        is_sell = action_lc.eq("sell") | action_lc.eq("sell all")
        is_hold = action_lc.eq("hold")

        very_liquid = df.get("Ticker", pd.Series([""]*len(df))).astype(str).str.upper().isin(list(tier_a))

        order_type = []
        for i in range(len(df)):
            if is_hold.iloc[i]:
                order_type.append("")
            else:
                order_type.append("Market" if very_liquid.iloc[i] else "Limit")
        df["Order Type"] = order_type

        buy_limit  = (px * (1.0 + eps)).round(2)
        sell_limit = (px * (1.0 - eps)).round(2)
        limit_price = pd.Series([None]*len(df))
        limit_price[is_buy]  = buy_limit[is_buy]
        limit_price[is_sell] = sell_limit[is_sell]
        limit_price = limit_price.where(pd.Series(order_type) == "Limit", other=pd.NA)
        df["Limit Price"] = limit_price

        # Limit Justification (short form, no 'Auto limit:')
        just = pd.Series([""]*len(df))
        just[is_buy]  = "last + " + eps_bps.astype(str) + " bps (buy)"
        just[is_sell] = "last - " + eps_bps.astype(str) + " bps (sell)"
        just = just.where(pd.Series(order_type) == "Limit", "")
        df["Limit Justification"] = just

        # Order Instruction (short form)
        instr = pd.Series([""]*len(df))
        instr[is_hold] = "No action"
        instr[(pd.Series(order_type) == "Market") & is_buy]  = "Market order"
        instr[(pd.Series(order_type) == "Market") & is_sell] = "Market order"
        instr[(pd.Series(order_type) == "Limit") & is_buy]  = "refresh next rebalance if unfilled"
        instr[(pd.Series(order_type) == "Limit") & is_sell] = "refresh next rebalance if unfilled"
        instr[action_lc.eq("sell all")] = "Liquidation"
        df["Order Instruction"] = instr

        # Column order priority
        preferred_cols = ["#", "Ticker", "Action", "Shares to Buy/Sell", "Order Type", "Limit Price", "Limit Justification", "Order Instruction"]
        ordered_cols = [c for c in preferred_cols if c in df.columns] + [c for c in df.columns if c not in preferred_cols]
        df = df[ordered_cols]

        # Row ordering
        _act = df.get("Action", pd.Series([""] * len(df)))
        _rank = _act.map(lambda x: 0 if str(x).lower()=="sell" or str(x).lower()=="sell all" else (1 if str(x).lower()=="buy" else 2))
        _mv = pd.to_numeric(df.get("Market Value", pd.Series(dtype=float)), errors="coerce").fillna(0)
        df = df.assign(_ActionRank=_rank, _MVsort=_mv).sort_values(by=["_ActionRank", "_MVsort"], ascending=[True, False]).drop(columns=["_ActionRank", "_MVsort"]).reset_index(drop=True)

        # Renumber #
        if "#" in df.columns:
            df["#"] = range(1, len(df) + 1)

        # Format Shares to Buy/Sell
        if "Shares to Buy/Sell" in df.columns:
            df["Shares to Buy/Sell"] = pd.to_numeric(df["Shares to Buy/Sell"], errors="coerce").fillna(0).round(2)

        # --- Append Total row back and blank out order fields ---
        df = pd.concat([df, total_df], ignore_index=True)
        if "#" in df.columns:
            df.loc[df.index[-1], "#"] = "Total"
        for col in ["Action","Order Type","Limit Price","Limit Justification","Order Instruction"]:
            if col in df.columns:
                df.loc[df.index[-1], col] = ""

        # Ensure all NaN/inf in Total row are written as empty strings
        df = df.replace([np.nan, np.inf, -np.inf], "")

        df.to_excel(writer, index=False, sheet_name="rebalance recommendation")

        # Cosmetic: full-width top border above Total row + bold
        try:
            wb = writer.book
            ws = writer.sheets["rebalance recommendation"]
            total_row_idx = len(df)  # after concat
            ncols = len(df.columns)
            total_fmt = wb.add_format({"bold": True, "top": 2})
            ws.set_row(total_row_idx, None, total_fmt)
            for c in range(ncols):
                ws.write(total_row_idx, c, df.iloc[-1, c], total_fmt)
        except Exception as _e:
            print("Formatting error (Total row):", _e)

# FIX #4: write trade summary below the table on same sheet
        try:
            mask_rows = df["#"] != "Total"
        except Exception:
            mask_rows = pd.Series([True]*len(df))
        buy_count = int((df.loc[mask_rows, "Shares to Buy/Sell"] > 0).sum()) if "Shares to Buy/Sell" in df.columns else 0
        sell_count = int((df.loc[mask_rows, "Shares to Buy/Sell"] < 0).sum()) if "Shares to Buy/Sell" in df.columns else 0
        cash_initial = float(config.get("Cash Contribution", 0))
        cash_remaining = float(cash)
        cash_used = round(cash_initial - cash_remaining, 2)
        trade_summary = pd.DataFrame({
            "Summary": ["Buy Count", "Sell Count", "Cash Used", "Cash Remaining"],
            "Value": [buy_count, sell_count, cash_used, cash_remaining]
        })
        start_row = len(df) + 3
        trade_summary.to_excel(writer, index=False, sheet_name="rebalance recommendation", startrow=start_row)


        workbook = writer.book
        worksheet = writer.sheets["rebalance recommendation"]
        yellow = workbook.add_format({'bg_color': '#FFFF00'})
        green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

        col_names = df.columns.tolist()
        for col in ["Shares to Buy/Sell", "Action"]:
            if col in col_names:
                ci = col_names.index(col)
                worksheet.set_column(ci, ci, None, yellow)

        if "Action" in col_names:
            ac = col_names.index("Action")
            worksheet.conditional_format(1, ac, len(df), ac, {
                'type': 'text', 'criteria': 'containing', 'value': 'Buy', 'format': green})
            worksheet.conditional_format(1, ac, len(df), ac, {
                'type': 'text', 'criteria': 'containing', 'value': 'Sell', 'format': red})

    print(f"✅ Export complete: {output_path}")


if __name__ == "__main__":
    main(None, None)

