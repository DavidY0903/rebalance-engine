# ===============================================================
# 📘 Rebalance Engine v1.5 — Final Full Fixed Edition v3
# Version: 1.5-final-full-v3 (2025-10-10)
# ---------------------------------------------------------------
# ✅ Restored pointer-based file picker (manual selection)
# ✅ Dynamic output naming (<input>_output.xlsx)
# ✅ Robust I/O + section parsing (from 9-27 baseline)
# ✅ Live Yahoo Finance bid/ask/last pricing
# ✅ Bilingual Excel sheets and formatting
# ✅ Correct handling of zero / negative cash
# ===============================================================

import os, re, time
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime
from tkinter import Tk, filedialog
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# ---------- Input Handling (Pointer File Picker) ----------
def _pick_input_filename() -> str:
    print("📂 Please select the Excel input file...")
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select the Rebalance Input File",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.update()
    if not file_path:
        raise FileNotFoundError("❌ No file selected. Please choose an input Excel file.")
    print(f"✅ Selected input file: {os.path.basename(file_path)}")
    return file_path

# ---------- Helpers ----------
def _norm_key(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def _extract_config(raw_df: pd.DataFrame) -> dict:
    cfg = raw_df.iloc[2:7, [1, 2]].copy()
    cfg.iloc[:, 0] = cfg.iloc[:, 0].astype(str).str.strip().str.lower()
    kv = {_norm_key(k): v for k, v in zip(cfg.iloc[:, 0], cfg.iloc[:, 1])}

    def getv(name, pct=False, yesno=False, default=None):
        k = _norm_key(name)
        if k not in kv:
            return default
        v = kv[k]
        if yesno:
            return str(v).strip().lower() in ("yes", "y", "true", "1")
        if pct:
            return float(str(v).replace("%", "")) / 100.0
        try:
            return float(str(v).replace(",", ""))
        except:
            return default

    return {
        "Cash Contribution": getv("cash contribution", default=0.0),
        "Upper Bound": getv("upper bound", pct=True, default=1.0),
        "Lower Bound": getv("lower bound", pct=True, default=0.0),
        "Relax Limit": getv("relax limit", pct=True, default=1.0),
        "Allow Partial Shares": getv("allow partial shares", yesno=True, default=False),
    }

def _extract_section(raw_df: pd.DataFrame, title: str, required_cols: list[str]) -> pd.DataFrame:
    m = raw_df[raw_df.iloc[:, 0].astype(str).str.contains(title, na=False)]
    if m.empty:
        raise ValueError(f"❌ Section not found: {title}")
    start = m.index[0]
    header = [str(x).strip() for x in raw_df.iloc[start + 1].tolist()]

    header_map = {}
    for h in header:
        if not h or h.lower() == "nan":
            continue
        if "ticker" in h.lower():
            header_map[h] = "Ticker"
        elif "weight" in h.lower():
            header_map[h] = "Weight"
        elif "quantity" in h.lower() or "shares" in h.lower():
            header_map[h] = "Quantity"
        else:
            header_map[h] = h

    data = []
    for i in range(start + 2, len(raw_df)):
        row = raw_df.iloc[i].tolist()
        if pd.isna(row[0]) or ("Portfolio" in str(row[0]) and i > start + 2):
            break
        data.append(row[:len(header)])

    df = pd.DataFrame(data, columns=header).dropna(how="all").reset_index(drop=True)
    df = df.rename(columns=header_map)
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"❌ Missing column '{col}' in section '{title}'. Found header={header}")
    return df

# ---------- Price Fetch ----------
def safe_fetch_price(ticker):
    try:
        data = yf.Ticker(ticker)
        info = data.info
        ask = info.get("ask") or 0
        bid = info.get("bid") or 0
        last = info.get("regularMarketPrice") or np.nan
        if ask <= 0: ask = last
        if bid <= 0: bid = last
        return float(ask or 0), float(bid or 0), float(last or 0)
    except Exception:
        return np.nan, np.nan, np.nan

# ---------- Excel Styling ----------
header_fill = PatternFill(start_color="FFDDEBF7", end_color="FFDDEBF7", fill_type="solid")
buy_fill = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
sell_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")

thin_top = Border(top=Side(style="thin"))
full_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

def format_sheet(ws):
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(name="新細明體", bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = full_border

    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(10, max_len + 2)

def style_action(ws):
    for r in range(2, ws.max_row + 1):
        val = ws[f"B{r}"].value or ws[f"C{r}"].value
        if val in ["Buy", "買進"]:
            for c in ws[r]:
                c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif val in ["Sell", "賣出", "Sell All"]:
            for c in ws[r]:
                c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# ---------- Main ----------
def main():
    print("📘 Rebalance Engine v1.5 — Final Full Fixed Edition v3")
    input_path = _pick_input_filename()
    xl = pd.ExcelFile(input_path)
    sheet_name = next((s for s in xl.sheet_names if s.strip().lower() == "current profile"), xl.sheet_names[0])
    raw_input_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    config = _extract_config(raw_input_df)
    target_df = _extract_section(raw_input_df, "Target Portfolio & Weight", ["Ticker", "Weight"])[["Ticker", "Weight"]]
    port_df = _extract_section(raw_input_df, "Current Portfolio & Quantity", ["Ticker", "Quantity"])[["Ticker", "Quantity"]]

    target_df.columns = ["Ticker", "Target Weight (%)"]
    target_df["Target Weight (%)"] = (
        target_df["Target Weight (%)"].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False)
    )
    target_df["Target Weight (%)"] = pd.to_numeric(target_df["Target Weight (%)"], errors="coerce").fillna(0.0)
    if target_df["Target Weight (%)"].sum() > 1.5:
        target_df["Target Weight (%)"] /= 100.0

    port_df.columns = ["Ticker", "Shares"]
    port_df["Shares"] = pd.to_numeric(
        port_df["Shares"].astype(str).str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0.0)

    df = pd.merge(target_df, port_df, on="Ticker", how="outer").fillna({"Target Weight (%)": 0.0, "Shares": 0.0})

    ask_list, bid_list, last_list = [], [], []
    for t in df["Ticker"]:
        ask, bid, last = safe_fetch_price(t)
        ask_list.append(ask)
        bid_list.append(bid)
        last_list.append(last)
        time.sleep(0.4)
    df["Ask"], df["Bid"], df["Last"] = ask_list, bid_list, last_list

    df["Used Price"] = df["Last"]
    df["Market Value"] = df["Shares"] * df["Used Price"]
    total_value = df["Market Value"].sum() + config["Cash Contribution"]
    df["Target Value"] = df["Target Weight (%)"] * total_value

    # ---------- Allocation math (C-2 LOGIC: proper negative-cash handling) ----------
    df["Diff Value"] = df["Target Value"] - df["Market Value"]

    # Step 1 — Start with standard proportional rebalance
    df["Shares to Buy/Sell"] = (df["Diff Value"] / df["Used Price"]).round(2)

    # Step 2 — If Cash Contribution is negative → override logic
    neg_cash = config["Cash Contribution"] < 0

    if neg_cash:
        needed = abs(config["Cash Contribution"])

        # Sell ONLY from positive-MV assets (ignore target weight)
        df["SellableValue"] = df["Market Value"]

        total_sellable = df["SellableValue"].sum()

        if total_sellable > 0:
            df["ProportionalToSell"] = df["SellableValue"] / total_sellable
            df["ValueToSell"] = df["ProportionalToSell"] * needed
            df["Shares to Buy/Sell"] = -(df["ValueToSell"] / df["Used Price"]).round(2)
        else:
            df["Shares to Buy/Sell"] = 0

    # Step 3 — Clean tiny floats
    df.loc[abs(df["Shares to Buy/Sell"]) < 0.01, "Shares to Buy/Sell"] = 0

    # Step 4 — Recompute Actions cleanly
    df["Action"] = np.where(df["Shares to Buy/Sell"] > 0, "Buy",
                    np.where(df["Shares to Buy/Sell"] < 0, "Sell", "Hold"))

    df.loc[df["Shares"] == 0, "Action"] = "Buy"
    df.loc[(df["Target Weight (%)"] == 0) & (df["Shares"] > 0), "Action"] = "Sell All"

    df["Order Type"] = "Limit"
    df["Limit Price"] = np.where(df["Action"].isin(["Buy", "買進"]), df["Ask"], df["Bid"])
    df["Limit Price"] = df["Limit Price"].fillna(df["Used Price"])

    # ---------- Summary & Output ----------
    buy_cash = (df.loc[df["Action"] == "Buy", "Shares to Buy/Sell"] * df["Used Price"]).sum()
    sell_cash = (-df.loc[df["Action"].isin(["Sell", "Sell All"]), "Shares to Buy/Sell"] * df["Used Price"]).sum()
    cash_remaining = round(sell_cash - buy_cash + config["Cash Contribution"], 2)

    # ----- Updated Summary Block (Bilingual + Styled + Wrap Text) -----
    summary = pd.DataFrame({
    "Summary": [
        "買進支付金額 (Cash Used for Buys)",
        "賣出收入 (Cash Received from Sells)",
        "現金結餘 (Net Cash Result)"
    ],
    "Value": [
        round(buy_cash, 2),
        round(sell_cash, 2),
        cash_remaining
    ]
})

    # ----- Create Rebalance Recommendation (Format B) -----
    cols = ["Ticker", "Target Weight (%)", "Shares", "Used Price",
            "Market Value", "Target Value", "Diff Value",
            "Shares to Buy/Sell", "Action", "Order Type"]
    rebalance_df = df[cols].copy()

    total_row = pd.DataFrame({
        "Ticker": ["Total"],
        "Market Value": [df["Market Value"].sum()],
        "Target Value": [df["Target Value"].sum()],
        "Diff Value": [config["Cash Contribution"]]
    })
    rebalance_df = pd.concat([rebalance_df, total_row], ignore_index=True)

    # ----- Excel Output -----
    input_dir = os.path.dirname(input_path) or os.getcwd()
    input_name = os.path.splitext(os.path.basename(input_path))[0]

    m = re.findall(r"\(.*?\)", input_name)
    tag = "".join(m) if m else ""
    timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")

    output_name = os.path.join(
        input_dir, f"rebalance recommendation {tag} ({timestamp}).xlsx"
    )

    wb = Workbook()
    wb.remove(wb.active)

    sort_order = {"Sell": 0, "Sell All": 0, "Buy": 1, "Hold": 2}
    df_sorted = df.copy()
    df_sorted["sort_key"] = df_sorted["Action"].map(sort_order)
    df_sorted = df_sorted.sort_values(by=["sort_key", "Diff Value"], ascending=[True, False])
    df_sorted = df_sorted.drop(columns=["sort_key"])

    # ----- Action sheet -----
    df_action = df_sorted[["Ticker", "Action", "Shares to Buy/Sell", "Order Type", "Limit Price"]]
    ws_action = wb.create_sheet("action")
    for row in dataframe_to_rows(df_action, index=False, header=True):
        ws_action.append(row)
    format_sheet(ws_action)
    style_action(ws_action)

    # ----- 中文執行 sheet -----
    df_exec = df_action.copy()
    df_exec.columns = ["股票代碼", "動作", "買賣股數", "委託類型", "委託價格"]

    df_exec["動作"] = df_exec["動作"].replace({
        "Buy": "買進",
        "Sell": "賣出",
        "Sell All": "全部賣出",
        "Hold": "持有"
    })
    df_exec["委託類型"] = df_exec["委託類型"].replace({
        "Limit": "限價",
        "Market": "市價"
    })

    ws_exec = wb.create_sheet("執行")
    for row in dataframe_to_rows(df_exec, index=False, header=True):
        ws_exec.append(row)
    format_sheet(ws_exec)
    style_action(ws_exec)

    # ----- Rebalance Recommendation sheet -----
    ordered_rebal = (
        rebalance_df[rebalance_df["Ticker"] != "Total"]
            .set_index("Ticker")
            .reindex(df_action["Ticker"])
            .reset_index()
    )

    total_rebal = rebalance_df[rebalance_df["Ticker"] == "Total"]

    ws_rebal = wb.create_sheet("Rebalance Recommendation")
    for row in dataframe_to_rows(pd.concat([ordered_rebal, total_rebal], ignore_index=True),
                                index=False, header=True):
        ws_rebal.append(row)

    format_sheet(ws_rebal)

    for r in range(2, ws_rebal.max_row + 1):
        ticker_val = str(ws_rebal[f"A{r}"].value or "").strip()
        if ticker_val.lower() == "total":
            for cell in ws_rebal[r]:
                cell.fill = PatternFill(fill_type=None)
                cell.border = Border(top=Side(style="medium"))
            continue

        action_val = ws_rebal[f"I{r}"].value
        if action_val in ("Buy", "買進"):
            for cell in ws_rebal[r]:
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif action_val in ("Sell", "賣出", "Sell All"):
            for cell in ws_rebal[r]:
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    ws_rebal.append([])
    ws_rebal.append([])

        # ----- Styled Summary Block (Bilingual + Auto Wrap + Borders + Align) -----
    for row_idx, row in enumerate(dataframe_to_rows(summary, index=False, header=True), start=1):
        ws_rebal.append(row)

    # Apply styles to the entire summary block
    start_row = ws_rebal.max_row - len(summary)  # first summary row
    end_row = ws_rebal.max_row                   # last summary row

    for r in range(start_row, end_row + 1):
        summary_cell = ws_rebal[f"A{r}"]
        value_cell = ws_rebal[f"B{r}"]

        # Summary column formatting
        summary_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )
        summary_cell.font = Font(name="新細明體", bold=False, size=11)
        summary_cell.border = full_border

        # Value column formatting (right-aligned)
        value_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )
        value_cell.font = Font(name="新細明體", bold=True, size=11)
        value_cell.border = full_border

    # ----- Keep original input file content as last sheet -----
    if "current input file" in wb.sheetnames:
        del wb["current input file"]

    ws_input = wb.create_sheet("current input file")

    for row in dataframe_to_rows(raw_input_df, index=False, header=False):
        ws_input.append(row)

    format_sheet(ws_input)

    wb.save(output_name)
    print(f"✅ Output saved: {output_name}")

if __name__ == "__main__":
    main()
