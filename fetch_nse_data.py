#!/usr/bin/env python3
"""
fetch_nse_data.py — pulls live NSE quotes for your tracked symbols and logs
them into Data_Analysis__daily.xlsx.

Run TWICE a day, via cron:
  09:15 IST  ->  python fetch_nse_data.py morning
  16:15 IST  ->  python fetch_nse_data.py close

MORNING run (09:15):
  - Fetches current live quotes (nse_eq) for every symbol in the
    'Data - MidCap', 'Range 100 to 1000' and 'Small Cap 100 to 1000' sheets.
  - Merges in DELIVERY % from the previous trading day's NSE
    "Bhavcopy with Delivery" report (sec_bhavdata_full), which is only
    published after that day's market close, so by 09:15 it's ready.
  - Filters to VOLUME (live) > 20,00,000 shares.
  - Overwrites 'Live Data' and 'NSE >20L' with today's snapshot.
  - Appends one row per symbol to 'Daily History' (TRADE DATE = previous
    trading day, since that's whose CLOSE/DELIVERY % this really is).

CLOSE run (16:15, ~45 min after the 15:30 market close):
  - Re-fetches live quotes (now at their final/closing values).
  - Filters to VOLUME (live) > 20,00,000 shares.
  - Appends one row per symbol to 'Close History' (TRADE DATE = today).
    No DELIVERY % column here — that day's bhavcopy isn't published yet.

Install once:
    pip install nsepython openpyxl requests

Edit WORKBOOK_PATH below to point at your actual file, then add to crontab
(see the bottom of this file for the exact lines).
"""

import sys
import time
import datetime
import requests
import openpyxl

# --------------------------------------------------------------------------
# CONFIG — edit this to the real path of your workbook on your machine
# --------------------------------------------------------------------------
WORKBOOK_PATH = "Data_Analysis__daily.xlsx"

VOLUME_THRESHOLD = 2_000_000  # 20 lakh shares
SYMBOL_SHEETS = ["Data - MidCap", "Range 100 to 1000", "Small Cap 100 to 1000"]

MASTER_SYMBOL_COL = "SYMBOL"  # header text to locate the symbol column in each sheet


# --------------------------------------------------------------------------
# Symbol universe
# --------------------------------------------------------------------------
def load_symbol_universe(wb):
    """Read every SYMBOL out of the reference sheets, de-duplicated, in order."""
    symbols = []
    seen = set()
    for sheet_name in SYMBOL_SHEETS:
        if sheet_name not in wb.sheetnames:
            print(f"  [warn] sheet '{sheet_name}' not found, skipping")
            continue
        ws = wb[sheet_name]
        header_row = None
        symbol_col = None
        for row in ws.iter_rows(min_row=1, max_row=5):
            for cell in row:
                if cell.value == MASTER_SYMBOL_COL:
                    header_row = cell.row
                    symbol_col = cell.column
                    break
            if header_row:
                break
        if not header_row:
            print(f"  [warn] no SYMBOL header found in '{sheet_name}', skipping")
            continue
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            val = row[symbol_col - 1].value
            if val and val not in seen:
                seen.add(val)
                symbols.append(val)
    return symbols


# --------------------------------------------------------------------------
# Live quote fetch (nsepython)
# --------------------------------------------------------------------------
def fetch_live_quotes(symbols, pause=0.6):
    """
    Returns dict: symbol -> {
        last_price, prev_close, pct_change, day_high, day_low,
        volume, w52_high, w52_low, status
    }
    Uses nsepython.nse_eq, which itself handles the NSE cookie/session dance.
    A small pause between calls avoids getting rate-limited/blocked.
    """
    from nsepython import nse_eq  # imported here so the script still loads without it

    out = {}
    for i, sym in enumerate(symbols, 1):
        try:
            data = nse_eq(sym)
            price_info = data.get("priceInfo", {})
            week_hl = price_info.get("weekHighLow", {})
            intraday = price_info.get("intraDayHighLow", {})
            out[sym] = {
                "last_price": price_info.get("lastPrice"),
                "prev_close": price_info.get("previousClose"),
                "pct_change": price_info.get("pChange"),
                "day_high": intraday.get("max"),
                "day_low": intraday.get("min"),
                "volume": data.get("marketDeptOrderBook", {})
                              .get("tradeInfo", {})
                              .get("totalTradedVolume"),
                "w52_high": week_hl.get("max"),
                "w52_low": week_hl.get("min"),
                "status": "OK",
            }
        except Exception as e:
            out[sym] = {"status": f"ERROR: {e}"}
        if i % 25 == 0:
            print(f"  fetched {i}/{len(symbols)}...")
        time.sleep(pause)
    return out


# --------------------------------------------------------------------------
# Previous day's delivery % (NSE bhavcopy-with-delivery report)
# --------------------------------------------------------------------------
def previous_trading_date(ref_date=None):
    """Naive previous-weekday lookup. Doesn't account for NSE holidays —
    if the previous day was a holiday, re-run with --trade-date to override."""
    d = ref_date or datetime.date.today()
    d -= datetime.timedelta(days=1)
    while d.weekday() >= 5:  # Sat/Sun
        d -= datetime.timedelta(days=1)
    return d


def fetch_delivery_pct(trade_date):
    """
    Downloads NSE's full bhavcopy-with-delivery CSV for trade_date and
    returns dict: symbol -> delivery_pct (float).
    """
    url = (
        "https://nsearchives.nseindia.com/products/content/"
        f"sec_bhavdata_full_{trade_date.strftime('%d%m%Y')}.csv"
    )
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=20)
    resp.raise_for_status()

    lines = resp.text.splitlines()
    header = [h.strip() for h in lines[0].split(",")]
    sym_i = header.index("SYMBOL")
    series_i = header.index("SERIES")
    deliv_i = header.index(" DELIV_PER") if " DELIV_PER" in header else header.index("DELIV_PER")

    out = {}
    for line in lines[1:]:
        cols = [c.strip() for c in line.split(",")]
        if len(cols) <= max(sym_i, series_i, deliv_i):
            continue
        if cols[series_i] != "EQ":
            continue
        try:
            out[cols[sym_i]] = float(cols[deliv_i])
        except ValueError:
            continue
    return out


# --------------------------------------------------------------------------
# Sheet writers
# --------------------------------------------------------------------------
def write_live_data_sheet(wb, quotes, symbols):
    ws = wb["Live Data"]
    ws.delete_rows(2, ws.max_row)  # keep header row 1, clear old rows
    r = 2
    for sym in symbols:
        q = quotes.get(sym, {"status": "MISSING"})
        ws.cell(r, 1, sym)
        ws.cell(r, 2, q.get("last_price"))
        ws.cell(r, 3, q.get("prev_close"))
        ws.cell(r, 4, q.get("pct_change"))
        ws.cell(r, 5, q.get("day_high"))
        ws.cell(r, 6, q.get("day_low"))
        ws.cell(r, 7, q.get("volume"))
        ws.cell(r, 8, q.get("w52_high"))
        ws.cell(r, 9, q.get("w52_low"))
        ws.cell(r, 10, q.get("status"))
        r += 1


def filtered_rows(quotes, symbols):
    """symbols with live volume > threshold, sorted by volume desc."""
    rows = []
    for sym in symbols:
        q = quotes.get(sym)
        if not q or q.get("status") != "OK":
            continue
        vol = q.get("volume")
        if vol and vol > VOLUME_THRESHOLD:
            rows.append((sym, q))
    rows.sort(key=lambda x: x[1]["volume"], reverse=True)
    return rows


def write_nse_20l_sheet(wb, rows, trade_date, fetched_at):
    ws = wb["NSE >20L"]
    ws.delete_rows(1, ws.max_row)
    ws.cell(1, 2,
            f"Trade date: {trade_date:%d-%b-%Y}   |   "
            f"Fetched: {fetched_at:%d-%b-%Y %H:%M}   |   "
            f"Filter: Volume > 2,000,000 shares   |   Stocks: {len(rows)}")
    headers = ["SYMBOL", "PREV. CLOSE", "CLOSE", "DAY %CHNG",
               "VOLUME (shares)", "VALUE (₹ Crores)", "DELIVERY %"]
    for j, h in enumerate(headers, start=2):
        ws.cell(2, j, h)
    r = 3
    for sym, q in rows:
        value_cr = (q["volume"] * (q["last_price"] or 0)) / 1e7
        ws.cell(r, 2, sym)
        ws.cell(r, 3, q.get("prev_close"))
        ws.cell(r, 4, q.get("last_price"))
        ws.cell(r, 5, q.get("pct_change"))
        ws.cell(r, 6, q.get("volume"))
        ws.cell(r, 7, round(value_cr, 2))
        ws.cell(r, 8, q.get("delivery_pct"))
        r += 1


def append_history(wb, sheet_name, rows, trade_date, fetched_at, include_delivery):
    ws = wb[sheet_name]
    r = ws.max_row + 1
    start_col = 2 if sheet_name == "Daily History" else 1  # matches existing layout
    for sym, q in rows:
        value_cr = (q["volume"] * (q["last_price"] or 0)) / 1e7
        vals = [trade_date, fetched_at, sym, q.get("prev_close"), q.get("last_price"),
                q.get("pct_change"), q.get("volume"), round(value_cr, 2)]
        if include_delivery:
            vals.append(q.get("delivery_pct"))
        for j, v in enumerate(vals):
            ws.cell(r, start_col + j, v)
        r += 1


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def run(mode):
    now = datetime.datetime.now()
    print(f"[{now:%Y-%m-%d %H:%M:%S}] starting '{mode}' fetch")

    wb = openpyxl.load_workbook(WORKBOOK_PATH)
    symbols = load_symbol_universe(wb)
    print(f"  {len(symbols)} symbols in universe")

    quotes = fetch_live_quotes(symbols)

    if mode == "morning":
        trade_date = previous_trading_date()
        try:
            deliv = fetch_delivery_pct(trade_date)
            for sym, q in quotes.items():
                if sym in deliv:
                    q["delivery_pct"] = deliv[sym]
        except Exception as e:
            print(f"  [warn] delivery % fetch failed: {e}")
    else:
        trade_date = now.date()

    write_live_data_sheet(wb, quotes, symbols)
    rows = filtered_rows(quotes, symbols)
    write_nse_20l_sheet(wb, rows, trade_date, now)

    if mode == "morning":
        append_history(wb, "Daily History", rows, trade_date, now, include_delivery=True)
    else:
        append_history(wb, "Close History", rows, trade_date, now, include_delivery=False)

    wb.save(WORKBOOK_PATH)
    print(f"  {len(rows)} stocks above 20L volume — saved.")


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in ("morning", "close"):
        print("Usage: python fetch_nse_data.py [morning|close]")
        sys.exit(1)
    run(sys.argv[1])

# --------------------------------------------------------------------------
# Crontab (edit paths, then: crontab -e)
# --------------------------------------------------------------------------
# 15 9  * * 1-5  /usr/bin/python3 /home/you/fetch_nse_data.py morning >> /home/you/nse_fetch.log 2>&1
# 15 16 * * 1-5  /usr/bin/python3 /home/you/fetch_nse_data.py close   >> /home/you/nse_fetch.log 2>&1
