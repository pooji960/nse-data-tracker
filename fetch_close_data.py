"""
Daily 4:15 PM (IST) closing-data fetch -> appends to 'Close History' sheet
in Data_Analysis_.xlsx

Setup:  pip install yfinance openpyxl pandas
Edit EXCEL_PATH below, then schedule with cron (see bottom).
"""
import sys
from datetime import datetime

import openpyxl
import pandas as pd
import yfinance as yf

import os

EXCEL_PATH = os.path.join(
    os.environ["GITHUB_WORKSPACE"],
    "Data_Analysis_daily.xlsx"
)  # <-- change this
SYMBOL_SHEET = "Live Data"        # symbols are read from column A of this sheet
OUT_SHEET = "Close History"
MIN_VOLUME = 2_000_000            # 20 lakh shares
HEADERS = ["TRADE DATE", "FETCHED AT", "SYMBOL", "PREV. CLOSE", "CLOSE",
           "DAY %CHNG", "VOLUME (shares)", "VALUE (₹ Crores)"]


def main():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    symbols = [r[0] for r in wb[SYMBOL_SHEET].iter_rows(min_row=2, max_col=1, values_only=True) if r[0]]
    tickers = [s + ".NS" for s in symbols]

    data = yf.download(tickers, period="5d", interval="1d",
                       group_by="ticker", threads=True, progress=False)

    trade_date = data.index[-1].date()
    if trade_date != datetime.now().date():
        print("Market closed today (holiday/weekend) - nothing to save.")
        return

    ws = wb[OUT_SHEET] if OUT_SHEET in wb.sheetnames else wb.create_sheet(OUT_SHEET)
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        ws.append(HEADERS)
    else:  # avoid duplicate run for the same date
        for r in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if r[0] and pd.Timestamp(r[0]).date() == trade_date:
                print("Today's data already saved.")
                return

    fetched_at = datetime.now().replace(microsecond=0)
    rows = []
    for sym, t in zip(symbols, tickers):
        try:
            d = data[t].dropna(subset=["Close"])
            if len(d) < 2:
                continue
            close, prev, vol = float(d["Close"].iloc[-1]), float(d["Close"].iloc[-2]), int(d["Volume"].iloc[-1])
            if vol <= MIN_VOLUME:
                continue
            rows.append([pd.Timestamp(trade_date).to_pydatetime(), fetched_at, sym,
                         round(prev, 2), round(close, 2),
                         round((close - prev) / prev * 100, 2), vol,
                         round(close * vol / 1e7, 2)])
        except Exception:
            continue

    rows.sort(key=lambda x: x[6], reverse=True)   # highest volume first
    for r in rows:
        ws.append(r)
    wb.save(EXCEL_PATH)      # Excel file must be CLOSED while this runs
    print(f"Saved {len(rows)} stocks for {trade_date}")


if __name__ == "__main__":
    sys.exit(main())

# ---- Cron (Mon-Fri, 4:15 PM IST) ----
# crontab -e   and add:
# CRON_TZ=Asia/Kolkata
# 15 16 * * 1-5 /usr/bin/python3 /full/path/to/fetch_close_data.py >> /full/path/to/close_log.txt 2>&1
