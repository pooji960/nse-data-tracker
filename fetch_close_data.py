
import sys
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import yfinance as yf


# =========================
# SETTINGS
# =========================

EXCEL_PATH = os.path.join(
    os.environ["GITHUB_WORKSPACE"],
    "Data_Analysis_daily.xlsx"
)

SYMBOL_SHEET = "Live Data"
OUT_SHEET = "Close History"

MIN_VOLUME = 2_000_000

HEADERS = [
    "TRADE DATE",
    "FETCHED AT",
    "SYMBOL",
    "PREV. CLOSE",
    "CLOSE",
    "DAY %CHNG",
    "VOLUME (shares)",
    "VALUE (₹ Crores)"
]

IST = ZoneInfo("Asia/Kolkata")


# =========================
# GET SYMBOLS
# =========================

def get_symbols(wb):
    ws = wb[SYMBOL_SHEET]

    symbols = []

    for row in ws.iter_rows(
        min_row=2,
        max_col=1,
        values_only=True
    ):
        if row[0]:
            symbols.append(str(row[0]).strip())

    return symbols


# =========================
# CHECK DUPLICATE SESSION
# =========================

def already_saved(ws, trade_date, session):
    for row in ws.iter_rows(
        min_row=2,
        max_col=2,
        values_only=True
    ):
        old_date = row[0]
        old_time = row[1]

        if not old_date or not old_time:
            continue

        try:
            old_date = pd.Timestamp(old_date).date()

            if isinstance(old_time, datetime):
                old_hour = old_time.hour
                old_minute = old_time.minute
            else:
                old_time = pd.Timestamp(old_time)
                old_hour = old_time.hour
                old_minute = old_time.minute

            if old_date == trade_date:

                # Morning session = 09:15
                if session == "MORNING":
                    if old_hour == 9 and old_minute == 15:
                        return True

                # Evening session = 16:15
                if session == "EVENING":
                    if old_hour == 16 and old_minute == 15:
                        return True

        except Exception:
            continue

    return False


# =========================
# MORNING 9:15 DATA
# =========================

def fetch_morning_data(symbols, trade_date):

    print("Fetching exact 9:15 AM IST data...")

    tickers = [symbol + ".NS" for symbol in symbols]

    data = yf.download(
        tickers,
        period="1d",
        interval="1m",
        group_by="ticker",
        threads=True,
        progress=False,
        auto_adjust=False
    )

    rows = []

    for symbol, ticker in zip(symbols, tickers):

        try:
            d = data[ticker].dropna(subset=["Close"])

            if d.empty:
                continue

            # Convert Yahoo timestamps to IST
            if d.index.tz is None:
                d.index = d.index.tz_localize("UTC")

            d.index = d.index.tz_convert("Asia/Kolkata")

            # Exact 9:15 AM candle
            morning = d[
                (d.index.date == trade_date) &
                (d.index.hour == 9) &
                (d.index.minute == 15)
            ]

            if morning.empty:
                print(f"No 9:15 data: {symbol}")
                continue

            candle = morning.iloc[0]

            close = float(candle["Close"])
            volume = int(candle["Volume"])

            # Previous trading day's close
            previous_rows = d[d.index.date < trade_date]

            if previous_rows.empty:
                continue

            prev_close = float(previous_rows["Close"].iloc[-1])

            if volume <= MIN_VOLUME:
                continue

            fetched_at = datetime(
                trade_date.year,
                trade_date.month,
                trade_date.day,
                9,
                15,
                0,
                tzinfo=IST
            )

            day_change = (
                (close - prev_close) / prev_close * 100
            )

            value_crores = (
                close * volume / 1e7
            )

            rows.append([
                datetime(
                    trade_date.year,
                    trade_date.month,
                    trade_date.day
                ),
                fetched_at.replace(tzinfo=None),
                symbol,
                round(prev_close, 2),
                round(close, 2),
                round(day_change, 2),
                volume,
                round(value_crores, 2)
            ])

        except Exception as e:
            print(f"Morning error {symbol}: {e}")

    return rows


# =========================
# EVENING 4:15 DATA
# =========================

def fetch_evening_data(symbols, trade_date):

    print("Fetching 4:15 PM closing data...")

    tickers = [symbol + ".NS" for symbol in symbols]

    data = yf.download(
        tickers,
        period="5d",
        interval="1d",
        group_by="ticker",
        threads=True,
        progress=False,
        auto_adjust=False
    )

    rows = []

    for symbol, ticker in zip(symbols, tickers):

        try:
            d = data[ticker].dropna(subset=["Close"])

            if len(d) < 2:
                continue

            last_date = pd.Timestamp(d.index[-1]).date()

            if last_date != trade_date:
                continue

            close = float(d["Close"].iloc[-1])
            prev_close = float(d["Close"].iloc[-2])
            volume = int(d["Volume"].iloc[-1])

            if volume <= MIN_VOLUME:
                continue

            fetched_at = datetime(
                trade_date.year,
                trade_date.month,
                trade_date.day,
                16,
                15,
                0,
                tzinfo=IST
            )

            day_change = (
                (close - prev_close) / prev_close * 100
            )

            value_crores = (
                close * volume / 1e7
            )

            rows.append([
                datetime(
                    trade_date.year,
                    trade_date.month,
                    trade_date.day
                ),
                fetched_at.replace(tzinfo=None),
                symbol,
                round(prev_close, 2),
                round(close, 2),
                round(day_change, 2),
                volume,
                round(value_crores, 2)
            ])

        except Exception as e:
            print(f"Evening error {symbol}: {e}")

    return rows


# =========================
# MAIN
# =========================

def main():

    now_ist = datetime.now(IST)

    trade_date = now_ist.date()

    print("IST time:", now_ist)
    print("Trade date:", trade_date)

    wb = openpyxl.load_workbook(EXCEL_PATH)

    symbols = get_symbols(wb)

    if not symbols:
        print("No symbols found in Live Data sheet.")
        return

    print(f"Found {len(symbols)} symbols.")

    if OUT_SHEET in wb.sheetnames:
        ws = wb[OUT_SHEET]
    else:
        ws = wb.create_sheet(OUT_SHEET)

    # Add header ONLY if sheet is empty
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        ws.append(HEADERS)

    # =========================
    # MORNING
    # =========================

    if not already_saved(
        ws,
        trade_date,
        "MORNING"
    ):

        morning_rows = fetch_morning_data(
            symbols,
            trade_date
        )

        for row in morning_rows:
            ws.append(row)

        print(
            f"Morning 9:15 AM: added "
            f"{len(morning_rows)} stocks."
        )

    else:
        print("Morning 9:15 AM data already exists.")


    # ========================


