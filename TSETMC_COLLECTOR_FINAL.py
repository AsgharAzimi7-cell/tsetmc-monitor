# -*- coding: utf-8 -*-
"""
TSETMC LIVE FAST + MARKET INDEX + INTRADAY FIX
================
Lightweight intraday updater for Pydroid 3 / Android.

NORMAL USE
----------
1) Keep this script in the same folder as your latest:
       TSETMC_FINAL_V4_DATA_*.xlsx
2) Run this script once.
   On the first run it imports the historical DAILY data from V4 into a small
   local SQLite database. This is ONE-TIME work.
3) From the second run onward it DOES NOT download the full price history.
   It only downloads today's changing data.
4) Upload only the newly generated:
       TSETMC_LIVE_YYYYMMDD_HHMM.xlsx

Install once:
    pip install requests openpyxl

Persistent local files - DO NOT DELETE:
    TSETMC_LOCAL_HISTORY.db

The light Excel contains:
    SUMMARY          = current market snapshot + comparison with history
    DAILY_260D       = compact last 260 trading sessions for technical analysis
    MONTHLY_ALL      = compact full monthly history for long-term trend/Elliott
    PV_5M_TODAY      = today's real TSETMC trade-price path in 5-minute buckets
    PV_QUALITY       = intraday data-quality check
    CLIENT_CURRENT   = current حقیقی/حقوقی
    BEST_LIMITS      = current order book
    HOLDERS          = cached/current major holders
    CODAL            = cached/recent Codal metadata
    MESSAGES         = cached/recent TSETMC messages
    API_STATUS       = request status

The full V4 history is therefore NOT downloaded on every run.
"""

import csv
import json
import math
import os
import random
import sqlite3
import time
import threading
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from zoneinfo import ZoneInfo


TSETMC_PROXY = os.getenv("TSETMC_PROXY", "").strip()
TSETMC_PROXIES = [x.strip() for x in os.getenv("TSETMC_PROXIES", "").replace(",", ";").split(";") if x.strip()]
if TSETMC_PROXY and TSETMC_PROXY not in TSETMC_PROXIES:
    TSETMC_PROXIES.insert(0, TSETMC_PROXY)
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "").strip()
COLLECTOR_MODE = os.getenv("COLLECTOR_MODE", "0") == "1"
SNAPSHOT_SECONDS = max(15, int(os.getenv("SNAPSHOT_SECONDS", "30")))
HEAVY_REFRESH_SECONDS = max(60, int(os.getenv("HEAVY_REFRESH_SECONDS", "300")))
TEHRAN_TZ = ZoneInfo("Asia/Tehran")
# ============================================================
# SETTINGS
# ============================================================

SYMBOLS = [
    "لابسا",
    "فاراک",
    "پتایر",
    "دروز",
    "وسدید",
    "غشان",
    "کاوه",
]

# Market indices are collected separately from stocks on every run.
# Codes are discovered dynamically from TSETMC; fallbacks are only used if
# the live index list is temporarily unavailable.
INDEX_TARGETS = {
    "شاخص کل": {
        "match": ("شاخصکل",),
        "fallback_codes": ("32097828799138957", "32097828820363860"),
    },
    "شاخص کل هم‌وزن": {
        "match": ("هموزن", "شاخص"),
        "fallback_codes": ("67130298613737946",),
    },
}
INDEX_HISTORY_CACHE_MINUTES = 30
INDEX_CONTEXT_DAYS = 260

BASE_URL = "https://cdn.tsetmc.com/api"

# Leave blank. The program searches for the newest V4 workbook beside itself.
# Or enter an absolute path manually.
BASELINE_XLSX = ""

DB_FILE = "TSETMC_LOCAL_HISTORY.db"

REQUEST_TIMEOUT = 18
MAX_RETRIES = 3
REQUEST_DELAY = (0.04, 0.12)

# Output historical context.
DAILY_CONTEXT_DAYS = 260

# Intraday aggregation.
INTRADAY_BUCKET_MINUTES = 5
MARKET_START_HHMMSS = 90000
MARKET_END_HHMMSS = 123000

# Slow-changing/current extras are cached to make repeated intraday runs fast.
HOLDERS_CACHE_MINUTES = 45
CODAL_CACHE_MINUTES = 30
MESSAGES_CACHE_MINUTES = 15

# Number of recent Codal items requested when cache expires.
CODAL_TOP = 12

# Core live endpoints.
FETCH_INTRADAY = True
FETCH_BEST_LIMITS = True
FETCH_CLIENT_TYPE = True

# Optional but useful. Cache prevents repeated downloading.
FETCH_HOLDERS = True
FETCH_CODAL = True
FETCH_MESSAGES = True

# Intraday source: real TSETMC trade history. This is the authoritative source
# for PRICE x TIME. Volume coverage may be lower than official cumulative volume,
# so current ClosingPriceInfo remains authoritative for total daily volume/value.
TRY_HEAVY_TRADE_FALLBACK = True

# If there is no V4 file on first run, allow a one-time history download.
# After bootstrap it is never used again unless DB is deleted.
ALLOW_ONE_TIME_API_BOOTSTRAP = True


# ============================================================
# PATHS
# ============================================================

try:
    APP_DIR = Path(__file__).resolve().parent
except Exception:
    APP_DIR = Path.cwd()

DB_PATH = APP_DIR / DB_FILE


# ============================================================
# HTTP
# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Linux; Android 16) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/140.0 Mobile Safari/537.36"
    ),
    "Accept": "application/json,text/plain,*/*",
    "Accept-Language": "fa-IR,fa;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.tsetmc.com/",
}

session = requests.Session()
session.headers.update(HEADERS)
_proxy_index = 0

def _proxy_label(proxy):
    if not proxy:
        return "DIRECT"
    # Avoid exposing credentials in logs.
    return proxy.rsplit("@", 1)[-1]

def activate_proxy(index=0):
    global _proxy_index
    if not TSETMC_PROXIES:
        session.proxies.clear()
        return None
    _proxy_index = index % len(TSETMC_PROXIES)
    proxy = TSETMC_PROXIES[_proxy_index]
    session.proxies.clear()
    session.proxies.update({"http": proxy, "https": proxy})
    return proxy

def rotate_proxy():
    if len(TSETMC_PROXIES) <= 1:
        return TSETMC_PROXIES[0] if TSETMC_PROXIES else None
    return activate_proxy(_proxy_index + 1)

# Only TSETMC traffic goes through an Iran proxy. Telegram bypasses it.
activate_proxy(0)

API_STATUS = []


def blocked_response(text):
    if not text:
        return False
    t = text.lower()
    markers = [
        "<html",
        "<!doctype html",
        "general error detected",
        "دسترسی شما",
        "مسدود",
    ]
    return any(x in t for x in markers)


def api_get(path, symbol="", endpoint="", allow_fail=True):
    url = path if path.startswith("http") else f"{BASE_URL}/{path.lstrip('/')}"
    last_error = None
    started = time.time()

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code in (429, 500, 502, 503, 504):
                raise RuntimeError(f"HTTP {r.status_code}")
            r.raise_for_status()

            if blocked_response(r.text):
                raise RuntimeError("Blocked/non-JSON TSETMC response")

            data = r.json()

            API_STATUS.append({
                "Symbol": symbol,
                "Endpoint": endpoint or path,
                "Status": "OK",
                "HTTP": r.status_code,
                "Attempts": attempt,
                "Seconds": round(time.time() - started, 2),
                "Error": "",
            })

            time.sleep(random.uniform(*REQUEST_DELAY))
            return data

        except Exception as e:
            last_error = e
            if attempt < MAX_RETRIES:
                old_proxy = TSETMC_PROXIES[_proxy_index] if TSETMC_PROXIES else None
                new_proxy = rotate_proxy()
                if new_proxy and new_proxy != old_proxy:
                    print(f"  Proxy retry: {_proxy_label(old_proxy)} -> {_proxy_label(new_proxy)}")
                time.sleep(0.5 * attempt + random.random() * 0.3)

    API_STATUS.append({
        "Symbol": symbol,
        "Endpoint": endpoint or path,
        "Status": "FAILED",
        "HTTP": "",
        "Attempts": MAX_RETRIES,
        "Seconds": round(time.time() - started, 2),
        "Error": str(last_error),
    })

    if allow_fail:
        return None

    raise RuntimeError(f"{endpoint or path}: {last_error}")


# ============================================================
# GENERIC HELPERS
# ============================================================

def fa_normalize(text):
    if text is None:
        return ""
    return (
        str(text).strip()
        .replace("ي", "ی")
        .replace("ك", "ک")
        .replace("\u200c", "")
        .replace(" ", "")
    )


def to_float(v, default=0.0):
    try:
        if v in (None, ""):
            return default
        return float(v)
    except Exception:
        return default


def to_int(v, default=0):
    try:
        if v in (None, ""):
            return default
        return int(float(v))
    except Exception:
        return default


def safe_div(a, b):
    try:
        if b in (None, 0, 0.0):
            return None
        return a / b
    except Exception:
        return None


def pct_change(new, old):
    if old in (None, 0, 0.0) or new is None:
        return None
    return (new / old - 1.0) * 100.0


def unwrap_list(payload, preferred=()):
    if payload is None:
        return []
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return []

    for k in preferred:
        x = payload.get(k)
        if isinstance(x, list):
            return x

    for x in payload.values():
        if isinstance(x, list):
            return x

    return []


def unwrap_dict(payload, preferred=()):
    if payload is None:
        return {}
    if not isinstance(payload, dict):
        return {}

    for k in preferred:
        x = payload.get(k)
        if isinstance(x, dict):
            return x

    return payload


def format_time(v):
    n = to_int(v, 0)
    h = n // 10000
    m = (n // 100) % 100
    s = n % 100

    if not (0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59):
        return ""

    return f"{h:02d}:{m:02d}:{s:02d}"


def minute_of_day(v):
    n = to_int(v, 0)
    return (n // 10000) * 60 + ((n // 100) % 100)


def bucket_time(v):
    mins = minute_of_day(v)
    b = (mins // INTRADAY_BUCKET_MINUTES) * INTRADAY_BUCKET_MINUTES
    return f"{b // 60:02d}:{b % 60:02d}"


def excel_safe(v):
    if v is None:
        return None
    if isinstance(v, (dict, list, tuple)):
        return json.dumps(v, ensure_ascii=False, separators=(",", ":"))
    return v


# ============================================================
# DATABASE
# ============================================================

def connect_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    conn.execute("""
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS daily (
            symbol TEXT NOT NULL,
            date INTEGER NOT NULL,
            inscode TEXT,
            heven INTEGER,
            price_min REAL,
            price_max REAL,
            price_yesterday REAL,
            price_first REAL,
            pclosing REAL,
            last_price REAL,
            trades REAL,
            volume REAL,
            value_rial REAL,
            finalized INTEGER DEFAULT 1,
            PRIMARY KEY(symbol, date)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS intraday_snapshots (
            symbol TEXT NOT NULL,
            date INTEGER NOT NULL,
            heven INTEGER NOT NULL,
            last_price REAL,
            pclosing REAL,
            price_yesterday REAL,
            cum_volume REAL,
            cum_value REAL,
            cum_trades REAL,
            PRIMARY KEY(symbol, date, heven)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS client_intraday (
            symbol TEXT NOT NULL,
            date INTEGER NOT NULL,
            captured_at TEXT NOT NULL,
            market_time TEXT,
            last_price REAL,
            closing_price REAL,
            volume REAL,
            real_buy_count REAL,
            real_sell_count REAL,
            legal_buy_count REAL,
            legal_sell_count REAL,
            real_buy_volume REAL,
            real_sell_volume REAL,
            legal_buy_volume REAL,
            legal_sell_volume REAL,
            avg_real_buy_volume REAL,
            avg_real_sell_volume REAL,
            buyer_power REAL,
            real_net_volume REAL,
            legal_net_volume REAL,
            PRIMARY KEY(symbol, date, captured_at)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS cache (
            cache_key TEXT PRIMARY KEY,
            updated_ts REAL NOT NULL,
            payload TEXT NOT NULL
        )
    """)

    conn.commit()
    return conn


def meta_get(conn, key, default=None):
    row = conn.execute(
        "SELECT value FROM meta WHERE key=?", (key,)
    ).fetchone()
    return row["value"] if row else default


def meta_set(conn, key, value):
    conn.execute(
        """
        INSERT INTO meta(key,value) VALUES(?,?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
        """,
        (key, str(value))
    )
    conn.commit()


def daily_count(conn):
    return conn.execute("SELECT COUNT(*) AS n FROM daily").fetchone()["n"]


def upsert_daily(conn, symbol, inscode, row, finalized):
    d = to_int(row.get("dEven") or row.get("Date"), 0)
    if not d:
        return

    conn.execute("""
        INSERT INTO daily(
            symbol,date,inscode,heven,
            price_min,price_max,price_yesterday,price_first,
            pclosing,last_price,trades,volume,value_rial,finalized
        )
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(symbol,date) DO UPDATE SET
            inscode=excluded.inscode,
            heven=excluded.heven,
            price_min=excluded.price_min,
            price_max=excluded.price_max,
            price_yesterday=excluded.price_yesterday,
            price_first=excluded.price_first,
            pclosing=excluded.pclosing,
            last_price=excluded.last_price,
            trades=excluded.trades,
            volume=excluded.volume,
            value_rial=excluded.value_rial,
            finalized=excluded.finalized
    """, (
        symbol,
        d,
        str(inscode),
        to_int(row.get("hEven"), 0),
        to_float(row.get("priceMin"), 0),
        to_float(row.get("priceMax"), 0),
        to_float(row.get("priceYesterday"), 0),
        to_float(row.get("priceFirst"), 0),
        to_float(row.get("pClosing"), 0),
        to_float(row.get("pDrCotVal"), 0),
        to_float(row.get("zTotTran"), 0),
        to_float(row.get("qTotTran5J"), 0),
        to_float(row.get("qTotCap"), 0),
        int(bool(finalized)),
    ))
    conn.commit()


def get_daily_rows(conn, symbol):
    rows = conn.execute("""
        SELECT * FROM daily
        WHERE symbol=?
        ORDER BY date
    """, (symbol,)).fetchall()

    return [dict(r) for r in rows]


def get_last_daily_row(conn, symbol):
    row = conn.execute("""
        SELECT * FROM daily
        WHERE symbol=? AND volume>0
        ORDER BY date DESC
        LIMIT 1
    """, (symbol,)).fetchone()
    return dict(row) if row else {}


def cache_get(conn, key, ttl_minutes):
    row = conn.execute("""
        SELECT updated_ts,payload FROM cache WHERE cache_key=?
    """, (key,)).fetchone()

    if not row:
        return None

    age_sec = time.time() - float(row["updated_ts"])
    if age_sec > ttl_minutes * 60:
        return None

    try:
        return json.loads(row["payload"])
    except Exception:
        return None


def cache_get_any(conn, key):
    row = conn.execute("""
        SELECT payload FROM cache WHERE cache_key=?
    """, (key,)).fetchone()

    if not row:
        return None

    try:
        return json.loads(row["payload"])
    except Exception:
        return None


def cache_set(conn, key, payload):
    conn.execute("""
        INSERT INTO cache(cache_key,updated_ts,payload)
        VALUES(?,?,?)
        ON CONFLICT(cache_key) DO UPDATE SET
            updated_ts=excluded.updated_ts,
            payload=excluded.payload
    """, (
        key,
        time.time(),
        json.dumps(payload, ensure_ascii=False),
    ))
    conn.commit()


def save_client_intraday(conn, live, captured_at=None):
    """Persist one cumulative client-type snapshot for later intraday analysis."""
    current = live.get("current") or {}
    client = live.get("client") or {}
    symbol = live.get("symbol") or ""
    date_value = to_int(current.get("dEven"), 0)
    if not symbol or not date_value:
        return

    captured_at = captured_at or datetime.now(TEHRAN_TZ).strftime("%H:%M:%S")
    market_time = format_time(current.get("hEven") or current.get("lastHEven"))

    rbv = client_metric(client, "buy_I_Volume")
    rsv = client_metric(client, "sell_I_Volume")
    lbv = client_metric(client, "buy_N_Volume")
    lsv = client_metric(client, "sell_N_Volume")
    rbc = client_metric(client, "buy_I_Count", "buy_CountI")
    rsc = client_metric(client, "sell_I_Count", "sell_CountI")
    lbc = client_metric(client, "buy_N_Count", "buy_CountN")
    lsc = client_metric(client, "sell_N_Count", "sell_CountN")

    avg_buy = safe_div(rbv, rbc)
    avg_sell = safe_div(rsv, rsc)
    bp = safe_div(avg_buy, avg_sell) if avg_buy is not None and avg_sell is not None else None

    conn.execute("""
        INSERT OR REPLACE INTO client_intraday(
            symbol,date,captured_at,market_time,last_price,closing_price,volume,
            real_buy_count,real_sell_count,legal_buy_count,legal_sell_count,
            real_buy_volume,real_sell_volume,legal_buy_volume,legal_sell_volume,
            avg_real_buy_volume,avg_real_sell_volume,buyer_power,
            real_net_volume,legal_net_volume
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        symbol, date_value, captured_at, market_time,
        to_float(current.get("pDrCotVal"), 0),
        to_float(current.get("pClosing"), 0),
        to_float(current.get("qTotTran5J"), 0),
        rbc, rsc, lbc, lsc, rbv, rsv, lbv, lsv,
        avg_buy, avg_sell, bp, rbv-rsv, lbv-lsv,
    ))
    conn.commit()


def get_client_intraday_rows(conn, date_value=None):
    if date_value:
        rows = conn.execute("""
            SELECT * FROM client_intraday WHERE date=?
            ORDER BY captured_at, symbol
        """, (date_value,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT * FROM client_intraday ORDER BY date, captured_at, symbol
        """).fetchall()
    return [dict(r) for r in rows]


# ============================================================
# ONE-TIME HISTORY BOOTSTRAP
# ============================================================

def locate_baseline_file():
    if BASELINE_XLSX:
        p = Path(BASELINE_XLSX)
        if p.exists():
            return p

    patterns = [
        "TSETMC_FINAL_V4_DATA_*.xlsx",
        "TSETMC_FINAL_DATA_*.xlsx",
        "TSETMC_FULL_DATA_*.xlsx",
    ]

    found = []
    for pat in patterns:
        found.extend(APP_DIR.glob(pat))

    found = [
        p for p in found
        if not p.name.startswith("~$")
        and not p.name.startswith("TSETMC_LIVE_")
    ]

    if not found:
        return None

    return max(found, key=lambda p: p.stat().st_mtime)


def import_v4_history(conn, baseline):
    print(f"One-time import from:\n{baseline}")

    wb = load_workbook(
        baseline,
        read_only=True,
        data_only=True,
    )

    if "ALL_DAILY" not in wb.sheetnames:
        wb.close()
        raise RuntimeError("ALL_DAILY sheet not found in baseline workbook.")

    ws = wb["ALL_DAILY"]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it))
    idx = {str(h): i for i, h in enumerate(headers) if h is not None}

    required = [
        "Symbol", "dEven", "insCode",
        "priceMin", "priceMax", "priceYesterday", "priceFirst",
        "pClosing", "pDrCotVal",
        "zTotTran", "qTotTran5J", "qTotCap",
    ]

    missing = [x for x in required if x not in idx]
    if missing:
        wb.close()
        raise RuntimeError(f"Missing columns in ALL_DAILY: {missing}")

    sql = """
        INSERT INTO daily(
            symbol,date,inscode,heven,
            price_min,price_max,price_yesterday,price_first,
            pclosing,last_price,trades,volume,value_rial,finalized
        )
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,1)
        ON CONFLICT(symbol,date) DO UPDATE SET
            inscode=excluded.inscode,
            heven=excluded.heven,
            price_min=excluded.price_min,
            price_max=excluded.price_max,
            price_yesterday=excluded.price_yesterday,
            price_first=excluded.price_first,
            pclosing=excluded.pclosing,
            last_price=excluded.last_price,
            trades=excluded.trades,
            volume=excluded.volume,
            value_rial=excluded.value_rial,
            finalized=1
    """

    batch = []
    count = 0

    for values in it:
        symbol = values[idx["Symbol"]]
        if symbol not in SYMBOLS:
            continue

        date_value = to_int(values[idx["dEven"]], 0)
        volume = to_float(values[idx["qTotTran5J"]], 0)
        close = to_float(values[idx["pClosing"]], 0)

        if not date_value or close <= 0 or volume <= 0:
            continue

        inscode = str(values[idx["insCode"]] or "")

        batch.append((
            symbol,
            date_value,
            inscode,
            to_int(values[idx.get("hEven", -1)], 0) if "hEven" in idx else 0,
            to_float(values[idx["priceMin"]], 0),
            to_float(values[idx["priceMax"]], 0),
            to_float(values[idx["priceYesterday"]], 0),
            to_float(values[idx["priceFirst"]], 0),
            close,
            to_float(values[idx["pDrCotVal"]], 0),
            to_float(values[idx["zTotTran"]], 0),
            volume,
            to_float(values[idx["qTotCap"]], 0),
        ))

        if inscode:
            meta_set(conn, f"inscode:{symbol}", inscode)

        if len(batch) >= 1000:
            conn.executemany(sql, batch)
            conn.commit()
            count += len(batch)
            batch = []

    if batch:
        conn.executemany(sql, batch)
        conn.commit()
        count += len(batch)

    wb.close()

    meta_set(conn, "bootstrap_source", str(baseline))
    meta_set(conn, "bootstrap_time", datetime.now().isoformat(timespec="seconds"))

    print(f"Imported {count:,} historical daily rows into local DB.")


def find_symbol_online(symbol):
    encoded = requests.utils.quote(symbol, safe="")

    p = api_get(
        f"Instrument/GetInstrumentSearch/{encoded}",
        symbol,
        "InstrumentSearch",
        allow_fail=False,
    )

    rows = unwrap_list(p, ("instrumentSearch",))
    norm = fa_normalize(symbol)

    exact = [
        x for x in rows
        if fa_normalize(x.get("lVal18AFC")) == norm
        and str(x.get("insCode", "")).isdigit()
    ]

    choices = exact or [
        x for x in rows
        if str(x.get("insCode", "")).isdigit()
    ]

    if not choices:
        raise RuntimeError(f"No InsCode for {symbol}")

    choices.sort(
        key=lambda x: to_int(x.get("lastDate"), 0),
        reverse=True
    )

    return str(choices[0]["insCode"])


def bootstrap_history_from_api(conn):
    print("No V4 workbook found.")
    print("Running ONE-TIME full-history bootstrap from TSETMC...")

    for symbol in SYMBOLS:
        print(f"  {symbol}")
        inscode = meta_get(conn, f"inscode:{symbol}")
        if not inscode:
            inscode = find_symbol_online(symbol)
            meta_set(conn, f"inscode:{symbol}", inscode)

        p = api_get(
            f"ClosingPrice/GetClosingPriceDailyList/{inscode}/0",
            symbol,
            "ONE_TIME_DailyHistory",
            allow_fail=False,
        )

        rows = unwrap_list(p, ("closingPriceDaily",))

        for r in rows:
            if not isinstance(r, dict):
                continue
            if (
                to_float(r.get("qTotTran5J"), 0) > 0
                and to_float(r.get("pClosing"), 0) > 0
            ):
                upsert_daily(conn, symbol, inscode, r, finalized=True)

    meta_set(conn, "bootstrap_source", "TSETMC API one-time")
    meta_set(conn, "bootstrap_time", datetime.now().isoformat(timespec="seconds"))


def ensure_bootstrap(conn):
    if daily_count(conn) > 0:
        return

    baseline = locate_baseline_file()

    if baseline:
        import_v4_history(conn, baseline)
        return

    if ALLOW_ONE_TIME_API_BOOTSTRAP:
        bootstrap_history_from_api(conn)
        return

    raise RuntimeError(
        "No local history DB and no V4 baseline workbook found."
    )


# ============================================================
# TECHNICAL CALCULATIONS
# ============================================================

def rolling_mean(values, period):
    out = []
    q = []

    for v in values:
        q.append(v)
        if len(q) > period:
            q.pop(0)

        valid = [float(x) for x in q if x is not None]
        out.append(sum(valid) / len(valid) if valid else None)

    return out


def ema(values, period):
    alpha = 2.0 / (period + 1.0)
    out = []
    prev = None

    for v in values:
        if v is None:
            out.append(prev)
            continue

        x = float(v)
        prev = x if prev is None else alpha * x + (1 - alpha) * prev
        out.append(prev)

    return out


def rsi_wilder(values, period=14):
    n = len(values)
    out = [None] * n

    if n <= period:
        return out

    gains = [0.0] * n
    losses = [0.0] * n

    for i in range(1, n):
        d = values[i] - values[i - 1]
        gains[i] = max(d, 0.0)
        losses[i] = max(-d, 0.0)

    avg_gain = sum(gains[1:period + 1]) / period
    avg_loss = sum(losses[1:period + 1]) / period

    out[period] = (
        100.0
        if avg_loss == 0
        else 100.0 - 100.0 / (1.0 + avg_gain / avg_loss)
    )

    for i in range(period + 1, n):
        avg_gain = ((period - 1) * avg_gain + gains[i]) / period
        avg_loss = ((period - 1) * avg_loss + losses[i]) / period

        out[i] = (
            100.0
            if avg_loss == 0
            else 100.0 - 100.0 / (1.0 + avg_gain / avg_loss)
        )

    return out


def enrich_history(raw_rows):
    rows = [
        dict(r) for r in raw_rows
        if to_float(r.get("volume"), 0) > 0
        and to_float(r.get("pclosing"), 0) > 0
    ]

    rows.sort(key=lambda r: to_int(r.get("date"), 0))

    if not rows:
        return []

    # Back-adjust using next-session priceYesterday / previous close.
    multipliers = [1.0] * len(rows)
    mult = 1.0

    for i in range(len(rows) - 1, 0, -1):
        later_y = to_float(rows[i].get("price_yesterday"), 0)
        older_c = to_float(rows[i - 1].get("pclosing"), 0)

        if later_y > 0 and older_c > 0:
            factor = later_y / older_c

            if abs(factor - 1.0) > 0.001:
                mult *= factor

        multipliers[i - 1] = mult

    closes = []
    volumes = []

    for i, r in enumerate(rows):
        m = multipliers[i]

        r["AdjFactor"] = m
        r["AdjOpen"] = (
            to_float(r.get("price_first"), 0) * m
            if to_float(r.get("price_first"), 0) > 0 else None
        )
        r["AdjHigh"] = (
            to_float(r.get("price_max"), 0) * m
            if to_float(r.get("price_max"), 0) > 0 else None
        )
        r["AdjLow"] = (
            to_float(r.get("price_min"), 0) * m
            if to_float(r.get("price_min"), 0) > 0 else None
        )
        r["AdjClose"] = to_float(r.get("pclosing"), 0) * m
        r["AdjLast"] = (
            to_float(r.get("last_price"), 0) * m
            if to_float(r.get("last_price"), 0) > 0 else None
        )

        closes.append(r["AdjClose"])
        volumes.append(to_float(r.get("volume"), 0))

    sma20 = rolling_mean(closes, 20)
    sma50 = rolling_mean(closes, 50)
    sma200 = rolling_mean(closes, 200)

    avgvol5 = rolling_mean(volumes, 5)
    avgvol20 = rolling_mean(volumes, 20)
    avgvol60 = rolling_mean(volumes, 60)

    e12 = ema(closes, 12)
    e26 = ema(closes, 26)

    macd = [
        a - b if a is not None and b is not None else None
        for a, b in zip(e12, e26)
    ]

    signal = ema(macd, 9)

    hist = [
        m - s if m is not None and s is not None else None
        for m, s in zip(macd, signal)
    ]

    rsi = rsi_wilder(closes, 14)

    obv = []
    obv_value = 0.0

    for i in range(len(rows)):
        if i == 0:
            obv_value = volumes[i]
        elif closes[i] > closes[i - 1]:
            obv_value += volumes[i]
        elif closes[i] < closes[i - 1]:
            obv_value -= volumes[i]

        obv.append(obv_value)

    for i, r in enumerate(rows):
        r["Return1D_Pct"] = (
            pct_change(closes[i], closes[i - 1])
            if i > 0 else None
        )
        r["SMA20"] = sma20[i]
        r["SMA50"] = sma50[i]
        r["SMA200"] = sma200[i]
        r["RSI14"] = rsi[i]
        r["MACD"] = macd[i]
        r["MACD_Signal"] = signal[i]
        r["MACD_Hist"] = hist[i]
        r["AvgVol5"] = avgvol5[i]
        r["AvgVol20"] = avgvol20[i]
        r["AvgVol60"] = avgvol60[i]
        r["VolumeRatio20"] = safe_div(volumes[i], avgvol20[i])
        r["OBV"] = obv[i]

        start = max(0, i - 251)
        hvals = [
            rows[j]["AdjHigh"]
            for j in range(start, i + 1)
            if rows[j]["AdjHigh"] is not None
        ]
        lvals = [
            rows[j]["AdjLow"]
            for j in range(start, i + 1)
            if rows[j]["AdjLow"] is not None
        ]

        r["High52W"] = max(hvals) if hvals else None
        r["Low52W"] = min(lvals) if lvals else None

    return rows


def compact_daily_rows(enriched, symbol, limit=260):
    out = []

    for r in enriched[-limit:]:
        out.append({
            "Symbol": symbol,
            "Date": r.get("date"),
            "Finalized": bool(r.get("finalized")),
            "Open": r.get("price_first"),
            "High": r.get("price_max"),
            "Low": r.get("price_min"),
            "Last": r.get("last_price"),
            "Close": r.get("pclosing"),
            "Yesterday": r.get("price_yesterday"),
            "Volume": r.get("volume"),
            "Value_Rial": r.get("value_rial"),
            "Trades": r.get("trades"),
            "AdjOpen": r.get("AdjOpen"),
            "AdjHigh": r.get("AdjHigh"),
            "AdjLow": r.get("AdjLow"),
            "AdjClose": r.get("AdjClose"),
            "Return1D_Pct": r.get("Return1D_Pct"),
            "AvgVol20": r.get("AvgVol20"),
            "AvgVol60": r.get("AvgVol60"),
            "VolumeRatio20": r.get("VolumeRatio20"),
            "RSI14": r.get("RSI14"),
            "MACD": r.get("MACD"),
            "MACD_Signal": r.get("MACD_Signal"),
            "MACD_Hist": r.get("MACD_Hist"),
            "SMA20": r.get("SMA20"),
            "SMA50": r.get("SMA50"),
            "SMA200": r.get("SMA200"),
            "OBV": r.get("OBV"),
            "High52W": r.get("High52W"),
            "Low52W": r.get("Low52W"),
        })

    return out


def monthly_history(enriched, symbol):
    groups = defaultdict(list)

    for r in enriched:
        d = to_int(r.get("date"), 0)
        if d <= 0:
            continue
        ym = d // 100
        groups[ym].append(r)

    out = []

    for ym in sorted(groups):
        xs = groups[ym]
        first = xs[0]
        last = xs[-1]

        highs = [x["AdjHigh"] for x in xs if x.get("AdjHigh") is not None]
        lows = [x["AdjLow"] for x in xs if x.get("AdjLow") is not None]

        out.append({
            "Symbol": symbol,
            "YearMonth": ym,
            "FirstDate": first.get("date"),
            "LastDate": last.get("date"),
            "AdjOpen": first.get("AdjOpen"),
            "AdjHigh": max(highs) if highs else None,
            "AdjLow": min(lows) if lows else None,
            "AdjClose": last.get("AdjClose"),
            "RawClose": last.get("pclosing"),
            "Volume": sum(to_float(x.get("volume"), 0) for x in xs),
            "Value_Rial": sum(to_float(x.get("value_rial"), 0) for x in xs),
            "Trades": sum(to_float(x.get("trades"), 0) for x in xs),
            "TradingDays": len(xs),
        })

    return out


# ============================================================
# CACHED ENDPOINTS
# ============================================================

def cached_endpoint(
    conn,
    cache_key,
    ttl_minutes,
    fetch_callable,
):
    cached = cache_get(conn, cache_key, ttl_minutes)
    if cached is not None:
        return cached, True

    data = fetch_callable()

    if data is not None:
        cache_set(conn, cache_key, data)
        return data, False

    old = cache_get_any(conn, cache_key)
    return (old if old is not None else None), True


# ============================================================
# INTRADAY SNAPSHOT CACHE + 5-MIN BARS
# ============================================================

def insert_intraday_snapshots(conn, symbol, date_value, rows):
    inserted = 0

    for r in rows:
        if not isinstance(r, dict):
            continue

        h = to_int(r.get("hEven"), 0)

        if not (MARKET_START_HHMMSS <= h <= MARKET_END_HHMMSS):
            continue

        last_price = to_float(r.get("pDrCotVal"), 0)
        cum_volume = to_float(r.get("qTotTran5J"), 0)

        # Ignore obvious premarket/empty snapshots.
        if last_price <= 0:
            continue

        conn.execute("""
            INSERT INTO intraday_snapshots(
                symbol,date,heven,last_price,pclosing,
                price_yesterday,cum_volume,cum_value,cum_trades
            )
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(symbol,date,heven) DO UPDATE SET
                last_price=excluded.last_price,
                pclosing=excluded.pclosing,
                price_yesterday=excluded.price_yesterday,
                cum_volume=MAX(intraday_snapshots.cum_volume,excluded.cum_volume),
                cum_value=MAX(intraday_snapshots.cum_value,excluded.cum_value),
                cum_trades=MAX(intraday_snapshots.cum_trades,excluded.cum_trades)
        """, (
            symbol,
            date_value,
            h,
            last_price,
            to_float(r.get("pClosing"), 0),
            to_float(r.get("priceYesterday"), 0),
            cum_volume,
            to_float(r.get("qTotCap"), 0),
            to_float(r.get("zTotTran"), 0),
        ))
        inserted += 1

    conn.commit()
    return inserted


def get_intraday_snapshots(conn, symbol, date_value):
    rows = conn.execute("""
        SELECT * FROM intraday_snapshots
        WHERE symbol=? AND date=?
        ORDER BY heven
    """, (symbol, date_value)).fetchall()

    return [dict(r) for r in rows]


def trades_to_snapshots(trades, current):
    """Convert TSETMC trade rows to cumulative time snapshots.

    TSETMC Trade/GetTradeHistory rows use fields such as:
        hEven    = HHMMSS
        pTran    = trade price
        qTitTran = trade volume

    There can be many trades in the same second. We aggregate those first,
    then build one cumulative snapshot per second so the existing 5-minute
    bar engine can be reused safely.

    Important: GetTradeHistory may not reproduce the official total daily
    volume perfectly. Therefore these snapshots are used for the PRICE/TIME
    path; ClosingPriceInfo remains the reference for official daily totals.
    """
    if not trades:
        return []

    by_second = defaultdict(list)

    for r in trades:
        if not isinstance(r, dict):
            continue

        h = to_int(r.get("hEven", r.get("heven", r.get("time"))), 0)
        if not (MARKET_START_HHMMSS <= h <= MARKET_END_HHMMSS):
            continue

        price = to_float(
            r.get("pTran", r.get("price", r.get("pDrCotVal"))),
            0,
        )
        vol = to_float(
            r.get("qTitTran", r.get("volume", r.get("qTotTran5J"))),
            0,
        )

        if price <= 0:
            continue

        by_second[h].append((price, max(vol, 0.0), r))

    if not by_second:
        return []

    py = to_float(current.get("priceYesterday"), 0)
    cum_vol = 0.0
    cum_val = 0.0
    cum_trades = 0.0
    snapshots = []

    for h in sorted(by_second):
        xs = by_second[h]
        # Keep the last trade price within the second; sum trade volumes.
        last_price = xs[-1][0]
        sec_vol = sum(x[1] for x in xs)
        sec_val = sum(x[0] * x[1] for x in xs)
        sec_trades = float(len(xs))

        cum_vol += sec_vol
        cum_val += sec_val
        cum_trades += sec_trades

        snapshots.append({
            "hEven": h,
            "pDrCotVal": last_price,
            "pClosing": 0,
            "priceYesterday": py,
            "qTotTran5J": cum_vol,
            "qTotCap": cum_val,
            "zTotTran": cum_trades,
        })

    return snapshots


def fetch_intraday_trade_path(inscode, market_date, symbol, current):
    """Fetch real intraday trade-price path from TSETMC.

    Priority:
      1) historical intraday trades for the exact date
      2) compact intraday endpoint for the latest trading day
      3) latest-day raw trades
    """
    sources = [
        (
            f"Trade/GetTradeHistory/{inscode}/{market_date}/false",
            "TradeHistoryToday",
            ("tradeHistory", "trade"),
        ),
        (
            f"Trade/GetTradeIntraDay/{inscode}",
            "TradeIntraDayToday",
            ("tradeIntraDay", "trade"),
        ),
        (
            f"Trade/GetTrade/{inscode}",
            "TradeToday",
            ("trade",),
        ),
    ]

    for path, endpoint, keys in sources:
        p = api_get(path, symbol, endpoint)
        rows = unwrap_list(p, keys)
        snaps = trades_to_snapshots(rows, current)
        if snaps:
            return snaps, endpoint, len(rows)

    return [], "NO_TRADE_SOURCE", 0


def aggregate_cached_intraday(
    snapshots,
    symbol,
    date_value,
    current_volume,
):
    if not snapshots:
        return [], {
            "Symbol": symbol,
            "Date": date_value,
            "SnapshotCount": 0,
            "FirstTime": None,
            "LastTime": None,
            "MaxGapMinutes": None,
            "CapturedCumVolume": 0,
            "CurrentVolume": current_volume,
            "CoveragePct": 0 if current_volume else None,
            "Status": "NO_DATA",
        }

    xs = sorted(snapshots, key=lambda r: to_int(r["heven"], 0))

    # Global cumulative deltas.
    points = []
    prev_vol = 0.0
    prev_val = 0.0
    prev_trades = 0.0
    prev_minute = None
    max_gap = 0

    for r in xs:
        h = to_int(r["heven"], 0)
        this_min = minute_of_day(h)

        if prev_minute is not None:
            max_gap = max(max_gap, this_min - prev_minute)

        cum_vol = to_float(r["cum_volume"], 0)
        cum_val = to_float(r["cum_value"], 0)
        cum_trades = to_float(r["cum_trades"], 0)

        dvol = max(cum_vol - prev_vol, 0)
        dval = max(cum_val - prev_val, 0)
        dtr = max(cum_trades - prev_trades, 0)

        points.append({
            **r,
            "Bucket": bucket_time(h),
            "DeltaVolume": dvol,
            "DeltaValue": dval,
            "DeltaTrades": dtr,
        })

        prev_vol = max(prev_vol, cum_vol)
        prev_val = max(prev_val, cum_val)
        prev_trades = max(prev_trades, cum_trades)
        prev_minute = this_min

    groups = defaultdict(list)
    for p in points:
        groups[p["Bucket"]].append(p)

    out = []
    day_open = None

    for b in sorted(
        groups,
        key=lambda x: int(x[:2]) * 60 + int(x[3:5])
    ):
        ps = groups[b]
        prices = [to_float(x["last_price"], 0) for x in ps]

        if day_open is None:
            day_open = prices[0]

        vol = sum(to_float(x["DeltaVolume"], 0) for x in ps)
        value = sum(to_float(x["DeltaValue"], 0) for x in ps)
        trades = sum(to_float(x["DeltaTrades"], 0) for x in ps)

        last_point = ps[-1]
        py = to_float(last_point.get("price_yesterday"), 0)

        out.append({
            "Symbol": symbol,
            "Date": date_value,
            "TimeBucket": b,
            "StartTime": format_time(ps[0]["heven"]),
            "EndTime": format_time(ps[-1]["heven"]),
            "Open": prices[0],
            "High": max(prices),
            "Low": min(prices),
            "Close": prices[-1],
            "IntervalVolume": vol,
            "IntervalValue_Rial": value,
            "VWAP": safe_div(value, vol),
            "IntervalTrades": trades,
            "CumVolume": to_float(last_point["cum_volume"], 0),
            "CumValue_Rial": to_float(last_point["cum_value"], 0),
            "CumTrades": to_float(last_point["cum_trades"], 0),
            "PriceVsYesterdayPct": pct_change(prices[-1], py),
            "PriceVsDayOpenPct": pct_change(prices[-1], day_open),
        })

    captured = max(to_float(x["cum_volume"], 0) for x in xs)
    coverage = (
        safe_div(captured, current_volume) * 100
        if current_volume > 0 else None
    )

    first_h = to_int(xs[0]["heven"], 0)
    last_h = to_int(xs[-1]["heven"], 0)

    status = "COMPLETE"

    if current_volume <= 0:
        status = "NO_TRADES_YET"
    elif coverage is None:
        status = "UNKNOWN"
    elif coverage < 90:
        status = "PARTIAL"
    elif max_gap > 15:
        status = "GAPS"
    elif coverage < 97:
        status = "GOOD_NOT_FULL"

    quality = {
        "Symbol": symbol,
        "Date": date_value,
        "SnapshotCount": len(xs),
        "FirstTime": format_time(first_h),
        "LastTime": format_time(last_h),
        "MaxGapMinutes": max_gap,
        "CapturedCumVolume": captured,
        "CurrentVolume": current_volume,
        "CoveragePct": coverage,
        "Status": status,
    }

    return out, quality


# ============================================================
# LIVE SYMBOL COLLECTION
# ============================================================

def collect_live_symbol(conn, symbol):
    inscode = meta_get(conn, f"inscode:{symbol}")

    if not inscode:
        inscode = find_symbol_online(symbol)
        meta_set(conn, f"inscode:{symbol}", inscode)

    # --------------------------------------------------------
    # Current price
    # --------------------------------------------------------
    p = api_get(
        f"ClosingPrice/GetClosingPriceInfo/{inscode}",
        symbol,
        "ClosingPriceInfo",
        allow_fail=False,
    )
    current = unwrap_dict(p, ("closingPriceInfo",))

    market_date = to_int(
        current.get("dEven"),
        to_int(datetime.now().strftime("%Y%m%d"), 0)
    )

    current_volume = to_float(current.get("qTotTran5J"), 0)
    current_close = to_float(current.get("pClosing"), 0)

    final_last_date = to_int(current.get("finalLastDate"), 0)

    is_final = (
        market_date > 0
        and final_last_date == market_date
    )

    # Store today's evolving daily bar only after actual trades start.
    if current_volume > 0 and current_close > 0:
        upsert_daily(
            conn,
            symbol,
            inscode,
            current,
            finalized=is_final,
        )

    # --------------------------------------------------------
    # Client type
    # --------------------------------------------------------
    client = {}

    if FETCH_CLIENT_TYPE:
        p = api_get(
            f"ClientType/GetClientType/{inscode}/1/0",
            symbol,
            "ClientTypeCurrent",
        )
        client = unwrap_dict(p, ("clientType",))

    # --------------------------------------------------------
    # Best limits
    # --------------------------------------------------------
    best = []

    if FETCH_BEST_LIMITS:
        p = api_get(
            f"BestLimits/{inscode}",
            symbol,
            "BestLimitsCurrent",
        )
        best = unwrap_list(p, ("bestLimits",))

        for r in best:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    # --------------------------------------------------------
    # Intraday snapshots
    # --------------------------------------------------------
    pv_rows = []
    pv_quality = {
        "Symbol": symbol,
        "Date": market_date,
        "Status": "DISABLED",
    }

    if FETCH_INTRADAY and market_date:
        # FIX: ClosingPrice/GetClosingPriceHistory is daily history and is not
        # the correct source for PRICE x TIME. Use TSETMC Trade endpoints.
        snapshots, intraday_source, raw_trade_rows = fetch_intraday_trade_path(
            inscode,
            market_date,
            symbol,
            current,
        )

        insert_intraday_snapshots(
            conn,
            symbol,
            market_date,
            snapshots,
        )

        cached_snaps = get_intraday_snapshots(
            conn,
            symbol,
            market_date,
        )

        pv_rows, pv_quality = aggregate_cached_intraday(
            cached_snaps,
            symbol,
            market_date,
            current_volume,
        )

        pv_quality["Source"] = intraday_source
        pv_quality["RawTradeRows"] = raw_trade_rows
        pv_quality["PricePathUsable"] = "YES" if pv_rows else "NO"

        for rr in pv_rows:
            rr["Source"] = intraday_source

    # --------------------------------------------------------
    # Holders - cached
    # --------------------------------------------------------
    holders = []

    if FETCH_HOLDERS:
        def fetch_holders():
            pp = api_get(
                f"Shareholder/GetInstrumentShareHolderLast/{inscode}",
                symbol,
                "MajorHolders",
            )
            return unwrap_list(pp, ("shareHolder",))

        holders, _ = cached_endpoint(
            conn,
            f"holders:{symbol}",
            HOLDERS_CACHE_MINUTES,
            fetch_holders,
        )

        holders = holders or []

        for r in holders:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    # --------------------------------------------------------
    # Codal - cached
    # --------------------------------------------------------
    codal = []

    if FETCH_CODAL:
        def fetch_codal():
            pp = api_get(
                f"Codal/GetPreparedDataByInsCode/{CODAL_TOP}/{inscode}",
                symbol,
                "CodalRecent",
            )
            return unwrap_list(pp, ("preparedData", "codalData", "codal"))

        codal, _ = cached_endpoint(
            conn,
            f"codal:{symbol}",
            CODAL_CACHE_MINUTES,
            fetch_codal,
        )

        codal = codal or []

        for r in codal:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    # --------------------------------------------------------
    # Messages - cached
    # --------------------------------------------------------
    messages = []

    if FETCH_MESSAGES:
        def fetch_messages():
            pp = api_get(
                f"Msg/GetMsgByInsCode/{inscode}",
                symbol,
                "MessagesRecent",
            )
            return unwrap_list(pp, ("msg", "tseMessage", "messages"))

        messages, _ = cached_endpoint(
            conn,
            f"messages:{symbol}",
            MESSAGES_CACHE_MINUTES,
            fetch_messages,
        )

        messages = messages or []

        for r in messages:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    return {
        "symbol": symbol,
        "inscode": inscode,
        "current": current,
        "client": client,
        "best": best,
        "holders": holders,
        "codal": codal,
        "messages": messages,
        "pv": pv_rows,
        "pv_quality": pv_quality,
    }


# ============================================================
# MARKET INDICES (TEDPIX + EQUAL WEIGHT)
# ============================================================

def _pick_index_row(rows, label, spec):
    """Pick the desired index from TSETMC's current index list."""
    normalized = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        name = fa_normalize(r.get("lVal30") or r.get("name") or "")
        normalized.append((name, r))

    # شاخص کل: exact match first so we do not accidentally select هم‌وزن.
    if label == "شاخص کل":
        exact_names = {"شاخصکل", "شاخصكل"}
        for name, r in normalized:
            if name in exact_names:
                return r
        for name, r in normalized:
            if "شاخصکل" in name and "هموزن" not in name:
                return r

    # Equal-weight: require both index and equal-weight wording.
    if label == "شاخص کل هم‌وزن":
        for name, r in normalized:
            if "هموزن" in name and "شاخص" in name:
                return r

    # Generic token fallback.
    tokens = tuple(spec.get("match", ()))
    for name, r in normalized:
        if all(tok in name for tok in tokens):
            return r
    return {}


def _index_value(row):
    """Read current/close-like index value across known TSETMC schemas."""
    for k in (
        "xDrNivJIdx004", "indexLastValue", "indexValue", "value",
        "xNivInuClMresIbs", "xNivInuPbMresIbs",
    ):
        if row.get(k) not in (None, ""):
            return to_float(row.get(k), 0)
    return 0.0


def _index_date(row):
    for k in ("dEven", "date", "DEven"):
        v = to_int(row.get(k), 0)
        if v:
            return v
    return 0


def _normalize_index_history(label, code, rows):
    """Keep raw index history fields and add stable helper columns."""
    out = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        rr = dict(r)
        rr["Index"] = label
        rr["IndexCode"] = str(code)
        rr["Date"] = _index_date(r)
        val = _index_value(r)
        if val:
            rr["IndexValue_Normalized"] = val
        out.append(rr)

    out.sort(key=lambda x: to_int(x.get("Date"), 0))
    return out[-INDEX_CONTEXT_DAYS:]


def collect_market_indices(conn):
    """Collect TEDPIX and equal-weight index on every run.

    Design goals:
    - MARKET_INDICES always contains two rows, even if TSETMC is temporarily unavailable.
    - Current values prefer MarketOverview because it exposes both indices in one response.
    - Dedicated Index endpoints are used as a secondary source and for 260-day history.
    - If live endpoints are empty after market close, the latest history row is used as fallback.
    """
    # 1) Current market overview -- most robust source for the two headline indices.
    overview = {}
    try:
        pp = api_get(
            "MarketData/GetMarketOverview/1",
            "MARKET",
            "MarketOverview",
        )
        overview = unwrap_dict(pp, ("marketOverview",)) or {}
    except Exception:
        overview = {}

    # 2) Current full index list -- useful for exact code/name/current high-low/time.
    current_all = []
    try:
        p = api_get(
            "Index/GetIndexB1LastAll/All/1",
            "MARKET",
            "IndexCurrentAll",
        )
        current_all = unwrap_list(p, ("indexB1",)) or []
    except Exception:
        current_all = []

    summary_rows = []
    history_rows = []

    for label, spec in INDEX_TARGETS.items():
        row = _pick_index_row(current_all, label, spec) or {}
        discovered_code = str(row.get("insCode") or "")

        candidate_codes = []
        if discovered_code:
            candidate_codes.append(discovered_code)
        for c in spec.get("fallback_codes", ()):
            c = str(c)
            if c and c not in candidate_codes:
                candidate_codes.append(c)

        # 3) History. Try all known codes until a non-empty series is returned.
        hist = []
        used_code = discovered_code
        for c in candidate_codes:
            def fetch_hist(c=c):
                hh = api_get(
                    f"Index/GetIndexB2History/{c}",
                    label,
                    "IndexHistory",
                )
                return unwrap_list(hh, ("indexB2",))

            try:
                cached, _ = cached_endpoint(
                    conn,
                    f"indexhist:{label}:{c}",
                    INDEX_HISTORY_CACHE_MINUTES,
                    fetch_hist,
                )
                candidate_hist = cached or []
            except Exception:
                candidate_hist = []

            if candidate_hist:
                hist = candidate_hist
                used_code = c
                break

        normalized_hist = _normalize_index_history(label, used_code, hist)
        history_rows.extend(normalized_hist)
        latest_hist = normalized_hist[-1] if normalized_hist else {}

        # Dedicated current-row values.
        value = _index_value(row) if row else 0.0
        change = to_float(row.get("indexChange"), 0) if row else 0.0
        change_pct = None
        if row and row.get("xVarIdxJRfV") not in (None, ""):
            change_pct = to_float(row.get("xVarIdxJRfV"), 0)
        high = to_float(row.get("xPhNivJIdx004"), 0) if row else 0.0
        low = to_float(row.get("xPbNivJIdx004"), 0) if row else 0.0
        heven = row.get("hEven") if row else None
        deven = _index_date(row) if row else 0
        source = "IndexB1LastAll" if value else ""

        # MarketOverview is preferred for the headline current values.
        if overview:
            if label == "شاخص کل":
                ov_value = to_float(overview.get("indexLastValue"), 0)
                ov_change = to_float(overview.get("indexChange"), 0)
            else:
                ov_value = to_float(overview.get("indexEqualWeightedLastValue"), 0)
                ov_change = to_float(overview.get("indexEqualWeightedChange"), 0)

            if ov_value:
                value = ov_value
                change = ov_change
                source = "MarketOverview"
                prev = value - change
                if prev:
                    change_pct = pct_change(value, prev)

        # After-hours/temporary-empty fallback: use the latest history close.
        if not value and latest_hist:
            value = to_float(
                latest_hist.get("IndexValue_Normalized") or
                latest_hist.get("xNivInuClMresIbs") or
                latest_hist.get("xDrNivJIdx004"),
                0,
            )
            # Index history schemas commonly expose daily change in xVarIdxJRfV.
            if latest_hist.get("xVarIdxJRfV") not in (None, ""):
                change_pct = to_float(latest_hist.get("xVarIdxJRfV"), 0)
            if not deven:
                deven = to_int(latest_hist.get("Date"), 0)
            source = "IndexHistoryFallback"

        if change_pct is None and value and change:
            prev = value - change
            if prev:
                change_pct = pct_change(value, prev)

        status = "OK" if value else "NO_DATA"
        summary_rows.append({
            "Index": label,
            "IndexCode": used_code or discovered_code,
            "Value": value if value else None,
            "Change": change if (value or change) else None,
            "ChangePct": change_pct,
            "High": high if high else None,
            "Low": low if low else None,
            "Date": deven or None,
            "Time": heven,
            "Source": source or "TSETMC",
            "Status": status,
        })

    # Stable ordering: total index first, equal-weight second.
    order = {"شاخص کل": 0, "شاخص کل هم‌وزن": 1}
    summary_rows.sort(key=lambda r: order.get(r.get("Index"), 99))
    history_rows.sort(key=lambda r: (order.get(r.get("Index"), 99), to_int(r.get("Date"), 0)))
    return summary_rows, history_rows

# ============================================================
# SUMMARY
# ============================================================

def client_metric(client, *keys):
    for k in keys:
        if k in client and client.get(k) not in (None, ""):
            return to_float(client.get(k), 0)
    return 0.0


def make_summary(conn, live):
    symbol = live["symbol"]
    current = live["current"]
    client = live["client"] or {}

    history_raw = get_daily_rows(conn, symbol)
    enriched = enrich_history(history_raw)

    market_date = to_int(current.get("dEven"), 0)
    live_volume = to_float(current.get("qTotTran5J"), 0)
    live_value = to_float(current.get("qTotCap"), 0)
    live_trades = to_float(current.get("zTotTran"), 0)
    live_last = to_float(current.get("pDrCotVal"), 0)
    live_close = to_float(current.get("pClosing"), 0)
    py = to_float(current.get("priceYesterday"), 0)

    # Previous finalized history for fair volume comparison.
    previous_rows = [
        r for r in enriched
        if to_int(r.get("date"), 0) < market_date
    ]

    prev20 = previous_rows[-20:]
    prev60 = previous_rows[-60:]

    avgvol20_prev = (
        sum(to_float(r.get("volume"), 0) for r in prev20) / len(prev20)
        if prev20 else None
    )

    avgvol60_prev = (
        sum(to_float(r.get("volume"), 0) for r in prev60) / len(prev60)
        if prev60 else None
    )

    tech = enriched[-1] if enriched else {}

    # If today's bar has not traded, technical values come from last traded day.
    last_hist = previous_rows[-1] if previous_rows else tech

    technical_source = tech
    provisional = False

    if live_volume > 0 and enriched and to_int(tech.get("date"), 0) == market_date:
        provisional = not bool(tech.get("finalized"))
    else:
        technical_source = last_hist

    buy_i_vol = client_metric(client, "buy_I_Volume")
    sell_i_vol = client_metric(client, "sell_I_Volume")
    buy_n_vol = client_metric(client, "buy_N_Volume")
    sell_n_vol = client_metric(client, "sell_N_Volume")

    buy_i_count = client_metric(client, "buy_I_Count", "buy_CountI")
    sell_i_count = client_metric(client, "sell_I_Count", "sell_CountI")

    buy_i_value = client_metric(client, "buy_I_Value")
    sell_i_value = client_metric(client, "sell_I_Value")

    real_net_vol = buy_i_vol - sell_i_vol

    real_net_value = (
        buy_i_value - sell_i_value
        if (buy_i_value or sell_i_value)
        else None
    )

    real_net_approx = (
        real_net_vol * live_last
        if live_last > 0 else None
    )

    avg_buy_vol_per_buyer = safe_div(buy_i_vol, buy_i_count)
    avg_sell_vol_per_seller = safe_div(sell_i_vol, sell_i_count)

    buyer_power = (
        safe_div(avg_buy_vol_per_buyer, avg_sell_vol_per_seller)
        if avg_buy_vol_per_buyer is not None
        and avg_sell_vol_per_seller is not None
        else None
    )

    holder_buy = 0.0
    holder_sell = 0.0

    for h in live.get("holders", []):
        if not isinstance(h, dict):
            continue
        ch = to_float(h.get("changeAmount"), 0)
        holder_buy += max(ch, 0)
        holder_sell += min(ch, 0)

    q = live.get("pv_quality", {})

    last_trade = get_last_daily_row(conn, symbol)

    return {
        "Symbol": symbol,
        "InsCode": live["inscode"],

        "MarketDate": market_date,
        "MarketTime": format_time(
            current.get("hEven") or current.get("lastHEven")
        ),
        "FinalLastDate": to_int(current.get("finalLastDate"), 0),
        "LastStoredTradeDate": last_trade.get("date"),

        "LastPrice": live_last,
        "ClosingPrice": live_close,
        "Yesterday": py,
        "LastChangePct": pct_change(live_last, py),
        "ClosingChangePct": pct_change(live_close, py),
        "Open": to_float(current.get("priceFirst"), 0),
        "Low": to_float(current.get("priceMin"), 0),
        "High": to_float(current.get("priceMax"), 0),

        "LiveVolume": live_volume,
        "LiveValue_Rial": live_value,
        "LiveValue_BnToman": live_value / 1e10,
        "LiveTrades": live_trades,

        "AvgVol20_Previous": avgvol20_prev,
        "AvgVol60_Previous": avgvol60_prev,
        "VolumeRatio20_Live": safe_div(live_volume, avgvol20_prev),
        "VolumeRatio60_Live": safe_div(live_volume, avgvol60_prev),

        "RealBuyVolume": buy_i_vol,
        "RealSellVolume": sell_i_vol,
        "RealNetVolume": real_net_vol,
        "RealBuyCount": buy_i_count,
        "RealSellCount": sell_i_count,
        "LegalBuyCount": client_metric(client, "buy_N_Count", "buy_CountN"),
        "LegalSellCount": client_metric(client, "sell_N_Count", "sell_CountN"),
        "AvgRealBuyVolume": avg_buy_vol_per_buyer,
        "AvgRealSellVolume": avg_sell_vol_per_seller,

        "RealNet_Rial": real_net_value,
        "RealNet_BnToman": (
            real_net_value / 1e10
            if real_net_value is not None else None
        ),

        "RealNetApprox_Rial": (
            real_net_approx
            if real_net_value is None else None
        ),
        "BuyerPower_Volume": buyer_power,

        "LegalBuyVolume": buy_n_vol,
        "LegalSellVolume": sell_n_vol,
        "LegalNetVolume": buy_n_vol - sell_n_vol,

        "RSI14": technical_source.get("RSI14"),
        "MACD": technical_source.get("MACD"),
        "MACD_Signal": technical_source.get("MACD_Signal"),
        "MACD_Hist": technical_source.get("MACD_Hist"),
        "SMA20": technical_source.get("SMA20"),
        "SMA50": technical_source.get("SMA50"),
        "SMA200": technical_source.get("SMA200"),
        "OBV": technical_source.get("OBV"),
        "High52W": technical_source.get("High52W"),
        "Low52W": technical_source.get("Low52W"),
        "IndicatorsProvisional": provisional,

        "MajorHolderBuyShares": holder_buy,
        "MajorHolderSellShares": holder_sell,
        "MajorHolderNetShares": holder_buy + holder_sell,

        "PV_SnapshotCount": q.get("SnapshotCount"),
        "PV_FirstTime": q.get("FirstTime"),
        "PV_LastTime": q.get("LastTime"),
        "PV_MaxGapMinutes": q.get("MaxGapMinutes"),
        "PV_CoveragePct": q.get("CoveragePct"),
        "PV_Status": q.get("Status"),
    }


# ============================================================
# EXCEL WRITER
# ============================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def clean_sheet_name(name):
    invalid = '[]:*?/\\'
    return "".join("_" if c in invalid else c for c in name)[:31]


def rows_to_table(rows):
    rows = rows or []

    if isinstance(rows, dict):
        rows = [rows]

    if not rows:
        return [], []

    dict_rows = [r for r in rows if isinstance(r, dict)]

    if not dict_rows:
        return ["value"], [[excel_safe(x)] for x in rows]

    headers = []
    seen = set()

    for r in dict_rows:
        for k in r:
            if k not in seen:
                seen.add(k)
                headers.append(k)

    matrix = [
        [excel_safe(r.get(h)) for h in headers]
        for r in dict_rows
    ]

    return headers, matrix


def write_sheet(wb, name, rows, rtl=False):
    ws = wb.create_sheet(clean_sheet_name(name))
    ws.sheet_view.rightToLeft = rtl

    headers, matrix = rows_to_table(rows)

    if not headers:
        ws["A1"] = "No data"
        return ws

    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for r_idx, row in enumerate(matrix, 2):
        for c_idx, value in enumerate(row, 1):
            ws.cell(r_idx, c_idx, value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    scan_rows = min(ws.max_row, 100)

    for c in range(1, ws.max_column + 1):
        max_len = 0

        for r in range(1, scan_rows + 1):
            v = ws.cell(r, c).value
            if v is not None:
                max_len = max(max_len, len(str(v)))

        ws.column_dimensions[get_column_letter(c)].width = min(
            max(max_len + 2, 10),
            32
        )

    return ws


def write_readme(wb):
    ws = wb.create_sheet("README", 0)

    rows = [
        ("TSETMC LIVE FAST", "Light intraday updater; full history is stored locally, not re-downloaded."),
        ("Persistent DB", str(DB_PATH)),
        ("SUMMARY", "Live price/volume/client-type + comparison against local history."),
        ("MARKET_INDICES", "Current TSETMC شاخص کل and شاخص کل هم‌وزن on every run."),
        ("INDEX_HISTORY_260D", "Last up to 260 historical rows for both market indices."),
        ("DAILY_260D", "Last 260 daily sessions: adjusted price, volume, RSI, MACD, SMA, OBV."),
        ("MONTHLY_ALL", "Compact complete monthly history for long-term trend and Elliott context."),
        ("PV_5M_TODAY", "Today's real TSETMC trade-price path in 5-minute buckets (TradeHistory/TradeIntraDay)."),
        ("PV_QUALITY", "Intraday price-path quality; official daily volume still comes from ClosingPriceInfo."),
        ("CLIENT_INTRADAY", "Time-stamped cumulative buyer/seller counts, volumes, BP and net flow during the day."),
        ("CLIENT_CURRENT", "Current real/legal participation."),
        ("BEST_LIMITS", "Current top bid/ask levels."),
        ("HOLDERS", "Major holders, cached for speed."),
        ("CODAL", "Recent Codal metadata, cached for speed."),
        ("MESSAGES", "Recent TSETMC messages, cached for speed."),
        ("IMPORTANT", "Run once near/after market close too so today's daily bar becomes as final as possible."),
        ("V4", "Do not run V4 during every market check. Keep it only as baseline/backup or when rebuilding local DB."),
    ]

    for i, (a, b) in enumerate(rows, 1):
        ws.cell(i, 1, a)
        ws.cell(i, 2, b)

    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 23
    ws.column_dimensions["B"].width = 105


def apply_summary_formats(ws):
    headers = {
        ws.cell(1, c).value: c
        for c in range(1, ws.max_column + 1)
    }

    fmt_cols = [
        "LastChangePct",
        "ClosingChangePct",
        "LiveValue_BnToman",
        "VolumeRatio20_Live",
        "VolumeRatio60_Live",
        "RealNet_BnToman",
        "BuyerPower_Volume",
        "RSI14",
        "MACD",
        "MACD_Signal",
        "MACD_Hist",
        "PV_CoveragePct",
    ]

    for h in fmt_cols:
        c = headers.get(h)
        if c:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = "0.00"


def build_output(conn, live_data, output_path, market_indices=None, index_history=None):
    wb = Workbook()
    wb.remove(wb.active)

    write_readme(wb)

    summaries = []
    all_daily = []
    all_monthly = []
    all_pv = []
    all_quality = []
    all_clients = []
    all_best = []
    all_holders = []
    all_codal = []
    all_messages = []

    for live in live_data:
        symbol = live["symbol"]

        summaries.append(make_summary(conn, live))

        enriched = enrich_history(get_daily_rows(conn, symbol))

        all_daily.extend(
            compact_daily_rows(
                enriched,
                symbol,
                DAILY_CONTEXT_DAYS
            )
        )

        all_monthly.extend(
            monthly_history(enriched, symbol)
        )

        all_pv.extend(live.get("pv", []))
        all_quality.append(live.get("pv_quality", {}))

        client = dict(live.get("client") or {})
        client["Symbol"] = symbol
        all_clients.append(client)

        all_best.extend(live.get("best", []))
        all_holders.extend(live.get("holders", []))
        all_codal.extend(live.get("codal", []))
        all_messages.extend(live.get("messages", []))

    ws = write_sheet(wb, "SUMMARY", summaries)
    apply_summary_formats(ws)

    ws_idx = write_sheet(wb, "MARKET_INDICES", market_indices or [], rtl=True)
    idx_headers = {ws_idx.cell(1, c).value: c for c in range(1, ws_idx.max_column + 1)}
    c = idx_headers.get("ChangePct")
    if c:
        for r in range(2, ws_idx.max_row + 1):
            ws_idx.cell(r, c).number_format = "0.00"
    write_sheet(wb, "INDEX_HISTORY_260D", index_history or [], rtl=True)

    write_sheet(wb, "PV_5M_TODAY", all_pv)
    write_sheet(wb, "PV_QUALITY", all_quality)
    write_sheet(wb, "DAILY_260D", all_daily)
    write_sheet(wb, "MONTHLY_ALL", all_monthly)

    market_dates = [to_int((x.get("current") or {}).get("dEven"), 0) for x in live_data]
    market_dates = [d for d in market_dates if d]
    client_intraday = get_client_intraday_rows(conn, max(market_dates) if market_dates else None)
    write_sheet(wb, "CLIENT_INTRADAY", client_intraday)

    write_sheet(wb, "CLIENT_CURRENT", all_clients)
    write_sheet(wb, "BEST_LIMITS", all_best)
    write_sheet(wb, "HOLDERS", all_holders, rtl=True)
    write_sheet(wb, "CODAL", all_codal, rtl=True)
    write_sheet(wb, "MESSAGES", all_messages, rtl=True)

    write_sheet(wb, "API_STATUS", API_STATUS)

    wb.save(output_path)


# ============================================================
# MAIN
# ============================================================

def tehran_now():
    return datetime.now(TEHRAN_TZ)


def hhmmss(dt):
    return dt.hour * 10000 + dt.minute * 100 + dt.second


def telegram_session():
    s = requests.Session()
    s.trust_env = False  # Never send Telegram traffic through the public TSETMC proxy.
    return s


def telegram_send_message(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False
    try:
        s = telegram_session()
        r = s.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
            data={"chat_id": TELEGRAM_CHAT_ID, "text": text},
            timeout=30,
        )
        r.raise_for_status()
        return True
    except Exception as e:
        print("Telegram message failed:", e)
        return False


def telegram_send_document(path, caption=""):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID or not Path(path).exists():
        return False
    try:
        s = telegram_session()
        with open(path, "rb") as f:
            r = s.post(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption},
                files={"document": f},
                timeout=120,
            )
        r.raise_for_status()
        return True
    except Exception as e:
        print("Telegram document failed:", e)
        return False


def telegram_get_updates(offset=None):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return [], offset
    try:
        s = telegram_session()
        params = {"timeout": 0}
        if offset is not None:
            params["offset"] = offset
        r = s.get(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates",
            params=params,
            timeout=20,
        )
        r.raise_for_status()
        result = r.json().get("result", [])
        next_offset = offset
        if result:
            next_offset = max(x.get("update_id", 0) for x in result) + 1
        return result, next_offset
    except Exception as e:
        print("Telegram getUpdates failed:", e)
        return [], offset


def build_status_text(live_data):
    lines = ["TSETMC live status"]
    for live in live_data or []:
        current = live.get("current") or {}
        client = live.get("client") or {}
        symbol = live.get("symbol") or "?"
        rbv = client_metric(client, "buy_I_Volume")
        rsv = client_metric(client, "sell_I_Volume")
        rbc = client_metric(client, "buy_I_Count", "buy_CountI")
        rsc = client_metric(client, "sell_I_Count", "sell_CountI")
        ab = safe_div(rbv, rbc)
        ass = safe_div(rsv, rsc)
        bp = safe_div(ab, ass) if ab is not None and ass is not None else None
        bp_txt = f"{bp:.2f}" if bp is not None else "-"
        lines.append(
            f"{symbol}: Last={to_int(current.get('pDrCotVal'),0):,} | "
            f"B/S={to_int(rbc,0)}/{to_int(rsc,0)} | BP={bp_txt} | "
            f"RealNet={to_int(rbv-rsv,0):,}"
        )
    return "\n".join(lines)


def collect_fast_symbol(conn, symbol):
    """Lightweight snapshot: price + client type + best limits only."""
    inscode = meta_get(conn, f"inscode:{symbol}")
    if not inscode:
        inscode = find_symbol_online(symbol)
        meta_set(conn, f"inscode:{symbol}", inscode)

    p = api_get(
        f"ClosingPrice/GetClosingPriceInfo/{inscode}",
        symbol,
        "FAST_ClosingPriceInfo",
        allow_fail=False,
    )
    current = unwrap_dict(p, ("closingPriceInfo",))

    market_date = to_int(
        current.get("dEven"),
        to_int(tehran_now().strftime("%Y%m%d"), 0),
    )
    current_volume = to_float(current.get("qTotTran5J"), 0)
    current_close = to_float(current.get("pClosing"), 0)
    final_last_date = to_int(current.get("finalLastDate"), 0)
    is_final = market_date > 0 and final_last_date == market_date

    if current_volume > 0 and current_close > 0:
        upsert_daily(conn, symbol, inscode, current, finalized=is_final)

    client = {}
    if FETCH_CLIENT_TYPE:
        p = api_get(
            f"ClientType/GetClientType/{inscode}/1/0",
            symbol,
            "FAST_ClientTypeCurrent",
        )
        client = unwrap_dict(p, ("clientType",))

    best = []
    if FETCH_BEST_LIMITS:
        p = api_get(
            f"BestLimits/{inscode}",
            symbol,
            "FAST_BestLimitsCurrent",
        )
        best = unwrap_list(p, ("bestLimits",))
        for r in best:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    return {
        "symbol": symbol,
        "inscode": inscode,
        "current": current,
        "client": client,
        "best": best,
        "holders": [],
        "codal": [],
        "messages": [],
        "pv": [],
        "pv_quality": {"Symbol": symbol, "Date": market_date, "Status": "FAST_ONLY"},
    }


def run_fast_snapshot(conn):
    """Collect only the fast-changing fields. Designed for ~30-second cadence."""
    global API_STATUS
    API_STATUS = []
    started = time.time()
    captured_at = tehran_now().strftime("%H:%M:%S")
    live_data = []

    for symbol in SYMBOLS:
        try:
            live = collect_fast_symbol(conn, symbol)
            live_data.append(live)
            save_client_intraday(conn, live, captured_at=captured_at)
        except Exception as e:
            print(f"FAST {symbol} failed: {e}")
            API_STATUS.append({
                "Symbol": symbol,
                "Endpoint": "FAST_SYMBOL_TOTAL",
                "Status": "FAILED",
                "HTTP": "",
                "Attempts": "",
                "Seconds": "",
                "Error": str(e),
            })

    print(f"Fast snapshot complete in {time.time()-started:.1f}s ({len(live_data)}/{len(SYMBOLS)} symbols)")
    return live_data


def run_once(conn, output_path=None):
    global API_STATUS
    API_STATUS = []
    start = time.time()
    now = tehran_now()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    if output_path is None:
        output_path = APP_DIR / f"TSETMC_LIVE_{stamp}.xlsx"

    live_data = []
    captured_at = now.strftime("%H:%M:%S")

    for symbol in SYMBOLS:
        print(f"Fetching {symbol} ...")
        try:
            live = collect_live_symbol(conn, symbol)
            live_data.append(live)
            save_client_intraday(conn, live, captured_at=captured_at)
        except Exception as e:
            print(f"  FAILED: {e}")
            API_STATUS.append({
                "Symbol": symbol, "Endpoint": "SYMBOL_TOTAL", "Status": "FAILED",
                "HTTP": "", "Attempts": "", "Seconds": "", "Error": str(e),
            })

    if not live_data:
        raise RuntimeError("No live symbol data collected.")

    print("Fetching market indices ...")
    try:
        market_indices, index_history = collect_market_indices(conn)
    except Exception as e:
        print(f"  INDEX FAILED: {e}")
        API_STATUS.append({
            "Symbol": "MARKET", "Endpoint": "INDEX_TOTAL", "Status": "FAILED",
            "HTTP": "", "Attempts": "", "Seconds": "", "Error": str(e),
        })
        market_indices, index_history = [], []

    build_output(
        conn, live_data, output_path,
        market_indices=market_indices,
        index_history=index_history,
    )
    print(f"Snapshot complete in {time.time()-start:.1f}s -> {output_path.name}")
    return Path(output_path), live_data


def handle_telegram_commands(conn, live_data, offset, report_state=None):
    updates, offset = telegram_get_updates(offset)
    for upd in updates:
        msg = upd.get("message") or upd.get("edited_message") or {}
        chat = str((msg.get("chat") or {}).get("id", ""))
        if chat != str(TELEGRAM_CHAT_ID):
            continue
        text = (msg.get("text") or "").strip().lower()
        if text.startswith("/report"):
            path = None
            if report_state:
                path = report_state.get("path")
            if path and Path(path).exists():
                age = int(max(0, time.time() - report_state.get("ts", time.time())) / 60)
                telegram_send_document(path, f"آخرین گزارش کامل TSETMC — حدود {age} دقیقه قبل")
            else:
                telegram_send_message("گزارش کامل هنوز ساخته نشده؛ حداکثر تا رفرش ۵ دقیقه‌ای بعدی آماده می‌شود.")
        elif text.startswith("/status"):
            telegram_send_message(build_status_text(live_data))
    return offset, live_data


def heavy_worker(stop_event, report_state, report_lock):
    """Run heavy collection/report generation every 5 minutes without blocking fast snapshots."""
    conn = connect_db()
    try:
        ensure_bootstrap(conn)
        while not stop_event.is_set():
            started = time.time()
            out = APP_DIR / f"TSETMC_HEAVY_{tehran_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            try:
                path, _ = run_once(conn, out)
                with report_lock:
                    report_state["path"] = str(path)
                    report_state["ts"] = time.time()
                print(f"Heavy refresh ready: {Path(path).name}")
            except Exception as e:
                print("Heavy refresh failed:", e)

            elapsed = time.time() - started
            wait_for = max(5, HEAVY_REFRESH_SECONDS - elapsed)
            stop_event.wait(wait_for)
    finally:
        conn.close()


def collector_main():
    print("=" * 58)
    print("TSETMC AUTOMATIC COLLECTOR — FAST 30s / HEAVY 5m")
    print("=" * 58)
    print("Local DB:", DB_PATH)
    print("Iran proxies configured:", len(TSETMC_PROXIES))
    print("Fast interval:", SNAPSHOT_SECONDS, "sec")
    print("Heavy interval:", HEAVY_REFRESH_SECONDS, "sec")

    conn = connect_db()
    stop_event = threading.Event()
    report_state = {"path": None, "ts": 0.0}
    report_lock = threading.Lock()
    heavy_thread = None

    try:
        ensure_bootstrap(conn)
        offset = None
        live_data = []
        last_fast_time = 0.0

        old, offset = telegram_get_updates(None)
        print(f"Telegram ready; ignored {len(old)} old update(s).")
        telegram_send_message("✅ Collector آماده است؛ Fast هر ۳۰ ثانیه و Heavy هر ۵ دقیقه فعال است.")

        now = tehran_now()
        if hhmmss(now) < MARKET_START_HHMMSS:
            telegram_send_message("TSETMC collector روشن شد و منتظر شروع بازار است.")

        while True:
            now = tehran_now()
            clock = hhmmss(now)

            if clock < MARKET_START_HHMMSS:
                offset, live_data = handle_telegram_commands(conn, live_data, offset, report_state)
                time.sleep(3)
                continue

            if clock > MARKET_END_HHMMSS:
                break

            if heavy_thread is None:
                heavy_thread = threading.Thread(
                    target=heavy_worker,
                    args=(stop_event, report_state, report_lock),
                    daemon=True,
                    name="tsetmc-heavy",
                )
                heavy_thread.start()

            # Fast snapshot cadence is measured from cycle start, not cycle end.
            now_ts = time.time()
            if now_ts - last_fast_time >= SNAPSHOT_SECONDS:
                last_fast_time = now_ts
                try:
                    live_data = run_fast_snapshot(conn)
                except Exception as e:
                    print("Fast cycle failed:", e)
                    telegram_send_message(f"هشدار Fast Collector: {e}")

            # Poll Telegram frequently, independent of 30-second market snapshots.
            offset, live_data = handle_telegram_commands(conn, live_data, offset, report_state)
            time.sleep(2)

        stop_event.set()
        if heavy_thread is not None:
            heavy_thread.join(timeout=10)

        final_path = APP_DIR / f"TSETMC_FINAL_{tehran_now().strftime('%Y%m%d_%H%M')}.xlsx"
        try:
            final_path, final_live = run_once(conn, final_path)
            telegram_send_document(final_path, "گزارش نهایی بازار امروز")
            telegram_send_message(build_status_text(final_live))
        except Exception as e:
            print("Final report failed:", e)
            with report_lock:
                fallback = report_state.get("path")
            if fallback and Path(fallback).exists():
                telegram_send_document(fallback, "آخرین گزارش کامل موجود امروز")

        print("Collector finished for today.")
    finally:
        stop_event.set()
        conn.close()


def main():
    if COLLECTOR_MODE:
        collector_main()
        return

    # Backward-compatible one-shot mode for manual/Pydroid runs.
    conn = connect_db()
    try:
        ensure_bootstrap(conn)
        output_path, _ = run_once(conn)
        print("\nDONE")
        print("Output:", output_path)
        print("Do NOT delete TSETMC_LOCAL_HISTORY.db.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
