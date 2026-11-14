# -*- coding: utf-8 -*-
"""
TSETMC LIVE FAST + MARKET INDEX + INTRADAY FIX + MULTI-MIN MOMENTUM
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
    PV_5M_TODAY      = today's real TSETMC trade-price path in 5-minute buckets
    PV_QUALITY       = intraday data-quality check
    CLIENT_CURRENT   = current حقیقی/حقوقی
    BEST_LIMITS      = current order book
    HOLDERS          = cached/current major holders
    CODAL            = cached/recent Codal metadata
    MESSAGES         = cached/recent TSETMC messages
    API_STATUS       = request status

Set TSETMC_EXPORT_HISTORY=1 for an occasional export with DAILY_260D,
MONTHLY_ALL and INDEX_HISTORY_260D. Price history remains in the database
and feeds the indicators in SUMMARY even when those sheets are omitted.

Alert volume and published daily volume use one matched ClientType side
(real + legal buys = real + legal sells). Their receipt times describe feed
updates, not individual execution times. Official volume remains archived
for diagnostics and order-cancellation estimates. Fast alert baselines are
isolated from Heavy collection; the existing local database upgrades in place.

Revision 2026-11-11.1:
  - Real flow: 200-500m toman = notable; over 500m = heavy, per side.
    Below 200m purchases require at least 150m estimated value per new buyer.
  - Significant sell-order additions outside existing queue states.
  - Persistent shares/event timeline normalizes prior positive-volume sessions
    to current capital; actual traded volume is never rescaled in messages.
  - Preopening book alerts are labeled; trading alerts start at 09:00.
  - Channel close reports are sent separately for each symbol, in watchlist order.

Revision 2026-11-11.2:
  - Persist successful per-recipient deliveries and subtract their exact ClientType
    intervals from later windows; averages use only the remaining intervals.
  - Combine simultaneous price reasons; later momentum starts after reported moves.
  - Phase crossings need both 1% of comparable prior-20-session volume and 50m toman.
  - Retry failed deliveries without consuming channel coverage. First activation
    starts at current counters; restarts retain that boundary and delivery history.

Revision 2026-11-14.1:
  - Concise queue growth/shrink alerts: before/after shares and toman value,
    estimated average toman VALUE per visible order; no order-count display.
  - Estimate a conservative lower bound of queue removal unrelated to trades,
    but only with fresh, monotonic official-volume samples on both book timestamps.
    Never equate queue shrinkage with confirmed buying or exact cancellations.
  - Do not announce a queue change when the newest book equals the prior book.
  - Include Thamid as the eighth watchlist symbol; use Jalali dates in
    channel closing messages without adding an external package dependency.

Revision 2026-11-14.2:
  - Add Khzamia (Zamyad ordinary share) as the ninth tracked symbol to
    Fast, Heavy, order-book alerts and per-symbol close summaries.
"""

import csv
import hashlib
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
TELEGRAM_CHANNEL_ID = os.getenv("TELEGRAM_CHANNEL_ID", "").strip()
COLLECTOR_MODE = os.getenv("COLLECTOR_MODE", "0") == "1"
SNAPSHOT_SECONDS = max(15, int(os.getenv("SNAPSHOT_SECONDS", "30")))
HEAVY_REFRESH_SECONDS = max(60, int(os.getenv("HEAVY_REFRESH_SECONDS", "300")))
EXPORT_HISTORY_SHEETS = os.getenv("TSETMC_EXPORT_HISTORY", "0") == "1"
COLLECTOR_VERSION = "2026-11-14.2"
PHASE_MIN_VOLUME_FRACTION = 0.01
PHASE_MIN_VALUE_TOMAN = 50_000_000

# Fast alerts: intentionally conservative to avoid Telegram spam.
ALERT_COOLDOWN_SECONDS = max(30, int(os.getenv("ALERT_COOLDOWN_SECONDS", "180")))
ALERT_MIN_SHARES = max(1, int(os.getenv("ALERT_MIN_SHARES", "200000")))
ALERT_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("ALERT_DAILY_VOLUME_FRACTION", "0.03")))
ORDER_ALERT_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("ORDER_ALERT_DAILY_VOLUME_FRACTION", "0.05")))
CONCENTRATED_MAX_NEW_PEOPLE = max(1, int(os.getenv("CONCENTRATED_MAX_NEW_PEOPLE", "4")))
REAL_FLOW_NOTICE_TOMAN = 200_000_000
REAL_FLOW_HEAVY_TOMAN = 500_000_000
SMALL_BUY_MIN_AVERAGE_TOMAN = 150_000_000
SUDDEN_PRICE_PCT_30S = max(0.1, float(os.getenv("SUDDEN_PRICE_PCT_30S", "0.60")))
# Multi-minute momentum alert: catches step-by-step aggressive buying/selling that a
# single 30-second jump can miss (e.g. a 2% rise spread over ~4 minutes).
MOMENTUM_MIN_SECONDS = max(60, int(os.getenv("MOMENTUM_MIN_SECONDS", "180")))
MOMENTUM_MAX_SECONDS = max(MOMENTUM_MIN_SECONDS + 30, int(os.getenv("MOMENTUM_MAX_SECONDS", "300")))
MOMENTUM_PRICE_PCT = max(0.2, float(os.getenv("MOMENTUM_PRICE_PCT", "1.20")))
MOMENTUM_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("MOMENTUM_DAILY_VOLUME_FRACTION", "0.03")))
MOMENTUM_ALERT_COOLDOWN_SECONDS = max(60, int(os.getenv("MOMENTUM_ALERT_COOLDOWN_SECONDS", "300")))
# Positive-to-stronger-positive alert: catches acceleration that is meaningful but
# smaller than the main 3-5 minute momentum trigger.
POSITIVE_STRENGTHEN_PCT = max(0.2, float(os.getenv("POSITIVE_STRENGTHEN_PCT", "0.80")))
POSITIVE_STRENGTHEN_VOLUME_FRACTION = max(0.001, float(os.getenv("POSITIVE_STRENGTHEN_VOLUME_FRACTION", "0.02")))
POSITIVE_STRENGTHEN_COOLDOWN_SECONDS = max(60, int(os.getenv("POSITIVE_STRENGTHEN_COOLDOWN_SECONDS", "300")))
QUEUE_MIN_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("QUEUE_MIN_DAILY_VOLUME_FRACTION", "0.03")))
# Queue growth/shrink alert: compare the visible queue with a 3-5 minute reference,
# so gradual queue changes are not missed by 30-second comparisons.
QUEUE_CHANGE_MIN_SECONDS = max(60, int(os.getenv("QUEUE_CHANGE_MIN_SECONDS", "180")))
QUEUE_CHANGE_MAX_SECONDS = max(QUEUE_CHANGE_MIN_SECONDS + 30, int(os.getenv("QUEUE_CHANGE_MAX_SECONDS", "300")))
QUEUE_CHANGE_PCT = max(5.0, float(os.getenv("QUEUE_CHANGE_PCT", "20.0")))
QUEUE_CHANGE_MIN_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("QUEUE_CHANGE_MIN_DAILY_VOLUME_FRACTION", "0.01")))
QUEUE_CHANGE_COOLDOWN_SECONDS = max(60, int(os.getenv("QUEUE_CHANGE_COOLDOWN_SECONDS", "180")))
# Require receipt-time alignment between archived official cumulative volumes and
# both order-book timestamps. Otherwise, do NOT infer deleted vs executed queue.
QUEUE_TRADE_MATCH_MAX_LAG_SECONDS = max(15, int(os.getenv("QUEUE_TRADE_MATCH_MAX_LAG_SECONDS", "90")))
# Institutional/legal flow alert. Triggered from ClientType cumulative deltas;
# never infer order ownership from the visible order book.
LEGAL_FLOW_DAILY_VOLUME_FRACTION = max(0.001, float(os.getenv("LEGAL_FLOW_DAILY_VOLUME_FRACTION", "0.03")))
LEGAL_FLOW_COOLDOWN_SECONDS = max(60, int(os.getenv("LEGAL_FLOW_COOLDOWN_SECONDS", "180")))
VOLUME_RATIO_ALERT = max(1.0, float(os.getenv("VOLUME_RATIO_ALERT", "2.0")))
VOLUME_BURST_SAMPLE_FRACTION = max(0.001, float(os.getenv("VOLUME_BURST_SAMPLE_FRACTION", "0.05")))
VOLUME_BURST_WINDOW_FRACTION = max(0.001, float(os.getenv("VOLUME_BURST_WINDOW_FRACTION", "0.20")))
VOLUME_BURST_WINDOW_SECONDS = 300
VOLUME_BURST_COOLDOWN_SECONDS = max(30, int(os.getenv("VOLUME_BURST_COOLDOWN_SECONDS", "60")))

CLIENT_FIELDS = {
    "real_buy_volume": ("buy_I_Volume",),
    "real_sell_volume": ("sell_I_Volume",),
    "legal_buy_volume": ("buy_N_Volume",),
    "legal_sell_volume": ("sell_N_Volume",),
    "real_buy_count": ("buy_I_Count", "buy_CountI"),
    "real_sell_count": ("sell_I_Count", "sell_CountI"),
    "legal_buy_count": ("buy_N_Count", "buy_CountN"),
    "legal_sell_count": ("sell_N_Count", "sell_CountN"),
}
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
    "ثامید",
    "خزامیا",
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
MARKET_START_HHMMSS = 84500
TRADING_START_HHMMSS = 90000
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


def jalali_date(value):
    """Convert a Gregorian datetime/date to YYYY/MM/DD Solar Hijri (Jalali)."""
    gy, gm, gd = value.year, value.month, value.day
    month_days = (0,31,59,90,120,151,181,212,243,273,304,334)
    gy2 = gy + (1 if gm > 2 else 0)
    days = (355666 + 365*gy + (gy2+3)//4 - (gy2+99)//100
            + (gy2+399)//400 + gd + month_days[gm-1])
    jy = -1595 + 33*(days//12053)
    days %= 12053
    jy += 4*(days//1461)
    days %= 1461
    if days > 365:
        jy += (days-1)//365
        days = (days-1)%365
    if days < 186:
        jm, jd = 1+days//31, 1+days%31
    else:
        jm, jd = 7+(days-186)//30, 1+(days-186)%30
    return f"{jy:04d}/{jm:02d}/{jd:02d}"


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
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")

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
        CREATE TABLE IF NOT EXISTS best_limits_intraday (
            symbol TEXT NOT NULL,
            date INTEGER NOT NULL,
            captured_at TEXT NOT NULL,
            level_no INTEGER NOT NULL,
            buy_price REAL,
            buy_volume REAL,
            buy_count REAL,
            sell_price REAL,
            sell_volume REAL,
            sell_count REAL,
            PRIMARY KEY(symbol, date, captured_at, level_no)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS cache (
            cache_key TEXT PRIMARY KEY,
            updated_ts REAL NOT NULL,
            payload TEXT NOT NULL
        )
    """)

    # Additive migration: retain official volumes and existing raw history.
    for table, columns in {
        "client_intraday": {
            "capture_source": "TEXT DEFAULT 'legacy'",
            "price_received_at": "TEXT",
            "client_received_at": "TEXT",
            "classified_volume": "REAL",
            "client_status": "TEXT",
        },
        "best_limits_intraday": {"capture_source": "TEXT DEFAULT 'legacy'"},
    }.items():
        existing = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
        for name, declaration in columns.items():
            if name not in existing:
                try:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {declaration}")
                except sqlite3.OperationalError:
                    # The Fast/Heavy connections may migrate concurrently.
                    if name not in {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}:
                        raise
    conn.execute("""
        CREATE TABLE IF NOT EXISTS alert_samples (
            symbol TEXT NOT NULL, date INTEGER NOT NULL, kind TEXT NOT NULL,
            captured_at TEXT NOT NULL, payload TEXT NOT NULL,
            PRIMARY KEY(symbol,date,kind,captured_at)
        )
    """)
    conn.execute("""CREATE INDEX IF NOT EXISTS alert_samples_lookup
                    ON alert_samples(symbol,date,kind,captured_at)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS instrument_shares (
        inscode TEXT NOT NULL, date INTEGER NOT NULL, shares REAL NOT NULL,
        PRIMARY KEY(inscode,date))""")
    conn.execute("""CREATE TABLE IF NOT EXISTS instrument_share_changes (
        inscode TEXT NOT NULL, date INTEGER NOT NULL,
        old_shares REAL NOT NULL, new_shares REAL NOT NULL,
        PRIMARY KEY(inscode,date,old_shares,new_shares))""")
    for table in ("market_alert_receipts", "market_alert_outbox"):
        conn.execute(f"""CREATE TABLE IF NOT EXISTS {table} (
            target TEXT NOT NULL, report_id TEXT NOT NULL, symbol TEXT NOT NULL,
            date INTEGER NOT NULL, text TEXT NOT NULL, payload TEXT NOT NULL,
            created REAL NOT NULL, PRIMARY KEY(target,report_id))""")
    conn.execute("""CREATE INDEX IF NOT EXISTS market_receipts_lookup
                    ON market_alert_receipts(target,symbol,date,created)""")
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


def received_clock():
    """Actual endpoint receipt time, rather than the start of a multi-symbol cycle."""
    return tehran_now().strftime("%H:%M:%S.%f")


def symbol_tag(symbol):
    return "#" + str(symbol or "?").lstrip("#").replace(" ", "_")


def client_snapshot(live, captured_at=None):
    """Validate both complete ClientType sides; missing packets are never zero trades."""
    client = live.get("client") or {}
    current = live.get("current") or {}
    result = {
        "captured_at": live.get("client_received_at") or captured_at or received_clock(),
        "last_price": to_float(current.get("pDrCotVal"), 0),
        "volume": to_float(current.get("qTotTran5J"), 0),  # official diagnostic only
        "ready": False, "status": "MISSING",
    }
    for field, keys in CLIENT_FIELDS.items():
        value = next((client[k] for k in keys if client.get(k) not in (None, "")), None)
        result[field] = to_float(value, float("nan")) if value is not None else None
    if not all(v is not None and math.isfinite(v) and v >= 0 and v == int(v)
               for v in (result[k] for k in CLIENT_FIELDS)):
        return result
    # Unknown participant categories cannot silently disappear from total volume.
    other = [to_float(v, float("nan")) for k, v in client.items()
             if k.endswith("_Volume") and k not in
             ("buy_I_Volume", "sell_I_Volume", "buy_N_Volume", "sell_N_Volume")]
    if any(not math.isfinite(v) or v != 0 for v in other):
        result["status"] = "OTHER_CATEGORY"
        return result
    buy = result["real_buy_volume"] + result["legal_buy_volume"]
    sell = result["real_sell_volume"] + result["legal_sell_volume"]
    if buy != sell:
        result["status"] = "SIDE_MISMATCH"
        return result
    for group in ("real", "legal"):
        for side in ("buy", "sell"):
            if result[f"{group}_{side}_volume"] > 0 and result[f"{group}_{side}_count"] <= 0:
                result["status"] = "COUNT_MISSING"
                return result
    result.update(classified_volume=buy, ready=buy > 0, status="READY" if buy > 0 else "EMPTY")
    return result


def client_interval(now, previous):
    """Only matched, monotonic client counters in the same correction epoch."""
    if not now or not previous or not now.get("ready") or not previous.get("ready"):
        return None
    if now.get("epoch", 0) != previous.get("epoch", 0):
        return None
    values = {key: now[key] - previous[key] for key in CLIENT_FIELDS}
    if any(v < 0 for v in values.values()):
        return None
    values["classified_volume"] = now["classified_volume"] - previous["classified_volume"]
    return values if values["classified_volume"] >= 0 else None


def save_alert_sample(conn, symbol, day, kind, sample):
    conn.execute("""INSERT OR REPLACE INTO alert_samples
                    (symbol,date,kind,captured_at,payload) VALUES(?,?,?,?,?)""",
                 (symbol, day, kind, sample["captured_at"],
                  json.dumps(sample, ensure_ascii=False, allow_nan=False)))
    conn.commit()


def alert_sample(conn, symbol, day, kind, before=None, lower=None, upper=None, oldest=False):
    sql = "SELECT payload FROM alert_samples WHERE symbol=? AND date=? AND kind=?"
    args = [symbol, day, kind]
    for operator, value in (("<", before), (">=", lower), ("<=", upper)):
        if value is not None:
            sql += f" AND captured_at{operator}?"
            args.append(value)
    sql += " ORDER BY captured_at " + ("ASC" if oldest else "DESC") + " LIMIT 1"
    row = conn.execute(sql, args).fetchone()
    return json.loads(row["payload"]) if row else None


def accept_client_snapshot(conn, live, captured_at):
    """Fast alone advances alert baselines. Reject old totals without lowering the baseline."""
    symbol = live.get("symbol") or "?"
    day = to_int((live.get("current") or {}).get("dEven"), 0)
    sample = client_snapshot(live, captured_at)
    previous = alert_sample(conn, symbol, day, "client")
    live["_flow_reference"] = previous
    live["_client_sample"] = None
    live["_client_status"] = sample["status"]
    if not sample["ready"]:
        return None
    if previous:
        if sample["captured_at"] <= previous["captured_at"] or sample["classified_volume"] < previous["classified_volume"]:
            live["_client_status"] = "ROLLBACK"
            return None
        # A redistribution/count correction is a new baseline, never a trade.
        correction = any(sample[k] < previous[k] for k in CLIENT_FIELDS)
        sample["epoch"] = previous.get("epoch", 0) + int(correction)
        if correction:
            live["_client_status"] = "REBASED"
    else:
        sample["epoch"] = 0
    live["_client_sample"] = sample
    save_alert_sample(conn, symbol, day, "client", sample)
    return sample


def published_client_snapshot(conn, live):
    """Coherent daily data; fall back to the accepted Fast reading on stale Heavy packets."""
    if not current_has_trades(live.get("current") or {}):
        return None
    if "_client_sample" in live:
        return live["_client_sample"]
    sample = client_snapshot(live)
    day = to_int((live.get("current") or {}).get("dEven"), 0)
    accepted = alert_sample(conn, live.get("symbol") or "?", day, "client")
    if accepted and (not sample["ready"] or sample["classified_volume"] < accepted["classified_volume"]
                     or sample["captured_at"] < accepted["captured_at"]):
        return accepted
    return sample if sample["ready"] else None


def store_share_changes(conn, inscode, changes):
    """Persist the dated share timeline; duplicate API events cannot compound it."""
    for row in changes:
        if not isinstance(row, dict):
            continue
        day = to_int(row.get("dEven"), 0)
        old = to_float(row.get("numberOfShareOld"), 0)
        new = to_float(row.get("numberOfShareNew"), 0)
        if 19000101 <= day <= 29991231 and all(math.isfinite(x) and x > 0 for x in (old, new)):
            conn.execute("INSERT OR IGNORE INTO instrument_share_changes VALUES(?,?,?,?)",
                         (str(inscode), day, old, new))
    conn.commit()


def refresh_share_context(conn, symbol, inscode, current):
    """Check each session and the first traded packet; recheck on Heavy refreshes.

    Only actual share-count changes affect volume. Dividend/price factors never do.
    Failed requests retain the saved timeline and retry after a short cache delay.
    """
    day = int(tehran_now().strftime("%Y%m%d"))
    phase = "traded" if current_has_trades(current) else "before_trade"
    key = f"share_check:{inscode}:{day}:{phase}"
    if cache_get(conn, key, 15) is not None:
        return
    if cache_get(conn, f"share_retry:{inscode}", 1) is not None:
        return
    cache_set(conn, f"share_retry:{inscode}", True)
    payload = api_get(f"Instrument/GetInstrumentInfo/{inscode}", symbol, "InstrumentShares")
    info = unwrap_dict(payload, ("instrumentInfo",))
    shares = to_float(info.get("zTitad"), 0)
    previous = conn.execute("SELECT shares FROM instrument_shares WHERE inscode=? ORDER BY date DESC LIMIT 1",
                            (str(inscode),)).fetchone()
    changed = shares > 0 and previous is not None and shares != previous["shares"]
    history_key = f"share_events:{inscode}:{day}"
    events_ok = cache_get(conn, history_key, 24 * 60) is not None
    if changed or not events_ok:
        events_ok = False
        if changed:
            conn.execute("DELETE FROM cache WHERE cache_key=?", (history_key,))
            conn.commit()
        events_payload = api_get(f"Instrument/GetInstrumentShareChange/{inscode}", symbol, "ShareChangeHistory")
        # An empty successful list is valid; a failed request is not a valid empty history.
        if isinstance(events_payload, list) or (isinstance(events_payload, dict)
                                               and isinstance(events_payload.get("instrumentShareChange"), list)):
            store_share_changes(conn, inscode, unwrap_list(events_payload, ("instrumentShareChange",)))
            cache_set(conn, history_key, True)
            events_ok = True
    if math.isfinite(shares) and shares > 0:
        conn.execute("INSERT OR REPLACE INTO instrument_shares VALUES(?,?,?)", (str(inscode), day, shares))
        conn.commit()
        if events_ok:
            cache_set(conn, key, True)


def share_count_on_date(conn, inscode, day):
    """Dated step function, including old capital before the first known event."""
    if not inscode:
        return None
    events = conn.execute("SELECT date,old_shares,new_shares FROM instrument_share_changes WHERE inscode=? ORDER BY date,old_shares",
                          (str(inscode),)).fetchall()
    anchors = [(r["date"], r["new_shares"]) for r in events if r["date"] <= day]
    observed = conn.execute("SELECT date,shares FROM instrument_shares WHERE inscode=? AND date<=? ORDER BY date DESC LIMIT 1",
                            (str(inscode), day)).fetchone()
    if observed:
        anchors.append((observed["date"], observed["shares"]))
    if anchors:
        return max(anchors, key=lambda r: r[0])[1]
    if events:
        return events[0]["old_shares"]
    first = conn.execute("SELECT shares FROM instrument_shares WHERE inscode=? ORDER BY date LIMIT 1",
                         (str(inscode),)).fetchone()
    return first["shares"] if first else None


def comparable_volume(conn, symbol, row, day):
    volume = to_float(row.get("volume"), 0)
    inscode = row.get("inscode") or meta_get(conn, f"inscode:{symbol}")
    shares_then = share_count_on_date(conn, inscode, to_int(row.get("date"), 0))
    shares_now = share_count_on_date(conn, inscode, day)
    return volume * shares_now / shares_then if shares_then and shares_now else volume


def previous_volume_average(conn, symbol, day, sessions=20):
    # Today and suspended/zero-volume sessions are never part of the denominator.
    rows = [r for r in get_daily_rows(conn, symbol)
            if to_int(r.get("date"), 0) < day and to_float(r.get("volume"), 0) > 0][-sessions:]
    return (sum(comparable_volume(conn, symbol, r, day) for r in rows)/len(rows), len(rows)) if rows else (None, 0)


def save_client_intraday(conn, live, captured_at=None):
    """Archive raw endpoint data, including its quality and origin; not an alert baseline."""
    current = live.get("current") or {}
    symbol = live.get("symbol") or ""
    day = to_int(current.get("dEven"), 0)
    if not symbol or not day:
        return
    sample = client_snapshot(live, captured_at)
    def clean(value):
        return value if value is None or math.isfinite(value) else None
    fields = {key: clean(sample[key]) for key in CLIENT_FIELDS}
    rbv, rsv = fields["real_buy_volume"], fields["real_sell_volume"]
    lbv, lsv = fields["legal_buy_volume"], fields["legal_sell_volume"]
    rbc, rsc = fields["real_buy_count"], fields["real_sell_count"]
    avg_buy = safe_div(rbv, rbc)
    avg_sell = safe_div(rsv, rsc)
    ready = sample["ready"]
    bp = real_buyer_power(rbv, rsv, rbc, rsc) if ready else None
    record = dict(
        symbol=symbol, date=day, captured_at=sample["captured_at"],
        market_time=format_time(current.get("hEven") or current.get("lastHEven")),
        last_price=to_float(current.get("pDrCotVal"), 0),
        closing_price=to_float(current.get("pClosing"), 0),
        volume=to_float(current.get("qTotTran5J"), 0),
        **fields, avg_real_buy_volume=avg_buy if ready else None,
        avg_real_sell_volume=avg_sell if ready else None, buyer_power=bp,
        real_net_volume=rbv-rsv if ready else None,
        legal_net_volume=lbv-lsv if ready else None,
        capture_source=live.get("capture_source", "manual"),
        price_received_at=live.get("price_received_at"),
        client_received_at=sample["captured_at"],
        classified_volume=sample.get("classified_volume") if sample["status"] in ("READY", "EMPTY") else None,
        client_status=live.get("_client_status", sample["status"]),
    )
    columns = ",".join(record)
    conn.execute(f"INSERT OR REPLACE INTO client_intraday ({columns}) VALUES ({','.join('?' for _ in record)})",
                 list(record.values()))
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



def save_best_limits_intraday(conn, live, captured_at=None):
    """Persist the visible order book every fast snapshot."""
    current = live.get("current") or {}
    symbol = live.get("symbol") or ""
    date_value = to_int(current.get("dEven"), 0)
    if not symbol or not date_value:
        return

    captured_at = live.get("book_received_at") or captured_at or received_clock()
    rows = live.get("best") or []
    for i, r in enumerate(rows, 1):
        if not isinstance(r, dict):
            continue
        level_no = to_int(r.get("number"), i) or i
        conn.execute("""
            INSERT OR REPLACE INTO best_limits_intraday(
                symbol,date,captured_at,level_no,
                buy_price,buy_volume,buy_count,
                sell_price,sell_volume,sell_count,capture_source
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (
            symbol, date_value, captured_at, level_no,
            to_float(r.get("pMeDem"), 0),
            to_float(r.get("qTitMeDem"), 0),
            to_float(r.get("zOrdMeDem"), 0),
            to_float(r.get("pMeOf"), 0),
            to_float(r.get("qTitMeOf"), 0),
            to_float(r.get("zOrdMeOf"), 0),
            live.get("capture_source", "manual"),
        ))
    conn.commit()


def get_best_limits_intraday_rows(conn, date_value=None):
    if date_value:
        rows = conn.execute("""
            SELECT * FROM best_limits_intraday
            WHERE date=?
            ORDER BY captured_at, symbol, level_no
        """, (date_value,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT * FROM best_limits_intraday
            ORDER BY date, captured_at, symbol, level_no
        """).fetchall()
    return [dict(r) for r in rows]


def previous_client_snapshot(conn, symbol, date_value, captured_at):
    return alert_sample(conn, symbol, date_value, "client", before=captured_at)


def _hms_to_seconds(value):
    try:
        hh, mm, ss = str(value).split(":")[:3]
        return int(hh)*3600 + int(mm)*60 + float(ss)
    except (ValueError, TypeError):
        return None


def _seconds_to_hms(value):
    value = max(0, min(86399, int(value)))
    hh, rem = divmod(value, 3600)
    mm, ss = divmod(rem, 60)
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


def momentum_reference_snapshot(conn, symbol, date_value, captured_at, kind="client"):
    """Fast-only references; Heavy timestamps never enter alert windows."""
    seconds = _hms_to_seconds(captured_at)
    if seconds is None:
        return None
    return alert_sample(conn, symbol, date_value, kind,
                        lower=_seconds_to_hms(seconds-MOMENTUM_MAX_SECONDS),
                        upper=_seconds_to_hms(seconds-MOMENTUM_MIN_SECONDS), oldest=True)


def latest_30s_delta(conn, symbol, date_value):
    rows = conn.execute("""SELECT payload FROM alert_samples
        WHERE symbol=? AND date=? AND kind='client'
        ORDER BY captured_at DESC LIMIT 2""", (symbol, date_value)).fetchall()
    if len(rows) < 2:
        return None
    cur, prev = [json.loads(row["payload"]) for row in rows]
    change = client_interval(cur, prev)
    if change is None:
        return None
    return {"captured_at": cur["captured_at"],
            "d_buy_v": change["real_buy_volume"], "d_sell_v": change["real_sell_volume"],
            "d_buy_c": change["real_buy_count"], "d_sell_c": change["real_sell_count"],
            "net": change["real_buy_volume"]-change["real_sell_volume"]}


def previous_book_snapshot(conn, symbol, date_value, captured_at):
    row = conn.execute("""
        SELECT captured_at FROM best_limits_intraday
        WHERE symbol=? AND date=? AND captured_at<? AND capture_source='fast'
        ORDER BY captured_at DESC LIMIT 1
    """, (symbol, date_value, captured_at)).fetchone()
    if not row:
        return []
    return [dict(r) for r in conn.execute("""
        SELECT * FROM best_limits_intraday
        WHERE symbol=? AND date=? AND captured_at=? AND capture_source='fast'
        ORDER BY level_no
    """, (symbol, date_value, row["captured_at"]))]


def queue_reference_snapshot(conn, symbol, date_value, captured_at):
    """Return a visible order-book snapshot roughly 3-5 minutes earlier."""
    cur_sec = _hms_to_seconds(captured_at)
    if cur_sec is None or cur_sec < QUEUE_CHANGE_MIN_SECONDS:
        return []

    lower_hms = _seconds_to_hms(max(0, cur_sec - QUEUE_CHANGE_MAX_SECONDS))
    upper_hms = _seconds_to_hms(max(0, cur_sec - QUEUE_CHANGE_MIN_SECONDS))
    row = conn.execute("""
        SELECT captured_at FROM best_limits_intraday
        WHERE symbol=? AND date=? AND capture_source='fast'
          AND captured_at>=? AND captured_at<=?
        ORDER BY captured_at ASC
        LIMIT 1
    """, (symbol, date_value, lower_hms, upper_hms)).fetchone()
    if not row:
        return []
    ts = row["captured_at"]
    rows = conn.execute("""
        SELECT * FROM best_limits_intraday
        WHERE symbol=? AND date=? AND capture_source='fast' AND captured_at=?
        ORDER BY level_no
    """, (symbol, date_value, ts)).fetchall()
    return [dict(r) for r in rows]


def fmt_compact_num(x):
    x = to_float(x, 0)
    ax = abs(x)
    if ax >= 1_000_000_000:
        return f"{x/1_000_000_000:.2f}B"
    if ax >= 1_000_000:
        return f"{x/1_000_000:.2f}M"
    if ax >= 1_000:
        return f"{x/1_000:.0f}K"
    return f"{x:,.0f}"


def fmt_toman(x):
    x = to_float(x, 0)
    ax = abs(x)
    sign = "-" if x < 0 else ""
    if ax >= 1_000_000_000:
        return f"{sign}{ax/1_000_000_000:.2f} میلیارد"
    if ax >= 1_000_000:
        return f"{sign}{ax/1_000_000:.1f} میلیون"
    if ax >= 1_000:
        return f"{sign}{ax/1_000:.0f} هزار"
    return f"{x:,.0f}"


def fmt_price_pct(price, yesterday=None, pct=None):
    """Format every displayed price with its day change in parentheses.

    Example: 2,328 (+2.96%).  If pct is supplied it is used directly;
    otherwise it is calculated versus yesterday's price.
    """
    price = to_float(price, 0)
    if price <= 0:
        return "-"
    if pct is None:
        yesterday = to_float(yesterday, 0)
        pct = pct_change(price, yesterday) if yesterday > 0 else 0.0
    else:
        pct = to_float(pct, 0)
    return f"{price:,.0f} ({pct:+.2f}%)"


def estimate_trade_value_toman(shares, price_rial):
    return to_float(shares, 0) * to_float(price_rial, 0) / 10.0


def real_buyer_power(buy_volume, sell_volume, buy_count, sell_count):
    """Buyer/seller mean-volume ratio; use cumulative day counts for daily BP.

    With count deltas this is only a proxy: traders already counted earlier
    may also trade in the interval, so the deltas are not active-trader counts.
    Missing/zero inputs must not turn into zero or infinite buyer power.
    """
    values = [to_float(x, 0) for x in (buy_volume, sell_volume, buy_count, sell_count)]
    if any(not math.isfinite(x) or x <= 0 for x in values):
        return None
    buy_volume, sell_volume, buy_count, sell_count = values
    return (buy_volume / buy_count) / (sell_volume / sell_count)


def format_real_buyer_power(buy_volume, sell_volume, buy_count, sell_count, previous):
    bp = real_buyer_power(buy_volume, sell_volume, buy_count, sell_count)
    old_bp = real_buyer_power(
        previous.get("real_buy_volume"), previous.get("real_sell_volume"),
        previous.get("real_buy_count"), previous.get("real_sell_count"),
    )
    if bp is None:
        return "قدرت خریدار امروز: نامشخص"
    display = alert_rtl_change(f"{old_bp:.2f}", f"{bp:.2f}") if old_bp is not None else alert_ltr(f"{bp:.2f}")
    return f"قدرت خریدار امروز: {display}"


def format_interval_buyer_power(buy, sell, buyers, sellers):
    label = "قدرت خریدار بازه (تخمینی): "
    if buyers == 0 and sellers == 0:
        return label + "نامشخص؛ خریدار و فروشنده جدید اضافه نشدند"
    if buyers == 0:
        return label + "نامشخص؛ خریدار جدید اضافه نشد"
    if sellers == 0:
        return label + "نامشخص؛ فروشنده جدید اضافه نشد"
    proxy = real_buyer_power(buy, sell, buyers, sellers)
    if proxy is None:
        return label + "نامشخص؛ حجم خرید یا فروش بازه صفر است"
    return label + alert_ltr(f"{proxy:.2f}")


def alert_ltr(value):
    """Keep numeric fragments in their written order inside Persian Telegram text."""
    return "\u2066" + str(value) + "\u2069"


def alert_rtl_change(before, now):
    """RTL reading: previous at right, current at left; isolate numeric fragments."""
    return "\u2067" + alert_ltr(before) + " ← " + alert_ltr(now) + "\u2069"


def alert_shares(value):
    return alert_ltr(fmt_compact_num(value)) + " سهم"


def book_shares(value):
    amount = fmt_compact_num(value)
    unit = {"K": "هزار", "M": "میلیون", "B": "میلیارد"}.get(amount[-1])
    return f"{amount[:-1]} {unit} سهم" if unit else amount + " سهم"


def alert_money(value):
    value = to_float(value, 0)
    if abs(value) >= 1_000_000_000:
        return alert_ltr(f"{value/1_000_000_000:.2f}") + " میلیارد تومان"
    if abs(value) >= 1_000_000:
        return alert_ltr(f"{value/1_000_000:.1f}") + " میلیون تومان"
    if abs(value) >= 1_000:
        return alert_ltr(f"{value/1_000:.1f}") + " هزار تومان"
    return alert_ltr(f"{value:,.0f}") + " تومان"


def format_estimated_side_average(side, volume, new_people, price, group="حقیقی"):
    """Show total and estimated per-new-person flow together, for either group."""
    line = f"{side} {group}: {alert_shares(volume)}"
    if price > 0:
        line += " ≈ " + alert_money(estimate_trade_value_toman(volume, price))
    if new_people > 0:
        average = volume/new_people
        line += f"\nسرانه تخمینی ({alert_ltr(f'{new_people:,.0f}')} نفر جدید): {alert_shares(average)}"
        if price > 0:
            line += " ≈ " + alert_money(estimate_trade_value_toman(average, price))
    elif group != "حقیقی":
        person = "خریدار" if side == "خرید" else "فروشنده"
        line += f"\n{person} جدید اضافه نشد"
    return line


def format_market_alert(title, symbol, captured_at, sections, minutes=None):
    time_line = "🕒 " + alert_ltr(str(captured_at).split(".")[0])
    if minutes is not None:
        if minutes < 1:
            time_line += " · " + alert_ltr(f"{minutes*60:.0f}") + " ثانیه"
        else:
            time_line += " · " + alert_ltr(f"{minutes:.1f}") + " دقیقه"
    header = f"{title} — {symbol_tag(symbol)}\n{time_line}"
    return "\n\n".join([header] + [s.strip() for s in sections if s and s.strip()])


def format_trade_window_alert(conn, live, reference, captured_at, title, extra="", flow_reference=None,
                              flow_segments=None, residual=False, duration_override=None):
    """Price window and received ClientType updates are deliberately labelled separately."""
    current = live.get("current") or {}
    reference = reference or {}
    sample = published_client_snapshot(conn, live)
    flow_reference = flow_reference if flow_reference is not None else reference
    change = client_interval(sample, flow_reference)
    if flow_segments is not None:
        change = sum_flow_segments(flow_segments)
        flow_reference = flow_segments[0][0] if flow_segments else {}
    price = to_float(current.get("pDrCotVal"), 0)
    old_price = to_float(reference.get("last_price"), 0)
    yesterday = to_float(current.get("priceYesterday"), 0)
    price_at = live.get("price_received_at") or captured_at
    cur_sec, ref_sec = _hms_to_seconds(price_at), _hms_to_seconds(reference.get("captured_at"))
    minutes = (cur_sec-ref_sec)/60 if cur_sec is not None and ref_sec is not None and cur_sec >= ref_sec else None
    if duration_override is not None:
        minutes = duration_override
    current_display = fmt_price_pct(price, yesterday)
    display = (alert_rtl_change(fmt_price_pct(old_price, yesterday), current_display)
               if old_price > 0 and old_price != price else alert_ltr(current_display))
    sections = [f"قیمت: {display}"]
    if change and change["classified_volume"] > 0:
        total = change["classified_volume"]
        flow_end = flow_segments[-1][1] if flow_segments else sample
        clock = alert_rtl_change(str(flow_reference["captured_at"]).split(".")[0],
                                 str(flow_end["captured_at"]).split(".")[0])
        note = "جمع بخش‌های گزارش‌نشده در پنجرهٔ ۵ دقیقه\n" if flow_segments and len(flow_segments) > 1 else ""
        volume_label = "حجم گزارش‌نشده" if residual else "حجم بازه"
        sections.append(f"زمان دریافت آمار: {clock}\n"
                        + note + f"{volume_label}: {alert_shares(total)} ≈ {alert_money(estimate_trade_value_toman(total, price))}")
        real_lines = [format_estimated_side_average(side, change[f"real_{key}_volume"],
                       change[f"real_{key}_count"], price)
                      for side, key in (("خرید", "buy"), ("فروش", "sell"))
                      if change[f"real_{key}_volume"] > 0]
        sections.append("\n\n".join(real_lines))
        legal = format_legal_flow(change["legal_buy_volume"], change["legal_sell_volume"],
                                  price, change["legal_buy_count"], change["legal_sell_count"])
        if legal:
            sections.append(legal)
        net = change["real_buy_volume"]-change["real_sell_volume"]
        if net:
            label = "خالص ورود حقیقی" if net > 0 else "خالص خروج حقیقی"
            sections.append(f"{label}: {alert_shares(abs(net))} ≈ {alert_money(estimate_trade_value_toman(abs(net), price))}")
        sections.append(format_interval_buyer_power(change["real_buy_volume"], change["real_sell_volume"],
                                                    change["real_buy_count"], change["real_sell_count"]))
    if sample:
        daily = [format_real_buyer_power(sample["real_buy_volume"], sample["real_sell_volume"],
                                        sample["real_buy_count"], sample["real_sell_count"],
                                        flow_reference)]
        day_volume = sample["classified_volume"]
        if change and change["classified_volume"] > 0:
            daily.append("سهم بازه از حجم کل: " + alert_ltr(f'{100*change["classified_volume"]/day_volume:.2f}%'))
        line = "حجم کل: " + alert_shares(day_volume)
        average, sessions = previous_volume_average(conn, live.get("symbol") or "?",
                                                     to_int(current.get("dEven"), 0))
        if average:
            digits = str(sessions).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
            line += f"؛ {alert_ltr(f'{day_volume/average:.2f}')} برابر میانگین{digits}روزه"
        daily.append(line)
        sections.append("\n".join(daily))
    if extra:
        sections.append(extra)
    return format_market_alert(title, live.get("symbol") or "?", price_at, sections, minutes=minutes)


def format_legal_flow(buy_v, sell_v, price_rial, buy_count_delta=0, sell_count_delta=0, label="حقوقی"):
    return "\n\n".join(format_estimated_side_average(side, volume, count, price_rial, label)
                          for side, volume, count in (("خرید", buy_v, buy_count_delta),
                                                       ("فروش", sell_v, sell_count_delta))
                          if volume > 0)


def format_book_alert(title, symbol, captured_at, yesterday, label="صف",
                      before=None, now=None, before_price=0, now_price=0, traded=None):
    preopen = is_preopen(captured_at)
    price_label = "قیمت سفارش" if preopen else "قیمت"
    if preopen:
        title += " — پیش‌گشایش"
    lines = []
    if before_price > 0 and now_price > 0 and before_price != now_price:
        lines.append(price_label + ": " + alert_rtl_change(fmt_price_pct(before_price, yesterday), fmt_price_pct(now_price, yesterday)))
    elif max(before_price, now_price) > 0:
        lines.append(price_label + ": " + alert_ltr(fmt_price_pct(now_price or before_price, yesterday)))
    if before is not None and now is not None:
        # Keep Persian units on both operands even if the client strips direction isolates.
        lines.append(f"حجم {label}: " + alert_rtl_change(book_shares(before), book_shares(now)))
        if before > 0:
            lines.append("تغییر: " + alert_ltr(f"{pct_change(now, before):+.1f}%"))
    elif now is not None:
        lines.append(f"حجم {label}: {alert_shares(now)}")
    elif before is not None:
        lines.append(f"حجم {label} قبلی: {alert_shares(before)}")
    if before is not None and now is not None and before_price > 0 and now_price > 0:
        lines.append("ارزش: " + alert_rtl_change(
            fmt_toman(estimate_trade_value_toman(before, before_price)),
            fmt_toman(estimate_trade_value_toman(now, now_price))) + " تومان")
    elif now is not None and now_price > 0:
        lines.append("ارزش: " + alert_money(estimate_trade_value_toman(now, now_price)))
    elif before is not None and before_price > 0:
        lines.append("ارزش قبلی: " + alert_money(estimate_trade_value_toman(before, before_price)))
    if not preopen and traded is not None and traded > 0:
        lines.append(f"معامله بازه: {alert_shares(traded)}")
    return format_market_alert(title, symbol, captured_at, ["\n".join(lines)])


def book_state(book):
    rows = [r for r in (book or []) if isinstance(r, dict)]
    total_buy = sum(max(0.0, to_float(r.get("qTitMeDem"), r.get("buy_volume", 0))) for r in rows)
    total_sell = sum(max(0.0, to_float(r.get("qTitMeOf"), r.get("sell_volume", 0))) for r in rows)
    best_buy = max([to_float(r.get("pMeDem"), r.get("buy_price", 0)) for r in rows] + [0.0])
    sell_prices = [to_float(r.get("pMeOf"), r.get("sell_price", 0)) for r in rows]
    sell_prices = [p for p in sell_prices if p > 0]
    best_sell = min(sell_prices) if sell_prices else 0.0
    if total_buy > 0 and total_sell <= 0:
        return "buy_queue", total_buy, best_buy
    if total_sell > 0 and total_buy <= 0:
        return "sell_queue", total_sell, best_sell
    return "normal", max(total_buy, total_sell), 0.0


def visible_queue_order_count(book, queue_kind):
    """Active order counts as reported by visible levels, not unique people.

    A count is usable only if every level with positive queue-side volume has
    a positive, finite order count. Missing depth/count means omit the average.
    """
    if queue_kind not in ("buy_queue", "sell_queue"):
        return None
    volume_key, count_key = (("qTitMeDem", "zOrdMeDem") if queue_kind == "buy_queue"
                             else ("qTitMeOf", "zOrdMeOf"))
    fallback_volume, fallback_count = (("buy_volume", "buy_count") if queue_kind == "buy_queue"
                                       else ("sell_volume", "sell_count"))
    total = 0
    present = False
    for row in book or []:
        if not isinstance(row, dict):
            continue
        volume = to_float(row.get(volume_key, row.get(fallback_volume)), 0)
        if volume <= 0:
            continue
        value = row.get(count_key, row.get(fallback_count))
        count = to_float(value, float("nan"))
        if not math.isfinite(count) or count <= 0 or int(count) != count:
            return None
        present = True
        total += int(count)
    return total if present and total > 0 else None


def average_queue_order_value_toman(volume, price, order_count):
    if not order_count or volume <= 0 or price <= 0:
        return None
    return estimate_trade_value_toman(volume, price)/order_count


def official_volume_between_books(conn, symbol, date_value, from_book_at, to_book_at):
    """Conservative executed-volume ceiling for a queue-change interval.

    Only same-day, fresh reception timestamps are compared; if either official
    observation is stale, missing, non-monotonic or temporally reversed, return
    None rather than inferring cancellations from an unreliable time window.
    Total trades are NOT necessarily trades at the queue price.
    """
    t0, t1 = _hms_to_seconds(from_book_at), _hms_to_seconds(to_book_at)
    if t0 is None or t1 is None or t1 <= t0:
        return None
    before = alert_sample(conn, symbol, date_value, "observation", upper=from_book_at)
    after = alert_sample(conn, symbol, date_value, "observation", upper=to_book_at)
    if not before or not after:
        return None
    b0, b1 = _hms_to_seconds(before.get("captured_at")), _hms_to_seconds(after.get("captured_at"))
    if b0 is None or b1 is None or b1 <= b0:
        return None
    if (t0-b0 > QUEUE_TRADE_MATCH_MAX_LAG_SECONDS
            or t1-b1 > QUEUE_TRADE_MATCH_MAX_LAG_SECONDS
            or b0 > t0 or b1 > t1):
        return None
    v0 = to_float(before.get("volume"), float("nan"))
    v1 = to_float(after.get("volume"), float("nan"))
    if not (math.isfinite(v0) and math.isfinite(v1) and v1 >= v0 >= 0):
        return None
    m0 = to_int(before.get("market_time"), 0)
    m1 = to_int(after.get("market_time"), 0)
    if m0 and m1 and m1 < m0:
        return None
    return v1-v0


def format_queue_change_alert(title, symbol, clock, yesterday,
                              before_volume, after_volume, before_price, after_price,
                              before_book, after_book, queue_kind, traded=None):
    """Compact, evidence-calibrated Telegram queue change alert."""
    lines = []
    price = after_price or before_price
    if price > 0:
        lines.append("قیمت: " + alert_ltr(fmt_price_pct(price, yesterday)))
    change = pct_change(after_volume, before_volume)
    volume_text = alert_rtl_change(book_shares(before_volume), book_shares(after_volume))
    if change is not None:
        volume_text += " (" + alert_ltr(f"{change:+.1f}%") + ")"
    lines.append("حجم صف: " + volume_text)
    if before_price > 0 and after_price > 0:
        lines.append("ارزش صف: " + alert_rtl_change(
            fmt_toman(estimate_trade_value_toman(before_volume, before_price)),
            fmt_toman(estimate_trade_value_toman(after_volume, after_price))) + " تومان")
    old_average = average_queue_order_value_toman(
        before_volume, before_price, visible_queue_order_count(before_book, queue_kind))
    new_average = average_queue_order_value_toman(
        after_volume, after_price, visible_queue_order_count(after_book, queue_kind))
    if old_average is not None and new_average is not None:
        lines.append("میانگین ارزش هر سفارش: " + alert_rtl_change(
            fmt_toman(old_average), fmt_toman(new_average)) + " تومان")
    elif new_average is not None:
        lines.append("میانگین ارزش هر سفارش: " + alert_money(new_average))
    if after_volume < before_volume and traded is not None:
        # A trade on another price cannot remove more than 'traded' shares from
        # this queue. New orders may obscure *additional* cancellations.
        nontrade_lower_bound = max(0, before_volume-after_volume-traded)
        if nontrade_lower_bound >= max(10_000, 0.10*(before_volume-after_volume)):
            lines.append("حداقل کاهش غیرمعامله‌ای: " + book_shares(nontrade_lower_bound))
    return format_market_alert(title, symbol, clock, ["\n".join(lines)])


def pct_emoji(x):
    x = to_float(x, 0)
    if x > 0:
        return "🟢"
    if x < 0:
        return "🔴"
    return "⚪️"



def alert_threshold_shares(conn, symbol, fraction=None, date_value=None):
    fraction = ALERT_DAILY_VOLUME_FRACTION if fraction is None else fraction
    day = date_value or int(tehran_now().strftime("%Y%m%d"))
    average, _ = previous_volume_average(conn, symbol, day)
    return max(float(ALERT_MIN_SHARES), (average or 0)*fraction)


def alert_allowed(conn, key, cooldown_seconds=None):
    now = time.time()
    cooldown = ALERT_COOLDOWN_SECONDS if cooldown_seconds is None else max(1, int(cooldown_seconds))
    last = to_float(meta_get(conn, f"alert_ts:{key}", "0"), 0)
    if now - last < cooldown:
        return False
    meta_set(conn, f"alert_ts:{key}", str(now))
    return True


def current_has_trades(current):
    market_time = to_int(current.get("hEven") or current.get("lastHEven"), 0)
    return (to_float(current.get("qTotTran5J"), 0) > 0
            and to_float(current.get("pDrCotVal"), 0) > 0 and market_time >= TRADING_START_HHMMSS)


def is_preopen(captured_at):
    seconds = _hms_to_seconds(captured_at)
    return seconds is not None and 8 * 3600 + 45 * 60 <= seconds < 9 * 3600


def real_flow_signal(change, price, threshold):
    """Per-side money bands after the existing volume/concentration conditions.

    The 150m estimated average exception is only for purchases below 200m.
    A side with no new people still has a valid volume/value but no invented average.
    """
    candidates = []
    for key, side in (("buy", "خرید"), ("sell", "فروش")):
        volume = change[f"real_{key}_volume"]
        count = change[f"real_{key}_count"]
        if volume < threshold or count > CONCENTRATED_MAX_NEW_PEOPLE or price <= 0:
            continue
        value = estimate_trade_value_toman(volume, price)
        if value > REAL_FLOW_HEAVY_TOMAN:
            candidates.append((2, side, "سنگین"))
        elif value >= REAL_FLOW_NOTICE_TOMAN:
            candidates.append((1, side, "قابل‌توجه"))
        elif key == "buy" and count > 0 and value / count >= SMALL_BUY_MIN_AVERAGE_TOMAN:
            candidates.append((0, side, "با سرانه قابل‌توجه"))
    if not candidates:
        return None
    rank = max(c[0] for c in candidates)
    selected = [c for c in candidates if c[0] == rank]
    side = " و ".join(c[1] for c in selected)
    emoji = "🚨" if rank == 2 else "🔎"
    if rank == 0:
        return "🔎 خرید حقیقی با سرانه قابل‌توجه"
    return f"{emoji} {side} {selected[0][2]} حقیقی"


class MarketAlert(str):
    """A normal Telegram string with durable delivery/coverage metadata."""
    def __new__(cls, text, report):
        value = super().__new__(cls, text)
        value.report = report
        value.report_id = hashlib.sha256((text + json.dumps(report, sort_keys=True,
                                         ensure_ascii=False)).encode()).hexdigest()
        return value


def market_targets():
    return list(dict.fromkeys(str(t).strip() for t in (TELEGRAM_CHANNEL_ID, TELEGRAM_CHAT_ID) if str(t or "").strip()))


def coverage_reports(conn, symbol, day, include_pending=False):
    targets = market_targets()
    target = targets[0] if targets else "preview"
    rows = []
    tables = ("market_alert_receipts", "market_alert_outbox") if include_pending else ("market_alert_receipts",)
    for table in tables:
        rows.extend(conn.execute(f"""SELECT payload,created,rowid FROM {table}
            WHERE target=? AND symbol=? AND date=?""", (target, symbol, day)).fetchall())
    return [json.loads(row["payload"]) for row in sorted(rows, key=lambda r:(r["created"], r["rowid"]))]


def sum_flow_segments(segments):
    total = dict.fromkeys(list(CLIENT_FIELDS) + ["classified_volume"], 0)
    for before, after in segments:
        delta = client_interval(after, before)
        if delta is None:
            return None
        for key in total:
            total[key] += delta[key]
    return total if segments else None


def unreported_flow(conn, symbol, day, before, after, planned=(), reserve_pending=False):
    """Subtract the union of actually delivered intervals, preserving exact counters."""
    if client_interval(after, before) is None:
        return []
    floor = meta_get(conn, f"coverage_client_floor:{day}:{symbol}")
    if floor:
        floor = json.loads(floor)
        if floor.get("epoch", 0) != after.get("epoch", 0):
            return []
        if floor["captured_at"] > before["captured_at"]:
            before = floor
    if client_interval(after, before) is None:
        return []
    remaining = [(before, after)]
    # Pending ranges remain in the durable retry queue, rather than being enqueued
    # a second time. They are never recorded as successfully delivered coverage.
    reports = coverage_reports(conn, symbol, day, include_pending=reserve_pending) + list(planned)
    for report in reports:
        for start, end in report.get("flow", []):
            if start.get("epoch", 0) != after.get("epoch", 0):
                continue
            updated = []
            for left, right in remaining:
                if end["captured_at"] <= left["captured_at"] or start["captured_at"] >= right["captured_at"]:
                    updated.append((left, right))
                    continue
                if left["captured_at"] < start["captured_at"] and client_interval(start, left) is not None:
                    updated.append((left, start))
                if end["captured_at"] < right["captured_at"] and client_interval(right, end) is not None:
                    updated.append((end, right))
            remaining = updated
    return [(a, b) for a, b in remaining if (client_interval(b, a) or {}).get("classified_volume", 0) > 0]


def fresh_price_reference(conn, symbol, day, reference, current):
    anchors = [reference]
    floor = meta_get(conn, f"coverage_price_floor:{day}:{symbol}")
    if floor:
        anchors.append(json.loads(floor))
    for report in coverage_reports(conn, symbol, day, include_pending=True):
        if report.get("price"):
            anchors.append(report["price"][1])
    return max((r for r in anchors if r and r["captured_at"] <= current["captured_at"]),
               key=lambda r: r["captured_at"])


def client_reference_for_price(conn, symbol, day, reference, sample):
    if not sample or not reference:
        return None
    row = alert_sample(conn, symbol, day, "client", lower=reference["captured_at"],
                       upper=sample["captured_at"], oldest=True)
    return row if client_interval(sample, row) is not None else None


def phase_volume_allowed(conn, symbol, day, reference, sample, price):
    client_ref = client_reference_for_price(conn, symbol, day, reference, sample)
    change = client_interval(sample, client_ref)
    average, _ = previous_volume_average(conn, symbol, day)
    return bool(change and average and change["classified_volume"] >= average * PHASE_MIN_VOLUME_FRACTION
                and estimate_trade_value_toman(change["classified_volume"], price) >= PHASE_MIN_VALUE_TOMAN)


def compose_market_reports(conn, live, flow_candidates, price_candidates, observation, sample):
    """One body per reported interval, with all coincident reasons in that body."""
    symbol = live["symbol"]
    day = to_int(live["current"].get("dEven"), 0)
    price = observation["last_price"]
    prepared = []
    # The short interval is handled first. Longer windows then see only the remainder.
    for candidate in sorted(flow_candidates, key=lambda r: r["reference"]["captured_at"], reverse=True):
        ref = candidate["reference"]
        segments = unreported_flow(conn, symbol, day, ref, sample, [p["report"] for p in prepared], reserve_pending=True)
        delta = sum_flow_segments(segments)
        kind, title = candidate["kind"], candidate["title"]
        if delta is None:
            # Only the exact same interval can contribute another reason to its message.
            for old in prepared:
                if old["reference"]["captured_at"] == ref["captured_at"] and kind != "volume":
                    old_delta = sum_flow_segments(old["segments"])
                    if kind == "real":
                        title = real_flow_signal(old_delta, price, candidate["threshold"])
                    elif kind == "legal":
                        net = old_delta["legal_buy_volume"] - old_delta["legal_sell_volume"]
                        title = (("🏢 خرید خالص سنگین حقوقی" if net > 0 else "🏢 فروش خالص سنگین حقوقی")
                                 if abs(net) >= candidate["threshold"] else None)
                    if title and title not in old["extras"] and title != old["title"]:
                        old["extras"].append(title)
            continue
        if kind == "real":
            title = real_flow_signal(delta, price, candidate["threshold"])
            valid = title is not None
        elif kind == "legal":
            net = delta["legal_buy_volume"] - delta["legal_sell_volume"]
            valid = abs(net) >= candidate["threshold"]
            title = "🏢 خرید خالص سنگین حقوقی" if net > 0 else "🏢 فروش خالص سنگین حقوقی"
        else:
            valid = delta["classified_volume"] >= candidate["threshold"]
        if not valid:
            continue
        residual = delta["classified_volume"] != client_interval(sample, ref)["classified_volume"]
        reference = dict(ref)
        reference["captured_at"] = segments[0][0]["captured_at"]
        reference["last_price"] = segments[0][0].get("last_price", price)
        report = dict(symbol=symbol, date=day, flow=segments)
        if reference.get("last_price", price) != price and segments[-1][1]["captured_at"] == sample["captured_at"]:
            report["price"] = [reference, observation]
        prior_volume = next((r for r in prepared if r.get("kind") == "volume"), None) if kind == "volume" else None
        if prior_volume:
            pieces = sorted(prior_volume["segments"] + segments, key=lambda r:r[0]["captured_at"])
            merged = []
            for a, b in pieces:
                if merged and merged[-1][1]["captured_at"] == a["captured_at"]:
                    merged[-1] = (merged[-1][0], b)
                else:
                    merged.append((a,b))
            prior_volume["segments"] = merged
            prior_volume["report"]["flow"] = merged
            prior_volume["reference"] = merged[0][0]
            prior_volume["residual"] = True
            if merged[0][0].get("last_price", price) != price:
                prior_volume["report"]["price"] = [merged[0][0], observation]
        else:
            prepared.append(dict(title=title, reference=reference, segments=segments, residual=residual,
                                 report=report, extras=[], kind=kind))

    groups = defaultdict(list)
    for candidate in price_candidates:
        groups[candidate["direction"]].append(candidate)
    for choices in groups.values():
        selected = max(choices, key=lambda r: r["priority"])
        ref = selected["reference"]
        # Reuse a simultaneous flow message rather than reproducing its numbers.
        combined = next((r for r in prepared if r["segments"] and not r.get("price_signal")
                         and r["segments"][-1][1]["captured_at"] == sample["captured_at"]), None) if sample else None
        if combined is None:
            client_ref = client_reference_for_price(conn, symbol, day, ref, sample)
            segments = unreported_flow(conn, symbol, day, client_ref, sample,
                                       [p["report"] for p in prepared], reserve_pending=True) if client_ref else []
            combined = dict(title=selected["title"], reference=ref, segments=segments,
                            residual=False, report=dict(symbol=symbol, date=day, flow=segments), extras=[])
            prepared.append(combined)
        else:
            combined["extras"].append(combined["title"])
            combined["title"] = selected["title"]
            combined["reference"] = ref
        combined["report"]["price"] = [ref, observation]
        combined["price_signal"] = True
        for item in choices:
            if item is not selected:
                combined["extras"].append(item["title"])

    messages = []
    for item in prepared:
        extras = list(dict.fromkeys(s.split(" ", 1)[-1] for s in item["extras"]))
        text = format_trade_window_alert(conn, live, item["reference"], observation["captured_at"],
                    item["title"], extra=" · ".join(extras), flow_segments=item["segments"], residual=item["residual"],
                    duration_override=(sum(_hms_to_seconds(b["captured_at"])-_hms_to_seconds(a["captured_at"])
                                           for a,b in item["segments"])/60
                                       if item["segments"] and not item.get("price_signal") else None))
        messages.append(MarketAlert(text, item["report"]))
    return messages


def acknowledge_market_alert(conn, message, target):
    r = message.report
    conn.execute("INSERT OR IGNORE INTO market_alert_receipts VALUES(?,?,?,?,?,?,?)",
                 (target, message.report_id, r["symbol"], r["date"], str(message),
                  json.dumps(r, ensure_ascii=False), time.time()))
    conn.execute("DELETE FROM market_alert_outbox WHERE target=? AND report_id=?", (target, message.report_id))
    conn.commit()


def deliver_market_alert(conn, message):
    if not isinstance(message, MarketAlert):
        return telegram_broadcast_message(message)
    sent = False
    r = message.report
    for target in market_targets():
        if conn.execute("SELECT 1 FROM market_alert_receipts WHERE target=? AND report_id=?",
                        (target, message.report_id)).fetchone():
            continue
        conn.execute("INSERT OR IGNORE INTO market_alert_outbox VALUES(?,?,?,?,?,?,?)",
                     (target, message.report_id, r["symbol"], r["date"], str(message),
                      json.dumps(r, ensure_ascii=False), time.time()))
        conn.commit()
        if telegram_send_message(str(message), chat_id=target):
            acknowledge_market_alert(conn, message, target)
            sent = True
    return sent


def retry_market_outbox(conn):
    day = int(tehran_now().strftime("%Y%m%d"))
    targets = market_targets()
    if not targets:
        return
    rows = conn.execute("""SELECT * FROM market_alert_outbox WHERE date=?
        ORDER BY (target=?) DESC,created,rowid LIMIT 20""", (day, targets[0])).fetchall()
    for row in rows:
        if row["target"] not in targets:
            continue
        if not telegram_send_message(row["text"], chat_id=row["target"]):
            break
        acknowledge_market_alert(conn, MarketAlert(row["text"], json.loads(row["payload"])), row["target"])


def detect_fast_alerts(conn, live, captured_at):
    """Independent price, classified-flow and book signals with separate Fast baselines."""
    alerts = []
    current = live.get("current") or {}
    symbol = live.get("symbol") or "?"
    date_value = to_int(current.get("dEven"), 0)
    if date_value != int(tehran_now().strftime("%Y%m%d")):
        return alerts
    price_at = live.get("price_received_at") or captured_at
    client_at = live.get("client_received_at") or captured_at
    book_at = live.get("book_received_at") or captured_at
    price_yesterday = to_float(current.get("priceYesterday"), 0)
    last_price = to_float(current.get("pDrCotVal"), 0)
    official_volume = to_float(current.get("qTotTran5J"), 0)
    previous_price = alert_sample(conn, symbol, date_value, "price")
    prev = alert_sample(conn, symbol, date_value, "observation")
    official_delta = official_volume-prev["volume"] if prev else None
    market_time = to_int(current.get("hEven") or current.get("lastHEven"), 0)
    if prev and (official_delta < 0 or market_time < prev.get("market_time", 0)):
        official_delta = None

    # The received clock also gates trading signals, even if the API carries old counters.
    trading_active = current_has_trades(current) and not is_preopen(price_at) and not is_preopen(client_at)
    sample = accept_client_snapshot(conn, live, client_at) if trading_active else None
    if not trading_active:
        live["_client_sample"] = None
        live["_client_status"] = "BEFORE_TRADE"
        live["_flow_reference"] = None
    flow_ref = live.get("_flow_reference")
    change = client_interval(sample, flow_ref)
    threshold = alert_threshold_shares(conn, symbol, date_value=date_value)
    flow_candidates = []
    price_candidates = []
    observation = dict(captured_at=price_at, volume=official_volume,
                       last_price=last_price, market_time=market_time)
    # A first activation starts here, even when an old collector populated this DB.
    floor_key = f"coverage_price_floor:{date_value}:{symbol}"
    if trading_active and not meta_get(conn, floor_key):
        meta_set(conn, floor_key, json.dumps(observation, ensure_ascii=False))
        previous_price = None
    client_floor_key = f"coverage_client_floor:{date_value}:{symbol}"
    floor = json.loads(meta_get(conn, client_floor_key, "null"))
    if sample and (not floor or floor.get("epoch", 0) != sample.get("epoch", 0)):
        meta_set(conn, client_floor_key, json.dumps(sample, ensure_ascii=False))
        change = None
        flow_ref = None
    if change:
        flow_title = real_flow_signal(change, last_price, threshold)
        if flow_title and alert_allowed(conn, f"{symbol}:realflow"):
            flow_candidates.append(dict(reference=flow_ref, title=flow_title, kind="real", threshold=threshold))
        legal_net = change["legal_buy_volume"] - change["legal_sell_volume"]
        legal_threshold = alert_threshold_shares(conn, symbol, LEGAL_FLOW_DAILY_VOLUME_FRACTION, date_value)
        if abs(legal_net) >= legal_threshold:
            side = "buy" if legal_net > 0 else "sell"
            if alert_allowed(conn, f"{symbol}:legalflow:{side}", LEGAL_FLOW_COOLDOWN_SECONDS):
                title = "🏢 خرید خالص سنگین حقوقی" if legal_net > 0 else "🏢 فروش خالص سنگین حقوقی"
                flow_candidates.append(dict(reference=flow_ref, title=title, kind="legal", threshold=legal_threshold))

        sample_threshold = alert_threshold_shares(conn, symbol, VOLUME_BURST_SAMPLE_FRACTION, date_value)
        window_threshold = alert_threshold_shares(conn, symbol, VOLUME_BURST_WINDOW_FRACTION, date_value)
        seconds = _hms_to_seconds(client_at)
        window_ref = alert_sample(conn, symbol, date_value, "client",
                                  lower=_seconds_to_hms(seconds-VOLUME_BURST_WINDOW_SECONDS),
                                  before=client_at, oldest=True)
        window_change = client_interval(sample, window_ref)
        bursts = []
        if change["classified_volume"] >= sample_threshold:
            bursts.append((flow_ref, sample_threshold))
        if window_change and window_change["classified_volume"] >= window_threshold:
            bursts.append((window_ref, window_threshold))
        pending_key = f"coverage_volume_pending:{date_value}:{symbol}"
        pending = json.loads(meta_get(conn, pending_key, "null"))
        if pending:
            pending_ref = pending["reference"]
            if window_ref and pending_ref["captured_at"] < window_ref["captured_at"]:
                pending_ref = window_ref
            if client_interval(sample, pending_ref) is not None:
                bursts.append((pending_ref, pending["threshold"]))
        if bursts and alert_allowed(conn, f"{symbol}:volume_burst", VOLUME_BURST_COOLDOWN_SECONDS):
            for ref, minimum in bursts:
                flow_candidates.append(dict(reference=ref, title="📊 حجم قابل‌توجه در بازهٔ کوتاه",
                                            kind="volume", threshold=minimum))
            meta_set(conn, pending_key, "null")
        elif bursts:
            ref, minimum = min(bursts, key=lambda r:r[0]["captured_at"])
            meta_set(conn, pending_key, json.dumps(dict(reference=ref, threshold=minimum), ensure_ascii=False))
        average, _ = previous_volume_average(conn, symbol, date_value)
        if average and flow_ref["classified_volume"]/average < VOLUME_RATIO_ALERT <= sample["classified_volume"]/average:
            if alert_allowed(conn, f"{symbol}:volumeratio"):
                flow_candidates.append(dict(reference=flow_ref, title="🔥 حجم غیرعادی", kind="ratio", threshold=1))

    price_valid = trading_active and price_yesterday > 0
    if previous_price and (price_at <= previous_price["captured_at"]
                           or official_volume < previous_price["volume"]
                           or market_time < previous_price["market_time"]):
        price_valid = False
    if price_valid and not previous_price:
        # Starting/restarting mid-session must not fabricate an opening event.
        legacy_active = conn.execute("""SELECT 1 FROM client_intraday
            WHERE symbol=? AND date=? AND volume>0 AND market_time>='09:00:00' LIMIT 1""",
                                     (symbol, date_value)).fetchone()
        opening_key = f"opening_seen:{date_value}:{symbol}"
        witnessed_open = ((prev is not None and prev["volume"] <= 0) or
                          (prev is None and not legacy_active and 90000 <= market_time <= 90200))
        if witnessed_open and not meta_get(conn, opening_key):
            opening = to_float(current.get("priceFirst"), 0) or last_price
            opening_pct = pct_change(opening, price_yesterday)
            direction = "مثبت" if opening_pct > 0 else "منفی" if opening_pct < 0 else "بدون تغییر"
            lines = ["قیمت گشایش: " + alert_ltr(fmt_price_pct(opening, price_yesterday))]
            if last_price != opening:
                lines.append("آخرین: " + alert_ltr(fmt_price_pct(last_price, price_yesterday)))
            alerts.append(format_market_alert(f"{pct_emoji(opening_pct)} گشایش {direction}",
                                             symbol, price_at, ["\n".join(lines)]))
        meta_set(conn, opening_key, "1")
    elif price_valid and previous_price:
        fresh_ref = fresh_price_reference(conn, symbol, date_value, previous_price, observation)
        previous_last = fresh_ref["last_price"]
        last_pct = pct_change(last_price, price_yesterday)
        previous_pct = pct_change(previous_last, price_yesterday)
        move = pct_change(last_price, previous_last)
        volume_ok = change is None or change["classified_volume"] == 0 or change["classified_volume"] >= max(50_000, threshold*.35)
        if abs(move) >= SUDDEN_PRICE_PCT_30S and volume_ok:
            up = move > 0
            key = "priceup" if up else "pricedown"
            if alert_allowed(conn, f"{symbol}:{key}"):
                price_candidates.append(dict(title="🚀 جهش رو به بالا" if up else "🔻 ریزش ناگهانی",
                                             reference=fresh_ref, direction=up, priority=2))
        price_ref = momentum_reference_snapshot(conn, symbol, date_value, price_at, kind="price")
        if price_ref:
            price_ref = fresh_price_reference(conn, symbol, date_value, price_ref, observation)
            client_ref = client_reference_for_price(conn, symbol, date_value, price_ref, sample)
            momentum_change = client_interval(sample, client_ref)
            volume = momentum_change["classified_volume"] if momentum_change else None
            movement = pct_change(last_price, price_ref["last_price"])
            ref_pct = pct_change(price_ref["last_price"], price_yesterday)
            momentum_threshold = alert_threshold_shares(conn, symbol, MOMENTUM_DAILY_VOLUME_FRACTION, date_value)
            volume_ok = volume is None or volume == 0 or volume >= momentum_threshold
            if abs(movement) >= MOMENTUM_PRICE_PCT and volume_ok:
                up = movement > 0
                key = "momentumup" if up else "momentumdown"
                if alert_allowed(conn, f"{symbol}:{key}", MOMENTUM_ALERT_COOLDOWN_SECONDS):
                    title = ("🟢 تقویت مثبت / خرید رو به بالا" if ref_pct > 0 and last_pct > 0
                             else "📈 خرید رو به بالا / مومنتوم مثبت") if up else "📉 فروش رو به پایین / مومنتوم منفی"
                    price_candidates.append(dict(title=title, reference=price_ref, direction=up, priority=1))
            strengthen_threshold = alert_threshold_shares(conn, symbol, POSITIVE_STRENGTHEN_VOLUME_FRACTION, date_value)
            if (ref_pct > 0 and last_pct > 0 and last_pct-ref_pct >= POSITIVE_STRENGTHEN_PCT
                    and movement < MOMENTUM_PRICE_PCT
                    and (volume is None or volume == 0 or volume >= strengthen_threshold)
                    and alert_allowed(conn, f"{symbol}:positive_strengthen", POSITIVE_STRENGTHEN_COOLDOWN_SECONDS)):
                price_candidates.append(dict(title="🟢 تقویت مثبت", reference=price_ref, direction=True, priority=1))
        if phase_volume_allowed(conn, symbol, date_value, fresh_ref, sample, last_price):
            if previous_pct <= 0 < last_pct and last_pct-previous_pct >= .40:
                if alert_allowed(conn, f"{symbol}:turnedgreen"):
                    price_candidates.append(dict(title="🟢 تغییر فاز به مثبت", reference=fresh_ref, direction=True, priority=3))
            elif previous_pct >= 0 > last_pct and previous_pct-last_pct >= .40:
                if alert_allowed(conn, f"{symbol}:turnedred"):
                    price_candidates.append(dict(title="🔴 تغییر فاز به منفی", reference=fresh_ref, direction=False, priority=3))

    alerts.extend(compose_market_reports(conn, live, flow_candidates, price_candidates, observation, sample))
    observation = dict(captured_at=price_at, volume=official_volume,
                       last_price=last_price, market_time=market_time)
    if price_valid:
        save_alert_sample(conn, symbol, date_value, "price", observation)
    # Store raw official observations separately, solely for order-execution estimates.
    if not prev or price_at > prev["captured_at"]:
        save_alert_sample(conn, symbol, date_value, "observation", observation)

    queue_shrink_reported = set()
    prev_book_for_state = previous_book_snapshot(conn, symbol, date_value, book_at)
    cur_book_for_state = live.get("best") or []
    if prev_book_for_state and cur_book_for_state:
        prev_state, prev_depth, prev_q_price = book_state(prev_book_for_state)
        cur_state, cur_depth, cur_q_price = book_state(cur_book_for_state)
        queue_threshold = alert_threshold_shares(conn, symbol, QUEUE_MIN_DAILY_VOLUME_FRACTION)

        if cur_state != prev_state:
            if cur_state == "buy_queue" and cur_depth >= queue_threshold:
                if alert_allowed(conn, f"{symbol}:buyqueue"):
                    alerts.append(format_book_alert(
                        "🟢 صف خرید تشکیل شد", symbol, book_at, price_yesterday,
                        now=cur_depth, now_price=cur_q_price
                    ))
            elif cur_state == "sell_queue" and cur_depth >= queue_threshold:
                if alert_allowed(conn, f"{symbol}:sellqueue"):
                    alerts.append(format_book_alert(
                        "🔴 صف فروش تشکیل شد", symbol, book_at, price_yesterday,
                        now=cur_depth, now_price=cur_q_price
                    ))
            elif prev_state == "buy_queue" and cur_state == "normal":
                if alert_allowed(conn, f"{symbol}:buyqueuebreak"):
                    alerts.append(format_book_alert(
                        "⚠️ صف خرید باز شد", symbol, book_at, price_yesterday,
                        before=prev_depth, before_price=prev_q_price
                    ))
            elif prev_state == "sell_queue" and cur_state == "normal":
                if alert_allowed(conn, f"{symbol}:sellqueuebreak"):
                    alerts.append(format_book_alert(
                        "⚠️ صف فروش جمع شد/باز شد", symbol, book_at, price_yesterday,
                        before=prev_depth, before_price=prev_q_price
                    ))

        # Queue size can change gradually without changing state. Compare with a
        # 3-5 minute reference so both heavier and lighter queues are reported.
        q_ref_book = queue_reference_snapshot(conn, symbol, date_value, book_at)
        if q_ref_book:
            ref_state, ref_depth, ref_q_price = book_state(q_ref_book)
            if cur_state == ref_state and cur_state in ("buy_queue", "sell_queue") and ref_depth > 0:
                q_delta = cur_depth - ref_depth
                q_change_pct = (q_delta / ref_depth) * 100.0
                q_abs_threshold = max(50_000.0, alert_threshold_shares(conn, symbol, QUEUE_CHANGE_MIN_DAILY_VOLUME_FRACTION))
                # Do not re-announce a 3-5 minute change on a later *unchanged*
                # 30-second book; the event may already have happened earlier.
                latest_depth_changed = abs(cur_depth-prev_depth) >= 1.0
                if latest_depth_changed and abs(q_change_pct) >= QUEUE_CHANGE_PCT and abs(q_delta) >= q_abs_threshold:
                    if cur_state == "sell_queue":
                        if q_delta > 0:
                            q_title = "🔴 صف فروش سنگین‌تر شد"
                            q_key = "sellqueuegrow"
                        else:
                            q_title = "🟡 صف فروش سبک‌تر شد"
                            q_key = "sellqueueshrink"
                    else:
                        if q_delta > 0:
                            q_title = "🟢 صف خرید سنگین‌تر شد"
                            q_key = "buyqueuegrow"
                        else:
                            q_title = "🟡 صف خرید سبک‌تر شد"
                            q_key = "buyqueueshrink"

                    if alert_allowed(conn, f"{symbol}:{q_key}", QUEUE_CHANGE_COOLDOWN_SECONDS):
                        traded_in_window = (official_volume_between_books(
                            conn, symbol, date_value, q_ref_book[0]["captured_at"], book_at)
                            if q_delta < 0 and not is_preopen(book_at) else None)
                        alerts.append(format_queue_change_alert(
                            q_title, symbol, book_at, price_yesterday,
                            ref_depth, cur_depth, ref_q_price, cur_q_price,
                            q_ref_book, cur_book_for_state, cur_state, traded=traded_in_window
                        ))
                        if q_delta < 0:
                            queue_shrink_reported.add("sell" if cur_state == "sell_queue" else "buy")

        # Significant SELL additions outside queue states. A newly visible deeper
        # price is not evidence of a new order; it may have been outside the old depth.
        if prev_state == cur_state == "normal":
            old_sells = {to_float(r.get("sell_price"), 0): to_float(r.get("sell_volume"), 0)
                         for r in prev_book_for_state if to_float(r.get("sell_price"), 0) > 0}
            worst_visible = max(old_sells, default=0)
            additions = []
            order_threshold = alert_threshold_shares(conn, symbol, ORDER_ALERT_DAILY_VOLUME_FRACTION, date_value)
            for row in cur_book_for_state:
                price = to_float(row.get("pMeOf"), 0)
                volume = to_float(row.get("qTitMeOf"), 0)
                if price <= 0 or price > worst_visible:
                    continue
                before = old_sells.get(price, 0)
                if volume - before >= order_threshold:
                    additions.append((price, before, volume))
            if additions and alert_allowed(conn, f"{symbol}:sellincrease"):
                if len(additions) == 1:
                    price, before, volume = additions[0]
                    alerts.append(format_book_alert(
                        "🔴 افزایش قابل‌توجه سفارش فروش", symbol, book_at, price_yesterday,
                        label="سفارش", before=before, now=volume, before_price=price, now_price=price))
                else:
                    preopen = is_preopen(book_at)
                    title = "🔴 افزایش قابل‌توجه سفارش فروش" + (" — پیش‌گشایش" if preopen else "")
                    sections = []
                    for price, before, volume in sorted(additions):
                        label = "قیمت سفارش" if preopen else "قیمت"
                        sections.append(label + ": " + alert_ltr(fmt_price_pct(price, price_yesterday))
                                        + "\nحجم سفارش: " + alert_rtl_change(book_shares(before), book_shares(volume))
                                        + "\nافزایش: " + alert_shares(volume-before) + " ≈ "
                                        + alert_money(estimate_trade_value_toman(volume-before, price)))
                    alerts.append(format_market_alert(title, symbol, book_at, sections))

    # Order deletion / movement: compare same-price visible levels with previous 30s book.
    prev_book = previous_book_snapshot(conn, symbol, date_value, book_at)
    cur_book = live.get("best") or []
    if prev_book and cur_book and official_delta is not None:
        cur_by_buy = {to_float(r.get("pMeDem"), 0): r for r in cur_book if isinstance(r, dict) and to_float(r.get("pMeDem"), 0) > 0}
        cur_by_sell = {to_float(r.get("pMeOf"), 0): r for r in cur_book if isinstance(r, dict) and to_float(r.get("pMeOf"), 0) > 0}

        day_volume = to_float(current.get("qTotTran5J"), 0)
        prev_day_volume = to_float(prev.get("volume"), 0) if prev else 0
        traded_delta = official_delta
        oth = alert_threshold_shares(conn, symbol, ORDER_ALERT_DAILY_VOLUME_FRACTION)

        for old in prev_book:
            bp = to_float(old.get("buy_price"), 0)
            bv = to_float(old.get("buy_volume"), 0)
            if bp > 0 and bv >= oth:
                newv = to_float((cur_by_buy.get(bp) or {}).get("qTitMeDem"), 0)
                drop = bv - newv
                if ("buy" not in queue_shrink_reported and drop >= oth
                        and traded_delta < drop * 0.5
                        and alert_allowed(conn, f"{symbol}:buycancel:{int(bp)}")):
                    alerts.append(format_book_alert(
                        "⚠️ حذف/جابجایی مشکوک سفارش خرید", symbol, book_at, price_yesterday,
                        label="سفارش", before=bv, now=newv,
                        before_price=bp, now_price=bp, traded=traded_delta
                    ))

            sp = to_float(old.get("sell_price"), 0)
            sv = to_float(old.get("sell_volume"), 0)
            if sp > 0 and sv >= oth:
                newv = to_float((cur_by_sell.get(sp) or {}).get("qTitMeOf"), 0)
                drop = sv - newv
                if ("sell" not in queue_shrink_reported and drop >= oth
                        and traded_delta < drop * 0.5
                        and alert_allowed(conn, f"{symbol}:sellcancel:{int(sp)}")):
                    alerts.append(format_book_alert(
                        "⚠️ حذف/جابجایی مشکوک سفارش فروش", symbol, book_at, price_yesterday,
                        label="سفارش", before=sv, now=newv,
                        before_price=sp, now_price=sp, traded=traded_delta
                    ))

    return alerts
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
    price_received_at = received_clock()

    refresh_share_context(conn, symbol, inscode, current)

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
    client_received_at = None

    if FETCH_CLIENT_TYPE:
        p = api_get(
            f"ClientType/GetClientType/{inscode}/1/0",
            symbol,
            "ClientTypeCurrent",
        )
        client = unwrap_dict(p, ("clientType",))
        client_received_at = received_clock()

    # --------------------------------------------------------
    # Best limits
    # --------------------------------------------------------
    best = []
    book_received_at = None

    if FETCH_BEST_LIMITS:
        p = api_get(
            f"BestLimits/{inscode}",
            symbol,
            "BestLimitsCurrent",
        )
        best = unwrap_list(p, ("bestLimits",))
        book_received_at = received_clock()

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
        "capture_source": "heavy",
        "price_received_at": price_received_at,
        "client_received_at": client_received_at,
        "book_received_at": book_received_at,
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
    official_volume = to_float(current.get("qTotTran5J"), 0)
    sample = published_client_snapshot(conn, live)
    live_volume = sample["classified_volume"] if sample else None
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

    avgvol20_prev, avgvol20_sessions = previous_volume_average(conn, symbol, market_date, 20)
    avgvol60_prev, avgvol60_sessions = previous_volume_average(conn, symbol, market_date, 60)

    tech = enriched[-1] if enriched else {}

    # If today's bar has not traded, technical values come from last traded day.
    last_hist = previous_rows[-1] if previous_rows else tech

    technical_source = tech
    provisional = False

    if official_volume > 0 and enriched and to_int(tech.get("date"), 0) == market_date:
        provisional = not bool(tech.get("finalized"))
    else:
        technical_source = last_hist

    buy_i_vol = sample["real_buy_volume"] if sample else None
    sell_i_vol = sample["real_sell_volume"] if sample else None
    buy_n_vol = sample["legal_buy_volume"] if sample else None
    sell_n_vol = sample["legal_sell_volume"] if sample else None

    buy_i_count = sample["real_buy_count"] if sample else None
    sell_i_count = sample["real_sell_count"] if sample else None

    buy_i_value = client_metric(client, "buy_I_Value")
    sell_i_value = client_metric(client, "sell_I_Value")
    raw_sample = client_snapshot(live)
    values_match_sample = sample and raw_sample["ready"] and all(
        sample[key] == raw_sample[key] for key in CLIENT_FIELDS
    )

    real_net_vol = buy_i_vol-sell_i_vol if sample else None

    real_net_value = (
        buy_i_value - sell_i_value
        if values_match_sample and (buy_i_value or sell_i_value)
        else None
    )

    real_net_approx = (
        real_net_vol * live_last
        if sample and live_last > 0 else None
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
        "OfficialVolume": official_volume,
        "VolumeBasis": "CLIENT_TYPE",
        "ClientStatus": "READY" if sample else "UNAVAILABLE",
        "ClientReceivedAt": sample["captured_at"] if sample else None,
        "LiveValue_Rial": live_value,
        "LiveValue_BnToman": live_value / 1e10,
        "LiveTrades": live_trades,

        "AvgVol20_Previous": avgvol20_prev,
        "AvgVol60_Previous": avgvol60_prev,
        "AvgVol20_Sessions": avgvol20_sessions,
        "AvgVol60_Sessions": avgvol60_sessions,
        "VolumeAverageBasis": ("SHARE_COUNT_COMPARABLE" if share_count_on_date(conn, live["inscode"], market_date)
                               else "RAW_NO_SHARE_CONTEXT"),
        "SharesOutstanding": share_count_on_date(conn, live["inscode"], market_date),
        "VolumeRatio20_Live": safe_div(live_volume, avgvol20_prev),
        "VolumeRatio60_Live": safe_div(live_volume, avgvol60_prev),

        "RealBuyVolume": buy_i_vol,
        "RealSellVolume": sell_i_vol,
        "RealNetVolume": real_net_vol,
        "RealBuyCount": buy_i_count,
        "RealSellCount": sell_i_count,
        "LegalBuyCount": sample["legal_buy_count"] if sample else None,
        "LegalSellCount": sample["legal_sell_count"] if sample else None,
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
        "LegalNetVolume": buy_n_vol-sell_n_vol if sample else None,

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
        ("SUMMARY", "LiveVolume/volume ratios use ClientType; OfficialVolume/LiveValue/LiveTrades retain the official feed. Daily history retains official finalized data."),
        ("MARKET_INDICES", "Current TSETMC شاخص کل and شاخص کل هم‌وزن on every run."),
        ("History export", "Set TSETMC_EXPORT_HISTORY=1 to include DAILY_260D, MONTHLY_ALL and INDEX_HISTORY_260D."),
        ("PV_5M_TODAY", "Today's real TSETMC trade-price path in 5-minute buckets (TradeHistory/TradeIntraDay)."),
        ("PV_QUALITY", "Intraday price-path quality; official daily volume still comes from ClosingPriceInfo."),
        ("CLIENT_INTRADAY", "Raw endpoint receipt times and source; volume=official, classified_volume=one ClientType side. Missing/stale packets carry quality flags."),
        ("CLIENT_CURRENT", "Current real/legal participation."),
        ("BEST_LIMITS", "Current top bid/ask levels."),
        ("BEST_LIMITS_INTRADAY", "30-second order-book history for cancellation/movement analysis."),
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

        if EXPORT_HISTORY_SHEETS:
            enriched = enrich_history(get_daily_rows(conn, symbol))
            all_daily.extend(compact_daily_rows(enriched, symbol, DAILY_CONTEXT_DAYS))
            all_monthly.extend(monthly_history(enriched, symbol))

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
    if EXPORT_HISTORY_SHEETS:
        write_sheet(wb, "INDEX_HISTORY_260D", index_history or [], rtl=True)

    write_sheet(wb, "PV_5M_TODAY", all_pv)
    write_sheet(wb, "PV_QUALITY", all_quality)
    if EXPORT_HISTORY_SHEETS:
        write_sheet(wb, "DAILY_260D", all_daily)
        write_sheet(wb, "MONTHLY_ALL", all_monthly)

    market_dates = [to_int((x.get("current") or {}).get("dEven"), 0) for x in live_data]
    market_dates = [d for d in market_dates if d]
    client_intraday = get_client_intraday_rows(conn, max(market_dates) if market_dates else None)
    write_sheet(wb, "CLIENT_INTRADAY", client_intraday)

    best_intraday = get_best_limits_intraday_rows(conn, max(market_dates) if market_dates else None)
    write_sheet(wb, "BEST_LIMITS_INTRADAY", best_intraday)

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


def telegram_send_message(text, chat_id=None):
    target = str(chat_id or TELEGRAM_CHAT_ID).strip()
    if not TELEGRAM_BOT_TOKEN or not target:
        return False
    try:
        s = telegram_session()
        r = s.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
            data={"chat_id": target, "text": text},
            timeout=30,
        )
        r.raise_for_status()
        return r.json().get("ok") is True
    except Exception as e:
        print(f"Telegram message failed for {target}:", e)
        return False


def telegram_send_document(path, caption="", chat_id=None):
    target = str(chat_id or TELEGRAM_CHAT_ID).strip()
    if not TELEGRAM_BOT_TOKEN or not target or not Path(path).exists():
        return False
    try:
        s = telegram_session()
        with open(path, "rb") as f:
            r = s.post(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
                data={"chat_id": target, "caption": caption},
                files={"document": f},
                timeout=120,
            )
        r.raise_for_status()
        return True
    except Exception as e:
        print(f"Telegram document failed for {target}:", e)
        return False


def telegram_broadcast_message(text):
    """Send intraday market alerts to the private chat and configured channel."""
    targets = []
    for target in (TELEGRAM_CHAT_ID, TELEGRAM_CHANNEL_ID):
        target = str(target or "").strip()
        if target and target not in targets:
            targets.append(target)
    ok = False
    for target in targets:
        ok = telegram_send_message(text, chat_id=target) or ok
    return ok


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



def build_status_text(conn, live_data):
    lines = ["📊 TSETMC — وضعیت لحظه‌ای"]
    for live in live_data or []:
        s = make_summary(conn, live)
        symbol = s.get("Symbol") or "?"
        last_price = to_float(s.get("LastPrice"), 0)
        close_price = to_float(s.get("ClosingPrice"), 0)
        last_pct = to_float(s.get("LastChangePct"), 0)
        close_pct = to_float(s.get("ClosingChangePct"), 0)
        volume = to_float(s.get("LiveVolume"), 0)
        vr20 = s.get("VolumeRatio20_Live")
        real_net = to_float(s.get("RealNetVolume"), 0)

        lines.append("")
        lines.append(f"▪️ {symbol_tag(symbol)}")
        lines.append(f"آخرین: {fmt_price_pct(last_price, pct=last_pct)} {pct_emoji(last_pct)}")
        lines.append(f"پایانی: {fmt_price_pct(close_price, pct=close_pct)} {pct_emoji(close_pct)}")
        if s.get("LiveVolume") is None:
            continue
        vr_txt = f"{to_float(vr20,0):.2f}×" if vr20 is not None else "-"
        lines.append(f"حجم کل: {fmt_compact_num(volume)} | {vr_txt} میانگین ۲۰روزه")
        if real_net:
            lines.append(f"خالص حقیقی: {fmt_compact_num(real_net)} {'🟢' if real_net>0 else '🔴'}")
        legal_buy = to_float(s.get("LegalBuyVolume"), 0)
        legal_sell = to_float(s.get("LegalSellVolume"), 0)
        legal_net = to_float(s.get("LegalNetVolume"), 0)
        legal_parts = [f"{side} {alert_shares(amount)}"
                       for side, amount in (("خرید", legal_buy), ("فروش", legal_sell)) if amount > 0]
        if legal_net:
            legal_side = "خالص خرید" if legal_net > 0 else "خالص فروش"
            legal_value = alert_money(estimate_trade_value_toman(abs(legal_net), last_price))
            legal_parts.append(f"{legal_side} {alert_shares(abs(legal_net))} ≈ {legal_value}")
        if legal_parts:
            lines.append("حقوقی: " + " | ".join(legal_parts))

        buy_count = to_float(s.get("RealBuyCount"), 0)
        sell_count = to_float(s.get("RealSellCount"), 0)
        buy_vol = to_float(s.get("RealBuyVolume"), 0)
        sell_vol = to_float(s.get("RealSellVolume"), 0)

        avg_buy_shares = safe_div(buy_vol, buy_count)
        avg_sell_shares = safe_div(sell_vol, sell_count)
        avg_buy_toman = estimate_trade_value_toman(avg_buy_shares or 0, last_price)
        avg_sell_toman = estimate_trade_value_toman(avg_sell_shares or 0, last_price)

        if buy_count > 0 or sell_count > 0:
            lines.append(
                f"سرانه حقیقی: خرید {fmt_toman(avg_buy_toman)} | فروش {fmt_toman(avg_sell_toman)} تومان"
            )
            if avg_buy_toman > 0 and avg_sell_toman > 0:
                if avg_buy_toman >= avg_sell_toman:
                    lines.append(f"برتری سرانه: خریدار {avg_buy_toman/avg_sell_toman:.2f}× 🟢")
                else:
                    lines.append(f"برتری سرانه: فروشنده {avg_sell_toman/avg_buy_toman:.2f}× 🔴")

        # Only surface exceptional quick-decision context.
        gap = last_pct - close_pct
        if abs(gap) >= 1.0:
            direction = "آخرین بالاتر از پایانی" if gap > 0 else "آخرین پایین‌تر از پایانی"
            lines.append(f"⚠️ {direction}: {gap:+.2f} واحد درصد")
        if vr20 is not None and to_float(vr20,0) >= 2.0:
            lines.append(f"🔥 حجم غیرعادی: {to_float(vr20,0):.2f}× میانگین ۲۰روزه")

    return "\n".join(lines)


def split_telegram_text(text, max_units=3500):
    """Split at paragraph/line boundaries within Telegram's UTF-16 text budget."""
    def units(value):
        return len(value.encode("utf-16-le")) // 2

    chunks = []
    current = ""
    for paragraph in str(text).strip().split("\n\n"):
        candidate = current + ("\n\n" if current else "") + paragraph
        if units(candidate) <= max_units:
            current = candidate
            continue
        if current:
            chunks.append(current)
            current = ""
        if units(paragraph) <= max_units:
            current = paragraph
            continue
        for line in paragraph.splitlines():
            candidate = current + ("\n" if current else "") + line
            if units(candidate) <= max_units:
                current = candidate
                continue
            if current:
                chunks.append(current)
                current = ""
            for char in line:
                if units(current + char) > max_units:
                    chunks.append(current)
                    current = ""
                current += char
    if current:
        chunks.append(current)
    return chunks


def eod_client_history(conn, symbol, market_date):
    accepted = conn.execute("""SELECT payload FROM alert_samples
        WHERE symbol=? AND date=? AND kind='client' ORDER BY captured_at""",
                            (symbol, market_date)).fetchall()
    if accepted:
        return [json.loads(row["payload"]) for row in accepted]
    # Compatibility for a pre-upgrade database. Raw legacy data is never used for live alert baselines.
    rows = conn.execute("""SELECT * FROM client_intraday WHERE symbol=? AND date=? AND volume>0
        ORDER BY captured_at""", (symbol, market_date)).fetchall()
    result = []
    for row in map(dict, rows):
        if any(row.get(k) is None or not math.isfinite(row[k]) or row[k] < 0 for k in CLIENT_FIELDS):
            continue
        total = row["real_buy_volume"]+row["legal_buy_volume"]
        if total <= 0 or total != row["real_sell_volume"]+row["legal_sell_volume"]:
            continue
        if result and any(row[k] < result[-1][k] for k in CLIENT_FIELDS):
            continue
        row.update(ready=True, classified_volume=total, epoch=0)
        result.append(row)
    return result


def eod_bp_history_line(rows, final_bp):
    """Use only actual same-day samples; never synthesize missing morning BP."""
    readings = []
    for row in rows:
        bp = real_buyer_power(
            row.get("real_buy_volume"), row.get("real_sell_volume"),
            row.get("real_buy_count"), row.get("real_sell_count"),
        )
        seconds = _hms_to_seconds(row.get("captured_at"))
        if bp is not None and seconds is not None:
            readings.append((seconds, row["captured_at"], bp))
    if final_bp is None:
        return "قدرت خریدار امروز: نامشخص"
    if not readings:
        return "قدرت خریدار در پایان: " + alert_ltr(f"{final_bp:.2f}")
    if len(readings) == 1:
        return "قدرت خریدار در پایان: " + alert_ltr(f"{final_bp:.2f}")

    selected = [readings[0]]
    # A representative middle reading and a late reading show the day's shape
    # without pretending that missing hourly checkpoints were observed.
    span = readings[-1][0]-readings[0][0]
    if span >= 1800:
        for target in ((readings[0][0]+readings[-1][0])/2, readings[-1][0]-900):
            closest = min(readings, key=lambda item: abs(item[0]-target))
            if closest not in selected and closest[0] < readings[-1][0]-120:
                selected.append(closest)
    selected.sort(key=lambda item: item[0])
    points = [alert_ltr(f"{clock[:5]} ({bp:.2f})") for _, clock, bp in selected]
    points.append("پایان " + alert_ltr(f"{final_bp:.2f}"))
    return "قدرت خریدار امروز: " + "، ".join(points)


def eod_late_reference(rows):
    if len(rows) < 2:
        return None
    last_sec = _hms_to_seconds(rows[-1].get("captured_at"))
    if last_sec is None:
        return None
    candidates = [r for r in rows if _hms_to_seconds(r.get("captured_at")) is not None]
    if not candidates:
        return None
    reference = min(candidates, key=lambda r: abs(_hms_to_seconds(r["captured_at"])-(last_sec-900)))
    return reference if last_sec-_hms_to_seconds(reference["captured_at"]) >= 180 else None


def book_volume_at_price(book, price, side):
    price_key, volume_key = ("pMeDem", "qTitMeDem") if side == "buy" else ("pMeOf", "qTitMeOf")
    return sum(max(0, to_float(row.get(volume_key), row.get(f"{side}_volume", 0)))
               for row in book if to_float(row.get(price_key), row.get(f"{side}_price", 0)) == price)


def eod_previous_low_offer(conn, symbol, day):
    """Use the previous traded session's closing book when it was recorded locally."""
    prior = conn.execute("""SELECT * FROM daily WHERE symbol=? AND date<? AND volume>0
        ORDER BY date DESC LIMIT 1""", (symbol, day)).fetchone()
    if not prior or not prior["finalized"] or prior["last_price"] != prior["price_min"]:
        return 0
    price = prior["last_price"]
    if price <= 0 or prior["price_yesterday"] <= 0 or pct_change(price, prior["price_yesterday"]) > -2:
        return 0
    rows = conn.execute("""SELECT * FROM best_limits_intraday WHERE symbol=? AND date=?
        AND captured_at=(SELECT MAX(captured_at) FROM best_limits_intraday
            WHERE symbol=? AND date=? AND captured_at>='12:25:00')""",
                        (symbol, prior["date"], symbol, prior["date"])).fetchall()
    offered = book_volume_at_price(map(dict, rows), price, "sell")
    if offered >= max(ALERT_MIN_SHARES, 0.10 * prior["volume"]):
        return offered
    return 0


def eod_headline(conn, symbol, day, summary, rows, book, bp):
    last = summary["LastPrice"]
    high, low, yesterday = (to_float(summary.get(key), 0) for key in ("High", "Low", "Yesterday"))
    volume = summary["LiveVolume"]
    ratio = to_float(summary.get("VolumeRatio20_Live"), 0)
    day_pct = summary["LastChangePct"]
    minimum_order = max(ALERT_MIN_SHARES, 0.10 * volume)
    buy_at_high = book_volume_at_price(book, high, "buy") if last == high else 0
    sell_at_low = book_volume_at_price(book, low, "sell") if last == low else 0
    details = []
    if buy_at_high:
        details.append("سفارش خرید در سقف: " + alert_shares(buy_at_high))
    if sell_at_low:
        details.append("سفارش فروش در کف: " + alert_shares(sell_at_low))

    if (yesterday > 0 and 0 < low <= 0.99 * yesterday and high > low
            and last >= 1.01 * yesterday and last >= 0.998 * high):
        details.insert(0, "کف روز: " + alert_ltr(fmt_price_pct(low, yesterday)))
        if 0 < ratio <= 0.60:
            return "برگشت از کف با حجم کمتر از معمول", details
        return ("برگشت از کف تا سقف روز" if last == high else "برگشت از کف تا نزدیکی سقف روز"), details
    if buy_at_high >= max(minimum_order, 0.25 * volume) and day_pct >= 2:
        prior_offer = eod_previous_low_offer(conn, symbol, day)
        if prior_offer:
            details.insert(0, "عرضه در کف جلسهٔ قبل: " + alert_shares(prior_offer))
            return "از عرضه در کف جلسهٔ قبل تا تقاضا در سقف", details
        return "سفارش خرید پرحجم در سقف", details
    has_offer = any(0 < to_float(row.get("pMeOf"), 0) <= high
                    and to_float(row.get("qTitMeOf"), 0) > 0 for row in book)
    if yesterday > 0 and high >= 1.02 * yesterday and last <= 0.99 * high and has_offer:
        prior_queue = conn.execute("""SELECT 1 FROM best_limits_intraday WHERE symbol=? AND date=?
            AND captured_at>='09:00:00' GROUP BY captured_at
            HAVING MAX(sell_volume)=0 AND SUM(CASE WHEN buy_price=? THEN buy_volume ELSE 0 END)>=?
            LIMIT 1""", (symbol, day, high, minimum_order)).fetchone()
        if prior_queue:
            details.insert(0, "سقف روز: " + alert_ltr(fmt_price_pct(high, yesterday)))
            return "صف خرید باز شد و قیمت از سقف عقب نشست", details
    if sell_at_low >= minimum_order and day_pct <= -2:
        return "عرضه پرحجم در کف روز", details
    if ratio >= 2:
        return "حجم معاملات بالاتر از معمول", details
    if ratio and ratio <= 0.60 and day_pct >= 1:
        return "رشد قیمت با حجم کمتر از معمول", details
    if rows and bp is not None:
        first = rows[0]
        first_bp = real_buyer_power(first["real_buy_volume"], first["real_sell_volume"],
                                    first["real_buy_count"], first["real_sell_count"])
        if first_bp and bp >= 1.20 and bp >= 1.25 * first_bp:
            return "تقویت قدرت خریدار در طول روز", details
        if first_bp and bp <= 0.85 and bp <= 0.80 * first_bp:
            return "افت قدرت خریدار در طول روز", details
    if bp is not None and bp <= 0.85:
        return "سرانه فروش بالاتر از خرید", details
    if bp is not None and bp >= 1.20:
        return "سرانه خرید بالاتر از فروش", details
    if day_pct > 0 and last == high:
        return "پایان معاملات در سقف روز", details
    return "مرور معاملات روز", details


def build_channel_eod_messages(conn, live_data, report_complete=True):
    """Factual Persian close commentary from collected prices, flows, BP and book.

    This intentionally does not send a workbook or reproduce the operational
    /status message. It interprets observed data without inventing news or BP
    history missing after a restart.
    """
    now = tehran_now()
    date_value = int(now.strftime("%Y%m%d"))
    title = "📊 جمع‌بندی پایان بازار" if report_complete else "📊 جمع‌بندی پایان بازار — آخرین داده موجود"
    header = [title, "تاریخ: " + alert_ltr(jalali_date(now))]
    try:
        indices = json.loads(meta_get(conn, f"market_indices:{date_value}", "[]"))
    except (ValueError, TypeError):
        indices = []
    for row in indices:
        if row.get("Status") == "OK" and row.get("ChangePct") is not None:
            index_pct = alert_ltr("{:+.2f}%".format(to_float(row["ChangePct"], 0)))
            header.append(f"{row.get('Index')}: {index_pct}")

    by_symbol = {live.get("symbol"): live for live in live_data or [] if live.get("symbol")}
    sections = ["\n".join(header)]
    for symbol in SYMBOLS:
        live = by_symbol.get(symbol)
        if not live:
            sections.append(f"▪️ {symbol_tag(symbol)}: داده پایان بازار دریافت نشد.")
            continue
        try:
            summary = make_summary(conn, live)
            current = live.get("current") or {}
            day = to_int(current.get("dEven"), 0)
            volume = to_float(summary.get("LiveVolume"), 0)
            if day != date_value:
                sections.append(f"▪️ {symbol_tag(symbol)}: داده تازه امروز موجود نیست.")
                continue
            if to_float(summary.get("OfficialVolume"), 0) <= 0:
                sections.append(f"▪️ {symbol_tag(symbol)}: امروز معامله‌ای ثبت نشد.")
                continue
            last = to_float(summary.get("LastPrice"), 0)
            close = to_float(summary.get("ClosingPrice"), 0)
            if last <= 0:
                sections.append(f"▪️ {symbol_tag(symbol)}: قیمت پایان بازار دریافت نشد.")
                continue
            day_pct = to_float(summary.get("LastChangePct"), 0)
            if summary.get("LiveVolume") is None:
                sections.append(f"{pct_emoji(day_pct)} {symbol_tag(symbol)}\nآخرین: " + alert_ltr(fmt_price_pct(last, pct=day_pct)))
                continue
            bp = real_buyer_power(
                summary.get("RealBuyVolume"), summary.get("RealSellVolume"),
                summary.get("RealBuyCount"), summary.get("RealSellCount"),
            )
            rows = eod_client_history(conn, symbol, day)
            book = [row for row in (live.get("best") or []) if isinstance(row, dict)]
            headline, details = eod_headline(conn, symbol, day, summary, rows, book, bp)
            lines = [f"{pct_emoji(day_pct)} {symbol_tag(symbol)} — {headline}"]
            lines.append("آخرین: " + alert_ltr(fmt_price_pct(last, pct=day_pct)))
            if close > 0 and close != last:
                lines.append("پایانی: " + alert_ltr(fmt_price_pct(close, pct=summary.get("ClosingChangePct"))))

            ratio = summary.get("VolumeRatio20_Live")
            activity = "حجم کل: " + alert_shares(volume)
            if ratio is not None:
                activity += "؛ " + alert_ltr(f"{ratio:.2f}") + " برابر میانگین ۲۰روزه"
            lines.extend((activity, eod_bp_history_line(rows, bp)))

            real_net = to_float(summary.get("RealNetVolume"), 0)
            if real_net:
                side = "ورود خالص حقیقی" if real_net > 0 else "خروج خالص حقیقی"
                value = alert_money(estimate_trade_value_toman(abs(real_net), last))
                lines.append(f"{side}: {alert_shares(abs(real_net))} ≈ {value}")
            legal_net = to_float(summary.get("LegalNetVolume"), 0)
            if legal_net:
                side = "خریدار" if legal_net > 0 else "فروشنده"
                lines.append(f"حقوقی‌ها {side} خالص روز بودند.")

            reference = eod_late_reference(rows)
            if reference and all(rows[-1].get(key) == summary.get(column) for key, column in (
                ("real_buy_volume", "RealBuyVolume"), ("real_sell_volume", "RealSellVolume")
            )):
                late = client_interval(rows[-1], reference)
                if late:
                    late_net = late["real_buy_volume"] - late["real_sell_volume"]
                    if late_net != real_net and abs(late_net) >= max(ALERT_MIN_SHARES, 0.05 * volume):
                        side = "خرید" if late_net > 0 else "فروش"
                        clock = alert_ltr(reference["captured_at"][:5])
                        details.append(f"خالص {side} حقیقی از {clock}: {alert_shares(abs(late_net))}")

            lines.extend(details)
            sections.append("\n".join(lines))
        except Exception as exc:
            print(f"EOD commentary unavailable for {symbol_tag(symbol)}: {exc}")
            sections.append(f"▪️ {symbol_tag(symbol)}: داده کافی برای جمع‌بندی موجود نیست.")
    # One complete, independently forwardable message per symbol, in SYMBOLS order.
    messages = [sections[0]]
    date_line = "📅 " + alert_ltr(jalali_date(now))
    for section in sections[1:]:
        first, sep, rest = section.partition("\n")
        message = first + "\n" + date_line + ("\n" + rest if sep else "")
        message += "\n\nقیمت‌ها ریال؛ ارزش‌ها برآورد با قیمت آخر.\n📡 نبض تابلو | هشدار بورس"
        messages.append(message)
    return messages


def build_channel_eod_text(conn, live_data, report_complete=True):
    """Compatibility for local/private previews; channel delivery uses individual messages."""
    return "\n\n".join(build_channel_eod_messages(conn, live_data, report_complete))


def deliver_end_of_day_reports(conn, live_data, report_path=None, report_complete=True):
    """Private Excel/status; channel receives analytical text only, including fallback."""
    private_id = str(TELEGRAM_CHAT_ID or "").strip()
    channel_id = str(TELEGRAM_CHANNEL_ID or "").strip()
    result = {"private_document": False, "private_status": False, "channel_analysis": False}
    # Never leak private reports into a channel if the two secrets were mixed up.
    private_ok = private_id and not private_id.startswith(("-", "@")) and private_id != channel_id
    if private_ok:
        if report_path and Path(report_path).exists():
            caption = "گزارش نهایی بازار امروز" if report_complete else "آخرین گزارش کامل موجود امروز"
            result["private_document"] = telegram_send_document(report_path, caption, chat_id=private_id)
        try:
            status = build_status_text(conn, live_data)
            outcomes = [telegram_send_message(part, chat_id=private_id) for part in split_telegram_text(status)]
            result["private_status"] = bool(outcomes) and all(outcomes)
        except Exception as exc:
            print("Private final status failed:", exc)
    elif private_id:
        print("Private final reports skipped: configure a separate private TELEGRAM_CHAT_ID.")

    if channel_id:
        try:
            messages = build_channel_eod_messages(conn, live_data, report_complete=report_complete)
            outcomes = [telegram_send_message(part, chat_id=channel_id)
                        for message in messages for part in split_telegram_text(message)]
            result["channel_analysis"] = bool(outcomes) and all(outcomes)
        except Exception as exc:
            print("Channel end-of-day analysis failed:", exc)
    return result


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
    price_received_at = received_clock()

    refresh_share_context(conn, symbol, inscode, current)

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
    client_received_at = None
    if FETCH_CLIENT_TYPE:
        p = api_get(
            f"ClientType/GetClientType/{inscode}/1/0",
            symbol,
            "FAST_ClientTypeCurrent",
        )
        client = unwrap_dict(p, ("clientType",))
        client_received_at = received_clock()

    best = []
    book_received_at = None
    if FETCH_BEST_LIMITS:
        p = api_get(
            f"BestLimits/{inscode}",
            symbol,
            "FAST_BestLimitsCurrent",
        )
        best = unwrap_list(p, ("bestLimits",))
        book_received_at = received_clock()
        for r in best:
            if isinstance(r, dict):
                r["Symbol"] = symbol

    return {
        "symbol": symbol,
        "capture_source": "fast",
        "price_received_at": price_received_at,
        "client_received_at": client_received_at,
        "book_received_at": book_received_at,
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
    live_data = []
    retry_market_outbox(conn)

    for symbol in SYMBOLS:
        try:
            live = collect_fast_symbol(conn, symbol)
            live_data.append(live)
            captured_at = live.get("client_received_at") or received_clock()

            # Compare against the previous 30-second snapshot before persisting the current one.
            for alert_text in detect_fast_alerts(conn, live, captured_at):
                deliver_market_alert(conn, alert_text)

            save_client_intraday(conn, live, captured_at=captured_at)
            save_best_limits_intraday(conn, live)
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

    for symbol in SYMBOLS:
        print(f"Fetching {symbol} ...")
        try:
            live = collect_live_symbol(conn, symbol)
            live_data.append(live)
            save_client_intraday(conn, live)
            save_best_limits_intraday(conn, live)
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

    if any(row.get("Status") == "OK" for row in market_indices):
        try:
            meta_set(conn, f"market_indices:{int(now.strftime('%Y%m%d'))}",
                     json.dumps(market_indices, ensure_ascii=False))
        except Exception as exc:
            print("Market index summary cache failed:", exc)

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
            telegram_send_message(build_status_text(conn, live_data))
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
    print(f"TSETMC AUTOMATIC COLLECTOR {COLLECTOR_VERSION} — FAST 30s / HEAVY 5m")
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
        telegram_send_message("✅ Collector آماده است؛ Fast و سفارش‌ها هر ۳۰ ثانیه، Heavy هر ۵ دقیقه، هشدار خرید/فروش درشت، جهش قیمت، صف و حذف سفارش فعال است.")

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
        final_live = live_data
        report_complete = False
        try:
            final_path, final_live = run_once(conn, final_path)
            report_complete = True
        except Exception as e:
            print("Final report failed:", e)
            with report_lock:
                fallback = report_state.get("path")
            final_path = fallback if fallback and Path(fallback).exists() else None

        deliver_end_of_day_reports(
            conn, final_live, report_path=final_path, report_complete=report_complete
        )

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
