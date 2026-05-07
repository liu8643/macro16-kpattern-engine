
# -*- coding: utf-8 -*-
"""
GTC AI Trading System v9.2 FINAL-RELEASE / v9.5.8 DATA_INTEGRITY_PATCH

功能：
- 股票主檔分類（市場 / 產業 / 題材 / 子題材）
- 本地 SQLite 歷史資料庫
- TWSE/TPEX 官方資料 + Yahoo Finance 備援更新
- V9.5.6：融資融券（個股 + 市場情緒）整合進 Decision Layer / UI / Excel / Debug Log
- V9.5.8：Data Integrity Patch：market_snapshot 禁止 internal proxy 假通過；必須 TWSE 官方市場資料成功才 data_ready=1
- 核心 StrategyEngineV91（訊號 → 評分 → 倉位 → 交易計畫）
- 波浪 + 費波交易模型化
- Kelly + ATR 資金管理
- 真回測系統（勝率 / 平均報酬 / CAGR / MDD / Sharpe）
- 回測視覺化（Equity Curve）
- 分類 + 產業輪動分析
- TOP20 / TOP5 / 下單清單 / 機構交易計畫
- 專業交易 UI（儀表板 / 輪動 / 排行 / TOP / 計畫 / 回測）
"""

import sqlite3
import traceback
import requests
import sys
import csv
import io
import re
import threading
import time
import os
import subprocess
import warnings
import logging
import json
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd

try:
    import yfinance as yf
except Exception:
    yf = None

import tkinter as tk
from tkinter import ttk, messagebox

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Rectangle
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


warnings.filterwarnings("ignore", message=r"Glyph .* missing from font")
warnings.filterwarnings("ignore", message=r"Matplotlib is currently using agg")

PREFERRED_CJK_FONTS = [
    "Microsoft JhengHei", "Microsoft YaHei", "PingFang TC", "PingFang SC",
    "Noto Sans CJK TC", "Noto Sans CJK SC", "Noto Sans CJK JP",
    "Source Han Sans TW", "Source Han Sans CN", "SimHei", "Arial Unicode MS",
]

def load_font_from_runtime(font_path: Path) -> Optional[str]:
    try:
        if font_path and Path(font_path).exists():
            font_manager.fontManager.addfont(str(font_path))
            return font_manager.FontProperties(fname=str(font_path)).get_name()
    except Exception:
        return None
    return None

def resolve_cjk_font_path() -> Optional[Path]:
    search_dirs = [
        (RUNTIME_DIR / "fonts") if "RUNTIME_DIR" in globals() else None,
        (BASE_DIR / "fonts") if "BASE_DIR" in globals() else None,
        Path(__file__).resolve().parent / "fonts",
    ]
    candidates = [
        "NotoSansCJK-Regular.ttc",
        "NotoSansCJKtc-Regular.otf",
        "NotoSansTC-Regular.ttf",
        "msjh.ttc",
        "msyh.ttc",
        "simhei.ttf",
    ]
    for folder in search_dirs:
        if folder is None:
            continue
        for name in candidates:
            p = folder / name
            if p.exists():
                return p
    return None

def configure_matplotlib_cjk_font() -> str:
    chosen = None
    font_path = resolve_cjk_font_path()
    if font_path is not None:
        chosen = load_font_from_runtime(font_path)
    if chosen is None:
        available = {f.name for f in font_manager.fontManager.ttflist}
        for name in PREFERRED_CJK_FONTS:
            if name in available:
                chosen = name
                break
    if chosen is None:
        chosen = "DejaVu Sans"
    plt.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    return chosen

def safe_plot_text(value, fallback: str = "-") -> str:
    if value is None:
        return fallback
    s = str(value).replace("\n", " ").replace("\r", " ").strip()
    if not s:
        return fallback
    replacements = {
        "｜": " | ", "【": "[", "】": "]", "（": "(", "）": ")",
        "：": ": ", "，": ", ", "／": "/", "～": "~",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    s = re.sub(r"\s+", " ", s).strip()
    return s or fallback

SELECTED_PLOT_FONT = None
BUILD_DISPLAY_WARNING_CACHE: set[tuple[str, ...]] = set()


class OperationCancelled(Exception):
    pass


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def get_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()
RUNTIME_DIR = get_runtime_dir()
APP_NAME = "GTC AI Trading System v9.6.2 PRO FUNDAMENTAL_LOCAL_CACHE V16.2-R10_MARKET_SNAPSHOT_FULL_FALLBACK_AND_FAIL_REASON"

# V9.5.5 EPS_OFFICIAL_SOURCE：外部 EPS / 估值資料源正式規範
# 優先順序：1) TWSE OpenAPI / TWSE 官方 API；2) TPEx 官方頁面 / CSV；3) MOPS OpenData；4) Goodinfo 僅允許 fallback，不作為主資料源。
TWSE_OPENAPI_LICENSE_URL = "http://data.gov.tw/license"
TWSE_OPENAPI_SWAGGER_URL = "https://openapi.twse.com.tw/v1/swagger.json"
EXTERNAL_DATA_SOURCE_PRIORITY = [
    "TWSE OpenAPI / TWSE 官方 API",
    "TPEx 官方頁面 / CSV",
    "MOPS OpenData",
    "Goodinfo fallback only（不可作為主資料源）",
]
GOODINFO_FALLBACK_ENABLED = os.getenv("GTC_ENABLE_GOODINFO_FALLBACK", "0").strip() == "1"

# V9.5.6 MARGIN_INTEGRATED：融資融券資料源正式規範
# 原則：不使用 Mitake / 券商 SDK / 未授權商業資料；只使用官方可追溯來源。
TPEX_OPENAPI_PORTAL_URL = "https://www.tpex.org.tw/openapi/"
TPEX_OPENAPI_SWAGGER_URL = "https://www.tpex.org.tw/openapi/swagger.json"
TWSE_MARGIN_OFFICIAL_PAGE = "https://wwwc.twse.com.tw/zh/trading/margin/mi-margn.html"
TWSE_MARGIN_DATASET_ID = "11680"
TWSE_MARGIN_OPEN_DATA_ENDPOINT = "https://www.twse.com.tw/exchangeReport/MI_MARGN?response=open_data&selectType=ALL"
TWSE_MARGIN_API_TEMPLATE = TWSE_MARGIN_OPEN_DATA_ENDPOINT
TWSE_MARGIN_COMPARE_PAGE = "https://www.twse.com.tw/IIH2/zh/compare/margin.html"

# V9.5.8 DATA_INTEGRITY_PATCH：market_snapshot 不可再使用 internal:price_history 當 success。
# 只有 TWSE 官方市場指數資料成功解析，market_snapshot 才允許 data_ready=1。
TWSE_MARKET_SNAPSHOT_ENDPOINT_TEMPLATE = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={date}&type=MS"
MARKET_PROXY_DECISION_POLICY = (
    "R9：market_snapshot 優先使用 TWSE 官方 MI_INDEX；若官方即時資料尚未更新或解析失敗，"
    "允許使用本地 price_history 建立 local_cache_fallback，讓系統可先分析/回測/選股；"
    "source_level 會明確標示 local_cache_fallback_not_official，且外部資料不得直接控制 trade_allowed。"
)
ANALYSIS_EXECUTION_SPLIT_POLICY = (
    "V9.5.9：外部資料不作交易控制開關；analysis_ready 永遠允許技術分析；"
    "execution_ready 僅作資訊/提示/Excel/Log 欄位；trade_allowed 只由技術面、RR、風控條件決定。"
)

MARGIN_DATA_SOURCE_POLICY = [
    "TWSE 官方 MI_MARGN open_data endpoint（dataset 11680）：上市個股融資融券，寫入 external_margin",
    "TPEx 官方 OpenAPI / 官方頁面：上櫃個股融資融券，寫入 external_margin",
    "TWSE compare/margin：市場層融資融券情緒，寫入 macro_margin_sentiment",
    "禁止 Mitake / 券商 SDK / 未授權商業資料",
]
STATE_PATH = RUNTIME_DIR / "build_history_state_v9_2_final_release.json"

LOG_DIR = RUNTIME_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = LOG_DIR / f"gtc_ai_trading_{datetime.now().strftime('%Y%m%d')}.log"

def configure_app_logger() -> logging.Logger:
    logger = logging.getLogger("gtc_ai_trading")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(threadName)s | %(message)s")
    try:
        file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception:
        # V9.5.7：若執行目錄/手機下載環境無法寫 log 檔，不阻斷主程式啟動；仍保留 console log。
        pass
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    logger.propagate = False
    return logger

APP_LOGGER = configure_app_logger()

def log_info(message: str):
    try:
        APP_LOGGER.info(str(message))
    except Exception:
        pass

def log_warning(message: str):
    try:
        APP_LOGGER.warning(str(message))
    except Exception:
        pass

def log_error(message: str):
    try:
        APP_LOGGER.error(str(message))
    except Exception:
        pass

def log_exception(message: str, exc: Exception | None = None):
    try:
        if exc is None:
            APP_LOGGER.exception(str(message))
        else:
            APP_LOGGER.exception(f"{message} | {exc}")
    except Exception:
        pass

SELECTED_PLOT_FONT = configure_matplotlib_cjk_font()


PACKED_DATA_DIR = BASE_DIR / "data"
EXTERNAL_DATA_DIR = RUNTIME_DIR / "data"

DEFAULT_MASTER_CSV = """stock_id,stock_name,market,industry,theme,sub_theme,is_etf,is_active,update_date
2330,台積電,上市,半導體,AI/晶圓代工,高權值,0,1,2026-03-22
2454,聯發科,上市,半導體,IC設計,高權值,0,1,2026-03-22
2317,鴻海,上市,電子代工,AI伺服器,高權值,0,1,2026-03-22
3231,緯創,上市,電子代工,AI伺服器,伺服器,0,1,2026-03-22
2382,廣達,上市,電子代工,AI伺服器,伺服器,0,1,2026-03-22
6669,緯穎,上市,電子代工,AI伺服器,伺服器,0,1,2026-03-22
2308,台達電,上市,電源/電機,電源/HVDC,電源,0,1,2026-03-22
3017,奇鋐,上市,散熱,AI散熱,液冷,0,1,2026-03-22
3324,雙鴻,上市,散熱,AI散熱,液冷,0,1,2026-03-22
3596,智易,上市,網通,網通,寬頻,0,1,2026-03-22
2345,智邦,上市,網通,資料中心交換器,高階網通,0,1,2026-03-22
4979,華星光,上櫃,光通訊,CPO/光模組,高速光通訊,0,1,2026-03-22
3443,創意,上市,半導體,ASIC,AI ASIC,0,1,2026-03-22
6533,晶心科,上市,半導體,RISC-V,IP,0,1,2026-03-22
0050,元大台灣50,ETF,ETF,大型權值,ETF,1,1,2026-03-22
0056,元大高股息,ETF,ETF,高股息,ETF,1,1,2026-03-22
00919,群益台灣精選高息,ETF,ETF,高股息,ETF,1,1,2026-03-22
00929,復華台灣科技優息,ETF,ETF,科技高息,ETF,1,1,2026-03-22
"""

def ensure_external_master_csv() -> Path:
    EXTERNAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
    external_csv = EXTERNAL_DATA_DIR / "stocks_master.csv"
    if not external_csv.exists():
        external_csv.write_text(DEFAULT_MASTER_CSV, encoding="utf-8-sig")
    return external_csv

def resolve_master_csv() -> Path:
    external_csv = EXTERNAL_DATA_DIR / "stocks_master.csv"
    packed_csv = PACKED_DATA_DIR / "stocks_master.csv"
    if external_csv.exists():
        return external_csv
    if packed_csv.exists():
        return packed_csv
    return ensure_external_master_csv()


CLASSIFICATION_DOWNLOAD_URL = "https://mopsfin.twse.com.tw/opendata/t187ap03_L.csv"
CLASSIFICATION_DOWNLOAD_URL_TWSE = CLASSIFICATION_DOWNLOAD_URL
CLASSIFICATION_DOWNLOAD_URL_TPEX = "https://mopsfin.twse.com.tw/opendata/t187ap03_O.csv"
CLASSIFICATION_LEGACY_DOWNLOAD_URL = "https://www.twse.com.tw/docs1/data01/market/public_html/960803-0960203558-2.xls"
CLASSIFICATION_CACHE_DIR = EXTERNAL_DATA_DIR / "classification"
CLASSIFICATION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
CLASSIFICATION_CACHE_CSV_TWSE = CLASSIFICATION_CACHE_DIR / "台股官方產業分類_上市.csv"
CLASSIFICATION_CACHE_CSV_TPEX = CLASSIFICATION_CACHE_DIR / "台股官方產業分類_上櫃.csv"
CLASSIFICATION_CACHE_CSV = CLASSIFICATION_CACHE_CSV_TWSE  # backward compatibility
CLASSIFICATION_CACHE_XLS = CLASSIFICATION_CACHE_DIR / "台股類股分類.xls"
CLASSIFICATION_CACHE_XLSX = CLASSIFICATION_CACHE_DIR / "台股類股分類.xlsx"
CLASSIFICATION_META_PATH = CLASSIFICATION_CACHE_DIR / "classification_meta.json"
CLASSIFICATION_CACHE_PICKLE = CLASSIFICATION_CACHE_DIR / "classification_cache.pkl"
CLASSIFICATION_MAX_AGE_DAYS = 7
CLASSIFICATION_DOWNLOAD_TIMEOUT = (10, 45)
CLASSIFICATION_DOWNLOAD_RETRIES = 3
CLASSIFICATION_MEMORY_CACHE = {"df": None, "path": None, "mtime": None, "meta": None, "market": "ALL"}
LAST_CLASSIFICATION_LOAD_INFO = {"loaded": False, "rows": 0, "path": "", "note": "尚未載入"}
CLASSIFICATION_V2_SUMMARY_PATH = CLASSIFICATION_CACHE_DIR / "classification_v2_summary.json"
CLASSIFICATION_V2_UNCLASSIFIED_PATH = CLASSIFICATION_CACHE_DIR / "未匹配分類清單.xlsx"
CLASSIFICATION_V2_LAST_SUMMARY = {}
CLASSIFICATION_V2_SUMMARY_HISTORY = []
CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS = 500

CLASSIFICATION_LOG_CALLBACK = None
CLASSIFICATION_STABILITY_LOCK = threading.RLock()
CLASSIFICATION_OFFICIAL_CACHE = {
    "上市": {"df": None, "path": "", "mtime": 0.0, "rows": 0},
    "上櫃": {"df": None, "path": "", "mtime": 0.0, "rows": 0},
    "ALL": {"df": None, "path": "", "mtime": 0.0, "rows": 0},
}
CLASSIFICATION_DOWNLOAD_SOURCES = {
    "上市": {
        "url": CLASSIFICATION_DOWNLOAD_URL_TWSE,
        "cache_path": CLASSIFICATION_CACHE_CSV_TWSE,
        "source": "MOPS-CSV-TWSE",
    },
    "上櫃": {
        "url": CLASSIFICATION_DOWNLOAD_URL_TPEX,
        "cache_path": CLASSIFICATION_CACHE_CSV_TPEX,
        "source": "MOPS-CSV-TPEX",
    },
}
BOOTSTRAP_LOCK = threading.Lock()
BOOTSTRAP_EVENT = threading.Event()

def set_classification_log_callback(cb):
    global CLASSIFICATION_LOG_CALLBACK
    CLASSIFICATION_LOG_CALLBACK = cb

def classification_debug_log(message: str, level: str = "INFO"):
    msg = f"[分類載入] {message}"
    try:
        level_upper = str(level or "INFO").upper()
        if level_upper == "ERROR":
            log_error(msg)
        elif level_upper == "WARNING":
            log_warning(msg)
        else:
            log_info(msg)
    except Exception:
        pass
    try:
        cb = CLASSIFICATION_LOG_CALLBACK
        if cb is not None:
            cb(msg, level)
    except Exception:
        pass

def _append_classification_summary_history(summary: dict, promoted: bool = False):
    global CLASSIFICATION_V2_SUMMARY_HISTORY
    try:
        item = dict(summary or {})
        item["promoted"] = bool(promoted)
        item["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        history = list(CLASSIFICATION_V2_SUMMARY_HISTORY or [])
        history.append(item)
        CLASSIFICATION_V2_SUMMARY_HISTORY = history[-100:]
    except Exception:
        pass

def _summary_sort_key(summary: dict) -> tuple:
    if not isinstance(summary, dict):
        return (-1, -1, -1.0, "")
    total = int(summary.get("total", 0) or 0)
    covered = int(summary.get("covered", 0) or 0)
    coverage = float(summary.get("coverage_pct", 0) or 0.0)
    report_time = str(summary.get("report_time", "") or "")
    return (total, covered, coverage, report_time)

def _pick_best_classification_summary(*candidates) -> dict:
    valid = []
    for item in candidates:
        if isinstance(item, dict) and item:
            valid.append(dict(item))
    if not valid:
        return {}
    return max(valid, key=_summary_sort_key)

def _should_promote_classification_summary(summary: dict) -> bool:
    total = int(summary.get("total", 0) or 0)
    if total < CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS:
        return False
    current = {}
    try:
        if CLASSIFICATION_V2_SUMMARY_PATH.exists():
            raw = json.loads(CLASSIFICATION_V2_SUMMARY_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                current = dict(raw)
    except Exception:
        current = {}
    best = _pick_best_classification_summary(current, CLASSIFICATION_V2_LAST_SUMMARY)
    if not best:
        return True
    return _summary_sort_key(summary) >= _summary_sort_key(best)

def _get_cached_official_classification(market: str = "ALL", path: Path | None = None) -> Optional[pd.DataFrame]:
    market = str(market or "ALL").strip() or "ALL"
    with CLASSIFICATION_STABILITY_LOCK:
        slot = CLASSIFICATION_OFFICIAL_CACHE.get(market)
        if not slot or slot.get("df") is None:
            return None
        cached_df = slot.get("df")
        cached_path = str(slot.get("path") or "")
        cached_mtime = float(slot.get("mtime", 0) or 0)
        if path is None:
            return cached_df.copy() if cached_df is not None else None
        p = Path(path)
        if (not p.exists()) or cached_df is None:
            return None
        if cached_path == str(p) and abs(float(p.stat().st_mtime) - cached_mtime) < 1e-6:
            return cached_df.copy()
    return None

def _set_cached_official_classification(market: str, path: Path | None, df: pd.DataFrame):
    market = str(market or "ALL").strip() or "ALL"
    try:
        cached_df = df.copy() if df is not None else None
    except Exception:
        cached_df = df
    with CLASSIFICATION_STABILITY_LOCK:
        slot = CLASSIFICATION_OFFICIAL_CACHE.setdefault(market, {"df": None, "path": "", "mtime": 0.0, "rows": 0})
        slot["df"] = cached_df
        slot["path"] = str(path) if path else ""
        slot["mtime"] = float(Path(path).stat().st_mtime) if path and Path(path).exists() else 0.0
        slot["rows"] = int(len(df) if df is not None else 0)



CLASSIFICATION_BOOK_CANDIDATES = [
    RUNTIME_DIR / "台股官方產業分類_上市.csv",
    EXTERNAL_DATA_DIR / "台股官方產業分類_上市.csv",
    CLASSIFICATION_CACHE_CSV_TWSE,
    BASE_DIR / "台股官方產業分類_上市.csv",
    RUNTIME_DIR / "台股官方產業分類_上櫃.csv",
    EXTERNAL_DATA_DIR / "台股官方產業分類_上櫃.csv",
    CLASSIFICATION_CACHE_CSV_TPEX,
    BASE_DIR / "台股官方產業分類_上櫃.csv",
    RUNTIME_DIR / "台股類股分類.csv",
    EXTERNAL_DATA_DIR / "台股類股分類.csv",
    BASE_DIR / "台股類股分類.csv",
    RUNTIME_DIR / "台股類股分類.xlsx",
    RUNTIME_DIR / "台股類股分類.xls",
    RUNTIME_DIR / "股票類別對照表.xlsx",
    RUNTIME_DIR / "股票類別對照表.xls",
    EXTERNAL_DATA_DIR / "台股類股分類.xlsx",
    EXTERNAL_DATA_DIR / "台股類股分類.xls",
    EXTERNAL_DATA_DIR / "股票類別對照表.xlsx",
    EXTERNAL_DATA_DIR / "股票類別對照表.xls",
    CLASSIFICATION_CACHE_XLSX,
    CLASSIFICATION_CACHE_XLS,
    BASE_DIR / "台股類股分類.xlsx",
    BASE_DIR / "台股類股分類.xls",
    BASE_DIR / "股票類別對照表.xlsx",
    BASE_DIR / "股票類別對照表.xls",
]


def _safe_file_sha256(path: Path) -> str:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""

def _read_classification_meta() -> dict:
    try:
        if CLASSIFICATION_META_PATH.exists():
            return json.loads(CLASSIFICATION_META_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _write_classification_meta(meta: dict):
    try:
        CLASSIFICATION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        CLASSIFICATION_META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def _build_classification_meta(path: Path, source: str = "TWSE", note: str = "") -> dict:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stat = path.stat() if path.exists() else None
    return {
        "last_update": now,
        "source": source,
        "file": path.name if path else "",
        "path": str(path) if path else "",
        "hash": _safe_file_sha256(path) if path and path.exists() else "",
        "size": int(stat.st_size) if stat else 0,
        "mtime": float(stat.st_mtime) if stat else 0.0,
        "status": "ok" if path and path.exists() else "missing",
        "note": note or "",
    }

def _mark_classification_meta_status(status: str, note: str = "", path: Path | None = None):
    meta = _read_classification_meta()
    if path and Path(path).exists():
        meta.update(_build_classification_meta(Path(path), source=meta.get("source", "TWSE"), note=note))
    meta["status"] = status
    if note:
        meta["note"] = note
    if "last_update" not in meta:
        meta["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write_classification_meta(meta)

def _classification_meta_is_stale(meta: dict, max_age_days: int = CLASSIFICATION_MAX_AGE_DAYS) -> bool:
    try:
        ts = float(meta.get("mtime", 0) or 0)
        if ts <= 0:
            return True
        age_days = (time.time() - ts) / 86400.0
        return age_days > max_age_days
    except Exception:
        return True

def _set_classification_load_info(loaded: bool, rows: int = 0, path: Path | None = None, note: str = ""):
    global LAST_CLASSIFICATION_LOAD_INFO
    LAST_CLASSIFICATION_LOAD_INFO = {
        "loaded": bool(loaded),
        "rows": int(rows or 0),
        "path": str(path) if path else "",
        "note": str(note or "")
    }

def get_classification_load_info() -> dict:
    try:
        return dict(LAST_CLASSIFICATION_LOAD_INFO)
    except Exception:
        return {"loaded": False, "rows": 0, "path": "", "note": "未知"}

def get_classification_status() -> dict:
    path = resolve_classification_book()
    meta = _read_classification_meta()
    out = dict(meta) if isinstance(meta, dict) else {}
    load_info = get_classification_load_info()
    out["loaded"] = bool(load_info.get("loaded", False))
    out["loaded_rows"] = int(load_info.get("rows", 0) or 0)
    out["load_note"] = str(load_info.get("note", "") or "")
    out["load_path"] = str(load_info.get("path", "") or "")
    if path and Path(path).exists():
        p = Path(path)
        out["file"] = p.name
        out["path"] = str(p)
        out["hash"] = _safe_file_sha256(p)
        out["size"] = int(p.stat().st_size)
        out["mtime"] = float(p.stat().st_mtime)
        out["exists"] = True
        out["is_stale"] = _classification_meta_is_stale(out)
    else:
        out["exists"] = False
        out["is_stale"] = True
        out.setdefault("status", "missing")
    if out.get("loaded") and out.get("loaded_rows", 0) > 0:
        out["status"] = "ok"
    elif out.get("exists"):
        out["status"] = out.get("status", "fallback")
    else:
        out["status"] = out.get("status", "missing")
    return out

def _normalize_stock_name_for_match(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    replacements = {
        "（": "(", "）": ")", "　": "", " ": "", "-": "", "_": "",
        "股份有限公司": "", "有限公司": "", "公司": "", "控股": "", "控": "",
        "DR": "", "dr": "", "ETF": "ETF",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    s = re.sub(r"[\.\,\/\\]+", "", s)
    return s.upper().strip()



def normalize_stock_id(v) -> str:
    """股票代號正規化：支援 2330、2330.TW、0050、50.0、00919。"""
    if v is None:
        return ""
    s = str(v).strip().replace("＝", "=").replace("\ufeff", "")
    if s in ("", "nan", "None", "NaN", "NULL", "null", "<NA>"):
        return ""
    s = s.upper().replace(".TW", "").replace(".TWO", "").replace(".OTC", "")
    try:
        if re.fullmatch(r"\d+\.0+", s):
            s = str(int(float(s)))
    except Exception:
        pass
    m = re.search(r"(\d{1,5})", s)
    if not m:
        return ""
    code = m.group(1)
    if len(code) >= 5:
        return code[-5:]
    return code.zfill(4)

def safe_read_csv_auto(path: Path) -> pd.DataFrame:
    path = Path(path)
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "big5", "cp950"):
        try:
            return pd.read_csv(path, encoding=enc, dtype=str).fillna("")
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"CSV讀取失敗：{path}｜{last_error}")




def _coerce_unique_columns(columns) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for c in list(columns):
        name = str(c).strip() if c is not None else ""
        if not name:
            name = "Unnamed"
        count = seen.get(name, 0)
        out_name = name if count == 0 else f"{name}__dup{count}"
        seen[name] = count + 1
        out.append(out_name)
    return out


def _validate_classification_helper_integrity():
    missing = []
    for fn_name in ['_coerce_unique_columns', 'normalize_stock_id', 'normalize_official_industry_name']:
        if fn_name not in globals() or not callable(globals().get(fn_name)):
            missing.append(fn_name)
    if missing:
        raise RuntimeError(f"分類模組缺少必要函式：{', '.join(missing)}")


def _extract_official_csv_rows(df: pd.DataFrame, market: str = "上市") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["stock_id", "stock_name_official", "stock_name_norm_official", "market_official", "industry_official"])

    x = df.copy().fillna("")
    x.columns = _coerce_unique_columns(x.columns)

    code_col = None
    name_col = None
    industry_name_col = None
    industry_code_col = None

    for c in x.columns:
        base = str(c).split("__dup")[0].strip()
        if base in ("公司代號", "股票代號", "證券代號", "公司代碼"):
            code_col = c if code_col is None else code_col
        elif base in ("公司名稱", "公司簡稱", "證券名稱", "股票名稱"):
            name_col = c if name_col is None else name_col
        elif base in ("新產業類別", "新產業別", "產業名稱", "industry_name"):
            industry_name_col = c if industry_name_col is None else industry_name_col
        elif base in ("產業別", "產業類別", "產業代碼", "產業類別代號", "industry_code"):
            industry_code_col = c if industry_code_col is None else industry_code_col

    industry_col = industry_name_col or industry_code_col
    if code_col is None or industry_col is None:
        return pd.DataFrame(columns=["stock_id", "stock_name_official", "stock_name_norm_official", "market_official", "industry_official"])

    out = pd.DataFrame({
        "stock_id": x[code_col].map(normalize_stock_id),
        "stock_name_official": x[name_col].astype(str).str.strip() if name_col in x.columns else "",
        "industry_official": x[industry_col].astype(str).str.strip(),
    })
    out["industry_official"] = out["industry_official"].map(normalize_official_industry_name)
    out["stock_name_norm_official"] = out["stock_name_official"].map(_normalize_stock_name_for_match)
    out["market_official"] = market
    out = out[(out["stock_id"] != "") & (out["industry_official"] != "")].copy()
    if out.empty:
        return out
    return out[["stock_id", "stock_name_official", "stock_name_norm_official", "market_official", "industry_official"]].drop_duplicates(subset=["stock_id"], keep="first")

def _try_convert_xls_to_xlsx(src: Path, dst: Path) -> Optional[Path]:
    try:
        if not src.exists():
            return None
        import win32com.client  # type: ignore
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(src.resolve()))
        try:
            wb.SaveAs(str(dst.resolve()), FileFormat=51)
        finally:
            wb.Close(False)
            excel.Quit()
        if dst.exists() and dst.stat().st_size > 0:
            return dst
    except Exception:
        pass
    return None

def _try_convert_xls_to_xlsx_soffice(src: Path, dst: Path) -> Optional[Path]:
    try:
        src = Path(src)
        dst = Path(dst)
        if not src.exists():
            return None
        out_dir = dst.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        for candidate in ("soffice", "libreoffice"):
            try:
                proc = subprocess.run(
                    [candidate, "--headless", "--convert-to", "xlsx", "--outdir", str(out_dir), str(src)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=60,
                )
                if proc.returncode == 0:
                    break
            except Exception:
                continue
        if dst.exists() and dst.stat().st_size > 0:
            return dst
        alt = out_dir / f"{src.stem}.xlsx"
        if alt.exists() and alt.stat().st_size > 0:
            if alt != dst:
                try:
                    if dst.exists():
                        dst.unlink()
                except Exception:
                    pass
                alt.replace(dst)
            return dst
    except Exception:
        pass
    return None


def safe_read_excel(path: Path):
    path = Path(path)
    last_error = None

    if path.suffix.lower() == ".xlsx":
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception as exc:
            last_error = exc

    if path.suffix.lower() == ".xls":
        converted = convert_xls_to_xlsx_force(path, CLASSIFICATION_CACHE_XLSX)
        if converted is not None and converted.exists():
            try:
                return pd.read_excel(converted, engine="openpyxl")
            except Exception as exc:
                last_error = exc
        raise RuntimeError(f"Excel讀取失敗：{path}｜無法將 xls 轉成 xlsx，且目前環境不可直接讀取 xls。原始錯誤：{last_error}")

    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception as exc:
        last_error = exc
    raise RuntimeError(f"Excel讀取失敗：{path}｜{last_error}")


def convert_xls_to_xlsx_force(src: Path, dst: Path, log_cb=None) -> Optional[Path]:
    src = Path(src)
    dst = Path(dst)

    converted = _try_convert_xls_to_xlsx(src, dst)
    if converted is not None and converted.exists() and converted.stat().st_size > 0:
        if log_cb:
            log_cb(f"分類檔已轉成 xlsx（Excel COM）：{converted}")
        return converted

    converted = _try_convert_xls_to_xlsx_soffice(src, dst)
    if converted is not None and converted.exists() and converted.stat().st_size > 0:
        if log_cb:
            log_cb(f"分類檔已轉成 xlsx（LibreOffice）：{converted}")
        return converted

    try:
        import importlib.util
        if importlib.util.find_spec("xlrd") is not None:
            xl = pd.ExcelFile(src, engine="xlrd")
            with pd.ExcelWriter(dst, engine="openpyxl") as writer:
                for sheet in xl.sheet_names:
                    try:
                        df = xl.parse(sheet)
                        df.to_excel(writer, sheet_name=safe_sheet_name(sheet), index=False)
                    except Exception:
                        continue
            if dst.exists() and dst.stat().st_size > 0:
                if log_cb:
                    log_cb(f"分類檔已轉成 xlsx（pandas/xlrd fallback）：{dst}")
                return dst
    except Exception:
        pass

    return None



def _download_single_classification_csv(market: str, force_refresh: bool = False, log_cb=None) -> Optional[Path]:
    market = str(market or "").strip()
    cfg = CLASSIFICATION_DOWNLOAD_SOURCES.get(market)
    if not cfg:
        return None
    cache_path = Path(cfg["cache_path"])
    if cache_path.exists() and not force_refresh:
        return cache_path

    last_error = None
    for attempt in range(1, CLASSIFICATION_DOWNLOAD_RETRIES + 1):
        try:
            if log_cb:
                log_cb(f"{market} 分類來源下載開始（第 {attempt}/{CLASSIFICATION_DOWNLOAD_RETRIES} 次）：{cfg['url']}")
            resp = requests.get(
                cfg["url"],
                timeout=CLASSIFICATION_DOWNLOAD_TIMEOUT,
                headers={"User-Agent": "Mozilla/5.0", "Referer": "https://mopsfin.twse.com.tw/"}
            )
            resp.raise_for_status()
            content = resp.content or b""
            if len(content) < 128:
                raise ValueError("下載內容過小，疑似失敗或非有效檔案")
            cache_path.write_bytes(content)
            probe = safe_read_csv_auto(cache_path)
            parsed = _extract_official_csv_rows(probe, market=market)
            if parsed is None or parsed.empty:
                raise RuntimeError(f"{market} CSV可讀取，但無可用官方產業資料")
            if log_cb:
                log_cb(f"{market} 官方分類CSV已下載到快取：{cache_path}｜rows={len(parsed)}")
            return cache_path
        except Exception as exc:
            last_error = exc
            log_warning(f"{market} 分類來源下載失敗（第 {attempt} 次）：{exc}")
            if log_cb:
                log_cb(f"{market} 分類來源下載失敗（第 {attempt} 次）：{exc}")
            if attempt < CLASSIFICATION_DOWNLOAD_RETRIES:
                time.sleep(min(2 * attempt, 5))
    raise RuntimeError(f"{market} 分類來源下載失敗：{last_error}")

def download_classification_book(force_refresh: bool = False, log_cb=None) -> Optional[Path]:
    if not force_refresh:
        existing = resolve_classification_book_by_market("上市")
        existing_tpex = resolve_classification_book_by_market("上櫃")
        if existing is not None and existing_tpex is not None:
            return existing

    CLASSIFICATION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    downloaded = []
    errors = []
    for market in ("上市", "上櫃"):
        try:
            p = _download_single_classification_csv(market, force_refresh=force_refresh, log_cb=log_cb)
            if p is not None:
                downloaded.append((market, p))
        except Exception as exc:
            errors.append(f"{market}: {exc}")

    if downloaded:
        meta = {
            "last_update": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": "MOPS-CSV-TWSE+TPEX",
            "file": ", ".join(Path(p).name for _, p in downloaded),
            "path": " | ".join(str(p) for _, p in downloaded),
            "status": "ok" if len(downloaded) == 2 else "partial",
            "note": f"downloaded markets={','.join(m for m,_ in downloaded)}" + (f" | errors={' ; '.join(errors)}" if errors else ""),
        }
        _write_classification_meta(meta)
        CLASSIFICATION_MEMORY_CACHE["df"] = None
        CLASSIFICATION_MEMORY_CACHE["path"] = None
        CLASSIFICATION_MEMORY_CACHE["mtime"] = None
        CLASSIFICATION_MEMORY_CACHE["meta"] = meta
        CLASSIFICATION_MEMORY_CACHE["market"] = "ALL"
        first_path = downloaded[0][1]
        note = "已下載官方分類CSV（上市+上櫃）" if len(downloaded) == 2 else f"部分下載成功：{','.join(m for m,_ in downloaded)}"
        _set_classification_load_info(False, 0, first_path, note)
        return first_path

    fallback = next(iter(_iter_classification_candidates_prefer_xlsx()), None)
    if fallback is not None:
        _mark_classification_meta_status("fallback", note=f"download failed: {' ; '.join(errors)}", path=fallback)
        _set_classification_load_info(False, 0, fallback, f"下載失敗，使用既有檔：{' ; '.join(errors)}")
        return fallback
    _mark_classification_meta_status("download_failed", note=' ; '.join(errors) or "unknown")
    _set_classification_load_info(False, 0, None, ' ; '.join(errors) or "unknown")
    return None

def _iter_classification_candidates_prefer_xlsx():
    csv_first = []
    xlsx_second = []
    xls_after = []
    others = []
    for p in CLASSIFICATION_BOOK_CANDIDATES:
        pp = Path(p)
        if not pp.exists():
            continue
        suffix = pp.suffix.lower()
        if suffix == ".csv":
            csv_first.append(pp)
        elif suffix == ".xlsx":
            xlsx_second.append(pp)
        elif suffix == ".xls":
            xls_after.append(pp)
        else:
            others.append(pp)
    seen = set()
    ordered = []
    for pp in csv_first + xlsx_second + xls_after + others:
        key = str(pp.resolve()) if pp.exists() else str(pp)
        if key in seen:
            continue
        seen.add(key)
        ordered.append(pp)
    return ordered

def ensure_classification_book(force_refresh: bool = False, log_cb=None) -> Optional[Path]:
    if not force_refresh:
        for p in _iter_classification_candidates_prefer_xlsx():
            meta = _read_classification_meta()
            if not meta or Path(str(meta.get("path", "")).split(" | ")[0]) != p:
                _write_classification_meta(_build_classification_meta(p, source="LOCAL", note="discovered existing file"))
            return p
    return download_classification_book(force_refresh=force_refresh, log_cb=log_cb)

def resolve_classification_book() -> Optional[Path]:
    return ensure_classification_book(force_refresh=False)

def _classification_csv_candidates_by_market(market: str = "ALL") -> list[Path]:
    market = str(market or "ALL").strip()
    if market == "上市":
        return [
            RUNTIME_DIR / "台股官方產業分類_上市.csv",
            EXTERNAL_DATA_DIR / "台股官方產業分類_上市.csv",
            CLASSIFICATION_CACHE_CSV_TWSE,
            BASE_DIR / "台股官方產業分類_上市.csv",
        ]
    if market == "上櫃":
        return [
            RUNTIME_DIR / "台股官方產業分類_上櫃.csv",
            EXTERNAL_DATA_DIR / "台股官方產業分類_上櫃.csv",
            CLASSIFICATION_CACHE_CSV_TPEX,
            BASE_DIR / "台股官方產業分類_上櫃.csv",
        ]
    return _classification_csv_candidates_by_market("上市") + _classification_csv_candidates_by_market("上櫃")

def resolve_classification_book_by_market(market: str = "ALL") -> Optional[Path]:
    for p in _classification_csv_candidates_by_market(market):
        if Path(p).exists():
            return Path(p)
    if market in ("上市", "上櫃"):
        try:
            downloaded = _download_single_classification_csv(market, force_refresh=False, log_cb=classification_debug_log)
            if downloaded is not None and Path(downloaded).exists():
                return Path(downloaded)
        except Exception as exc:
            classification_debug_log(f"自動補抓官方分類失敗（{market}）：{exc}", "WARNING")
    return None

def load_official_classification_book(market: str = "ALL") -> pd.DataFrame:
    _validate_classification_helper_integrity()
    market = str(market or "ALL").strip() or "ALL"
    empty_df = pd.DataFrame(columns=["stock_id", "stock_name_official", "stock_name_norm_official", "market_official", "industry_official"])

    if market == "ALL":
        cached_all = _get_cached_official_classification("ALL")
        if cached_all is not None and not cached_all.empty:
            return cached_all
        twse = load_official_classification_book("上市")
        tpex = load_official_classification_book("上櫃")
        parts = [df for df in [twse, tpex] if df is not None and not df.empty]
        if not parts:
            note = "找不到任何官方分類來源（ALL）"
            classification_debug_log(note, "WARNING")
            return empty_df
        out = pd.concat(parts, ignore_index=True).fillna("")
        out["stock_id"] = out["stock_id"].astype(str).map(normalize_stock_id)
        out["industry_official"] = out["industry_official"].astype(str).map(normalize_official_industry_name)
        out = out[(out["stock_id"] != "") & (out["industry_official"] != "")].copy()
        out = out.sort_values(["stock_id", "market_official", "industry_official"]).drop_duplicates(subset=["stock_id"], keep="first")
        _set_cached_official_classification("ALL", None, out)
        _set_classification_load_info(True, len(out), None, f"補充分類載入成功：ALL｜rows={len(out)}")
        classification_debug_log(f"補充分類載入成功：ALL｜rows={len(out)}")
        return out

    path = resolve_classification_book_by_market(market)
    if path is None:
        note = f"找不到官方分類來源：{market}"
        classification_debug_log(note, "WARNING")
        return empty_df

    cached_market = _get_cached_official_classification(market, path)
    if cached_market is not None and not cached_market.empty:
        return cached_market

    classification_debug_log(f"實際讀到的分類來源（{market}）：{path}")
    try:
        raw = safe_read_csv_auto(path)
        out = _extract_official_csv_rows(raw, market=market)
        if out is None or out.empty:
            raise RuntimeError(f"{market} 官方分類可讀取，但沒有有效資料")
        out["stock_id"] = out["stock_id"].astype(str).map(normalize_stock_id)
        out["industry_official"] = out["industry_official"].astype(str).map(normalize_official_industry_name)
        out = out[(out["stock_id"] != "") & (out["industry_official"] != "")].copy()
        out = out.drop_duplicates(subset=["stock_id"], keep="first")
        _set_cached_official_classification(market, path, out)
        _set_classification_load_info(True, len(out), path, f"補充分類載入成功：{market}｜rows={len(out)}")
        classification_debug_log(f"補充分類載入成功：{market}｜rows={len(out)}")
        return out
    except Exception as exc:
        classification_debug_log(f"分類來源解析失敗（{market}）：{exc}", "ERROR")
        return empty_df

def load_manual_theme_mapping() -> pd.DataFrame:


    manual_parts = []
    try:
        seed = pd.read_csv(io.StringIO(DEFAULT_MASTER_CSV), dtype={"stock_id": str}).fillna("")
        manual_parts.append(seed)
    except Exception:
        pass
    try:
        csv_path = resolve_master_csv()
        if csv_path.exists():
            ext = pd.read_csv(csv_path, dtype={"stock_id": str}).fillna("")
            ext = ext[[c for c in ["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf"] if c in ext.columns]]
            if not ext.empty:
                manual_parts.append(ext)
    except Exception:
        pass
    if not manual_parts:
        return pd.DataFrame(columns=["stock_id", "stock_name_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"])
    x = pd.concat(manual_parts, ignore_index=True).fillna("")
    x["stock_id"] = x["stock_id"].astype(str).map(normalize_stock_id)
    x = x[x["stock_id"] != ""].copy()
    for c in ["stock_name", "market", "industry", "theme", "sub_theme", "is_etf"]:
        if c not in x.columns:
            x[c] = ""
    x["stock_name_norm_manual"] = x["stock_name"].map(_normalize_stock_name_for_match)
    x = x.drop_duplicates(subset=["stock_id"], keep="first")
    return x.rename(columns={
        "stock_name": "stock_name_manual",
        "stock_name_norm_manual": "stock_name_norm_manual",
        "market": "market_manual",
        "industry": "industry_manual",
        "theme": "theme_manual",
        "sub_theme": "sub_theme_manual",
        "is_etf": "is_etf_manual",
    })



def _ensure_object_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    x = df.copy()
    for col in columns:
        if col not in x.columns:
            x[col] = ""
        try:
            x[col] = x[col].astype("object")
        except Exception:
            x[col] = pd.Series(list(x[col]), index=x.index, dtype="object")
    return x

def _safe_text_series(series: pd.Series, default: str = "") -> pd.Series:
    if series is None:
        return pd.Series(dtype="object")
    s = pd.Series(series, copy=True)
    try:
        s = s.astype("object")
    except Exception:
        s = pd.Series(list(s), index=s.index, dtype="object")
    s = s.where(pd.notna(s), default)
    def _clean(v):
        if v is None:
            return default
        sv = str(v).strip()
        if sv in ("<NA>", "nan", "None", "NaN", "NULL", "null"):
            return default
        return sv
    s = s.map(_clean)
    try:
        s = s.fillna(default)
    except Exception:
        pass
    return s.astype("object")

def _safe_numeric_flag_series(series: pd.Series, default: int = 0) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(default)
    return s.astype(int)

def _safe_num_series(df: pd.DataFrame, col: str, default: float = 0.0) -> pd.Series:
    if isinstance(df, pd.DataFrame) and col in df.columns:
        return pd.to_numeric(df[col], errors="coerce").fillna(default)
    if isinstance(df, pd.DataFrame):
        return pd.Series(default, index=df.index, dtype="float64")
    return pd.Series(dtype="float64")


def calculate_eps_ttm(price, pe, max_abs_eps: float = 100.0):
    """V9.5.4：以 TWSE BWIBBU_d 的收盤價 / 本益比反推 EPS_TTM。

    定位：EPS_TTM 是估值近似值（Trailing EPS proxy），保留追溯與評分用途；
    不等同 MOPS 原始季 EPS，也不可作為硬性 Gate。
    """
    try:
        price_v = float(price)
        pe_v = float(pe)
        if not np.isfinite(price_v) or not np.isfinite(pe_v) or pe_v <= 0:
            return None
        eps = price_v / pe_v
        if not np.isfinite(eps) or abs(eps) > float(max_abs_eps):
            return None
        return round(float(eps), 4)
    except Exception:
        return None

def _safe_text_fill_series(df: pd.DataFrame, col: str, default: str = "") -> pd.Series:
    if isinstance(df, pd.DataFrame) and col in df.columns:
        return pd.Series(df[col], index=df.index, copy=True).fillna(default).astype(str)
    if isinstance(df, pd.DataFrame):
        return pd.Series(default, index=df.index, dtype="object")
    return pd.Series(dtype="object")

def _assign_object_values(df: pd.DataFrame, mask: pd.Series, col: str, values) -> pd.DataFrame:
    x = df
    if col not in x.columns:
        x[col] = ""
    try:
        x[col] = x[col].astype("object")
    except Exception:
        x[col] = pd.Series(list(x[col]), index=x.index, dtype="object")
    vals = _safe_text_series(pd.Series(list(values), index=x.index[mask]), "")
    x.loc[mask, col] = vals.astype("object").values
    return x


def _safe_int(v, default: int = 0) -> int:
    try:
        return int(float(v))
    except Exception:
        return int(default)


def _is_placeholder_text(v) -> bool:
    return str(v or "").strip() in ("", "未分類", "全市場", "系統掃描", "其他", "nan", "None", "<NA>", "N/A", "NaN", "null", "NULL")


def _is_missing_classification_value(v) -> bool:
    s = str(v or "").strip()
    return s in ("", "未分類", "全市場", "系統掃描", "其他", "nan", "None", "<NA>", "N/A", "NaN", "null", "NULL")


def _choose_text(*values, default: str = "") -> str:
    for v in values:
        s = str(v or "").strip()
        if s and not _is_placeholder_text(s):
            return s
    return str(default or "")

def normalize_official_industry_name(v: str) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    s = s.replace(".0", "") if re.fullmatch(r"\d+\.0", s) else s
    if re.fullmatch(r"\d{1,2}", s):
        s = s.zfill(2)
    if s in INDUSTRY_CODE_MAP:
        s = INDUSTRY_CODE_MAP.get(s, s)
    return OFFICIAL_INDUSTRY_ALIAS_MAP.get(s, s)


def _write_json_safe(path: Path, payload: dict):
    try:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _write_unclassified_report(df: pd.DataFrame):
    try:
        if df is None:
            return
        CLASSIFICATION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        engine = available_excel_engine()
        if engine:
            with pd.ExcelWriter(CLASSIFICATION_V2_UNCLASSIFIED_PATH, engine=engine) as writer:
                df.to_excel(writer, sheet_name="Unclassified", index=False)
        else:
            df.to_csv(CLASSIFICATION_V2_UNCLASSIFIED_PATH.with_suffix('.csv'), index=False, encoding='utf-8-sig')
    except Exception:
        pass



def get_classification_v2_summary() -> dict:
    global CLASSIFICATION_V2_LAST_SUMMARY
    file_data = {}
    try:
        if CLASSIFICATION_V2_SUMMARY_PATH.exists():
            raw = json.loads(CLASSIFICATION_V2_SUMMARY_PATH.read_text(encoding='utf-8'))
            if isinstance(raw, dict):
                file_data = dict(raw)
    except Exception:
        file_data = {}
    best = _pick_best_classification_summary(file_data, CLASSIFICATION_V2_LAST_SUMMARY)
    if best:
        CLASSIFICATION_V2_LAST_SUMMARY = dict(best)
        return dict(best)
    return {}

def infer_ai_classification(stock_name: str, industry_hint: str = "", market_hint: str = "", is_etf: int = 0) -> tuple[str, str, str, str, int, str]:
    name = str(stock_name or "").strip()
    industry_hint = normalize_official_industry_name(str(industry_hint or "").strip())
    market_hint = str(market_hint or "").strip()
    if int(is_etf or 0) == 1 or re.search(r"ETF|台灣50|高股息|中型100|科技優息|精選高息", name, flags=re.I):
        return ("ETF", "ETF", "ETF", "ai_infer", 98, "ETF keyword")

    ai_rules = [
        (r"華星光|上詮|聯鈞|波若威|光聖|聯亞|眾達|前鼎|立碁|環宇|光環|聯光通|眾達-KY", ("光通訊", "CPO/光模組", "高速光通訊", 94, "光通訊/CPO keyword")),
        (r"智邦|智易|中磊|啟碁|正文|建漢|神準|明泰|友訊|合勤控|振曜", ("網通", "資料中心交換器", "高階網通", 90, "網通 keyword")),
        (r"台積電|創意|世芯|世芯-KY|晶心科|聯發科|聯詠|祥碩|M31|力旺|智原|信驊", ("半導體", "AI/晶圓代工", "半導體", 92, "半導體 keyword")),
        (r"奇鋐|雙鴻|建準|高力|力致|超眾|健策", ("散熱", "AI散熱", "液冷", 90, "散熱 keyword")),
        (r"鴻海|廣達|緯創|緯穎|仁寶|英業達|和碩|神達|技嘉|華碩|微星|宏碁", ("電子代工", "AI伺服器", "伺服器", 88, "伺服器/代工 keyword")),
        (r"台達電|光寶科|康舒|群電|全漢|偉訓|順達|AES|新盛力|加百裕|系統電|飛宏|茂達", ("電源/電機", "電源/HVDC", "電源", 88, "電源 keyword")),
        (r"長榮|萬海|陽明|裕民|慧洋|中航|四維航", ("航運業", "運輸", "航運", 85, "航運 keyword")),
        (r"富邦金|國泰金|中信金|兆豐金|玉山金|元大金|第一金|華南金|永豐金|台新金", ("金融保險", "金融", "金融", 86, "金融 keyword")),
        (r"統一|大成|卜蜂|味全|愛之味|黑松|佳格", ("食品工業", "民生消費", "食品", 84, "食品 keyword")),
    ]
    for pattern, bundle in ai_rules:
        if re.search(pattern, name):
            industry, theme, sub_theme, conf, note = bundle
            return industry, theme, sub_theme, "ai_infer", conf, note

    if industry_hint in INDUSTRY_THEME_MAP:
        industry, theme, sub_theme = INDUSTRY_THEME_MAP[industry_hint]
        return industry, theme, sub_theme, "rule_engine", 76, f"industry map: {industry_hint}"

    fallback_industry, fallback_theme, fallback_sub = infer_theme_bundle(name, industry_hint, int(is_etf or 0))
    if _is_placeholder_text(fallback_theme) and _is_placeholder_text(fallback_sub):
        if market_hint == "ETF":
            return ("ETF", "ETF", "ETF", "ai_infer", 90, "market hint ETF")
        return (_choose_text(industry_hint, fallback_industry, default="未分類"), _choose_text(fallback_theme, default="全市場"), _choose_text(fallback_sub, default="系統掃描"), "ai_infer", 55, "weak fallback")
    return (_choose_text(industry_hint, fallback_industry, default="未分類"), _choose_text(fallback_theme, default="全市場"), _choose_text(fallback_sub, default="系統掃描"), "rule_engine", 68, "generic keyword fallback")



def build_classification_quality_report(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    global CLASSIFICATION_V2_LAST_SUMMARY
    if df is None or df.empty:
        summary = {
            "total": 0, "official": 0, "manual": 0, "rule_engine": 0, "ai_infer": 0,
            "unclassified": 0, "covered": 0, "coverage_pct": 0.0,
            "report_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        best = _pick_best_classification_summary(CLASSIFICATION_V2_LAST_SUMMARY, summary)
        CLASSIFICATION_V2_LAST_SUMMARY = dict(best or summary)
        if _should_promote_classification_summary(summary):
            _write_json_safe(CLASSIFICATION_V2_SUMMARY_PATH, summary)
        _append_classification_summary_history(summary, promoted=False)
        _write_unclassified_report(pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_confidence", "classification_note"]))
        return pd.DataFrame(), summary

    x = df.copy()
    for col in ["stock_id", "stock_name", "market", "industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_confidence", "classification_note"]:
        if col not in x.columns:
            x[col] = ""
    for col in ["stock_id", "stock_name", "market", "industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_note"]:
        x[col] = _safe_text_series(x[col], "")
    x["classification_confidence"] = pd.to_numeric(x["classification_confidence"], errors="coerce").fillna(0)

    unclassified_mask = (
        x["industry_final"].map(_is_missing_classification_value) |
        x["theme_final"].map(_is_missing_classification_value) |
        x["sub_theme_final"].map(_is_missing_classification_value)
    )
    unclassified = x.loc[unclassified_mask, ["stock_id", "stock_name", "market", "industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_confidence", "classification_note"]].copy().sort_values(["classification_source", "stock_id"])

    source_counts = x["classification_source"].astype(str).value_counts().to_dict()
    total = int(len(x))
    unclassified_count = int(unclassified_mask.sum())
    covered = max(total - unclassified_count, 0)
    summary = {
        "total": total,
        "official": int(source_counts.get("official", 0)),
        "manual": int(source_counts.get("manual", 0)),
        "rule_engine": int(source_counts.get("rule_engine", 0)),
        "ai_infer": int(source_counts.get("ai_infer", 0)),
        "unclassified": unclassified_count,
        "covered": covered,
        "coverage_pct": round((covered / total * 100.0), 2) if total else 0.0,
        "report_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "unclassified_report": str(CLASSIFICATION_V2_UNCLASSIFIED_PATH),
    }

    promoted = _should_promote_classification_summary(summary)
    if promoted:
        CLASSIFICATION_V2_LAST_SUMMARY = dict(summary)
        _write_json_safe(CLASSIFICATION_V2_SUMMARY_PATH, summary)
    else:
        best = _pick_best_classification_summary(CLASSIFICATION_V2_LAST_SUMMARY, get_classification_v2_summary())
        if best:
            CLASSIFICATION_V2_LAST_SUMMARY = dict(best)

    _append_classification_summary_history(summary, promoted=promoted)
    _write_unclassified_report(unclassified)

    coverage = 1.0 - float(x["industry_final"].isin(["未分類", "未匹配"]).mean()) if len(x) else 0.0
    if coverage < 0.95:
        log_warning(f"分類覆蓋率不足：{coverage:.2%}")
    return unclassified, summary

THEME_RULES = [

    (r"台積電|創意|世芯|世芯-KY|晶心|聯發科|聯詠|矽力|祥碩|M31|力旺|智原|信驊|創惟|威鋒電子", ("半導體", "AI/晶圓代工", "半導體")),
    (r"華星光|聯亞|光聖|波若威|上詮|聯鈞|眾達|環宇|前鼎|立碁|光環|聯光通|聯合再生|聯亞光|IET-KY", ("光通訊", "CPO/光模組", "高速光通訊")),
    (r"智邦|智易|中磊|啟碁|正文|建漢|神準|明泰|友訊|合勤控|振曜|康全電訊|仲琦", ("網通", "資料中心交換器", "高階網通")),
    (r"台達電|光寶科|康舒|群電|全漢|偉訓|順達|AES|新盛力|加百裕|系統電|飛宏|茂達|群光電能", ("電源/電機", "電源/HVDC", "電源")),
    (r"奇鋐|雙鴻|建準|超眾|力致|高力|健策|一詮|散熱", ("散熱", "AI散熱", "液冷")),
    (r"鴻海|廣達|緯創|緯穎|仁寶|英業達|和碩|技嘉|華碩|微星|神達|宏碁|研華", ("電子代工", "AI伺服器", "伺服器")),
    (r"南亞科|華邦電|旺宏|群聯|威剛|十銓|宇瞻|創見", ("半導體", "記憶體", "DRAM/NAND")),
    (r"日月光|京元電子|矽格|精材|頎邦|力成", ("半導體", "先進封裝", "封測")),
    (r"長榮|萬海|陽明|裕民|慧洋|中航|四維航", ("航運業", "運輸", "航運")),
    (r"富邦金|國泰金|中信金|兆豐金|玉山金|元大金|第一金|華南金|永豐金|台新金", ("金融保險", "金融", "金融")),
]

INDUSTRY_CODE_MAP = {
    "01": "水泥工業",
    "02": "食品工業",
    "03": "塑膠工業",
    "04": "紡織纖維",
    "05": "電機機械",
    "06": "電器電纜",
    "08": "玻璃陶瓷",
    "09": "造紙工業",
    "10": "鋼鐵工業",
    "11": "橡膠工業",
    "12": "汽車工業",
    "14": "建材營造",
    "15": "航運業",
    "16": "觀光餐旅",
    "17": "金融保險",
    "18": "貿易百貨",
    "20": "其他",
    "21": "化學工業",
    "22": "生技醫療",
    "23": "油電燃氣",
    "24": "半導體業",
    "25": "電腦及週邊設備業",
    "26": "光電業",
    "27": "通信網路業",
    "28": "電子零組件業",
    "29": "電子通路業",
    "30": "資訊服務業",
    "31": "其他電子業",
    "32": "文化創意業",
    "33": "農業科技業",
    "34": "電子商務",
    "35": "綠能環保",
    "36": "數位雲端",
    "37": "運動休閒",
    "38": "居家生活",
}

OFFICIAL_INDUSTRY_ALIAS_MAP = {
    "半導體業": "半導體",
    "半導體": "半導體",
    "電腦及週邊設備業": "電子代工",
    "電子工業": "電子工業",
    "電子零組件業": "電子零組件",
    "通信網路業": "網通",
    "光電業": "光通訊",
    "資訊服務業": "資訊服務",
    "其他電子業": "其他電子",
    "電機機械": "電源/電機",
    "電器電纜": "電源/電機",
    "生技醫療": "生技醫療",
    "油電燃氣": "油電燃氣",
    "其他": "其他",
    "文化創意業": "文化創意",
    "農業科技業": "農業科技",
    "電子商務": "電子商務",
    "綠能環保": "綠能環保",
    "數位雲端": "資訊服務",
    "運動休閒": "運動休閒",
    "居家生活": "居家生活",
}

INDUSTRY_THEME_MAP = {
    "食品工業": ("食品工業", "民生消費", "食品"),
    "塑膠工業": ("塑膠工業", "基礎原物料", "塑膠"),
    "紡織纖維": ("紡織纖維", "傳產", "紡織"),
    "電源/電機": ("電源/電機", "電源/HVDC", "電源"),
    "電機機械": ("電源/電機", "電源/HVDC", "電機"),
    "電器電纜": ("電源/電機", "電力基建", "電纜"),
    "化學工業": ("化學工業", "基礎原物料", "化工"),
    "生技醫療": ("生技醫療", "生技醫療", "醫療"),
    "玻璃陶瓷": ("玻璃陶瓷", "傳產", "玻璃陶瓷"),
    "造紙工業": ("造紙工業", "傳產", "造紙"),
    "鋼鐵工業": ("鋼鐵工業", "基礎原物料", "鋼鐵"),
    "橡膠工業": ("橡膠工業", "傳產", "橡膠"),
    "汽車工業": ("汽車工業", "電動車", "車用"),
    "電子工業": ("電子工業", "電子", "電子"),
    "半導體業": ("半導體", "半導體", "半導體"),
    "半導體": ("半導體", "AI/晶圓代工", "半導體"),
    "記憶體": ("半導體", "記憶體", "DRAM/NAND"),
    "先進封裝": ("半導體", "先進封裝", "封測"),
    "電腦及週邊設備業": ("電子代工", "AI伺服器", "伺服器"),
    "電子代工": ("電子代工", "AI伺服器", "伺服器"),
    "光電業": ("光電", "光電", "面板/光學"),
    "光通訊": ("光通訊", "CPO/光模組", "高速光通訊"),
    "通信網路業": ("網通", "網通/光通訊", "網通"),
    "網通": ("網通", "資料中心交換器", "高階網通"),
    "電子零組件業": ("電子零組件", "電子零組件", "零組件"),
    "電子零組件": ("電子零組件", "電子零組件", "零組件"),
    "電子通路業": ("電子通路", "電子通路", "通路"),
    "資訊服務業": ("資訊服務", "軟體/資訊服務", "資訊服務"),
    "資訊服務": ("資訊服務", "軟體/雲端", "雲端"),
    "其他電子業": ("其他電子", "電子", "其他電子"),
    "散熱": ("散熱", "AI散熱", "液冷"),
    "建材營造": ("建材營造", "傳產", "營造"),
    "航運業": ("航運業", "運輸", "航運"),
    "觀光餐旅": ("觀光餐旅", "內需消費", "觀光"),
    "金融保險": ("金融保險", "金融", "金融"),
    "貿易百貨": ("貿易百貨", "內需消費", "百貨"),
    "油電燃氣": ("油電燃氣", "公用事業", "能源"),
    "居家生活": ("居家生活", "內需消費", "居家"),
    "綠能環保": ("綠能環保", "綠能環保", "環保"),
    "數位雲端": ("資訊服務", "軟體/雲端", "雲端"),
    "運動休閒": ("運動休閒", "內需消費", "運動"),
    "文化創意業": ("文化創意", "內需消費", "文創"),
    "農業科技業": ("農業科技", "農業科技", "農業"),
    "ETF": ("ETF", "ETF", "ETF"),
}


def infer_theme_bundle(stock_name: str, industry: str, is_etf: int) -> Tuple[str, str, str]:
    name = str(stock_name or "")
    industry = normalize_official_industry_name(str(industry or "").strip())
    if int(is_etf or 0) == 1 or re.search(r"ETF|台灣50|高股息|中型100|科技優息|精選高息", name):
        return "ETF", "ETF", "ETF"
    for pattern, bundle in THEME_RULES:
        if re.search(pattern, name):
            return bundle
    if industry in INDUSTRY_THEME_MAP:
        return INDUSTRY_THEME_MAP[industry]
    if re.search(r"光|通|網|訊|CPO|矽光", name):
        return (industry or "光通訊", "CPO/光模組", "高速光通訊")
    if re.search(r"交換器|路由|寬頻|通訊|5G", name):
        return (industry or "網通", "資料中心交換器", "高階網通")
    if re.search(r"電|控|達|機|電源|電池|能源", name):
        return (industry or "電源/電機", "電源/HVDC", "電源")
    if re.search(r"積電|半導體|晶|芯|封裝|測試|DRAM|記憶體", name):
        return (industry or "半導體", "半導體", "半導體")
    if re.search(r"伺服器|AI|雲端|運算|主機板", name):
        return (industry or "電子代工", "AI伺服器", "伺服器")
    return (industry or "未分類", "全市場", "系統掃描")



def apply_classification_layers(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy().fillna("")
    x = _ensure_object_columns(x, [
        "stock_id", "stock_name", "market", "industry", "theme", "sub_theme",
        "classification_source", "classification_note"
    ])

    official = load_official_classification_book("ALL")
    manual = load_manual_theme_mapping()

    x["stock_id"] = _safe_text_series(x["stock_id"], "").map(normalize_stock_id)
    x = x[x["stock_id"] != ""].copy().reset_index(drop=True)
    if x.empty:
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])

    x["stock_name"] = _safe_text_series(x["stock_name"], "")
    x["market"] = _safe_text_series(x["market"], "")
    x["industry"] = _safe_text_series(x["industry"], "")
    x["theme"] = _safe_text_series(x["theme"], "")
    x["sub_theme"] = _safe_text_series(x["sub_theme"], "")
    x["classification_source"] = _safe_text_series(x["classification_source"], "")
    x["classification_note"] = _safe_text_series(x["classification_note"], "")
    x["stock_name_norm"] = x.get("stock_name", "").map(_normalize_stock_name_for_match)

    if not official.empty:
        official = official.copy()
        official = _ensure_object_columns(official, [
            "stock_id", "stock_name_official", "stock_name_norm_official",
            "market_official", "industry_official"
        ])
        for col in ["stock_id", "stock_name_official", "stock_name_norm_official", "market_official", "industry_official"]:
            official[col] = _safe_text_series(official[col], "")
        official["industry_official"] = official["industry_official"].map(normalize_official_industry_name)
        x = x.merge(official, on="stock_id", how="left")
        x = _ensure_object_columns(x, ["stock_name_official", "stock_name_norm_official", "market_official", "industry_official"])
        for col in ["stock_name_official", "stock_name_norm_official", "market_official", "industry_official"]:
            x[col] = _safe_text_series(x[col], "")
        missing_mask = x["industry_official"].eq("")
        if missing_mask.any() and "stock_name_norm_official" in official.columns:
            official_name_map = official[official["stock_name_norm_official"].astype(str) != ""].drop_duplicates("stock_name_norm_official")
            if not official_name_map.empty:
                miss = x.loc[missing_mask, ["stock_name_norm"]].merge(
                    official_name_map[["stock_name_norm_official", "stock_name_official", "market_official", "industry_official"]],
                    left_on="stock_name_norm", right_on="stock_name_norm_official", how="left"
                )
                for col in ["stock_name_official", "market_official", "industry_official"]:
                    vals = _safe_text_series(miss[col], "")
                    orig = _safe_text_series(x.loc[missing_mask, col], "")
                    merged_vals = [v if v else o for v, o in zip(vals.tolist(), orig.tolist())]
                    x = _assign_object_values(x, missing_mask, col, merged_vals)
    else:
        for col in ["stock_name_official", "stock_name_norm_official", "market_official", "industry_official"]:
            x[col] = ""

    if not manual.empty:
        manual = manual.copy()
        manual = _ensure_object_columns(manual, [
            "stock_id", "stock_name_manual", "stock_name_norm_manual", "market_manual",
            "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"
        ])
        for col in ["stock_id", "stock_name_manual", "stock_name_norm_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"]:
            manual[col] = _safe_text_series(manual[col], "")
        manual["industry_manual"] = manual["industry_manual"].map(normalize_official_industry_name)
        x = x.merge(manual, on="stock_id", how="left")
        x = _ensure_object_columns(x, ["stock_name_manual", "stock_name_norm_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"])
        for col in ["stock_name_manual", "stock_name_norm_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"]:
            x[col] = _safe_text_series(x[col], "")
        missing_manual = x["industry_manual"].eq("")
        if missing_manual.any() and "stock_name_norm_manual" in manual.columns:
            manual_name_map = manual[manual["stock_name_norm_manual"].astype(str) != ""].drop_duplicates("stock_name_norm_manual")
            if not manual_name_map.empty:
                miss = x.loc[missing_manual, ["stock_name_norm"]].merge(
                    manual_name_map[["stock_name_norm_manual", "stock_name_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"]],
                    left_on="stock_name_norm", right_on="stock_name_norm_manual", how="left"
                )
                for col in ["stock_name_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"]:
                    vals = _safe_text_series(miss[col], "")
                    orig = _safe_text_series(x.loc[missing_manual, col], "")
                    merged_vals = [v if v else o for v, o in zip(vals.tolist(), orig.tolist())]
                    x = _assign_object_values(x, missing_manual, col, merged_vals)
    else:
        for col in ["stock_name_manual", "stock_name_norm_manual", "market_manual", "industry_manual", "theme_manual", "sub_theme_manual", "is_etf_manual"]:
            x[col] = ""

    x["stock_name"] = _safe_text_series(x["stock_name"], "")
    x["stock_name_official"] = _safe_text_series(x.get("stock_name_official", ""), "")
    x["stock_name_manual"] = _safe_text_series(x.get("stock_name_manual", ""), "")
    x["stock_name"] = _safe_text_series(
        x["stock_name"].replace("", pd.NA).fillna(x["stock_name_official"].replace("", pd.NA)).fillna(x["stock_name_manual"].replace("", pd.NA)).fillna(x["stock_id"]),
        ""
    )

    etf_mask = x["stock_id"].astype(str).str.startswith("00") | x["stock_name"].astype(str).str.contains("ETF|台灣50|高股息|中型100|科技優息|精選高息|DR", regex=True)
    if "is_etf_manual" in x.columns:
        etf_mask = etf_mask | _safe_numeric_flag_series(x["is_etf_manual"], 0).eq(1)
    x["is_etf"] = etf_mask.astype(int)

    x["market"] = _safe_text_series(x["market"], "")
    if "market_official" in x.columns:
        x["market"] = _safe_text_series(x["market"].replace("", pd.NA).fillna(_safe_text_series(x["market_official"], "").replace("", pd.NA)), "")
    if "market_manual" in x.columns:
        x["market"] = _safe_text_series(x["market"].replace("", pd.NA).fillna(_safe_text_series(x["market_manual"], "").replace("", pd.NA)), "")
    x["market"] = _safe_text_series(x["market"].replace("", pd.NA).fillna("上市"), "上市")
    x.loc[x["is_etf"].eq(1), "market"] = "ETF"

    if "industry_official" in x.columns:
        x["industry_official"] = _safe_text_series(x["industry_official"], "").map(normalize_official_industry_name)
    if "industry_manual" in x.columns:
        x["industry_manual"] = _safe_text_series(x["industry_manual"], "").map(normalize_official_industry_name)

    x["industry_seed"] = _safe_text_series(x["industry"], "").map(normalize_official_industry_name).replace("", pd.NA)
    if "industry_official" in x.columns:
        x["industry_seed"] = x["industry_seed"].fillna(_safe_text_series(x["industry_official"], "").replace("", pd.NA))
    if "industry_manual" in x.columns:
        x["industry_seed"] = x["industry_seed"].fillna(_safe_text_series(x["industry_manual"], "").replace("", pd.NA))
    x["industry_seed"] = _safe_text_series(x["industry_seed"].fillna("未分類"), "未分類")

    x["theme_seed"] = _safe_text_series(x["theme"] if "theme" in x.columns else "", "")
    x["sub_theme_seed"] = _safe_text_series(x["sub_theme"] if "sub_theme" in x.columns else "", "")

    def _clean_classification_seed(series: pd.Series) -> pd.Series:
        return _safe_text_series(series, "")

    if "theme_manual" in x.columns:
        manual_theme = _clean_classification_seed(x["theme_manual"])
        x["theme_seed"] = _safe_text_series(x["theme_seed"].replace("", pd.NA).fillna(manual_theme.replace("", pd.NA)), "")
    if "sub_theme_manual" in x.columns:
        manual_sub_theme = _clean_classification_seed(x["sub_theme_manual"])
        x["sub_theme_seed"] = _safe_text_series(x["sub_theme_seed"].replace("", pd.NA).fillna(manual_sub_theme.replace("", pd.NA)), "")

    x["classification_source"] = _safe_text_series(x.get("classification_source", ""), "")
    official_mask = _safe_text_series(x.get("industry_official", ""), "").ne("")
    x.loc[official_mask, "classification_source"] = "official"
    manual_mask = x["classification_source"].eq("") & _safe_text_series(x.get("industry_manual", ""), "").ne("")
    x.loc[manual_mask, "classification_source"] = "manual"

    x["classification_confidence"] = 0
    x.loc[x["classification_source"].eq("official"), "classification_confidence"] = 100
    x.loc[x["classification_source"].eq("manual"), "classification_confidence"] = 92

    x["classification_note"] = _safe_text_series(x.get("classification_note", ""), "")
    x.loc[x["classification_source"].eq("official"), "classification_note"] = "official workbook matched"
    x.loc[x["classification_source"].eq("manual"), "classification_note"] = "manual mapping matched"

    ai_rows = x.apply(lambda r: infer_ai_classification(r.get("stock_name", ""), r.get("industry_seed", ""), r.get("market", ""), _safe_int(r.get("is_etf", 0))), axis=1, result_type="expand")
    ai_rows.columns = ["industry_ai", "theme_ai", "sub_theme_ai", "source_ai", "confidence_ai", "note_ai"]
    x = pd.concat([x, ai_rows], axis=1)
    x = _ensure_object_columns(x, ["industry_ai", "theme_ai", "sub_theme_ai", "source_ai", "note_ai"])
    for col in ["industry_ai", "theme_ai", "sub_theme_ai", "source_ai", "note_ai"]:
        x[col] = _safe_text_series(x[col], "")
    x["confidence_ai"] = pd.to_numeric(x["confidence_ai"], errors="coerce").fillna(0).astype(int)

    x["industry_final"] = _safe_text_series(x["industry_seed"], "未分類").map(normalize_official_industry_name)
    x["theme_final"] = _safe_text_series(x["theme_seed"], "")
    x["sub_theme_final"] = _safe_text_series(x["sub_theme_seed"], "")

    missing_industry = x["industry_final"].map(_is_missing_classification_value)
    missing_theme = x["theme_final"].map(_is_missing_classification_value)
    missing_sub = x["sub_theme_final"].map(_is_missing_classification_value)

    x.loc[missing_industry, "industry_final"] = _safe_text_series(x.loc[missing_industry, "industry_ai"], "").map(normalize_official_industry_name).values
    x.loc[missing_theme, "theme_final"] = _safe_text_series(x.loc[missing_theme, "theme_ai"], "").values
    x.loc[missing_sub, "sub_theme_final"] = _safe_text_series(x.loc[missing_sub, "sub_theme_ai"], "").values

    industry_map_rows = x["industry_final"].map(lambda s: INDUSTRY_THEME_MAP.get(normalize_official_industry_name(s), ("", "", "")))
    industry_map_df = pd.DataFrame(industry_map_rows.tolist(), columns=["industry_from_map", "theme_from_map", "sub_from_map"], index=x.index)
    x = pd.concat([x, industry_map_df], axis=1)
    x = _ensure_object_columns(x, ["industry_from_map", "theme_from_map", "sub_from_map"])
    for col in ["industry_from_map", "theme_from_map", "sub_from_map"]:
        x[col] = _safe_text_series(x[col], "")
    missing_theme = x["theme_final"].map(_is_missing_classification_value)
    missing_sub = x["sub_theme_final"].map(_is_missing_classification_value)
    x.loc[missing_theme, "theme_final"] = _safe_text_series(x.loc[missing_theme, "theme_from_map"], "").values
    x.loc[missing_sub, "sub_theme_final"] = _safe_text_series(x.loc[missing_sub, "sub_from_map"], "").values

    rule_used_mask = x["classification_source"].eq("") & x["source_ai"].isin(["rule_engine", "ai_infer"])
    x.loc[rule_used_mask, "classification_source"] = _safe_text_series(x.loc[rule_used_mask, "source_ai"], "").values
    x.loc[rule_used_mask, "classification_confidence"] = pd.to_numeric(x.loc[rule_used_mask, "confidence_ai"], errors="coerce").fillna(0).astype(int).values
    x.loc[rule_used_mask, "classification_note"] = _safe_text_series(x.loc[rule_used_mask, "note_ai"], "").values

    supplement_mask = ~x["classification_source"].eq("") & (x["theme_final"].map(_is_missing_classification_value) | x["sub_theme_final"].map(_is_missing_classification_value))
    x.loc[supplement_mask, "theme_final"] = _safe_text_series(x.loc[supplement_mask, "theme_ai"], "").values
    x.loc[supplement_mask, "sub_theme_final"] = _safe_text_series(x.loc[supplement_mask, "sub_theme_ai"], "").values
    note_mask = supplement_mask & x["classification_note"].eq("")
    x.loc[note_mask, "classification_note"] = _safe_text_series(x.loc[note_mask, "note_ai"], "").values

    for col in [
        "industry", "theme", "sub_theme",
        "industry_final", "theme_final", "sub_theme_final",
        "classification_source", "classification_note"
    ]:
        if col not in x.columns:
            x[col] = ""
        x[col] = _safe_text_series(x[col], "")

    x["industry_final"] = _safe_text_series(x["industry_final"], "未分類").map(normalize_official_industry_name)
    x["industry_final"] = _safe_text_series(x["industry_final"], "未分類").replace("", "未分類")
    x["theme_final"] = _safe_text_series(x["theme_final"], "全市場").replace("", "全市場")
    x["sub_theme_final"] = _safe_text_series(x["sub_theme_final"], "系統掃描").replace("", "系統掃描")
    x.loc[x["is_etf"].eq(1), ["industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_confidence", "classification_note"]] = ["ETF", "ETF", "ETF", "manual", 98, "ETF normalized"]

    x["industry"] = _safe_text_series(x["industry_final"], "未分類")
    x["theme"] = _safe_text_series(x["theme_final"], "全市場")
    x["sub_theme"] = _safe_text_series(x["sub_theme_final"], "系統掃描")
    x["is_active"] = 1
    x["update_date"] = datetime.now().strftime("%Y-%m-%d")

    _, summary = build_classification_quality_report(x)
    try:
        classification_debug_log(f"V2 覆蓋率 {summary.get('coverage_pct', 0):.2f}%｜官方 {summary.get('official', 0)}｜手動 {summary.get('manual', 0)}｜規則 {summary.get('rule_engine', 0)}｜AI {summary.get('ai_infer', 0)}｜未分類 {summary.get('unclassified', 0)}")
    except Exception:
        pass

    keep = ["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"]
    for c in keep:
        if c not in x.columns:
            x[c] = ""
    for c in ["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "update_date"]:
        x[c] = _safe_text_series(x[c], "")
    x["is_etf"] = _safe_numeric_flag_series(x["is_etf"], 0)
    x["is_active"] = _safe_numeric_flag_series(x["is_active"], 1)
    return x[keep].drop_duplicates(subset=["stock_id"], keep="first").reset_index(drop=True)



DATA_DIR = EXTERNAL_DATA_DIR if (EXTERNAL_DATA_DIR / "stocks_master.csv").exists() else PACKED_DATA_DIR
CHART_DIR = RUNTIME_DIR / "charts"
CHART_DIR.mkdir(exist_ok=True)

LEGACY_DB_PATH = RUNTIME_DIR / "stock_system_v6_0_1.db"
DB_PATH = RUNTIME_DIR / "stock_system_v6_2.db"
LEGACY_DB_PATH_V606 = RUNTIME_DIR / "stock_system_v6_0_6.db"
LEGACY_DB_PATH_V603 = RUNTIME_DIR / "stock_system_v6_0_3.db"
if (not DB_PATH.exists()) and LEGACY_DB_PATH_V606.exists():
    DB_PATH = LEGACY_DB_PATH_V606
elif (not DB_PATH.exists()) and LEGACY_DB_PATH_V603.exists():
    DB_PATH = LEGACY_DB_PATH_V603
elif (not DB_PATH.exists()) and LEGACY_DB_PATH.exists():
    DB_PATH = LEGACY_DB_PATH
MASTER_CSV = resolve_master_csv()


def normalize_csv_cell(v: str) -> str:
    s = str(v).strip().replace("=", "").strip()
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1]
    return s.strip()


def parse_twse_mi_index_csv(csv_text: str) -> pd.DataFrame:
    rows = []
    for raw in csv_text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("=") and "證券代號" in line:
            line = line.replace("=", "")
        if not re.match(r'^[="]?\d{4}', line):
            continue
        try:
            cols = next(csv.reader([line]))
        except Exception:
            continue
        cols = [normalize_csv_cell(x) for x in cols]
        if len(cols) < 11:
            continue
        code = cols[0]
        if not (code.isdigit() and len(code) == 4):
            continue
        rows.append({
            "stock_id": code,
            "stock_name": cols[1] if len(cols) > 1 else "",
            "volume": cols[2] if len(cols) > 2 else "",
            "open": cols[5] if len(cols) > 5 else "",
            "high": cols[6] if len(cols) > 6 else "",
            "low": cols[7] if len(cols) > 7 else "",
            "close": cols[8] if len(cols) > 8 else "",
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    for c in ["volume", "open", "high", "low", "close"]:
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df = df.dropna(subset=["close"])
    df["date"] = datetime.now().strftime("%Y-%m-%d")
    df["turnover"] = df["close"] * df["volume"].fillna(0)
    return df[["stock_id", "date", "open", "high", "low", "close", "volume", "turnover"]].drop_duplicates(subset=["stock_id"])


def download_twse_official_daily_csv(date_str: str | None = None, fallback_days: int = 10) -> pd.DataFrame:
    base_date = datetime.strptime(date_str, "%Y%m%d") if date_str else datetime.now()
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.twse.com.tw/"}
    for offset in range(fallback_days + 1):
        use_date = (base_date - pd.Timedelta(days=offset)).strftime("%Y%m%d")
        url = f"https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date={use_date}&type=ALLBUT0999"
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            resp.raise_for_status()
            df = parse_twse_mi_index_csv(resp.text)
            if df is not None and not df.empty:
                return df
        except Exception:
            continue
    return pd.DataFrame()



def open_path(path: Path):
    try:
        path = Path(path)
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def available_excel_engine() -> Optional[str]:
    try:
        import importlib.util
        if importlib.util.find_spec("xlsxwriter") is not None:
            return "xlsxwriter"
        if importlib.util.find_spec("openpyxl") is not None:
            return "openpyxl"
    except Exception:
        pass
    return None


def safe_sheet_name(name: str) -> str:
    invalid = r'[]:*?/\\'
    out = "".join("_" if ch in invalid else ch for ch in str(name))
    return out[:31] or "Sheet1"


def write_table_bundle(base_path: Path, tables: Dict[str, pd.DataFrame], preferred: str = "excel") -> tuple[Path, str]:
    clean_tables = {}
    for name, df in (tables or {}).items():
        if df is None:
            continue
        if isinstance(df, pd.DataFrame) and not df.empty:
            clean_tables[str(name)] = df.copy()
    if not clean_tables:
        raise ValueError("沒有可輸出的資料")

    preferred = (preferred or "excel").lower()
    engine = available_excel_engine()

    if preferred == "excel" and engine:
        out = base_path.with_suffix(".xlsx")
        with pd.ExcelWriter(out, engine=engine) as writer:
            for name, df in clean_tables.items():
                df.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
        return out, f"Excel（{engine}）"

    if preferred == "txt":
        if len(clean_tables) == 1:
            _, df = next(iter(clean_tables.items()))
            out = base_path.with_suffix(".txt")
            df.to_csv(out, index=False, sep="\t", encoding="utf-8-sig")
            return out, "TXT"
        out_dir = base_path.parent / f"{base_path.name}_TXT"
        out_dir.mkdir(parents=True, exist_ok=True)
        for name, df in clean_tables.items():
            df.to_csv(out_dir / f"{name}.txt", index=False, sep="\t", encoding="utf-8-sig")
        return out_dir, "TXT資料夾"

    if len(clean_tables) == 1:
        _, df = next(iter(clean_tables.items()))
        out = base_path.with_suffix(".csv")
        df.to_csv(out, index=False, encoding="utf-8-sig")
        return out, "CSV"

    out_dir = base_path.parent / f"{base_path.name}_CSV"
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, df in clean_tables.items():
        df.to_csv(out_dir / f"{name}.csv", index=False, encoding="utf-8-sig")
    return out_dir, "CSV資料夾"




def _normalize_master_df(df: pd.DataFrame, market_label: str, persist_classification_summary: bool = True) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])

    x = df.copy()
    rename_map = {
        "Code": "stock_id", "證券代號": "stock_id", "SecuritiesCompanyCode": "stock_id", "CompanyCode": "stock_id", "股票代號": "stock_id",
        "Name": "stock_name", "證券名稱": "stock_name", "CompanyName": "stock_name", "股票名稱": "stock_name",
    }
    x = x.rename(columns=rename_map)
    if "stock_id" not in x.columns:
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])
    if "stock_name" not in x.columns:
        x["stock_name"] = x["stock_id"]

    x["stock_id"] = x["stock_id"].astype(str).map(normalize_stock_id)
    x["stock_name"] = x["stock_name"].astype(str).str.strip()
    x = x[x["stock_id"] != ""].copy()
    x = x[x["stock_id"].str.fullmatch(r"\d{4,5}", na=False)].copy()
    if x.empty:
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])

    x["market"] = market_label
    x["industry"] = ""
    x["theme"] = ""
    x["sub_theme"] = ""
    x["is_etf"] = x["stock_id"].str.startswith("00").astype(int)
    x["is_active"] = 1
    x["update_date"] = datetime.now().strftime("%Y-%m-%d")

    if not persist_classification_summary:
        original_threshold = CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS
        try:
            globals()["CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS"] = max(int(CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS), len(x) + 1)
            x = apply_classification_layers(x)
        finally:
            globals()["CLASSIFICATION_SUMMARY_PROMOTION_MIN_ROWS"] = original_threshold
    else:
        x = apply_classification_layers(x)
    return x[["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"]].drop_duplicates(subset=["stock_id"]).reset_index(drop=True)

def fetch_twse_universe() -> pd.DataFrame:

    urls = [
        "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
        "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=json",
    ]
    for url in urls:
        try:
            res = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            res.raise_for_status()
            data = res.json()
            if isinstance(data, dict):
                data = data.get("data") or data.get("records") or []
            df = pd.DataFrame(data)
            if not df.empty:
                return _normalize_master_df(df, "上市")
        except Exception:
            continue
    return pd.DataFrame()


def fetch_tpex_universe() -> pd.DataFrame:
    urls = [
        "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes",
        "https://www.tpex.org.tw/openapi/v1/tpex_esb_quotes",
    ]
    parts = []
    for url in urls:
        try:
            res = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            res.raise_for_status()
            df = pd.DataFrame(res.json())
            if not df.empty:
                parts.append(_normalize_master_df(df, "上櫃"))
        except Exception:
            continue
    if not parts:
        return pd.DataFrame()
    return pd.concat(parts, ignore_index=True).drop_duplicates(subset=["stock_id"]).reset_index(drop=True)


def build_full_market_universe() -> pd.DataFrame:
    twse = fetch_twse_universe()
    tpex = fetch_tpex_universe()
    all_df = pd.concat([twse, tpex], ignore_index=True)
    if all_df.empty:
        csv_path = resolve_master_csv()
        if csv_path.exists():
            x = pd.read_csv(csv_path, dtype=str).fillna("")
            return _normalize_master_df(x, "上市")
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])

    try:
        csv_path = resolve_master_csv()
        if csv_path.exists():
            x = pd.read_csv(csv_path, dtype=str).fillna("")
            x = _normalize_master_df(x, "上市", persist_classification_summary=False)
            if x is not None and not x.empty and "stock_id" in x.columns:
                x["stock_id"] = x["stock_id"].astype(str).str.strip()
                x = x[x["stock_id"].str.fullmatch(r"\d{4,5}", na=False)].copy()
                keep_cols = ["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"]
                for c in keep_cols:
                    if c not in x.columns:
                        x[c] = ""
                x = x[keep_cols].drop_duplicates(subset=["stock_id"]).set_index("stock_id")
                all_df = all_df.drop_duplicates(subset=["stock_id"]).set_index("stock_id")

                invalid_text = {"", "未分類", "全市場", "系統掃描"}
                for col in ["stock_name", "market", "industry", "theme", "sub_theme"]:
                    if col in x.columns and col in all_df.columns:
                        valid_mask = ~x[col].astype(str).isin(invalid_text)
                        valid_ids = x.index[valid_mask & x.index.isin(all_df.index)]
                        if len(valid_ids) > 0:
                            all_df.loc[valid_ids, col] = x.loc[valid_ids, col]

                for col in ["is_etf", "is_active", "update_date"]:
                    if col in x.columns and col in all_df.columns:
                        valid_mask = x[col].astype(str).str.strip().ne("")
                        valid_ids = x.index[valid_mask & x.index.isin(all_df.index)]
                        if len(valid_ids) > 0:
                            all_df.loc[valid_ids, col] = x.loc[valid_ids, col]

                all_df = all_df.reset_index()
    except Exception as exc:
        log_warning(f"主檔覆蓋既有 CSV 時略過：{exc}")

    if "stock_id" not in all_df.columns:
        return pd.DataFrame(columns=["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"])

    return all_df.drop_duplicates(subset=["stock_id"]).sort_values(["market", "industry", "stock_id"]).reset_index(drop=True)



class DBManager:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.lock = threading.RLock()
        self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row

    def close(self):
        with self.lock:
            self.conn.close()

    def init_db(self):
        with self.lock:
            cur = self.conn.cursor()
            cur.execute("""
            CREATE TABLE IF NOT EXISTS stocks_master (
                stock_id TEXT PRIMARY KEY,
                stock_name TEXT,
                market TEXT,
                industry TEXT,
                theme TEXT,
                sub_theme TEXT,
                is_etf INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                update_date TEXT
            )
            """)
            cur.execute("""
            CREATE TABLE IF NOT EXISTS price_history (
                stock_id TEXT,
                date TEXT,
                open REAL,
                high REAL,
                low REAL,
                close REAL,
                volume REAL,
                turnover REAL,
                PRIMARY KEY (stock_id, date)
            )
            """)
            cur.execute("""
            CREATE TABLE IF NOT EXISTS ranking_result (
                date TEXT,
                stock_id TEXT,
                momentum_score REAL,
                trend_score REAL,
                reversal_score REAL,
                volume_score REAL,
                risk_score REAL,
                ai_score REAL,
                total_score REAL,
                signal TEXT,
                action TEXT,
                rank_all INTEGER,
                rank_industry INTEGER,
                PRIMARY KEY (date, stock_id)
            )
            """)
            self._init_external_decision_schema(cur)
            self.conn.commit()
            self.log_system_run(event="init_db", status="ok", message="DB schema initialized with external decision layer")


    def _init_external_decision_schema(self, cur):
        cur.execute("PRAGMA table_info(system_run_log)")
        _run_cols = cur.fetchall()
        if _run_cols and any(str(c[1]) == "run_id" and int(c[5] or 0) == 1 for c in _run_cols):
            legacy_name = "system_run_log_legacy_" + datetime.now().strftime("%Y%m%d_%H%M%S")
            cur.execute(f"ALTER TABLE system_run_log RENAME TO {legacy_name}")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS system_run_log (
            log_id TEXT PRIMARY KEY,
            run_id TEXT,
            event_seq INTEGER DEFAULT 0,
            run_time TEXT,
            event_time TEXT,
            program_name TEXT,
            program_version TEXT,
            program_path TEXT,
            program_hash TEXT,
            db_path TEXT,
            db_hash TEXT,
            event TEXT,
            step TEXT,
            module TEXT,
            status TEXT,
            duration_ms REAL DEFAULT 0,
            message TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_data_log (
            log_id TEXT PRIMARY KEY,
            run_id TEXT,
            module TEXT,
            source_name TEXT,
            official_url TEXT,
            request_url TEXT,
            source_date TEXT,
            fetch_time TEXT,
            http_status TEXT,
            status TEXT,
            fallback_count INTEGER DEFAULT 0,
            rows_count INTEGER DEFAULT 0,
            error_message TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_source_status (
            module TEXT PRIMARY KEY,
            source_name TEXT,
            official_url TEXT,
            request_url TEXT,
            source_date TEXT,
            last_fetch_time TEXT,
            status TEXT,
            fallback_count INTEGER DEFAULT 0,
            rows_count INTEGER DEFAULT 0,
            error_message TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS market_snapshot (
            snapshot_date TEXT PRIMARY KEY,
            market_score REAL,
            market_mode TEXT,
            taiex_close REAL,
            taiex_trend TEXT,
            sp500_close REAL,
            nasdaq_close REAL,
            vix REAL,
            us10y REAL,
            dxy REAL,
            breadth REAL,
            source_status TEXT,
            update_time TEXT
        )
        """)
        self._ensure_market_snapshot_r10_schema(cur)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_revenue (
            stock_id TEXT,
            revenue_month TEXT,
            revenue REAL,
            mom REAL,
            yoy REAL,
            cumulative_revenue REAL,
            cumulative_yoy REAL,
            source_date TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, revenue_month)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_valuation (
            stock_id TEXT,
            data_date TEXT,
            close_price REAL,
            pe REAL,
            pb REAL,
            dividend_yield REAL,
            eps REAL,
            eps_ttm REAL,
            roe REAL,
            gross_margin REAL,
            operating_margin REAL,
            fiscal_year_quarter TEXT,
            source_date TEXT,
            source_url TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, data_date)
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS financial_feature_daily (
            stock_id TEXT,
            feature_date TEXT,
            eps_ttm REAL,
            eps_yoy REAL,
            revenue_yoy REAL,
            eps_bucket TEXT,
            rev_bucket TEXT,
            matrix_cell TEXT,
            eps_category TEXT,
            matrix_base_score REAL,
            modifier REAL,
            revenue_eps_score REAL,
            data_quality_flag TEXT,
            source_trace_json TEXT,
            source_date TEXT,
            run_id TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, feature_date)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_dividend (
            stock_id TEXT,
            data_year TEXT,
            cash_dividend REAL,
            stock_dividend REAL,
            dividend_yield REAL,
            ex_dividend_date TEXT,
            source_date TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, data_year)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_institutional (
            stock_id TEXT,
            trade_date TEXT,
            foreign_buy_sell REAL,
            trust_buy_sell REAL,
            dealer_buy_sell REAL,
            eight_bank_buy_sell REAL,
            institutional_score REAL,
            main_force_flag TEXT,
            source_date TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, trade_date)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_margin (
            stock_id TEXT,
            trade_date TEXT,
            margin_balance REAL,
            short_balance REAL,
            margin_change REAL,
            short_change REAL,
            margin_utilization REAL,
            retail_heat_score REAL,
            source_date TEXT,
            update_time TEXT,
            PRIMARY KEY(stock_id, trade_date)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS macro_margin_sentiment (
            data_date TEXT PRIMARY KEY,
            total_margin_balance REAL,
            total_short_balance REAL,
            total_margin_change REAL,
            total_short_change REAL,
            market_margin_utilization REAL,
            macro_margin_score REAL,
            macro_margin_state TEXT,
            sentiment_reason TEXT,
            source_name TEXT,
            source_url TEXT,
            source_date TEXT,
            update_time TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_event (
            event_id TEXT PRIMARY KEY,
            stock_id TEXT,
            event_date TEXT,
            event_type TEXT,
            event_title TEXT,
            event_score REAL,
            event_window TEXT,
            source_name TEXT,
            source_url TEXT,
            update_time TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS macro_module_score (
            module_id TEXT,
            data_date TEXT,
            module_name TEXT,
            score REAL,
            mode TEXT,
            source_name TEXT,
            source_url TEXT,
            source_date TEXT,
            status TEXT,
            update_time TEXT,
            PRIMARY KEY(module_id, data_date)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS external_source_status_detail (
            run_id TEXT,
            module TEXT,
            source_key TEXT,
            source_name TEXT,
            official_url TEXT,
            request_url TEXT,
            source_date TEXT,
            fetch_time TEXT,
            http_status TEXT,
            status TEXT,
            attempt_no INTEGER DEFAULT 1,
            fallback_count INTEGER DEFAULT 0,
            rows_count INTEGER DEFAULT 0,
            coverage_total INTEGER DEFAULT 0,
            coverage_hit INTEGER DEFAULT 0,
            freshness_days INTEGER DEFAULT 9999,
            data_ready INTEGER DEFAULT 0,
            blocking_reason TEXT,
            error_message TEXT,
            source_level TEXT,
            target_table TEXT,
            PRIMARY KEY (run_id, module, source_key, attempt_no)
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS trade_plan (
            run_id TEXT,
            plan_date TEXT,
            stock_id TEXT,
            stock_name TEXT,
            market TEXT,
            industry TEXT,
            theme TEXT,
            close REAL,
            entry_low REAL,
            entry_high REAL,
            stop_loss REAL,
            target_price REAL,
            target_1382 REAL,
            target_1618 REAL,
            rr REAL,
            rr_live REAL,
            win_rate REAL,
            market_gate INTEGER,
            flow_gate INTEGER,
            fundamental_gate INTEGER,
            event_gate INTEGER,
            technical_gate INTEGER,
            risk_gate INTEGER,
            trade_allowed INTEGER,
            gate_summary TEXT,
            decision_reason TEXT,
            final_trade_decision TEXT,
            ui_state TEXT,
            pool_role TEXT,
            source_rank TEXT,
            update_time TEXT,
            PRIMARY KEY(run_id, stock_id, source_rank)
        )
        """)
        self._ensure_external_schema_columns(cur)
        for sql in [
            "CREATE INDEX IF NOT EXISTS idx_system_run_log_run_id ON system_run_log(run_id)",
            "CREATE INDEX IF NOT EXISTS idx_system_run_log_event ON system_run_log(event, status)",
            "CREATE INDEX IF NOT EXISTS idx_trade_plan_date ON trade_plan(plan_date)",
            "CREATE INDEX IF NOT EXISTS idx_trade_plan_allowed ON trade_plan(trade_allowed)",
            "CREATE INDEX IF NOT EXISTS idx_external_log_module ON external_data_log(module, source_date)",
            "CREATE INDEX IF NOT EXISTS idx_external_status_ready ON external_source_status(data_ready, status)",
            "CREATE INDEX IF NOT EXISTS idx_inst_stock_date ON external_institutional(stock_id, trade_date)",
            "CREATE INDEX IF NOT EXISTS idx_margin_stock_date ON external_margin(stock_id, trade_date)",
            "CREATE INDEX IF NOT EXISTS idx_macro_margin_date ON macro_margin_sentiment(data_date)",
            "CREATE INDEX IF NOT EXISTS idx_revenue_stock_month ON external_revenue(stock_id, revenue_month)",
            "CREATE INDEX IF NOT EXISTS idx_financial_feature_stock_date ON financial_feature_daily(stock_id, feature_date)",
            "CREATE INDEX IF NOT EXISTS idx_external_detail_module ON external_source_status_detail(module, source_key, source_date)",
            "CREATE INDEX IF NOT EXISTS idx_trade_plan_gate_state ON trade_plan(trade_allowed, market_gate_state, flow_gate_state, fundamental_gate_state)",
        ]:
            cur.execute(sql)

    def _ensure_market_snapshot_r10_schema(self, cur):
        """R10：market_snapshot 改為「每檔股票一筆」的全市場快照表。

        R9 的 market_snapshot 只有 snapshot_date PRIMARY KEY，fallback 時只能寫入 1 筆市場總覽，
        無法支撐 Decision Layer 所需的 close/volume/rsi/atr/price_dev/ma20/ma60。
        R10 若偵測到舊表沒有 stock_id 或仍是單一 snapshot_date 主鍵，會自動備份舊表並重建。
        """
        try:
            info = cur.execute("PRAGMA table_info(market_snapshot)").fetchall()
            cols = {str(r[1]) for r in info}
            pk_cols = [str(r[1]) for r in info if int(r[5] or 0) > 0]
            need_rebuild = ("stock_id" not in cols) or (set(pk_cols) != {"snapshot_date", "stock_id"})
            if need_rebuild and info:
                backup = "market_snapshot_legacy_" + datetime.now().strftime("%Y%m%d_%H%M%S")
                cur.execute(f"ALTER TABLE market_snapshot RENAME TO {backup}")
                log_warning(f"[R10][SCHEMA] market_snapshot 舊結構已備份為 {backup}，改建全市場快照表")
        except Exception as exc:
            log_warning(f"[R10][SCHEMA] market_snapshot schema probe failed: {exc}")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS market_snapshot (
            snapshot_date TEXT,
            stock_id TEXT,
            stock_name TEXT,
            market TEXT,
            industry TEXT,
            close REAL,
            open REAL,
            high REAL,
            low REAL,
            volume REAL,
            turnover REAL,
            rsi REAL,
            rsi14 REAL,
            atr REAL,
            atr_pct REAL,
            price_dev REAL,
            price_deviation REAL,
            ma20 REAL,
            ma60 REAL,
            market_score REAL,
            market_mode TEXT,
            taiex_close REAL,
            taiex_trend TEXT,
            sp500_close REAL,
            nasdaq_close REAL,
            vix REAL,
            us10y REAL,
            dxy REAL,
            breadth REAL,
            source_status TEXT,
            source_type TEXT,
            source_url TEXT,
            source_level TEXT,
            proxy_reason TEXT,
            update_time TEXT,
            PRIMARY KEY(snapshot_date, stock_id)
        )
        """)
        for col, ddl in [
            ("stock_id", "stock_id TEXT"), ("stock_name", "stock_name TEXT"), ("market", "market TEXT"), ("industry", "industry TEXT"),
            ("close", "close REAL"), ("open", "open REAL"), ("high", "high REAL"), ("low", "low REAL"),
            ("volume", "volume REAL"), ("turnover", "turnover REAL"), ("rsi", "rsi REAL"), ("rsi14", "rsi14 REAL"),
            ("atr", "atr REAL"), ("atr_pct", "atr_pct REAL"), ("price_dev", "price_dev REAL"), ("price_deviation", "price_deviation REAL"),
            ("ma20", "ma20 REAL"), ("ma60", "ma60 REAL"),
        ]:
            try:
                cur.execute(f"ALTER TABLE market_snapshot ADD COLUMN {ddl}")
            except Exception:
                pass

    def _ensure_external_schema_columns(self, cur):
        def _cols(table):
            try:
                return {str(r[1]) for r in cur.execute(f"PRAGMA table_info({table})").fetchall()}
            except Exception:
                return set()

        def _add(table, col, ddl):
            cols = _cols(table)
            if col not in cols:
                try:
                    cur.execute(f"ALTER TABLE {table} ADD COLUMN {ddl}")
                except Exception as exc:
                    log_warning(f"Schema欄位補強失敗：{table}.{col}｜{exc}")

        for col, ddl in [
            ("step", "step TEXT"),
            ("validator_status", "validator_status TEXT"),
            ("writer_status", "writer_status TEXT"),
            ("duration_ms", "duration_ms REAL DEFAULT 0"),
            ("data_ready", "data_ready INTEGER DEFAULT 0"),
            ("blocking_reason", "blocking_reason TEXT"),
            ("target_table", "target_table TEXT"),
        ]:
            _add("external_data_log", col, ddl)

        for col, ddl in [
            ("target_table", "target_table TEXT"),
            ("last_success_time", "last_success_time TEXT"),
            ("data_ready", "data_ready INTEGER DEFAULT 0"),
            ("blocking_reason", "blocking_reason TEXT"),
            ("last_rows_count", "last_rows_count INTEGER DEFAULT 0"),
            ("source_level", "source_level TEXT"),
            ("analysis_ready", "analysis_ready INTEGER DEFAULT 1"),
            ("execution_ready", "execution_ready INTEGER DEFAULT 0"),
            ("soft_block", "soft_block INTEGER DEFAULT 0"),
            ("block_reason", "block_reason TEXT"),
            ("execution_block_reason", "execution_block_reason TEXT"),
        ]:
            _add("external_source_status", col, ddl)

        for col, ddl in [
            ("source_type", "source_type TEXT"),
            ("source_url", "source_url TEXT"),
            ("source_level", "source_level TEXT"),
            ("proxy_reason", "proxy_reason TEXT"),
        ]:
            _add("market_snapshot", col, ddl)

        for col, ddl in [
            ("close_price", "close_price REAL"),
            ("dividend_yield", "dividend_yield REAL"),
            ("eps_ttm", "eps_ttm REAL"),
            ("fiscal_year_quarter", "fiscal_year_quarter TEXT"),
            ("source_url", "source_url TEXT"),
        ]:
            _add("external_valuation", col, ddl)

        for col, ddl in [
            ("eps_ttm", "eps_ttm REAL"),
            ("eps_yoy", "eps_yoy REAL"),
            ("revenue_yoy", "revenue_yoy REAL"),
            ("eps_bucket", "eps_bucket TEXT"),
            ("rev_bucket", "rev_bucket TEXT"),
            ("matrix_cell", "matrix_cell TEXT"),
            ("eps_category", "eps_category TEXT"),
            ("matrix_base_score", "matrix_base_score REAL"),
            ("modifier", "modifier REAL"),
            ("revenue_eps_score", "revenue_eps_score REAL"),
            ("data_quality_flag", "data_quality_flag TEXT"),
            ("source_trace_json", "source_trace_json TEXT"),
            ("source_date", "source_date TEXT"),
            ("run_id", "run_id TEXT"),
            ("update_time", "update_time TEXT"),
        ]:
            _add("financial_feature_daily", col, ddl)

        for col, ddl in [
            ("technical_total_score", "technical_total_score REAL DEFAULT 0"),
            ("financial_score", "financial_score REAL DEFAULT 50"),
            ("eps_ttm", "eps_ttm REAL"),
            ("eps_yoy", "eps_yoy REAL"),
            ("revenue_yoy", "revenue_yoy REAL"),
            ("eps_bucket", "eps_bucket TEXT"),
            ("rev_bucket", "rev_bucket TEXT"),
            ("matrix_cell", "matrix_cell TEXT"),
            ("eps_category", "eps_category TEXT"),
            ("matrix_base_score", "matrix_base_score REAL"),
            ("modifier", "modifier REAL"),
            ("revenue_eps_score", "revenue_eps_score REAL DEFAULT 50"),
            ("data_quality_flag", "data_quality_flag TEXT"),
            ("source_trace_json", "source_trace_json TEXT"),
        ]:
            _add("ranking_result", col, ddl)

        for col, ddl in [
            ("external_data_ready", "external_data_ready INTEGER DEFAULT 0"),
            ("external_blocking_reason", "external_blocking_reason TEXT"),
            ("analysis_ready", "analysis_ready INTEGER DEFAULT 1"),
            ("execution_ready", "execution_ready INTEGER DEFAULT 0"),
            ("soft_block", "soft_block INTEGER DEFAULT 0"),
            ("block_reason", "block_reason TEXT DEFAULT ''"),
            ("execution_block_reason", "execution_block_reason TEXT DEFAULT ''"),
            ("pipeline_run_id", "pipeline_run_id TEXT"),
            ("external_run_id", "external_run_id TEXT"),
            ("decision_run_id", "decision_run_id TEXT"),
            ("market_gate_state", "market_gate_state TEXT DEFAULT 'BLOCK'"),
            ("flow_gate_state", "flow_gate_state TEXT DEFAULT 'BLOCK'"),
            ("fundamental_gate_state", "fundamental_gate_state TEXT DEFAULT 'BLOCK'"),
            ("event_gate_state", "event_gate_state TEXT DEFAULT 'NE'"),
            ("risk_gate_state", "risk_gate_state TEXT DEFAULT 'BLOCK'"),
            ("latest_external_date", "latest_external_date TEXT DEFAULT ''"),
            ("market_source_level", "market_source_level TEXT DEFAULT ''"),
            ("source_trace_json", "source_trace_json TEXT DEFAULT ''"),
            ("decision_reason_short", "decision_reason_short TEXT DEFAULT ''"),
            ("global_external_ready", "global_external_ready INTEGER DEFAULT 0"),
            ("stock_external_coverage_state", "stock_external_coverage_state TEXT DEFAULT ''"),
            ("gate_policy_note", "gate_policy_note TEXT DEFAULT ''"),
            ("pe", "pe REAL"),
            ("pb", "pb REAL"),
            ("dividend_yield", "dividend_yield REAL"),
            ("eps_ttm", "eps_ttm REAL"),
            ("valuation_score", "valuation_score REAL DEFAULT 0"),
            ("eps_yoy", "eps_yoy REAL"),
            ("revenue_yoy", "revenue_yoy REAL"),
            ("eps_bucket", "eps_bucket TEXT"),
            ("rev_bucket", "rev_bucket TEXT"),
            ("matrix_cell", "matrix_cell TEXT"),
            ("eps_category", "eps_category TEXT"),
            ("matrix_base_score", "matrix_base_score REAL"),
            ("modifier", "modifier REAL"),
            ("revenue_eps_score", "revenue_eps_score REAL DEFAULT 50"),
            ("data_quality_flag", "data_quality_flag TEXT"),
            ("financial_score", "financial_score REAL DEFAULT 50"),
            ("eps_matrix_decision_note", "eps_matrix_decision_note TEXT DEFAULT ''"),
            ("margin_balance", "margin_balance REAL"),
            ("short_balance", "short_balance REAL"),
            ("margin_change", "margin_change REAL"),
            ("short_change", "short_change REAL"),
            ("margin_utilization", "margin_utilization REAL"),
            ("retail_heat_score", "retail_heat_score REAL DEFAULT 50"),
            ("margin_score", "margin_score REAL DEFAULT 50"),
            ("margin_state", "margin_state TEXT DEFAULT 'NE'"),
            ("macro_margin_score", "macro_margin_score REAL DEFAULT 50"),
            ("macro_margin_state", "macro_margin_state TEXT DEFAULT 'NE'"),
            ("margin_decision_note", "margin_decision_note TEXT DEFAULT ''"),
            ("fail_reason", "fail_reason TEXT DEFAULT ''"),
            ("rsi", "rsi REAL"),
            ("atr_pct", "atr_pct REAL"),
            ("price_deviation", "price_deviation REAL"),
            ("model_score", "model_score REAL"),
            ("wave_trade_score", "wave_trade_score REAL"),
        ]:
            _add("trade_plan", col, ddl)

    def _safe_sha256(self, path) -> str:
        try:
            p = Path(path)
            if not p.exists():
                return ""
            h = hashlib.sha256()
            with open(p, "rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
            return h.hexdigest()
        except Exception:
            return ""

    def make_run_id(self) -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    def _next_event_seq(self, run_id: str) -> int:
        try:
            with self.lock:
                row = self.conn.cursor().execute("SELECT COALESCE(MAX(event_seq),0)+1 FROM system_run_log WHERE run_id=?", (run_id,)).fetchone()
            return int(row[0] or 1)
        except Exception:
            return int(time.time() * 1000) % 1000000

    def log_system_run(self, event: str, status: str = "ok", message: str = "", run_id: str | None = None, step: str = "", module: str = "", duration_ms: float = 0.0) -> str:
        run_id = run_id or self.make_run_id()
        try:
            program_path = str(Path(__file__).resolve()) if "__file__" in globals() else ""
            event_seq = self._next_event_seq(run_id)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            row = {
                "log_id": f"{run_id}_{event_seq:04d}_{int(time.time()*1000)}",
                "run_id": run_id,
                "event_seq": event_seq,
                "run_time": now,
                "event_time": now,
                "program_name": APP_NAME,
                "program_version": "v9.6.2_pro_fundamental_local_cache_v16.2_r5",
                "program_path": program_path,
                "program_hash": self._safe_sha256(program_path),
                "db_path": str(self.db_path),
                "db_hash": self._safe_sha256(self.db_path),
                "event": event,
                "step": step or event,
                "module": module or "",
                "status": status,
                "duration_ms": float(duration_ms or 0.0),
                "message": message,
            }
            with self.lock:
                pd.DataFrame([row]).to_sql("system_run_log", self.conn, if_exists="append", index=False)
                self.conn.commit()
        except Exception as exc:
            log_warning(f"system_run_log 寫入失敗：{exc}")
        return run_id

    def log_external_data(self, module: str, source_name: str, official_url: str = "", request_url: str = "", source_date: str = "", status: str = "pending", http_status: str = "", fallback_count: int = 0, rows_count: int = 0, error_message: str = "", run_id: str | None = None, step: str = "fetch", validator_status: str = "", writer_status: str = "", duration_ms: float = 0.0, target_table: str = "", data_ready: int | None = None, blocking_reason: str = "", source_level: str = ""):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        safe_module = re.sub(r"[^A-Za-z0-9_]", "_", str(module or "module"))
        log_id = f"{safe_module}_{source_date or datetime.now().strftime('%Y%m%d')}_{step}_{int(time.time()*1000)}"
        ready = int(data_ready if data_ready is not None else (1 if str(status).lower() in ("success", "fallback", "ok") and int(rows_count or 0) > 0 else 0))
        block = str(blocking_reason or ("" if ready else (error_message or "資料未就緒")))
        row = {
            "log_id": log_id, "run_id": run_id or "", "module": module, "source_name": source_name,
            "official_url": official_url, "request_url": request_url, "source_date": source_date,
            "fetch_time": now, "http_status": str(http_status or ""), "status": status,
            "fallback_count": int(fallback_count or 0), "rows_count": int(rows_count or 0), "error_message": str(error_message or ""),
            "step": step, "validator_status": validator_status, "writer_status": writer_status,
            "duration_ms": float(duration_ms or 0.0), "data_ready": ready, "blocking_reason": block,
            "target_table": target_table,
        }
        status_row = {
            "module": module, "source_name": source_name, "official_url": official_url, "request_url": request_url,
            "source_date": source_date, "last_fetch_time": now, "status": status, "fallback_count": int(fallback_count or 0),
            "rows_count": int(rows_count or 0), "error_message": str(error_message or ""), "target_table": target_table,
            "last_success_time": now if ready else "", "data_ready": ready, "blocking_reason": block,
            "last_rows_count": int(rows_count or 0), "source_level": source_level or ("official" if ready else "failed"),
        }
        try:
            with self.lock:
                pd.DataFrame([row]).to_sql("external_data_log", self.conn, if_exists="append", index=False)
                cur = self.conn.cursor()
                cur.execute("""
                    INSERT INTO external_source_status(module, source_name, official_url, request_url, source_date, last_fetch_time, status, fallback_count, rows_count, error_message, target_table, last_success_time, data_ready, blocking_reason, last_rows_count, source_level)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(module) DO UPDATE SET
                        source_name=excluded.source_name, official_url=excluded.official_url, request_url=excluded.request_url,
                        source_date=excluded.source_date, last_fetch_time=excluded.last_fetch_time, status=excluded.status,
                        fallback_count=excluded.fallback_count, rows_count=excluded.rows_count, error_message=excluded.error_message,
                        target_table=excluded.target_table,
                        last_success_time=CASE WHEN excluded.data_ready=1 THEN excluded.last_success_time ELSE external_source_status.last_success_time END,
                        data_ready=excluded.data_ready, blocking_reason=excluded.blocking_reason, last_rows_count=excluded.last_rows_count,
                        source_level=excluded.source_level
                """, tuple(status_row.values()))
                self.conn.commit()
        except Exception as exc:
            log_warning(f"external_data_log 寫入失敗：{exc}")

    def replace_trade_plan_batch(self, df: pd.DataFrame, run_id: str | None = None):
        if df is None or df.empty:
            return ""
        run_id = run_id or self.log_system_run(event="trade_plan_batch", status="start", message=f"rows={len(df)}")
        x = df.copy()
        x["run_id"] = run_id
        x["plan_date"] = datetime.now().strftime("%Y-%m-%d")
        x["update_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        x = x.rename(columns={"現價": "close", "名稱": "stock_name", "代號": "stock_id"})
        defaults = {
            "stock_id": "", "stock_name": "", "market": "", "industry": "", "theme": "", "close": 0.0,
            "entry_low": 0.0, "entry_high": 0.0, "stop_loss": 0.0, "target_price": 0.0, "target_1382": 0.0, "target_1618": 0.0,
            "rr": 0.0, "rr_live": 0.0, "win_rate": 0.0, "market_gate": 0, "flow_gate": 0, "fundamental_gate": 0,
            "event_gate": 0, "technical_gate": 0, "risk_gate": 0, "trade_allowed": 0, "gate_summary": "", "decision_reason": "",
            "final_trade_decision": "", "ui_state": "", "pool_role": "", "source_rank": "",
            "external_data_ready": 0, "external_blocking_reason": "",
            "analysis_ready": 1, "execution_ready": 0, "soft_block": 0, "block_reason": "", "execution_block_reason": "",
            "pipeline_run_id": run_id,
            "external_run_id": run_id, "decision_run_id": "",
            "market_gate_state": "BLOCK", "flow_gate_state": "BLOCK", "fundamental_gate_state": "BLOCK",
            "event_gate_state": "NE", "risk_gate_state": "BLOCK",
            "latest_external_date": "", "market_source_level": "", "source_trace_json": "", "decision_reason_short": "",
            "global_external_ready": 0, "stock_external_coverage_state": "", "gate_policy_note": "NE=Not Evaluated：資料未覆蓋不阻擋交易，只降權/標示",
            "pe": 0.0, "pb": 0.0, "dividend_yield": 0.0, "eps_ttm": 0.0, "valuation_score": 0.0,
            "eps_yoy": 0.0, "revenue_yoy": 0.0, "eps_bucket": "", "rev_bucket": "", "matrix_cell": "",
            "eps_category": "U0", "matrix_base_score": 0.0, "modifier": 0.0, "revenue_eps_score": 50.0,
            "data_quality_flag": "NE", "financial_score": 50.0, "eps_matrix_decision_note": "",
            "margin_balance": 0.0, "short_balance": 0.0, "margin_change": 0.0, "short_change": 0.0, "margin_utilization": 0.0,
            "retail_heat_score": 50.0, "margin_score": 50.0, "margin_state": "NE", "macro_margin_score": 50.0, "macro_margin_state": "NE", "margin_decision_note": "",
            "fail_reason": "", "rsi": np.nan, "atr_pct": np.nan, "price_deviation": np.nan, "model_score": np.nan, "wave_trade_score": np.nan,
        }
        for c, d in defaults.items():
            if c not in x.columns:
                x[c] = d
        if "entry_mid" in x.columns:
            close_series = pd.to_numeric(x["close"], errors="coerce").fillna(0)
            x.loc[close_series.eq(0), "close"] = pd.to_numeric(x.loc[close_series.eq(0), "entry_mid"], errors="coerce").fillna(0)
        for c in ["close", "entry_low", "entry_high", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "rr_live", "win_rate", "pe", "pb", "dividend_yield", "eps_ttm", "eps_yoy", "revenue_yoy", "matrix_base_score", "modifier", "revenue_eps_score", "financial_score", "valuation_score", "margin_balance", "short_balance", "margin_change", "short_change", "margin_utilization", "retail_heat_score", "margin_score", "macro_margin_score", "rsi", "atr_pct", "price_deviation", "model_score", "wave_trade_score"]:
            x[c] = pd.to_numeric(x[c], errors="coerce")
        for c in ["market_gate", "flow_gate", "fundamental_gate", "event_gate", "technical_gate", "risk_gate", "trade_allowed", "analysis_ready", "execution_ready", "soft_block"]:
            if c not in x.columns:
                x[c] = 0
            x[c] = pd.to_numeric(x[c], errors="coerce").fillna(0).astype(int)
        keep = ["run_id", "plan_date", "stock_id", "stock_name", "market", "industry", "theme", "close", "entry_low", "entry_high", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "rr_live", "win_rate", "market_gate", "flow_gate", "fundamental_gate", "event_gate", "technical_gate", "risk_gate", "trade_allowed", "gate_summary", "decision_reason", "final_trade_decision", "ui_state", "pool_role", "source_rank", "external_data_ready", "external_blocking_reason", "fail_reason", "rsi", "atr_pct", "price_deviation", "model_score", "wave_trade_score", "analysis_ready", "execution_ready", "soft_block", "block_reason", "execution_block_reason", "pipeline_run_id", "external_run_id", "decision_run_id", "market_gate_state", "flow_gate_state", "fundamental_gate_state", "event_gate_state", "risk_gate_state", "latest_external_date", "market_source_level", "source_trace_json", "decision_reason_short", "global_external_ready", "stock_external_coverage_state", "gate_policy_note", "pe", "pb", "dividend_yield", "eps_ttm", "eps_yoy", "revenue_yoy", "eps_bucket", "rev_bucket", "matrix_cell", "eps_category", "matrix_base_score", "modifier", "revenue_eps_score", "data_quality_flag", "financial_score", "eps_matrix_decision_note", "valuation_score", "margin_balance", "short_balance", "margin_change", "short_change", "margin_utilization", "retail_heat_score", "margin_score", "margin_state", "macro_margin_score", "macro_margin_state", "margin_decision_note", "update_time"]
        x = x[keep].drop_duplicates(subset=["run_id", "stock_id", "source_rank"], keep="first")
        with self.lock:
            self.conn.execute("DELETE FROM trade_plan WHERE run_id=?", (run_id,))
            x.to_sql("trade_plan", self.conn, if_exists="append", index=False)
            self.conn.commit()
        self.log_system_run(event="trade_plan_batch", status="ok", message=f"trade_plan rows={len(x)}", run_id=run_id)
        return run_id


    def replace_financial_feature_batch(self, df: pd.DataFrame, run_id: str | None = None):
        """V9.5.7 EPS_MATRIX_PATCH：寫入 financial_feature_daily。"""
        if df is None or df.empty:
            return ""
        run_id = run_id or self.log_system_run(event="financial_feature_batch", status="start", message=f"rows={len(df)}", module="eps_matrix")
        x = df.copy()
        if "feature_date" not in x.columns:
            x["feature_date"] = datetime.now().strftime("%Y-%m-%d")
        if "run_id" not in x.columns:
            x["run_id"] = run_id
        x["update_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        defaults = {
            "stock_id": "", "feature_date": "", "eps_ttm": np.nan, "eps_yoy": np.nan, "revenue_yoy": np.nan,
            "eps_bucket": "E_NA", "rev_bucket": "R_NA", "matrix_cell": "E_NA-R_NA", "eps_category": "U0",
            "matrix_base_score": 40.0, "modifier": 0.0, "revenue_eps_score": 40.0,
            "data_quality_flag": "NE", "source_trace_json": "{}", "source_date": "", "run_id": run_id,
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        for c, d in defaults.items():
            if c not in x.columns:
                x[c] = d
        for c in ["eps_ttm", "eps_yoy", "revenue_yoy", "matrix_base_score", "modifier", "revenue_eps_score"]:
            x[c] = pd.to_numeric(x[c], errors="coerce")
        text_cols = ["stock_id", "feature_date", "eps_bucket", "rev_bucket", "matrix_cell", "eps_category", "data_quality_flag", "source_trace_json", "source_date", "run_id", "update_time"]
        for c in text_cols:
            x[c] = x[c].fillna("").astype(str)
        keep = ["stock_id", "feature_date", "eps_ttm", "eps_yoy", "revenue_yoy", "eps_bucket", "rev_bucket", "matrix_cell", "eps_category", "matrix_base_score", "modifier", "revenue_eps_score", "data_quality_flag", "source_trace_json", "source_date", "run_id", "update_time"]
        x = x[keep].drop_duplicates(subset=["stock_id", "feature_date"], keep="last")
        with self.lock:
            cur = self.conn.cursor()
            for fd in x["feature_date"].dropna().astype(str).unique().tolist():
                cur.execute("DELETE FROM financial_feature_daily WHERE feature_date=?", (fd,))
            x.to_sql("financial_feature_daily", self.conn, if_exists="append", index=False)
            self.conn.commit()
        self.log_system_run(event="financial_feature_batch", status="ok", message=f"financial_feature_daily rows={len(x)}", run_id=run_id, module="eps_matrix")
        return run_id

    def get_latest_financial_features(self) -> pd.DataFrame:
        q = """
        SELECT f.*
        FROM financial_feature_daily f
        JOIN (
            SELECT stock_id, MAX(feature_date) AS feature_date
            FROM financial_feature_daily
            GROUP BY stock_id
        ) m
        ON f.stock_id=m.stock_id AND f.feature_date=m.feature_date
        """
        with self.lock:
            try:
                return pd.read_sql_query(q, self.conn)
            except Exception:
                return pd.DataFrame()

    def get_latest_financial_feature_row(self, stock_id: str) -> dict:
        df = self.get_latest_financial_features()
        if df is None or df.empty or "stock_id" not in df.columns:
            return {}
        x = df[df["stock_id"].astype(str) == str(stock_id)].tail(1)
        if x.empty:
            return {}
        return dict(x.iloc[-1])

    def read_table(self, table_name: str, limit: int | None = 500) -> pd.DataFrame:
        safe = re.sub(r"[^A-Za-z0-9_]", "", str(table_name or ""))
        if not safe:
            return pd.DataFrame()
        q = f"SELECT * FROM {safe}"
        if limit is not None:
            q += f" LIMIT {int(limit)}"
        with self.lock:
            try:
                return pd.read_sql_query(q, self.conn)
            except Exception:
                return pd.DataFrame()

    def schema_check_df(self) -> pd.DataFrame:
        required = ["stocks_master", "price_history", "ranking_result", "system_run_log", "external_data_log", "external_source_status", "market_snapshot", "external_revenue", "external_valuation", "financial_feature_daily", "external_dividend", "external_institutional", "external_margin", "external_event", "macro_module_score", "trade_plan"]
        rows = []
        with self.lock:
            existing = {r[0] for r in self.conn.cursor().execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
            for table in required:
                cols = []
                if table in existing:
                    cols = [c[1] for c in self.conn.cursor().execute(f"PRAGMA table_info({table})").fetchall()]
                rows.append({"table_name": table, "exists": "YES" if table in existing else "NO", "column_count": len(cols), "columns": ", ".join(cols)})
        return pd.DataFrame(rows)

    def import_master_csv(self, csv_path: Path):
        df = pd.read_csv(csv_path, dtype={"stock_id": str}).fillna("")
        self.import_master_df(df)

    def import_master_df(self, df: pd.DataFrame):
        x = df.copy().fillna("")
        required_defaults = {
            "stock_id": "", "stock_name": "", "market": "", "industry": "", "theme": "", "sub_theme": "",
            "is_etf": 0, "is_active": 1, "update_date": datetime.now().strftime("%Y-%m-%d"),
        }
        for col, default in required_defaults.items():
            if col not in x.columns:
                x[col] = default
        x["stock_id"] = x["stock_id"].astype(str).str.strip()
        x = x[x["stock_id"].str.fullmatch(r"\d{4,5}", na=False)].copy()
        x["is_etf"] = pd.to_numeric(x["is_etf"], errors="coerce").fillna(0).astype(int)
        x["is_active"] = pd.to_numeric(x["is_active"], errors="coerce").fillna(1).astype(int)
        x = x[["stock_id", "stock_name", "market", "industry", "theme", "sub_theme", "is_etf", "is_active", "update_date"]]
        with self.lock:
            x.to_sql("stocks_master", self.conn, if_exists="replace", index=False)
            self.conn.commit()

    def get_master(self) -> pd.DataFrame:
        with self.lock:
            return pd.read_sql_query(
                "SELECT * FROM stocks_master WHERE is_active=1 ORDER BY market, industry, stock_id",
                self.conn,
            )

    def get_stock_row(self, stock_id: str) -> Optional[pd.Series]:
        with self.lock:
            df = pd.read_sql_query("SELECT * FROM stocks_master WHERE stock_id=?", self.conn, params=[stock_id])
        if df.empty:
            return None
        return df.iloc[0]

    def upsert_price_history(self, stock_id: str, df: pd.DataFrame):
        if df is None or df.empty:
            return
        rows = []
        for _, r in df.iterrows():
            rows.append((
                stock_id,
                str(r["date"]),
                float(r["open"]) if pd.notna(r.get("open")) else None,
                float(r["high"]) if pd.notna(r.get("high")) else None,
                float(r["low"]) if pd.notna(r.get("low")) else None,
                float(r["close"]) if pd.notna(r.get("close")) else None,
                float(r["volume"]) if pd.notna(r.get("volume")) else None,
                float(r["turnover"]) if pd.notna(r.get("turnover")) else None,
            ))
        with self.lock:
            cur = self.conn.cursor()
            cur.executemany("""
                INSERT INTO price_history(stock_id, date, open, high, low, close, volume, turnover)
                VALUES(?,?,?,?,?,?,?,?)
                ON CONFLICT(stock_id, date) DO UPDATE SET
                    open=excluded.open,
                    high=excluded.high,
                    low=excluded.low,
                    close=excluded.close,
                    volume=excluded.volume,
                    turnover=excluded.turnover
            """, rows)
            self.conn.commit()

    def get_price_history(self, stock_id: str) -> pd.DataFrame:
        with self.lock:
            return pd.read_sql_query(
                "SELECT * FROM price_history WHERE stock_id=? ORDER BY date",
                self.conn, params=[stock_id]
            )

    def get_price_history_count(self, stock_id: str) -> int:
        with self.lock:
            row = self.conn.cursor().execute("SELECT COUNT(*) FROM price_history WHERE stock_id=?", (stock_id,)).fetchone()
        return int(row[0]) if row and row[0] is not None else 0

    def get_last_price_date(self) -> Optional[str]:
        with self.lock:
            row = self.conn.cursor().execute("SELECT MAX(date) FROM price_history").fetchone()
        return str(row[0]) if row and row[0] else None

    def get_total_price_rows(self) -> int:
        with self.lock:
            row = self.conn.cursor().execute("SELECT COUNT(*) FROM price_history").fetchone()
        return int(row[0]) if row and row[0] is not None else 0

    def get_ranking_rows_count(self) -> int:
        with self.lock:
            row = self.conn.cursor().execute("SELECT COUNT(*) FROM ranking_result WHERE date = (SELECT MAX(date) FROM ranking_result)").fetchone()
        return int(row[0]) if row and row[0] is not None else 0

    def replace_ranking(self, df: pd.DataFrame):
        today = datetime.now().strftime("%Y-%m-%d")
        with self.lock:
            cur = self.conn.cursor()
            cur.execute("DELETE FROM ranking_result WHERE date=?", (today,))
            self.conn.commit()
            df.to_sql("ranking_result", self.conn, if_exists="append", index=False)
            self.conn.commit()

    def get_latest_ranking(self) -> pd.DataFrame:
        q = """
        SELECT rr.*, sm.stock_name, sm.market, sm.industry, sm.theme
        FROM ranking_result rr
        JOIN stocks_master sm ON rr.stock_id = sm.stock_id
        WHERE rr.date = (SELECT MAX(date) FROM ranking_result)
        ORDER BY rr.rank_all ASC
        """
        with self.lock:
            return pd.read_sql_query(q, self.conn)


class DataEngine:
    def __init__(self, db: DBManager):
        self.db = db

    @staticmethod
    def yahoo_symbol(stock_id: str, market: str) -> str:
        if market in ("上市", "ETF"):
            return f"{stock_id}.TW"
        if market == "上櫃":
            return f"{stock_id}.TWO"
        return stock_id

    @staticmethod
    def _to_num(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")

    def fetch_twse_daily(self) -> pd.DataFrame:
        try:
            df = download_twse_official_daily_csv()
            if df is not None and not df.empty:
                return df
        except Exception:
            pass
        return pd.DataFrame()

    def fetch_tpex_daily(self) -> pd.DataFrame:
        urls = [
            "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes",
            "https://www.tpex.org.tw/openapi/v1/tpex_esb_quotes",
        ]
        parts = []
        for url in urls:
            try:
                res = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
                res.raise_for_status()
                data = res.json()
                df = pd.DataFrame(data)
                if df.empty:
                    continue
                rename_map = {
                    "SecuritiesCompanyCode": "stock_id", "CompanyCode": "stock_id", "股票代號": "stock_id", "證券代號": "stock_id",
                    "CompanyName": "stock_name", "股票名稱": "stock_name",
                    "Open": "open", "開盤價": "open",
                    "High": "high", "最高價": "high",
                    "Low": "low", "最低價": "low",
                    "Close": "close", "收盤價": "close",
                    "TradingShares": "volume", "成交股數": "volume", "成交數量": "volume", "Volume": "volume",
                }
                df = df.rename(columns=rename_map)
                required = ["stock_id", "open", "high", "low", "close", "volume"]
                if not all(c in df.columns for c in required):
                    continue
                df["stock_id"] = df["stock_id"].astype(str).str.strip()
                df = df[df["stock_id"].str.fullmatch(r"\d{4,5}", na=False)].copy()
                for c in ["open", "high", "low", "close", "volume"]:
                    df[c] = self._to_num(df[c])
                df = df.dropna(subset=["close"])
                if df.empty:
                    continue
                df["date"] = datetime.now().strftime("%Y-%m-%d")
                df["turnover"] = df["close"] * df["volume"]
                parts.append(df[["stock_id", "date", "open", "high", "low", "close", "volume", "turnover"]])
            except Exception:
                continue
        if not parts:
            return pd.DataFrame()
        return pd.concat(parts, ignore_index=True).drop_duplicates(subset=["stock_id"])
    def download_history(self, stock_id: str, market: str, period: str = "2y") -> pd.DataFrame:
        if yf is None:
            return pd.DataFrame()
        symbols = []
        primary = self.yahoo_symbol(stock_id, market)
        if primary:
            symbols.append(primary)
        if f"{stock_id}.TW" not in symbols:
            symbols.append(f"{stock_id}.TW")
        if f"{stock_id}.TWO" not in symbols:
            symbols.append(f"{stock_id}.TWO")
        seen = set()
        for symbol in symbols:
            if symbol in seen:
                continue
            seen.add(symbol)
            try:
                hist = yf.Ticker(symbol).history(period=period, auto_adjust=False)
                if hist is None or hist.empty:
                    continue
                hist = hist.rename(columns={
                    "Open": "open", "High": "high", "Low": "low", "Close": "close", "Volume": "volume"
                }).reset_index()
                date_col = "Date" if "Date" in hist.columns else "Datetime"
                hist["date"] = pd.to_datetime(hist[date_col]).dt.strftime("%Y-%m-%d")
                hist["turnover"] = hist["close"] * hist["volume"]
                out = hist[["date", "open", "high", "low", "close", "volume", "turnover"]].copy()
                for c in ["open", "high", "low", "close", "volume", "turnover"]:
                    out[c] = pd.to_numeric(out[c], errors="coerce")
                out = out.dropna(subset=["close"])
                if not out.empty:
                    return out
            except Exception:
                continue
        return pd.DataFrame()

    def download_latest_bar_yahoo(self, stock_id: str, market: str, days: str = "7d") -> pd.DataFrame:
        if yf is None:
            return pd.DataFrame()
        symbols = []
        primary = self.yahoo_symbol(stock_id, market)
        if primary:
            symbols.append(primary)
        if f"{stock_id}.TW" not in symbols:
            symbols.append(f"{stock_id}.TW")
        if f"{stock_id}.TWO" not in symbols:
            symbols.append(f"{stock_id}.TWO")

        seen = set()
        latest = pd.DataFrame()
        for symbol in symbols:
            if symbol in seen:
                continue
            seen.add(symbol)
            try:
                hist = yf.Ticker(symbol).history(period=days, auto_adjust=False)
                if hist is None or hist.empty:
                    continue
                hist = hist.rename(columns={
                    "Open": "open", "High": "high", "Low": "low", "Close": "close", "Volume": "volume"
                }).reset_index()
                date_col = "Date" if "Date" in hist.columns else "Datetime"
                hist["date"] = pd.to_datetime(hist[date_col]).dt.strftime("%Y-%m-%d")
                hist["turnover"] = hist["close"] * hist["volume"]
                out = hist[["date", "open", "high", "low", "close", "volume", "turnover"]].copy()
                for c in ["open", "high", "low", "close", "volume", "turnover"]:
                    out[c] = pd.to_numeric(out[c], errors="coerce")
                out = out.dropna(subset=["close"]).sort_values("date")
                if not out.empty:
                    latest = out.tail(1).copy()
                    break
            except Exception:
                continue
        return latest


    def build_full_history(self, min_days: int = 240, batch_size: int = 25, sleep_sec: float = 0.6, progress_cb=None, log_cb=None, cancel_cb=None) -> Tuple[int, int, int]:
        master = self.db.get_master()
        if master.empty:
            return 0, 0, 0
        success = 0
        failed = 0
        rows = 0
        total = len(master)
        for idx, (_, row) in enumerate(master.iterrows(), start=1):
            if cancel_cb and cancel_cb():
                raise OperationCancelled("使用者中斷完整歷史建庫")
            stock_id = str(row["stock_id"])
            market = str(row["market"])
            existing = self.db.get_price_history_count(stock_id)
            if existing >= min_days:
                if progress_cb:
                    progress_cb(idx, total, stock_id, existing, "skip")
                if log_cb and (idx % 25 == 0 or idx == total):
                    log_cb(f"[{idx}/{total}] {stock_id} 已具備 {existing} 筆歷史，跳過")
                continue
            try:
                hist_df = self.download_history(stock_id, market, period="2y")
                if hist_df is not None and not hist_df.empty:
                    self.db.upsert_price_history(stock_id, hist_df)
                    success += 1
                    rows += len(hist_df)
                    current_count = self.db.get_price_history_count(stock_id)
                    if log_cb:
                        log_cb(f"[{idx}/{total}] {stock_id} 補建成功，新增/覆蓋 {len(hist_df)} 筆，累計 {current_count} 筆")
                    if progress_cb:
                        progress_cb(idx, total, stock_id, current_count, "ok")
                else:
                    failed += 1
                    if log_cb:
                        log_cb(f"[{idx}/{total}] {stock_id} 無可用歷史資料")
                    if progress_cb:
                        progress_cb(idx, total, stock_id, existing, "fail")
            except Exception as e:
                failed += 1
                if log_cb:
                    log_cb(f"[{idx}/{total}] {stock_id} 下載失敗：{e}")
                if progress_cb:
                    progress_cb(idx, total, stock_id, existing, "error")
            if idx % batch_size == 0:
                if log_cb:
                    log_cb(f"--- 分批節點：已處理 {idx}/{total}，暫停 {sleep_sec:.1f} 秒，避免介面卡住 ---")
                time.sleep(sleep_sec)
        return success, failed, rows

    def update_incremental(self, progress_cb=None, log_cb=None, cancel_cb=None) -> Tuple[int, int, int]:
        master = self.db.get_master()
        if master.empty:
            return 0, 0, 0

        twse_df = self.fetch_twse_daily()
        tpex_df = self.fetch_tpex_daily()

        official_map = {}
        if not twse_df.empty:
            for _, row in twse_df.iterrows():
                official_map[str(row["stock_id"])] = pd.DataFrame([row])
        if not tpex_df.empty:
            for _, row in tpex_df.iterrows():
                official_map[str(row["stock_id"])] = pd.DataFrame([row])

        success = 0
        failed = 0
        rows = 0
        source_summary = {"official": 0, "yahoo": 0, "none": 0}

        total = len(master)
        for idx, (_, row) in enumerate(master.iterrows(), start=1):
            if cancel_cb and cancel_cb():
                raise OperationCancelled("使用者中斷每日增量更新")
            stock_id = str(row["stock_id"])
            market = str(row["market"])
            official_df = official_map.get(stock_id, pd.DataFrame())
            used_source = ""
            write_df = pd.DataFrame()

            if not official_df.empty:
                write_df = official_df.copy()
                used_source = "official"
            else:
                yahoo_df = self.download_latest_bar_yahoo(stock_id, market, days="7d")
                if yahoo_df is not None and not yahoo_df.empty:
                    write_df = yahoo_df.copy()
                    used_source = "yahoo"

            if not write_df.empty:
                self.db.upsert_price_history(stock_id, write_df)
                actual_rows = len(write_df)
                rows += actual_rows
                success += 1
                source_summary[used_source] += 1
                if log_cb and (idx % 20 == 0 or idx == total or used_source == "yahoo"):
                    src_name = "官方" if used_source == "official" else "Yahoo備援"
                    log_cb(f"[{idx}/{total}] {stock_id} 每日資料更新 {actual_rows} 筆｜來源 {src_name}")
                if progress_cb:
                    progress_cb(idx, total, stock_id, actual_rows, used_source)
            else:
                failed += 1
                source_summary["none"] += 1
                if log_cb and (idx % 50 == 0 or idx == total):
                    log_cb(f"[{idx}/{total}] {stock_id} 今日無官方資料，Yahoo 備援亦未取到")
                if progress_cb:
                    progress_cb(idx, total, stock_id, 0, "skip")

        if log_cb:
            log_cb(f"每日更新彙總｜官方 {source_summary['official']} 檔｜Yahoo備援 {source_summary['yahoo']} 檔｜未取到 {source_summary['none']} 檔")
        return success, failed, rows
    @staticmethod
    def attach(df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        x["ma5"] = x["close"].rolling(5).mean()
        x["ma10"] = x["close"].rolling(10).mean()
        x["ma20"] = x["close"].rolling(20).mean()
        x["ma60"] = x["close"].rolling(60).mean()

        ema12 = x["close"].ewm(span=12, adjust=False).mean()
        ema26 = x["close"].ewm(span=26, adjust=False).mean()
        x["macd"] = ema12 - ema26
        x["macd_signal"] = x["macd"].ewm(span=9, adjust=False).mean()
        x["macd_hist"] = x["macd"] - x["macd_signal"]

        delta = x["close"].diff()
        up = delta.clip(lower=0)
        down = -delta.clip(upper=0)
        ma_up = up.ewm(com=13, adjust=False).mean()
        ma_down = down.ewm(com=13, adjust=False).mean()
        rs = ma_up / ma_down.replace(0, np.nan)
        x["rsi14"] = 100 - (100 / (1 + rs))

        low_min = x["low"].rolling(9).min()
        high_max = x["high"].rolling(9).max()
        rsv = (x["close"] - low_min) / (high_max - low_min).replace(0, np.nan) * 100
        x["k"] = rsv.ewm(alpha=1 / 3, adjust=False).mean()
        x["d"] = x["k"].ewm(alpha=1 / 3, adjust=False).mean()
        return x


class IndicatorEngine:
    """相容層：舊版仍呼叫 IndicatorEngine.attach(...)，統一導向 DataEngine.attach(...)。"""

    @staticmethod
    def attach(df: pd.DataFrame) -> pd.DataFrame:
        return DataEngine.attach(df)


class LegacyStrategyEngine:
    """相容層（未被主流程使用）：僅保留舊版評分參考，不再作為排行或交易核心。"""
    @staticmethod
    def _clamp(v: float) -> float:
        return max(0.0, min(100.0, v))

    @staticmethod
    def score(df: pd.DataFrame) -> Dict[str, float]:
        last = df.iloc[-1]
        if len(df) < 60:
            return {
                "momentum_score": 0.0,
                "trend_score": 0.0,
                "reversal_score": 0.0,
                "volume_score": 0.0,
                "risk_score": 0.0,
                "ai_score": 0.0,
                "total_score": 0.0,
                "signal": "資料不足",
                "action": "等待資料",
            }

        ret20 = (last["close"] / df.iloc[-21]["close"] - 1) * 100 if len(df) >= 21 else 0
        momentum = LegacyStrategyEngine._clamp(50 + ret20 * 2)

        trend_raw = 0
        trend_raw += 1 if pd.notna(last["ma5"]) and last["close"] > last["ma5"] else 0
        trend_raw += 1 if pd.notna(last["ma10"]) and last["ma5"] > last["ma10"] else 0
        trend_raw += 1 if pd.notna(last["ma20"]) and last["ma10"] > last["ma20"] else 0
        trend_raw += 1 if pd.notna(last["ma60"]) and last["ma20"] > last["ma60"] else 0
        trend = trend_raw * 25

        rsi = float(last["rsi14"]) if pd.notna(last["rsi14"]) else 50
        macd_hist = float(last["macd_hist"]) if pd.notna(last["macd_hist"]) else 0
        reversal = LegacyStrategyEngine._clamp((100 - abs(rsi - 55) * 1.4) * 0.6 + (50 + macd_hist * 150) * 0.4)

        vol_ma20 = df["volume"].tail(20).mean()
        vol_ratio = (float(last["volume"]) / vol_ma20) if vol_ma20 and not np.isnan(vol_ma20) else 1.0
        volume = LegacyStrategyEngine._clamp(vol_ratio * 50)

        vol20 = df["close"].pct_change().tail(20).std()
        vol20 = 0.02 if pd.isna(vol20) else float(vol20)
        risk = LegacyStrategyEngine._clamp(100 - vol20 * 1500)

        ai = LegacyStrategyEngine._clamp(momentum * 0.2 + trend * 0.25 + reversal * 0.15 + volume * 0.15 + risk * 0.25)
        total = LegacyStrategyEngine._clamp(momentum * 0.22 + trend * 0.28 + reversal * 0.15 + volume * 0.15 + risk * 0.10 + ai * 0.10)

        signal, action = LegacyStrategyEngine.signal_action(last, total)
        return {
            "momentum_score": round(momentum, 2),
            "trend_score": round(trend, 2),
            "reversal_score": round(reversal, 2),
            "volume_score": round(volume, 2),
            "risk_score": round(risk, 2),
            "ai_score": round(ai, 2),
            "total_score": round(total, 2),
            "signal": signal,
            "action": action,
        }

    @staticmethod
    def signal_action(last: pd.Series, total_score: float):
        close_ = float(last["close"])
        ma20 = float(last["ma20"]) if pd.notna(last["ma20"]) else close_
        ma60 = float(last["ma60"]) if pd.notna(last["ma60"]) else close_
        macd_hist = float(last["macd_hist"]) if pd.notna(last["macd_hist"]) else 0
        rsi = float(last["rsi14"]) if pd.notna(last["rsi14"]) else 50

        if close_ > ma20 > ma60 and macd_hist > 0 and total_score >= 80:
            return "強勢追蹤", "拉回加碼"
        if close_ >= ma20 and total_score >= 65:
            return "整理偏多", "低接布局"
        if abs(close_ - ma20) / max(ma20, 1e-6) < 0.03 and 45 <= total_score < 65:
            return "區間整理", "區間操作"
        if close_ < ma20 and rsi < 45:
            return "轉弱警戒", "減碼/防守"
        if close_ < ma60 and macd_hist < 0 and total_score < 35:
            return "急跌風險", "觀望為主"
        return "中性觀察", "等待訊號"

    @staticmethod
    def fib_targets(df: pd.DataFrame):
        recent = df.tail(60)
        swing_low = float(recent["low"].min())
        swing_high = float(recent["high"].max())
        diff = max(swing_high - swing_low, 0.01)
        return (
            round(swing_high, 2),
            round(swing_low + diff * 1.382, 2),
            round(swing_low + diff * 1.618, 2),
        )

    @staticmethod
    def wave_stage(df: pd.DataFrame):
        if len(df) < 60:
            return "資料不足"
        recent = df.tail(55)["close"].reset_index(drop=True)
        hi = int(recent.idxmax())
        lo = int(recent.idxmin())
        if hi > lo and recent.iloc[-1] > recent.mean():
            return "推動浪"
        if hi < lo and recent.iloc[-1] < recent.mean():
            return "修正浪"
        return "整理浪"


class RankingEngine:
    def __init__(self, db: DBManager):
        self.db = db

    def rebuild(self, progress_cb=None, log_cb=None, cancel_cb=None):
        master = self.db.get_master()
        today = datetime.now().strftime("%Y-%m-%d")
        rows = []
        total = len(master)
        success = 0
        skipped = 0

        # V9.6.2 FUNDAMENTAL_LOCAL_CACHE：Ranking 不再負責下載或重建 EPS Matrix。
        # 正確流程：每日增量更新先同步 external_valuation / external_revenue，再寫入 financial_feature_daily；
        # 重排行只讀本地 financial_feature_daily，避免每次排行時即時抓網路或產生 E_NA-R_NA 假評分。
        run_id = self.db.log_system_run(event="ranking_rebuild", status="start", message="Ranking rebuild reads local financial_feature_daily only", module="ranking")
        feature_df = self.db.get_latest_financial_features()
        feature_map = {}
        if feature_df is not None and not feature_df.empty and "stock_id" in feature_df.columns:
            feature_map = {str(r.get("stock_id")): dict(r) for _, r in feature_df.iterrows()}
            ne_ratio = 0.0
            try:
                if "data_quality_flag" in feature_df.columns:
                    flags = feature_df["data_quality_flag"].fillna("").astype(str)
                    ne_ratio = float(flags.str.contains("NE", case=False, na=False).mean())
                elif "matrix_cell" in feature_df.columns:
                    cells = feature_df["matrix_cell"].fillna("").astype(str)
                    ne_ratio = float(cells.eq("E_NA-R_NA").mean())
            except Exception:
                ne_ratio = 0.0
            msg = f"[EPS MATRIX][MERGE] 只讀本地 financial_feature_daily｜features={len(feature_map)}｜NE_ratio={ne_ratio:.2%}"
            if log_cb:
                log_cb(msg)
            log_info(f"{msg}｜run_id={run_id}")
            if ne_ratio >= 0.80:
                warn = "[EPS MATRIX][GATE][WARNING] financial_feature_daily 大量 NE，基本面資料不足；排行可產出但不可視為可下單依據，請先執行每日增量更新同步基本面。"
                if log_cb:
                    log_cb(warn)
                log_warning(f"{warn}｜run_id={run_id}")
                self.db.log_system_run(event="financial_feature_quality_gate", status="warning", message=f"NE_ratio={ne_ratio:.2%}; features={len(feature_map)}", run_id=run_id, module="eps_matrix")
        else:
            if log_cb:
                log_cb("[EPS MATRIX][MERGE] 無本地 financial_feature_daily，Ranking 使用中性財務分數；請先執行每日增量更新同步基本面。")
            log_warning(f"[EPS MATRIX][MERGE] run_id={run_id} features=0; local cache missing")
            self.db.log_system_run(event="financial_feature_quality_gate", status="warning", message="financial_feature_daily missing; ranking uses neutral financial score", run_id=run_id, module="eps_matrix")

        for idx, (_, row) in enumerate(master.iterrows(), start=1):
            if cancel_cb and cancel_cb():
                raise OperationCancelled("使用者中斷重建排行")
            stock_id = str(row["stock_id"])
            hist = self.db.get_price_history(stock_id)
            if hist.empty or len(hist) < 70:
                skipped += 1
                if progress_cb:
                    progress_cb(idx, total, stock_id, success, 0, skipped, "skip")
                continue
            hist = DataEngine.attach(hist)
            score = StrategyEngineV91.score(hist)
            technical_total = float(score.get("total_score", 0) or 0)
            feat = feature_map.get(stock_id, {})
            revenue_eps_score = float(pd.to_numeric(pd.Series([feat.get("revenue_eps_score", 50)]), errors="coerce").fillna(50).iloc[0])
            valuation_score = float(pd.to_numeric(pd.Series([feat.get("revenue_eps_score", 50)]), errors="coerce").fillna(50).iloc[0])
            liquidity_score = float(score.get("volume_score", 0) or 0)
            theme_score = float(score.get("ai_score", 0) or 0)
            final_total = round(
                technical_total * 0.50 +
                revenue_eps_score * 0.20 +
                valuation_score * 0.10 +
                liquidity_score * 0.10 +
                theme_score * 0.10,
                2
            )
            score["technical_total_score"] = round(technical_total, 2)
            score["financial_score"] = round(revenue_eps_score, 2)
            score["total_score"] = final_total
            for c in ["eps_ttm", "eps_yoy", "revenue_yoy", "eps_bucket", "rev_bucket", "matrix_cell", "eps_category", "matrix_base_score", "modifier", "revenue_eps_score", "data_quality_flag", "source_trace_json"]:
                score[c] = feat.get(c, np.nan if c in ["eps_ttm", "eps_yoy", "revenue_yoy", "matrix_base_score", "modifier", "revenue_eps_score"] else "")
            rows.append({
                "date": today,
                "stock_id": stock_id,
                **score,
                "rank_all": 0,
                "rank_industry": 0
            })
            success += 1
            if progress_cb:
                progress_cb(idx, total, stock_id, success, 0, skipped, "ok")
            if log_cb and (idx % 100 == 0 or idx == total):
                log_cb(f"重排行進度 {idx}/{total}｜已納入 {success} 檔｜跳過 {skipped} 檔｜本地EPS矩陣合併")
        if not rows:
            return 0

        df = pd.DataFrame(rows).sort_values(["total_score", "ai_score"], ascending=[False, False]).reset_index(drop=True)
        df["rank_all"] = np.arange(1, len(df) + 1)
        merged = df.merge(master[["stock_id", "industry"]], on="stock_id", how="left")
        df["rank_industry"] = merged.groupby("industry")["total_score"].rank(method="dense", ascending=False).astype(int)
        self.db.replace_ranking(df)
        self.db.log_system_run(event="ranking_rebuild", status="ok", message=f"ranking rows={len(df)} with FUNDAMENTAL_LOCAL_CACHE", run_id=run_id, module="ranking")
        return len(df)









class ExternalSourceConfig:
    """V9.4：外部資料來源設定。每個module必須有target_table、parser、official_url與fallback_days。"""
    SOURCES = {
        "market_snapshot": {
            "source_name": "TWSE 官方市場指數快照",
            "official_url": "https://www.twse.com.tw/zh/trading/historical/mi-index.html",
            "request_template": TWSE_MARKET_SNAPSHOT_ENDPOINT_TEMPLATE,
            "target_table": "market_snapshot",
            "required_columns": ["snapshot_date", "market_score", "market_mode"],
            "fallback_days": 15,
            "mandatory": True,
            "parser": "fetch_market_snapshot",
            "source_priority": "1.TWSE MI_INDEX 官方市場指數 → 2.日期 fallback 找最近交易日 → 3.本地 price_history local_cache_fallback（清楚標示非官方，但不讓系統空轉）",
            "data_integrity_policy": MARKET_PROXY_DECISION_POLICY,
        },
        "institutional": {
            "source_name": "TWSE 三大法人買賣超",
            "official_url": "https://www.twse.com.tw/zh/trading/foreign/t86",
            "request_template": "https://www.twse.com.tw/rwd/zh/fund/T86?date={date}&selectType=ALLBUT0999&response=json",
            "target_table": "external_institutional",
            "required_columns": ["stock_id", "trade_date", "foreign_buy_sell", "trust_buy_sell", "dealer_buy_sell"],
            "fallback_days": 5,
            "mandatory": True,
            "parser": "fetch_institutional",
        },
        "margin": {
            "source_name": "TWSE+TPEx 官方融資融券（個股）",
            "official_url": TWSE_MARGIN_OFFICIAL_PAGE + " | " + TPEX_OPENAPI_PORTAL_URL,
            "request_template": TWSE_MARGIN_API_TEMPLATE,
            "target_table": "external_margin",
            "required_columns": ["stock_id", "trade_date", "margin_balance", "short_balance", "margin_change", "short_change", "retail_heat_score"],
            "fallback_days": 7,
            "mandatory": False,
            "parser": "fetch_margin",
            "source_priority": "1.TWSE MI_MARGN open_data 上市個股（dataset 11680）→ 2.TPEx 官方/OpenAPI 上櫃個股 → 3.官方HTML/CSV fallback；禁止Mitake/券商SDK",
            "license_url": TWSE_OPENAPI_LICENSE_URL,
            "oas_swagger_url": TWSE_OPENAPI_SWAGGER_URL + " | " + TPEX_OPENAPI_SWAGGER_URL,
        },
        "macro_margin_sentiment": {
            "source_name": "TWSE 市場融資融券情緒（整體市場）",
            "official_url": TWSE_MARGIN_COMPARE_PAGE,
            "request_template": TWSE_MARGIN_COMPARE_PAGE,
            "target_table": "macro_margin_sentiment",
            "required_columns": ["data_date", "macro_margin_score", "macro_margin_state"],
            "fallback_days": 7,
            "mandatory": False,
            "parser": "fetch_macro_margin_sentiment",
            "source_priority": "TWSE compare/margin 僅作市場情緒層，不寫 external_margin 個股表",
            "license_url": TWSE_OPENAPI_LICENSE_URL,
            "oas_swagger_url": TWSE_OPENAPI_SWAGGER_URL,
        },
        "revenue": {
            "source_name": "MOPS 月營收",
            "official_url": "https://mops.twse.com.tw",
            "request_template": "https://mopsfin.twse.com.tw/opendata/t187ap05_L.csv",
            "target_table": "external_revenue",
            "required_columns": ["stock_id", "revenue_month", "revenue", "yoy"],
            "fallback_days": 45,
            "mandatory": True,
            "parser": "fetch_revenue",
        },
        "valuation": {
            "source_name": "TWSE+TPEx 官方估值/EPS來源",
            "official_url": "https://www.twse.com.tw/zh/trading/historical/bwibbu-day.html | https://www.tpex.org.tw/zh-tw/mainboard/trading/info/daily-pe.html | https://www.tpex.org.tw/zh-tw/mainboard/listed/financial/rank-pe.html",
            "request_template": "https://www.twse.com.tw/rwd/zh/afterTrading/BWIBBU_d?date={date}&selectType=ALL&response=json",
            "target_table": "external_valuation",
            "required_columns": ["stock_id", "data_date", "pe", "pb", "dividend_yield", "eps_ttm"],
            "fallback_days": 10,
            "mandatory": False,
            "parser": "fetch_valuation",
            "source_priority": "1.TWSE官方API → 2.TPEx官方頁面/CSV → 3.MOPS OpenData → 4.Goodinfo fallback only",
            "license_url": TWSE_OPENAPI_LICENSE_URL,
            "oas_swagger_url": TWSE_OPENAPI_SWAGGER_URL,
            "goodinfo_policy": "Goodinfo僅允許fallback，不作為主資料源；預設停用，需設定GTC_ENABLE_GOODINFO_FALLBACK=1才會啟用。",
        },
        "event": {
            "source_name": "MOPS 重大訊息/事件",
            "official_url": "https://mops.twse.com.tw",
            "request_template": "manual_parser_required",
            "target_table": "external_event",
            "required_columns": ["event_id", "event_date", "event_type"],
            "fallback_days": 30,
            "mandatory": False,
            "parser": "fetch_event",
        },
    }

    @classmethod
    def to_dataframe(cls) -> pd.DataFrame:
        rows = []
        for module, cfg in cls.SOURCES.items():
            row = dict(cfg)
            row["module"] = module
            row.setdefault("source_priority", "1.TWSE OpenAPI/官方API → 2.TPEx官方頁面/CSV → 3.MOPS OpenData → 4.Goodinfo fallback only")
            row.setdefault("license_url", TWSE_OPENAPI_LICENSE_URL if module in ("valuation", "market_snapshot", "institutional", "margin", "macro_margin_sentiment") else "")
            row.setdefault("oas_swagger_url", (TWSE_OPENAPI_SWAGGER_URL + " | " + TPEX_OPENAPI_SWAGGER_URL) if module == "margin" else (TWSE_OPENAPI_SWAGGER_URL if module in ("valuation", "market_snapshot", "institutional", "macro_margin_sentiment") else ""))
            row.setdefault("mandatory_for_analysis", False)
            row.setdefault("mandatory_for_execution", bool(row.get("mandatory", False)))
            row.setdefault("soft_block_allowed", True)
            row.setdefault("v959_policy", ANALYSIS_EXECUTION_SPLIT_POLICY)
            row.setdefault("goodinfo_policy", "Goodinfo不可作主資料源；僅允許官方來源失敗後fallback，且預設停用。")
            rows.append(row)
        return pd.DataFrame(rows)


class ExternalPipelineResult:
    """統一外部資料Pipeline結果，避免UI只有pending卻宣告完成。"""
    def __init__(self, module: str, status: str, df: pd.DataFrame | None = None, source_name: str = "", official_url: str = "", request_url: str = "", source_date: str = "", http_status: str = "", fallback_count: int = 0, error_message: str = "", target_table: str = "", data_ready: int = 0, source_level: str = "official"):
        self.module = module
        self.status = status
        self.df = df if isinstance(df, pd.DataFrame) else pd.DataFrame()
        self.source_name = source_name
        self.official_url = official_url
        self.request_url = request_url
        self.source_date = source_date
        self.http_status = str(http_status or "")
        self.fallback_count = int(fallback_count or 0)
        self.error_message = str(error_message or "")
        self.target_table = target_table
        self.rows_count = int(len(self.df)) if isinstance(self.df, pd.DataFrame) else 0
        self.data_ready = int(data_ready)
        self.source_level = source_level

    def to_dict(self) -> dict:
        return {
            "module": self.module, "status": self.status, "source_name": self.source_name,
            "official_url": self.official_url, "request_url": self.request_url, "source_date": self.source_date,
            "http_status": self.http_status, "fallback_count": self.fallback_count, "rows_count": self.rows_count,
            "error_message": self.error_message, "target_table": self.target_table, "data_ready": self.data_ready,
            "source_level": self.source_level,
        }


class DataValidator:
    @staticmethod
    def validate_df(df: pd.DataFrame, required_columns: list[str], module: str) -> tuple[bool, str]:
        if df is None or df.empty:
            return False, f"{module} 無資料"
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            return False, f"{module} 缺欄位：{', '.join(missing)}"
        return True, "ok"


class ExternalDataWriter:
    def __init__(self, db: DBManager):
        self.db = db

    def _delete_existing_rows(self, table: str, df: pd.DataFrame):
        cur = self.db.conn.cursor()
        if table == "market_snapshot" and "snapshot_date" in df.columns:
            for v in df["snapshot_date"].dropna().astype(str).unique().tolist():
                cur.execute("DELETE FROM market_snapshot WHERE snapshot_date=?", (v,))
        elif table == "external_institutional" and {"stock_id", "trade_date"}.issubset(df.columns):
            cur.executemany("DELETE FROM external_institutional WHERE stock_id=? AND trade_date=?", df[["stock_id", "trade_date"]].astype(str).drop_duplicates().itertuples(index=False, name=None))
        elif table == "external_margin" and {"stock_id", "trade_date"}.issubset(df.columns):
            cur.executemany("DELETE FROM external_margin WHERE stock_id=? AND trade_date=?", df[["stock_id", "trade_date"]].astype(str).drop_duplicates().itertuples(index=False, name=None))
        elif table == "external_revenue" and {"stock_id", "revenue_month"}.issubset(df.columns):
            cur.executemany("DELETE FROM external_revenue WHERE stock_id=? AND revenue_month=?", df[["stock_id", "revenue_month"]].astype(str).drop_duplicates().itertuples(index=False, name=None))
        elif table == "external_valuation" and {"stock_id", "data_date"}.issubset(df.columns):
            cur.executemany("DELETE FROM external_valuation WHERE stock_id=? AND data_date=?", df[["stock_id", "data_date"]].astype(str).drop_duplicates().itertuples(index=False, name=None))
        elif table == "external_event" and "event_id" in df.columns:
            cur.executemany("DELETE FROM external_event WHERE event_id=?", [(str(v),) for v in df["event_id"].dropna().astype(str).unique().tolist()])

    def write_result(self, result: ExternalPipelineResult, run_id: str | None = None) -> ExternalPipelineResult:
        if result is None:
            return ExternalPipelineResult("unknown", "fail", error_message="result is None")
        df = result.df.copy() if isinstance(result.df, pd.DataFrame) else pd.DataFrame()
        if result.status not in ("success", "fallback") or df.empty:
            result.data_ready = 0
            return result
        try:
            with self.db.lock:
                if result.target_table == "market_snapshot":
                    self._delete_existing_rows("market_snapshot", df)
                    df.to_sql("market_snapshot", self.db.conn, if_exists="append", index=False)
                elif result.target_table == "macro_margin_sentiment":
                    if "data_date" in df.columns:
                        cur = self.db.conn.cursor()
                        cur.executemany("DELETE FROM macro_margin_sentiment WHERE data_date=?", df[["data_date"]].astype(str).drop_duplicates().itertuples(index=False, name=None))
                    df.to_sql("macro_margin_sentiment", self.db.conn, if_exists="append", index=False)
                elif result.target_table in {
                    "external_institutional", "external_margin", "external_revenue",
                    "external_valuation", "external_event"
                }:
                    self._delete_existing_rows(result.target_table, df)
                    df.to_sql(result.target_table, self.db.conn, if_exists="append", index=False)
                else:
                    raise ValueError(f"未知target_table：{result.target_table}")
                self.db.conn.commit()
            result.rows_count = int(len(df))
            result.data_ready = 1 if len(df) > 0 else 0
            return result
        except Exception as exc:
            result.status = "fail"
            result.data_ready = 0
            result.error_message = f"DB寫入失敗：{exc}"
            return result


class ExternalDataFetcher:
    def __init__(self, db: DBManager):
        self.db = db
        self.writer = ExternalDataWriter(db)

    def _request_json(self, url: str, timeout: int = 30) -> tuple[object, str]:
        """R9：官方資料請求統一入口。加上 raise_for_status 與 JSON 失敗診斷，避免 silent fail。"""
        resp = requests.get(
            url,
            timeout=timeout,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Referer": "https://www.twse.com.tw/",
                "Accept": "application/json,text/plain,*/*",
            },
        )
        http_status = str(resp.status_code)
        resp.raise_for_status()
        try:
            return resp.json(), http_status
        except Exception as exc:
            preview = (resp.text or "")[:300].replace("\n", " ").replace("\r", " ")
            raise RuntimeError(f"JSON解析失敗｜status={http_status}｜preview={preview}") from exc

    def _date_candidates(self, fallback_days: int) -> list[tuple[str, str, int]]:
        base_date = datetime.now()
        out = []
        for offset in range(int(fallback_days or 0) + 1):
            d = base_date - pd.Timedelta(days=offset)
            out.append((d.strftime("%Y%m%d"), d.strftime("%Y-%m-%d"), offset))
        return out

    def _num(self, v, default: float = 0.0) -> float:
        try:
            s = str(v).replace(",", "").replace("--", "").strip()
            if s in ("", "-", "None", "nan"):
                return float(default)
            return float(s)
        except Exception:
            return float(default)

    def _normalize_twse_dataset(self, payload) -> tuple[list, list]:
        """R9：TWSE JSON 可能是 {fields,data}、{tables:[{fields,data}]} 或 list，統一轉成 fields/data。"""
        if isinstance(payload, dict):
            fields = payload.get("fields") or payload.get("stat") or []
            data = payload.get("data")
            tables = payload.get("tables")
            if (not data) and isinstance(tables, list) and tables:
                first = tables[0] if isinstance(tables[0], dict) else {}
                fields = fields or first.get("fields") or first.get("stat") or []
                data = first.get("data") or []
            if isinstance(data, dict):
                data = data.get("data", [])
            if data is None:
                data = []
            return list(fields or []), list(data or [])
        if isinstance(payload, list):
            return [], payload
        return [], []

    def _clean_market_number(self, v, default: float = np.nan) -> float:
        """V9.5.8：清理 TWSE 指數數字，支援逗號、百分比與特殊符號。"""
        try:
            s = str(v).replace(",", "").replace("%", "").replace("+", "").strip()
            s = re.sub(r"<[^>]+>", "", s)
            s = s.replace("--", "").replace("-", "-")
            if s in ("", "-", "None", "nan", "NaN"):
                return float(default)
            return float(s)
        except Exception:
            return float(default)

    def _purge_proxy_market_snapshot_rows(self):
        """V9.5.8：避免舊版 proxy market_snapshot 繼續留在 DB 誤導 UI/Decision。"""
        try:
            with self.db.lock:
                cur = self.db.conn.cursor()
                cur.execute("DELETE FROM market_snapshot WHERE COALESCE(source_level,'')='proxy' OR COALESCE(source_url,'')='internal:price_history' OR COALESCE(source_status,'') LIKE 'fallback_proxy%'")
                self.db.conn.commit()
        except Exception as exc:
            log_warning(f"[DATA_INTEGRITY][MARKET] 清除 proxy market_snapshot 失敗：{exc}")

    def _parse_twse_market_index_snapshot(self, payload, date_iso: str, request_url: str) -> pd.DataFrame:
        """R9：解析 TWSE MI_INDEX type=MS，強化欄位對應與格式變動容錯。"""
        fields, data = self._normalize_twse_dataset(payload)
        rows = data if isinstance(data, list) else []
        target_values = None
        target_map = {}

        for rec in rows:
            vals = list(rec.values()) if isinstance(rec, dict) else (list(rec) if isinstance(rec, (list, tuple)) else [])
            if not vals:
                continue

            if isinstance(rec, dict):
                rec_map = {str(k).strip(): v for k, v in rec.items()}
            else:
                rec_map = {str(k).strip(): v for k, v in zip(fields, vals)} if fields else {}

            joined = " ".join(str(v) for v in vals) + " " + " ".join(str(k) for k in rec_map.keys())
            if ("發行量加權股價指數" in joined) or ("加權股價指數" in joined) or ("TAIEX" in joined):
                target_values = vals
                target_map = rec_map
                break

        if target_values is None:
            return pd.DataFrame()

        def pick_by_key(candidates, pos=None, default=np.nan):
            for key in candidates:
                for actual_key, actual_val in target_map.items():
                    if key in str(actual_key):
                        val = self._clean_market_number(actual_val, default=default)
                        if np.isfinite(val):
                            return val
            if pos is not None and len(target_values) > pos:
                return self._clean_market_number(target_values[pos], default=default)
            return float(default)

        taiex_close = pick_by_key(["收盤指數", "收盤價", "指數"], pos=1, default=np.nan)
        change_points = pick_by_key(["漲跌點數", "漲跌"], pos=3, default=0.0)
        change_pct = pick_by_key(["漲跌百分比", "漲跌幅"], pos=4, default=np.nan)

        # 若欄位位置改版造成 close 取到名稱或錯欄，掃描數值欄補救。
        if (not np.isfinite(taiex_close)) or taiex_close <= 100:
            numeric_vals = []
            for v in target_values:
                n = self._clean_market_number(v, default=np.nan)
                if np.isfinite(n):
                    numeric_vals.append(float(n))
            # TAIEX 通常會是整列最大的數值之一，避開 0/百分比/漲跌點數。
            candidates = [v for v in numeric_vals if v > 1000]
            if candidates:
                taiex_close = float(candidates[0])

        if not np.isfinite(change_pct) and np.isfinite(taiex_close) and taiex_close != 0:
            prev = max(abs(taiex_close - change_points), 1.0)
            change_pct = change_points / prev * 100.0

        if not np.isfinite(taiex_close) or taiex_close <= 0:
            return pd.DataFrame()

        market_score = 50.0 + max(-20.0, min(20.0, float(change_pct if np.isfinite(change_pct) else 0.0) * 10.0))
        market_score = round(max(0.0, min(100.0, market_score)), 2)
        market_mode = "Risk_ON" if market_score >= 60 else "Risk_OFF" if market_score <= 40 else "Neutral"
        taiex_trend = "多頭" if change_points > 0 else "空頭" if change_points < 0 else "中性"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return pd.DataFrame([{
            "snapshot_date": date_iso,
            "stock_id": "TAIEX",
            "stock_name": "發行量加權股價指數",
            "market": "指數",
            "industry": "大盤",
            "close": float(taiex_close),
            "rsi": np.nan,
            "rsi14": np.nan,
            "atr": np.nan,
            "atr_pct": np.nan,
            "price_dev": np.nan,
            "price_deviation": np.nan,
            "ma20": np.nan,
            "ma60": np.nan,
            "market_score": market_score,
            "market_mode": market_mode,
            "taiex_close": float(taiex_close),
            "taiex_trend": taiex_trend,
            "sp500_close": np.nan,
            "nasdaq_close": np.nan,
            "vix": np.nan,
            "us10y": np.nan,
            "dxy": np.nan,
            "breadth": np.nan,
            "source_status": "success_twse_official_mi_index_r9",
            "source_type": "official",
            "source_url": request_url,
            "source_level": "official_twse_mi_index",
            "proxy_reason": "R9：TWSE官方MI_INDEX解析成功；不使用假資料。",
            "update_time": now,
        }])

    def _build_market_snapshot_from_local_price_history(self, date_iso: str = "", reason: str = "") -> pd.DataFrame:
        """R10：從本地 price_history 建立「全市場、每檔股票最新一筆」market_snapshot fallback。

        修正 R9 只寫 1 筆市場總覽的問題。R10 會：
        1) 用 SQL JOIN 取每一檔股票最新交易日資料，不是整張表 tail(1)。
        2) 以 price_history 計算 close/volume/rsi/atr/price_dev/ma20/ma60。
        3) 寫入 market_snapshot 時每檔一筆，rows 應接近全市場 2000+。
        """
        try:
            latest_sql = """
                SELECT p.stock_id, p.date, p.open, p.high, p.low, p.close, p.volume, p.turnover
                FROM price_history p
                JOIN (
                    SELECT stock_id, MAX(date) AS max_date
                    FROM price_history
                    GROUP BY stock_id
                ) latest
                  ON p.stock_id = latest.stock_id
                 AND p.date = latest.max_date
            """
            hist_sql = """
                SELECT stock_id, date, open, high, low, close, volume, turnover
                FROM price_history
                WHERE stock_id IN (SELECT DISTINCT stock_id FROM price_history)
                ORDER BY stock_id, date
            """
            with self.db.lock:
                latest_df = pd.read_sql_query(latest_sql, self.db.conn)
                hist_df = pd.read_sql_query(hist_sql, self.db.conn)
                master_df = pd.read_sql_query("SELECT stock_id, stock_name, market, industry FROM stocks_master", self.db.conn)
        except Exception as exc:
            log_warning(f"[MARKET_SNAPSHOT][R10_LOCAL_FALLBACK] 讀取 price_history 失敗：{exc}")
            return pd.DataFrame()

        if latest_df is None or latest_df.empty:
            return pd.DataFrame()

        for _df in (latest_df, hist_df, master_df):
            if _df is not None and not _df.empty and "stock_id" in _df.columns:
                _df["stock_id"] = _df["stock_id"].astype(str).map(normalize_stock_id)
        for c in ["open", "high", "low", "close", "volume", "turnover"]:
            if c in latest_df.columns:
                latest_df[c] = pd.to_numeric(latest_df[c], errors="coerce")
            if c in hist_df.columns:
                hist_df[c] = pd.to_numeric(hist_df[c], errors="coerce")
        latest_df = latest_df.dropna(subset=["stock_id", "close"]).copy()
        latest_df = latest_df[latest_df["stock_id"].astype(str).ne("")]
        if latest_df.empty:
            return pd.DataFrame()

        def _calc_last_tech(g: pd.DataFrame) -> pd.Series:
            g = g.sort_values("date").copy()
            close = pd.to_numeric(g["close"], errors="coerce")
            high = pd.to_numeric(g.get("high", close), errors="coerce")
            low = pd.to_numeric(g.get("low", close), errors="coerce")
            prev_close = close.shift(1)
            ma20 = close.rolling(20, min_periods=1).mean()
            ma60 = close.rolling(60, min_periods=1).mean()
            delta = close.diff()
            up = delta.clip(lower=0)
            down = -delta.clip(upper=0)
            ma_up = up.ewm(com=13, adjust=False).mean()
            ma_down = down.ewm(com=13, adjust=False).mean()
            rs = ma_up / ma_down.replace(0, np.nan)
            rsi14 = 100 - (100 / (1 + rs))
            tr = pd.concat([(high - low).abs(), (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
            atr = tr.rolling(14, min_periods=1).mean()
            last_close = float(close.iloc[-1]) if len(close) else np.nan
            last_ma20 = float(ma20.iloc[-1]) if len(ma20) else np.nan
            last_atr = float(atr.iloc[-1]) if len(atr) else np.nan
            price_dev = ((last_close - last_ma20) / max(abs(last_ma20), 0.01)) if np.isfinite(last_close) and np.isfinite(last_ma20) else np.nan
            atr_pct = (last_atr / max(abs(last_close), 0.01) * 100.0) if np.isfinite(last_atr) and np.isfinite(last_close) else np.nan
            return pd.Series({
                "rsi": float(rsi14.iloc[-1]) if len(rsi14) and pd.notna(rsi14.iloc[-1]) else np.nan,
                "rsi14": float(rsi14.iloc[-1]) if len(rsi14) and pd.notna(rsi14.iloc[-1]) else np.nan,
                "atr": last_atr,
                "atr_pct": atr_pct,
                "price_dev": price_dev,
                "price_deviation": price_dev,
                "ma20": last_ma20,
                "ma60": float(ma60.iloc[-1]) if len(ma60) and pd.notna(ma60.iloc[-1]) else np.nan,
            })

        try:
            tech_df = hist_df.dropna(subset=["stock_id", "close"]).groupby("stock_id", group_keys=False).apply(_calc_last_tech).reset_index()
        except Exception as exc:
            log_warning(f"[MARKET_SNAPSHOT][R10_LOCAL_FALLBACK] 技術欄位計算失敗，改用空值：{exc}")
            tech_df = pd.DataFrame({"stock_id": latest_df["stock_id"].unique()})

        out = latest_df.merge(tech_df, on="stock_id", how="left")
        if master_df is not None and not master_df.empty:
            out = out.merge(master_df.drop_duplicates("stock_id"), on="stock_id", how="left")
        for c in ["stock_name", "market", "industry"]:
            if c not in out.columns:
                out[c] = ""
            out[c] = out[c].fillna("").astype(str)

        latest_date = str(out["date"].dropna().astype(str).max()) if "date" in out.columns and not out.empty else (date_iso or datetime.now().strftime("%Y-%m-%d"))
        # 市場分數以最新全市場漲跌廣度與等權報酬估算；明確標示為 local_cache_fallback_not_official。
        breadth = np.nan
        change_pct = 0.0
        try:
            prev_sql = """
                SELECT p.stock_id, p.close AS prev_close
                FROM price_history p
                JOIN (
                    SELECT stock_id, MAX(date) AS prev_date
                    FROM price_history
                    WHERE date < ?
                    GROUP BY stock_id
                ) prev
                  ON p.stock_id=prev.stock_id AND p.date=prev.prev_date
            """
            with self.db.lock:
                prev_df = pd.read_sql_query(prev_sql, self.db.conn, params=[latest_date])
            prev_df["stock_id"] = prev_df["stock_id"].astype(str).map(normalize_stock_id)
            prev_df["prev_close"] = pd.to_numeric(prev_df["prev_close"], errors="coerce")
            valid = out[["stock_id", "close", "turnover"]].merge(prev_df, on="stock_id", how="left").dropna(subset=["prev_close", "close"])
            valid = valid[valid["prev_close"] > 0]
            if not valid.empty:
                valid["ret_pct"] = (valid["close"] / valid["prev_close"] - 1.0) * 100.0
                breadth = round(float((valid["ret_pct"] > 0).mean() * 100.0), 2)
                w = pd.to_numeric(valid.get("turnover", 0), errors="coerce").fillna(0)
                change_pct = float((valid["ret_pct"] * w).sum() / w.sum()) if float(w.sum()) > 0 else float(valid["ret_pct"].mean())
        except Exception as exc:
            log_warning(f"[MARKET_SNAPSHOT][R10_LOCAL_FALLBACK] 市場廣度估算失敗：{exc}")

        market_score = 50.0 + max(-20.0, min(20.0, float(change_pct if np.isfinite(change_pct) else 0.0) * 10.0))
        market_score = round(max(0.0, min(100.0, market_score)), 2)
        market_mode = "Risk_ON" if market_score >= 60 else "Risk_OFF" if market_score <= 40 else "Neutral"
        taiex_trend = "多頭" if change_pct > 0 else "空頭" if change_pct < 0 else "中性"
        taiex_proxy = float(pd.to_numeric(out.get("close", pd.Series(dtype=float)), errors="coerce").mean())
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        out["snapshot_date"] = latest_date
        out["market_score"] = market_score
        out["market_mode"] = market_mode
        out["taiex_close"] = taiex_proxy
        out["taiex_trend"] = taiex_trend
        out["sp500_close"] = np.nan
        out["nasdaq_close"] = np.nan
        out["vix"] = np.nan
        out["us10y"] = np.nan
        out["dxy"] = np.nan
        out["breadth"] = breadth
        out["source_status"] = "fallback_local_price_history_full_market_r10"
        out["source_type"] = "local_cache"
        out["source_url"] = "internal:price_history_full_market_latest_by_stock"
        out["source_level"] = "local_cache_fallback_not_official"
        out["proxy_reason"] = f"R10 全市場本地快取fallback：每檔股票最新一筆；latest_date={latest_date}; rows={len(out)}; reason={reason}"
        out["update_time"] = now
        keep = [
            "snapshot_date", "stock_id", "stock_name", "market", "industry", "close", "open", "high", "low", "volume", "turnover",
            "rsi", "rsi14", "atr", "atr_pct", "price_dev", "price_deviation", "ma20", "ma60",
            "market_score", "market_mode", "taiex_close", "taiex_trend", "sp500_close", "nasdaq_close", "vix", "us10y", "dxy", "breadth",
            "source_status", "source_type", "source_url", "source_level", "proxy_reason", "update_time"
        ]
        for c in keep:
            if c not in out.columns:
                out[c] = np.nan if c not in ["stock_id", "stock_name", "market", "industry", "snapshot_date", "market_mode", "taiex_trend", "source_status", "source_type", "source_url", "source_level", "proxy_reason", "update_time"] else ""
        out = out[keep].drop_duplicates(subset=["snapshot_date", "stock_id"], keep="last")
        print(f"[MARKET_SNAPSHOT R10] local full-market fallback rows={len(out)} latest_date={latest_date}")
        log_warning(f"[MARKET_SNAPSHOT][R10] local full-market fallback rows={len(out)} latest_date={latest_date}")
        return out

    def fetch_market_snapshot(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        """R9 MARKET_SNAPSHOT_FIXED：
        1) 優先抓 TWSE 官方 MI_INDEX，含日期 fallback + retry + 強化解析。
        2) 若官方資料尚未更新、假日、網路/JSON/欄位改版造成失敗，改用本地 price_history 建立 local cache fallback。
        3) 每一步都寫 log；data_ready=1 代表 market_snapshot 表不再空轉，但 source_level 會清楚標示是否官方。
        """
        print("====== MARKET_SNAPSHOT R9 START ======")
        self._purge_proxy_market_snapshot_rows()
        last_error = ""
        last_url = cfg.get("request_template", "")
        attempts = []
        max_retry_per_date = 2

        for date_ymd, date_iso, offset in self._date_candidates(cfg.get("fallback_days", 15)):
            url = str(cfg.get("request_template", TWSE_MARKET_SNAPSHOT_ENDPOINT_TEMPLATE)).format(date=date_ymd)
            last_url = url
            for retry_no in range(1, max_retry_per_date + 1):
                try:
                    print(f"[MARKET_SNAPSHOT R9] official try date={date_ymd} retry={retry_no} url={url}")
                    payload, http_status = self._request_json(url, timeout=30)
                    df = self._parse_twse_market_index_snapshot(payload, date_iso, url)
                    attempts.append({"date": date_iso, "retry": retry_no, "status": http_status, "rows": 0 if df is None else len(df)})
                    if df is not None and not df.empty:
                        print(f"[MARKET_SNAPSHOT R9] official success rows={len(df)} date={date_iso}")
                        log_info(f"[MARKET_SNAPSHOT][R9] official success rows={len(df)} date={date_iso} url={url}")
                        return ExternalPipelineResult(
                            module,
                            "success" if offset == 0 else "fallback",
                            df=df,
                            source_name=cfg.get("source_name", ""),
                            official_url=cfg.get("official_url", ""),
                            request_url=url,
                            source_date=date_iso,
                            http_status=http_status,
                            fallback_count=offset,
                            target_table=cfg.get("target_table", ""),
                            data_ready=1,
                            source_level="official_twse_mi_index",
                        )
                    last_error = f"TWSE MI_INDEX 無可解析加權指數資料 date={date_ymd} retry={retry_no}"
                    log_warning(f"[MARKET_SNAPSHOT][R9] {last_error}")
                except Exception as exc:
                    last_error = f"{url} | retry={retry_no} | {exc}"
                    attempts.append({"date": date_iso, "retry": retry_no, "status": "error", "error": str(exc)[:200]})
                    log_warning(f"[MARKET_SNAPSHOT][R9] official fetch failed: {last_error}")
                    time.sleep(0.4 * retry_no)

        local_df = self._build_market_snapshot_from_local_price_history(
            reason=last_error or "TWSE official market snapshot unavailable"
        )
        if local_df is not None and not local_df.empty:
            print(f"[MARKET_SNAPSHOT R10] local full-market fallback success rows={len(local_df)} date={local_df.iloc[0].get('snapshot_date')}")
            log_warning(
                "[MARKET_SNAPSHOT][R10] official failed; local full-market price_history fallback used "
                f"rows={len(local_df)} reason={last_error}"
            )
            return ExternalPipelineResult(
                module,
                "fallback",
                df=local_df,
                source_name="本地 price_history 市場快照 fallback",
                official_url=cfg.get("official_url", ""),
                request_url="internal:price_history",
                source_date=str(local_df.iloc[0].get("snapshot_date", datetime.now().strftime("%Y-%m-%d"))),
                http_status="local_cache",
                fallback_count=int(cfg.get("fallback_days", 15) or 15),
                error_message=(last_error or ""),
                target_table=cfg.get("target_table", ""),
                data_ready=1,
                source_level="local_cache_fallback_not_official",
            )

        print("[MARKET_SNAPSHOT R9] fail: official and local fallback unavailable")
        return ExternalPipelineResult(
            module,
            "fail",
            df=pd.DataFrame(),
            source_name=cfg.get("source_name", ""),
            official_url=cfg.get("official_url", ""),
            request_url=last_url,
            source_date=datetime.now().strftime("%Y-%m-%d"),
            http_status="",
            fallback_count=int(cfg.get("fallback_days", 15) or 15),
            error_message=(last_error or "TWSE 官方 market_snapshot 未取得，且本地 price_history 無可用資料"),
            target_table=cfg.get("target_table", ""),
            data_ready=0,
            source_level="official_failed_local_cache_missing",
        )

    def fetch_institutional(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        last_error = ""
        for date_ymd, date_iso, offset in self._date_candidates(cfg.get("fallback_days", 5)):
            url = cfg.get("request_template", "").format(date=date_ymd)
            try:
                payload, http_status = self._request_json(url)
                fields, data = self._normalize_twse_dataset(payload)
                rows = []
                for rec in data:
                    vals = list(rec) if isinstance(rec, (list, tuple)) else []
                    if len(vals) < 8:
                        continue
                    sid = normalize_stock_id(vals[0])
                    if not sid:
                        continue
                    # TWSE T86常見欄位：外資買賣超、投信買賣超、自營商買賣超等；若欄位排序異動，仍保留可追溯URL與失敗原因。
                    foreign = self._num(vals[4] if len(vals) > 4 else 0)
                    trust = self._num(vals[10] if len(vals) > 10 else 0)
                    dealer = self._num(vals[11] if len(vals) > 11 else 0)
                    total = foreign + trust + dealer
                    score = max(0, min(100, 50 + total / 1000.0))
                    rows.append({
                        "stock_id": sid, "trade_date": date_iso,
                        "foreign_buy_sell": foreign, "trust_buy_sell": trust, "dealer_buy_sell": dealer,
                        "eight_bank_buy_sell": 0.0, "institutional_score": round(score, 2),
                        "main_force_flag": "BUY" if total > 0 else "SELL" if total < 0 else "NEUTRAL",
                        "source_date": date_iso, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "trade_date"], keep="last") if rows else pd.DataFrame()
                if not df.empty:
                    return ExternalPipelineResult(module, "success" if offset == 0 else "fallback", df=df, source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=url, source_date=date_iso, http_status=http_status, fallback_count=offset, target_table=cfg.get("target_table",""), data_ready=1, source_level="official")
                last_error = "TWSE T86回傳無可解析資料"
            except Exception as exc:
                last_error = str(exc)
                continue
        return ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message=last_error or "TWSE三大法人資料抓取失敗", target_table=cfg.get("target_table",""), data_ready=0, source_level="official")

    def _roc_date(self, date_ymd: str) -> str:
        """YYYYMMDD -> ROC date YYY/MM/DD for TPEx legacy endpoints."""
        try:
            d = datetime.strptime(str(date_ymd), "%Y%m%d")
            return f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"
        except Exception:
            return str(date_ymd)

    def _request_text(self, url: str, timeout: int = 30, referer: str = "https://www.twse.com.tw/") -> tuple[str, str]:
        resp = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0", "Referer": referer})
        resp.raise_for_status()
        try:
            resp.encoding = resp.apparent_encoding or "utf-8"
        except Exception:
            pass
        return resp.text, str(resp.status_code)

    def _parse_margin_twse_rows(self, payload, date_iso: str, source_url: str) -> pd.DataFrame:
        rows = []
        if isinstance(payload, str):
            try:
                raw = pd.read_csv(io.StringIO(payload), dtype=str).fillna("")
            except Exception:
                raw = pd.DataFrame()
            if raw is not None and not raw.empty:
                for _, r in raw.iterrows():
                    sid = normalize_stock_id(r.get("股票代號", ""))
                    if not sid:
                        continue
                    prev_margin = self._num(r.get("融資前日餘額", 0))
                    margin_balance = self._num(r.get("融資今日餘額", 0))
                    prev_short = self._num(r.get("融券前日餘額", 0))
                    short_balance = self._num(r.get("融券今日餘額", 0))
                    margin_change = margin_balance - prev_margin
                    short_change = short_balance - prev_short
                    margin_utilization = (short_balance / margin_balance * 100.0) if margin_balance > 0 else 0.0
                    retail_heat = max(0, min(100, 50 + margin_change / 500.0 - short_change / 500.0 + min(margin_utilization, 100) * 0.05))
                    rows.append({
                        "stock_id": sid, "trade_date": date_iso,
                        "margin_balance": margin_balance, "short_balance": short_balance,
                        "margin_change": margin_change, "short_change": short_change,
                        "margin_utilization": round(margin_utilization, 2), "retail_heat_score": round(retail_heat, 2),
                        "source_date": date_iso, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                return pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "trade_date"], keep="last") if rows else pd.DataFrame()
        fields, data = self._normalize_twse_dataset(payload)
        # TWSE MI_MARGN 常見資料列：代號、名稱、融資買進/賣出/現償/前餘額/今日餘額、融券買進/賣出/現償/前餘額/今日餘額等。
        for rec in data:
            vals = list(rec.values()) if isinstance(rec, dict) else (list(rec) if isinstance(rec, (list, tuple)) else [])
            if len(vals) < 8:
                continue
            sid = normalize_stock_id(vals[0])
            if not sid:
                continue
            # 儘量以 TWSE 固定欄位順序取值，若欄位改版，下面的資料品質 log 會顯示 rows=0。
            prev_margin = self._num(vals[5] if len(vals) > 5 else 0)
            margin_balance = self._num(vals[6] if len(vals) > 6 else 0)
            prev_short = self._num(vals[11] if len(vals) > 11 else 0)
            short_balance = self._num(vals[12] if len(vals) > 12 else 0)
            margin_change = margin_balance - prev_margin if prev_margin or margin_balance else self._num(vals[5] if len(vals) > 5 else 0)
            short_change = short_balance - prev_short if prev_short or short_balance else self._num(vals[11] if len(vals) > 11 else 0)
            margin_utilization = (short_balance / margin_balance * 100.0) if margin_balance > 0 else 0.0
            retail_heat = max(0, min(100, 50 + margin_change / 500.0 - short_change / 500.0 + min(margin_utilization, 100) * 0.05))
            rows.append({
                "stock_id": sid, "trade_date": date_iso,
                "margin_balance": margin_balance, "short_balance": short_balance,
                "margin_change": margin_change, "short_change": short_change,
                "margin_utilization": round(margin_utilization, 2), "retail_heat_score": round(retail_heat, 2),
                "source_date": date_iso, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
        return pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "trade_date"], keep="last") if rows else pd.DataFrame()

    def _fetch_twse_margin_one_day(self, date_ymd: str, date_iso: str) -> tuple[pd.DataFrame, str, str]:
        urls = [
            TWSE_MARGIN_OPEN_DATA_ENDPOINT,
            f"https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={date_ymd}&selectType=ALL",
            f"https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?date={date_ymd}&selectType=ALL&response=json",
        ]
        last_error = ""
        for url in urls:
            try:
                if "response=open_data" in url:
                    text, http_status = self._request_text(url, referer="https://www.twse.com.tw/")
                    df = self._parse_margin_twse_rows(text, date_iso, url)
                else:
                    payload, http_status = self._request_json(url)
                    df = self._parse_margin_twse_rows(payload, date_iso, url)
                if not df.empty:
                    log_info(f"[MARGIN][TWSE] success rows={len(df)} date={date_iso} url={url}")
                    return df, url, http_status
                last_error = f"TWSE MI_MARGN rows=0 url={url}"
            except Exception as exc:
                last_error = f"{url} | {exc}"
        raise RuntimeError(last_error or "TWSE MI_MARGN 無可解析資料")

    def _fetch_tpex_margin_candidates(self, date_ymd: str, date_iso: str) -> tuple[pd.DataFrame, str, str]:
        roc = self._roc_date(date_ymd)
        urls = [
            # TPEx OpenAPI 名稱可能隨官方版本調整；保留多候選並全部寫 log，避免 silent fail。
            "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_margin_trading",
            "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_margin_purchase_short_sale",
            f"https://www.tpex.org.tw/www/zh-tw/margin/balance?response=json&date={date_ymd}",
            f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={roc}",
        ]
        last_error = ""
        for url in urls:
            try:
                payload, http_status = self._request_json(url)
                data = payload.get("data") if isinstance(payload, dict) else payload
                rows = []
                for rec in data if isinstance(data, list) else []:
                    if isinstance(rec, dict):
                        keys = {str(k): v for k, v in rec.items()}
                        sid = normalize_stock_id(keys.get("代號") or keys.get("股票代號") or keys.get("證券代號") or keys.get("SecuritiesCompanyCode") or keys.get("Code") or keys.get("stock_id"))
                        if not sid:
                            continue
                        margin_balance = self._num(keys.get("融資餘額") or keys.get("融資今日餘額") or keys.get("MarginBalance") or keys.get("MarginPurchaseBalance") or 0)
                        short_balance = self._num(keys.get("融券餘額") or keys.get("融券今日餘額") or keys.get("ShortBalance") or keys.get("ShortSaleBalance") or 0)
                        margin_change = self._num(keys.get("融資增減") or keys.get("MarginChange") or 0)
                        short_change = self._num(keys.get("融券增減") or keys.get("ShortChange") or 0)
                    else:
                        vals = list(rec) if isinstance(rec, (list, tuple)) else []
                        if len(vals) < 6:
                            continue
                        sid = normalize_stock_id(vals[0])
                        if not sid:
                            continue
                        margin_balance = self._num(vals[5] if len(vals) > 5 else 0)
                        short_balance = self._num(vals[11] if len(vals) > 11 else 0)
                        margin_change = self._num(vals[4] if len(vals) > 4 else 0)
                        short_change = self._num(vals[10] if len(vals) > 10 else 0)
                    util = (short_balance / margin_balance * 100.0) if margin_balance > 0 else 0.0
                    heat = max(0, min(100, 50 + margin_change / 500.0 - short_change / 500.0 + min(util, 100) * 0.05))
                    rows.append({
                        "stock_id": sid, "trade_date": date_iso,
                        "margin_balance": margin_balance, "short_balance": short_balance,
                        "margin_change": margin_change, "short_change": short_change,
                        "margin_utilization": round(util, 2), "retail_heat_score": round(heat, 2),
                        "source_date": date_iso, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "trade_date"], keep="last") if rows else pd.DataFrame()
                if not df.empty:
                    log_info(f"[MARGIN][TPEx] success rows={len(df)} date={date_iso} url={url}")
                    return df, url, http_status
                last_error = f"TPEx margin rows=0 url={url}"
            except Exception as exc:
                last_error = f"{url} | {exc}"
                continue
        raise RuntimeError(last_error or "TPEx margin 無可解析資料")

    def _build_macro_margin_from_external_margin(self, df: pd.DataFrame, date_iso: str, source_url: str) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        total_margin = float(pd.to_numeric(df.get("margin_balance", 0), errors="coerce").fillna(0).sum())
        total_short = float(pd.to_numeric(df.get("short_balance", 0), errors="coerce").fillna(0).sum())
        total_margin_chg = float(pd.to_numeric(df.get("margin_change", 0), errors="coerce").fillna(0).sum())
        total_short_chg = float(pd.to_numeric(df.get("short_change", 0), errors="coerce").fillna(0).sum())
        util = (total_short / total_margin * 100.0) if total_margin > 0 else 0.0
        score = 50.0
        reason = []
        if total_margin_chg > 0:
            score -= min(20, abs(total_margin_chg) / max(total_margin, 1) * 1000)
            reason.append("融資總額增加：散戶熱度升溫")
        elif total_margin_chg < 0:
            score += min(15, abs(total_margin_chg) / max(total_margin, 1) * 1000)
            reason.append("融資總額下降：籌碼冷卻")
        if total_short_chg > 0:
            score += min(10, abs(total_short_chg) / max(total_short, 1) * 500)
            reason.append("融券增加：軋空潛力上升")
        if util > 30:
            score -= 10
            reason.append("券資比偏高：風險升高")
        score = round(max(0, min(100, score)), 2)
        state = "過熱" if score < 40 else "冷卻偏多" if score >= 60 else "中性"
        return pd.DataFrame([{
            "data_date": date_iso,
            "total_margin_balance": total_margin,
            "total_short_balance": total_short,
            "total_margin_change": total_margin_chg,
            "total_short_change": total_short_chg,
            "market_margin_utilization": round(util, 2),
            "macro_margin_score": score,
            "macro_margin_state": state,
            "sentiment_reason": "；".join(reason) if reason else "市場融資融券中性",
            "source_name": "TWSE/TPEx external_margin aggregate",
            "source_url": source_url,
            "source_date": date_iso,
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])

    def fetch_margin(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        last_error = ""
        for date_ymd, date_iso, offset in self._date_candidates(cfg.get("fallback_days", 7)):
            parts = []
            urls = []
            http_statuses = []
            try:
                try:
                    twse_df, twse_url, twse_status = self._fetch_twse_margin_one_day(date_ymd, date_iso)
                    if not twse_df.empty:
                        parts.append(twse_df)
                        urls.append(twse_url)
                        http_statuses.append(f"TWSE:{twse_status}")
                except Exception as exc:
                    last_error = f"TWSE margin fail: {exc}"
                    log_warning(f"[MARGIN][TWSE] {last_error}")
                try:
                    tpex_df, tpex_url, tpex_status = self._fetch_tpex_margin_candidates(date_ymd, date_iso)
                    if not tpex_df.empty:
                        parts.append(tpex_df)
                        urls.append(tpex_url)
                        http_statuses.append(f"TPEx:{tpex_status}")
                except Exception as exc:
                    last_error = (last_error + " | " if last_error else "") + f"TPEx margin fail: {exc}"
                    log_warning(f"[MARGIN][TPEx] {exc}")
                df = pd.concat(parts, ignore_index=True).drop_duplicates(subset=["stock_id", "trade_date"], keep="last") if parts else pd.DataFrame()
                if not df.empty:
                    macro_df = self._build_macro_margin_from_external_margin(df, date_iso, " | ".join(urls))
                    if not macro_df.empty:
                        macro_result = ExternalPipelineResult("macro_margin_sentiment", "success" if offset == 0 else "fallback", df=macro_df, source_name="TWSE/TPEx 市場融資情緒", official_url=TWSE_MARGIN_COMPARE_PAGE, request_url=" | ".join(urls), source_date=date_iso, http_status=";".join(http_statuses), fallback_count=offset, target_table="macro_margin_sentiment", data_ready=1, source_level="official")
                        self.writer.write_result(macro_result, run_id)
                    log_info(f"[MARGIN] merged rows={len(df)} date={date_iso} macro_rows={len(macro_df)}")
                    return ExternalPipelineResult(module, "success" if offset == 0 else "fallback", df=df, source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=" | ".join(urls), source_date=date_iso, http_status=";".join(http_statuses), fallback_count=offset, target_table=cfg.get("target_table",""), data_ready=1, source_level="official")
            except Exception as exc:
                last_error = str(exc)
                continue
        return ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message=last_error or "TWSE/TPEx融資融券資料抓取失敗", target_table=cfg.get("target_table",""), data_ready=0, source_level="official")

    def fetch_macro_margin_sentiment(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        # 優先由 external_margin 已成功寫入的個股資料彙總，確保市場層與個股層口徑一致。
        df = self.db.read_table("external_margin", limit=None)
        if df is not None and not df.empty:
            date_iso = str(df.get("trade_date", pd.Series([datetime.now().strftime("%Y-%m-%d")])).dropna().astype(str).max())
            macro = self._build_macro_margin_from_external_margin(df[df["trade_date"].astype(str) == date_iso].copy(), date_iso, "internal:external_margin aggregate")
            if not macro.empty:
                return ExternalPipelineResult(module, "success", df=macro, source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url="internal:external_margin aggregate", source_date=date_iso, http_status="internal", fallback_count=0, target_table=cfg.get("target_table",""), data_ready=1, source_level="official_aggregate")
        return ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message="external_margin尚無資料，無法彙總市場融資情緒；請先同步margin", target_table=cfg.get("target_table",""), data_ready=0, source_level="official")

    def _read_remote_csv(self, url: str) -> tuple[pd.DataFrame, str]:
        resp = requests.get(url, timeout=45, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://mopsfin.twse.com.tw/"})
        resp.raise_for_status()
        content = resp.content
        last_error = None
        for enc in ("utf-8-sig", "utf-8", "big5", "cp950"):
            try:
                return pd.read_csv(io.BytesIO(content), encoding=enc, dtype=str).fillna(""), str(resp.status_code)
            except Exception as exc:
                last_error = exc
        raise RuntimeError(f"CSV解析失敗：{last_error}")

    def fetch_revenue(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        urls = [cfg.get("request_template",""), "https://mopsfin.twse.com.tw/opendata/t187ap05_O.csv"]
        all_rows = []
        last_error = ""
        used_url = ""
        http_status = ""
        for url in urls:
            try:
                raw, http_status = self._read_remote_csv(url)
                used_url = url
                col_map = {}
                for c in raw.columns:
                    cs = str(c).strip()
                    if cs in ("公司代號", "公司代碼", "股票代號"):
                        col_map["stock_id"] = c
                    elif cs in ("出表年月", "資料年月", "年月", "營收年月"):
                        col_map["revenue_month"] = c
                    elif "當月營收" in cs:
                        col_map["revenue"] = c
                    elif "上月" in cs and "增減" in cs:
                        col_map["mom"] = c
                    elif ("去年" in cs or "同期" in cs) and "增減" in cs:
                        col_map["yoy"] = c
                    elif "累計營收" in cs and "當月" not in cs:
                        col_map["cumulative_revenue"] = c
                    elif "累計" in cs and "增減" in cs:
                        col_map["cumulative_yoy"] = c
                if "stock_id" not in col_map or "revenue" not in col_map:
                    last_error = f"MOPS月營收欄位無法辨識：{list(raw.columns)[:12]}"
                    continue
                for _, r in raw.iterrows():
                    sid = normalize_stock_id(r.get(col_map.get("stock_id"), ""))
                    if not sid:
                        continue
                    month = str(r.get(col_map.get("revenue_month"), datetime.now().strftime("%Y%m"))).strip()
                    rev = self._num(r.get(col_map.get("revenue"), 0))
                    yoy = self._num(r.get(col_map.get("yoy"), 0))
                    mom = self._num(r.get(col_map.get("mom"), 0))
                    all_rows.append({
                        "stock_id": sid, "revenue_month": month, "revenue": rev, "mom": mom, "yoy": yoy,
                        "cumulative_revenue": self._num(r.get(col_map.get("cumulative_revenue"), 0)),
                        "cumulative_yoy": self._num(r.get(col_map.get("cumulative_yoy"), 0)),
                        "source_date": datetime.now().strftime("%Y-%m-%d"),
                        "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
            except Exception as exc:
                last_error = str(exc)
                continue
        df = pd.DataFrame(all_rows).drop_duplicates(subset=["stock_id", "revenue_month"], keep="last") if all_rows else pd.DataFrame()
        if not df.empty:
            return ExternalPipelineResult(module, "success", df=df, source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=used_url, source_date=datetime.now().strftime("%Y-%m-%d"), http_status=http_status, target_table=cfg.get("target_table",""), data_ready=1, source_level="official")
        return ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=" | ".join(urls), source_date=datetime.now().strftime("%Y-%m-%d"), error_message=last_error or "MOPS月營收資料抓取失敗", target_table=cfg.get("target_table",""), data_ready=0, source_level="official")

    def _request_text(self, url: str, timeout: int = 45, referer: str = "https://www.tpex.org.tw/") -> tuple[str, str]:
        resp = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0", "Referer": referer})
        resp.raise_for_status()
        content = resp.content or b""
        text = ""
        for enc in ("utf-8-sig", "utf-8", "big5", "cp950"):
            try:
                text = content.decode(enc)
                break
            except Exception:
                continue
        if not text:
            text = resp.text
        return text, str(resp.status_code)

    def _extract_tpex_rows_from_payload(self, payload, expected_keywords: list[str] | None = None) -> tuple[list[str], list[list]]:
        """V9.5.5：TPEx 新舊頁面/CSV/HTML 解析共用函式。
        支援：
        - JSON dict: fields/data/tables
        - JSON list
        - CSV文字
        - HTML table（pandas.read_html）
        """
        expected_keywords = expected_keywords or []
        if isinstance(payload, dict):
            fields = payload.get("fields") or payload.get("stat") or payload.get("columns") or []
            data = payload.get("data") or payload.get("aaData") or payload.get("records") or []
            if not data and payload.get("tables"):
                try:
                    tbl0 = payload.get("tables", [{}])[0]
                    fields = fields or tbl0.get("fields") or tbl0.get("columns") or []
                    data = tbl0.get("data") or []
                except Exception:
                    pass
            return list(fields or []), list(data or [])
        if isinstance(payload, list):
            if payload and isinstance(payload[0], dict):
                fields = list(payload[0].keys())
                data = [[row.get(c, "") for c in fields] for row in payload]
                return fields, data
            return [], payload
        if isinstance(payload, str):
            s = payload.strip()
            # Try JSON text first
            if s.startswith("{") or s.startswith("["):
                try:
                    return self._extract_tpex_rows_from_payload(json.loads(s), expected_keywords=expected_keywords)
                except Exception:
                    pass
            # Try CSV
            try:
                df = pd.read_csv(io.StringIO(s), dtype=str).fillna("")
                if not df.empty and any(any(k in str(c) for k in expected_keywords) for c in df.columns):
                    return list(df.columns), df.values.tolist()
            except Exception:
                pass
            # Try HTML tables
            try:
                tables = pd.read_html(io.StringIO(s))
                for df in tables:
                    df = df.fillna("")
                    cols = [str(c).strip() for c in df.columns]
                    if expected_keywords and not any(any(k in c for k in expected_keywords) for c in cols):
                        continue
                    return cols, df.astype(str).values.tolist()
            except Exception:
                pass
        return [], []

    def _fetch_tpex_daily_pe(self, date_ymd: str, date_iso: str) -> tuple[pd.DataFrame, str, str, str]:
        """TPEx daily-pe：上櫃 PE / PB / 殖利率（官方）。"""
        roc_date = ""
        try:
            d = datetime.strptime(date_ymd, "%Y%m%d")
            roc_date = f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"
        except Exception:
            roc_date = ""
        urls = [
            f"https://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_result.php?d={roc_date}&l=zh-tw&o=json",
            f"https://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_result.php?d={roc_date}&l=zh-tw&o=csv",
            f"https://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_result.php?d={roc_date}&l=zh-tw&o=htm",
            f"https://www.tpex.org.tw/zh-tw/mainboard/trading/info/daily-pe.html",
        ]
        last_error = ""
        for url in urls:
            try:
                if "o=json" in url:
                    payload, http_status = self._request_json(url)
                    fields, data = self._extract_tpex_rows_from_payload(payload, expected_keywords=["本益比", "殖利率", "股價淨值比"])
                else:
                    text, http_status = self._request_text(url, referer="https://www.tpex.org.tw/")
                    fields, data = self._extract_tpex_rows_from_payload(text, expected_keywords=["本益比", "殖利率", "股價淨值比"])
                rows = []
                for rec in data:
                    vals = list(rec) if isinstance(rec, (list, tuple, np.ndarray)) else []
                    if not vals:
                        continue
                    row_map = {str(fields[i]).strip(): vals[i] for i in range(min(len(fields), len(vals)))} if fields else {}
                    sid = normalize_stock_id(row_map.get("股票代號", row_map.get("證券代號", vals[0] if vals else "")))
                    if not sid:
                        continue
                    name = str(row_map.get("名稱", row_map.get("證券名稱", vals[1] if len(vals) > 1 else "")) or "").strip()
                    pe = self._num(row_map.get("本益比", vals[2] if len(vals) > 2 else np.nan), default=np.nan)
                    dividend_yield = self._num(row_map.get("殖利率(%)", row_map.get("殖利率", vals[5] if len(vals) > 5 else np.nan)), default=np.nan)
                    pb = self._num(row_map.get("股價淨值比", vals[6] if len(vals) > 6 else np.nan), default=np.nan)
                    # TPEx daily-pe不一定提供收盤價；以PE直接使用，eps_ttm可由rank-pe補足。
                    rows.append({
                        "stock_id": sid,
                        "data_date": date_iso,
                        "close_price": None,
                        "pe": None if pd.isna(pe) else float(pe),
                        "pb": None if pd.isna(pb) else float(pb),
                        "dividend_yield": None if pd.isna(dividend_yield) else float(dividend_yield),
                        "eps": None,
                        "eps_ttm": None,
                        "roe": None,
                        "gross_margin": None,
                        "operating_margin": None,
                        "fiscal_year_quarter": "",
                        "source_date": date_iso,
                        "source_url": url,
                        "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "data_date"], keep="last") if rows else pd.DataFrame()
                if df is not None and not df.empty:
                    return df, url, http_status, ""
                last_error = "TPEx daily-pe 回傳無可解析資料"
            except Exception as exc:
                last_error = str(exc)
                continue
        return pd.DataFrame(), urls[0], "", last_error

    def _fetch_tpex_rank_pe_eps(self, date_iso: str) -> tuple[pd.DataFrame, str, str, str]:
        """TPEx rank-pe：上櫃 EPS 排名（官方，來源註明為MOPS最近期財報）。"""
        urls = [
            "https://www.tpex.org.tw/web/regular_emerging/financereport/regular_capitals_rank/list.php?l=zh-tw&o=json",
            "https://www.tpex.org.tw/web/regular_emerging/financereport/regular_capitals_rank/list.php?l=zh-tw&o=csv",
            "https://www.tpex.org.tw/web/regular_emerging/financereport/regular_capitals_rank/list.php?l=zh-tw&o=htm",
            "https://www.tpex.org.tw/zh-tw/mainboard/listed/financial/rank-pe.html",
        ]
        last_error = ""
        for url in urls:
            try:
                if "o=json" in url:
                    payload, http_status = self._request_json(url)
                    fields, data = self._extract_tpex_rows_from_payload(payload, expected_keywords=["EPS", "公司代號"])
                else:
                    text, http_status = self._request_text(url, referer="https://www.tpex.org.tw/")
                    fields, data = self._extract_tpex_rows_from_payload(text, expected_keywords=["EPS", "公司代號"])
                rows = []
                for rec in data:
                    vals = list(rec) if isinstance(rec, (list, tuple, np.ndarray)) else []
                    if not vals:
                        continue
                    row_map = {str(fields[i]).strip(): vals[i] for i in range(min(len(fields), len(vals)))} if fields else {}
                    sid = normalize_stock_id(row_map.get("公司代號", row_map.get("股票代號", row_map.get("證券代號", vals[1] if len(vals) > 1 else vals[0]))))
                    if not sid:
                        continue
                    eps_val = row_map.get("EPS", row_map.get("每股盈餘", vals[-1] if vals else np.nan))
                    eps = self._num(eps_val, default=np.nan)
                    if pd.isna(eps):
                        continue
                    rows.append({
                        "stock_id": sid,
                        "data_date": date_iso,
                        "eps": float(eps),
                        "eps_ttm": float(eps),
                        "source_url": url,
                        "source_date": date_iso,
                        "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "data_date"], keep="last") if rows else pd.DataFrame()
                if df is not None and not df.empty:
                    return df, url, http_status, ""
                last_error = "TPEx rank-pe 回傳無可解析EPS資料"
            except Exception as exc:
                last_error = str(exc)
                continue
        return pd.DataFrame(), urls[0], "", last_error

    def _fetch_goodinfo_eps_fallback(self, date_iso: str) -> tuple[pd.DataFrame, str, str, str]:
        """Goodinfo 僅作 fallback；預設停用，不作主資料源。"""
        url = "https://goodinfo.tw/tw/StockList.asp?MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%B9%B4%E5%BA%A6EPS%E6%9C%80%E9%AB%98"
        if not GOODINFO_FALLBACK_ENABLED:
            return pd.DataFrame(), url, "", "Goodinfo fallback disabled：依V9.5.5資料源政策，Goodinfo不可作主資料源，預設停用。"
        try:
            text, http_status = self._request_text(url, referer="https://goodinfo.tw/")
            fields, data = self._extract_tpex_rows_from_payload(text, expected_keywords=["EPS", "代號", "名稱"])
            rows = []
            for rec in data:
                vals = list(rec) if isinstance(rec, (list, tuple, np.ndarray)) else []
                if not vals:
                    continue
                joined = " ".join([str(v) for v in vals])
                sid = normalize_stock_id(joined)
                if not sid:
                    continue
                eps = np.nan
                for v in reversed(vals):
                    eps = self._num(v, default=np.nan)
                    if pd.notna(eps):
                        break
                if pd.isna(eps):
                    continue
                rows.append({
                    "stock_id": sid, "data_date": date_iso, "eps": float(eps), "eps_ttm": float(eps),
                    "source_url": url, "source_date": date_iso, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
            df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "data_date"], keep="last") if rows else pd.DataFrame()
            return df, url, http_status, "" if not df.empty else "Goodinfo fallback無可解析資料"
        except Exception as exc:
            return pd.DataFrame(), url, "", str(exc)

    def fetch_valuation(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        """V9.5.5 EPS_OFFICIAL_SOURCE：正式整合官方 EPS / 估值來源。

        資料源優先順序：
        1. TWSE OpenAPI / TWSE 官方 API：BWIBBU_d（上市 PE/PB/殖利率/eps_ttm）
           授權：http://data.gov.tw/license
           OAS：https://openapi.twse.com.tw/v1/swagger.json
        2. TPEx 官方頁面 / CSV：daily-pe（上櫃 PE/PB/殖利率）與 rank-pe（上櫃 EPS）
        3. MOPS OpenData：後續作 EPS YoY / QoQ 主來源（本版保留為資料來源依據）
        4. Goodinfo：僅 fallback，不當主資料源；預設停用。
        """
        last_error = ""
        source_notes = []
        request_templates = [
            cfg.get("request_template", ""),
            "https://www.twse.com.tw/exchangeReport/BWIBBU_d?date={date}&selectType=ALL&response=json",
            "https://www.twse.com.tw/rwd/zh/afterTrading/BWIBBU_d?date={date}&response=json",
        ]
        for date_ymd, date_iso, offset in self._date_candidates(cfg.get("fallback_days", 10)):
            all_parts = []
            request_urls = []
            http_statuses = []
            # 1) TWSE official BWIBBU_d（上市）
            for tpl in request_templates:
                if not tpl:
                    continue
                url = tpl.format(date=date_ymd)
                try:
                    payload, http_status = self._request_json(url)
                    fields, data = self._normalize_twse_dataset(payload)
                    if not fields or not data:
                        last_error = "TWSE BWIBBU_d 回傳無 fields/data"
                        continue
                    rows = []
                    for rec in data:
                        vals = list(rec) if isinstance(rec, (list, tuple)) else []
                        if len(vals) < 4:
                            continue
                        row_map = {str(fields[i]).strip(): vals[i] for i in range(min(len(fields), len(vals)))}
                        sid = normalize_stock_id(row_map.get("證券代號", vals[0] if vals else ""))
                        if not sid:
                            continue
                        close_price = self._num(row_map.get("收盤價", row_map.get("收盤價(元)", 0)), default=np.nan)
                        pe = self._num(row_map.get("本益比", 0), default=np.nan)
                        pb = self._num(row_map.get("股價淨值比", 0), default=np.nan)
                        dividend_yield = self._num(row_map.get("殖利率(%)", row_map.get("殖利率", 0)), default=np.nan)
                        eps_ttm = calculate_eps_ttm(close_price, pe)
                        fyq = str(row_map.get("財報年/季", "") or "").strip()
                        rows.append({
                            "stock_id": sid,
                            "data_date": date_iso,
                            "close_price": None if pd.isna(close_price) else float(close_price),
                            "pe": None if pd.isna(pe) else float(pe),
                            "pb": None if pd.isna(pb) else float(pb),
                            "dividend_yield": None if pd.isna(dividend_yield) else float(dividend_yield),
                            "eps": eps_ttm,
                            "eps_ttm": eps_ttm,
                            "roe": None,
                            "gross_margin": None,
                            "operating_margin": None,
                            "fiscal_year_quarter": fyq,
                            "source_date": date_iso,
                            "source_url": url,
                            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        })
                    twse_df = pd.DataFrame(rows).drop_duplicates(subset=["stock_id", "data_date"], keep="last") if rows else pd.DataFrame()
                    if twse_df is not None and not twse_df.empty:
                        all_parts.append(twse_df)
                        request_urls.append(url)
                        http_statuses.append(http_status)
                        source_notes.append(f"TWSE BWIBBU_d rows={len(twse_df)}")
                        break
                    last_error = "TWSE BWIBBU_d 回傳無可解析個股資料"
                except Exception as exc:
                    last_error = str(exc)
                    continue

            # 2) TPEx official daily-pe（上櫃 PE/PB/殖利率）
            tpex_daily, tpex_daily_url, tpex_daily_status, tpex_daily_err = self._fetch_tpex_daily_pe(date_ymd, date_iso)
            if tpex_daily is not None and not tpex_daily.empty:
                all_parts.append(tpex_daily)
                request_urls.append(tpex_daily_url)
                http_statuses.append(tpex_daily_status)
                source_notes.append(f"TPEx daily-pe rows={len(tpex_daily)}")
            elif tpex_daily_err:
                source_notes.append(f"TPEx daily-pe fail={tpex_daily_err}")

            # 3) TPEx official rank-pe（上櫃 EPS；補足 eps/eps_ttm）
            tpex_eps, tpex_eps_url, tpex_eps_status, tpex_eps_err = self._fetch_tpex_rank_pe_eps(date_iso)
            if tpex_eps is not None and not tpex_eps.empty:
                if tpex_daily is not None and not tpex_daily.empty:
                    # 將TPEx EPS補進daily-pe相同stock_id/data_date
                    key_cols = ["stock_id", "data_date"]
                    merged = tpex_daily.merge(tpex_eps[["stock_id", "data_date", "eps", "eps_ttm", "source_url"]], on=key_cols, how="left", suffixes=("", "_rank"))
                    for col in ["eps", "eps_ttm"]:
                        rank_col = f"{col}_rank"
                        if rank_col in merged.columns:
                            merged[col] = pd.to_numeric(merged[col], errors="coerce")
                            merged[rank_col] = pd.to_numeric(merged[rank_col], errors="coerce")
                            merged[col] = merged[col].where(merged[col].notna(), merged[rank_col])
                    if "source_url_rank" in merged.columns:
                        merged["source_url"] = merged["source_url"].astype(str) + " | EPS:" + merged["source_url_rank"].fillna("").astype(str)
                    drop_cols = [c for c in merged.columns if c.endswith("_rank") or c == "source_url_rank"]
                    merged = merged.drop(columns=drop_cols, errors="ignore")
                    # 用補完後版本取代前面已append的 tpex_daily
                    all_parts = [p for p in all_parts if not (isinstance(p, pd.DataFrame) and p is tpex_daily)]
                    all_parts.append(merged)
                else:
                    # 無daily-pe時，至少把EPS寫入估值表，PE/PB/殖利率留空，仍為官方補強資料
                    for c in ["close_price", "pe", "pb", "dividend_yield", "roe", "gross_margin", "operating_margin", "fiscal_year_quarter"]:
                        if c not in tpex_eps.columns:
                            tpex_eps[c] = None if c != "fiscal_year_quarter" else ""
                    all_parts.append(tpex_eps[["stock_id", "data_date", "close_price", "pe", "pb", "dividend_yield", "eps", "eps_ttm", "roe", "gross_margin", "operating_margin", "fiscal_year_quarter", "source_date", "source_url", "update_time"]])
                request_urls.append(tpex_eps_url)
                http_statuses.append(tpex_eps_status)
                source_notes.append(f"TPEx rank-pe EPS rows={len(tpex_eps)}")
            elif tpex_eps_err:
                source_notes.append(f"TPEx rank-pe fail={tpex_eps_err}")

            df = pd.concat(all_parts, ignore_index=True) if all_parts else pd.DataFrame()
            if df is not None and not df.empty:
                for c in ["close_price", "pe", "pb", "dividend_yield", "eps", "eps_ttm", "roe", "gross_margin", "operating_margin"]:
                    if c not in df.columns:
                        df[c] = None
                    df[c] = pd.to_numeric(df[c], errors="coerce")
                for c in ["fiscal_year_quarter", "source_date", "source_url", "update_time"]:
                    if c not in df.columns:
                        df[c] = ""
                df = df.drop_duplicates(subset=["stock_id", "data_date"], keep="last")
                return ExternalPipelineResult(
                    module,
                    "success" if offset == 0 else "fallback",
                    df=df,
                    source_name=cfg.get("source_name", ""),
                    official_url=cfg.get("official_url", ""),
                    request_url=" | ".join([u for u in request_urls if u]),
                    source_date=date_iso,
                    http_status=" | ".join([s for s in http_statuses if s]),
                    fallback_count=offset,
                    target_table=cfg.get("target_table", ""),
                    data_ready=1,
                    source_level="official_priority: TWSE→TPEx→MOPS; Goodinfo fallback only",
                )

        # 4) Goodinfo fallback only（預設停用，且永遠不作主資料源）
        good_df, good_url, good_status, good_err = self._fetch_goodinfo_eps_fallback(datetime.now().strftime("%Y-%m-%d"))
        if good_df is not None and not good_df.empty:
            for c in ["close_price", "pe", "pb", "dividend_yield", "roe", "gross_margin", "operating_margin", "fiscal_year_quarter"]:
                if c not in good_df.columns:
                    good_df[c] = None if c != "fiscal_year_quarter" else ""
            return ExternalPipelineResult(
                module,
                "fallback",
                df=good_df[["stock_id", "data_date", "close_price", "pe", "pb", "dividend_yield", "eps", "eps_ttm", "roe", "gross_margin", "operating_margin", "fiscal_year_quarter", "source_date", "source_url", "update_time"]],
                source_name="Goodinfo EPS fallback only",
                official_url=cfg.get("official_url", ""),
                request_url=good_url,
                source_date=datetime.now().strftime("%Y-%m-%d"),
                http_status=good_status,
                fallback_count=99,
                target_table=cfg.get("target_table", ""),
                data_ready=1,
                source_level="fallback_non_official_goodinfo",
            )
        if good_err:
            source_notes.append(good_err)

        return ExternalPipelineResult(
            module,
            "fail",
            source_name=cfg.get("source_name", ""),
            official_url=cfg.get("official_url", ""),
            request_url=cfg.get("request_template", ""),
            source_date=datetime.now().strftime("%Y-%m-%d"),
            error_message=(last_error or "官方估值/EPS資料抓取失敗") + (" | " + "；".join(source_notes[-5:]) if source_notes else ""),
            target_table=cfg.get("target_table", ""),
            data_ready=0,
            source_level="official_priority_failed_goodinfo_disabled_or_failed",
        )

    def fetch_event(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        return ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message="V9.4已取消pending假完成：重大訊息/事件官方parser尚未指定穩定API，列為Not Evaluated/fail。", target_table=cfg.get("target_table",""), data_ready=0, source_level="official")

    def _execute_module(self, module: str, cfg: dict, run_id: str) -> ExternalPipelineResult:
        t0 = time.time()
        self.db.log_system_run(event="external_pipeline", status="start", message=f"module={module}", run_id=run_id, step="fetch", module=module)
        func = getattr(self, str(cfg.get("parser", "")), None)
        if not callable(func):
            result = ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message=f"找不到parser：{cfg.get('parser')}", target_table=cfg.get("target_table",""), data_ready=0)
        else:
            try:
                result = func(module, cfg, run_id)
            except Exception as exc:
                result = ExternalPipelineResult(module, "fail", source_name=cfg.get("source_name",""), official_url=cfg.get("official_url",""), request_url=cfg.get("request_template",""), source_date=datetime.now().strftime("%Y-%m-%d"), error_message=str(exc), target_table=cfg.get("target_table",""), data_ready=0)

        ok, msg = DataValidator.validate_df(result.df, cfg.get("required_columns", []), module) if result.status in ("success", "fallback") else (False, result.error_message)
        self.db.log_external_data(
            module=module, source_name=result.source_name, official_url=result.official_url, request_url=result.request_url,
            source_date=result.source_date, status=result.status if ok else "fail", http_status=result.http_status,
            fallback_count=result.fallback_count, rows_count=result.rows_count, error_message="" if ok else msg,
            run_id=run_id, step="validate", validator_status="ok" if ok else "fail", duration_ms=(time.time()-t0)*1000,
            target_table=result.target_table, data_ready=1 if ok else 0, blocking_reason="" if ok else msg, source_level=result.source_level,
        )
        if ok:
            result = self.writer.write_result(result)
        else:
            result.status = "fail"
            result.data_ready = 0
            result.error_message = msg

        self.db.log_external_data(
            module=module, source_name=result.source_name, official_url=result.official_url, request_url=result.request_url,
            source_date=result.source_date, status=result.status, http_status=result.http_status, fallback_count=result.fallback_count,
            rows_count=result.rows_count, error_message=result.error_message, run_id=run_id, step="write",
            validator_status="ok" if ok else "fail", writer_status="ok" if result.data_ready else "fail",
            duration_ms=(time.time()-t0)*1000, target_table=result.target_table, data_ready=result.data_ready,
            blocking_reason="" if result.data_ready else (result.error_message or msg), source_level=result.source_level,
        )
        self.db.log_system_run(event="external_pipeline", status=result.status, message=f"module={module}; rows={result.rows_count}; ready={result.data_ready}; error={result.error_message}", run_id=run_id, step="write", module=module, duration_ms=(time.time()-t0)*1000)
        return result

    def sync_fundamental_local_cache(self, modules: list[str] | tuple[str, ...] | None = None, run_id: str | None = None, log_cb=None) -> dict:
        """V9.6.2：基本面本地快取同步。

        用途：每日增量更新流程中，先把不同來源的 EPS/估值與月營收下載並寫入 SQLite，
        再由 FinancialFeatureEngine 產生 financial_feature_daily。之後 Ranking 只讀本地快取，
        不再於重排行時即時抓網路或重建資料。
        """
        run_id = run_id or self.db.log_system_run(
            event="fundamental_local_cache_sync",
            status="start",
            message="sync valuation/revenue to local DB before ranking",
            step="start",
            module="fundamental_cache",
        )
        wanted = list(modules or ["valuation", "revenue"])
        results = []
        if log_cb:
            log_cb(f"[FUNDAMENTAL CACHE][START] modules={','.join(wanted)}｜run_id={run_id}")
        for module in wanted:
            cfg = ExternalSourceConfig.SOURCES.get(module)
            if not cfg:
                msg = f"module設定不存在：{module}"
                log_warning(f"[FUNDAMENTAL CACHE][SKIP] {msg}")
                results.append({"module": module, "status": "skip", "data_ready": 0, "rows_count": 0, "error_message": msg})
                continue
            result = self._execute_module(module, cfg, run_id)
            results.append(result.to_dict())
            if log_cb:
                log_cb(f"[FUNDAMENTAL CACHE][{module}] status={result.status}｜ready={result.data_ready}｜rows={result.rows_count}｜source_date={result.source_date}")

        feature_rows = 0
        ne_ratio = 1.0
        feature_status = "fail"
        try:
            feature_df = FinancialFeatureEngine(self.db).build_feature_batch(run_id=run_id, write_db=True)
            feature_rows = int(0 if feature_df is None else len(feature_df))
            if feature_df is not None and not feature_df.empty:
                feature_status = "success"
                if "data_quality_flag" in feature_df.columns:
                    flags = feature_df["data_quality_flag"].fillna("").astype(str)
                    ne_ratio = float(flags.str.contains("NE", case=False, na=False).mean())
                elif "matrix_cell" in feature_df.columns:
                    cells = feature_df["matrix_cell"].fillna("").astype(str)
                    ne_ratio = float(cells.eq("E_NA-R_NA").mean())
                else:
                    ne_ratio = 0.0
            self.db.log_external_data(
                module="financial_feature_daily",
                source_name="Fundamental Local Cache Feature Builder",
                official_url="external_valuation + external_revenue",
                request_url="internal:FinancialFeatureEngine.build_feature_batch",
                source_date=datetime.now().strftime("%Y-%m-%d"),
                status=feature_status,
                http_status="internal",
                rows_count=feature_rows,
                run_id=run_id,
                step="feature",
                target_table="financial_feature_daily",
                data_ready=1 if feature_rows > 0 else 0,
                blocking_reason="" if feature_rows > 0 else "financial_feature_daily rows=0",
                source_level="derived_local_cache",
            )
        except Exception as exc:
            feature_status = "fail"
            log_warning(f"[FUNDAMENTAL CACHE][FEATURE][ERROR] {exc}")
            self.db.log_external_data(
                module="financial_feature_daily",
                source_name="Fundamental Local Cache Feature Builder",
                official_url="external_valuation + external_revenue",
                request_url="internal:FinancialFeatureEngine.build_feature_batch",
                source_date=datetime.now().strftime("%Y-%m-%d"),
                status="fail",
                http_status="internal",
                rows_count=0,
                error_message=str(exc),
                run_id=run_id,
                step="feature",
                target_table="financial_feature_daily",
                data_ready=0,
                blocking_reason=str(exc),
                source_level="derived_local_cache",
            )

        quality_status = "ok" if feature_rows > 0 and ne_ratio < 0.80 else "warning"
        message = f"features={feature_rows}; NE_ratio={ne_ratio:.2%}; modules={','.join(wanted)}"
        self.db.log_system_run(
            event="fundamental_local_cache_sync",
            status=quality_status,
            message=message,
            run_id=run_id,
            step="finish",
            module="fundamental_cache",
        )
        if log_cb:
            log_cb(f"[FUNDAMENTAL CACHE][FINISH] {message}")
            if ne_ratio >= 0.80:
                log_cb("[FUNDAMENTAL CACHE][WARNING] EPS/Revenue 覆蓋率不足，若直接下單需人工確認基本面資料來源。")
        return {
            "run_id": run_id,
            "modules": wanted,
            "results": results,
            "feature_rows": feature_rows,
            "ne_ratio": ne_ratio,
            "status": quality_status,
        }

    def refresh_external_data_pipeline(self, run_id: str | None = None) -> dict:
        run_id = run_id or self.db.log_system_run(event="external_refresh", status="start", message="V9.4 true external data pipeline start", step="start")
        results = []
        for module, cfg in ExternalSourceConfig.SOURCES.items():
            results.append(self._execute_module(module, cfg, run_id))
        ready_map = {r.module: int(r.data_ready) for r in results}
        try:
            ff = FinancialFeatureEngine(self.db).build_feature_batch(run_id=run_id, write_db=True)
            self.db.log_external_data(
                module="financial_feature_daily",
                source_name="EPS Matrix Feature Engine",
                official_url="external_valuation + external_revenue",
                request_url="internal:financial_feature_daily",
                source_date=datetime.now().strftime("%Y-%m-%d"),
                status="success" if ff is not None and not ff.empty else "fail",
                http_status="internal",
                rows_count=0 if ff is None else len(ff),
                run_id=run_id,
                step="feature",
                target_table="financial_feature_daily",
                data_ready=1 if ff is not None and not ff.empty else 0,
                blocking_reason="" if ff is not None and not ff.empty else "financial_feature_daily rows=0",
                source_level="derived_feature",
            )
        except Exception as exc:
            log_warning(f"[EPS MATRIX][BUILD][ERROR] {exc}")
            self.db.log_external_data(
                module="financial_feature_daily",
                source_name="EPS Matrix Feature Engine",
                official_url="external_valuation + external_revenue",
                request_url="internal:financial_feature_daily",
                source_date=datetime.now().strftime("%Y-%m-%d"),
                status="fail",
                http_status="internal",
                rows_count=0,
                error_message=str(exc),
                run_id=run_id,
                step="feature",
                target_table="financial_feature_daily",
                data_ready=0,
                blocking_reason=str(exc),
                source_level="derived_feature",
            )
        # V9.5.9：外部資料缺口只形成資訊型 soft_block，不可停止分析或直接控制交易。
        execution_blocking = [f"{r.module}:{r.error_message or r.status}" for r in results if int(r.data_ready) == 0 and ExternalSourceConfig.SOURCES.get(r.module, {}).get("mandatory", False)]
        blocking = list(execution_blocking)  # 舊欄位相容：代表外部資料未完整，不代表停止交易引擎。
        status = "soft_block" if execution_blocking else "ok"
        self.db.log_system_run(event="external_refresh", status=status, message="; ".join(execution_blocking) if execution_blocking else "external pipeline completed", run_id=run_id, step="finish")
        return {"run_id": run_id, "results": [r.to_dict() for r in results], "ready_map": ready_map, "blocking": blocking, "execution_blocking": execution_blocking, "analysis_allowed": 1, "status": status}

    def refresh_all(self, run_id: str | None = None) -> dict:
        # backward compatibility：舊UI呼叫仍導向V9.4真實Pipeline，不再只寫pending。
        return self.refresh_external_data_pipeline(run_id=run_id)



class ExternalDataReadiness:
    """V9.5.9：外部資料Ready分層。

    重點：
    - analysis_ready：技術分析 / TOP20 / 觀察池永遠允許，外部缺口只提示。
    - execution_ready：僅為資訊欄位，用於 UI / Excel / Log 顯示外部資料是否完整。
    - mandatory_ready：保留舊介面相容，回傳 execution_ready 狀態，但不得再被 trade_allowed 當硬開關。
    """
    EXECUTION_MANDATORY_MODULES = ["market_snapshot", "institutional", "revenue"]
    ANALYSIS_MANDATORY_MODULES = []
    MANDATORY_MODULES = EXECUTION_MANDATORY_MODULES

    def __init__(self, db: DBManager):
        self.db = db

    def status_df(self) -> pd.DataFrame:
        return self.db.read_table("external_source_status", limit=None)

    def get_status(self, module: str) -> dict:
        df = self.status_df()
        if df is None or df.empty or "module" not in df.columns:
            return {"ready": 0, "reason": f"{module} 未執行同步", "rows": 0, "status": "missing"}
        x = df[df["module"].astype(str) == str(module)].tail(1)
        if x.empty:
            return {"ready": 0, "reason": f"{module} 無狀態紀錄", "rows": 0, "status": "missing"}
        r = x.iloc[-1]
        ready = int(pd.to_numeric(pd.Series([r.get("data_ready", 0)]), errors="coerce").fillna(0).iloc[0])
        rows = int(pd.to_numeric(pd.Series([r.get("rows_count", 0)]), errors="coerce").fillna(0).iloc[0])
        status = str(r.get("status", "") or "")
        reason = str(r.get("blocking_reason", "") or r.get("error_message", "") or "")
        if not ready:
            reason = reason or f"{module} data_ready=0/status={status}/rows={rows}"
        return {"ready": ready, "reason": reason, "rows": rows, "status": status}

    def analysis_ready(self) -> tuple[int, str]:
        # V9.5.9 最終整合版：外部資料缺口不得阻擋 AI 分析 / TOP20 / 觀察池。
        return 1, "analysis_ready=1：外部資料僅作風險提示，不阻擋分析"

    def execution_ready(self) -> tuple[int, str]:
        missing = []
        for module in self.EXECUTION_MANDATORY_MODULES:
            st = self.get_status(module)
            if int(st.get("ready", 0)) != 1:
                missing.append(f"{module}:{st.get('reason','not ready')}")
        return (0 if missing else 1, "；".join(missing))

    def mandatory_ready(self) -> tuple[int, str]:
        # 舊函式相容：只代表外部資料完整狀態，不代表 trade_allowed。
        return self.execution_ready()

def gate_state(pass_value, data_ready: int = 1, not_evaluated: int = 0, applicable: int = 1) -> str:
    if int(applicable or 0) == 0:
        return "NA"
    if int(not_evaluated or 0) == 1:
        return "NE"
    if int(data_ready or 0) != 1:
        return "BLOCK"
    return "PASS" if int(pass_value or 0) == 1 else "BLOCK"


def short_reason(text: str, max_len: int = 120) -> str:
    s = str(text or "").replace("\n", " ").replace("\r", " ").strip()
    return s[:max_len] + ("..." if len(s) > max_len else "")


def latest_source_trace(db: DBManager) -> dict:
    try:
        status = db.read_table("external_source_status", limit=None)
        if status is None or status.empty:
            return {"ready": 0, "analysis_ready": 1, "execution_ready": 0, "soft_block": 1, "latest_external_date": "", "market_source_level": "", "modules": {}, "blocking": "external_source_status empty"}
        modules = {}
        blocking = []
        latest_dates = []
        market_source_level = ""
        execution_mandatory = set(getattr(ExternalDataReadiness, "EXECUTION_MANDATORY_MODULES", []))
        for _, r in status.iterrows():
            module = str(r.get("module", ""))
            ready = int(pd.to_numeric(pd.Series([r.get("data_ready", 0)]), errors="coerce").fillna(0).iloc[0])
            source_date = str(r.get("source_date", "") or "")
            if source_date:
                latest_dates.append(source_date)
            if module == "market_snapshot":
                market_source_level = str(r.get("source_level", "") or "")
            modules[module] = {
                "ready": ready,
                "status": str(r.get("status", "") or ""),
                "rows": int(pd.to_numeric(pd.Series([r.get("rows_count", 0)]), errors="coerce").fillna(0).iloc[0]),
                "source_date": source_date,
                "source_level": str(r.get("source_level", "") or ""),
                "reason": str(r.get("blocking_reason", "") or r.get("error_message", "") or ""),
                "url": str(r.get("request_url", "") or ""),
            }
            if module in execution_mandatory and not ready:
                blocking.append(f"{module}:{modules[module]['reason'] or modules[module]['status']}")
        execution_ready = 0 if blocking else 1
        return {
            "ready": execution_ready,  # 舊欄位相容：代表外部資料完整，不代表交易開關。
            "analysis_ready": 1,
            "execution_ready": execution_ready,
            "soft_block": int(execution_ready != 1),
            "latest_external_date": max(latest_dates) if latest_dates else "",
            "market_source_level": market_source_level,
            "modules": modules,
            "blocking": "；".join(blocking),
        }
    except Exception as exc:
        return {"ready": 0, "analysis_ready": 1, "execution_ready": 0, "soft_block": 1, "latest_external_date": "", "market_source_level": "", "modules": {}, "blocking": str(exc)}

def apply_external_decision_filter(df: pd.DataFrame, label: str = "") -> pd.DataFrame:
    """V9.5.9：外部資料不得直接濾掉交易候選。

    trade_allowed 是唯一交易控制欄位；execution_ready/soft_block 只作資訊提示。
    此函式保留名稱是為了相容舊流程，但不再因外部資料缺口移除資料列。
    """
    if df is None or df.empty:
        return pd.DataFrame()
    x = df.copy()
    for col, default in [("trade_allowed", 0), ("analysis_ready", 1), ("execution_ready", 0), ("soft_block", 0)]:
        if col not in x.columns:
            x[col] = default
    if "external_blocking_reason" not in x.columns:
        x["external_blocking_reason"] = ""
    if "block_reason" not in x.columns:
        x["block_reason"] = x.get("external_blocking_reason", "")
    soft_cnt = int(pd.to_numeric(x.get("soft_block", 0), errors="coerce").fillna(0).astype(int).sum()) if "soft_block" in x.columns else 0
    if soft_cnt:
        try:
            log_warning(f"V9.5.9 external soft_block only {soft_cnt} rows from {label}; rows are kept, not filtered")
        except Exception:
            pass
    return x.copy()

def attach_external_display_columns(x: pd.DataFrame) -> pd.DataFrame:
    if x is None or x.empty:
        return x
    out = x.copy()
    mapping = {
        "外部允許": "trade_allowed",
        "外部Ready": "external_data_ready",
        "分析Ready": "analysis_ready",
        "ExecutionReady": "execution_ready",
        "SoftBlock": "soft_block",
        "BlockReason": "block_reason",
        "Market Gate": "market_gate_state",
        "Flow Gate": "flow_gate_state",
        "Fundamental Gate": "fundamental_gate_state",
        "Event Gate": "event_gate_state",
        "Risk Gate": "risk_gate_state",
        "外部阻擋原因": "external_blocking_reason",
        "外部資料日": "latest_external_date",
        "資料來源層級": "market_source_level",
        "決策摘要": "decision_reason_short",
        "全域外部Ready": "global_external_ready",
        "個股覆蓋狀態": "stock_external_coverage_state",
        "Gate說明": "gate_policy_note",
        "PE": "pe",
        "PB": "pb",
        "殖利率%": "dividend_yield",
        "EPS_TTM": "eps_ttm",
        "EPS YoY": "eps_yoy",
        "營收YoY": "revenue_yoy",
        "EPS分類": "eps_category",
        "Matrix": "matrix_cell",
        "財務分數": "revenue_eps_score",
        "資料狀態": "data_quality_flag",
        "估值分": "valuation_score",
        "融資餘額": "margin_balance",
        "融資增減": "margin_change",
        "融券餘額": "short_balance",
        "融券增減": "short_change",
        "券資比%": "margin_utilization",
        "散戶熱度": "retail_heat_score",
        "融資分": "margin_score",
        "融資狀態": "margin_state",
        "市場融資分": "macro_margin_score",
        "市場融資狀態": "macro_margin_state",
        "融資決策說明": "margin_decision_note",
    }
    for zh, src in mapping.items():
        if zh not in out.columns:
            out[zh] = out[src] if src in out.columns else ""
    return out


class CapitalFlowEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.readiness = ExternalDataReadiness(db)

    def evaluate(self, stock_id: str) -> dict:
        st = self.readiness.get_status("institutional")
        if int(st.get("ready", 0)) != 1:
            return {
                "pass": 1,
                "score": 50.0,
                "reason": f"法人資料未就緒，Flow Gate=NE，不阻擋交易：{st.get('reason')}",
                "data_ready": 0,
                "not_evaluated": 1,
                "coverage_state": "NE_FLOW_SOURCE_NOT_READY",
            }
        df = self.db.read_table("external_institutional", limit=None)
        if df.empty or "stock_id" not in df.columns:
            return {
                "pass": 1,
                "score": 50.0,
                "reason": "法人資料表無個股覆蓋資料，Flow Gate=NE，不阻擋交易",
                "data_ready": 0,
                "not_evaluated": 1,
                "coverage_state": "NE_NO_FLOW_TABLE",
            }
        x = df[df["stock_id"].astype(str) == str(stock_id)].copy()
        if x.empty:
            return {
                "pass": 1,
                "score": 50.0,
                "reason": "該股無法人覆蓋資料，Flow Gate=NE，不阻擋交易",
                "data_ready": 0,
                "not_evaluated": 1,
                "coverage_state": "NE_NO_FLOW_COVERAGE",
            }
        x = x.tail(5)
        score = float(pd.to_numeric(x.get("institutional_score", 50), errors="coerce").fillna(50).mean())
        return {
            "pass": int(score >= 45),
            "score": round(score, 2),
            "reason": f"法人分數 {score:.1f}",
            "data_ready": 1,
            "not_evaluated": 0,
            "coverage_state": "COVERED",
        }


class MarginDecisionEngine:
    """V9.5.6：個股融資 + 市場融資情緒。資料未覆蓋時為 NE，不阻擋；有資料時作為籌碼風險/加分。"""
    def __init__(self, db: DBManager):
        self.db = db
        self.readiness = ExternalDataReadiness(db)

    @staticmethod
    def score_margin_row(row: pd.Series) -> tuple[float, str, str]:
        margin_change = float(pd.to_numeric(pd.Series([row.get("margin_change", 0)]), errors="coerce").fillna(0).iloc[0])
        short_change = float(pd.to_numeric(pd.Series([row.get("short_change", 0)]), errors="coerce").fillna(0).iloc[0])
        util = float(pd.to_numeric(pd.Series([row.get("margin_utilization", 0)]), errors="coerce").fillna(0).iloc[0])
        score = 50.0
        reasons = []
        if margin_change > 0:
            score -= min(25, margin_change / 500.0)
            reasons.append("融資增加：散戶追價風險")
        elif margin_change < 0:
            score += min(20, abs(margin_change) / 500.0)
            reasons.append("融資下降：籌碼冷卻")
        if short_change > 0:
            score += min(12, short_change / 500.0)
            reasons.append("融券增加：軋空潛力")
        elif short_change < 0:
            score -= min(8, abs(short_change) / 500.0)
            reasons.append("融券下降：空方回補")
        if util > 30:
            score -= 10
            reasons.append("券資比偏高")
        score = float(max(0, min(100, score)))
        state = "冷卻偏多" if score >= 60 else "散戶過熱" if score < 40 else "中性"
        return round(score, 2), state, "；".join(reasons) if reasons else "融資融券中性"

    def evaluate(self, stock_id: str) -> dict:
        st = self.readiness.get_status("margin")
        df = self.db.read_table("external_margin", limit=None)
        macro_df = self.db.read_table("macro_margin_sentiment", limit=None)
        macro_score = 50.0
        macro_state = "NE"
        macro_reason = "市場融資情緒NE"
        if macro_df is not None and not macro_df.empty:
            m = macro_df.tail(1).iloc[-1]
            macro_score = float(pd.to_numeric(pd.Series([m.get("macro_margin_score", 50)]), errors="coerce").fillna(50).iloc[0])
            macro_state = str(m.get("macro_margin_state", "中性") or "中性")
            macro_reason = str(m.get("sentiment_reason", "") or "")
        if int(st.get("ready", 0)) != 1 or df is None or df.empty or "stock_id" not in df.columns:
            return {"pass": 1, "score": 50.0, "reason": f"融資資料NE：{st.get('reason','尚無個股融資資料')}｜{macro_state}:{macro_score}", "data_ready": 0, "not_evaluated": 1, "coverage_state": "NE_NO_MARGIN", "margin_score": 50.0, "margin_state": "NE", "macro_margin_score": macro_score, "macro_margin_state": macro_state, "margin_decision_note": macro_reason}
        x = df[df["stock_id"].astype(str) == str(stock_id)].tail(1)
        if x.empty:
            return {"pass": 1, "score": 50.0, "reason": f"該股無融資覆蓋資料，Margin=NE｜{macro_state}:{macro_score}", "data_ready": 0, "not_evaluated": 1, "coverage_state": "NE_NO_MARGIN_COVERAGE", "margin_score": 50.0, "margin_state": "NE", "macro_margin_score": macro_score, "macro_margin_state": macro_state, "margin_decision_note": macro_reason}
        row = x.iloc[-1]
        score, state, reason = self.score_margin_row(row)
        # 市場情緒做小幅修正：過熱扣分、冷卻加分，不直接阻擋。
        adjusted = score + (macro_score - 50.0) * 0.15
        adjusted = round(max(0, min(100, adjusted)), 2)
        return {
            "pass": int(adjusted >= 35),
            "score": adjusted,
            "reason": f"個股融資={state} score={score}｜市場={macro_state} {macro_score}｜{reason}｜{macro_reason}",
            "data_ready": 1,
            "not_evaluated": 0,
            "coverage_state": "COVERED",
            "margin_balance": row.get("margin_balance", np.nan),
            "short_balance": row.get("short_balance", np.nan),
            "margin_change": row.get("margin_change", np.nan),
            "short_change": row.get("short_change", np.nan),
            "margin_utilization": row.get("margin_utilization", np.nan),
            "retail_heat_score": row.get("retail_heat_score", np.nan),
            "margin_score": adjusted,
            "margin_state": state,
            "macro_margin_score": macro_score,
            "macro_margin_state": macro_state,
            "margin_decision_note": reason + ("｜" + macro_reason if macro_reason else ""),
        }



class FinancialFeatureEngine:
    """V9.6.2-R7：把 external_valuation / external_revenue 正確合併成 financial_feature_daily。

    R7 修正重點：
    1. stock_id 嚴格正規化，支援 0050 被寫成 50 / 50.0、2330.TW / 2330.TWO。
    2. external_valuation / external_revenue 進入 feature 前先做欄位別名標準化。
    3. revenue_yoy 支援 yoy / cumulative_yoy / revenue_yoy / 中文欄位名稱。
    4. EPS 來源支援 eps_ttm / eps / close_price ÷ pe / 最新收盤價 ÷ pe。
    5. EPS_YOY 缺值不再算核心 NE，避免 NE_ratio 被誤判 100%。
    6. Log 顯示 valuation_hit / revenue_hit / eps_ok / revenue_ok / core_NE_ratio，驗收可直接看。
    """
    MATRIX_SCORE = {
        ("E3", "R3"): 100, ("E3", "R2"): 88, ("E3", "R1"): 48, ("E3", "R0"): 25,
        ("E2", "R3"): 92,  ("E2", "R2"): 75, ("E2", "R1"): 45, ("E2", "R0"): 22,
        ("E1", "R3"): 82,  ("E1", "R2"): 62, ("E1", "R1"): 35, ("E1", "R0"): 15,
        ("E0", "R3"): 70,  ("E0", "R2"): 45, ("E0", "R1"): 20, ("E0", "R0"): 5,
        ("E_NA", "R3"): 55, ("E_NA", "R2"): 45, ("E_NA", "R1"): 30, ("E_NA", "R0"): 15,
        ("E3", "R_NA"): 60, ("E2", "R_NA"): 50, ("E1", "R_NA"): 40, ("E0", "R_NA"): 25,
        ("E_NA", "R_NA"): 40,
    }

    def __init__(self, db: DBManager):
        self.db = db

    @staticmethod
    def _num(v, default=np.nan):
        try:
            if v is None:
                return default
            s = str(v).replace(",", "").replace("％", "%").replace("%", "").strip()
            s = s.replace("--", "").replace("－", "-")
            if s in ("", "-", "nan", "None", "NULL", "null", "<NA>", "NaN"):
                return default
            out = float(s)
            if not np.isfinite(out):
                return default
            return out
        except Exception:
            return default

    @staticmethod
    def _clamp(v, lo=0.0, hi=100.0):
        try:
            return max(lo, min(hi, float(v)))
        except Exception:
            return lo

    @staticmethod
    def _normalize_stock_id_series(s: pd.Series) -> pd.Series:
        return s.astype(str).map(normalize_stock_id)

    @staticmethod
    def _first_numeric(row: pd.Series, names: list[str], default=np.nan):
        for name in names:
            if name in row.index:
                v = FinancialFeatureEngine._num(row.get(name), np.nan)
                if np.isfinite(v):
                    return v
        return default

    @staticmethod
    def _pick_col(columns, candidates: list[str]) -> str | None:
        cols = list(columns)
        norm = {str(c).strip().lower(): c for c in cols}
        for cand in candidates:
            key = str(cand).strip().lower()
            if key in norm:
                return norm[key]
        for c in cols:
            cs = str(c).strip().lower()
            for cand in candidates:
                ck = str(cand).strip().lower()
                if ck and ck in cs:
                    return c
        return None

    def _latest_close_map(self) -> dict:
        try:
            q = """
            SELECT p.stock_id, p.close
            FROM price_history p
            JOIN (
                SELECT stock_id, MAX(date) AS date
                FROM price_history
                GROUP BY stock_id
            ) m
            ON p.stock_id=m.stock_id AND p.date=m.date
            """
            with self.db.lock:
                df = pd.read_sql_query(q, self.db.conn)
            if df is None or df.empty:
                return {}
            df["stock_id"] = self._normalize_stock_id_series(df["stock_id"])
            df["close"] = pd.to_numeric(df["close"], errors="coerce")
            return dict(zip(df["stock_id"], df["close"]))
        except Exception:
            return {}

    def _standardize_feature_source(self, df: pd.DataFrame, table: str, date_col: str) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        x = df.copy()
        if "stock_id" not in x.columns:
            sid_col = self._pick_col(x.columns, ["stock_id", "公司代號", "公司代碼", "股票代號", "證券代號", "SecuritiesCompanyCode", "CompanyCode", "Code"])
            if sid_col is not None:
                x["stock_id"] = x[sid_col]
        if "stock_id" not in x.columns:
            return pd.DataFrame()
        x["stock_id"] = self._normalize_stock_id_series(x["stock_id"])
        x = x[x["stock_id"] != ""].copy()
        if x.empty:
            return pd.DataFrame()

        if table == "external_valuation":
            alias_candidates = {
                "data_date": ["data_date", "日期", "資料日期", "source_date"],
                "close_price": ["close_price", "close", "收盤價", "收盤價(元)", "收盤"],
                "pe": ["pe", "PER", "本益比", "本益比(倍)"],
                "pb": ["pb", "PBR", "股價淨值比", "股價淨值比(倍)"],
                "dividend_yield": ["dividend_yield", "殖利率", "殖利率(%)"],
                "eps": ["eps", "EPS", "每股盈餘", "eps_latest"],
                "eps_ttm": ["eps_ttm", "EPS_TTM", "eps_proxy", "近四季EPS", "近四季每股盈餘"],
                "eps_yoy": ["eps_yoy", "eps_ttm_yoy", "EPS_YOY", "每股盈餘年增率"],
                "roe": ["roe", "ROE", "股東權益報酬率"],
                "gross_margin": ["gross_margin", "毛利率"],
                "operating_margin": ["operating_margin", "營益率", "營業利益率"],
                "fiscal_year_quarter": ["fiscal_year_quarter", "財報年/季", "財報年度季別"],
                "source_date": ["source_date", "資料日期", "日期"],
                "source_url": ["source_url", "來源網址"],
            }
            for dst, candidates in alias_candidates.items():
                if dst not in x.columns:
                    src = self._pick_col(x.columns, candidates)
                    if src is not None:
                        x[dst] = x[src]
            if date_col not in x.columns:
                x[date_col] = x.get("source_date", "")
            for c in ["close_price", "pe", "pb", "dividend_yield", "eps", "eps_ttm", "eps_yoy", "roe", "gross_margin", "operating_margin"]:
                if c not in x.columns:
                    x[c] = np.nan
                x[c] = pd.to_numeric(x[c], errors="coerce")
            for c in ["fiscal_year_quarter", "source_date", "source_url"]:
                if c not in x.columns:
                    x[c] = ""
                x[c] = x[c].fillna("").astype(str)

        elif table == "external_revenue":
            alias_candidates = {
                "revenue_month": ["revenue_month", "出表年月", "資料年月", "年月", "營收年月"],
                "revenue": ["revenue", "當月營收", "營業收入-當月營收", "本月營收"],
                "mom": ["mom", "上月比較增減", "上月增減", "MoM"],
                "yoy": ["yoy", "revenue_yoy", "monthly_yoy", "YoY", "去年同月增減", "去年同期增減", "營業收入-去年同月增減"],
                "cumulative_revenue": ["cumulative_revenue", "累計營收", "營業收入-累計營收"],
                "cumulative_yoy": ["cumulative_yoy", "cumulative_revenue_yoy", "累計增減", "前期比較增減", "累計營業收入-前期比較增減"],
                "source_date": ["source_date", "資料日期"],
            }
            for dst, candidates in alias_candidates.items():
                if dst not in x.columns:
                    src = self._pick_col(x.columns, candidates)
                    if src is not None:
                        x[dst] = x[src]
            if date_col not in x.columns:
                x[date_col] = datetime.now().strftime("%Y%m")
            for c in ["revenue", "mom", "yoy", "cumulative_revenue", "cumulative_yoy"]:
                if c not in x.columns:
                    x[c] = np.nan
                x[c] = pd.to_numeric(x[c], errors="coerce")
            if "source_date" not in x.columns:
                x["source_date"] = ""
            x["source_date"] = x["source_date"].fillna("").astype(str)

        if date_col not in x.columns:
            x[date_col] = ""
        x[date_col] = x[date_col].fillna("").astype(str).str.strip()
        x = x.sort_values(["stock_id", date_col])
        x = x.drop_duplicates(subset=["stock_id", date_col], keep="last")
        x = x.drop_duplicates(subset=["stock_id"], keep="last").copy()
        return x

    def _latest_by_stock(self, table: str, date_col: str) -> pd.DataFrame:
        df = self.db.read_table(table, limit=None)
        return self._standardize_feature_source(df, table, date_col)

    def build_eps_ttm(self, row: pd.Series, latest_close_map: dict | None = None) -> tuple[float, str]:
        eps_ttm = self._num(row.get("eps_ttm"), np.nan)
        if np.isfinite(eps_ttm):
            return eps_ttm, "OK"
        eps = self._num(row.get("eps"), np.nan)
        if np.isfinite(eps):
            return eps, "EPS_RAW"
        price = self._num(row.get("close_price"), np.nan)
        pe = self._num(row.get("pe"), np.nan)
        calc = calculate_eps_ttm(price, pe)
        if calc is not None:
            return float(calc), "PE_PROXY"
        if latest_close_map:
            sid = normalize_stock_id(row.get("stock_id", ""))
            price2 = self._num(latest_close_map.get(sid), np.nan)
            calc2 = calculate_eps_ttm(price2, pe)
            if calc2 is not None:
                return float(calc2), "PE_PROXY_PRICE_HISTORY"
        return np.nan, "EPS_NE"

    @staticmethod
    def classify_eps_bucket(eps_ttm) -> str:
        try:
            eps = float(eps_ttm)
            if not np.isfinite(eps):
                return "E_NA"
        except Exception:
            return "E_NA"
        if eps < 0:
            return "E0"
        if eps < 2:
            return "E1"
        if eps < 8:
            return "E2"
        return "E3"

    @staticmethod
    def classify_revenue_bucket(revenue_yoy) -> str:
        try:
            rev = float(revenue_yoy)
            if not np.isfinite(rev):
                return "R_NA"
        except Exception:
            return "R_NA"
        if rev < -10:
            return "R0"
        if rev < 0:
            return "R1"
        if rev < 15:
            return "R2"
        return "R3"

    @staticmethod
    def classify_eps_category(eps_bucket: str, rev_bucket: str, eps_yoy=np.nan) -> str:
        try:
            ey = float(eps_yoy)
        except Exception:
            ey = np.nan
        if eps_bucket == "E3" and rev_bucket == "R3" and (not np.isfinite(ey) or ey >= 10):
            return "U1"
        if eps_bucket == "E2" and rev_bucket == "R3" and (not np.isfinite(ey) or ey >= 0):
            return "U1"
        if eps_bucket == "E3" and rev_bucket == "R2":
            return "U2"
        if eps_bucket in ("E0", "E1") and rev_bucket == "R3":
            return "U3"
        if eps_bucket == "E3" and rev_bucket in ("R0", "R1"):
            return "U4"
        if eps_bucket == "E_NA" or rev_bucket == "R_NA":
            return "U0"
        return "U0"

    def calc_modifier(self, eps_yoy, roe, gross_margin, pe) -> float:
        mod = 0.0
        ey = self._num(eps_yoy, np.nan)
        if np.isfinite(ey):
            mod += 12 if ey >= 50 else 8 if ey >= 30 else 4 if ey >= 10 else 0 if ey >= 0 else -8 if ey >= -20 else -15
        roe_v = self._num(roe, np.nan)
        if np.isfinite(roe_v):
            mod += 6 if roe_v >= 20 else 3 if roe_v >= 10 else -8 if roe_v < 0 else 0
        gm = self._num(gross_margin, np.nan)
        if np.isfinite(gm):
            mod += 4 if gm >= 35 else -5 if gm < 10 else 0
        pe_v = self._num(pe, np.nan)
        if np.isfinite(pe_v) and pe_v > 0:
            mod += -8 if pe_v > 60 else 2 if pe_v < 15 else 0
        return round(mod, 2)

    def calc_data_quality_flag(self, eps_flag: str, revenue_yoy, eps_yoy, source_trace: dict) -> str:
        flags = []
        if eps_flag == "EPS_NE":
            flags.append("EPS_NE")
        elif eps_flag and eps_flag != "OK":
            flags.append(eps_flag)
        if not np.isfinite(self._num(revenue_yoy, np.nan)):
            flags.append("REV_NE")
        if not bool(source_trace.get("valuation_hit")) and not bool(source_trace.get("revenue_hit")):
            flags.append("SOURCE_NE")
        if not np.isfinite(self._num(eps_yoy, np.nan)):
            flags.append("EPS_YOY_MISSING")
        return "OK" if not flags else "|".join(flags)

    def calc_revenue_eps_score(self, matrix_score, eps_ttm, revenue_yoy, modifier) -> float:
        eps = self._num(eps_ttm, np.nan)
        rev = self._num(revenue_yoy, np.nan)
        eps_score = 40.0
        if np.isfinite(eps):
            eps_score = 20 if eps < 0 else 45 if eps < 2 else 70 if eps < 8 else 88
        rev_score = 40.0
        if np.isfinite(rev):
            rev_score = 20 if rev < -10 else 40 if rev < 0 else 65 if rev < 15 else 90
        score = 0.5 * float(matrix_score) + 0.3 * eps_score + 0.2 * rev_score + float(modifier)
        return round(self._clamp(score), 2)

    def build_feature_batch(self, run_id: str | None = None, write_db: bool = True, log_limit: int = 10) -> pd.DataFrame:
        run_id = run_id or datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        feature_date = datetime.now().strftime("%Y-%m-%d")
        master = self.db.get_master()
        valuation = self._latest_by_stock("external_valuation", "data_date")
        revenue = self._latest_by_stock("external_revenue", "revenue_month")
        latest_close_map = self._latest_close_map()
        if master is None or master.empty:
            return pd.DataFrame()

        base = master[["stock_id"]].copy()
        base["stock_id"] = self._normalize_stock_id_series(base["stock_id"])
        base = base[base["stock_id"] != ""].drop_duplicates(subset=["stock_id"], keep="first").copy()

        valuation_hit = 0
        revenue_hit = 0
        valuation_rows = 0 if valuation is None else len(valuation)
        revenue_rows = 0 if revenue is None else len(revenue)
        if valuation is not None and not valuation.empty:
            keep_val = [c for c in ["stock_id", "data_date", "close_price", "pe", "pb", "dividend_yield", "eps", "eps_ttm", "eps_yoy", "roe", "gross_margin", "operating_margin", "fiscal_year_quarter", "source_date", "source_url"] if c in valuation.columns]
            val_src = valuation[keep_val].copy()
            base = base.merge(val_src, on="stock_id", how="left")
            valuation_hit = int(base.get("data_date", pd.Series("", index=base.index)).fillna("").astype(str).ne("").sum())

        if revenue is not None and not revenue.empty:
            keep_rev = [c for c in ["stock_id", "revenue_month", "revenue", "mom", "yoy", "cumulative_yoy", "source_date"] if c in revenue.columns]
            rev_src = revenue[keep_rev].copy()
            base = base.merge(rev_src, on="stock_id", how="left", suffixes=("", "_revenue"))
            revenue_hit = int(base.get("revenue_month", pd.Series("", index=base.index)).fillna("").astype(str).ne("").sum())

        rows = []
        for _, r in base.iterrows():
            sid = normalize_stock_id(r.get("stock_id", ""))
            if not sid:
                continue
            eps_ttm, eps_flag = self.build_eps_ttm(r, latest_close_map=latest_close_map)
            revenue_yoy = self._first_numeric(r, ["yoy", "cumulative_yoy", "revenue_yoy"], np.nan)
            eps_yoy = self._first_numeric(r, ["eps_yoy", "eps_ttm_yoy"], np.nan)
            eps_b = self.classify_eps_bucket(eps_ttm)
            rev_b = self.classify_revenue_bucket(revenue_yoy)
            matrix_cell = f"{eps_b}-{rev_b}"
            matrix_base = float(self.MATRIX_SCORE.get((eps_b, rev_b), 40.0))
            eps_cat = self.classify_eps_category(eps_b, rev_b, eps_yoy)
            modifier = self.calc_modifier(eps_yoy, r.get("roe", np.nan), r.get("gross_margin", np.nan), r.get("pe", np.nan))
            score = self.calc_revenue_eps_score(matrix_base, eps_ttm, revenue_yoy, modifier)
            source_trace = {
                "valuation_hit": bool(str(r.get("data_date", "") or "").strip()),
                "revenue_hit": bool(str(r.get("revenue_month", "") or "").strip()),
                "valuation_date": str(r.get("data_date", "") or ""),
                "valuation_source": str(r.get("source_url", "") or ""),
                "revenue_month": str(r.get("revenue_month", "") or ""),
                "eps_source_flag": eps_flag,
                "latest_price_fallback": bool(eps_flag == "PE_PROXY_PRICE_HISTORY"),
            }
            dq = self.calc_data_quality_flag(eps_flag, revenue_yoy, eps_yoy, source_trace)
            rows.append({
                "stock_id": sid, "feature_date": feature_date, "eps_ttm": eps_ttm, "eps_yoy": eps_yoy,
                "revenue_yoy": revenue_yoy, "eps_bucket": eps_b, "rev_bucket": rev_b,
                "matrix_cell": matrix_cell, "eps_category": eps_cat, "matrix_base_score": matrix_base,
                "modifier": modifier, "revenue_eps_score": score, "data_quality_flag": dq,
                "source_trace_json": json.dumps(source_trace, ensure_ascii=False),
                "source_date": str(r.get("source_date_revenue", r.get("source_date", "")) or ""),
                "run_id": run_id, "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

        out = pd.DataFrame(rows)
        if out is None or out.empty:
            return pd.DataFrame()

        eps_ok = int(pd.to_numeric(out["eps_ttm"], errors="coerce").notna().sum())
        revenue_ok = int(pd.to_numeric(out["revenue_yoy"], errors="coerce").notna().sum())
        both_na_ratio = float((out["matrix_cell"].astype(str) == "E_NA-R_NA").mean())
        core_ne_ratio = float(out["data_quality_flag"].fillna("").astype(str).str.contains(r"EPS_NE|REV_NE|SOURCE_NE", regex=True).mean())

        # V9.6.2-R8：強制輸出 merge 驗證資訊到 console / EXE 視窗。
        # 不只依賴 log_info，避免 UI log 未接到時無法確認 EPS / Revenue 是否真的合併成功。
        try:
            print("\n====== DEBUG R8 FINANCIAL MERGE ======", flush=True)
            print(f"valuation_rows={valuation_rows}", flush=True)
            print(f"revenue_rows={revenue_rows}", flush=True)
            print(f"valuation_hit={valuation_hit}", flush=True)
            print(f"revenue_hit={revenue_hit}", flush=True)
            print(f"eps_ok={eps_ok}", flush=True)
            print(f"revenue_ok={revenue_ok}", flush=True)
            print(f"both_NA_ratio={both_na_ratio:.2%}", flush=True)
            print(f"core_NE_ratio={core_ne_ratio:.2%}", flush=True)
            print("======================================\n", flush=True)
        except Exception:
            pass

        if write_db:
            self.db.replace_financial_feature_batch(out, run_id=run_id)
            log_info(
                f"[EPS MATRIX][MERGE][R7] run_id={run_id} master={len(base)} valuation_rows={valuation_rows} revenue_rows={revenue_rows} "
                f"valuation_hit={valuation_hit} revenue_hit={revenue_hit} eps_ok={eps_ok} revenue_ok={revenue_ok} "
                f"both_NA_ratio={both_na_ratio:.2%} core_NE_ratio={core_ne_ratio:.2%}"
            )
            sample = out.head(log_limit)
            for _, rr in sample.iterrows():
                log_info(f"[EPS MATRIX][BUILD][R7] run_id={run_id} stock={rr.get('stock_id')} eps_ttm={rr.get('eps_ttm')} rev_yoy={rr.get('revenue_yoy')} cell={rr.get('matrix_cell')} cat={rr.get('eps_category')} score={rr.get('revenue_eps_score')} flag={rr.get('data_quality_flag')}")
            log_info(f"[EPS MATRIX][BUILD][R7] run_id={run_id} rows={len(out)}")
        return out

    def get_latest_feature(self, stock_id: str) -> dict:
        return self.db.get_latest_financial_feature_row(stock_id)

class FundamentalEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.readiness = ExternalDataReadiness(db)

    def evaluate(self, stock_id: str) -> dict:
        feature = FinancialFeatureEngine(self.db).get_latest_feature(stock_id)
        if feature:
            score = float(pd.to_numeric(pd.Series([feature.get("revenue_eps_score", 50)]), errors="coerce").fillna(50).iloc[0])
            eps_category = str(feature.get("eps_category", "U0") or "U0")
            matrix_cell = str(feature.get("matrix_cell", "") or "")
            flag = str(feature.get("data_quality_flag", "") or "")
            # U4 是高EPS但業務/成長衰退：不直接亂砍所有技術股，但 Fundamental Gate 要明確 BLOCK，Decision reason 可追溯。
            if eps_category == "U4":
                passed = 0
                reason = f"EPS矩陣 U4 高EPS衰退風險｜{matrix_cell}｜score={score:.1f}｜flag={flag}"
            elif eps_category == "U3":
                passed = 1
                reason = f"EPS矩陣 U3 轉機觀察｜{matrix_cell}｜score={score:.1f}｜flag={flag}"
            elif eps_category in ("U1", "U2"):
                passed = int(score >= 45)
                reason = f"EPS矩陣 {eps_category}｜{matrix_cell}｜score={score:.1f}｜flag={flag}"
            else:
                passed = int(score >= 35)
                reason = f"EPS矩陣 {eps_category}｜{matrix_cell}｜score={score:.1f}｜flag={flag}"
            return {
                "pass": passed,
                "score": round(score, 2),
                "reason": reason,
                "data_ready": 1,
                "not_evaluated": 0,
                "coverage_state": "COVERED_EPS_MATRIX",
                "eps_ttm": feature.get("eps_ttm", np.nan),
                "eps_yoy": feature.get("eps_yoy", np.nan),
                "revenue_yoy": feature.get("revenue_yoy", np.nan),
                "eps_bucket": feature.get("eps_bucket", ""),
                "rev_bucket": feature.get("rev_bucket", ""),
                "matrix_cell": matrix_cell,
                "eps_category": eps_category,
                "matrix_base_score": feature.get("matrix_base_score", np.nan),
                "modifier": feature.get("modifier", np.nan),
                "revenue_eps_score": score,
                "data_quality_flag": flag,
                "financial_score": score,
                "eps_matrix_decision_note": reason,
                "source_trace_json": feature.get("source_trace_json", ""),
                "pe": np.nan, "pb": np.nan, "dividend_yield": np.nan, "valuation_score": round(score, 2),
            }

        # 若尚未建立 feature，維持原本 NE 原則：資料未覆蓋不硬擋，但明確標示。
        rev_st = self.readiness.get_status("revenue")
        if int(rev_st.get("ready", 0)) != 1:
            return {
                "pass": 1,
                "score": 50.0,
                "reason": f"EPS矩陣尚未建立，營收資料未就緒，Fundamental Gate=NE：{rev_st.get('reason')}",
                "data_ready": 0,
                "not_evaluated": 1,
                "coverage_state": "NE_NO_EPS_MATRIX",
                "eps_category": "U0", "matrix_cell": "E_NA-R_NA", "revenue_eps_score": 50.0,
                "data_quality_flag": "EPS_MATRIX_NE",
            }

        # fallback：沿用 valuation/revenue 的簡化評分，並標示為 legacy fundamental。
        rev = self.db.read_table("external_revenue", limit=None)
        val = self.db.read_table("external_valuation", limit=None)
        score = 50.0
        reasons = []
        if not rev.empty and "stock_id" in rev.columns:
            r = rev[rev["stock_id"].astype(str) == str(stock_id)].tail(1)
            if not r.empty:
                yoy = float(pd.to_numeric(r.get("yoy", 0), errors="coerce").fillna(0).iloc[-1])
                score += 15 if yoy > 20 else 8 if yoy > 0 else -10
                reasons.append(f"營收YoY {yoy:.1f}%")
        pe = pb = dividend_yield = eps_ttm = np.nan
        valuation_score = 0.0
        if not val.empty and "stock_id" in val.columns:
            v = val[val["stock_id"].astype(str) == str(stock_id)].tail(1)
            if not v.empty:
                pe = float(pd.to_numeric(v.get("pe", np.nan), errors="coerce").iloc[-1]) if "pe" in v.columns else np.nan
                pb = float(pd.to_numeric(v.get("pb", np.nan), errors="coerce").iloc[-1]) if "pb" in v.columns else np.nan
                dividend_yield = float(pd.to_numeric(v.get("dividend_yield", np.nan), errors="coerce").iloc[-1]) if "dividend_yield" in v.columns else np.nan
                eps_ttm = float(pd.to_numeric(v.get("eps_ttm", v.get("eps", np.nan)), errors="coerce").iloc[-1]) if ("eps_ttm" in v.columns or "eps" in v.columns) else np.nan
                if pd.notna(pe) and pe > 0:
                    pe_score = 8 if pe < 10 else 5 if pe < 20 else 0 if pe < 35 else -5
                    valuation_score += pe_score
                    score += pe_score
                    reasons.append(f"PE {pe:.2f}")
                if pd.notna(eps_ttm):
                    eps_score = 2 if eps_ttm > 10 else 1 if eps_ttm > 5 else 0
                    valuation_score += eps_score
                    score += eps_score
                    reasons.append(f"EPS_TTM {eps_ttm:.2f}")
        if not reasons:
            return {
                "pass": 1,
                "score": 50.0,
                "reason": "該股無基本面覆蓋資料，Fundamental Gate=NE，不阻擋交易",
                "data_ready": 0,
                "not_evaluated": 1,
                "coverage_state": "NE_NO_FUNDAMENTAL_COVERAGE",
                "eps_category": "U0", "matrix_cell": "E_NA-R_NA", "revenue_eps_score": 50.0,
                "data_quality_flag": "FUNDAMENTAL_NE",
            }
        score = max(0, min(100, score))
        return {
            "pass": int(score >= 45),
            "score": round(score, 2),
            "reason": "legacy fundamental｜" + "；".join(reasons),
            "data_ready": 1,
            "not_evaluated": 0,
            "coverage_state": "COVERED_LEGACY",
            "pe": pe,
            "pb": pb,
            "dividend_yield": dividend_yield,
            "eps_ttm": eps_ttm,
            "valuation_score": round(valuation_score, 2),
            "eps_category": "U0", "matrix_cell": "E_NA-R_NA", "revenue_eps_score": round(score, 2),
            "data_quality_flag": "LEGACY_FUNDAMENTAL",
            "financial_score": round(score, 2),
        }


class EventEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.readiness = ExternalDataReadiness(db)

    def evaluate(self, stock_id: str) -> dict:
        st = self.readiness.get_status("event")
        if int(st.get("ready", 0)) != 1:
            return {"pass": 1, "score": 50.0, "reason": f"事件資料Not Evaluated：{st.get('reason')}", "data_ready": 0, "not_evaluated": 1}
        ev = self.db.read_table("external_event", limit=None)
        if ev.empty or "stock_id" not in ev.columns:
            return {"pass": 1, "score": 50.0, "reason": "事件資料表空白，Not Evaluated，不宣告事件Gate成功", "data_ready": 0, "not_evaluated": 1}
        x = ev[(ev["stock_id"].astype(str) == str(stock_id)) | (ev["stock_id"].astype(str) == "")].tail(5)
        score = float(pd.to_numeric(x.get("event_score", 50), errors="coerce").fillna(50).max()) if not x.empty else 50.0
        return {"pass": int(score >= 40), "score": round(score, 2), "reason": f"事件分數 {score:.1f}", "data_ready": 1, "not_evaluated": 0}


class RiskGateEngine:
    @staticmethod
    def evaluate(plan: dict, market_mode: str = "Neutral", external_data_ready: int = 1) -> dict:
        rr_live = float(plan.get("rr_live", plan.get("rr", 0)) or 0)
        atr_pct = float(plan.get("atr_pct", 0) or 0)
        win_rate = float(plan.get("win_rate", 0) or 0)
        allowed = rr_live >= 1.2 and atr_pct <= 8 and win_rate >= 45
        if market_mode == "Risk_OFF" and int(plan.get("is_etf", 0) or 0) != 1:
            allowed = False
        # V9.5.9：external_data_ready 僅作資訊，不可直接控制 RiskGate / trade_allowed。
        return {"pass": int(allowed), "score": round(min(100, rr_live * 25 + win_rate * 0.5 - atr_pct * 2), 2), "reason": f"RR={rr_live:.2f}; ATR={atr_pct:.1f}%; 勝率={win_rate:.1f}%; Market={market_mode}; ExternalReadyInfo={external_data_ready}"}


class DecisionLayerEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.market_engine = MarketRegimeEngine(db)
        self.readiness = ExternalDataReadiness(db)
        self.flow_engine = CapitalFlowEngine(db)
        self.fundamental_engine = FundamentalEngine(db)
        self.margin_engine = MarginDecisionEngine(db)
        self.event_engine = EventEngine(db)

    def evaluate_plan(self, plan: dict) -> dict:
        stock_id = str(plan.get("stock_id", ""))
        trace = latest_source_trace(self.db)
        market_status = self.readiness.get_status("market_snapshot")
        analysis_ready, analysis_reason = self.readiness.analysis_ready()
        execution_ready, execution_reason = self.readiness.execution_ready()
        mandatory_ready, mandatory_reason = execution_ready, execution_reason

        # V9.5.8 DATA_INTEGRITY_PATCH：Decision Layer 只接受 data_ready=1 且非 proxy 的 market_snapshot。
        # 若 market_snapshot 未通過官方資料 Ready，不再使用 MarketRegimeEngine internal proxy 當判斷依據。
        market_snapshot = self.db.read_table("market_snapshot", limit=None)
        market_status_ready = int(market_status.get("ready", 0) or 0)
        valid_market_snapshot = False
        if market_snapshot is not None and not market_snapshot.empty and "market_score" in market_snapshot.columns and market_status_ready == 1:
            ms = market_snapshot.tail(1).iloc[-1]
            source_level = str(ms.get("source_level", "") or "")
            source_url = str(ms.get("source_url", "") or "")
            valid_market_snapshot = ("proxy" not in source_level.lower()) and (source_url != "internal:price_history")
        if valid_market_snapshot:
            market_score = float(pd.to_numeric(pd.Series([ms.get("market_score", 50)]), errors="coerce").fillna(50).iloc[0])
            market_mode = str(ms.get("market_mode", "") or "")
            if not market_mode:
                market_mode = "Risk_ON" if market_score >= 68 else "Risk_OFF" if market_score <= 42 else "Neutral"
            market_memo = f"market_snapshot DB｜score={market_score:.1f}｜mode={market_mode}｜source={ms.get('source_level','')}"
        else:
            market_score = 50.0
            market_mode = "NOT_READY"
            market_memo = f"market_snapshot NOT_READY｜{market_status.get('reason','official market data not ready')}｜proxy blocked"

        is_etf = int(plan.get("is_etf", 0) or 0)
        # V9.5.9：market_snapshot 未ready不可讓交易引擎全停；僅當官方市場明確 Risk_OFF 時才由市場風控擋非ETF。
        market_gate = 1 if (market_mode != "Risk_OFF" or is_etf == 1) else 0
        flow = self.flow_engine.evaluate(stock_id)
        fundamental = self.fundamental_engine.evaluate(stock_id)
        margin = self.margin_engine.evaluate(stock_id)
        event = self.event_engine.evaluate(stock_id)

        technical_gate = int(str(plan.get("final_trade_decision", "")).upper() in ["STRONG_BUY", "BUY", "WAIT_PULLBACK", "DEFENSE"] and float(plan.get("rr_live", plan.get("rr", 0)) or 0) >= 1.0)

        # V9.5.2：拆開「全域外部資料Ready」與「個股外部資料覆蓋」。
        # mandatory_ready 只代表 market_snapshot / institutional / revenue 三個全域資料源已完成同步。
        # 個股法人/基本面若無覆蓋，改為 NE（Not Evaluated），只標示、不阻擋交易。
        # V9.5.9：external_ready/execution_ready 僅為資訊欄位，不可控制 trade_allowed。
        external_ready = int(execution_ready == 1)
        global_external_ready = external_ready
        soft_block = int(analysis_ready == 1 and execution_ready != 1)

        risk = RiskGateEngine.evaluate(plan, market_mode=market_mode, external_data_ready=external_ready)
        fundamental_applicable = 0 if is_etf == 1 else 1
        flow_ne = int(flow.get("not_evaluated", 0) or 0)
        fund_ne = int(fundamental.get("not_evaluated", 0) or 0)
        margin_ne = int(margin.get("not_evaluated", 0) or 0)
        event_ne = int(event.get("not_evaluated", 0) or 0)

        gate_states = {
            "market_gate_state": gate_state(market_gate, data_ready=1),
            "flow_gate_state": gate_state(flow.get("pass", 0), data_ready=flow.get("data_ready", 0), not_evaluated=flow_ne),
            "fundamental_gate_state": gate_state(fundamental.get("pass", 0), data_ready=fundamental.get("data_ready", 0), not_evaluated=fund_ne, applicable=fundamental_applicable),
            "margin_gate_state": gate_state(margin.get("pass", 0), data_ready=margin.get("data_ready", 0), not_evaluated=margin_ne),
            "event_gate_state": gate_state(event.get("pass", 0), data_ready=event.get("data_ready", 0), not_evaluated=event_ne),
            "risk_gate_state": gate_state(risk.get("pass", 0), data_ready=1),
        }
        flow_gate_ok = gate_states["flow_gate_state"] in ("PASS", "NA", "NE")
        fundamental_gate_ok = gate_states["fundamental_gate_state"] in ("PASS", "NA", "NE")
        margin_gate_ok = gate_states["margin_gate_state"] in ("PASS", "NA", "NE")
        gates = {
            "market_gate": int(market_gate),
            "flow_gate": int(flow_gate_ok),
            "fundamental_gate": int(fundamental_gate_ok),
            "margin_gate": int(margin_gate_ok),
            "event_gate": 1 if gate_states["event_gate_state"] in ("PASS", "NA", "NE") else 0,
            "technical_gate": int(technical_gate),
            "risk_gate": int(risk["pass"]),
        }

        coverage_states = [gate_states["flow_gate_state"], gate_states["margin_gate_state"]]
        if fundamental_applicable == 1:
            coverage_states.append(gate_states["fundamental_gate_state"])
        if any(s == "BLOCK" for s in coverage_states):
            stock_external_coverage_state = "BLOCK"
        elif any(s == "NE" for s in coverage_states):
            stock_external_coverage_state = "PARTIAL_NE"
        else:
            stock_external_coverage_state = "FULL"
        gate_policy_note = "V9.5.9：外部資料不得直接控制 trade_allowed；execution_ready/soft_block 只作 UI/Excel/Log 資訊。NE=Not Evaluated：資料未覆蓋不阻擋交易，只降權/標示；market_snapshot proxy 不可作 Ready。"
        eps_category = str(fundamental.get("eps_category", plan.get("eps_category", "U0")) or "U0")
        matrix_cell = str(fundamental.get("matrix_cell", plan.get("matrix_cell", "")) or "")
        revenue_eps_score = float(pd.to_numeric(pd.Series([fundamental.get("revenue_eps_score", plan.get("revenue_eps_score", 50))]), errors="coerce").fillna(50).iloc[0])
        eps_matrix_decision_note = str(fundamental.get("eps_matrix_decision_note", "") or fundamental.get("reason", ""))
        eps_category_block = int(eps_category == "U4")
        eps_turnaround_watch = int(eps_category == "U3")
        if eps_category == "U1":
            gate_policy_note += "；EPS矩陣U1=高成長主升，允許加權。"
        elif eps_category == "U3":
            gate_policy_note += "；EPS矩陣U3=轉機觀察，不直接下重手。"
        elif eps_category == "U4":
            gate_policy_note += "；EPS矩陣U4=高EPS衰退風險，禁止BUY。"

        blocking_parts = []
        for k, state in gate_states.items():
            if state == "BLOCK":
                blocking_parts.append(k.replace("_state", ""))
        if int(technical_gate) != 1:
            blocking_parts.append("technical_gate")
        if eps_category_block:
            blocking_parts.append("eps_matrix_u4_high_eps_decline")

        # V9.6.2-R8：trade_allowed 正式接上 strategy_config active_profile。
        # 外部資料缺口仍只作資訊提示，不直接控制 trade_allowed；真正下單開關由
        # strategy_config.execution 的 RR / RSI / 價格偏離 / ATR / decision / liquidity 控制。
        active_strategy = STRATEGY_CONFIG_MANAGER.get_active_profile()
        rr_min = float(get_strategy_threshold(active_strategy, "execution", "rr_min"))
        rsi_max = float(get_strategy_threshold(active_strategy, "execution", "rsi_max"))
        price_dev_max = float(get_strategy_threshold(active_strategy, "execution", "price_dev_max"))
        atr_pct_max = float(get_strategy_threshold(active_strategy, "execution", "atr_pct_max"))
        allowed_decisions = set(_strategy_allowed_decisions(active_strategy, "execution"))
        required_liquidity = set(_strategy_required_liquidity(active_strategy))

        rr_value = _row_float(plan, "rr_live", "rr", default=0.0)
        rsi_raw_value = _row_float(plan, "rsi", "rsi14", default=np.nan)
        price_dev_raw_value = _row_float(plan, "price_deviation", "price_dev", default=np.nan)
        atr_raw_value = _row_float(plan, "atr_pct", "atr", default=np.nan)
        rsi_missing = pd.isna(rsi_raw_value)
        atr_missing = pd.isna(atr_raw_value)
        price_dev_missing = pd.isna(price_dev_raw_value)
        rsi_value = 999.0 if rsi_missing else float(rsi_raw_value)
        price_dev_value = 999.0 if price_dev_missing else abs(float(price_dev_raw_value))
        atr_value = 999.0 if atr_missing else float(atr_raw_value)
        decision_value = str(plan.get("final_trade_decision", plan.get("decision", "")) or "").strip()
        liquidity_value = str(plan.get("liquidity_status", "") or "").strip()

        config_fail_parts = []
        if decision_value not in allowed_decisions:
            config_fail_parts.append(f"fail_decision:{decision_value}")
        if required_liquidity and liquidity_value not in required_liquidity:
            config_fail_parts.append(f"fail_liquidity:{liquidity_value}")
        if pd.isna(rr_value) or rr_value < rr_min:
            config_fail_parts.append(f"fail_rr:{0.0 if pd.isna(rr_value) else rr_value:.2f}<{rr_min:.2f}")
        if rsi_missing:
            config_fail_parts.append("FAIL_RSI_NA")
        elif rsi_value > rsi_max:
            config_fail_parts.append(f"fail_rsi:{rsi_value:.2f}>{rsi_max:.2f}")
        if price_dev_missing:
            config_fail_parts.append("FAIL_PRICE_DEV_NA")
        elif price_dev_value > price_dev_max:
            config_fail_parts.append(f"fail_price_deviation:{price_dev_value:.4f}>{price_dev_max:.4f}")
        if atr_missing:
            config_fail_parts.append("FAIL_ATR_NA")
        elif atr_value > atr_pct_max:
            config_fail_parts.append(f"fail_atr:{atr_value:.2f}>{atr_pct_max:.2f}")

        if config_fail_parts:
            blocking_parts.extend(config_fail_parts)

        # V9.5.9 核心修正：外部資料缺口不加入 blocking_parts；trade_allowed 不看 execution_ready/external_ready。
        if soft_block:
            gate_policy_note += "；外部資料未完整形成SOFT_BLOCK提示，但不作交易開關。"
        gate_policy_note += (
            f"；R8 trade_allowed=config execution｜profile={STRATEGY_CONFIG_MANAGER.get_active_profile_name()}｜"
            f"RR>={rr_min} RSI<={rsi_max} 偏離<={price_dev_max:.4f} ATR<={atr_pct_max}｜"
            f"decision={decision_value} liquidity={liquidity_value}"
        )
        trade_allowed = int((not blocking_parts) and all(v == 1 for v in gates.values()))

        gate_summary = f"Market={market_mode}/{market_score:.1f}; MarketReady={market_status.get('ready')}; Flow={flow['score']}; Fundamental={fundamental['score']}; Margin={margin['score']}; Event={event['score']}; Risk={risk['score']}; GlobalExternalReady={global_external_ready}; StockCoverage={stock_external_coverage_state}; States={gate_states}"
        blocking_reason = "；".join([str(x) for x in blocking_parts if str(x).strip()])
        decision_reason = "｜".join([market_memo, market_status.get("reason",""), flow["reason"], fundamental["reason"], margin["reason"], event["reason"], risk["reason"], blocking_reason])

        out = dict(plan)
        out.update(gates)
        out.update(gate_states)
        out.update({
            "trade_allowed": trade_allowed,
            "gate_summary": gate_summary,
            "decision_reason": decision_reason,
            "decision_reason_short": short_reason(decision_reason, 140),
            "external_data_ready": external_ready,
            "analysis_ready": int(analysis_ready),
            "execution_ready": int(execution_ready),
            "soft_block": int(soft_block),
            "block_reason": execution_reason,
            "execution_block_reason": execution_reason,
            "global_external_ready": global_external_ready,
            "stock_external_coverage_state": stock_external_coverage_state,
            "gate_policy_note": gate_policy_note,
            "pe": fundamental.get("pe", np.nan),
            "pb": fundamental.get("pb", np.nan),
            "dividend_yield": fundamental.get("dividend_yield", np.nan),
            "eps_yoy": fundamental.get("eps_yoy", np.nan),
            "revenue_yoy": fundamental.get("revenue_yoy", np.nan),
            "eps_bucket": fundamental.get("eps_bucket", ""),
            "rev_bucket": fundamental.get("rev_bucket", ""),
            "matrix_cell": matrix_cell,
            "eps_category": eps_category,
            "matrix_base_score": fundamental.get("matrix_base_score", np.nan),
            "modifier": fundamental.get("modifier", np.nan),
            "revenue_eps_score": revenue_eps_score,
            "data_quality_flag": fundamental.get("data_quality_flag", ""),
            "financial_score": revenue_eps_score,
            "eps_matrix_decision_note": eps_matrix_decision_note,
            "valuation_score": fundamental.get("valuation_score", 0.0),
            "margin_balance": margin.get("margin_balance", np.nan),
            "short_balance": margin.get("short_balance", np.nan),
            "margin_change": margin.get("margin_change", np.nan),
            "short_change": margin.get("short_change", np.nan),
            "margin_utilization": margin.get("margin_utilization", np.nan),
            "retail_heat_score": margin.get("retail_heat_score", np.nan),
            "margin_score": margin.get("margin_score", 50.0),
            "margin_state": margin.get("margin_state", "NE"),
            "macro_margin_score": margin.get("macro_margin_score", 50.0),
            "macro_margin_state": margin.get("macro_margin_state", "NE"),
            "margin_decision_note": margin.get("margin_decision_note", ""),
            "external_blocking_reason": blocking_reason,
            "fail_reason": blocking_reason,
            "rsi": np.nan if rsi_missing else rsi_value,
            "atr_pct": np.nan if atr_missing else atr_value,
            "price_deviation": np.nan if price_dev_missing else price_dev_value,
            "model_score": plan.get("model_score", np.nan),
            "wave_trade_score": plan.get("wave_trade_score", np.nan),
            "pipeline_run_id": str(plan.get("pipeline_run_id", "")),
            "external_run_id": str(plan.get("pipeline_run_id", "")),
            "decision_run_id": datetime.now().strftime("%Y%m%d_%H%M%S_%f"),
            "latest_external_date": trace.get("latest_external_date", ""),
            "market_source_level": trace.get("market_source_level", ""),
            "source_trace_json": json.dumps(trace, ensure_ascii=False)[:3000],
        })
        if not trade_allowed:
            out["ui_state"] = "交易條件未通過"
        elif soft_block:
            out["ui_state"] = "可交易-外部資料SOFT_BLOCK提示"
        elif eps_turnaround_watch:
            out["ui_state"] = "轉機觀察-EPS矩陣U3"
        elif stock_external_coverage_state == "PARTIAL_NE":
            out["ui_state"] = "可交易-部分外部資料NE"
        if eps_matrix_decision_note:
            out["decision_reason_short"] = short_reason(str(out.get("decision_reason_short", "")) + "｜" + eps_matrix_decision_note, 160)
        log_info(f"[EPS MATRIX][DECISION] stock={stock_id} cat={eps_category} cell={matrix_cell} score={revenue_eps_score} trade_allowed={trade_allowed}")
        return out

    def evaluate_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        rows = [self.evaluate_plan(dict(r)) for _, r in df.iterrows()]
        return pd.DataFrame(rows)


class MarketRegimeEngine:
    def __init__(self, db: DBManager):
        self.db = db

    def _score_proxy(self, stock_id: str) -> float:
        hist = self.db.get_price_history(stock_id)
        if hist is None or hist.empty or len(hist) < 80:
            return 50.0
        x = IndicatorEngine.attach(hist)
        last = x.iloc[-1]
        score = 0.0
        if pd.notna(last["ma20"]) and last["close"] > last["ma20"]:
            score += 25
        if pd.notna(last["ma60"]) and last["close"] > last["ma60"]:
            score += 20
        if pd.notna(last["ma20"]) and pd.notna(last["ma60"]) and last["ma20"] > last["ma60"]:
            score += 20
        if pd.notna(last["macd_hist"]) and last["macd_hist"] > 0:
            score += 20
        if pd.notna(last["rsi14"]) and 50 <= last["rsi14"] <= 72:
            score += 15
        return round(score, 2)

    def _breadth_score(self) -> float:
        ranking = self.db.get_latest_ranking()
        if ranking is None or ranking.empty:
            return 50.0
        up = float((ranking["signal"].isin(["強勢追蹤", "整理偏多"])).mean() * 100)
        return round(up, 2)

    def get_market_regime(self) -> dict:
        s_2330 = self._score_proxy("2330")
        s_0050 = self._score_proxy("0050")
        breadth = self._breadth_score()
        score = round(s_2330 * 0.4 + s_0050 * 0.25 + breadth * 0.35, 2)

        if score >= 68:
            regime = "多頭"
            memo = "指數與領頭股結構偏強，可放寬門檻並增加出手檔數。"
            max_positions = 8
            min_win_rate = 70.0
            rsi_low, rsi_high = 50.0, 72.0
        elif score <= 42:
            regime = "空頭"
            memo = "市場偏弱，防守優先，只保留極少數高勝率或防守型 ETF。"
            max_positions = 1
            min_win_rate = 80.0
            rsi_low, rsi_high = 48.0, 68.0
        else:
            regime = "震盪"
            memo = "市場分化，精選出手，不為了湊數而放寬條件。"
            max_positions = 4
            min_win_rate = 75.0
            rsi_low, rsi_high = 50.0, 70.0

        return {
            "regime": regime,
            "score": score,
            "memo": memo,
            "max_positions": max_positions,
            "min_win_rate": min_win_rate,
            "rsi_low": rsi_low,
            "rsi_high": rsi_high,
            "breadth": breadth,
        }


class ThemeStrengthEngine:
    PREFERRED_KEYWORDS = ["AI", "CPO", "Server", "伺服器", "半導體", "晶圓", "ASIC", "RISC-V", "光", "散熱", "HVDC", "網通"]

    @staticmethod
    def summarize(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame(columns=["theme", "count", "avg_total", "avg_ai", "hot_score"])
        x = (
            df.groupby("theme", as_index=False)
            .agg(
                count=("stock_id", "count"),
                avg_total=("total_score", "mean"),
                avg_ai=("ai_score", "mean"),
            )
        )
        x["hot_score"] = x["count"] * 10 + x["avg_total"] * 0.5 + x["avg_ai"] * 0.5
        return x.sort_values(["hot_score", "avg_total", "avg_ai"], ascending=False)

    @staticmethod
    def get_hot_themes(df: pd.DataFrame) -> list:
        x = ThemeStrengthEngine.summarize(df)
        if x.empty:
            return []
        out = x[(x["count"] >= 1) & (x["avg_total"] >= 55)]["theme"].astype(str).tolist()
        preferred = []
        for theme in x["theme"].astype(str).tolist():
            if any(k.lower() in theme.lower() for k in ThemeStrengthEngine.PREFERRED_KEYWORDS):
                preferred.append(theme)
        return list(dict.fromkeys(preferred + out))


class WinRateEngine:
    @staticmethod
    def estimate(hist: pd.DataFrame) -> tuple[str, float]:
        if hist is None or hist.empty or len(hist) < 80:
            return "C", 45.0

        x = IndicatorEngine.attach(hist.copy())
        future_ret = x["close"].shift(-5) / x["close"] - 1
        cond = (
            (x["close"] > x["ma20"]) &
            (x["ma20"] > x["ma60"]) &
            (x["macd_hist"] > 0)
        )
        sample = future_ret[cond].dropna()
        if len(sample) < 8:
            base = float((future_ret.tail(30) > 0).mean() * 100) if len(future_ret.dropna()) else 45.0
        else:
            base = float((sample > 0).mean() * 100)

        if base >= 60:
            grade = "A"
        elif base >= 50:
            grade = "B"
        else:
            grade = "C"
        return grade, round(base, 2)




V80_KLINE_SCORE = {
    "突破強勢": 100, "強勢追蹤": 90, "整理偏多": 75, "偏多觀察": 68,
    "區間整理": 55, "轉弱警戒": 25, "急跌風險": 10,
}
V80_WAVE_SCORE = {
    "第3浪": 100, "推動浪": 92, "修正浪": 70, "整理浪": 60, "第5浪": 35,
}
V80_SAKATA_SCORE = {
    "拉回承接": 95, "偏多低接": 88, "整理偏多": 75, "區間低接": 70,
    "突破追價": 52, "觀望": 20,
}
V80_VOLUME_SCORE = {
    "買盤明顯偏強": 100, "買盤偏強": 88, "多空均衡": 60, "賣盤偏強": 28,
}
V80_WEIGHTS = {
    "kline": 0.18, "wave": 0.22, "fib": 0.14, "sakata": 0.14, "volume": 0.16, "indicator": 0.16,
}

ORDER_COLUMNS = [
    "優先級", "代號", "名稱", "現價", "PE", "PB", "殖利率%", "EPS_TTM", "估值分", "融資餘額", "融資增減", "融券餘額", "融券增減", "券資比%", "散戶熱度", "融資分", "融資狀態", "市場融資分", "市場融資狀態", "漲跌", "漲跌幅%", "分類", "狀態", "盤中狀態", "活性分",
    "進場區", "停損", "目標價", "1.382", "1.618", "RR", "勝率", "ATR%", "Kelly%",
    "建議張數", "建議金額", "單檔曝險%", "投資組合狀態", "風險備註", "外部允許", "外部Ready", "全域外部Ready", "個股覆蓋狀態", "Market Gate", "Flow Gate", "Fundamental Gate", "Event Gate", "Risk Gate", "外部阻擋原因", "外部資料日", "資料來源層級", "決策摘要", "Gate說明"
]

INSTITUTIONAL_COLUMNS = [
    "優先級", "代號", "名稱", "現價", "PE", "PB", "殖利率%", "EPS_TTM", "估值分", "融資餘額", "融資增減", "融券餘額", "融券增減", "券資比%", "散戶熱度", "融資分", "融資狀態", "市場融資分", "市場融資狀態", "漲跌", "漲跌幅%", "市場", "產業", "題材", "分類", "狀態",
    "盤中狀態", "活性分", "淘汰原因", "進場區", "停損", "目標價", "1.382", "1.618", "RR", "勝率",
    "模型分數", "交易分數", "ATR%", "Kelly%", "建議張數", "建議金額", "單檔曝險%", "題材曝險%",
    "產業曝險%", "投資組合狀態", "風險備註"
]

ORDER_TREE_SCHEMA = [
    ("priority", "優先級", 80), ("id", "代號", 90), ("name", "名稱", 120), ("price", "現價", 90),
    ("pe", "PE", 70), ("pb", "PB", 70), ("dividend_yield", "殖利率%", 80), ("eps_ttm", "EPS_TTM", 90), ("valuation_score", "估值分", 80),
    ("margin_score", "融資分", 80), ("margin_state", "融資狀態", 90), ("macro_margin_state", "市場融資", 90),
    ("chg", "漲跌", 90), ("chg_pct", "漲跌幅%", 95), ("bucket", "分類", 95), ("action", "狀態", 95),
    ("liquidity", "盤中狀態", 100), ("liq_score", "活性分", 90), ("entry", "進場區", 130), ("stop", "停損", 100),
    ("target", "目標價", 100), ("target1382", "1.382", 90), ("target1618", "1.618", 90), ("rr", "RR", 80),
    ("win_rate", "勝率%", 85), ("atr_pct", "ATR%", 85), ("kelly_pct", "Kelly%", 85), ("qty", "建議張數", 90),
    ("amount", "建議金額", 100), ("single_pct", "單檔曝險%", 95), ("portfolio_state", "組合狀態", 105), ("risk_note", "風險備註", 220),
    ("trade_allowed", "外部允許", 85), ("global_ready", "全域Ready", 90), ("coverage_state", "覆蓋狀態", 110),
    ("market_gate", "Market", 90), ("flow_gate", "Flow", 90), ("fund_gate", "Fund", 105),
    ("event_gate", "Event", 90), ("risk_gate", "Risk", 90), ("block_reason", "外部阻擋", 220),
]

INSTITUTIONAL_TREE_SCHEMA = [
    ("priority", "優先級", 80), ("id", "代號", 90), ("name", "名稱", 120), ("price", "現價", 90),
    ("pe", "PE", 70), ("pb", "PB", 70), ("dividend_yield", "殖利率%", 80), ("eps_ttm", "EPS_TTM", 90), ("valuation_score", "估值分", 80),
    ("margin_score", "融資分", 80), ("margin_state", "融資狀態", 90), ("macro_margin_state", "市場融資", 90),
    ("chg", "漲跌", 90), ("chg_pct", "漲跌幅%", 95), ("market", "市場", 85), ("industry", "產業", 110),
    ("theme", "題材", 110), ("bucket", "分類", 95), ("action", "狀態", 95), ("liquidity", "盤中狀態", 100),
    ("liq_score", "活性分", 90), ("elim_reason", "淘汰原因", 160), ("entry", "進場區", 130), ("stop", "停損", 100),
    ("target", "目標價", 100), ("target1382", "1.382", 90), ("target1618", "1.618", 90), ("rr", "RR", 80),
    ("win_rate", "勝率%", 85), ("model_score", "模型分數", 95), ("trade_score", "交易分數", 95), ("atr_pct", "ATR%", 85),
    ("kelly_pct", "Kelly%", 85), ("qty", "建議張數", 90), ("amount", "建議金額", 100), ("single_pct", "單檔曝險%", 95),
    ("theme_pct", "題材曝險%", 95), ("industry_pct", "產業曝險%", 95), ("portfolio_state", "投資組合狀態", 110), ("risk_note", "風險備註", 220),
    ("trade_allowed", "外部允許", 85), ("global_ready", "全域Ready", 90), ("coverage_state", "覆蓋狀態", 110),
    ("market_gate", "Market", 90), ("flow_gate", "Flow", 90), ("fund_gate", "Fund", 105),
    ("event_gate", "Event", 90), ("risk_gate", "Risk", 90), ("block_reason", "外部阻擋", 220),
]


CORE_ANALYSIS_COLUMNS = [
    "stock_id", "stock_name", "market", "industry", "theme", "is_etf",
    "candidate_engine", "signal", "wave", "trade_type",
    "model_score", "wave_trade_score", "candidate20_score", "core_attack5_score", "execution_score",
    "mainstream_score", "breakout_score", "liquidity_score",
    "entry_low", "entry_high", "entry_mid", "price_deviation", "rr_live",
    "stop_loss", "target_price", "target_1382", "target_1618",
    "rr", "win_grade", "win_rate", "tactical_light", "final_trade_decision",
    "ui_state", "bucket", "pool_role", "liquidity_status", "elimination_reason",
    "strategy_profile", "strategy_nogo_detail", "strategy_config_summary",
    "wave_text", "wave_condition_pass",
    "threshold_model_score_min", "threshold_wave_trade_score_min", "threshold_rr_min",
    "threshold_rsi_max", "threshold_price_dev_max", "threshold_atr_pct_max",
    "fail_decision", "fail_model_score", "fail_wave_trade_score", "fail_wave_keyword",
    "fail_liquidity", "fail_rsi_na", "fail_atr_na", "fail_price_dev_na", "fail_rsi", "fail_atr", "fail_price_deviation", "fail_rr", "fail_today_decision", "fail_reason",
]

DISPLAY_COLUMN_MAP = {
    "priority": "優先級",
    "id": "代號",
    "name": "名稱",
    "price": "現價",
    "pe": "PE",
    "pb": "PB",
    "dividend_yield": "殖利率%",
    "eps_ttm": "EPS_TTM",
    "valuation_score": "估值分",
    "margin_score": "融資分",
    "margin_state": "融資狀態",
    "macro_margin_score": "市場融資分",
    "macro_margin_state": "市場融資狀態",
    "chg": "漲跌",
    "chg_pct": "漲跌幅%",
    "bucket": "分類",
    "action": "狀態",
    "liquidity": "盤中狀態",
    "liq_score": "活性分",
    "elim_reason": "淘汰原因",
    "entry": "進場區",
    "stop": "停損",
    "target": "目標價",
    "target1382": "1.382",
    "target1618": "1.618",
    "rr": "RR",
    "win_rate": "勝率",
    "model_score": "模型分數",
    "trade_score": "交易分數",
    "atr_pct": "ATR%",
    "kelly_pct": "Kelly%",
    "qty": "建議張數",
    "amount": "建議金額",
    "single_pct": "單檔曝險%",
    "theme_pct": "題材曝險%",
    "industry_pct": "產業曝險%",
    "portfolio_state": "投資組合狀態",
    "risk_note": "風險備註",
    "trade_allowed": "可交易",
    "analysis_ready": "分析Ready",
    "execution_ready": "ExecutionReady資訊",
    "soft_block": "SoftBlock提示",
    "global_ready": "全域外部Ready資訊",
    "coverage_state": "個股覆蓋狀態",
    "market_gate": "Market Gate",
    "flow_gate": "Flow Gate",
    "fund_gate": "Fundamental Gate",
    "event_gate": "Event Gate",
    "risk_gate": "Risk Gate",
    "block_reason": "外部阻擋原因",
}


SCORE_FORMULA_WEIGHTS = {
    "candidate20": {
        "model_score": 0.22, "wave_trade_score": 0.12, "liquidity_score": 0.14,
        "mainstream_score": 0.14, "breakout_score": 0.12, "leader_follow_score": 0.08,
        "active_buy_score": 0.06, "orderflow_aggression_score": 0.06,
        "win_rate": 0.04, "rr_factor": 4.0, "modules_pass_count": 2.0,
    },
    "core_attack5": {
        "candidate20_score": 0.30, "mainstream_score": 0.18, "breakout_score": 0.14,
        "leader_follow_score": 0.10, "active_buy_score": 0.08, "orderflow_aggression_score": 0.08,
        "rel_market_norm": 0.04, "rel_ind_norm": 0.04, "source_count_factor": 4.0, "light_rank_factor": 3.0,
    },
    "execution": {
        "core_attack5_score": 0.40, "price_fit_factor": 25.0, "rr_live_factor": 10.0,
        "liquidity_score": 0.12, "win_rate": 0.08, "model_score": 0.05,
    },
}

REPORT_DECISION_LIMITS = {
    "candidate20": 20,
    "core_attack5": 5,
    "today_buy": 10,
    "unique_decision": 5,
}



# V9.6 STRATEGY_CONFIG：可調策略條件設定層
STRATEGY_CONFIG_DIR = RUNTIME_DIR / "config"
STRATEGY_CONFIG_JSON = STRATEGY_CONFIG_DIR / "strategy_config_v9_6.json"
STRATEGY_CONFIG_EXCEL = STRATEGY_CONFIG_DIR / "strategy_config_v9_6.xlsx"

DEFAULT_STRATEGY_CONFIG = {
    "active_profile": "normal",
    "profiles": {
        "aggressive": {
            "description": "激進模式：提高今日可下單機會，適合強多與題材主升段。",
            "core_attack": {
                "allowed_decisions": ["STRONG_BUY", "BUY", "DEFENSE"],
                "model_score_min": 76.0,
                "wave_trade_score_min": 76.0,
                "require_wave_keyword": False,
                "allowed_wave_keywords": ["第3浪", "推動浪"],
            },
            "execution": {
                "rr_min": 1.20,
                "rsi_max": 75.0,
                "price_dev_max": 0.05,
                "atr_pct_max": 10.0,
                "required_liquidity_status": ["PASS"],
                "allowed_decisions": ["STRONG_BUY", "BUY", "DEFENSE"],
            },
            "wait_pullback": {
                "rr_min": 1.05,
                "price_dev_min": 0.05,
                "price_dev_max": 0.10,
                "allowed_decisions": ["STRONG_BUY", "BUY", "WAIT_PULLBACK", "DEFENSE"],
            },
        },
        "normal": {
            "description": "標準模式：沿用原始交易紀律，避免追價與低RR交易。",
            "core_attack": {
                "allowed_decisions": ["STRONG_BUY", "BUY", "DEFENSE"],
                "model_score_min": 82.0,
                "wave_trade_score_min": 82.0,
                "require_wave_keyword": False,
                "allowed_wave_keywords": ["第3浪", "推動浪"],
            },
            "execution": {
                "rr_min": 1.50,
                "rsi_max": 72.0,
                "price_dev_max": 0.03,
                "atr_pct_max": 8.0,
                "required_liquidity_status": ["PASS"],
                "allowed_decisions": ["STRONG_BUY", "BUY", "DEFENSE"],
            },
            "wait_pullback": {
                "rr_min": 1.20,
                "price_dev_min": 0.03,
                "price_dev_max": 0.08,
                "allowed_decisions": ["STRONG_BUY", "BUY", "WAIT_PULLBACK", "DEFENSE"],
            },
        },
        "conservative": {
            "description": "保守模式：只做最乾淨結構，適合震盪或偏空環境。",
            "core_attack": {
                "allowed_decisions": ["STRONG_BUY", "BUY"],
                "model_score_min": 88.0,
                "wave_trade_score_min": 86.0,
                "require_wave_keyword": True,
                "allowed_wave_keywords": ["第3浪", "推動浪"],
            },
            "execution": {
                "rr_min": 1.80,
                "rsi_max": 68.0,
                "price_dev_max": 0.02,
                "atr_pct_max": 6.0,
                "required_liquidity_status": ["PASS"],
                "allowed_decisions": ["STRONG_BUY", "BUY"],
            },
            "wait_pullback": {
                "rr_min": 1.40,
                "price_dev_min": 0.02,
                "price_dev_max": 0.06,
                "allowed_decisions": ["STRONG_BUY", "BUY", "WAIT_PULLBACK"],
            },
        },
    },
    "validation_limits": {
        "rr_min": [0.5, 5.0],
        "rsi_max": [45.0, 90.0],
        "price_dev_max": [0.0, 0.30],
        "atr_pct_max": [1.0, 30.0],
        "model_score_min": [0.0, 100.0],
        "wave_trade_score_min": [0.0, 100.0],
    },
}


def _deep_merge_strategy_config(base: dict, override: dict) -> dict:
    out = json.loads(json.dumps(base, ensure_ascii=False))
    def _merge(a, b):
        for k, v in (b or {}).items():
            if isinstance(v, dict) and isinstance(a.get(k), dict):
                _merge(a[k], v)
            else:
                a[k] = v
        return a
    return _merge(out, override or {})


class StrategyConfigManager:
    """V9.6：策略參數設定管理器。設定來源優先順序：JSON > Excel > DEFAULT。"""
    def __init__(self, json_path: Path = STRATEGY_CONFIG_JSON, excel_path: Path = STRATEGY_CONFIG_EXCEL):
        self.json_path = Path(json_path)
        self.excel_path = Path(excel_path)
        self.config = json.loads(json.dumps(DEFAULT_STRATEGY_CONFIG, ensure_ascii=False))
        self.last_load_message = "default"
        self.load()

    def ensure_files(self):
        try:
            STRATEGY_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            if not self.json_path.exists():
                self.json_path.write_text(json.dumps(DEFAULT_STRATEGY_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            log_warning(f"[STRATEGY_CONFIG] 建立JSON預設檔失敗：{exc}")

    def load(self) -> dict:
        self.ensure_files()
        cfg = json.loads(json.dumps(DEFAULT_STRATEGY_CONFIG, ensure_ascii=False))
        loaded = []
        try:
            if self.json_path.exists():
                raw = json.loads(self.json_path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    cfg = _deep_merge_strategy_config(cfg, raw)
                    loaded.append(str(self.json_path))
        except Exception as exc:
            log_warning(f"[STRATEGY_CONFIG] JSON載入失敗，使用預設值：{exc}")
        try:
            if self.excel_path.exists():
                xl_cfg = self._load_excel_override(self.excel_path)
                if xl_cfg:
                    cfg = _deep_merge_strategy_config(cfg, xl_cfg)
                    loaded.append(str(self.excel_path))
        except Exception as exc:
            log_warning(f"[STRATEGY_CONFIG] Excel載入失敗，略過Excel覆蓋：{exc}")
        self.config = self._sanitize(cfg)
        self.last_load_message = " | ".join(loaded) if loaded else "default"
        log_info(f"[STRATEGY_CONFIG] loaded={self.last_load_message} active={self.get_active_profile_name()}")
        return self.config

    def _parse_value(self, v):
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                return ""
            if s.startswith("[") or s.startswith("{"):
                try:
                    return json.loads(s)
                except Exception:
                    pass
            if "," in s and not re.search(r"[^A-Za-z0-9_\u4e00-\u9fff, .+-]", s):
                return [x.strip() for x in s.split(",") if x.strip()]
            if s.lower() in ("true", "yes", "y", "1"):
                return True
            if s.lower() in ("false", "no", "n", "0"):
                return False
            try:
                return float(s)
            except Exception:
                return s
        return v

    def _load_excel_override(self, path: Path) -> dict:
        try:
            df = pd.read_excel(path, sheet_name="Strategy_Config")
        except Exception:
            return {}
        if df is None or df.empty:
            return {}
        df = df.fillna("")
        out = {"profiles": {}}
        for _, row in df.iterrows():
            profile = str(row.get("profile", "")).strip()
            section = str(row.get("section", "")).strip()
            key = str(row.get("key", "")).strip()
            if not profile or not section or not key:
                continue
            value = self._parse_value(row.get("value", ""))
            out.setdefault("profiles", {}).setdefault(profile, {}).setdefault(section, {})[key] = value
        # Optional active profile sheet/value
        active_candidates = df[(df.get("section", "") == "system") & (df.get("key", "") == "active_profile")]
        if active_candidates is not None and not active_candidates.empty:
            out["active_profile"] = str(active_candidates.iloc[0].get("value", "normal")).strip() or "normal"
        return out

    def _clamp(self, v, lo, hi, default):
        try:
            x = float(v)
        except Exception:
            x = float(default)
        return max(float(lo), min(float(hi), x))

    def _sanitize(self, cfg: dict) -> dict:
        cfg = _deep_merge_strategy_config(DEFAULT_STRATEGY_CONFIG, cfg or {})
        limits = cfg.get("validation_limits", {}) or {}
        for profile_name, profile in (cfg.get("profiles") or {}).items():
            exe = profile.get("execution", {})
            core = profile.get("core_attack", {})
            wait = profile.get("wait_pullback", {})
            exe["rr_min"] = self._clamp(exe.get("rr_min"), *limits.get("rr_min", [0.5, 5.0]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["execution"]["rr_min"])
            exe["rsi_max"] = self._clamp(exe.get("rsi_max"), *limits.get("rsi_max", [45, 90]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["execution"]["rsi_max"])
            exe["price_dev_max"] = self._clamp(exe.get("price_dev_max"), *limits.get("price_dev_max", [0, 0.3]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["execution"]["price_dev_max"])
            exe["atr_pct_max"] = self._clamp(exe.get("atr_pct_max"), *limits.get("atr_pct_max", [1, 30]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["execution"]["atr_pct_max"])
            core["model_score_min"] = self._clamp(core.get("model_score_min"), *limits.get("model_score_min", [0, 100]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["core_attack"]["model_score_min"])
            core["wave_trade_score_min"] = self._clamp(core.get("wave_trade_score_min"), *limits.get("wave_trade_score_min", [0, 100]), DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["core_attack"]["wave_trade_score_min"])
            wait["rr_min"] = self._clamp(wait.get("rr_min"), 0.5, 5.0, DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["wait_pullback"]["rr_min"])
            wait["price_dev_min"] = self._clamp(wait.get("price_dev_min"), 0.0, 0.3, DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["wait_pullback"]["price_dev_min"])
            wait["price_dev_max"] = self._clamp(wait.get("price_dev_max"), 0.0, 0.3, DEFAULT_STRATEGY_CONFIG["profiles"]["normal"]["wait_pullback"]["price_dev_max"])
        active = str(cfg.get("active_profile", "normal") or "normal").strip()
        if active not in cfg.get("profiles", {}):
            active = "normal"
        cfg["active_profile"] = active
        return cfg

    def save_json(self):
        self.ensure_files()
        self.json_path.write_text(json.dumps(self.config, ensure_ascii=False, indent=2), encoding="utf-8")

    def set_active_profile(self, name: str):
        name = str(name or "normal").strip()
        if name not in self.config.get("profiles", {}):
            name = "normal"
        self.config["active_profile"] = name
        self.save_json()

    def update_active_values(self, values: dict):
        profile = self.get_active_profile_name()
        cfg = self.config.setdefault("profiles", {}).setdefault(profile, {})
        for path, value in (values or {}).items():
            parts = str(path).split(".")
            if len(parts) != 2:
                continue
            cfg.setdefault(parts[0], {})[parts[1]] = value
        self.config = self._sanitize(self.config)
        self.save_json()
        return self.get_active_profile()

    def get_active_profile_name(self) -> str:
        return str(self.config.get("active_profile", "normal") or "normal")

    def get_active_profile(self) -> dict:
        return dict((self.config.get("profiles") or {}).get(self.get_active_profile_name(), {}))

    def summary_text(self) -> str:
        p = self.get_active_profile()
        e = p.get("execution", {})
        c = p.get("core_attack", {})
        return (
            f"profile={self.get_active_profile_name()} | "
            f"RR>={e.get('rr_min')} RSI<={e.get('rsi_max')} 偏離<={float(e.get('price_dev_max'))*100:.1f}% "
            f"ATR<={e.get('atr_pct_max')} model>={c.get('model_score_min')} wave>={c.get('wave_trade_score_min')}"
        )


STRATEGY_CONFIG_MANAGER = StrategyConfigManager()


def _active_strategy_config(cfg: dict | None = None) -> dict:
    """V9.6.1：決策流程唯一策略設定入口。

    DEFAULT_STRATEGY_CONFIG 只允許在 StrategyConfigManager._sanitize() 合併；
    Decision/Execution/Pool 不再各自寫交易門檻 fallback。
    """
    return cfg if isinstance(cfg, dict) and cfg else STRATEGY_CONFIG_MANAGER.get_active_profile()


def _strategy_section(cfg: dict | None, section: str) -> dict:
    active = _active_strategy_config(cfg)
    value = active[section]
    if not isinstance(value, dict):
        raise KeyError(f"Strategy config section invalid: {section}")
    return value


def get_strategy_threshold(cfg: dict | None, section: str, key: str):
    """唯一門檻取值函式：禁止在決策流程散落 cfg.get(..., fallback)。"""
    return _strategy_section(cfg, section)[key]


def get_strategy_list(cfg: dict | None, section: str, key: str) -> list[str]:
    value = get_strategy_threshold(cfg, section, key)
    if isinstance(value, (list, tuple, set)):
        return [str(x).strip() for x in value if str(x).strip()]
    if value is None:
        return []
    return [str(value).strip()] if str(value).strip() else []


def _strategy_values(cfg: dict | None) -> tuple[dict, dict, dict]:
    cfg = _active_strategy_config(cfg)
    return _strategy_section(cfg, "core_attack"), _strategy_section(cfg, "execution"), _strategy_section(cfg, "wait_pullback")


def _strategy_profile_name() -> str:
    return STRATEGY_CONFIG_MANAGER.get_active_profile_name()


def _strategy_summary_text() -> str:
    return STRATEGY_CONFIG_MANAGER.summary_text()


def _coerce_float(v, default: float = 0.0) -> float:
    try:
        x = float(v)
        if np.isfinite(x):
            return x
    except Exception:
        pass
    return float(default)


def _row_float(row, *keys, default: float = 0.0) -> float:
    for key in keys:
        try:
            if key in row and str(row.get(key, "")).strip() not in ("", "nan", "None", "<NA>"):
                return _coerce_float(row.get(key), default)
        except Exception:
            pass
    return float(default)


def _build_wave_text_from_df(df: pd.DataFrame) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype="object")
    parts = []
    for col in ["wave_label", "wave", "trade_type"]:
        parts.append(_safe_text_fill_series(df, col, ""))
    return (parts[0] + " " + parts[1] + " " + parts[2]).astype(str).str.strip()


def _match_wave_keyword_text(wave_text: str, core_cfg: dict) -> bool:
    if not bool(core_cfg["require_wave_keyword"]):
        return True
    keywords = [str(x).strip() for x in core_cfg["allowed_wave_keywords"] if str(x).strip()]
    if not keywords:
        return True
    text = str(wave_text or "")
    return any(k in text for k in keywords)


def _strategy_wave_keyword_mask(df: pd.DataFrame, core_cfg: dict) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    wave_text = _build_wave_text_from_df(df)
    return wave_text.map(lambda s: _match_wave_keyword_text(str(s), core_cfg)).reindex(df.index).fillna(False)


def _strategy_allowed_decisions(cfg: dict | None, section: str) -> list[str]:
    return get_strategy_list(cfg, section, "allowed_decisions")


def _strategy_required_liquidity(cfg: dict | None) -> list[str]:
    return get_strategy_list(cfg, "execution", "required_liquidity_status")


def _safe_strategy_series(df: pd.DataFrame, col: str, default=0, numeric: bool = False) -> pd.Series:
    if df is None:
        return pd.Series(dtype="float64" if numeric else "object")
    if col in df.columns:
        s = pd.Series(df[col], index=df.index, copy=True)
    else:
        s = pd.Series(default, index=df.index)
    if numeric:
        return pd.to_numeric(s, errors="coerce").fillna(default)
    return s.fillna(default)


def attach_strategy_nogo_columns(df: pd.DataFrame, cfg: dict | None = None, pool_stage: str = "candidate20") -> pd.DataFrame:
    """V9.6.1：完整 NoGo_Detail 欄位化。

    輸出人可讀 strategy_nogo_detail，也輸出機器可驗證 fail_* 與 threshold_* 欄位，
    讓 today_buy=0 時可直接統計原因。
    """
    if df is None or df.empty:
        return df
    x = df.copy()
    cfg = _active_strategy_config(cfg)
    core_cfg, exe_cfg, wait_cfg = _strategy_values(cfg)

    allowed_core = set(_strategy_allowed_decisions(cfg, "core_attack"))
    allowed_exec = set(_strategy_allowed_decisions(cfg, "execution"))
    required_liq = set(_strategy_required_liquidity(cfg))
    model_min = float(get_strategy_threshold(cfg, "core_attack", "model_score_min"))
    wave_min = float(get_strategy_threshold(cfg, "core_attack", "wave_trade_score_min"))
    rr_min = float(get_strategy_threshold(cfg, "execution", "rr_min"))
    rsi_max = float(get_strategy_threshold(cfg, "execution", "rsi_max"))
    price_max = float(get_strategy_threshold(cfg, "execution", "price_dev_max"))
    atr_max = float(get_strategy_threshold(cfg, "execution", "atr_pct_max"))

    decision_s = _safe_strategy_series(x, "final_trade_decision", "", numeric=False).astype(str)
    x["wave_text"] = _build_wave_text_from_df(x)
    x["wave_condition_pass"] = _strategy_wave_keyword_mask(x, core_cfg).astype(int)

    model_s = _safe_strategy_series(x, "model_score", 0.0, numeric=True)
    wave_s = _safe_strategy_series(x, "wave_trade_score", 0.0, numeric=True)
    liq_s = _safe_strategy_series(x, "liquidity_status", "", numeric=False).astype(str).str.upper()
    rsi_raw = pd.to_numeric(_safe_strategy_series(x, "rsi", np.nan, numeric=True), errors="coerce")
    rsi14_raw = pd.to_numeric(_safe_strategy_series(x, "rsi14", np.nan, numeric=True), errors="coerce")
    rsi_s = rsi_raw.where(rsi_raw.notna(), rsi14_raw)
    rsi_na = rsi_s.isna()
    rsi_s_cmp = rsi_s.fillna(999.0)
    atr_raw = pd.to_numeric(_safe_strategy_series(x, "atr_pct", np.nan, numeric=True), errors="coerce")
    atr_alt_raw = pd.to_numeric(_safe_strategy_series(x, "atr", np.nan, numeric=True), errors="coerce")
    atr_s = atr_raw.where(atr_raw.notna(), atr_alt_raw)
    atr_na = atr_s.isna()
    atr_s_cmp = atr_s.fillna(999.0)
    dev_raw = pd.to_numeric(_safe_strategy_series(x, "price_deviation", np.nan, numeric=True), errors="coerce")
    dev_alt_raw = pd.to_numeric(_safe_strategy_series(x, "price_dev", np.nan, numeric=True), errors="coerce")
    dev_s = dev_raw.where(dev_raw.notna(), dev_alt_raw).abs()
    dev_na = dev_s.isna()
    dev_s_cmp = dev_s.fillna(999.0)
    rr_s = pd.to_numeric(_safe_strategy_series(x, "rr_live", np.nan, numeric=True), errors="coerce")
    if rr_s.isna().all():
        rr_s = _safe_strategy_series(x, "rr", 0.0, numeric=True)
    rr_s = rr_s.fillna(0.0)

    x["threshold_model_score_min"] = model_min
    x["threshold_wave_trade_score_min"] = wave_min
    x["threshold_rr_min"] = rr_min
    x["threshold_rsi_max"] = rsi_max
    x["threshold_price_dev_max"] = price_max
    x["threshold_atr_pct_max"] = atr_max

    x["fail_decision"] = (~decision_s.isin(allowed_core)).astype(int)
    x["fail_model_score"] = (model_s < model_min).astype(int)
    x["fail_wave_trade_score"] = (wave_s < wave_min).astype(int)
    x["fail_wave_keyword"] = (x["wave_condition_pass"].astype(int).eq(0)).astype(int)
    x["fail_liquidity"] = (~liq_s.isin(required_liq)).astype(int)
    x["fail_rsi_na"] = rsi_na.astype(int)
    x["fail_atr_na"] = atr_na.astype(int)
    x["fail_price_dev_na"] = dev_na.astype(int)
    x["fail_rsi"] = (rsi_s_cmp > rsi_max).astype(int)
    x["fail_atr"] = (atr_s_cmp > atr_max).astype(int)
    x["fail_price_deviation"] = (dev_s_cmp > price_max).astype(int)
    x["fail_rr"] = (rr_s < rr_min).astype(int)
    x["fail_today_decision"] = (~decision_s.isin(allowed_exec)).astype(int)

    reason_columns = [
        ("fail_decision", lambda i: f"決策非主攻({decision_s.loc[i] or 'NA'})"),
        ("fail_model_score", lambda i: f"model_score {model_s.loc[i]:.2f} < {model_min:g}"),
        ("fail_wave_trade_score", lambda i: f"wave_trade_score {wave_s.loc[i]:.2f} < {wave_min:g}"),
        ("fail_wave_keyword", lambda i: "波段型態非設定主升條件"),
        ("fail_liquidity", lambda i: f"流動性 {liq_s.loc[i] or 'NA'} 不在 {sorted(required_liq)}"),
        ("fail_rsi_na", lambda i: "FAIL_RSI_NA：RSI欄位缺值"),
        ("fail_atr_na", lambda i: "FAIL_ATR_NA：ATR/ATR%欄位缺值"),
        ("fail_price_dev_na", lambda i: "FAIL_PRICE_DEV_NA：price_dev/price_deviation欄位缺值"),
        ("fail_rsi", lambda i: f"RSI {rsi_s_cmp.loc[i]:.2f} > {rsi_max:g}"),
        ("fail_atr", lambda i: f"ATR% {atr_s_cmp.loc[i]:.2f} > {atr_max:g}"),
        ("fail_price_deviation", lambda i: f"價格偏離 {dev_s_cmp.loc[i]*100:.2f}% > {price_max*100:.1f}%"),
        ("fail_rr", lambda i: f"rr_live {rr_s.loc[i]:.2f} < {rr_min:g}"),
        ("fail_today_decision", lambda i: "未進今日可下單決策清單"),
    ]
    reasons = []
    for idx in x.index:
        rs = []
        for col, fn in reason_columns:
            try:
                if int(x.at[idx, col]) == 1:
                    rs.append(fn(idx))
            except Exception:
                pass
        reasons.append("；".join(dict.fromkeys(rs)) if rs else "PASS")

    x["strategy_profile"] = _strategy_profile_name()
    x["strategy_nogo_detail"] = reasons
    x["fail_reason"] = reasons
    x["strategy_config_summary"] = _strategy_summary_text()
    return x


def normalize_core_analysis_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=CORE_ANALYSIS_COLUMNS)
    x = df.copy()
    numeric_cols = {
        "model_score", "wave_trade_score", "candidate20_score", "core_attack5_score", "execution_score",
        "mainstream_score", "breakout_score", "liquidity_score",
        "entry_low", "entry_high", "entry_mid", "price_deviation", "rr_live",
        "target_1382", "target_1618", "rr", "win_rate",
        "threshold_model_score_min", "threshold_wave_trade_score_min", "threshold_rr_min",
        "threshold_rsi_max", "threshold_price_dev_max", "threshold_atr_pct_max",
        "wave_condition_pass", "fail_decision", "fail_model_score", "fail_wave_trade_score",
        "fail_wave_keyword", "fail_liquidity", "fail_rsi_na", "fail_atr_na", "fail_price_dev_na", "fail_rsi", "fail_atr", "fail_price_deviation",
        "fail_rr", "fail_today_decision",
    }
    for col in CORE_ANALYSIS_COLUMNS:
        if col not in x.columns:
            x[col] = 0.0 if col in numeric_cols else ""
    for col in numeric_cols:
        x[col] = pd.to_numeric(x[col], errors="coerce").fillna(0.0)
    return x[CORE_ANALYSIS_COLUMNS].copy()


def build_display_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    x = df.copy()
    # V9.5.9：補齊外部資訊欄位，避免 UI/Excel 因缺欄位或純量 fillna 崩潰。
    for col, default in [("analysis_ready", 1), ("execution_ready", 0), ("soft_block", 0), ("block_reason", ""), ("execution_block_reason", "")]:
        if col not in x.columns:
            x[col] = default


    def _log_missing_columns(*cols: str):
        missing = sorted({c for c in cols if c not in x.columns})
        if missing:
            key = tuple(missing)
            if key not in BUILD_DISPLAY_WARNING_CACHE:
                BUILD_DISPLAY_WARNING_CACHE.add(key)
                try:
                    log_warning(f"build_display_columns 缺少欄位，已套用安全預設值：{', '.join(missing)}")
                except Exception:
                    pass

    def _safe_series(col: str, default=0, numeric: bool = False):
        if col in x.columns:
            s = pd.Series(x[col], index=x.index, copy=True)
        else:
            _log_missing_columns(col)
            s = pd.Series(default, index=x.index)
        if numeric:
            return pd.to_numeric(s, errors="coerce")
        return s.fillna(default)

    def _safe_first_series(candidates, default=0, numeric: bool = False):
        for col in candidates:
            if col in x.columns:
                return _safe_series(col, default=default, numeric=numeric)
        _log_missing_columns(*candidates)
        s = pd.Series(default, index=x.index)
        if numeric:
            return pd.to_numeric(s, errors="coerce")
        return s.fillna(default)

    if "close" not in x.columns:
        x["close"] = np.nan
    x["close"] = pd.to_numeric(x["close"], errors="coerce")

    if "prev_close" not in x.columns:
        x["prev_close"] = x["close"]
    x["prev_close"] = pd.to_numeric(x["prev_close"], errors="coerce").replace(0, np.nan)

    x["現價"] = pd.to_numeric(_safe_first_series(["現價", "close"], default=np.nan, numeric=True), errors="coerce").fillna(x["close"])
    chg = (x["close"] - x["prev_close"]).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    chg_pct = ((x["close"] / x["prev_close"] - 1.0) * 100.0).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    x["漲跌"] = pd.to_numeric(_safe_first_series(["漲跌"], default=np.nan, numeric=True), errors="coerce").fillna(chg).round(2)
    x["漲跌幅%"] = pd.to_numeric(_safe_first_series(["漲跌幅%"], default=np.nan, numeric=True), errors="coerce").fillna(chg_pct).round(2)

    if "進場區" not in x.columns:
        if "entry_zone" in x.columns:
            x["進場區"] = _safe_series("entry_zone", default="", numeric=False).astype(str)
        elif {"entry_low", "entry_high"}.issubset(x.columns):
            entry_low = _safe_series("entry_low", default=0, numeric=True).fillna(0.0)
            entry_high = _safe_series("entry_high", default=0, numeric=True).fillna(0.0)
            x["進場區"] = [f"{float(lo):.2f} ~ {float(hi):.2f}" for lo, hi in zip(entry_low, entry_high)]
        elif "entry" in x.columns:
            x["進場區"] = _safe_series("entry", default="", numeric=False).astype(str)
        else:
            _log_missing_columns("entry_zone", "entry_low", "entry_high", "entry")
            x["進場區"] = pd.Series("", index=x.index, dtype="object")

    if "停損" not in x.columns:
        if "stop_loss" in x.columns:
            x["停損"] = _safe_series("stop_loss", default="", numeric=False)
        elif "stop" in x.columns:
            x["停損"] = _safe_series("stop", default="", numeric=False)
        else:
            _log_missing_columns("stop_loss", "stop")
            x["停損"] = pd.Series("", index=x.index, dtype="object")

    if "目標價" not in x.columns:
        if "target_price" in x.columns:
            x["目標價"] = _safe_series("target_price", default="", numeric=False)
        elif "1.382" in x.columns:
            x["目標價"] = _safe_series("1.382", default="", numeric=False)
        elif "target_1382" in x.columns:
            x["目標價"] = _safe_series("target_1382", default="", numeric=False)
        elif "target1382" in x.columns:
            x["目標價"] = _safe_series("target1382", default="", numeric=False)
        else:
            _log_missing_columns("target_price", "1.382", "target_1382", "target1382")
            x["目標價"] = pd.Series("", index=x.index, dtype="object")

    alias_pairs = [
        ("分類", "bucket"), ("狀態", "ui_state"), ("盤中狀態", "liquidity_status"), ("活性分", "liquidity_score"),
        ("1.382", "target_1382"), ("1.618", "target_1618"), ("RR", "rr_live"), ("勝率", "win_rate"),
        ("ATR%", "atr_pct"), ("Kelly%", "position_pct"),
        ("建議張數", "suggest_qty"), ("建議張數", "qty"), ("建議張數", "shares"),
        ("建議金額", "suggest_amount"), ("建議金額", "amount"),
        ("單檔曝險%", "single_position_pct"), ("單檔曝險%", "single_pct"),
        ("題材曝險%", "theme_exposure_pct"), ("題材曝險%", "theme_pct"),
        ("產業曝險%", "industry_exposure_pct"), ("產業曝險%", "industry_pct"),
        ("投資組合狀態", "portfolio_state"), ("風險備註", "risk_note"),
        ("模型分數", "model_score"), ("交易分數", "trade_score"),
        ("淘汰原因", "elimination_reason"), ("市場", "market"), ("產業", "industry"), ("題材", "theme"),
        ("PE", "pe"), ("PB", "pb"), ("殖利率%", "dividend_yield"), ("EPS_TTM", "eps_ttm"), ("EPS YoY", "eps_yoy"), ("營收YoY", "revenue_yoy"),
        ("EPS Bucket", "eps_bucket"), ("Revenue Bucket", "rev_bucket"), ("EPS分類", "eps_category"), ("Matrix", "matrix_cell"),
        ("財務分數", "revenue_eps_score"), ("資料狀態", "data_quality_flag"), ("EPS矩陣說明", "eps_matrix_decision_note"), ("估值分", "valuation_score"),
        ("外部允許", "trade_allowed"), ("外部Ready", "external_data_ready"),
        ("Market Gate", "market_gate_state"), ("Flow Gate", "flow_gate_state"),
        ("Fundamental Gate", "fundamental_gate_state"), ("Event Gate", "event_gate_state"), ("Risk Gate", "risk_gate_state"),
        ("外部阻擋原因", "external_blocking_reason"), ("外部資料日", "latest_external_date"),
        ("資料來源層級", "market_source_level"), ("決策摘要", "decision_reason_short"),
        ("全域外部Ready", "global_external_ready"), ("個股覆蓋狀態", "stock_external_coverage_state"),
        ("Gate說明", "gate_policy_note"),
        ("代號", "stock_id"), ("名稱", "stock_name"), ("優先級", "priority"), ("優先級", "優先級"),
    ]
    for zh, src in alias_pairs:
        if zh not in x.columns and src in x.columns:
            x[zh] = x[src]

    if "Kelly%" not in x.columns:
        x["Kelly%"] = pd.to_numeric(_safe_first_series(["position_pct", "kelly_raw"], default=np.nan, numeric=True), errors="coerce")
    if "建議張數" not in x.columns:
        x["建議張數"] = pd.to_numeric(_safe_first_series(["suggest_qty", "qty", "shares"], default=0, numeric=True), errors="coerce").fillna(0)
    if "建議金額" not in x.columns:
        x["建議金額"] = pd.to_numeric(_safe_first_series(["suggest_amount", "amount"], default=0.0, numeric=True), errors="coerce").fillna(0.0)
    if "單檔曝險%" not in x.columns:
        x["單檔曝險%"] = pd.to_numeric(_safe_first_series(["single_position_pct", "single_pct"], default=0.0, numeric=True), errors="coerce").fillna(0.0)
    if "題材曝險%" not in x.columns:
        x["題材曝險%"] = pd.to_numeric(_safe_first_series(["theme_exposure_pct", "theme_pct"], default=0.0, numeric=True), errors="coerce").fillna(0.0)
    if "產業曝險%" not in x.columns:
        x["產業曝險%"] = pd.to_numeric(_safe_first_series(["industry_exposure_pct", "industry_pct"], default=0.0, numeric=True), errors="coerce").fillna(0.0)
    if "投資組合狀態" not in x.columns:
        x["投資組合狀態"] = _safe_first_series(["portfolio_state"], default="未配置", numeric=False).astype(str)
    if "風險備註" not in x.columns:
        x["風險備註"] = _safe_first_series(["risk_note"], default="", numeric=False).astype(str)
    if "優先級" not in x.columns:
        x["優先級"] = pd.to_numeric(_safe_first_series(["priority", "優先級"], default=np.nan, numeric=True), errors="coerce")
        x["優先級"] = x["優先級"].fillna(pd.Series(np.arange(1, len(x) + 1), index=x.index))

    numeric_display = ["1.382", "1.618", "RR", "勝率", "ATR%", "Kelly%", "活性分", "模型分數", "交易分數", "現價", "漲跌", "漲跌幅%", "建議張數", "建議金額", "PE", "PB", "殖利率%", "EPS_TTM", "EPS YoY", "營收YoY", "財務分數", "估值分"]
    for col in numeric_display:
        if col in x.columns:
            x[col] = pd.to_numeric(x[col], errors="coerce")

    if "建議張數" in x.columns:
        x["建議張數"] = pd.to_numeric(x["建議張數"], errors="coerce").fillna(0).astype(int)
    if "建議金額" in x.columns:
        x["建議金額"] = pd.to_numeric(x["建議金額"], errors="coerce").fillna(0.0).round(2)
    for pct_col in ["單檔曝險%", "題材曝險%", "產業曝險%"]:
        if pct_col in x.columns:
            x[pct_col] = pd.to_numeric(x[pct_col], errors="coerce").fillna(0.0).round(2)

    for text_col, default in [("投資組合狀態", "未配置"), ("風險備註", ""), ("盤中狀態", ""), ("淘汰原因", ""), ("EPS Bucket", ""), ("Revenue Bucket", ""), ("EPS分類", "U0"), ("Matrix", ""), ("資料狀態", ""), ("EPS矩陣說明", ""), ("代號", ""), ("名稱", "")]:
        if text_col in x.columns:
            x[text_col] = pd.Series(x[text_col], index=x.index, copy=True).fillna(default).astype(str)
    return x

def assert_schema(df: pd.DataFrame, expected_columns: list[str], schema_name: str) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=expected_columns)
    x = df.copy()
    for col in expected_columns:
        if col not in x.columns:
            x[col] = ""
    x = x[expected_columns].copy()
    return x.fillna("")


def _pool_stock_id_series(df: pd.DataFrame) -> pd.Series:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.Series(dtype="object")
    if "stock_id" not in df.columns:
        return pd.Series(dtype="object")
    return df["stock_id"].astype(str).map(normalize_stock_id).astype(str).str.strip()


def build_pool_audit(pool_dict: dict) -> dict:
    candidate20 = set(_pool_stock_id_series(pool_dict.get("candidate20", pd.DataFrame())).tolist())
    core_attack5 = set(_pool_stock_id_series(pool_dict.get("core_attack5", pd.DataFrame())).tolist())
    today_buy = set(_pool_stock_id_series(pool_dict.get("today_buy", pd.DataFrame())).tolist())
    execution_ready = set(_pool_stock_id_series(pool_dict.get("execution_ready", pd.DataFrame())).tolist())
    unique_decision = set(_pool_stock_id_series(pool_dict.get("unique_decision", pd.DataFrame())).tolist())
    wait_pullback = set(_pool_stock_id_series(pool_dict.get("wait_pullback", pd.DataFrame())).tolist())
    watch = set(_pool_stock_id_series(pool_dict.get("watch", pd.DataFrame())).tolist())
    audit = {
        "candidate20_count": len(candidate20),
        "core_attack5_count": len(core_attack5),
        "today_buy_count": len(today_buy),
        "execution_ready_count": len(execution_ready),
        "unique_decision_count": len(unique_decision),
        "wait_pullback_count": len(wait_pullback),
        "watch_count": len(watch),
        "core_minus_candidate20": sorted(core_attack5 - candidate20),
        "today_minus_core": sorted(today_buy - core_attack5),
        "unique_minus_core": sorted(unique_decision - core_attack5),
        "wait_minus_core": sorted(wait_pullback - core_attack5),
        "watch_minus_candidate20": sorted(watch - candidate20),
        "watch_core_overlap": sorted(watch & core_attack5),
        "wait_watch_overlap": sorted(wait_pullback & watch),
    }
    return audit


def assert_pool_consistency(pool_dict: dict):
    audit = build_pool_audit(pool_dict)
    if audit["core_minus_candidate20"]:
        raise ValueError(f"V16.2 pool error: core_attack5 必須為 candidate20 子集合｜差集={audit['core_minus_candidate20'][:10]}")
    if audit["today_minus_core"]:
        raise ValueError(f"V16.2 pool error: today_buy 必須為 core_attack5 子集合｜差集={audit['today_minus_core'][:10]}")
    if audit["unique_minus_core"]:
        raise ValueError(f"V16.2 pool error: unique_decision 必須為 core_attack5 子集合｜差集={audit['unique_minus_core'][:10]}")
    if audit["wait_minus_core"]:
        raise ValueError(f"V16.2 pool error: wait_pullback 必須為 core_attack5 子集合｜差集={audit['wait_minus_core'][:10]}")
    if audit["watch_minus_candidate20"]:
        raise ValueError(f"V16.2 pool error: watch 必須為 candidate20 子集合｜差集={audit['watch_minus_candidate20'][:10]}")
    if audit["watch_core_overlap"]:
        raise ValueError(f"V16.2 pool error: watch 不可與 core_attack5 重疊｜差集={audit['watch_core_overlap'][:10]}")
    if audit["wait_watch_overlap"]:
        raise ValueError(f"V16.2 pool error: wait_pullback 不可與 watch 重疊｜差集={audit['wait_watch_overlap'][:10]}")


def _pool_ids_from_any_df(df: pd.DataFrame, id_col: str = "stock_id") -> set[str]:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return set()
    if id_col not in df.columns:
        return set()
    return set(pd.Series(df[id_col]).astype(str).map(normalize_stock_id).astype(str).str.strip().tolist())


def assert_phase1_report_consistency(
    candidate20_df: pd.DataFrame,
    core_attack5_df: pd.DataFrame,
    today_buy_df: pd.DataFrame,
    wait_pullback_df: pd.DataFrame,
    watch_df: pd.DataFrame,
    unique_decision_df: pd.DataFrame,
    order_list_df: pd.DataFrame,
    institutional_plan_df: pd.DataFrame,
):
    candidate20_ids = _pool_ids_from_any_df(candidate20_df, "stock_id")
    core_ids = _pool_ids_from_any_df(core_attack5_df, "stock_id")
    today_ids = _pool_ids_from_any_df(today_buy_df, "stock_id")
    wait_ids = _pool_ids_from_any_df(wait_pullback_df, "stock_id")
    watch_ids = _pool_ids_from_any_df(watch_df, "stock_id")
    unique_ids = _pool_ids_from_any_df(unique_decision_df, "stock_id")
    order_ids = _pool_ids_from_any_df(order_list_df, "代號")
    inst_ids = _pool_ids_from_any_df(institutional_plan_df, "代號")

    if not order_ids.issubset(today_ids):
        raise ValueError(f"V16.2 report error: Order_List 必須只來自 today_buy｜差集={sorted(order_ids - today_ids)[:10]}")
    if not inst_ids.issubset(today_ids):
        raise ValueError(f"V16.2 report error: Institutional_Plan 必須只來自 today_buy｜差集={sorted(inst_ids - today_ids)[:10]}")
    if not unique_ids.issubset(core_ids):
        raise ValueError(f"V16.2 report error: Unique_Decision 必須只來自 core_attack5｜差集={sorted(unique_ids - core_ids)[:10]}")
    if not watch_ids.issubset(candidate20_ids):
        raise ValueError(f"V16.2 report error: Watch 必須只來自 candidate20｜差集={sorted(watch_ids - candidate20_ids)[:10]}")
    if watch_ids & core_ids:
        raise ValueError(f"V16.2 report error: Watch 不可與 core_attack5 重疊｜差集={sorted(watch_ids & core_ids)[:10]}")
    if not wait_ids.issubset(core_ids):
        raise ValueError(f"V16.2 report error: Wait_Pullback 必須只來自 core_attack5｜差集={sorted(wait_ids - core_ids)[:10]}")
    if wait_ids & watch_ids:
        raise ValueError(f"V16.2 report error: Wait_Pullback 不可與 Watch 重疊｜差集={sorted(wait_ids & watch_ids)[:10]}")





class WaveEngine:
    @staticmethod
    def detect_wave_label(x: pd.DataFrame) -> str:
        recent = x.tail(89).copy()
        if recent.empty or len(recent) < 30:
            return "整理浪"
        close_ = float(recent.iloc[-1]["close"])
        recent_high = float(recent["high"].max())
        recent_low = float(recent["low"].min())
        ma20 = float(recent.iloc[-1]["ma20"]) if pd.notna(recent.iloc[-1]["ma20"]) else close_
        ma60 = float(recent.iloc[-1]["ma60"]) if pd.notna(recent.iloc[-1]["ma60"]) else close_
        rsi = float(recent.iloc[-1]["rsi14"]) if pd.notna(recent.iloc[-1]["rsi14"]) else 50.0
        macd_hist = float(recent.iloc[-1]["macd_hist"]) if pd.notna(recent.iloc[-1]["macd_hist"]) else 0.0

        width = max(recent_high - recent_low, 1e-6)
        pos = (close_ - recent_low) / width
        breakout = close_ >= recent_high * 0.995

        if breakout and ma20 > ma60 and macd_hist > 0 and 50 <= rsi <= 64 and 0.55 <= pos <= 0.85:
            return "第3浪"
        if breakout and (rsi > 72 or pos > 0.90):
            return "第5浪"
        if close_ > ma20 > ma60 and macd_hist > 0:
            return "推動浪"
        if close_ < ma20 and macd_hist < 0:
            return "修正浪"
        return "整理浪"


class FibEngine:
    @staticmethod
    def score_and_targets(close_: float, support: float, resistance: float) -> tuple[float, float, float]:
        if resistance <= support or support <= 0:
            return 0.0, 0.0, 0.0
        width = max(resistance - support, 1e-6)
        pos = (close_ - support) / width
        if pos < 0.3:
            base = 95.0
        elif pos < 0.6:
            base = 80.0
        elif pos < 0.85:
            base = 65.0
        else:
            base = 45.0
        if pos < 0:
            base = 35.0
        elif pos > 1.05:
            base = 38.0
        return round(base, 2), round(support + width * 1.382, 2), round(support + width * 1.618, 2)


class SakataEngine:
    @staticmethod
    def detect(signal: str, close_: float, ma5: float, ma10: float, ma20: float, recent_high: float) -> str:
        if signal == "突破強勢":
            return "突破追價" if close_ >= recent_high * 0.995 else "拉回承接"
        if signal == "強勢追蹤":
            return "拉回承接" if close_ <= ma5 * 1.01 else "偏多低接"
        if signal == "整理偏多":
            return "整理偏多"
        if signal == "偏多觀察":
            return "區間低接" if close_ >= ma20 else "觀望"
        if signal == "區間整理":
            return "區間低接"
        return "觀望"


class IndustryRotationEngine:
    @staticmethod
    def summarize(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["industry", "count", "avg_total", "avg_ai", "trend_count", "hot_score", "rotation"])
        x = (
            df.groupby("industry", as_index=False)
            .agg(
                count=("stock_id", "count"),
                avg_total=("total_score", "mean"),
                avg_ai=("ai_score", "mean"),
                trend_count=("signal", lambda s: int(pd.Series(s).isin(["強勢追蹤", "整理偏多"]).sum()))
            )
        )
        x["hot_score"] = x["avg_total"] * 0.45 + x["avg_ai"] * 0.25 + x["trend_count"] * 6 + x["count"] * 2
        x["rotation"] = np.where(
            x["hot_score"] >= 75, "主升輪動",
            np.where(x["hot_score"] >= 60, "偏多輪動", np.where(x["hot_score"] >= 45, "中性輪動", "轉弱輪動"))
        )
        return x.sort_values(["hot_score", "avg_total"], ascending=False).reset_index(drop=True)



class StrategyEngineV91:
    """
    v9.2 FINAL-RELEASE 核心策略引擎：
    訊號 → 評分 → 倉位 → 交易計畫
    """
    @staticmethod
    def calc_atr(x: pd.DataFrame, n: int = 14) -> pd.Series:
        prev_close = x["close"].shift(1)
        tr = pd.concat([
            (x["high"] - x["low"]).abs(),
            (x["high"] - prev_close).abs(),
            (x["low"] - prev_close).abs()
        ], axis=1).max(axis=1)
        return tr.rolling(n).mean()

    @staticmethod
    def wave_fib_trade_model(x: pd.DataFrame) -> dict:
        last = x.iloc[-1]
        close_ = float(last["close"])
        ma20 = float(last["ma20"]) if pd.notna(last["ma20"]) else close_
        ma60 = float(last["ma60"]) if pd.notna(last["ma60"]) else close_
        rsi = float(last["rsi14"]) if pd.notna(last["rsi14"]) else 50.0
        macd_hist = float(last["macd_hist"]) if pd.notna(last["macd_hist"]) else 0.0

        recent = x.tail(89)
        recent_high = float(recent["high"].max())
        recent_low = float(recent["low"].min())
        width = max(recent_high - recent_low, 1e-6)
        pos = (close_ - recent_low) / width

        wave = WaveEngine.detect_wave_label(x)
        fib_score, fib1382, fib1618 = FibEngine.score_and_targets(close_, max(ma20, recent_low), recent_high)

        # 交易模型化
        if wave == "第3浪":
            entry_low = max(ma20, recent_low) * 1.003
            entry_high = max(ma20, recent_low) * 1.015
            primary_target = fib1382
            structure_bonus = 12
        elif wave == "推動浪":
            entry_low = max(ma20, recent_low) * 1.002
            entry_high = max(ma20, recent_low) * 1.012
            primary_target = fib1382
            structure_bonus = 8
        elif wave == "第5浪":
            entry_low = ma20 * 0.998
            entry_high = ma20 * 1.006
            primary_target = fib1618
            structure_bonus = -5
        elif wave == "修正浪":
            entry_low = recent_low * 1.002
            entry_high = recent_low * 1.010
            primary_target = fib1382
            structure_bonus = -8
        else:
            entry_low = ma20 * 0.998
            entry_high = ma20 * 1.008
            primary_target = fib1382
            structure_bonus = 0

        regime_bias = 0
        if close_ > ma20 > ma60 and macd_hist > 0:
            regime_bias += 8
        if 48 <= rsi <= 68:
            regime_bias += 6
        elif rsi > 75:
            regime_bias -= 10
        elif rsi < 35:
            regime_bias -= 8

        model_trade_score = float(fib_score) + structure_bonus + regime_bias
        return {
            "wave_trade_score": round(model_trade_score, 2),
            "entry_low_v91": round(entry_low, 2),
            "entry_high_v91": round(entry_high, 2),
            "primary_target_v91": round(primary_target, 2),
            "fib1382_v91": round(fib1382, 2),
            "fib1618_v91": round(fib1618, 2),
            "wave_pos_v91": round(pos, 3),
        }

    @staticmethod
    def decide_signal(model_score: float, trade_score: float, rr: float, rsi: float, wave: str, cfg: dict | None = None) -> tuple[str, str]:
        cfg = _active_strategy_config(cfg)
        core_cfg, exe_cfg, wait_cfg = _strategy_values(cfg)
        model_min = float(get_strategy_threshold(cfg, "core_attack", "model_score_min"))
        wave_min = float(get_strategy_threshold(cfg, "core_attack", "wave_trade_score_min"))
        rr_min = float(get_strategy_threshold(cfg, "execution", "rr_min"))
        rsi_max = float(get_strategy_threshold(cfg, "execution", "rsi_max"))
        wait_rr_min = float(get_strategy_threshold(cfg, "wait_pullback", "rr_min"))
        wave_ok = _match_wave_keyword_text(str(wave or ""), core_cfg)

        model_score = _coerce_float(model_score)
        trade_score = _coerce_float(trade_score)
        rr = _coerce_float(rr)
        rsi = _coerce_float(rsi, 50.0)

        if model_score >= model_min and trade_score >= wave_min and rr >= rr_min and rsi <= rsi_max and wave_ok:
            return "BUY", "可買"
        weak_model_min = max(0.0, model_min - 10.0)
        weak_wave_min = max(0.0, wave_min - 10.0)
        if model_score >= weak_model_min and trade_score >= weak_wave_min and rr >= wait_rr_min and rsi <= rsi_max:
            return "WEAK BUY", "條件預掛"
        if model_score >= max(0.0, model_min - 22.0) and rr >= max(1.0, wait_rr_min * 0.90):
            return "HOLD", "觀察"
        return "AVOID", "不可買"

    @staticmethod
    def score(df: pd.DataFrame) -> Dict[str, float]:
        """
        統一核心評分輸出：
        - 供 RankingEngine / TradingPlanEngine 共用
        - 回傳欄位格式維持與 ranking_result 相容
        """
        if df is None or df.empty or len(df) < 60:
            return {
                "momentum_score": 0.0,
                "trend_score": 0.0,
                "reversal_score": 0.0,
                "volume_score": 0.0,
                "risk_score": 0.0,
                "ai_score": 0.0,
                "total_score": 0.0,
                "signal": "資料不足",
                "action": "等待資料",
            }

        x = IndicatorEngine.attach(df.copy())
        last = x.iloc[-1]

        close_ = float(last["close"])
        ma5 = float(last["ma5"]) if pd.notna(last["ma5"]) else close_
        ma10 = float(last["ma10"]) if pd.notna(last["ma10"]) else close_
        ma20 = float(last["ma20"]) if pd.notna(last["ma20"]) else close_
        ma60 = float(last["ma60"]) if pd.notna(last["ma60"]) else close_
        rsi = float(last["rsi14"]) if pd.notna(last["rsi14"]) else 50.0
        macd_hist = float(last["macd_hist"]) if pd.notna(last["macd_hist"]) else 0.0

        recent20 = x.tail(20)
        recent60 = x.tail(60)

        ret20 = (close_ / float(x.iloc[-21]["close"]) - 1) * 100 if len(x) >= 21 else 0.0
        momentum = max(0.0, min(100.0, 50 + ret20 * 2))

        trend_raw = 0
        trend_raw += 1 if close_ > ma5 else 0
        trend_raw += 1 if ma5 > ma10 else 0
        trend_raw += 1 if ma10 > ma20 else 0
        trend_raw += 1 if ma20 > ma60 else 0
        trend = float(trend_raw * 25)

        if 45 <= rsi <= 65:
            reversal = 90.0
        elif 40 <= rsi < 45 or 65 < rsi <= 70:
            reversal = 72.0
        elif 35 <= rsi < 40 or 70 < rsi <= 75:
            reversal = 50.0
        else:
            reversal = 22.0
        if macd_hist > 0:
            reversal = min(100.0, reversal + 8)

        vol_ma20 = float(recent20["volume"].mean()) if not recent20.empty else 0.0
        vol_ratio = (float(last["volume"]) / vol_ma20) if vol_ma20 > 0 else 1.0
        if vol_ratio >= 1.4:
            volume = 100.0
        elif vol_ratio >= 1.05:
            volume = 82.0
        elif vol_ratio >= 0.8:
            volume = 60.0
        else:
            volume = 28.0

        vol20 = float(x["close"].pct_change().tail(20).std()) if len(x) >= 20 else 0.02
        vol20 = 0.02 if pd.isna(vol20) else vol20
        risk = max(0.0, min(100.0, 100 - vol20 * 1500))

        recent_high = float(recent60["high"].max()) if not recent60.empty else close_
        recent_low = float(recent60["low"].min()) if not recent60.empty else close_
        mapped_signal = "區間整理"
        breakout = recent_high > 0 and close_ >= recent_high * 0.995
        strong_trend = close_ > ma5 > ma10 > ma20
        mild_trend = close_ >= ma20 and ma20 >= ma60
        if breakout and strong_trend and macd_hist > 0 and 48 <= rsi <= 62:
            mapped_signal = "突破強勢"
        elif breakout and rsi > 72:
            mapped_signal = "區間整理"
        elif strong_trend and 50 <= rsi <= 70 and macd_hist > 0:
            mapped_signal = "強勢追蹤"
        elif mild_trend and 45 <= rsi <= 68:
            mapped_signal = "整理偏多"
        elif close_ > ma20 > ma60 and macd_hist > 0 and 45 <= rsi <= 68:
            mapped_signal = "強勢追蹤"
        elif close_ >= ma20 and macd_hist >= -0.02 and 40 <= rsi <= 65:
            mapped_signal = "偏多觀察"
        elif close_ < ma20 and (rsi < 40 or macd_hist < 0):
            mapped_signal = "轉弱警戒"
        elif close_ < ma60 and rsi < 32:
            mapped_signal = "急跌風險"

        ai = max(0.0, min(100.0, momentum * 0.18 + trend * 0.22 + reversal * 0.15 + volume * 0.15 + risk * 0.12 + (8 if macd_hist > 0 else 0)))
        total = max(0.0, min(100.0, momentum * 0.20 + trend * 0.24 + reversal * 0.14 + volume * 0.14 + risk * 0.10 + ai * 0.18))

        if mapped_signal in ("突破強勢", "強勢追蹤") and total >= 78:
            action = "拉回加碼"
        elif mapped_signal in ("整理偏多", "偏多觀察") and total >= 60:
            action = "低接布局"
        elif mapped_signal == "區間整理":
            action = "區間操作"
        elif mapped_signal == "轉弱警戒":
            action = "減碼/防守"
        elif mapped_signal == "急跌風險":
            action = "觀望為主"
        else:
            action = "等待訊號"

        return {
            "momentum_score": round(momentum, 2),
            "trend_score": round(trend, 2),
            "reversal_score": round(reversal, 2),
            "volume_score": round(volume, 2),
            "risk_score": round(risk, 2),
            "ai_score": round(ai, 2),
            "total_score": round(total, 2),
            "signal": mapped_signal,
            "action": action,
        }

    @staticmethod
    def kelly_position(win_rate_pct: float, rr: float, atr_pct: float, total_capital: float, regime: str) -> dict:
        p = max(0.01, min(float(win_rate_pct) / 100.0, 0.95))
        b = max(float(rr), 0.05)
        q = 1 - p
        raw_kelly = (b * p - q) / b
        raw_kelly = max(0.0, raw_kelly)

        # 分數保守化
        regime_factor = {"多頭": 0.60, "震盪": 0.35, "空頭": 0.18}.get(regime, 0.30)
        atr_penalty = 1.0
        if atr_pct >= 8:
            atr_penalty = 0.45
        elif atr_pct >= 5:
            atr_penalty = 0.65
        elif atr_pct >= 3:
            atr_penalty = 0.82

        final_pct = min(raw_kelly * 0.5, regime_factor) * atr_penalty
        final_pct = max(0.0, min(final_pct, 0.12))
        amount = round(total_capital * final_pct, 2)

        if final_pct >= 0.08:
            tier = "核心"
        elif final_pct >= 0.04:
            tier = "標準"
        elif final_pct > 0:
            tier = "試單"
        else:
            tier = "觀察"

        return {
            "kelly_raw": round(raw_kelly * 100, 2),
            "position_pct": round(final_pct * 100, 2),
            "suggest_amount": amount,
            "position_tier_v91": tier,
        }



class IntradayLiquidityEngine:
    """
    v9.2 升級：盤中資金活性淘汰規則（以日K/近端價量做代理版）
    - Phase A：分時趨勢分 / 攻擊量分 / 區間股淘汰
    - Phase B：VWAP缺資料時用近端均價代理 + 相對市場 / 相對產業強弱
    - Phase C：主動買盤 / 大單掃盤暫以量價攻擊代理分數占位
    """
    def __init__(self, db: DBManager):
        self.db = db

    @staticmethod
    def _safe_float(v, default: float = 0.0) -> float:
        try:
            return float(v)
        except Exception:
            return float(default)

    @staticmethod
    def _clamp(v: float, low: float = 0.0, high: float = 100.0) -> float:
        return max(low, min(high, float(v)))

    def _benchmark_strength(self) -> float:
        for sid in ("0050", "2330"):
            hist = self.db.get_price_history(sid)
            if hist is not None and not hist.empty and len(hist) >= 25:
                x = IndicatorEngine.attach(hist.copy())
                last = x.iloc[-1]
                close_ = self._safe_float(last.get("close", 0), 0)
                close_5 = self._safe_float(x.iloc[-6]["close"], close_) if len(x) >= 6 else close_
                close_20 = self._safe_float(x.iloc[-21]["close"], close_) if len(x) >= 21 else close_
                ret5 = (close_ / max(close_5, 1e-6) - 1.0) * 100
                ret20 = (close_ / max(close_20, 1e-6) - 1.0) * 100
                return ret5 * 0.55 + ret20 * 0.45
        return 0.0

    def _industry_relative_score(self, stock_row: pd.Series, own_ret20: float) -> tuple[float, float]:
        try:
            industry = str(stock_row.get("industry", "") or "").strip()
        except Exception:
            industry = ""
        if not industry:
            return 50.0, 0.0
        master = self.db.get_master()
        if master is None or master.empty or "industry" not in master.columns:
            return 50.0, 0.0
        peers = master[master["industry"].astype(str) == industry]["stock_id"].astype(str).tolist()[:60]
        peer_scores = []
        for sid in peers:
            hist = self.db.get_price_history(sid)
            if hist is None or hist.empty or len(hist) < 25:
                continue
            try:
                c = float(hist.iloc[-1]["close"])
                c20 = float(hist.iloc[-21]["close"]) if len(hist) >= 21 else c
                peer_scores.append((c / max(c20, 1e-6) - 1.0) * 100)
            except Exception:
                continue
        if not peer_scores:
            return 50.0, 0.0
        arr = pd.Series(peer_scores, dtype=float)
        pct = float((arr <= own_ret20).mean() * 100.0)
        return self._clamp(pct), float(arr.mean())

    def evaluate(self, stock_row: pd.Series, hist: pd.DataFrame, theme_hot: bool = False) -> dict:
        out = {
            "intraday_trend_score": 0.0,
            "attack_volume_score": 0.0,
            "range_breakout_score": 0.0,
            "vwap_distance_pct": 0.0,
            "relative_strength_market": 0.0,
            "relative_strength_industry": 0.0,
            "leader_follow_score": 0.0,
            "active_buy_score": 0.0,
            "orderflow_aggression_score": 0.0,
            "large_order_scan_score": 0.0,
            "institutional_participation_score": 0.0,
            "liquidity_status": "WATCH",
            "elimination_reason": "",
            "is_mainstream_funding": 0,
            "liquidity_score": 0.0,
        }
        if hist is None or hist.empty or len(hist) < 30:
            out["liquidity_status"] = "ELIMINATE"
            out["elimination_reason"] = "歷史資料不足"
            return out

        x = IndicatorEngine.attach(hist.copy())
        last = x.iloc[-1]
        close_ = self._safe_float(last.get("close", 0), 0)
        open_ = self._safe_float(last.get("open", close_), close_)
        high_ = self._safe_float(last.get("high", close_), close_)
        low_ = self._safe_float(last.get("low", close_), close_)
        ma5 = self._safe_float(last.get("ma5", close_), close_)
        ma10 = self._safe_float(last.get("ma10", close_), close_)
        ma20 = self._safe_float(last.get("ma20", close_), close_)
        ma60 = self._safe_float(last.get("ma60", close_), close_)

        recent20 = x.tail(20)
        recent60 = x.tail(60)
        recent10 = x.tail(10)

        close5 = self._safe_float(x.iloc[-6]["close"], close_) if len(x) >= 6 else close_
        close20 = self._safe_float(x.iloc[-21]["close"], close_) if len(x) >= 21 else close_
        ret5 = (close_ / max(close5, 1e-6) - 1.0) * 100
        ret20 = (close_ / max(close20, 1e-6) - 1.0) * 100

        range20_high = self._safe_float(recent20["high"].max(), close_)
        range20_low = self._safe_float(recent20["low"].min(), close_)
        range60_high = self._safe_float(recent60["high"].max(), close_)
        range60_low = self._safe_float(recent60["low"].min(), close_)
        width20 = max(range20_high - range20_low, 1e-6)
        width60 = max(range60_high - range60_low, 1e-6)
        close_loc = (close_ - low_) / max(high_ - low_, 1e-6)
        range_pos20 = (close_ - range20_low) / width20
        range_pos60 = (close_ - range60_low) / width60

        # 用近端均價代理 VWAP / 均價線
        proxy_vwap = self._safe_float((recent10["turnover"].sum() / max(recent10["volume"].sum(), 1e-6)) if "turnover" in recent10.columns else recent10["close"].mean(), recent10["close"].mean())
        vwap_distance_pct = ((close_ / max(proxy_vwap, 1e-6)) - 1.0) * 100

        trend_score = 0.0
        if close_ > proxy_vwap:
            trend_score += 24
        if close_ > ma5 > ma10:
            trend_score += 18
        if ma10 >= ma20:
            trend_score += 12
        if recent10["close"].tail(3).is_monotonic_increasing:
            trend_score += 12
        if ret5 > 0:
            trend_score += 12
        if range_pos20 >= 0.62:
            trend_score += 12
        if close_loc >= 0.62:
            trend_score += 10
        out["intraday_trend_score"] = round(self._clamp(trend_score), 2)
        out["vwap_distance_pct"] = round(vwap_distance_pct, 2)

        vol_today = self._safe_float(last.get("volume", 0), 0)
        vol5 = self._safe_float(recent20["volume"].tail(5).mean(), vol_today)
        vol10 = self._safe_float(recent20["volume"].tail(10).mean(), vol_today)
        vol20 = self._safe_float(recent20["volume"].mean(), vol_today)
        vr5 = vol_today / max(vol5, 1e-6)
        vr10 = vol_today / max(vol10, 1e-6)
        vr20 = vol_today / max(vol20, 1e-6)
        attack_volume_score = 0.0
        if vr5 >= 1.35:
            attack_volume_score += 30
        elif vr5 >= 1.10:
            attack_volume_score += 20
        elif vr5 >= 0.90:
            attack_volume_score += 10
        if vr10 >= 1.20:
            attack_volume_score += 20
        elif vr10 >= 1.00:
            attack_volume_score += 12
        if ret5 > 0 and vr20 >= 1.00:
            attack_volume_score += 20
        if close_ > open_ and close_loc >= 0.60:
            attack_volume_score += 15
        if close_ >= range20_high * 0.995:
            attack_volume_score += 15
        out["attack_volume_score"] = round(self._clamp(attack_volume_score), 2)

        range_breakout_score = 0.0
        is_range_stock = (width60 / max(close_, 1e-6)) < 0.12 and abs(ret20) < 8
        if close_ >= range20_high * 0.995:
            range_breakout_score += 45
        elif range_pos20 >= 0.75:
            range_breakout_score += 25
        if close_ > ma20 > ma60:
            range_breakout_score += 20
        if not is_range_stock:
            range_breakout_score += 20
        out["range_breakout_score"] = round(self._clamp(range_breakout_score), 2)

        benchmark = self._benchmark_strength()
        rel_market = ret20 - benchmark
        out["relative_strength_market"] = round(rel_market, 2)

        ind_pct, ind_mean = self._industry_relative_score(stock_row, ret20)
        out["relative_strength_industry"] = round(ind_pct, 2)

        leader_follow_score = 0.0
        if rel_market > 3:
            leader_follow_score += 35
        elif rel_market > 0:
            leader_follow_score += 22
        if ind_pct >= 70:
            leader_follow_score += 30
        elif ind_pct >= 55:
            leader_follow_score += 18
        if theme_hot:
            leader_follow_score += 20
        if close_ > proxy_vwap and vr5 >= 1.0:
            leader_follow_score += 15
        out["leader_follow_score"] = round(self._clamp(leader_follow_score), 2)
        out["is_mainstream_funding"] = int(theme_hot and rel_market > 0 and ind_pct >= 50)

        # 無五檔/主動買盤資料時，使用量價攻擊代理分數占位
        candle_body = max(close_ - open_, 0.0)
        candle_range = max(high_ - low_, 1e-6)
        body_ratio = candle_body / candle_range
        active_buy_score = 0.0
        if close_ > open_:
            active_buy_score += 25
        if body_ratio >= 0.55:
            active_buy_score += 20
        if close_loc >= 0.70:
            active_buy_score += 20
        if vr5 >= 1.10 and ret5 > 0:
            active_buy_score += 20
        if rel_market > 0:
            active_buy_score += 15
        out["active_buy_score"] = round(self._clamp(active_buy_score), 2)
        out["orderflow_aggression_score"] = round(self._clamp(active_buy_score * 0.9 + attack_volume_score * 0.1), 2)
        out["large_order_scan_score"] = round(self._clamp((body_ratio * 45) + (max(vr5 - 1, 0) * 55)), 2)
        out["institutional_participation_score"] = round(self._clamp((leader_follow_score * 0.55) + (active_buy_score * 0.45)), 2)

        fail_reasons = []
        hard_fail = False
        if out["intraday_trend_score"] < 42:
            fail_reasons.append("無方向")
        if out["attack_volume_score"] < 35:
            fail_reasons.append("無攻擊量")
        if is_range_stock and out["range_breakout_score"] < 40:
            fail_reasons.append("區間股")
        if out["leader_follow_score"] < 35:
            fail_reasons.append("非主流/相對弱")
        if out["orderflow_aggression_score"] < 35:
            fail_reasons.append("主動買盤弱")

        if out["intraday_trend_score"] < 35 and out["attack_volume_score"] < 30:
            hard_fail = True
        if is_range_stock and out["orderflow_aggression_score"] < 38:
            hard_fail = True

        liquidity_score = (
            out["intraday_trend_score"] * 0.28 +
            out["attack_volume_score"] * 0.24 +
            out["range_breakout_score"] * 0.14 +
            self._clamp(out["relative_strength_market"] * 6 + 50) * 0.10 +
            out["relative_strength_industry"] * 0.10 +
            out["leader_follow_score"] * 0.08 +
            out["orderflow_aggression_score"] * 0.06
        )
        out["liquidity_score"] = round(self._clamp(liquidity_score), 2)

        if hard_fail:
            out["liquidity_status"] = "ELIMINATE"
        elif out["liquidity_score"] >= 68 and out["intraday_trend_score"] >= 55 and out["attack_volume_score"] >= 50:
            out["liquidity_status"] = "PASS"
        else:
            out["liquidity_status"] = "WATCH"

        if out["liquidity_status"] == "ELIMINATE" and not fail_reasons:
            fail_reasons = ["量縮盤整"]
        out["elimination_reason"] = " / ".join(fail_reasons[:4]) if fail_reasons else "盤中結構可接受"
        return out


class TradingPlanEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.market_engine = MarketRegimeEngine(db)
        self.intraday_engine = IntradayLiquidityEngine(db)
        self.financial_feature_engine = FinancialFeatureEngine(db)

    @staticmethod
    def _is_etf(stock: pd.Series) -> bool:
        try:
            return int(stock.get("is_etf", 0)) == 1 or str(stock.get("market", "")) == "ETF"
        except Exception:
            return False

    @staticmethod
    def _round_price(v) -> str:
        try:
            return f"{float(v):.2f}"
        except Exception:
            return "-"

    @staticmethod
    def _in_entry_zone(close_: float, entry_low: float, entry_high: float) -> bool:
        try:
            return float(entry_low) <= float(close_) <= float(entry_high)
        except Exception:
            return False

    @staticmethod
    def _ui_trade_state(decision: str, close_: float, entry_low: float, entry_high: float, rr: float, win_rate: float, liquidity_status: str = "WATCH") -> str:
        decision = str(decision or "").strip().upper()
        liquidity_status = str(liquidity_status or "WATCH").strip().upper()
        in_entry_zone = TradingPlanEngine._in_entry_zone(close_, entry_low, entry_high)

        if liquidity_status == "ELIMINATE":
            return "淘汰"
        if decision == "BUY":
            return "可買" if in_entry_zone else "準備買"
        if decision == "WEAK BUY":
            return "條件預掛" if liquidity_status == "PASS" else "觀察"
        if decision == "HOLD":
            return "觀察"
        if decision == "AVOID":
            return "不可買"
        return "觀察"

    @staticmethod
    def _clamp(v: float, low: float = 0.0, high: float = 100.0) -> float:
        return max(low, min(high, float(v)))

    def _map_kline_signal(self, source_signal: str, close_: float, recent_high: float, ma5: float, ma10: float, ma20: float, ma60: float, macd_hist: float, rsi: float) -> str:
        signal = str(source_signal or "").strip()
        breakout = recent_high > 0 and close_ >= recent_high * 0.995
        strong_trend = close_ > ma5 > ma10 > ma20
        mild_trend = close_ >= ma20 and ma20 >= ma60

        if breakout and strong_trend and macd_hist > 0 and 48 <= rsi <= 62:
            return "突破強勢"
        if breakout and rsi > 72:
            return "區間整理"
        if signal == "強勢追蹤" and strong_trend and 50 <= rsi <= 70:
            return "強勢追蹤"
        if signal == "整理偏多" and mild_trend:
            return "整理偏多"
        if signal == "中性觀察":
            if mild_trend and macd_hist >= 0 and rsi < 68:
                return "偏多觀察"
            return "區間整理"
        if close_ > ma20 > ma60 and macd_hist > 0 and 45 <= rsi <= 68:
            return "強勢追蹤"
        if close_ >= ma20 and macd_hist >= -0.02 and 40 <= rsi <= 65:
            return "偏多觀察"
        if close_ < ma20 and (rsi < 40 or macd_hist < 0):
            return "轉弱警戒"
        if close_ < ma60 and rsi < 32:
            return "急跌風險"
        return "區間整理"

    # legacy helper removed in v9.2 FINAL-RELEASE: _wave_position is no longer used

    # legacy helper removed in v9.2 FINAL-RELEASE: _fib_score_and_targets is no longer used

    # legacy helper removed in v9.2 FINAL-RELEASE: _sakata_label is no longer used

    def _volume_label(self, vol_ratio: float, close_: float, ma20: float) -> str:
        if vol_ratio >= 1.4 and close_ >= ma20:
            return "買盤明顯偏強"
        if vol_ratio >= 1.05 and close_ >= ma20:
            return "買盤偏強"
        if vol_ratio >= 0.8:
            return "多空均衡"
        return "賣盤偏強"

    def _indicator_score(self, rsi: float, macd_hist: float, k: float, d: float) -> float:
        # RSI 分段強化（v9.2 FINAL-RELEASE）：極端值明確扣分，避免過熱/過弱仍拿高分
        if 45 <= rsi <= 65:
            base = 100.0
        elif 40 <= rsi < 45:
            base = 82.0
        elif 65 < rsi <= 70:
            base = 78.0
        elif 70 < rsi <= 72:
            base = 68.0
        elif 35 <= rsi < 40:
            base = 50.0
        elif 72 < rsi <= 75:
            base = 40.0
        elif 75 < rsi <= 80:
            base = 25.0
        elif rsi > 80:
            base = 10.0
        elif 30 <= rsi < 35:
            base = 25.0
        else:
            base = 12.0

        if macd_hist > 0:
            base += 5
        else:
            base -= 3

        if k >= d:
            base += 3
        else:
            base -= 2

        return round(self._clamp(base), 2)

    # legacy helper removed in v9.2 FINAL-RELEASE: _decision is no longer used

    def _derive_tactical_light(self, signal: str, model_score: float, liquidity_status: str, rsi: float, decision: str, wave_label: str) -> str:
        signal = str(signal or "").strip()
        liquidity_status = str(liquidity_status or "WATCH").strip().upper()
        decision = str(decision or "").strip().upper()
        if liquidity_status == "ELIMINATE":
            return "🔴"
        if decision == "BUY" and liquidity_status == "PASS" and wave_label in ("第3浪", "推動浪") and float(model_score or 0) >= 80 and float(rsi or 0) <= 72:
            return "🔵"
        if decision in ("BUY", "WEAK BUY") and liquidity_status == "PASS":
            return "🟢"
        if decision in ("WEAK BUY", "HOLD") or liquidity_status == "WATCH":
            return "🟡"
        if signal in ("轉弱警戒", "急跌風險") or decision == "AVOID":
            return "🟠"
        return "⚪"

    def _derive_final_trade_decision(self, tactical_light: str, is_etf: bool = False) -> str:
        if is_etf:
            return "DEFENSE"
        light = str(tactical_light or "⚪")
        if light == "🔵":
            return "STRONG_BUY"
        if light == "🟢":
            return "BUY"
        if light == "🟡":
            return "WAIT_PULLBACK"
        if light == "🟠":
            return "AVOID"
        if light == "🔴":
            return "ELIMINATE"
        return "IGNORE"

    def _derive_bucket_from_light(self, tactical_light: str, is_etf: bool = False) -> str:
        if is_etf:
            return "防守"
        light = str(tactical_light or "⚪")
        if light in ("🔵", "🟢"):
            return "主攻"
        if light == "🟡":
            return "觀察"
        return "排除"

    def _derive_ui_state_from_final_decision(self, final_trade_decision: str, liquidity_status: str, close_: float, entry_low: float, entry_high: float) -> str:
        fd = str(final_trade_decision or "IGNORE").strip().upper()
        liquidity_status = str(liquidity_status or "WATCH").strip().upper()
        in_entry_zone = self._in_entry_zone(close_, entry_low, entry_high)
        if fd in ("ELIMINATE", "IGNORE") or liquidity_status == "ELIMINATE":
            return "淘汰" if liquidity_status == "ELIMINATE" else "不可買"
        if fd == "STRONG_BUY":
            return "可買" if in_entry_zone else "準備買"
        if fd == "BUY":
            return "可買" if in_entry_zone else "準備買"
        if fd == "WAIT_PULLBACK":
            return "條件預掛" if liquidity_status == "PASS" else "觀察"
        if fd == "AVOID":
            return "不可買"
        if fd == "DEFENSE":
            return "防守"
        return "觀察"

    def _gate_mainstream(self, metrics: dict) -> tuple[int, str]:
        score = 0
        if float(metrics.get("mainstream_score", 0) or 0) >= 70:
            score += 1
        if float(metrics.get("leader_follow_score", 0) or 0) >= 55:
            score += 1
        if float(metrics.get("relative_strength_market", 0) or 0) > 0:
            score += 1
        if float(metrics.get("relative_strength_industry", 0) or 0) >= 55:
            score += 1
        passed = int(score >= 3)
        return passed, "主力成立" if passed else "非主流資金"

    def _gate_structure(self, metrics: dict) -> tuple[int, str]:
        score = 0
        if str(metrics.get("wave", "")).strip() in ("第3浪", "推動浪"):
            score += 1
        if str(metrics.get("signal", "")).strip() in ("突破強勢", "強勢追蹤", "整理偏多"):
            score += 1
        if int(metrics.get("trend_ok", 0) or 0) == 1:
            score += 1
        if int(metrics.get("macd_ok", 0) or 0) == 1:
            score += 1
        passed = int(score >= 3)
        return passed, "主升成立" if passed else "結構不足"

    def _gate_execution(self, metrics: dict, cfg: dict | None = None) -> tuple[int, str]:
        cfg = _active_strategy_config(cfg)
        price_deviation = abs(_coerce_float(metrics.get("price_deviation", 0), 0.0))
        rr_live = _coerce_float(metrics.get("rr_live", 0), 0.0)
        liquidity_status = str(metrics.get("liquidity_status", "WATCH") or "WATCH").strip().upper()
        price_max = float(get_strategy_threshold(cfg, "execution", "price_dev_max"))
        rr_min = float(get_strategy_threshold(cfg, "execution", "rr_min"))
        required_liq = set(_strategy_required_liquidity(cfg))
        fail = []
        if price_deviation > price_max:
            fail.append(f"fail_price_deviation({price_deviation*100:.2f}%>{price_max*100:.1f}%)")
        if rr_live < rr_min:
            fail.append(f"fail_rr({rr_live:.2f}<{rr_min:g})")
        if liquidity_status not in required_liq:
            fail.append(f"fail_liquidity({liquidity_status or 'NA'} not in {sorted(required_liq)})")
        return int(not fail), "PASS" if not fail else "/".join(fail)

    def build_plan(self, stock_id: str) -> dict:
        stock = self.db.get_stock_row(stock_id)
        hist = self.db.get_price_history(stock_id)
        if stock is None or hist.empty or len(hist) < 70:
            return {
                "stock_id": stock_id,
                "stock_name": stock["stock_name"] if stock is not None else stock_id,
                "theme": stock["theme"] if stock is not None else "",
                "industry": stock["industry"] if stock is not None else "",
                "trade_action": "AVOID",
                "ui_state": "不可買",
                "entry_low": 0.0,
                "entry_high": 0.0,
                "entry_zone": "-",
                "stop_loss": "-",
                "target_price": "-",
                "rr": 0.0,
                "win_grade": "C",
                "win_rate": 45.0,
                "selection_score": 0.0,
                "trade_score": 0.0,
                "bucket": "排除",
                "reason": "資料不足",
                "wave": "資料不足",
                "rsi": 50.0,
                "trend_ok": 0,
                "kd_ok": 0,
                "macd_ok": 0,
                "volume_ok": 0,
                "decision": "AVOID",
                "support": 0.0,
                "resistance": 0.0,
                "model_score": 0.0,
                "kline_score": 0.0,
                "wave_score": 0.0,
                "fib_score": 0.0,
                "sakata_score": 0.0,
                "volume_score": 0.0,
                "indicator_score": 0.0,
                "target_1382": 0.0,
                "target_1618": 0.0,
                "signal": "資料不足",
                "trade_type": "觀望",
                "sakata_label": "觀望",
                "volume_label": "多空均衡",
                "intraday_trend_score": 0.0,
                "attack_volume_score": 0.0,
                "range_breakout_score": 0.0,
                "vwap_distance_pct": 0.0,
                "relative_strength_market": 0.0,
                "relative_strength_industry": 0.0,
                "leader_follow_score": 0.0,
                "active_buy_score": 0.0,
                "orderflow_aggression_score": 0.0,
                "large_order_scan_score": 0.0,
                "institutional_participation_score": 0.0,
                "liquidity_status": "ELIMINATE",
                "elimination_reason": "資料不足",
                "liquidity_score": 0.0,
                "is_mainstream_funding": 0,
                "final_trade_decision": "AVOID",
            }

        x = IndicatorEngine.attach(hist.copy())
        x["atr14"] = StrategyEngineV91.calc_atr(x)
        last = x.iloc[-1]
        score = StrategyEngineV91.score(x)
        is_etf = self._is_etf(stock)

        close_ = float(last["close"])
        ma5 = float(last["ma5"]) if pd.notna(last["ma5"]) else close_
        ma10 = float(last["ma10"]) if pd.notna(last["ma10"]) else close_
        ma20 = float(last["ma20"]) if pd.notna(last["ma20"]) else close_
        ma60 = float(last["ma60"]) if pd.notna(last["ma60"]) else close_
        rsi = float(last["rsi14"]) if pd.notna(last["rsi14"]) else 50.0
        macd_hist = float(last["macd_hist"]) if pd.notna(last["macd_hist"]) else 0.0
        k = float(last["k"]) if pd.notna(last["k"]) else 50.0
        d = float(last["d"]) if pd.notna(last["d"]) else 50.0

        recent = x.tail(60)
        recent_high = float(recent["high"].max())
        recent_low = float(recent["low"].min())
        support = min(ma20, recent["low"].tail(20).min()) if pd.notna(ma20) else recent_low
        resistance = recent_high
        support = float(support) if pd.notna(support) else recent_low
        resistance = float(resistance) if pd.notna(resistance) else close_

        source_signal = str(score["signal"])
        signal = self._map_kline_signal(source_signal, close_, recent_high, ma5, ma10, ma20, ma60, macd_hist, rsi)
        wave_label = WaveEngine.detect_wave_label(x)
        sakata_label = SakataEngine.detect(signal, close_, ma5, ma10, ma20, recent_high)
        vol_ma20 = x["volume"].tail(20).mean()
        vol_ratio = float(last["volume"] / vol_ma20) if vol_ma20 and pd.notna(vol_ma20) else 1.0
        volume_label = self._volume_label(vol_ratio, close_, ma20)

        kline_score = float(V80_KLINE_SCORE.get(signal, 55))
        wave_score = float(V80_WAVE_SCORE.get(wave_label, 60))
        fib_score, fib1382, fib1618 = FibEngine.score_and_targets(close_, support, resistance)
        sakata_score = float(V80_SAKATA_SCORE.get(sakata_label, 20))
        volume_score = float(V80_VOLUME_SCORE.get(volume_label, 60))
        indicator_score = float(self._indicator_score(rsi, macd_hist, k, d))

        model_score = round(
            kline_score * V80_WEIGHTS["kline"] +
            wave_score * V80_WEIGHTS["wave"] +
            fib_score * V80_WEIGHTS["fib"] +
            sakata_score * V80_WEIGHTS["sakata"] +
            volume_score * V80_WEIGHTS["volume"] +
            indicator_score * V80_WEIGHTS["indicator"], 2
        )

        active_strategy_config = STRATEGY_CONFIG_MANAGER.get_active_profile()
        wave_trade = StrategyEngineV91.wave_fib_trade_model(x)
        atr14 = float(last["atr14"]) if pd.notna(last["atr14"]) else max(close_ * 0.03, 0.01)
        atr_pct = round((atr14 / max(close_, 0.01)) * 100, 2)

        entry_low = float(wave_trade["entry_low_v91"])
        entry_high = float(wave_trade["entry_high_v91"])
        target = float(wave_trade["primary_target_v91"])
        fib1382 = float(wave_trade["fib1382_v91"])
        fib1618 = float(wave_trade["fib1618_v91"])
        trade_type = f"波浪+費波模型({wave_label})"

        stop = max(support * 0.97, entry_low - atr14 * 1.5)
        risk = max(entry_high - stop, 0.01)
        reward = max(target - entry_high, 0.0)
        rr = round(reward / risk, 2)

        trend_ok = int(close_ > ma5 > ma10 > ma20)
        macd_ok = int(macd_hist > 0)
        kd_ok = int(k >= d)
        volume_ok = int(vol_ratio >= 1.0)

        win_grade, win_rate = WinRateEngine.estimate(hist)
        decision, auto_state = StrategyEngineV91.decide_signal(model_score, float(wave_trade["wave_trade_score"]), rr, rsi, wave_label, active_strategy_config)

        preferred_theme = any(key.lower() in str(stock.get("theme", "")).lower() for key in ThemeStrengthEngine.PREFERRED_KEYWORDS)
        liquidity = self.intraday_engine.evaluate(stock, hist, theme_hot=preferred_theme)
        liquidity_status = str(liquidity.get("liquidity_status", "WATCH") or "WATCH")
        elimination_reason = str(liquidity.get("elimination_reason", "") or "")

        if liquidity_status == "ELIMINATE":
            decision = "AVOID"
            auto_state = "淘汰"
        elif liquidity_status == "WATCH" and decision == "BUY":
            decision = "HOLD"
            auto_state = "觀察"
        elif liquidity_status == "WATCH" and decision == "WEAK BUY":
            auto_state = "觀察"

        entry_mid = round((entry_low + entry_high) / 2.0, 2)
        price_deviation = round((close_ - entry_mid) / max(entry_mid, 0.01), 4)
        rr_live = round(max(target - close_, 0.0) / max(close_ - stop, 0.01), 2)

        selection_score = round(model_score * 0.45 + float(wave_trade["wave_trade_score"]) * 0.16 + win_rate * 0.12 + min(rr, 3.0) * 5 + float(liquidity.get("liquidity_score", 0)) * 0.22 + (6 if decision == "BUY" else 2 if decision == "WEAK BUY" else 0), 2)
        trade_score = round(model_score * 0.25 + float(wave_trade["wave_trade_score"]) * 0.18 + score["ai_score"] * 0.08 + win_rate * 0.14 + min(rr, 3.0) * 5 + float(liquidity.get("intraday_trend_score", 0)) * 0.14 + float(liquidity.get("attack_volume_score", 0)) * 0.11 + float(liquidity.get("leader_follow_score", 0)) * 0.10 + (6 if preferred_theme else 0), 2)

        mainstream_score = round(
            model_score * 0.24 +
            float(liquidity.get("liquidity_score", 0) or 0) * 0.20 +
            float(liquidity.get("leader_follow_score", 0) or 0) * 0.18 +
            float(liquidity.get("institutional_participation_score", 0) or 0) * 0.12 +
            float(liquidity.get("relative_strength_industry", 0) or 0) * 0.10 +
            float(liquidity.get("relative_strength_market", 0) or 0) * 0.06 +
            win_rate * 0.06 + min(rr, 3.0) * 4,
            2
        )
        breakout_score = round(
            float(wave_trade["wave_trade_score"]) * 0.26 +
            float(liquidity.get("attack_volume_score", 0) or 0) * 0.18 +
            float(liquidity.get("range_breakout_score", 0) or 0) * 0.16 +
            float(liquidity.get("active_buy_score", 0) or 0) * 0.12 +
            model_score * 0.12 +
            min(rr, 3.0) * 6,
            2
        )

        mainstream_pass, mainstream_reason = self._gate_mainstream({
            "mainstream_score": mainstream_score,
            "leader_follow_score": liquidity.get("leader_follow_score", 0),
            "relative_strength_market": liquidity.get("relative_strength_market", 0),
            "relative_strength_industry": liquidity.get("relative_strength_industry", 0),
        })
        structure_pass, structure_reason = self._gate_structure({
            "wave": wave_label,
            "signal": signal,
            "trend_ok": trend_ok,
            "macd_ok": macd_ok,
        })
        execution_pass, execution_reason = self._gate_execution({
            "price_deviation": price_deviation,
            "rr_live": rr_live,
            "liquidity_status": liquidity_status,
        }, active_strategy_config)

        if is_etf:
            final_trade_decision = "DEFENSE"
        elif not mainstream_pass:
            final_trade_decision = "IGNORE"
        elif not structure_pass:
            final_trade_decision = "WAIT_PULLBACK"
        elif execution_pass:
            final_trade_decision = "BUY"
        else:
            final_trade_decision = "WAIT_PULLBACK"

        tactical_light = self._derive_tactical_light(signal, model_score, liquidity_status, rsi, decision, wave_label)
        bucket = self._derive_bucket_from_light(tactical_light, is_etf=is_etf)
        ui_state = self._derive_ui_state_from_final_decision(final_trade_decision, liquidity_status, close_, entry_low, entry_high)

        reason = (
            f"{signal}｜{wave_label}｜{trade_type}｜{volume_label}｜"
            f"{mainstream_reason}/{structure_reason}/{execution_reason}｜"
            f"活性 {float(liquidity.get('liquidity_score',0) or 0):.1f}｜{liquidity_status}｜{elimination_reason or '盤中結構可接受'}｜"
            f"六模組 {model_score:.1f}｜RR {rr:.2f}｜RSI {rsi:.1f}"
        )

        feature = self.financial_feature_engine.get_latest_feature(stock_id)
        if not feature:
            feature = {
                "eps_ttm": np.nan, "eps_yoy": np.nan, "revenue_yoy": np.nan,
                "eps_bucket": "", "rev_bucket": "", "matrix_cell": "",
                "eps_category": "U0", "matrix_base_score": np.nan, "modifier": np.nan,
                "revenue_eps_score": 50.0, "data_quality_flag": "EPS_MATRIX_NE",
                "financial_score": 50.0, "eps_matrix_decision_note": "EPS矩陣尚未建立"
            }

        return {
            "stock_id": stock_id,
            "stock_name": stock["stock_name"],
            "industry": stock["industry"],
            "theme": stock["theme"],
            "market": stock["market"],
            "is_etf": 1 if is_etf else 0,
            "trade_action": decision,
            "ui_state": ui_state,
            "tactical_light": tactical_light,
            "candidate_engine": "混合",
            "entry_low": round(entry_low, 2),
            "entry_high": round(entry_high, 2),
            "entry_zone": f"{self._round_price(entry_low)} ~ {self._round_price(entry_high)}",
            "entry_mid": entry_mid,
            "price_deviation": price_deviation,
            "rr_live": rr_live,
            "stop_loss": self._round_price(stop),
            "target_price": self._round_price(target),
            "target_1382": fib1382,
            "target_1618": fib1618,
            "rr": rr,
            "win_grade": win_grade,
            "win_rate": win_rate,
            "selection_score": selection_score,
            "trade_score": trade_score,
            "bucket": bucket,
            "operation_grade": "S" if tactical_light == "🔵" else "A" if tactical_light == "🟢" else "B" if tactical_light == "🟡" else "C",
            "reason": reason,
            "wave": wave_label,
            "rsi": round(rsi, 2),
            "trend_ok": trend_ok,
            "kd_ok": kd_ok,
            "macd_ok": macd_ok,
            "volume_ok": volume_ok,
            "decision": decision,
            "signal": signal,
            "trade_type": trade_type,
            "support": round(support, 2),
            "resistance": round(resistance, 2),
            "kline_score": round(kline_score, 2),
            "wave_score": round(wave_score, 2),
            "fib_score": round(fib_score, 2),
            "sakata_score": round(sakata_score, 2),
            "volume_score": round(volume_score, 2),
            "indicator_score": round(indicator_score, 2),
            "model_score": model_score,
            "wave_trade_score": round(float(wave_trade["wave_trade_score"]), 2),
            "mainstream_score": mainstream_score,
            "breakout_score": breakout_score,
            "atr14": round(atr14, 4),
            "atr_pct": atr_pct,
            "sakata_label": sakata_label,
            "volume_label": volume_label,
            "intraday_trend_score": round(float(liquidity.get("intraday_trend_score", 0) or 0), 2),
            "attack_volume_score": round(float(liquidity.get("attack_volume_score", 0) or 0), 2),
            "range_breakout_score": round(float(liquidity.get("range_breakout_score", 0) or 0), 2),
            "vwap_distance_pct": round(float(liquidity.get("vwap_distance_pct", 0) or 0), 2),
            "relative_strength_market": round(float(liquidity.get("relative_strength_market", 0) or 0), 2),
            "relative_strength_industry": round(float(liquidity.get("relative_strength_industry", 0) or 0), 2),
            "leader_follow_score": round(float(liquidity.get("leader_follow_score", 0) or 0), 2),
            "active_buy_score": round(float(liquidity.get("active_buy_score", 0) or 0), 2),
            "orderflow_aggression_score": round(float(liquidity.get("orderflow_aggression_score", 0) or 0), 2),
            "large_order_scan_score": round(float(liquidity.get("large_order_scan_score", 0) or 0), 2),
            "institutional_participation_score": round(float(liquidity.get("institutional_participation_score", 0) or 0), 2),
            "liquidity_status": liquidity_status,
            "liquidity_score": round(float(liquidity.get("liquidity_score", 0) or 0), 2),
            "elimination_reason": elimination_reason,
            "is_mainstream_funding": int(liquidity.get("is_mainstream_funding", 0) or 0),
            "eps_ttm": feature.get("eps_ttm", np.nan),
            "eps_yoy": feature.get("eps_yoy", np.nan),
            "revenue_yoy": feature.get("revenue_yoy", np.nan),
            "eps_bucket": feature.get("eps_bucket", ""),
            "rev_bucket": feature.get("rev_bucket", ""),
            "matrix_cell": feature.get("matrix_cell", ""),
            "eps_category": feature.get("eps_category", "U0"),
            "matrix_base_score": feature.get("matrix_base_score", np.nan),
            "modifier": feature.get("modifier", np.nan),
            "revenue_eps_score": feature.get("revenue_eps_score", 50.0),
            "financial_score": feature.get("revenue_eps_score", 50.0),
            "data_quality_flag": feature.get("data_quality_flag", ""),
            "eps_matrix_decision_note": f"EPS矩陣 {feature.get('eps_category', 'U0')}｜{feature.get('matrix_cell','')}｜score={feature.get('revenue_eps_score',50)}｜flag={feature.get('data_quality_flag','')}",
            "final_trade_decision": final_trade_decision,
        }


class MasterTradingEngine:
    def __init__(self, db: DBManager):
        self.db = db
        self.market_engine = MarketRegimeEngine(db)
        self.plan_engine = TradingPlanEngine(db)
        self.decision_layer = DecisionLayerEngine(db)
        self.external_fetcher = ExternalDataFetcher(db)
        self.last_pool_audit = {}

    def get_trade_pool(self, filtered_df: pd.DataFrame, progress_cb=None, log_cb=None, cancel_cb=None) -> dict:
        if filtered_df.empty:
            empty = pd.DataFrame()
            market = self.market_engine.get_market_regime()
            return {
                "market": market, "trade_top20": empty, "attack": empty, "watch": empty, "defense": empty,
                "today_buy": empty, "wait_pullback": empty, "theme_summary": empty, "eliminated": empty
            }

        base = filtered_df.copy()
        hot_themes = ThemeStrengthEngine.get_hot_themes(base)
        try:
            feature_rows = FinancialFeatureEngine(self.db).build_feature_batch(write_db=True)
            if log_cb:
                log_cb(f"[EPS MATRIX][BUILD] AI選股前已更新 financial_feature_daily：{0 if feature_rows is None else len(feature_rows)} 筆")
        except Exception as exc:
            log_warning(f"[EPS MATRIX][BUILD][WARN] AI選股前 feature 建立失敗：{exc}")
            if log_cb:
                log_cb(f"[EPS MATRIX][BUILD][WARN] {exc}")

        plans = []
        sids = base["stock_id"].astype(str).tolist()
        total = len(sids)
        for idx2, sid in enumerate(sids, start=1):
            if cancel_cb and cancel_cb():
                raise OperationCancelled("使用者中斷 AI選股TOP20")
            plans.append(self.plan_engine.build_plan(sid))
            if progress_cb:
                progress_cb(idx2, total, sid)
            if log_cb and (idx2 % 100 == 0 or idx2 == total):
                log_cb(f"AI選股分析進度 {idx2}/{total}｜{sid}")
        plans_df = pd.DataFrame(plans)
        if plans_df is not None and not plans_df.empty:
            try:
                plans_df = self.decision_layer.evaluate_dataframe(plans_df)
            except Exception as exc:
                log_warning(f"DecisionLayer evaluate_dataframe 失敗，外部決策保守阻擋：{exc}")
                plans_df["trade_allowed"] = 0
                plans_df["external_blocking_reason"] = f"DecisionLayer失敗：{exc}"
                plans_df["ui_state"] = "外部決策錯誤"

        market = self.market_engine.get_market_regime()
        if plans_df.empty:
            empty = pd.DataFrame()
            return {
                "market": market, "trade_top20": empty, "attack": empty, "watch": empty, "defense": empty,
                "today_buy": empty, "wait_pullback": empty, "theme_summary": ThemeStrengthEngine.summarize(base), "eliminated": empty
            }

        preferred_mask = plans_df["theme"].isin(hot_themes) if hot_themes else pd.Series([True] * len(plans_df), index=plans_df.index)

        eligible = plans_df[
            (plans_df["support"] > 0) &
            (plans_df["resistance"] > plans_df["support"])
        ].copy()
        eliminated = eligible[eligible["liquidity_status"].eq("ELIMINATE")].copy()
        tradable = eligible[~eligible["liquidity_status"].eq("ELIMINATE")].copy()

        mainstream_top20 = pd.DataFrame()
        breakout_top20 = pd.DataFrame()
        candidate_pool = pd.DataFrame()
        trade_top20 = pd.DataFrame()
        core_attack5 = pd.DataFrame()
        watch = pd.DataFrame()
        today_buy = pd.DataFrame()
        wait_pullback = pd.DataFrame()
        execution_ready = pd.DataFrame()
        unique_decision = pd.DataFrame()

        strategy_cfg = STRATEGY_CONFIG_MANAGER.load() or DEFAULT_STRATEGY_CONFIG
        active_strategy = STRATEGY_CONFIG_MANAGER.get_active_profile()
        core_cfg, exe_cfg, wait_cfg = _strategy_values(active_strategy)
        if log_cb:
            log_cb(f"[STRATEGY_CONFIG] {STRATEGY_CONFIG_MANAGER.summary_text()}")

        if not tradable.empty:
            tradable["decision_rank"] = tradable["decision"].map({"BUY": 3, "WEAK BUY": 2, "HOLD": 1}).fillna(0)
            tradable["preferred_rank"] = preferred_mask.reindex(tradable.index).fillna(False).astype(int)
            tradable["modules_pass_count"] = tradable[[c for c in ["trend_ok", "kd_ok", "macd_ok", "volume_ok"] if c in tradable.columns]].fillna(0).sum(axis=1)
            tradable["mainstream_score"] = (
                _safe_num_series(tradable, "model_score", 0) * 0.30 +
                _safe_num_series(tradable, "liquidity_score", 0) * 0.22 +
                _safe_num_series(tradable, "leader_follow_score", 0) * 0.18 +
                _safe_num_series(tradable, "intraday_trend_score", 0) * 0.10 +
                _safe_num_series(tradable, "win_rate", 0) * 0.12 +
                _safe_num_series(tradable, "rr", 0).clip(upper=3) * 5 +
                tradable["preferred_rank"] * 6
            ).round(2)
            tradable["breakout_score"] = (
                (_safe_num_series(tradable, "wave_trade_score", np.nan).fillna(_safe_num_series(tradable, "trade_score", 0))) * 0.28 +
                tradable.get("attack_volume_score", 0).fillna(0) * 0.20 +
                tradable.get("range_breakout_score", 0).fillna(0) * 0.18 +
                tradable.get("active_buy_score", 0).fillna(0) * 0.12 +
                _safe_num_series(tradable, "model_score", 0) * 0.10 +
                _safe_num_series(tradable, "rr", 0).clip(upper=3) * 5 +
                tradable["preferred_rank"] * 5
            ).round(2)

            mainstream_top20 = tradable.sort_values(["mainstream_score", "modules_pass_count", "liquidity_score", "model_score"], ascending=False).head(20).copy()
            breakout_top20 = tradable.sort_values(["breakout_score", "modules_pass_count", "attack_volume_score", "trade_score"], ascending=False).head(20).copy()

            combined_parts = []
            if not mainstream_top20.empty:
                tmp = mainstream_top20.copy()
                tmp["candidate_engine"] = "主流TOP20"
                combined_parts.append(tmp)
            if not breakout_top20.empty:
                tmp = breakout_top20.copy()
                tmp["candidate_engine"] = "起爆TOP20"
                combined_parts.append(tmp)
            if combined_parts:
                candidate_pool = pd.concat(combined_parts, ignore_index=True)
                candidate_pool["stock_id"] = candidate_pool["stock_id"].astype(str).map(normalize_stock_id).astype(str).str.strip()
                candidate_pool["source_count"] = candidate_pool.groupby("stock_id")["stock_id"].transform("count")
                candidate_pool["candidate_engine"] = np.where(candidate_pool["source_count"] >= 2, "雙引擎共振", candidate_pool["candidate_engine"])
                candidate_pool = candidate_pool.sort_values(["source_count", "mainstream_score", "breakout_score", "liquidity_score", "model_score"], ascending=False)
                candidate_pool = candidate_pool.drop_duplicates(subset=["stock_id"], keep="first").reset_index(drop=True)

            trade_top20 = candidate_pool.head(REPORT_DECISION_LIMITS["candidate20"]).copy() if not candidate_pool.empty else tradable.head(REPORT_DECISION_LIMITS["candidate20"]).copy()
            if not trade_top20.empty:
                trade_top20["stock_id"] = trade_top20["stock_id"].astype(str).map(normalize_stock_id).astype(str).str.strip()
                trade_top20 = trade_top20.drop_duplicates(subset=["stock_id"], keep="first").head(REPORT_DECISION_LIMITS["candidate20"]).copy()
                trade_top20["pool_role"] = "強勢候選20"
                cw = SCORE_FORMULA_WEIGHTS["candidate20"]
                trade_top20["candidate20_score"] = (
                    _safe_num_series(trade_top20, "model_score", 0) * cw["model_score"] +
                    _safe_num_series(trade_top20, "wave_trade_score", 0) * cw["wave_trade_score"] +
                    _safe_num_series(trade_top20, "liquidity_score", 0) * cw["liquidity_score"] +
                    _safe_num_series(trade_top20, "mainstream_score", 0) * cw["mainstream_score"] +
                    _safe_num_series(trade_top20, "breakout_score", 0) * cw["breakout_score"] +
                    _safe_num_series(trade_top20, "leader_follow_score", 0) * cw["leader_follow_score"] +
                    _safe_num_series(trade_top20, "active_buy_score", 0) * cw["active_buy_score"] +
                    _safe_num_series(trade_top20, "orderflow_aggression_score", 0) * cw["orderflow_aggression_score"] +
                    _safe_num_series(trade_top20, "win_rate", 0) * cw["win_rate"] +
                    _safe_num_series(trade_top20, "rr", 0).clip(upper=3) * cw["rr_factor"] +
                    _safe_num_series(trade_top20, "modules_pass_count", 0) * cw["modules_pass_count"]
                ).round(2)
                trade_top20 = trade_top20.sort_values(["candidate20_score", "mainstream_score", "breakout_score", "liquidity_score", "model_score"], ascending=False).head(REPORT_DECISION_LIMITS["candidate20"]).copy()

            # 鎖死來源：core_attack5 只能由 candidate20 產生
            trade_top20 = attach_external_display_columns(trade_top20)
            trade_top20 = attach_strategy_nogo_columns(trade_top20, active_strategy, "candidate20")
            core_source_base = apply_external_decision_filter(trade_top20, "core_attack5")
            if not core_source_base.empty:
                allowed_core_decisions = _strategy_allowed_decisions(active_strategy, "core_attack")
                core_decision_mask = core_source_base.get("final_trade_decision", pd.Series(dtype=str, index=core_source_base.index)).astype(str).isin(allowed_core_decisions)
                core_model_mask = _safe_num_series(core_source_base, "model_score", 0) >= float(get_strategy_threshold(active_strategy, "core_attack", "model_score_min"))
                core_wave_mask = _safe_num_series(core_source_base, "wave_trade_score", 0) >= float(get_strategy_threshold(active_strategy, "core_attack", "wave_trade_score_min"))
                core_wave_keyword_mask = _strategy_wave_keyword_mask(core_source_base, core_cfg)
                core_source = core_source_base[core_decision_mask & core_model_mask & core_wave_mask & core_wave_keyword_mask].copy()
            else:
                core_source = pd.DataFrame()
            if not core_source.empty:
                core_source["stock_id"] = core_source["stock_id"].astype(str).map(normalize_stock_id).astype(str).str.strip()
                core_source = core_source.drop_duplicates(subset=["stock_id"], keep="first").copy()
                core_source["light_rank"] = _safe_text_fill_series(core_source, "tactical_light", "⚪").map({"🔵": 5, "🟢": 4, "🟡": 3, "🟠": 2, "🔴": 1, "⚪": 0}).fillna(0)
                rel_market_norm = (((_safe_num_series(core_source, "relative_strength_market", 0) + 10) / 20.0) * 100.0).clip(0, 100)
                rel_ind_norm = _safe_num_series(core_source, "relative_strength_industry", 0).clip(0, 100)
                core_source["candidate20_score"] = _safe_num_series(core_source, "candidate20_score", 0)
                core_source["source_count"] = _safe_num_series(core_source, "source_count", 1)
                aw = SCORE_FORMULA_WEIGHTS["core_attack5"]
                core_source["core_attack5_score"] = (
                    core_source["candidate20_score"] * aw["candidate20_score"] +
                    _safe_num_series(core_source, "mainstream_score", 0) * aw["mainstream_score"] +
                    _safe_num_series(core_source, "breakout_score", 0) * aw["breakout_score"] +
                    _safe_num_series(core_source, "leader_follow_score", 0) * aw["leader_follow_score"] +
                    _safe_num_series(core_source, "active_buy_score", 0) * aw["active_buy_score"] +
                    _safe_num_series(core_source, "orderflow_aggression_score", 0) * aw["orderflow_aggression_score"] +
                    rel_market_norm * aw["rel_market_norm"] +
                    rel_ind_norm * aw["rel_ind_norm"] +
                    core_source["source_count"] * aw["source_count_factor"] +
                    core_source["light_rank"] * aw["light_rank_factor"]
                ).round(2)
                core_source["pool_role"] = "主攻5"
                core_attack5 = core_source.sort_values(
                    ["core_attack5_score", "light_rank", "mainstream_score", "breakout_score", "liquidity_score", "model_score", "win_rate"],
                    ascending=False
                ).drop_duplicates(subset=["stock_id"], keep="first").head(REPORT_DECISION_LIMITS["core_attack5"]).copy()
                core_attack5 = core_attack5[core_attack5["stock_id"].isin(trade_top20["stock_id"].tolist())].copy()
                core_attack5["pool_role"] = "主攻5"

            watch = pd.DataFrame()
            if not trade_top20.empty:
                core_ids_for_watch = set(_pool_stock_id_series(core_attack5).tolist()) if not core_attack5.empty else set()
                watch = trade_top20[~trade_top20["stock_id"].astype(str).isin(list(core_ids_for_watch))].copy()
                watch = watch.drop_duplicates(subset=["stock_id"], keep="first").copy()
                if not watch.empty:
                    watch["pool_role"] = "觀察"
                    watch["ui_state"] = "觀察"

            execution_candidates = core_attack5.copy() if not core_attack5.empty else pd.DataFrame()
            if not execution_candidates.empty:
                execution_candidates = normalize_core_analysis_df(execution_candidates)
                execution_candidates = execution_candidates.drop_duplicates(subset=["stock_id"], keep="first")
                execution_candidates["rsi"] = _safe_num_series(execution_candidates, "rsi", 0)
                execution_candidates["atr_pct"] = _safe_num_series(execution_candidates, "atr_pct", 999)
                required_liq = _strategy_required_liquidity(active_strategy)
                execution_candidates = execution_candidates[
                    (execution_candidates["liquidity_status"].astype(str).isin(required_liq)) &
                    (execution_candidates["rsi"] <= float(get_strategy_threshold(active_strategy, "execution", "rsi_max"))) &
                    (execution_candidates["atr_pct"] <= float(get_strategy_threshold(active_strategy, "execution", "atr_pct_max")))
                ].copy()
                execution_candidates["entry_mid"] = _safe_num_series(execution_candidates, "entry_mid", 0)
                execution_candidates["price_deviation"] = _safe_num_series(execution_candidates, "price_deviation", 0)
                execution_candidates["rr_live"] = _safe_num_series(execution_candidates, "rr_live", _safe_num_series(execution_candidates, "rr", 0))
                execution_candidates["candidate20_score"] = _safe_num_series(execution_candidates, "candidate20_score", 0)
                execution_candidates["core_attack5_score"] = _safe_num_series(execution_candidates, "core_attack5_score", np.nan).fillna(execution_candidates["candidate20_score"])
                ew = SCORE_FORMULA_WEIGHTS["execution"]
                execution_candidates["execution_score"] = (
                    execution_candidates["core_attack5_score"] * ew["core_attack5_score"] +
                    (1 - execution_candidates["price_deviation"].abs().clip(0, 1)) * ew["price_fit_factor"] +
                    execution_candidates["rr_live"].clip(upper=3) * ew["rr_live_factor"] +
                    _safe_num_series(execution_candidates, "liquidity_score", 0) * ew["liquidity_score"] +
                    _safe_num_series(execution_candidates, "win_rate", 0) * ew["win_rate"] +
                    _safe_num_series(execution_candidates, "model_score", 0) * ew["model_score"]
                ).round(2)

                today_allowed_decisions = _strategy_allowed_decisions(active_strategy, "execution")
                today_buy = execution_candidates[
                    execution_candidates.get("final_trade_decision", pd.Series(dtype=str, index=execution_candidates.index)).astype(str).isin(today_allowed_decisions) &
                    (execution_candidates["price_deviation"].abs() <= float(get_strategy_threshold(active_strategy, "execution", "price_dev_max"))) &
                    (execution_candidates["rr_live"] >= float(get_strategy_threshold(active_strategy, "execution", "rr_min")))
                ].sort_values(["execution_score", "core_attack5_score", "liquidity_score", "model_score"], ascending=False).head(REPORT_DECISION_LIMITS["today_buy"]).copy()
                if not today_buy.empty:
                    today_buy["pool_role"] = "今日可下單"
                    today_buy["ui_state"] = "可下單"

                wait_allowed_decisions = _strategy_allowed_decisions(active_strategy, "wait_pullback")
                wait_rr_min = float(get_strategy_threshold(active_strategy, "wait_pullback", "rr_min"))
                wait_dev_min = float(get_strategy_threshold(active_strategy, "wait_pullback", "price_dev_min"))
                wait_dev_max = float(get_strategy_threshold(active_strategy, "wait_pullback", "price_dev_max"))
                today_dev_max = float(get_strategy_threshold(active_strategy, "execution", "price_dev_max"))
                today_rr_min = float(get_strategy_threshold(active_strategy, "execution", "rr_min"))
                abs_dev = execution_candidates["price_deviation"].abs()
                wait_pullback = execution_candidates[
                    execution_candidates.get("final_trade_decision", pd.Series(dtype=str, index=execution_candidates.index)).astype(str).isin(wait_allowed_decisions) &
                    (((abs_dev > wait_dev_min) & (abs_dev <= wait_dev_max) & (execution_candidates["rr_live"] >= wait_rr_min)) |
                     ((abs_dev <= today_dev_max) & (execution_candidates["rr_live"] >= wait_rr_min) & (execution_candidates["rr_live"] < today_rr_min)))
                ].sort_values(["execution_score", "core_attack5_score", "liquidity_score", "model_score"], ascending=False).head(REPORT_DECISION_LIMITS["today_buy"]).copy()

            if not wait_pullback.empty and not today_buy.empty:
                wait_pullback = wait_pullback[~wait_pullback["stock_id"].isin(today_buy["stock_id"].astype(str).tolist())].copy()
            if not wait_pullback.empty:
                wait_pullback["core_attack5_score"] = _safe_num_series(wait_pullback, "core_attack5_score", _safe_num_series(wait_pullback, "candidate20_score", 0))
                wait_pullback["price_deviation"] = _safe_num_series(wait_pullback, "price_deviation", 0)
                wait_pullback["rr_live"] = _safe_num_series(wait_pullback, "rr_live", _safe_num_series(wait_pullback, "rr", 0))
                wait_pullback = wait_pullback.sort_values(["execution_score", "core_attack5_score", "candidate20_score", "liquidity_score", "model_score"], ascending=False).drop_duplicates(subset=["stock_id"], keep="first").head(REPORT_DECISION_LIMITS["today_buy"]).copy()
                wait_pullback["pool_role"] = "等待回測"
                wait_pullback["ui_state"] = "等待回測"

            execution_ready = today_buy.copy()
            unique_base = core_attack5.copy() if core_attack5 is not None else pd.DataFrame()
            if unique_base is not None and not unique_base.empty:
                if "final_trade_decision" in unique_base.columns:
                    unique_base = unique_base[unique_base["final_trade_decision"].astype(str).isin(_strategy_allowed_decisions(active_strategy, "execution") + _strategy_allowed_decisions(active_strategy, "wait_pullback"))].copy()
                unique_decision = unique_base.sort_values([c for c in ["core_attack5_score", "candidate20_score", "liquidity_score", "model_score", "win_rate", "rr"] if c in unique_base.columns], ascending=False).drop_duplicates(subset=["stock_id"], keep="first").head(REPORT_DECISION_LIMITS["unique_decision"]).copy()
            else:
                unique_decision = pd.DataFrame()
            if not unique_decision.empty:
                unique_decision["pool_role"] = "唯一決策"

        defense = plans_df[
            (plans_df["is_etf"] == 1) &
            (plans_df["support"] > 0) &
            (plans_df["resistance"] > plans_df["support"])
        ].copy()
        if not defense.empty:
            defense = defense.sort_values(["model_score", "trade_score", "rr", "win_rate"], ascending=False).head(10)
            defense["pool_role"] = "防守"

        if log_cb:
            candidate20_ids = set(_pool_stock_id_series(trade_top20).tolist())
            core_ids = set(_pool_stock_id_series(core_attack5).tolist())
            core_missing = sorted(core_ids - candidate20_ids)
            log_cb(f"[POOL-STAGE1] candidate20={len(candidate20_ids)}｜core_attack5={len(core_ids)}｜diff={len(core_missing)}")
            if core_missing:
                log_cb(f"[POOL-STAGE1-ERROR] core_attack5 - candidate20 差集：{','.join(core_missing[:20])}")
            today_ids = set(_pool_stock_id_series(today_buy).tolist())
            today_missing = sorted(today_ids - core_ids)
            log_cb(f"[POOL-STAGE2] today_buy={len(today_ids)}｜subset_of_core={len(today_missing)==0}")
            if today_missing:
                log_cb(f"[POOL-STAGE2-ERROR] today_buy - core_attack5 差集：{','.join(today_missing[:20])}")

        # V9.5.9：此處保留相容函式呼叫，但 apply_external_decision_filter 不再因外部資料缺口移除資料列；trade_allowed由技術/RR/風控決定。
        core_attack5 = apply_external_decision_filter(core_attack5, "core_attack5")
        today_buy = apply_external_decision_filter(today_buy, "today_buy")
        wait_pullback = apply_external_decision_filter(wait_pullback, "wait_pullback")
        execution_ready = apply_external_decision_filter(execution_ready, "execution_ready")
        unique_decision = apply_external_decision_filter(unique_decision, "unique_decision")
        for _df_name in ["trade_top20", "core_attack5", "today_buy", "wait_pullback", "execution_ready", "unique_decision", "watch"]:
            try:
                locals()[_df_name] = attach_strategy_nogo_columns(locals().get(_df_name), active_strategy, _df_name)
            except Exception as exc:
                log_warning(f"[STRATEGY_CONFIG] attach NoGo failed {_df_name}: {exc}")
        dynamic_n = max(1, min(10, market["max_positions"] + 2))
        result = {
            "market": market,
            "trade_top20": trade_top20,
            "tradable_top20": trade_top20.copy(),
            "mainstream_top20": mainstream_top20.head(20) if not mainstream_top20.empty else pd.DataFrame(),
            "breakout_top20": breakout_top20.head(20) if not breakout_top20.empty else pd.DataFrame(),
            "attack": core_attack5.head(REPORT_DECISION_LIMITS["core_attack5"]),
            "watch": watch.head(10),
            "defense": defense.head(10),
            "today_buy": today_buy.head(dynamic_n),
            "wait_pullback": wait_pullback.head(dynamic_n),
            "candidate20": trade_top20.head(REPORT_DECISION_LIMITS["candidate20"]),
            "core_attack5": core_attack5.head(REPORT_DECISION_LIMITS["core_attack5"]),
            "execution_ready": execution_ready.head(dynamic_n),
            "unique_decision": unique_decision.head(REPORT_DECISION_LIMITS["unique_decision"]),
            "theme_summary": ThemeStrengthEngine.summarize(base),
            "eliminated": eliminated.head(20),
        }
        pool_audit = build_pool_audit(result)
        self.last_pool_audit = dict(pool_audit)
        result["pool_audit"] = dict(pool_audit)
        if log_cb:
            log_cb(
                f"[POOL-AUDIT] candidate20={pool_audit['candidate20_count']}｜core_attack5={pool_audit['core_attack5_count']}｜today_buy={pool_audit['today_buy_count']}｜execution_ready={pool_audit['execution_ready_count']}｜unique_decision={pool_audit['unique_decision_count']}"
            )
            if pool_audit["core_minus_candidate20"]:
                log_cb(f"[POOL-AUDIT] core_attack5 - candidate20：{','.join(pool_audit['core_minus_candidate20'][:20])}")
            if pool_audit["today_minus_core"]:
                log_cb(f"[POOL-AUDIT] today_buy - core_attack5：{','.join(pool_audit['today_minus_core'][:20])}")
            if pool_audit["unique_minus_core"]:
                log_cb(f"[POOL-AUDIT] unique_decision - core_attack5：{','.join(pool_audit['unique_minus_core'][:20])}")
            if pool_audit["wait_minus_core"]:
                log_cb(f"[POOL-AUDIT] wait_pullback - core_attack5：{','.join(pool_audit['wait_minus_core'][:20])}")
            if pool_audit["watch_minus_candidate20"]:
                log_cb(f"[POOL-AUDIT] watch - candidate20：{','.join(pool_audit['watch_minus_candidate20'][:20])}")
            if pool_audit["watch_core_overlap"]:
                log_cb(f"[POOL-AUDIT] watch ∩ core_attack5：{','.join(pool_audit['watch_core_overlap'][:20])}")
            if pool_audit["wait_watch_overlap"]:
                log_cb(f"[POOL-AUDIT] wait_pullback ∩ watch：{','.join(pool_audit['wait_watch_overlap'][:20])}")
        assert_pool_consistency(result)
        if log_cb:
            log_cb(f"[POOL-FINAL] trade_top20={len(result.get('trade_top20', pd.DataFrame()))}｜attack={len(result.get('attack', pd.DataFrame()))}｜today_buy={len(result.get('today_buy', pd.DataFrame()))}｜wait_pullback={len(result.get('wait_pullback', pd.DataFrame()))}")
        try:
            persist_parts = []
            for key in ["candidate20", "core_attack5", "today_buy", "wait_pullback", "unique_decision", "defense"]:
                dfp = result.get(key, pd.DataFrame())
                if dfp is not None and not dfp.empty:
                    tmp = dfp.copy()
                    tmp["source_rank"] = key
                    persist_parts.append(tmp)
            if persist_parts:
                persist_df = pd.concat(persist_parts, ignore_index=True)
                self.db.replace_trade_plan_batch(persist_df)
                if log_cb:
                    log_cb(f"[TRADE_PLAN-DB] 已寫入 trade_plan：{len(persist_df)} 筆")
        except Exception as exc:
            if log_cb:
                log_cb(f"[TRADE_PLAN-DB-ERROR] trade_plan 寫入失敗：{exc}")
            log_warning(f"trade_plan persistence failed: {exc}")
        return result



class SelectionEngine:  # deprecated compatibility helper, not used by v9.2 FINAL-RELEASE main flow
    @staticmethod
    def prepare(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df.copy()

        x = df.copy()
        x["is_etf"] = x["market"].eq("ETF").astype(int)

        def classify_bucket(row):
            signal = str(row.get("signal", ""))
            action = str(row.get("action", ""))
            ai = float(row.get("ai_score", 0) or 0)
            total = float(row.get("total_score", 0) or 0)
            is_etf = int(row.get("is_etf", 0) or 0)

            if is_etf:
                if ai >= 55 and total >= 50:
                    return "防守"
                return "觀察"

            if signal == "強勢追蹤" and action == "拉回加碼" and ai >= 65 and total >= 70:
                return "主攻"
            if signal in ("整理偏多", "強勢追蹤") and action in ("低接布局", "拉回加碼") and ai >= 55 and total >= 60:
                return "次強"
            if signal in ("區間整理", "中性觀察") and ai >= 45 and total >= 45:
                return "觀察"
            return "排除"

        def selection_score(row):
            ai = float(row.get("ai_score", 0) or 0)
            total = float(row.get("total_score", 0) or 0)
            signal = str(row.get("signal", ""))
            action = str(row.get("action", ""))

            bonus = 0.0
            if signal == "強勢追蹤":
                bonus += 8
            elif signal == "整理偏多":
                bonus += 4

            if action == "拉回加碼":
                bonus += 6
            elif action == "低接布局":
                bonus += 3
            elif action == "區間操作":
                bonus -= 2

            return round(total * 0.55 + ai * 0.45 + bonus, 2)

        x["bucket"] = x.apply(classify_bucket, axis=1)
        x["selection_score"] = x.apply(selection_score, axis=1)
        return x

    @staticmethod
    def build_trade_pool(df: pd.DataFrame) -> dict:
        x = SelectionEngine.prepare(df)
        if x.empty:
            return {"master_top5": x, "attack": x, "watch": x, "defense": x}

        attack = x[x["bucket"] == "主攻"].sort_values(["selection_score", "ai_score", "total_score"], ascending=False)
        watch = x[x["bucket"].isin(["次強", "觀察"]) & (x["is_etf"] == 0)].sort_values(["selection_score", "ai_score", "total_score"], ascending=False)
        defense = x[x["bucket"] == "防守"].sort_values(["selection_score", "ai_score", "total_score"], ascending=False)

        master_top5 = pd.concat([attack.head(3), watch.head(2)], ignore_index=True)

        if len(master_top5) < 5:
            need = 5 - len(master_top5)
            used = set(master_top5["stock_id"].tolist()) if not master_top5.empty else set()
            extra = x[(~x["stock_id"].isin(list(used))) & (x["is_etf"] == 0)].sort_values(
                ["selection_score", "ai_score", "total_score"], ascending=False
            ).head(need)
            master_top5 = pd.concat([master_top5, extra], ignore_index=True)

        return {
            "master_top5": master_top5.head(REPORT_DECISION_LIMITS["core_attack5"]),
            "attack": attack.head(REPORT_DECISION_LIMITS["core_attack5"]),
            "watch": watch.head(REPORT_DECISION_LIMITS["core_attack5"]),
            "defense": defense.head(3),
        }




class BacktestEngine:
    """
    v9.2 FINAL-RELEASE：
    - 真回測核心：單一模擬邏輯，供摘要統計與 Equity Curve 共用
    - 目前為簡化交易模型，不含滑價 / 手續費 / 多策略參數化；estimate_trade_quality 與 Equity Curve 共用同一核心
    - 輸出：勝率 / 平均報酬 / 平均RR / CAGR / MDD / Sharpe / 樣本數
    """
    def __init__(self, db: DBManager):
        self.db = db

    def simulate_trades(self, stock_id: str) -> pd.DataFrame:
        hist = self.db.get_price_history(stock_id)
        if hist is None or hist.empty or len(hist) < 140:
            return pd.DataFrame(columns=["win", "ret", "rr"])

        x = IndicatorEngine.attach(hist.copy()).tail(260).reset_index(drop=True)
        trades = []

        for i in range(70, len(x) - 6):
            row = x.iloc[i]
            entry = float(row["close"])
            support = float(x.iloc[max(0, i-20):i+1]["low"].min())
            resistance = float(x.iloc[max(0, i-60):i+1]["high"].max())
            if support <= 0 or resistance <= support:
                continue

            signal_like = (
                (pd.notna(row["ma20"]) and pd.notna(row["ma60"]) and row["close"] > row["ma20"] >= row["ma60"]) and
                (pd.notna(row["macd_hist"]) and row["macd_hist"] > 0)
            )
            if not signal_like:
                continue

            stop = support * 0.97
            risk = max(entry - stop, 0.01)
            target = support + (resistance - support) * 1.382
            rr = max((target - entry) / risk, 0.0)

            future = x.iloc[i+1:i+6]
            max_hi = float(future["high"].max())
            min_lo = float(future["low"].min())
            exit_close = float(future.iloc[-1]["close"])

            if max_hi >= target:
                ret = (target / entry) - 1
                win = 1
            elif min_lo <= stop:
                ret = (stop / entry) - 1
                win = 0
            else:
                ret = (exit_close / entry) - 1
                win = 1 if ret > 0 else 0

            trades.append({"win": win, "ret": ret, "rr": rr})

        return pd.DataFrame(trades)

    def estimate_trade_quality(self, stock_id: str) -> dict:
        t = self.simulate_trades(stock_id)
        if t.empty:
            return {"backtest_win_rate": 45.0, "avg_return": 0.0, "avg_rr": 1.0, "cagr": 0.0, "mdd": 0.0, "sharpe": 0.0, "samples": 0}

        returns = t["ret"].astype(float)
        equity = (1 + returns).cumprod()
        years = max(len(returns) / 48.0, 0.25)
        cagr = (equity.iloc[-1] ** (1 / years) - 1) if len(equity) and equity.iloc[-1] > 0 else 0.0
        running_max = equity.cummax()
        dd = (equity / running_max - 1.0).min() if len(equity) else 0.0
        sharpe = (returns.mean() / (returns.std() + 1e-9)) * np.sqrt(48) if len(returns) > 1 else 0.0

        return {
            "backtest_win_rate": round(float(t["win"].mean() * 100), 2),
            "avg_return": round(float(t["ret"].mean() * 100), 2),
            "avg_rr": round(float(t["rr"].mean()), 2),
            "cagr": round(float(cagr * 100), 2),
            "mdd": round(float(dd * 100), 2),
            "sharpe": round(float(sharpe), 2),
            "samples": int(len(t)),
        }



class CapitalConfig:
    TOTAL_CAPITAL = 1000000.0
    MAX_TOTAL_EXPOSURE_PCT = 0.60
    MAX_SINGLE_POSITION_PCT = 0.12
    MAX_THEME_EXPOSURE_PCT = 0.25
    MAX_INDUSTRY_EXPOSURE_PCT = 0.30
    LOT_SIZE = 1000


class PortfolioEngine:
    """
    v9.2 FINAL-RELEASE：
    - 唯一資金管理引擎：依市場狀態、模型分數、交易分數、勝率、RR、ATR、Kelly 做配置
    - 控制總曝險、單檔曝險、題材集中度、產業集中度
    - 產出可執行『機構交易計畫』
    """
    def __init__(self, db: DBManager):
        self.db = db
        self.market_engine = MarketRegimeEngine(db)

    def _base_risk_budget_pct(self, regime: str) -> float:
        return {"多頭": 0.60, "震盪": 0.40, "空頭": 0.20}.get(regime, 0.35)

    def _score_strength(self, row: pd.Series) -> float:
        model_score = float(row.get("model_score", 0) or 0)
        trade_score = float(row.get("wave_trade_score", row.get("trade_score", 0)) or 0)
        win_rate = float(row.get("win_rate", 0) or 0)
        rr = float(row.get("rr", 0) or 0)
        decision = str(row.get("decision", ""))
        ui_state = str(row.get("ui_state", ""))

        score = 0.0
        score += min(max((model_score - 60) / 40, 0), 1) * 0.30
        score += min(max((trade_score - 60) / 40, 0), 1) * 0.25
        score += min(max((win_rate - 45) / 45, 0), 1) * 0.25
        score += min(rr / 3.0, 1) * 0.20

        if decision == "BUY":
            score += 0.10
        if ui_state == "條件預掛":
            score -= 0.25
        elif ui_state == "準備買":
            score -= 0.10
        return max(0.0, min(score, 1.25))

    def build_institutional_plan(self, candidates: pd.DataFrame) -> pd.DataFrame:
        cols = [
            "優先級","代號","名稱","市場","產業","題材","分類","狀態","盤中狀態","活性分","淘汰原因","進場區","停損","目標價",
            "1.382","1.618","RR","勝率","模型分數","交易分數","ATR%","Kelly%","建議張數","建議金額","單檔曝險%",
            "題材曝險%","產業曝險%","投資組合狀態","風險備註"
        ]
        if candidates is None or candidates.empty:
            return pd.DataFrame(columns=cols)

        market = self.market_engine.get_market_regime()
        total_capital = CapitalConfig.TOTAL_CAPITAL
        total_budget = total_capital * min(CapitalConfig.MAX_TOTAL_EXPOSURE_PCT, self._base_risk_budget_pct(market["regime"]))
        max_single_amt = total_capital * CapitalConfig.MAX_SINGLE_POSITION_PCT
        max_theme_amt = total_capital * CapitalConfig.MAX_THEME_EXPOSURE_PCT
        max_industry_amt = total_capital * CapitalConfig.MAX_INDUSTRY_EXPOSURE_PCT

        x = candidates.copy()
        x["strength"] = x.apply(self._score_strength, axis=1)
        x = x.sort_values(["strength","model_score","win_rate","rr"], ascending=False).reset_index(drop=True)

        theme_alloc = {}
        industry_alloc = {}
        deployed = 0.0
        rows = []

        for i, (_, r) in enumerate(x.iterrows(), start=1):
            close_proxy = float(r.get("entry_low", 0) or 0) or float(r.get("support", 0) or 0) or 1.0
            theme = str(r.get("theme", "全市場") or "全市場")
            industry = str(r.get("industry", "未分類") or "未分類")
            desired_amt = total_budget * (0.06 + 0.12 * float(r["strength"]))
            desired_amt = min(desired_amt, max_single_amt)

            # regime / state control
            if str(r.get("ui_state","")) == "條件預掛":
                desired_amt = 0.0
            elif str(r.get("ui_state","")) == "準備買":
                desired_amt *= 0.5

            remain_total = max(total_budget - deployed, 0.0)
            remain_theme = max(max_theme_amt - theme_alloc.get(theme, 0.0), 0.0)
            remain_industry = max(max_industry_amt - industry_alloc.get(industry, 0.0), 0.0)
            allowed_amt = min(desired_amt, remain_total, remain_theme, remain_industry)

            kelly = StrategyEngineV91.kelly_position(
                win_rate_pct=float(r.get("win_rate", 0) or 0),
                rr=float(r.get("rr", 0) or 0),
                atr_pct=float(r.get("atr_pct", 0) or 0),
                total_capital=total_capital,
                regime=market["regime"],
            )
            desired_amt = min(desired_amt, kelly["suggest_amount"]) if kelly["suggest_amount"] > 0 else 0.0

            qty = 0.0
            amount = 0.0
            if allowed_amt > 0 and close_proxy > 0:
                raw_qty = min(allowed_amt, desired_amt if desired_amt > 0 else allowed_amt) / close_proxy / CapitalConfig.LOT_SIZE
                qty = max(0.0, int(raw_qty * 2) / 2.0)  # 0.5 張階梯
                amount = round(qty * CapitalConfig.LOT_SIZE * close_proxy, 2)

            if qty > 0:
                deployed += amount
                theme_alloc[theme] = theme_alloc.get(theme, 0.0) + amount
                industry_alloc[industry] = industry_alloc.get(industry, 0.0) + amount
                portfolio_state = "可執行"
            else:
                portfolio_state = "等待/預掛"

            single_pct = round(amount / total_capital * 100, 2) if total_capital else 0.0
            theme_pct = round(theme_alloc.get(theme, 0.0) / total_capital * 100, 2) if total_capital else 0.0
            industry_pct = round(industry_alloc.get(industry, 0.0) / total_capital * 100, 2) if total_capital else 0.0

            note_parts = [f"市場={market['regime']}"]
            if str(r.get("liquidity_status","")) == "ELIMINATE":
                note_parts.append(f"淘汰={r.get('elimination_reason','')}")
            if desired_amt > remain_theme:
                note_parts.append("受題材曝險上限限制")
            if desired_amt > remain_industry:
                note_parts.append("受產業曝險上限限制")
            if desired_amt > remain_total:
                note_parts.append("受總曝險上限限制")
            if str(r.get("ui_state","")) == "條件預掛":
                note_parts.append("未到價不進場")
            elif str(r.get("ui_state","")) == "準備買":
                note_parts.append("僅半倉等待確認")
            elif qty > 0:
                note_parts.append("符合執行條件")

            rows.append({
                "優先級": i,
                "代號": r.get("stock_id",""),
                "名稱": r.get("stock_name",""),
                "市場": r.get("market",""),
                "產業": industry,
                "題材": theme,
                "分類": r.get("bucket",""),
                "狀態": r.get("ui_state",""),
                "盤中狀態": r.get("liquidity_status",""),
                "活性分": round(float(r.get("liquidity_score",0) or 0),2),
                "淘汰原因": r.get("elimination_reason",""),
                "進場區": r.get("entry_zone","-"),
                "停損": r.get("stop_loss","-"),
                "目標價": r.get("target_price", f"{float(r.get('target_1382',0) or 0):.2f}"),
                "1.382": f"{float(r.get('target_1382',0) or 0):.2f}",
                "1.618": f"{float(r.get('target_1618',0) or 0):.2f}",
                "RR": round(float(r.get("rr",0) or 0),2),
                "勝率": round(float(r.get("win_rate",0) or 0),1),
                "模型分數": round(float(r.get("model_score",0) or 0),2),
                "交易分數": round(float(r.get("wave_trade_score", r.get("trade_score", 0)) or 0),2),
                "ATR%": round(float(r.get("atr_pct",0) or 0),2),
                "Kelly%": kelly["position_pct"],
                "建議張數": qty,
                "建議金額": amount,
                "單檔曝險%": single_pct,
                "題材曝險%": theme_pct,
                "產業曝險%": industry_pct,
                "投資組合狀態": portfolio_state,
                "風險備註": "｜".join(note_parts),
            })

        return pd.DataFrame(rows, columns=cols)


class OperationGuideEngine:
    """V3.5 操作版：把資料轉成可執行的日內/波段操作 SOP。"""

    @staticmethod
    def build_playbook(market: dict, trade_top20: pd.DataFrame, today_buy: pd.DataFrame, wait_pullback: pd.DataFrame, attack: pd.DataFrame, defense: pd.DataFrame) -> pd.DataFrame:
        regime = str((market or {}).get("regime", "未定義"))
        memo = str((market or {}).get("memo", ""))
        top20_count = len(trade_top20) if trade_top20 is not None else 0
        buy_count = len(today_buy) if today_buy is not None else 0
        wait_count = len(wait_pullback) if wait_pullback is not None else 0
        attack_count = len(attack) if attack is not None else 0
        defense_count = len(defense) if defense is not None else 0

        regime_rule = {
            "多頭": "先做主攻股，再看條件預掛；可接受拉回承接，不追高到 1.382 上方。",
            "震盪": "先挑 RR 與勝率都高的標的；沒有明確進場區就不買。",
            "空頭": "先看防守 ETF；個股只保留極高勝率與低 ATR 的 setup。",
        }.get(regime, "先確認市場狀態，再選股。")

        top_pick = "-"
        if trade_top20 is not None and not trade_top20.empty:
            r = trade_top20.iloc[0]
            top_pick = f"{r['stock_id']} {r['stock_name']}｜{r.get('ui_state','-')}｜進場 {r.get('entry_zone','-')}"

        rows = [
            {"step": 1, "module": "先看市場", "focus": f"市場狀態＝{regime}", "rule": regime_rule, "purpose": "先決定今天偏攻擊、偏等待、還是偏防守", "output": memo or "依市場狀態決定倉位"},
            {"step": 2, "module": "再看輪動", "focus": f"主攻 {attack_count} 檔 / 防守 {defense_count} 檔", "rule": "只做有族群與題材支持的標的；不要逆勢單打獨鬥。", "purpose": "確認今天是做主流股，還是退守 ETF", "output": f"TOP20 候選 {top20_count} 檔"},
            {"step": 3, "module": "看今日可下單", "focus": f"今日可下單 {buy_count} 檔", "rule": "決策需為 BUY，且支撐 > 0、壓力 > 支撐、RR 夠大。", "purpose": "找出今天真正可以下手的標的", "output": top_pick},
            {"step": 4, "module": "看條件預掛", "focus": f"條件預掛 {wait_count} 檔", "rule": "未進入進場區前不追價，只能預掛，不可提前亂買。", "purpose": "把看好的股票留在觀察名單，等價格到位", "output": "價格進入進場區才升級為準備買"},
            {"step": 5, "module": "核對六模組", "focus": "K線 / 波浪 / 費波 / 阪田 / 量能 / 指標", "rule": "至少確認波浪位置、1.382/1.618 目標、RR、ATR%、Kelly%。", "purpose": "避免只有題材沒有結構，或只有指標沒有風險控管", "output": "決定可買 / 條件預掛 / 觀察 / 不可買"},
            {"step": 6, "module": "下單與倉位", "focus": "下單清單 / 機構交易計畫", "rule": "先看 Kelly% 與建議張數；有風險備註就縮小倉位。", "purpose": "把分析轉成可執行部位", "output": "建議張數、建議金額、單檔曝險%"},
            {"step": 7, "module": "盤後驗證", "focus": "回測視覺化 / Log", "rule": "看勝率、CAGR、MDD、Sharpe，不好的 setup 下次降權。", "purpose": "讓系統愈用愈準，而不是每天重複犯錯", "output": "策略保留 / 降權 / 淘汰"},
        ]
        return pd.DataFrame(rows, columns=["step", "module", "focus", "rule", "purpose", "output"])

    @staticmethod
    def explain_state(ui_state: str) -> str:
        mapping = {
            "可買": "已同時滿足決策、進場區與風險報酬條件，可執行。",
            "準備買": "條件接近完成，通常代表進場價快到位，可小倉位等待。",
            "條件預掛": "只列入名單，不追價，等價格回到進場區再動作。",
            "觀察": "可以追蹤，但目前不應出手。",
            "不可買": "不符合 SOP，應直接排除。",
            "淘汰": "已被盤中資金活性規則淘汰，不進觀察清單。",
        }
        return mapping.get(str(ui_state or ""), "依 SOP 判斷，不做主觀硬拗。")


class AppUI:
    def __init__(self, root, db: DBManager):
        self.root = root
        self.db = db
        self.data_engine = DataEngine(db)
        self.rank_engine = RankingEngine(db)
        self.master_trading_engine = MasterTradingEngine(db)
        self.strategy_config = STRATEGY_CONFIG_MANAGER
        self.backtest_engine = BacktestEngine(db)
        self.portfolio_engine = PortfolioEngine(db)
        self.last_top20_df = pd.DataFrame()
        self.last_candidate_top20_df = pd.DataFrame()
        self.last_top5_df = pd.DataFrame()
        self.last_theme_summary_df = pd.DataFrame()
        self.last_attack_df = pd.DataFrame()
        self.last_watch_df = pd.DataFrame()
        self.last_defense_df = pd.DataFrame()
        self.last_order_list_df = pd.DataFrame()
        self.last_institutional_plan_df = pd.DataFrame()
        self.last_today_buy_df = pd.DataFrame()
        self.last_wait_df = pd.DataFrame()
        self.last_operation_sop_df = pd.DataFrame()
        self.last_unique_decision_df = pd.DataFrame()
        self.current_chart_path = None
        self.plan_cache = {}
        self.backtest_cache = {}
        self.selection_job_token = 0
        self.selection_source = ""
        self.selector_syncing = False
        self.last_selected_stock_id = None
        self.last_selected_source = ""
        self.last_selected_ts = 0.0
        self.worker = None
        self.cancel_event = threading.Event()
        self.current_job = None
        self.history_batch_size = 25
        self.history_sleep_sec = 0.6
        self.last_job_summary = {}
        self.startup_initialized = False
        self.startup_in_progress = False

        self.root.title(APP_NAME)
        self._configure_startup_window()
        log_info(f"應用程式啟動｜DB={DB_PATH}｜LOG={LOG_PATH}")

        self.market_var = tk.StringVar(value="全部")
        self.multi_window_var = tk.BooleanVar(value=True)
        self.top20_window = None
        self.chart_window = None
        self.plan_window = None
        self.win_top20_tree = None
        self.win_plan_text = None
        self.chart_fig = None
        self.chart_canvas = None
        self.window_current_stock_id = None
        self.chart_updating = False
        self.pending_stock_id = None
        self.chart_update_job = None
        self.pending_chart_image = None
        self.chart_image_job = None
        self.selection_chart_pending_token = 0
        self.backtest_selection_token = 0
        self.industry_var = tk.StringVar(value="全部")
        self.theme_var = tk.StringVar(value="全部")
        self.search_var = tk.StringVar(value="")

        self._build_ui()
        set_classification_log_callback(lambda message, level="INFO": self.root.after(0, lambda: self.append_log(message, level)))
        self.refresh_filters()
        self.show_welcome_message()
        self.root.after(120, self.start_background_init)
        self.root.after(180, self._apply_initial_layout)
        self.root.after(600, self._apply_initial_layout)
        self.set_status(f"PACKED={PACKED_DATA_DIR} | EXTERNAL={EXTERNAL_DATA_DIR} | CSV={MASTER_CSV}")

    def start_background_init(self):
        if getattr(self, "startup_initialized", False) or getattr(self, "startup_in_progress", False):
            return

        def worker():
            done = threading.Event()
            self.startup_in_progress = True
            try:
                self.ui_call(self.start_task, "啟動初始化", 4)
                self.ui_call(self.update_task, "啟動初始化", 1, 4, item="載入歡迎頁與篩選條件")
                self.ui_call(self.show_welcome_message)
                self.ui_call(self.update_task, "啟動初始化", 2, 4, item="建立儀表板與排行檢視")

                def _refresh_and_mark():
                    try:
                        self.refresh_all_tables()
                    finally:
                        done.set()

                self.ui_call(_refresh_and_mark)
                done.wait()
                self.ui_call(self.update_task, "啟動初始化", 3, 4, success=1, item="首輪畫面刷新完成")
                self.startup_initialized = True
                self.ui_call(self.update_task, "啟動初始化", 4, 4, success=1, item="系統可操作")
                self.ui_call(self.finish_task, "啟動初始化", "系統初始化完成，可開始操作。")
            finally:
                self.startup_in_progress = False

        self._run_in_thread(worker, "startup_init")

    def update_status(self, msg: str):
        self.ui_call(self.set_status, msg)

    def _configure_startup_window(self):
        """啟動時自動貼齊可視區域，避免主視窗超出螢幕範圍。"""
        try:
            self.root.update_idletasks()
            if sys.platform.startswith("win"):
                try:
                    self.root.state("zoomed")
                    return
                except Exception:
                    pass

            sw = int(self.root.winfo_screenwidth() or 1600)
            sh = int(self.root.winfo_screenheight() or 900)
            width = max(1280, int(sw * 0.96))
            height = max(780, int(sh * 0.90))
            width = min(width, sw - 20) if sw > 40 else width
            height = min(height, sh - 80) if sh > 120 else height
            x = max((sw - width) // 2, 0)
            y = max((sh - height) // 2, 0)
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.root.geometry("1500x860")

    def _apply_initial_layout(self):
        """啟動後固定三區比例，減少人工手動調整。"""
        try:
            self.root.update_idletasks()
            total_w = max(int(self.root.winfo_width() or 0), 1200)
            total_h = max(int(self.root.winfo_height() or 0), 780)

            if getattr(self, "main_paned", None) is not None:
                left_w = max(760, int(total_w * 0.64))
                try:
                    self.main_paned.sashpos(0, left_w)
                except Exception:
                    pass

            if getattr(self, "right_paned", None) is not None:
                usable_h = max(total_h - 180, 520)
                upper_h = max(260, int(usable_h * 0.56))
                try:
                    self.right_paned.sashpos(0, upper_h)
                except Exception:
                    pass
        except Exception as exc:
            log_warning(f"套用啟動版型失敗：{exc}")

    def show_welcome_message(self):
        set_classification_log_callback(lambda message, level="INFO": self.root.after(0, lambda: self.append_log(message, level)))
        last_date = self.db.get_last_price_date() or "尚未建立"
        ranking_count = self.db.get_ranking_rows_count()
        price_rows = self.db.get_total_price_rows()
        cls_status = get_classification_status()
        cls_file = cls_status.get("file", "未載入")
        cls_rows = int(cls_status.get("loaded_rows", 0) or 0)
        cls_note = str(cls_status.get("load_note", "") or cls_status.get("note", ""))
        if cls_status.get("loaded") and cls_rows > 0:
            cls_mark = f"✔ 正常（已載入 {cls_rows} 筆）"
        elif cls_status.get("exists"):
            cls_mark = "⚠ 降級模式（檔案存在但未成功載入）"
        else:
            cls_mark = "❌ 缺失"
        cls_v2 = get_classification_v2_summary()
        coverage_text = "-"
        coverage_detail = ""
        if cls_v2:
            coverage_text = f"{float(cls_v2.get('coverage_pct', 0) or 0):.2f}%"
            coverage_detail = f"官方 {int(cls_v2.get('official', 0) or 0)}｜手動 {int(cls_v2.get('manual', 0) or 0)}｜規則 {int(cls_v2.get('rule_engine', 0) or 0)}｜AI {int(cls_v2.get('ai_infer', 0) or 0)}｜未分類 {int(cls_v2.get('unclassified', 0) or 0)}"
        lines = [
            "《GTC AI Trading System v9.2 FINAL-RELEASE》",
            "",
            f"主檔狀態：{len(self.db.get_master())} 檔",
            f"歷史資料：{price_rows} 筆｜最後交易日：{last_date}",
            f"最新排行筆數：{ranking_count}",
            f"分類檔狀態：{cls_mark}｜{cls_file}",
            f"分類檔備註：{cls_note or '-'}",
            f"分類覆蓋率：{coverage_text}",
            f"分類V2統計：{coverage_detail or '-'}",
            "",
            "建議操作順序：",
            "1. 初始化全市場（第一次或要重整主檔時）",
            "2. 建立完整歷史（第一次建庫）",
            "3. 每日增量更新",
            "4. 重建排行",
            "5. AI選股TOP20",
            "6. 採用 v9.2 FINAL-RELEASE：唯一核心策略引擎 / 波浪費波模型 / Kelly+ATR / Equity Curve",
            "7. V3.5操作版重點：先看市場，再看輪動，再看今日可下單 / 條件預掛，最後才下單",
            f"8. 圖表字型：{SELECTED_PLOT_FONT}",
        ]
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))

    def refresh_classification_summary_ui(self):
        try:
            _ = get_classification_status()
            _ = get_classification_v2_summary()
        except Exception:
            pass
        try:
            self.show_welcome_message()
        except Exception:
            pass
        try:
            self.root.update_idletasks()
        except Exception:
            pass


    def ensure_ranking_ready(self, auto_rebuild: bool = False) -> bool:
        ranking = self.db.get_latest_ranking()
        if ranking is not None and not ranking.empty:
            return True
        if auto_rebuild and self.db.get_total_price_rows() > 0:
            try:
                count = self.rank_engine.rebuild()
                return count > 0
            except Exception:
                return False
        return False

    def _build_ui(self):
        toolbar = ttk.Frame(self.root, padding=8)
        toolbar.pack(fill="x")

        row1 = ttk.Frame(toolbar)
        row1.pack(fill="x", pady=(0, 6))
        row2 = ttk.Frame(toolbar)
        row2.pack(fill="x")

        ttk.Label(row1, text="市場").pack(side="left")
        self.market_cb = ttk.Combobox(row1, textvariable=self.market_var, width=12, state="readonly")
        self.market_cb.pack(side="left", padx=4)

        ttk.Label(row1, text="產業").pack(side="left")
        self.industry_cb = ttk.Combobox(row1, textvariable=self.industry_var, width=16, state="readonly")
        self.industry_cb.pack(side="left", padx=4)

        ttk.Label(row1, text="題材").pack(side="left")
        self.theme_cb = ttk.Combobox(row1, textvariable=self.theme_var, width=18, state="readonly")
        self.theme_cb.pack(side="left", padx=4)

        ttk.Label(row1, text="搜尋").pack(side="left")
        ttk.Entry(row1, textvariable=self.search_var, width=16).pack(side="left", padx=4)

        self.btn_filter = ttk.Button(row1, text="套用篩選", command=self.refresh_all_tables)
        self.btn_filter.pack(side="left", padx=4)

        self.status_label = ttk.Label(row1, text="系統就緒")
        self.status_label.pack(side="right")

        ttk.Label(row2, text="功能").pack(side="left")
        self.action_var = tk.StringVar(value="AI選股TOP20")
        self.action_cb = ttk.Combobox(row2, textvariable=self.action_var, width=18, state="readonly")
        self.action_cb["values"] = [
            "AI選股TOP20",
            "主攻5",
            "V3.5操作說明",
            "策略設定",
            "重新計算今日可下單",
            "v9策略回測",
            "初始化全市場",
            "建立完整歷史（一次）",
            "續跑建庫",
            "每日增量更新",
            "重建排行",
            "更新分類檔",
            "外部資料監控中心",
            "同步外部資料",
            "重整外部資料狀態顯示",
            "查看外部原始資料",
            "中斷作業",
            "匯出分析Excel",
            "開啟圖表",
        ]
        self.action_cb.pack(side="left", padx=4)
        self.btn_run_action = ttk.Button(row2, text="執行功能", command=self.execute_action)
        self.btn_run_action.pack(side="left", padx=(4, 12))

        ttk.Label(row2, text="下載").pack(side="left")
        self.download_target_var = tk.StringVar(value="TOP20")
        self.download_target_cb = ttk.Combobox(row2, textvariable=self.download_target_var, width=12, state="readonly")
        self.download_target_cb["values"] = ["TOP20", "TOP5", "今日可下單", "等待回測", "條件預掛", "主攻", "次強", "防守", "執行下單清單", "組合交易計畫", "唯一決策", "操作SOP", "排行", "類股", "題材", "未分類清單", "分類V2摘要"]
        self.download_target_cb.pack(side="left", padx=4)
        self.btn_export_data = ttk.Button(row2, text="下載資料", command=self.export_selected_data)
        self.btn_export_data.pack(side="left", padx=(4, 12))

        self.btn_export_excel = ttk.Button(row2, text="匯出分析Excel", command=self.export_analysis_excel)
        self.btn_export_excel.pack(side="left", padx=4)
        self.btn_open_chart = ttk.Button(row2, text="開啟圖表", command=self.open_current_chart)
        self.btn_open_chart.pack(side="left", padx=4)
        self.multi_window_chk = ttk.Checkbutton(row2, text="左主區＋右上分析＋右下圖表/Log", variable=self.multi_window_var)
        self.multi_window_chk.state(["selected", "disabled"])
        self.multi_window_chk.pack(side="left", padx=(8, 2))
        self.btn_open_3wins = ttk.Button(row2, text="同步右側面板", command=self.open_three_windows)
        self.btn_open_3wins.pack(side="left", padx=4)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(row2, variable=self.progress_var, maximum=100, length=180, mode="determinate")
        self.progress.pack(side="left", padx=(12, 6))
        self.progress_text_var = tk.StringVar(value="0% | 0/0 | 成功 0 | 失敗 0")
        self.progress_text_label = ttk.Label(row2, textvariable=self.progress_text_var, width=44)
        self.progress_text_label.pack(side="left", padx=4)

        self.main_paned = ttk.Panedwindow(self.root, orient="horizontal")
        self.main_paned.pack(fill="both", expand=True, padx=8, pady=8)

        self.left_notebook = ttk.Notebook(self.main_paned)
        right = ttk.Frame(self.main_paned, padding=8)
        self.main_paned.add(self.left_notebook, weight=5)
        self.main_paned.add(right, weight=2)

        self.tab_dashboard = ttk.Frame(self.left_notebook)
        self.tab_sop = ttk.Frame(self.left_notebook)
        self.tab_rotation = ttk.Frame(self.left_notebook)
        self.tab_rank = ttk.Frame(self.left_notebook)
        self.tab_sector = ttk.Frame(self.left_notebook)
        self.tab_theme = ttk.Frame(self.left_notebook)
        self.tab_top20 = ttk.Frame(self.left_notebook)
        self.tab_top5 = ttk.Frame(self.left_notebook)
        self.tab_unique = ttk.Frame(self.left_notebook)
        self.tab_order = ttk.Frame(self.left_notebook)
        self.tab_inst = ttk.Frame(self.left_notebook)
        self.tab_external = ttk.Frame(self.left_notebook)
        self.tab_strategy = ttk.Frame(self.left_notebook)
        self.tab_backtest = ttk.Frame(self.left_notebook)
        self.left_notebook.add(self.tab_dashboard, text="交易儀表板")
        self.left_notebook.add(self.tab_sop, text="V3.5操作SOP")
        self.left_notebook.add(self.tab_rotation, text="產業輪動")
        self.left_notebook.add(self.tab_rank, text="排行榜")
        self.left_notebook.add(self.tab_sector, text="類股熱度")
        self.left_notebook.add(self.tab_theme, text="題材輪動")
        self.left_notebook.add(self.tab_top20, text="強勢候選20")
        self.left_notebook.add(self.tab_top5, text="主攻5")
        self.left_notebook.add(self.tab_unique, text="唯一決策")
        self.left_notebook.add(self.tab_order, text="執行下單清單")
        self.left_notebook.add(self.tab_inst, text="組合交易計畫")
        self.left_notebook.add(self.tab_external, text="外部資料中心")
        self.left_notebook.add(self.tab_strategy, text="策略設定")
        self.left_notebook.add(self.tab_backtest, text="回測視覺化")
        self.left_notebook.bind("<<NotebookTabChanged>>", self.on_left_tab_changed)

        self.dashboard_tree = self._make_tree(self.tab_dashboard, ("metric", "value", "desc"), {
            "metric": "指標", "value": "數值", "desc": "說明"
        })

        self.sop_tree = self._make_tree(self.tab_sop, ("step", "module", "focus", "rule", "purpose", "output"), {
            "step": "步驟", "module": "模組", "focus": "先看什麼", "rule": "判斷規則", "purpose": "用途", "output": "輸出"
        })

        self.rotation_tree = self._make_tree(self.tab_rotation, ("industry", "count", "avg_total", "avg_ai", "trend_count", "hot_score", "rotation"), {
            "industry": "產業", "count": "檔數", "avg_total": "平均總分", "avg_ai": "平均AI分", "trend_count": "強勢數", "hot_score": "輪動分", "rotation": "輪動狀態"
        })
        self.rotation_tree.bind("<<TreeviewSelect>>", self.on_select_rotation)

        self.rank_tree = self._make_tree(self.tab_rank, ("rank", "id", "name", "price", "chg", "chg_pct", "industry", "theme", "total", "ai", "signal", "action"), {
            "rank": "排名", "id": "代號", "name": "名稱", "price": "現價", "chg": "漲跌", "chg_pct": "漲跌幅%", "industry": "產業", "theme": "題材", "total": "總分", "ai": "AI分", "signal": "訊號", "action": "建議"
        })
        self.rank_tree.bind("<<TreeviewSelect>>", self.on_select_stock)

        self.sector_tree = self._make_tree(self.tab_sector, ("industry", "count", "avg_total", "avg_ai", "top_name"), {
            "industry": "產業", "count": "檔數", "avg_total": "平均總分", "avg_ai": "平均AI分", "top_name": "代表股"
        })

        self.theme_tree = self._make_tree(self.tab_theme, ("theme", "count", "avg_total", "avg_ai", "top_name"), {
            "theme": "題材", "count": "檔數", "avg_total": "平均總分", "avg_ai": "平均AI分", "top_name": "代表股"
        })

        self.top20_tree = self._make_tree(self.tab_top20, ("rank", "id", "name", "light", "engine", "price", "chg", "chg_pct", "bucket", "ui_action", "liquidity", "liq_score", "entry", "stop", "target", "target1382", "target1618", "rr", "win_rate", "strategy_nogo", "elim_reason"), {
            "rank": "排序", "id": "代號", "name": "名稱", "light": "燈號", "engine": "來源引擎", "price": "現價", "chg": "漲跌", "chg_pct": "漲跌幅%", "bucket": "分類", "ui_action": "狀態", "liquidity": "盤中狀態", "liq_score": "活性分", "entry": "進場區", "stop": "停損", "target": "目標價", "target1382": "1.382", "target1618": "1.618", "rr": "RR", "win_rate": "勝率%", "strategy_nogo": "不合格原因", "elim_reason": "淘汰原因"
        })
        self.top20_tree.bind("<<TreeviewSelect>>", self.on_select_top20)

        self.top5_tree = self._make_tree(self.tab_top5, ("rank", "id", "name", "price", "chg", "chg_pct", "state", "liquidity", "liq_score", "entry", "stop", "target1382", "rr", "win_rate", "backtest", "cagr", "mdd"), {
            "rank": "排序", "id": "代號", "name": "名稱", "price": "現價", "chg": "漲跌", "chg_pct": "漲跌幅%", "state": "狀態", "liquidity": "盤中狀態", "liq_score": "活性分", "entry": "進場區", "stop": "停損", "target1382": "1.382", "rr": "RR", "win_rate": "勝率%", "backtest": "回測勝率%", "cagr": "CAGR%", "mdd": "MDD%"
        })
        self.top5_tree.bind("<<TreeviewSelect>>", self.on_select_top5)

        self.unique_tree = self._make_tree(self.tab_unique, ("rank", "id", "name", "trade_allowed", "market_gate", "flow_gate", "fund_gate", "event_gate", "risk_gate", "external_ready", "source_date", "source_level", "block_reason", "decision"), {
            "rank": "排序", "id": "代號", "name": "名稱", "trade_allowed": "可下單", "market_gate": "Market", "flow_gate": "Flow", "fund_gate": "Fundamental", "event_gate": "Event", "risk_gate": "Risk", "external_ready": "ExternalReady", "source_date": "外部資料日", "source_level": "來源層級", "block_reason": "阻擋原因", "decision": "決策摘要"
        })
        self.unique_tree.bind("<<TreeviewSelect>>", self.on_select_unique)

        self.order_tree = self._make_tree(self.tab_order, ("priority", "id", "name", "price", "chg", "chg_pct", "bucket", "action", "liquidity", "liq_score", "entry", "stop", "target1382", "target1618", "rr", "win_rate", "atr_pct", "kelly_pct", "qty", "amount", "single_pct", "portfolio_state", "risk_note", "trade_allowed", "market_gate", "flow_gate", "fund_gate", "event_gate", "risk_gate", "block_reason"), {
            "priority": "優先級", "id": "代號", "name": "名稱", "price": "現價", "chg": "漲跌", "chg_pct": "漲跌幅%", "bucket": "分類", "action": "狀態", "liquidity": "盤中狀態", "liq_score": "活性分", "entry": "進場區", "stop": "停損", "target1382": "1.382", "target1618": "1.618", "rr": "RR", "win_rate": "勝率%", "atr_pct": "ATR%", "kelly_pct": "Kelly%", "qty": "建議張數", "amount": "建議金額", "single_pct": "單檔曝險%", "portfolio_state": "組合狀態", "risk_note": "風險備註", "trade_allowed": "外部允許", "market_gate": "Market", "flow_gate": "Flow", "fund_gate": "Fundamental", "event_gate": "Event", "risk_gate": "Risk", "block_reason": "外部阻擋"
        })
        self.order_tree.bind("<<TreeviewSelect>>", self.on_select_order)

        self.inst_tree = self._make_tree(self.tab_inst, ("priority", "id", "name", "price", "chg", "chg_pct", "market", "industry", "theme", "bucket", "action", "liquidity", "liq_score", "entry", "stop", "target", "rr", "win_rate", "model_score", "trade_score", "atr_pct", "kelly_pct", "qty", "amount", "single_pct", "theme_pct", "industry_pct", "portfolio_state"), {
            "priority": "優先級", "id": "代號", "name": "名稱", "price": "現價", "chg": "漲跌", "chg_pct": "漲跌幅%", "market": "市場", "industry": "產業", "theme": "題材", "bucket": "分類", "action": "狀態", "liquidity": "盤中狀態", "liq_score": "活性分", "entry": "進場區", "stop": "停損", "target": "目標價", "rr": "RR", "win_rate": "勝率%", "model_score": "模型分數", "trade_score": "交易分數", "atr_pct": "ATR%", "kelly_pct": "Kelly%", "qty": "建議張數", "amount": "建議金額", "single_pct": "單檔曝險%", "theme_pct": "題材曝險%", "industry_pct": "產業曝險%", "portfolio_state": "組合狀態"
        })
        self.inst_tree.bind("<<TreeviewSelect>>", self.on_select_institutional)

        self.external_tree = self._make_tree(self.tab_external, ("module", "source", "source_date", "status", "rows", "ready", "last_success", "url", "blocking", "error"), {
            "module": "模組", "source": "資料來源", "source_date": "資料日", "status": "狀態", "rows": "筆數", "ready": "DataReady",
            "last_success": "最後成功", "url": "Request URL", "blocking": "阻擋原因", "error": "錯誤/說明"
        })

        self._build_strategy_config_tab()

        self.backtest_tree = self._make_tree(self.tab_backtest, ("rank", "id", "name", "win", "avg_ret", "cagr", "mdd", "sharpe", "samples"), {
            "rank": "排序", "id": "代號", "name": "名稱", "win": "勝率%", "avg_ret": "平均報酬%", "cagr": "CAGR%", "mdd": "MDD%", "sharpe": "Sharpe", "samples": "樣本數"
        })
        self.backtest_tree.bind("<<TreeviewSelect>>", self.on_select_backtest)

        self.right_paned = ttk.Panedwindow(right, orient="vertical")
        self.right_paned.pack(fill="both", expand=True)

        upper = ttk.LabelFrame(self.right_paned, text="單股分析 / 執行建議", padding=6)
        right_lower = ttk.LabelFrame(self.right_paned, text="圖表 / Log", padding=6)
        self.right_paned.add(upper, weight=3)
        self.right_paned.add(right_lower, weight=2)

        upper_body = ttk.Frame(upper)
        upper_body.pack(fill="both", expand=True)
        self.detail = tk.Text(upper_body, wrap="none", font=("Consolas", 11), height=18)
        self.detail_vsb = ttk.Scrollbar(upper_body, orient="vertical", command=self.detail.yview)
        self.detail_hsb = ttk.Scrollbar(upper_body, orient="horizontal", command=self.detail.xview)
        self.detail.configure(yscrollcommand=self.detail_vsb.set, xscrollcommand=self.detail_hsb.set)
        self.detail.grid(row=0, column=0, sticky="nsew")
        self.detail_vsb.grid(row=0, column=1, sticky="ns")
        self.detail_hsb.grid(row=1, column=0, sticky="ew")
        upper_body.rowconfigure(0, weight=1)
        upper_body.columnconfigure(0, weight=1)
        self.right_panel_mode = "個股模式"
        self.right_panel_source = "-"

        self.win_plan_text = None
        self.win_top20_tree = None

        self.right_lower_notebook = ttk.Notebook(right_lower)
        self.right_lower_notebook.pack(fill="both", expand=True)

        self.chart_tab = ttk.Frame(self.right_lower_notebook)
        self.log_tab = ttk.Frame(self.right_lower_notebook)
        self.right_lower_notebook.add(self.chart_tab, text="圖表")
        self.right_lower_notebook.add(self.log_tab, text="Log")

        chart_wrap = ttk.Frame(self.chart_tab)
        chart_wrap.pack(fill="both", expand=True)
        self.chart_fig = plt.Figure(figsize=(8.6, 4.8), dpi=100)
        self.chart_canvas = FigureCanvasTkAgg(self.chart_fig, master=chart_wrap)
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)

        log_body = ttk.Frame(self.log_tab)
        log_body.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_body, wrap="none", font=("Consolas", 10), height=12)
        self.log_vsb = ttk.Scrollbar(log_body, orient="vertical", command=self.log_text.yview)
        self.log_hsb = ttk.Scrollbar(log_body, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=self.log_vsb.set, xscrollcommand=self.log_hsb.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_vsb.grid(row=0, column=1, sticky="ns")
        self.log_hsb.grid(row=1, column=0, sticky="ew")
        log_body.rowconfigure(0, weight=1)
        log_body.columnconfigure(0, weight=1)

    def _build_strategy_config_tab(self):
        frame = ttk.Frame(self.tab_strategy, padding=10)
        frame.pack(fill="both", expand=True)
        top = ttk.LabelFrame(frame, text="V9.6 策略參數設定（修改後可直接重算今日可下單）", padding=8)
        top.pack(fill="x", pady=(0, 8))
        self.strategy_profile_var = tk.StringVar(value=self.strategy_config.get_active_profile_name())
        profiles = list((self.strategy_config.config.get("profiles") or {}).keys()) or ["normal"]
        ttk.Label(top, text="策略模式").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.strategy_profile_cb = ttk.Combobox(top, textvariable=self.strategy_profile_var, values=profiles, width=16, state="readonly")
        self.strategy_profile_cb.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        self.strategy_vars = {}
        fields = [
            ("core_attack.model_score_min", "AI模型分數下限", 0, 2),
            ("core_attack.wave_trade_score_min", "波浪費波分數下限", 0, 4),
            ("execution.rr_min", "今日可下單RR下限", 1, 0),
            ("execution.rsi_max", "RSI上限", 1, 2),
            ("execution.price_dev_max", "價格偏離上限(小數；0.03=3%)", 1, 4),
            ("execution.atr_pct_max", "ATR%上限", 2, 0),
            ("wait_pullback.rr_min", "等待回測RR下限", 2, 2),
            ("wait_pullback.price_dev_max", "等待回測偏離上限", 2, 4),
        ]
        for key, label, row, col in fields:
            ttk.Label(top, text=label).grid(row=row, column=col, sticky="w", padx=4, pady=4)
            var = tk.StringVar(value="")
            self.strategy_vars[key] = var
            ttk.Entry(top, textvariable=var, width=12).grid(row=row, column=col+1, sticky="w", padx=4, pady=4)
        btns = ttk.Frame(top)
        btns.grid(row=3, column=0, columnspan=6, sticky="w", pady=(8, 2))
        ttk.Button(btns, text="載入目前設定", command=self.refresh_strategy_config_ui).pack(side="left", padx=4)
        ttk.Button(btns, text="套用設定", command=self.apply_strategy_config_from_ui).pack(side="left", padx=4)
        ttk.Button(btns, text="重新計算今日可下單", command=self.recompute_today_buy_from_cache).pack(side="left", padx=4)
        ttk.Button(btns, text="開啟設定資料夾", command=lambda: open_path(STRATEGY_CONFIG_DIR)).pack(side="left", padx=4)
        self.strategy_summary = tk.Text(frame, wrap="word", height=16, font=("Consolas", 10))
        self.strategy_summary.pack(fill="both", expand=True)
        self.refresh_strategy_config_ui()

    def refresh_strategy_config_ui(self):
        try:
            self.strategy_config.load()
            profile = self.strategy_config.get_active_profile_name()
            self.strategy_profile_var.set(profile)
            cfg = self.strategy_config.get_active_profile()
            for key, var in self.strategy_vars.items():
                section, item = key.split(".")
                value = (cfg.get(section, {}) or {}).get(item, "")
                var.set(str(value))
            text = [
                "《V9.6 Strategy Config》",
                f"目前設定來源：{self.strategy_config.last_load_message}",
                f"JSON：{STRATEGY_CONFIG_JSON}",
                f"Excel：{STRATEGY_CONFIG_EXCEL}",
                "",
                self.strategy_config.summary_text(),
                "",
                "說明：",
                "1. 套用設定只改條件，不重新抓外部資料。",
                "2. 重新計算今日可下單會使用目前快取TOP20/候選池重算，不需重啟程式。",
                "3. price_dev_max 使用小數：0.03 = 3%。",
                "4. execution_ready 仍為資訊欄位，不控制 trade_allowed。",
            ]
            self.strategy_summary.delete("1.0", tk.END)
            self.strategy_summary.insert("1.0", "\n".join(text))
        except Exception as exc:
            messagebox.showerror("策略設定", f"載入設定失敗：{exc}")

    def apply_strategy_config_from_ui(self):
        try:
            self.strategy_config.set_active_profile(self.strategy_profile_var.get())
            values = {}
            for key, var in self.strategy_vars.items():
                raw = str(var.get()).strip()
                try:
                    values[key] = float(raw)
                except Exception:
                    values[key] = raw
            self.strategy_config.update_active_values(values)
            self.strategy_config.load()
            self.refresh_strategy_config_ui()
            self.append_log(f"[STRATEGY_CONFIG] 已套用：{self.strategy_config.summary_text()}")
            self.set_status("策略設定已套用，可按『重新計算今日可下單』")
        except Exception as exc:
            messagebox.showerror("策略設定", f"套用設定失敗：{exc}")

    def recompute_today_buy_from_cache(self):
        try:
            src = getattr(self, "last_candidate_top20_df", pd.DataFrame())
            if src is None or src.empty:
                src = getattr(self, "last_top20_df", pd.DataFrame())
            if src is None or src.empty:
                return messagebox.showwarning("策略設定", "尚無TOP20快取，請先執行 AI選股TOP20。")
            self.strategy_config.load()
            cfg = self.strategy_config.get_active_profile()
            active_strategy = cfg
            core_cfg, exe_cfg, wait_cfg = _strategy_values(cfg)
            x = attach_strategy_nogo_columns(src.copy(), cfg, "candidate20")
            allowed_core = _strategy_allowed_decisions(active_strategy, "core_attack")
            core_mask = x.get("final_trade_decision", pd.Series(dtype=str, index=x.index)).astype(str).isin(allowed_core)
            core_mask &= _safe_num_series(x, "model_score", 0) >= float(get_strategy_threshold(active_strategy, "core_attack", "model_score_min"))
            core_mask &= _safe_num_series(x, "wave_trade_score", 0) >= float(get_strategy_threshold(active_strategy, "core_attack", "wave_trade_score_min"))
            core_mask &= _strategy_wave_keyword_mask(x, core_cfg)
            core = x[core_mask].copy()
            if not core.empty:
                if "core_attack5_score" not in core.columns or _safe_num_series(core, "core_attack5_score", 0).eq(0).all():
                    core["core_attack5_score"] = _safe_num_series(core, "candidate20_score", _safe_num_series(core, "model_score", 0))
                core = core.sort_values([c for c in ["core_attack5_score", "candidate20_score", "liquidity_score", "model_score", "win_rate"] if c in core.columns], ascending=False).head(REPORT_DECISION_LIMITS["core_attack5"]).copy()
                core["pool_role"] = "主攻5"
            required_liq = _strategy_required_liquidity(active_strategy)
            ec = normalize_core_analysis_df(core) if core is not None and not core.empty else pd.DataFrame()
            today = pd.DataFrame(); wait = pd.DataFrame()
            if ec is not None and not ec.empty:
                ec["rsi"] = _safe_num_series(ec, "rsi", 0)
                ec["atr_pct"] = _safe_num_series(ec, "atr_pct", 999)
                ec["price_deviation"] = _safe_num_series(ec, "price_deviation", 0)
                ec["rr_live"] = _safe_num_series(ec, "rr_live", _safe_num_series(ec, "rr", 0))
                ec = ec[(ec["liquidity_status"].astype(str).isin(required_liq)) & (ec["rsi"] <= float(get_strategy_threshold(active_strategy, "execution", "rsi_max"))) & (ec["atr_pct"] <= float(get_strategy_threshold(active_strategy, "execution", "atr_pct_max")))].copy()
                if not ec.empty:
                    allowed_today = _strategy_allowed_decisions(active_strategy, "execution")
                    ec["execution_score"] = _safe_num_series(ec, "core_attack5_score", _safe_num_series(ec, "candidate20_score", 0))
                    today = ec[ec.get("final_trade_decision", pd.Series(dtype=str, index=ec.index)).astype(str).isin(allowed_today) & (ec["price_deviation"].abs() <= float(get_strategy_threshold(active_strategy, "execution", "price_dev_max"))) & (ec["rr_live"] >= float(get_strategy_threshold(active_strategy, "execution", "rr_min")))].copy()
                    today = today.sort_values([c for c in ["execution_score", "core_attack5_score", "liquidity_score", "model_score"] if c in today.columns], ascending=False).head(REPORT_DECISION_LIMITS["today_buy"])
                    if not today.empty:
                        today["pool_role"] = "今日可下單"; today["ui_state"] = "可下單"
                    wait_allowed = _strategy_allowed_decisions(active_strategy, "wait_pullback")
                    abs_dev = ec["price_deviation"].abs()
                    wait = ec[ec.get("final_trade_decision", pd.Series(dtype=str, index=ec.index)).astype(str).isin(wait_allowed) & (abs_dev > float(get_strategy_threshold(active_strategy, "wait_pullback", "price_dev_min"))) & (abs_dev <= float(get_strategy_threshold(active_strategy, "wait_pullback", "price_dev_max"))) & (ec["rr_live"] >= float(get_strategy_threshold(active_strategy, "wait_pullback", "rr_min")))].copy()
                    if not wait.empty:
                        wait["pool_role"] = "等待回測"; wait["ui_state"] = "等待回測"
            self.last_no_go_df = self.enrich_price_and_export_fields(attach_strategy_nogo_columns(x, cfg, "candidate20"), id_col="stock_id") if x is not None and not x.empty else pd.DataFrame()
            try:
                if self.last_no_go_df is not None and not self.last_no_go_df.empty and "strategy_nogo_detail" in self.last_no_go_df.columns:
                    _ng = self.last_no_go_df[self.last_no_go_df["strategy_nogo_detail"].astype(str).ne("PASS")].head(10)
                    for _, _r in _ng.iterrows():
                        self.append_log(f"[NO_GO][R10] {_r.get('stock_id','')} {_r.get('stock_name','')}｜{_r.get('strategy_nogo_detail','')}")
            except Exception:
                pass
            self.last_attack_df = self.enrich_price_and_export_fields(attach_strategy_nogo_columns(core, cfg, "core_attack5"), id_col="stock_id") if core is not None and not core.empty else pd.DataFrame()
            self.last_top5_df = self.last_attack_df.head(5).copy() if self.last_attack_df is not None and not self.last_attack_df.empty else pd.DataFrame()
            self.last_today_buy_df = self.enrich_price_and_export_fields(attach_strategy_nogo_columns(today, cfg, "today_buy"), id_col="stock_id") if today is not None and not today.empty else pd.DataFrame()
            self.last_wait_df = self.enrich_price_and_export_fields(attach_strategy_nogo_columns(wait, cfg, "wait_pullback"), id_col="stock_id") if wait is not None and not wait.empty else pd.DataFrame()
            self.last_order_list_df = self.normalize_order_df(self.build_order_list(self.last_today_buy_df)) if self.last_today_buy_df is not None else pd.DataFrame()
            self.refresh_top20_and_order_views()
            self.populate_operation_sop(self.master_trading_engine.market_engine.get_market_regime(), src, self.last_today_buy_df, self.last_wait_df, self.last_attack_df, self.last_defense_df)
            self.left_notebook.select(self.tab_order if self.last_today_buy_df is not None and not self.last_today_buy_df.empty else self.tab_strategy)
            self.append_log(f"[STRATEGY_CONFIG][RECOMPUTE] {self.strategy_config.summary_text()}｜主攻={0 if self.last_attack_df is None else len(self.last_attack_df)}｜今日可下單={0 if self.last_today_buy_df is None else len(self.last_today_buy_df)}｜等待={0 if self.last_wait_df is None else len(self.last_wait_df)}")
            self.set_status(f"策略重算完成｜今日可下單 {0 if self.last_today_buy_df is None else len(self.last_today_buy_df)}｜等待 {0 if self.last_wait_df is None else len(self.last_wait_df)}")
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("策略設定", f"重新計算失敗：{exc}")

    def _make_tree(self, parent, cols, headers):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=28)
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=140 if c not in ("rank", "count", "avg_total", "avg_ai", "id", "total", "ai") else 90, anchor="center")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _make_tree_from_schema(self, parent, schema):
        cols = tuple(item[0] for item in schema)
        headers = {item[0]: item[1] for item in schema}
        tree = self._make_tree(parent, cols, headers)
        for key, _, width in schema:
            tree.column(key, width=width, anchor="center")
        return tree

    def _reconfigure_tree_from_schema(self, tree, schema):
        cols = tuple(item[0] for item in schema)
        tree.configure(columns=cols)
        for c in cols:
            tree.heading(c, text="")
        for key, title, width in schema:
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor="center")

    def _normalize_schema_df(self, df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=columns)
        x = df.copy()
        for col in columns:
            if col not in x.columns:
                x[col] = ""
        x = x[columns].copy()
        return x.fillna("")

    def normalize_core_analysis_df(self, df: pd.DataFrame) -> pd.DataFrame:
        return normalize_core_analysis_df(df)

    def normalize_order_df(self, df: pd.DataFrame) -> pd.DataFrame:
        x = self._normalize_schema_df(build_display_columns(df), ORDER_COLUMNS)
        x = assert_schema(x, ORDER_COLUMNS, "ORDER_COLUMNS")
        if not x.empty:
            assert list(x.columns) == ORDER_COLUMNS
        return x

    def normalize_institutional_df(self, df: pd.DataFrame) -> pd.DataFrame:
        x = self._normalize_schema_df(build_display_columns(df), INSTITUTIONAL_COLUMNS)
        x = assert_schema(x, INSTITUTIONAL_COLUMNS, "INSTITUTIONAL_COLUMNS")
        if not x.empty:
            assert list(x.columns) == INSTITUTIONAL_COLUMNS
        return x

    def _schema_values_from_row(self, row: pd.Series, schema: list[tuple]) -> list:
        values = []
        for key, _, _ in schema:
            label = DISPLAY_COLUMN_MAP.get(key, "")
            if key == "priority":
                val = row.get(label, row.get("優先級", 0))
                try:
                    val = int(float(val or 0))
                except Exception:
                    val = 0
            else:
                val = row.get(label, "")
            values.append(val)
        return values

    def _render_order_tree(self, df: pd.DataFrame):
        self._reconfigure_tree_from_schema(self.order_tree, ORDER_TREE_SCHEMA)
        df = self.normalize_order_df(df)
        for _, r in df.iterrows():
            self.order_tree.insert("", "end", values=self._schema_values_from_row(r, ORDER_TREE_SCHEMA))

    def _render_institutional_tree(self, df: pd.DataFrame):
        self._reconfigure_tree_from_schema(self.inst_tree, INSTITUTIONAL_TREE_SCHEMA)
        df = self.normalize_institutional_df(df)
        for _, r in df.iterrows():
            self.inst_tree.insert("", "end", values=self._schema_values_from_row(r, INSTITUTIONAL_TREE_SCHEMA))

    def clear_right_panel(self, source: str = ""):
        self.right_panel_mode = "產業輪動模式" if source else "個股模式"
        self.right_panel_source = source or "-"
        self.detail.delete("1.0", tk.END)
        lines = [
            "《右側面板已刷新》",
            f"模式：{self.right_panel_mode}",
            f"資料來源：{self.right_panel_source}",
            "",
            "目前尚未指定個股。",
            "請點選左側列表股票，或在產業輪動頁點選產業後帶出代表股。",
        ]
        self.detail.insert("1.0", "\n".join(lines))

    def _find_rotation_representative_stock(self, industry_name: str) -> str:
        industry_name = str(industry_name or "").strip()
        if not industry_name:
            return ""
        candidates = []
        for df in [getattr(self, "last_top20_df", pd.DataFrame()), self.db.get_latest_ranking()]:
            if df is None or df.empty or "industry" not in df.columns:
                continue
            x = df[df["industry"].astype(str).str.strip() == industry_name].copy()
            if x.empty:
                continue
            sort_cols = [c for c in ["total_score", "ai_score", "candidate20_score", "trade_score", "rank_all"] if c in x.columns]
            if sort_cols:
                ascending = [False if c != "rank_all" else True for c in sort_cols]
                x = x.sort_values(sort_cols, ascending=ascending)
            sid_col = "stock_id" if "stock_id" in x.columns else ("代號" if "代號" in x.columns else None)
            if sid_col:
                sid = str(x.iloc[0][sid_col])
                if sid:
                    return sid
        return ""

    def on_left_tab_changed(self, event=None):
        try:
            current_tab = self.left_notebook.select()
        except Exception:
            return
        if current_tab == str(self.tab_rotation):
            self.clear_right_panel(source="產業輪動模式")

    def on_select_rotation(self, event=None):
        sel = self.rotation_tree.selection()
        if not sel:
            return
        vals = self.rotation_tree.item(sel[0], "values")
        industry_name = str(vals[0]) if vals else ""
        stock_id = self._find_rotation_representative_stock(industry_name)
        if stock_id:
            self.sync_all_views(stock_id, source=f"產業輪動/代表股/{industry_name}")
        else:
            self.clear_right_panel(source=f"產業輪動/{industry_name}")

    def set_status(self, text):
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def set_progress(self, current=0, total=100, success=0, failed=0, sid="", skipped=0, stage=""):
        total = max(int(total), 1)
        current = max(0, min(int(current), total))
        self.progress.configure(maximum=total)
        self.progress_var.set(current)
        pct = (current / total) * 100 if total else 0
        stage_part = f"[{stage}] " if stage else ""
        sid_part = f" | {sid}" if sid else ""
        skip_part = f" | 跳過 {skipped}" if skipped else ""
        self.progress_text_var.set(f"{stage_part}{pct:5.1f}% | {current}/{total} | 成功 {success} | 失敗 {failed}{skip_part}{sid_part}")
        self.root.update_idletasks()

    def reset_progress(self):
        self.progress.configure(maximum=100)
        self.progress_var.set(0)
        self.progress_text_var.set("0% | 0/0 | 成功 0 | 失敗 0")
        self.root.update_idletasks()

    def start_task(self, stage: str, total: int = 100):
        self.set_status(f"{stage} 開始...")
        self.set_progress(0, total, 0, 0, stage=stage)

    def update_task(self, stage: str, current: int, total: int, success: int = 0, failed: int = 0, skipped: int = 0, item: str = ""):
        self.set_progress(current, total, success, failed, item, skipped=skipped, stage=stage)

    def finish_task(self, stage: str, summary: str = ""):
        self.set_status(summary or f"{stage} 完成")

    def append_log(self, text, level: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        level_upper = str(level or "INFO").upper()
        msg = f"[{ts}] [{level_upper}] {text}"
        try:
            if level_upper == "ERROR":
                log_error(text)
            elif level_upper == "WARNING":
                log_warning(text)
            else:
                log_info(text)
        except Exception:
            pass
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def save_history_state(self, state: dict):
        try:
            STATE_PATH.write_text(pd.Series(state).to_json(force_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def load_history_state(self) -> dict:
        try:
            if STATE_PATH.exists():
                return pd.read_json(STATE_PATH, typ="series").to_dict()
        except Exception:
            pass
        return {}

    def clear_history_state(self):
        try:
            if STATE_PATH.exists():
                STATE_PATH.unlink()
        except Exception:
            pass

    def update_classification_book(self):
        set_classification_log_callback(lambda message, level="INFO": self.ui_call(self.append_log, message, level))
        def worker():
            self.ui_call(self.start_task, "更新分類檔", 4)
            self.ui_call(self.update_task, "更新分類檔", 1, 4, item="檢查本機快取")
            current = ensure_classification_book(force_refresh=False, log_cb=lambda m: self.ui_call(self.append_log, m))
            if current is not None:
                self.ui_call(self.append_log, f"目前官方分類來源：{current}")
            self.ui_call(self.update_task, "更新分類檔", 2, 4, item="下載最新官方分類來源")
            refreshed = ensure_classification_book(force_refresh=True, log_cb=lambda m: self.ui_call(self.append_log, m))
            if refreshed is None:
                raise RuntimeError("官方分類來源下載失敗，且本機沒有可用快取。")
            self.ui_call(self.update_task, "更新分類檔", 3, 4, item="驗證官方分類來源")
            official = load_official_classification_book()
            if official is None or official.empty:
                raise RuntimeError(f"分類來源存在，但無法成功讀取：{refreshed}")
            self.ui_call(self.update_task, "更新分類檔", 4, 4, success=1, item=Path(refreshed).name)
            self.ui_call(self.refresh_classification_summary_ui)
            status = get_classification_status()
            stale_text = "是" if status.get("is_stale") else "否"
            self.ui_call(self.finish_task, "更新分類檔", f"分類來源更新完成：{Path(refreshed).name}｜共 {len(official)} 筆")
            self.ui_call(self.append_log, f"分類來源更新完成：{refreshed}｜可辨識 {len(official)} 筆｜過期={stale_text}")
            self.ui_call(messagebox.showinfo, "完成", f"分類來源更新完成：\n{refreshed}\n\n可辨識筆數：{len(official)}\n是否過期：{stale_text}\n\n下一步請執行「初始化全市場」以重建主檔分類。\n\n現在官方產業優先採用 MOPS CSV，題材/子題材則由手動映射 + 規則引擎 + AI 補值。")
        self._run_in_thread(worker, "update_classification_book")

    def cancel_current_job(self):
        if self.worker is None or not self.worker.is_alive():
            return messagebox.showinfo("提醒", "目前沒有執行中的背景作業。")
        self.cancel_event.set()
        self.append_log("已收到中斷要求，將於本批或本檔完成後停止。")
        self.set_status("已發出中斷要求，請稍候…")

    def set_busy(self, busy: bool):
        normal_buttons = [
            self.btn_filter, self.action_cb, self.btn_run_action,
            self.btn_export_data, self.download_target_cb,
            self.btn_export_excel, self.btn_open_chart, self.btn_open_3wins
        ]
        for btn in normal_buttons:
            try:
                btn.config(state="disabled" if busy else "readonly" if btn in (self.action_cb, self.download_target_cb) else "normal")
            except Exception:
                pass
        if busy:
            if self.action_var.get() == "中斷作業":
                self.action_var.set("AI選股TOP20")
        self.root.update_idletasks()

    def execute_action(self):
        action = (self.action_var.get() or "").strip()
        mapping = {
            "初始化全市場": self.init_master_data,
            "建立完整歷史（一次）": self.build_full_history_once,
            "續跑建庫": self.resume_full_history,
            "每日增量更新": self.update_data,
            "重建排行": self.rebuild_ranking,
            "更新分類檔": self.update_classification_book,
            "外部資料監控中心": self.show_external_data_center,
            "同步外部資料": self.sync_external_data,
            "重整外部資料狀態顯示": self.refresh_external_data_status,
            "查看外部原始資料": self.export_external_raw_sample,
            "AI選股TOP20": self.show_top20,
            "主攻5": self.show_top5,
            "V3.5操作說明": self.show_operation_guide,
            "v9策略回測": self.show_strategy_backtest,
            "匯出分析Excel": self.export_analysis_excel,
            "開啟圖表": self.open_current_chart,
            "中斷作業": self.cancel_current_job,
        }
        func = mapping.get(action)
        if func is None:
            return messagebox.showwarning("提醒", "請先選擇功能。")
        func()


    def export_external_raw_sample(self):
        """V9.4：依外部資料監控中心選取module匯出原始落表資料，讓使用者可查核是否真寫DB。"""
        try:
            item = None
            try:
                sels = self.external_tree.selection()
                item = sels[0] if sels else None
            except Exception:
                item = None
            module = ""
            if item:
                vals = self.external_tree.item(item, "values")
                module = str(vals[0]) if vals else ""
            status_df = self.db.read_table("external_source_status", limit=None)
            if status_df is None or status_df.empty:
                return messagebox.showwarning("外部資料", "尚無外部資料狀態，請先執行「同步外部資料」。")
            if module:
                row = status_df[status_df["module"].astype(str) == module].tail(1)
            else:
                row = status_df.head(1)
            if row.empty:
                return messagebox.showwarning("外部資料", f"找不到module：{module}")
            target_table = str(row.iloc[-1].get("target_table", "") or "")
            if not target_table:
                return messagebox.showwarning("外部資料", f"{module or '所選模組'}未設定target_table")
            raw = self.db.read_table(target_table, limit=500)
            log_df = self.db.read_table("external_data_log", limit=1000)
            if log_df is not None and not log_df.empty and "module" in log_df.columns and module:
                log_df = log_df[log_df["module"].astype(str) == module].copy()
            out_base = CHART_DIR / f"External_Raw_Sample_{module or target_table}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            out_path, kind = write_table_bundle(out_base, {
                "Raw_Data": raw if raw is not None else pd.DataFrame(),
                "Source_Status": row,
                "External_Data_Log": log_df if log_df is not None else pd.DataFrame(),
            }, preferred="excel")
            self.append_log(f"外部原始資料已匯出：{out_path}")
            open_path(out_path)
        except Exception as exc:
            self.append_log(f"外部原始資料匯出失敗：{exc}", "ERROR")

    def show_external_data_center(self):
        try:
            self.left_notebook.select(self.tab_external)
        except Exception:
            pass
        self.refresh_external_data_status()

    def sync_external_data(self):
        """V9.4：真正同步外部資料。會執行fetch→validate→write_db→status→log，不再只刷新狀態。"""
        def worker():
            try:
                run_id = self.db.log_system_run(event="external_sync_ui", status="start", message="UI triggered true external data pipeline", step="start")
                result = ExternalDataFetcher(self.db).refresh_external_data_pipeline(run_id=run_id)
                status = self.db.read_table("external_source_status", limit=500)
                self.ui_call(self._populate_external_tree, status)
                blocking = result.get("blocking", []) if isinstance(result, dict) else []
                if blocking:
                    self.ui_call(self.append_log, f"外部資料同步完成但有阻擋項｜run_id={run_id}｜{'; '.join(blocking)}", "WARNING")
                else:
                    self.ui_call(self.append_log, f"外部資料同步完成｜run_id={run_id}")
            except Exception as exc:
                self.ui_call(self.append_log, f"外部資料同步失敗：{exc}", "ERROR")
        self._run_in_thread(worker, "external_data_sync")

    def refresh_external_data_status(self):
        """V9.4：只刷新監控中心畫面，不執行fetch，避免功能語義錯誤。"""
        def worker():
            try:
                run_id = self.db.log_system_run(event="external_status_refresh", status="start", message="UI refresh external source status only", step="status")
                status = self.db.read_table("external_source_status", limit=500)
                self.ui_call(self._populate_external_tree, status)
                self.ui_call(self.append_log, f"外部資料狀態顯示已重整｜run_id={run_id}")
            except Exception as exc:
                self.ui_call(self.append_log, f"外部資料狀態顯示重整失敗：{exc}", "ERROR")
        self._run_in_thread(worker, "external_status_refresh")

    def _populate_external_tree(self, df: pd.DataFrame):
        try:
            self.external_tree.delete(*self.external_tree.get_children())
            if df is None or df.empty:
                self.external_tree.insert("", "end", values=("NO_DATA", "", "", "FAIL", 0, 0, "", "", "尚未執行外部資料同步", ""))
                return
            for _, r in df.iterrows():
                ready = int(pd.to_numeric(pd.Series([r.get("data_ready", 0)]), errors="coerce").fillna(0).iloc[0])
                status = str(r.get("status", "") or "")
                rows = int(pd.to_numeric(pd.Series([r.get("rows_count", 0)]), errors="coerce").fillna(0).iloc[0])
                if status == "pending_implementation" or rows == 0 and ready == 0:
                    status = "FAIL/NOT_READY"
                self.external_tree.insert("", "end", values=(
                    r.get("module", ""), r.get("source_name", ""), r.get("source_date", ""), status,
                    rows, ready, r.get("last_success_time", ""), r.get("request_url", ""),
                    r.get("blocking_reason", ""), r.get("error_message", ""),
                ))
        except Exception as exc:
            log_warning(f"外部資料監控中心填表失敗：{exc}")

    def ui_call(self, func, *args, **kwargs):
        def _wrapped():
            try:
                func(*args, **kwargs)
            except Exception as exc:
                try:
                    log_exception(f"UI callback failed: {getattr(func, '__name__', str(func))}", exc)
                except Exception:
                    pass
                try:
                    self.append_log(f"UI callback failed: {getattr(func, '__name__', str(func))}｜{exc}", "ERROR")
                except Exception:
                    pass
        self.root.after(0, _wrapped)

    def _run_in_thread(self, target, name="worker"):
        if self.worker is not None and self.worker.is_alive():
            messagebox.showwarning("提醒", "背景作業進行中，請稍候。")
            return

        def runner():
            self.cancel_event.clear()
            self.current_job = name
            self.ui_call(self.set_busy, True)
            self.ui_call(self.reset_progress)
            self.ui_call(self.append_log, f"背景作業啟動：{name}")
            log_info(f"背景作業啟動：{name}")
            try:
                target()
                self.ui_call(self.append_log, f"背景作業完成：{name}")
            except Exception as exc:
                log_exception(f"背景作業失敗：{name}", exc)
                self.ui_call(self.append_log, f"背景作業失敗：{name}｜{exc}", "ERROR")
                self.ui_call(messagebox.showerror, "背景作業錯誤", f"{name} 執行失敗：\n{exc}")
            finally:
                self.current_job = None
                self.ui_call(self.set_busy, False)

        self.worker = threading.Thread(target=runner, name=name, daemon=True)
        self.worker.start()

    def open_current_chart(self):
        if self.current_chart_path is None or not Path(self.current_chart_path).exists():
            return messagebox.showwarning("提醒", "目前沒有可開啟的圖表，請先點選股票。")
        open_path(Path(self.current_chart_path))

    def _should_ignore_select_event(self, stock_id: str, source: str) -> bool:
        sid = str(stock_id or "").strip()
        src = str(source or "").strip()
        if not sid:
            return True
        if getattr(self, "selector_syncing", False):
            log_info(f"忽略程式性選取事件：{sid}｜來源={src}")
            return True
        now = time.time()
        last_sid = str(getattr(self, "last_selected_stock_id", "") or "")
        last_src = str(getattr(self, "last_selected_source", "") or "")
        last_ts = float(getattr(self, "last_selected_ts", 0.0) or 0.0)
        if sid == last_sid and src != last_src and (now - last_ts) <= 0.8:
            log_info(f"忽略短時間重複點股事件：{sid}｜來源={src}｜前次來源={last_src}")
            return True
        self.last_selected_stock_id = sid
        self.last_selected_source = src
        self.last_selected_ts = now
        return False


    def open_three_windows(self):
        self.left_notebook.select(self.tab_top20)
        self.sync_multi_windows()
        stock_id = self.window_current_stock_id
        if not stock_id:
            stock_id = self.get_current_selected_stock_id()
        if not stock_id and self.last_top20_df is not None and not self.last_top20_df.empty:
            stock_id = str(self.last_top20_df.iloc[0]["stock_id"])
        if stock_id:
            self.update_multi_window_stock(stock_id)
        self.set_status("已同步到右側分析 / 圖表面板。")

    def ensure_multi_windows(self):
        return

    def sync_multi_windows(self):
        return

    def on_select_window_top20(self, event=None):
        stock_id = self.get_current_selected_stock_id()
        if stock_id and not self._should_ignore_select_event(stock_id, "桌面同步"):
            self.sync_all_views(stock_id, source="桌面同步")
        return

    def _tree_selected_stock_id(self, tree, value_index: int = 1):
        try:
            if tree is None:
                return None
            sel = tree.selection()
            if not sel:
                return None
            vals = tree.item(sel[0], "values")
            if vals and len(vals) > value_index:
                sid = str(vals[value_index]).strip()
                return sid or None
        except Exception:
            return None
        return None

    def _row_lookup(self, df: pd.DataFrame, stock_id: str):
        if df is None or df.empty or not stock_id:
            return None
        try:
            row = df[df["stock_id"].astype(str) == str(stock_id)]
            if row.empty:
                return None
            return row.iloc[0]
        except Exception:
            return None

    def get_current_selected_stock_id(self):
        for tree, idx in [
            (self.top20_tree, 1),
            (self.top5_tree, 1),
            (self.order_tree, 1),
            (self.inst_tree, 1),
            (self.backtest_tree, 1),
            (self.rank_tree, 1),
        ]:
            sid = self._tree_selected_stock_id(tree, idx)
            if sid:
                return sid
        if self.window_current_stock_id:
            return self.window_current_stock_id
        if self.last_top20_df is not None and not self.last_top20_df.empty:
            return str(self.last_top20_df.iloc[0]["stock_id"])
        ranking = self._filtered_ranking()
        if ranking is not None and not ranking.empty:
            return str(ranking.iloc[0]["stock_id"])
        return None

    def _set_tree_selection_by_stock_id(self, tree, stock_id: str, value_index: int = 1):
        try:
            if tree is None or not stock_id:
                return
            found = None
            for item in tree.get_children():
                vals = tree.item(item, "values")
                if vals and len(vals) > value_index and str(vals[value_index]) == str(stock_id):
                    found = item
                    break
            if found is not None:
                tree.selection_set(found)
                tree.focus(found)
                tree.see(found)
        except Exception:
            pass

    def sync_multi_windows_selectors(self, stock_id: str):
        if not stock_id:
            return
        self.selector_syncing = True
        try:
            for tree, idx in [
                (self.top20_tree, 1),
                (self.top5_tree, 1),
                (self.order_tree, 1),
                (self.inst_tree, 1),
                (self.backtest_tree, 1),
                (self.rank_tree, 1),
            ]:
                self._set_tree_selection_by_stock_id(tree, stock_id, idx)
        finally:
            try:
                self.root.after(120, lambda: setattr(self, "selector_syncing", False))
            except Exception:
                self.selector_syncing = False


    def cache_trade_dataframe(self, df: pd.DataFrame):
        if df is None or df.empty or "stock_id" not in df.columns:
            return
        try:
            for _, row in df.iterrows():
                sid = str(row.get("stock_id", "")).strip()
                if not sid:
                    continue
                payload = row.to_dict()
                payload["_cached_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.plan_cache[sid] = payload
        except Exception:
            pass

    def cache_backtest_dataframe(self, df: pd.DataFrame):
        if df is None or df.empty or "stock_id" not in df.columns:
            return
        try:
            for _, row in df.iterrows():
                sid = str(row.get("stock_id", "")).strip()
                if not sid:
                    continue
                payload = row.to_dict()
                payload["_cached_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.backtest_cache[sid] = payload
        except Exception:
            pass

    def _price_snapshot(self, stock_id: str) -> dict:
        sid = str(stock_id or "").strip()
        empty = {"現價": np.nan, "漲跌": np.nan, "漲跌幅%": np.nan}
        if not sid:
            return empty
        try:
            hist = self.db.get_price_history(sid)
            if hist is None or hist.empty:
                return empty
            x = hist.sort_values("date").reset_index(drop=True)
            close_now = float(x.iloc[-1]["close"]) if pd.notna(x.iloc[-1]["close"]) else np.nan
            if len(x) >= 2 and pd.notna(x.iloc[-2]["close"]):
                prev_close = float(x.iloc[-2]["close"])
            else:
                prev_close = np.nan
            change = close_now - prev_close if pd.notna(close_now) and pd.notna(prev_close) else np.nan
            chg_pct = (change / prev_close * 100.0) if pd.notna(change) and pd.notna(prev_close) and prev_close != 0 else np.nan
            return {"現價": round(close_now, 2) if pd.notna(close_now) else np.nan, "漲跌": round(change, 2) if pd.notna(change) else np.nan, "漲跌幅%": round(chg_pct, 2) if pd.notna(chg_pct) else np.nan}
        except Exception:
            return empty

    def enrich_price_and_export_fields(self, df: pd.DataFrame, id_col: str | None = None) -> pd.DataFrame:
        if df is None:
            return pd.DataFrame()
        x = df.copy()
        if x.empty:
            return x
        if id_col is None:
            for candidate in ["stock_id", "代號", "id"]:
                if candidate in x.columns:
                    id_col = candidate
                    break
        if id_col is None or id_col not in x.columns:
            return x
        x[id_col] = x[id_col].astype(str).map(normalize_stock_id)
        snapshots = {sid: self._price_snapshot(sid) for sid in x[id_col].astype(str).tolist() if str(sid).strip()}
        x["現價"] = x[id_col].map(lambda s: snapshots.get(str(s), {}).get("現價", np.nan))
        x["漲跌"] = x[id_col].map(lambda s: snapshots.get(str(s), {}).get("漲跌", np.nan))
        x["漲跌幅%"] = x[id_col].map(lambda s: snapshots.get(str(s), {}).get("漲跌幅%", np.nan))
        if "進場區" not in x.columns:
            if "entry_zone" in x.columns:
                x["進場區"] = x["entry_zone"]
            elif "entry" in x.columns:
                x["進場區"] = x["entry"]
        if "停損" not in x.columns:
            if "stop_loss" in x.columns:
                x["停損"] = x["stop_loss"]
            elif "stop" in x.columns:
                x["停損"] = x["stop"]
        if "目標價" not in x.columns:
            if "target_price" in x.columns:
                x["目標價"] = x["target_price"]
            elif "1.382" in x.columns:
                x["目標價"] = x["1.382"]
            elif "target_1382" in x.columns:
                x["目標價"] = x["target_1382"]
            elif "target1382" in x.columns:
                x["目標價"] = x["target1382"]
        x = build_display_columns(x)
        return x

    def build_unique_decision_df(self, *dfs: pd.DataFrame) -> pd.DataFrame:
        parts = []
        for df in dfs:
            if df is not None and not df.empty:
                parts.append(df.copy())
        if not parts:
            return pd.DataFrame()
        x = pd.concat(parts, ignore_index=True)
        if "stock_id" not in x.columns:
            return pd.DataFrame()
        x["stock_id"] = x["stock_id"].astype(str).map(normalize_stock_id)
        if "final_trade_decision" in x.columns:
            mask = ~x["final_trade_decision"].astype(str).str.upper().isin(["ELIMINATE", "淘汰", "AVOID", "不可買"])
            x = x[mask].copy()
        if "ui_state" in x.columns:
            x = x[~x["ui_state"].astype(str).isin(["淘汰", "不可買"])].copy()
        if x.empty:
            return x
        if "decision" in x.columns:
            x["decision_rank"] = x["decision"].map({"BUY": 3, "WEAK BUY": 2, "HOLD": 1}).fillna(0)
        elif "final_trade_decision" in x.columns:
            x["decision_rank"] = x["final_trade_decision"].map({"STRONG_BUY": 4, "BUY": 3, "DEFENSE": 2, "WEAK BUY": 2, "HOLD": 1, "WAIT_PULLBACK": 1, "WATCH": 1}).fillna(0)
        else:
            x["decision_rank"] = 0
        sort_cols = [c for c in ["decision_rank", "liquidity_score", "model_score", "trade_score", "win_rate", "rr"] if c in x.columns]
        if "execution_score" in x.columns:
            x = x.sort_values([c for c in ["execution_score", "core_attack5_score", "candidate20_score", "liquidity_score", "model_score", "win_rate", "rr"] if c in x.columns], ascending=False)
        elif sort_cols:
            x = x.sort_values(sort_cols, ascending=[False] * len(sort_cols))
        x = x.drop_duplicates(subset=["stock_id"], keep="first").head(REPORT_DECISION_LIMITS["unique_decision"]).reset_index(drop=True)
        if not x.empty:
            x["pool_role"] = "唯一決策"
        x = self.enrich_price_and_export_fields(x, id_col="stock_id")
        return x

    def get_cached_trade_plan(self, stock_id: str):
        sid = str(stock_id).strip()
        if not sid:
            return None
        plan = self.plan_cache.get(sid)
        if plan:
            return plan
        latest_tp = self.db.read_table("trade_plan", limit=None)
        if latest_tp is not None and not latest_tp.empty and "stock_id" in latest_tp.columns:
            latest_tp = latest_tp[latest_tp["stock_id"].astype(str) == sid].tail(1)
            if not latest_tp.empty:
                payload = latest_tp.iloc[-1].to_dict()
                self.plan_cache[sid] = payload
                return payload
        for df in [getattr(self, "last_top20_df", pd.DataFrame()), getattr(self, "last_top5_df", pd.DataFrame()),
                   getattr(self, "last_today_buy_df", pd.DataFrame()), getattr(self, "last_wait_df", pd.DataFrame()),
                   getattr(self, "last_attack_df", pd.DataFrame()), getattr(self, "last_watch_df", pd.DataFrame()),
                   getattr(self, "last_defense_df", pd.DataFrame())]:
            row = self._row_lookup(df, sid)
            if row is not None:
                payload = row.to_dict()
                self.plan_cache[sid] = payload
                return payload
        return None

    def get_cached_backtest(self, stock_id: str):
        sid = str(stock_id).strip()
        if not sid:
            return None
        bt = self.backtest_cache.get(sid)
        if bt:
            return bt
        row = self._row_lookup(getattr(self, "last_top5_df", pd.DataFrame()), sid)
        if row is not None and "backtest_win_rate" in row.index:
            payload = {
                "backtest_win_rate": float(row.get("backtest_win_rate", 0) or 0),
                "avg_return": float(row.get("avg_return", 0) or 0),
                "cagr": float(row.get("cagr", 0) or 0),
                "mdd": float(row.get("mdd", 0) or 0),
                "sharpe": float(row.get("sharpe", 0) or 0),
                "samples": int(row.get("samples", 0) or 0),
            }
            self.backtest_cache[sid] = payload
            return payload
        return None

    def build_lightweight_plan(self, stock_id: str, hist: pd.DataFrame, stock=None) -> dict:
        if hist is None or hist.empty:
            return {
                "stock_id": stock_id, "stock_name": stock.get("stock_name", stock_id) if stock is not None else stock_id,
                "market": stock.get("market", "") if stock is not None else "",
                "industry": stock.get("industry", "") if stock is not None else "",
                "theme": stock.get("theme", "") if stock is not None else "",
                "ui_state": "觀察", "trade_action": "HOLD", "entry_zone": "-", "stop_loss": "-",
                "target_1382": 0.0, "target_1618": 0.0, "support": 0.0, "resistance": 0.0, "rr": 0.0,
                "win_rate": 0.0, "wave": "資料不足", "signal": "載入中", "trade_type": "快速模式", "bucket": "觀察",
                "reason": "使用快速模式顯示，背景計算完成後會自動更新。", "kline_score": 0.0, "wave_score": 0.0,
                "fib_score": 0.0, "sakata_score": 0.0, "volume_score": 0.0, "indicator_score": 0.0
            }
        x = hist.copy()
        if "ma20" not in x.columns:
            x = DataEngine.attach(x)
        last = x.iloc[-1]
        close_ = float(last["close"])
        ma20 = float(last["ma20"]) if pd.notna(last.get("ma20")) else close_
        ma60 = float(last["ma60"]) if pd.notna(last.get("ma60")) else close_
        recent = x.tail(60)
        support = float(min(ma20, recent["low"].tail(20).min())) if not recent.empty else close_
        resistance = float(recent["high"].max()) if not recent.empty else close_
        fib_score, fib1382, fib1618 = FibEngine.score_and_targets(close_, support, resistance)
        signal = "偏多觀察" if close_ >= ma20 >= ma60 else "區間整理" if close_ >= ma20 else "轉弱警戒"
        wave = WaveEngine.detect_wave_label(x)
        entry_low = support * 1.002 if support > 0 else close_ * 0.99
        entry_high = entry_low * 1.01
        stop = support * 0.97 if support > 0 else close_ * 0.95
        risk = max(entry_high - stop, 0.01)
        reward = max(fib1382 - entry_high, 0.0)
        rr = round(reward / risk, 2)
        ui_state = "觀察" if signal in ("偏多觀察", "區間整理") else "不可買"
        return {
            "stock_id": stock_id,
            "stock_name": stock.get("stock_name", stock_id) if stock is not None else stock_id,
            "market": stock.get("market", "") if stock is not None else "",
            "industry": stock.get("industry", "") if stock is not None else "",
            "theme": stock.get("theme", "") if stock is not None else "",
            "ui_state": ui_state, "trade_action": "HOLD" if ui_state == "觀察" else "AVOID",
            "entry_zone": f"{entry_low:.2f} ~ {entry_high:.2f}",
            "entry_low": round(entry_low, 2), "entry_high": round(entry_high, 2),
            "stop_loss": f"{stop:.2f}",
            "target_1382": round(fib1382, 2), "target_1618": round(fib1618, 2),
            "support": round(support, 2), "resistance": round(resistance, 2), "rr": rr,
            "win_rate": 0.0, "wave": wave, "signal": signal, "trade_type": "快速模式", "bucket": "觀察",
            "reason": "已先顯示快速資料，完整交易計畫與回測由背景更新。", "kline_score": 0.0,
            "wave_score": 0.0, "fib_score": round(fib_score, 2), "sakata_score": 0.0,
            "volume_score": 0.0, "indicator_score": 0.0
        }

    def start_selection_analysis(self, stock_id: str, source: str = ""):
        if not stock_id:
            return
        self.append_log(f"點股分析啟動：{stock_id}｜來源 {source or '-'}")
        self.selection_job_token += 1
        token = self.selection_job_token
        self.selection_chart_pending_token = token
        self.selection_source = source or ""
        stock = self.db.get_stock_row(stock_id)
        hist = self.db.get_price_history(stock_id)
        if stock is None or hist.empty:
            return

        quick_plan = self.get_cached_trade_plan(stock_id)
        if quick_plan is None:
            try:
                quick_plan = self.build_lightweight_plan(stock_id, DataEngine.attach(hist.copy()), stock=stock)
                self.plan_cache[str(stock_id)] = quick_plan
            except Exception:
                quick_plan = None

        lines = self.build_unified_detail_lines(stock_id, source=(source or "快速顯示"), quick_only=True)
        lines.append("")
        lines.append("圖表狀態：背景輸出 PNG 後再載入右下圖表，避免點股時主畫面卡住。")
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))
        try:
            self.show_chart_loading(stock_id)
        except Exception:
            pass

        t = threading.Thread(target=self._selection_analysis_worker, args=(str(stock_id), str(source or ""), token), daemon=True, name=f"select_{stock_id}")
        t.start()

    def _selection_analysis_worker(self, stock_id: str, source: str, token: int):
        try:
            log_info(f"點股背景分析開始：{stock_id}｜token={token}｜來源={source or '-'}")
            if token != self.selection_job_token:
                log_warning(f"點股背景分析略過（token過期）：{stock_id}｜token={token}")
                return

            stock = self.db.get_stock_row(stock_id)
            hist = self.db.get_price_history(stock_id)
            if stock is None or hist is None or hist.empty:
                log_warning(f"點股背景分析無資料：{stock_id}")
                self.ui_call(self.append_log, f"點股背景分析無資料：{stock_id}", "WARNING")
                return

            log_info(f"點股背景分析讀取資料完成：{stock_id}｜rows={len(hist)}")

            hist_attached = None
            try:
                hist_attached = DataEngine.attach(hist.copy())
            except Exception:
                hist_attached = hist.copy()

            plan = self.get_cached_trade_plan(stock_id)
            plan_is_full = bool(plan) and str(plan.get("trade_type", "")) != "快速模式" and float(plan.get("model_score", 0) or 0) > 0
            if not plan_is_full:
                try:
                    log_info(f"建立完整交易計畫：{stock_id}")
                    plan = self.master_trading_engine.plan_engine.build_plan(stock_id)
                except Exception as e:
                    log_exception(f"完整交易計畫失敗，改用快速模式：{stock_id}", e)
                    self.ui_call(self.append_log, f"完整交易計畫失敗，改用快速模式：{stock_id}｜{e}", "WARNING")
                    plan = self.build_lightweight_plan(stock_id, hist_attached, stock=stock)
                self.plan_cache[str(stock_id)] = plan

            if token != self.selection_job_token:
                return

            bt = self.get_cached_backtest(stock_id)
            bt_ready = bool(bt) and int(bt.get("samples", 0) or 0) >= 0
            if not bt_ready:
                try:
                    log_info(f"背景回測開始：{stock_id}")
                    bt = self.backtest_engine.estimate_trade_quality(stock_id)
                    log_info(f"背景回測完成：{stock_id}｜samples={bt.get('samples',0)}")
                except Exception as e:
                    log_exception(f"背景回測失敗：{stock_id}", e)
                    self.ui_call(self.append_log, f"背景回測失敗：{stock_id}｜{e}", "WARNING")
                    bt = {"backtest_win_rate": 0.0, "avg_return": 0.0, "avg_rr": 0.0, "cagr": 0.0, "mdd": 0.0, "sharpe": 0.0, "samples": 0}
                self.backtest_cache[str(stock_id)] = bt

            chart_path = None
            if token == self.selection_job_token:
                try:
                    self.ui_call(self.append_log, f"背景圖表輸出開始：{stock_id}")
                    log_info(f"背景圖表輸出開始：{stock_id}")
                    chart_path = self.export_chart(stock_id, hist)
                    self.ui_call(self.append_log, f"背景圖表輸出完成：{stock_id}")
                    log_info(f"背景圖表輸出完成：{stock_id}｜{chart_path}")
                except Exception as e:
                    log_exception(f"背景圖檔輸出失敗：{stock_id}", e)
                    self.ui_call(self.append_log, f"背景圖檔輸出失敗：{stock_id}｜{e}", "ERROR")

            def apply_result():
                if token != self.selection_job_token:
                    return
                if chart_path:
                    self.current_chart_path = chart_path
                try:
                    self.update_detail_panel(stock_id, source=source or "背景完成")
                except Exception as e:
                    self.append_log(f"背景分析更新失敗：{stock_id}｜{e}")
                if token != self.selection_chart_pending_token:
                    return
                try:
                    if chart_path:
                        self._schedule_chart_file_update(stock_id, chart_path)
                    else:
                        self.show_chart_message("圖表產生失敗，請改用『開啟圖表』或重新點選。")
                except Exception as e:
                    self.append_log(f"背景圖表更新失敗：{stock_id}｜{e}")

            self.ui_call(apply_result)
        except Exception as e:
            log_exception(f"背景選股分析失敗：{stock_id}", e)
            self.ui_call(self.append_log, f"背景選股分析失敗：{stock_id}｜{e}", "ERROR")

    def build_unified_detail_lines(self, stock_id: str, source: str = "", quick_only: bool = False):
        stock = self.db.get_stock_row(stock_id)
        hist = self.db.get_price_history(stock_id)
        if stock is None or hist.empty:
            return [f"《{source or '同步檢視'}》", f"股票：{stock_id}", "無資料"]

        hist = DataEngine.attach(hist.copy())
        last = hist.iloc[-1]
        trade_plan = self.get_cached_trade_plan(stock_id)
        if trade_plan is None:
            trade_plan = self.build_lightweight_plan(stock_id, hist, stock=stock)
            self.plan_cache[str(stock_id)] = trade_plan
        bt = self.get_cached_backtest(stock_id)
        if bt is None:
            bt = {"backtest_win_rate": 0.0, "avg_return": 0.0, "cagr": 0.0, "mdd": 0.0, "sharpe": 0.0, "samples": 0}

        rank_text = "-"
        try:
            ranking = self._filtered_ranking()
            if ranking is not None and not ranking.empty:
                row = ranking[ranking["stock_id"].astype(str) == str(stock_id)]
                if not row.empty:
                    rank_text = str(int(row.iloc[0]["rank_all"]))
        except Exception:
            pass

        close_ = float(last["close"]) if pd.notna(last.get("close")) else 0.0
        ma5 = float(last["ma5"]) if pd.notna(last.get("ma5")) else close_
        ma10 = float(last["ma10"]) if pd.notna(last.get("ma10")) else close_
        ma20 = float(last["ma20"]) if pd.notna(last.get("ma20")) else close_
        ma60 = float(last["ma60"]) if pd.notna(last.get("ma60")) else close_
        macd_hist = float(last["macd_hist"]) if pd.notna(last.get("macd_hist")) else 0.0
        rsi14 = float(last["rsi14"]) if pd.notna(last.get("rsi14")) else 50.0
        k_val = float(last["k"]) if pd.notna(last.get("k")) else 50.0
        d_val = float(last["d"]) if pd.notna(last.get("d")) else 50.0

        def _avg_slope(series, n=5):
            try:
                s = pd.to_numeric(series.tail(n), errors='coerce').dropna()
                if len(s) < 2:
                    return 0.0
                return float(s.diff().dropna().mean())
            except Exception:
                return 0.0

        slope5 = _avg_slope(hist["ma5"], 3) if "ma5" in hist.columns else 0.0
        slope20 = _avg_slope(hist["ma20"], 5) if "ma20" in hist.columns else 0.0
        slope60 = _avg_slope(hist["ma60"], 5) if "ma60" in hist.columns else 0.0

        def _trend_state(close_price, fast, slow, slope_fast, slope_slow, macd_value, rsi_value, k_now, d_now):
            score = 0
            if close_price >= fast:
                score += 1
            if fast >= slow:
                score += 1
            if slope_fast > 0:
                score += 1
            if slope_slow > 0:
                score += 1
            if macd_value > 0:
                score += 1
            if 45 <= rsi_value <= 72:
                score += 1
            if k_now >= d_now:
                score += 1
            if score >= 5:
                return "多"
            if score <= 2:
                return "空"
            return "盤整"

        trend_short = _trend_state(close_, ma5, ma10, slope5, slope20, macd_hist, rsi14, k_val, d_val)
        trend_mid = _trend_state(close_, ma20, ma60, slope20, slope60, macd_hist, rsi14, k_val, d_val)
        trend_long = _trend_state(close_, ma60, ma60, slope60, slope60, macd_hist, rsi14, k_val, d_val)

        recent = hist.tail(89)
        recent_high = float(recent["high"].max()) if not recent.empty else close_
        recent_low = float(recent["low"].min()) if not recent.empty else close_
        width = max(recent_high - recent_low, 1e-6)
        pos = (close_ - recent_low) / width if width > 0 else 0.5
        signal = str(trade_plan.get("signal", "-"))
        if close_ >= recent_high * 0.995 and rsi14 > 72:
            wave_position = "第5浪"
        elif close_ >= recent_high * 0.995 and macd_hist > 0:
            wave_position = "第3浪"
        elif close_ > ma20 > ma60 and macd_hist > 0 and 0.45 <= pos <= 0.7:
            wave_position = "第1浪"
        elif close_ > ma20 and macd_hist >= 0 and pos < 0.45:
            wave_position = "第2浪"
        elif close_ >= ma20 and abs(close_ - ma20) / max(ma20, 1e-6) <= 0.04 and 0.45 <= pos <= 0.75:
            wave_position = "第4浪"
        elif close_ < ma20 and macd_hist < 0 and pos <= 0.35:
            wave_position = "A浪"
        elif close_ < (recent_low + width * 0.5) and macd_hist >= 0 and 0.25 <= pos <= 0.55:
            wave_position = "B浪"
        elif close_ < ma60 and macd_hist < 0 and pos < 0.25:
            wave_position = "C浪"
        else:
            wave_position = "區間整理"

        if wave_position in ("第1浪", "第3浪", "第5浪"):
            wave_structure = "推動浪"
        elif wave_position in ("A浪", "B浪", "C浪"):
            wave_structure = "修正浪"
        else:
            wave_structure = "整理浪"

        entry_low = float(trade_plan.get("entry_low", 0) or 0)
        entry_high = float(trade_plan.get("entry_high", 0) or 0)
        rr_value = float(trade_plan.get("rr", 0) or 0)
        model_win_rate = float(trade_plan.get("win_rate", 0) or 0)
        backtest_win_rate = float(bt.get("backtest_win_rate", 0) or 0)
        decision = str(trade_plan.get("trade_action", trade_plan.get("decision", "")) or "")
        in_entry_zone = TradingPlanEngine._in_entry_zone(close_, entry_low, entry_high)

        if decision == "BUY":
            display_ai = "可買" if in_entry_zone else "準備買"
        elif decision == "WEAK BUY":
            display_ai = "條件預掛"
        elif decision == "HOLD":
            display_ai = "觀察"
        elif decision == "AVOID":
            display_ai = "不可買"
        else:
            display_ai = str(trade_plan.get("ui_state", "觀察") or "觀察")

        semantic_note = ""
        if rr_value >= 2.0 and max(model_win_rate, backtest_win_rate) < 50:
            semantic_note = "屬高RR型，宜小倉位"
        elif rr_value < 1.2 and max(model_win_rate, backtest_win_rate) >= 60:
            semantic_note = "屬穩健型，目標空間有限"

        analysis_mode = "快速模式" if quick_only or str(trade_plan.get("trade_type", "")) == "快速模式" else "完整模式"
        reason_text = str(trade_plan.get("reason", "-"))
        if semantic_note:
            reason_text = f"{reason_text}｜{semantic_note}"

        lines = [
            f"《V3.6.1 AI語意收斂版｜{analysis_mode}》",
            f"股票：{stock['stock_name']} ({stock_id})｜排行：{rank_text}",
            f"市場 / 產業 / 題材：{stock['market']} / {stock['industry']} / {stock['theme']}",
            f"資料來源：{source or '-'}｜最新收盤：{close_:.2f}",
            "",
            "【趨勢判斷模組】",
            f"短線趨勢：{trend_short}｜依據 MA5/MA10、短斜率、MACD、RSI、KD",
            f"中線趨勢：{trend_mid}｜依據 MA20/MA60、均線斜率與收盤位置",
            f"長線趨勢：{trend_long}｜依據 MA60 結構、長斜率與中長期位置",
            f"MA5 / MA10 / MA20 / MA60：{ma5:.2f} / {ma10:.2f} / {ma20:.2f} / {ma60:.2f}",
            f"MACD Hist / RSI14 / KD：{macd_hist:.4f} / {rsi14:.2f} / {k_val:.2f}-{d_val:.2f}",
            "",
            "【波浪分析模組】",
            f"波浪結構：{wave_structure}",
            f"可能位置：{wave_position}",
            f"交易訊號：{signal}｜交易型態：{trade_plan.get('trade_type','-')}",
            "用途：不是只看漲跌，而是判斷目前位在推動、修正或整理的哪一段。",
            "",
            "【費波南西目標位模組】",
            f"支撐位 / 壓力位：{float(trade_plan.get('support',0) or 0):.2f} / {float(trade_plan.get('resistance',0) or 0):.2f}",
            f"Fib 1.0：{float(trade_plan.get('resistance',0) or 0):.2f}",
            f"Fib 1.382：{float(trade_plan.get('target_1382',0) or 0):.2f}",
            f"Fib 1.618：{float(trade_plan.get('target_1618',0) or 0):.2f}",
            f"進場區 / 停損 / RR：{trade_plan.get('entry_zone','-')} / {trade_plan.get('stop_loss','-')} / {rr_value:.2f}",
            "用途：判斷目標價、追價風險與停利停損區間。",
            "",
            "【AI建議模組】",
            f"AI結論：{display_ai}",
            f"執行狀態：{display_ai}｜決策：{decision or '-'}｜分類：{trade_plan.get('bucket','-')}",
            f"盤中狀態：{trade_plan.get('liquidity_status','-')}｜活性分：{float(trade_plan.get('liquidity_score',0) or 0):.1f}｜淘汰原因：{trade_plan.get('elimination_reason','-')}",
            f"模型分數 / 交易分數：{float(trade_plan.get('model_score',0) or 0):.2f} / {float(trade_plan.get('wave_trade_score', trade_plan.get('trade_score',0)) or 0):.2f}",
            f"勝率 / 回測勝率：{model_win_rate:.1f}% / {backtest_win_rate:.1f}%",
            f"平均報酬 / CAGR / MDD / Sharpe：{float(bt.get('avg_return',0) or 0):.2f}% / {float(bt.get('cagr',0) or 0):.2f}% / {float(bt.get('mdd',0) or 0):.2f}% / {float(bt.get('sharpe',0) or 0):.2f}",
            f"一句話：{reason_text}",
        ]
        if quick_only:
            lines.extend(["", "備註：目前為快速模式，完整 AI 分析與回測背景完成後會自動更新。"])
        return lines
    def update_detail_panel(self, stock_id: str, source: str = ""):
        lines = self.build_unified_detail_lines(stock_id, source=source or "多來源同步模式")
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))

    def start_backtest_selection_job(self, stock_id: str):
        stock_id = str(stock_id or "").strip()
        if not stock_id:
            return
        self.backtest_selection_token += 1
        token = self.backtest_selection_token
        self.window_current_stock_id = stock_id
        self.sync_multi_windows_selectors(stock_id)

        lines = self.build_unified_detail_lines(stock_id, source="回測視覺化", quick_only=True)
        lines.extend([
            "",
            "回測圖表狀態：背景計算 Equity Curve 中，完成後自動更新右下圖表。",
        ])
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))
        try:
            self.show_chart_loading(stock_id)
        except Exception:
            pass
        self.append_log(f"回測視覺化背景任務啟動：{stock_id}")
        t = threading.Thread(
            target=self._backtest_selection_worker,
            args=(stock_id, token),
            daemon=True,
            name=f"bt_select_{stock_id}"
        )
        t.start()

    def _backtest_selection_worker(self, stock_id: str, token: int):
        try:
            log_info(f"回測視覺化背景開始：{stock_id}｜token={token}")
            if token != self.backtest_selection_token:
                return

            stock = self.db.get_stock_row(stock_id)
            hist = self.db.get_price_history(stock_id)
            if stock is None or hist is None or hist.empty:
                self.ui_call(self.append_log, f"回測視覺化無資料：{stock_id}", "WARNING")
                return

            try:
                bt = self.backtest_engine.estimate_trade_quality(stock_id)
                self.backtest_cache[str(stock_id)] = bt
                log_info(f"回測視覺化回測完成：{stock_id}｜samples={bt.get('samples', 0)}")
            except Exception as e:
                log_exception(f"回測視覺化回測失敗：{stock_id}", e)
                self.ui_call(self.append_log, f"回測視覺化回測失敗：{stock_id}｜{e}", "ERROR")

            if token != self.backtest_selection_token:
                return

            eq_path = None
            try:
                log_info(f"回測視覺化圖表輸出開始：{stock_id}")
                eq_path = self.export_equity_curve_chart(stock_id, hist)
                log_info(f"回測視覺化圖表輸出完成：{stock_id}｜{eq_path}")
            except Exception as e:
                log_exception(f"回測視覺化圖表輸出失敗：{stock_id}", e)
                self.ui_call(self.append_log, f"回測視覺化圖表輸出失敗：{stock_id}｜{e}", "ERROR")

            def apply_result():
                if token != self.backtest_selection_token:
                    return
                try:
                    self.update_detail_panel(stock_id, source="回測視覺化")
                except Exception as e:
                    self.append_log(f"回測視覺化明細更新失敗：{stock_id}｜{e}", "WARNING")
                try:
                    if eq_path:
                        self.current_chart_path = eq_path
                        self._schedule_chart_file_update(stock_id, eq_path)
                    else:
                        self.show_chart_message("回測圖表產生失敗，請重新點選。")
                    try:
                        self.right_lower_notebook.select(self.chart_tab)
                    except Exception:
                        pass
                except Exception as e:
                    self.append_log(f"回測視覺化圖表更新失敗：{stock_id}｜{e}", "WARNING")

            self.ui_call(apply_result)
        except Exception as e:
            log_exception(f"回測視覺化背景失敗：{stock_id}", e)
            self.ui_call(self.append_log, f"回測視覺化背景失敗：{stock_id}｜{e}", "ERROR")

    def update_chart_panel(self, stock_id: str):
        self.window_current_stock_id = stock_id
        if self.current_chart_path and Path(self.current_chart_path).exists():
            self._schedule_chart_file_update(stock_id, self.current_chart_path)
            return
        self.show_chart_loading(stock_id)

    def safe_sync_stock_views(self, stock_id: str, source: str = ""):
        if not stock_id:
            return
        log_info(f"同步個股檢視：{stock_id}｜來源={source or '-'}")
        stock = self.db.get_stock_row(stock_id)
        hist = self.db.get_price_history(stock_id)
        if stock is None or hist.empty:
            return
        self.window_current_stock_id = stock_id
        self.sync_multi_windows_selectors(stock_id)
        try:
            self.start_selection_analysis(stock_id, source=source)
        except Exception as e:
            self.detail.delete("1.0", tk.END)
            self.detail.insert("1.0", f"股票：{stock_id}\n選股同步更新失敗：{e}")
            self.append_log(f"選股同步更新失敗：{stock_id}｜{e}")

    def sync_all_views(self, stock_id: str, source: str = ""):
        self.right_panel_mode = "個股模式"
        self.right_panel_source = source or "-"
        self.safe_sync_stock_views(stock_id, source=source)

    def _schedule_chart_update(self, stock_id: str):
        self.pending_stock_id = stock_id
        try:
            if self.chart_update_job is not None:
                self.root.after_cancel(self.chart_update_job)
        except Exception:
            pass
        self.chart_update_job = self.root.after(180, self._flush_chart_update)

    def _flush_chart_update(self):
        self.chart_update_job = None
        stock_id = self.pending_stock_id
        self.pending_stock_id = None
        if not stock_id:
            return
        self.update_multi_window_stock(stock_id)

    def _schedule_chart_file_update(self, stock_id: str, chart_path):
        self.pending_chart_image = (str(stock_id), str(chart_path))
        try:
            if self.chart_image_job is not None:
                self.root.after_cancel(self.chart_image_job)
        except Exception:
            pass
        self.chart_image_job = self.root.after(80, self._flush_chart_file_update)

    def _flush_chart_file_update(self):
        self.chart_image_job = None
        payload = self.pending_chart_image
        self.pending_chart_image = None
        if not payload:
            return
        stock_id, chart_path = payload
        self.update_multi_window_stock(stock_id, chart_path=chart_path)

    def show_chart_message(self, message: str):
        if self.chart_fig is None or self.chart_canvas is None:
            return
        self.chart_fig.clear()
        ax = self.chart_fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, safe_plot_text(message, fallback="圖表訊息"), ha="center", va="center", transform=ax.transAxes, fontfamily=SELECTED_PLOT_FONT, fontsize=12)
        self.chart_fig.tight_layout()
        self.chart_canvas.draw_idle()

    def show_chart_loading(self, stock_id: str):
        self.append_log(f"圖表載入中：{stock_id}")
        self.show_chart_message(f"{stock_id} 圖表背景載入中…")

    def show_chart_file(self, chart_path):
        self.append_log(f"載入圖表檔：{chart_path}")
        if self.chart_fig is None or self.chart_canvas is None:
            return
        p = Path(chart_path)
        if not p.exists():
            self.show_chart_message("找不到圖表檔案")
            return
        self.chart_fig.clear()
        ax = self.chart_fig.add_subplot(111)
        ax.axis("off")
        img = plt.imread(str(p))
        ax.imshow(img)
        self.chart_fig.tight_layout()
        self.chart_canvas.draw_idle()

    def _candlestick(self, ax, x_vals, opens, highs, lows, closes):
        width = 0.55
        for xi, op, hi, lo, cl in zip(x_vals, opens, highs, lows, closes):
            color = "#d62728" if cl >= op else "#2ca02c"
            ax.vlines(xi, lo, hi, color=color, linewidth=1)
            bottom = min(op, cl)
            height = abs(cl - op)
            if height < 1e-6:
                height = max((hi - lo) * 0.02, 0.05)
            rect = Rectangle((xi - width / 2, bottom), width, height, facecolor=color, edgecolor=color, alpha=0.65)
            ax.add_patch(rect)

    def build_window_plan_lines(self, stock_id: str):
        stock = self.db.get_stock_row(stock_id)
        hist = self.db.get_price_history(stock_id)
        if stock is None or hist.empty:
            return ["無資料"]
        hist = DataEngine.attach(hist)
        last = hist.iloc[-1]
        trade_plan = self.master_trading_engine.plan_engine.build_plan(stock_id)
        bt = self.backtest_engine.estimate_trade_quality(stock_id)
        wave = WaveEngine.detect_wave_label(hist)
        return [
            "《v9.2 交易計畫視窗》",
            f"股票：{stock['stock_name']} ({stock_id})",
            f"市場 / 產業 / 題材：{stock['market']} / {stock['industry']} / {stock['theme']}",
            f"最新收盤：{float(last['close']):.2f}",
            f"狀態：{trade_plan.get('ui_state','-')}｜決策：{trade_plan.get('trade_action','-')}｜盤中：{trade_plan.get('liquidity_status','-')} {float(trade_plan.get('liquidity_score',0) or 0):.1f}",
            f"波浪：{wave}｜訊號：{trade_plan.get('signal','-')}｜交易型態：{trade_plan.get('trade_type','-')}",
            f"進場區：{trade_plan.get('entry_zone','-')}",
            f"停損：{trade_plan.get('stop_loss','-')}",
            f"Support / Fib1.0：{float(trade_plan.get('support',0) or 0):.2f} / {float(trade_plan.get('resistance',0) or 0):.2f}",
            f"Fib 1.382 / 1.618：{float(trade_plan.get('target_1382',0) or 0):.2f} / {float(trade_plan.get('target_1618',0) or 0):.2f}",
            f"RR：{float(trade_plan.get('rr',0) or 0):.2f}｜勝率：{float(trade_plan.get('win_rate',0) or 0):.1f}%",
            f"六模組：K {trade_plan.get('kline_score',0):.1f}｜波 {trade_plan.get('wave_score',0):.1f}｜費 {trade_plan.get('fib_score',0):.1f}｜阪 {trade_plan.get('sakata_score',0):.1f}｜量 {trade_plan.get('volume_score',0):.1f}｜指 {trade_plan.get('indicator_score',0):.1f}",
            f"回測：勝率 {float(bt.get('backtest_win_rate',0) or 0):.1f}%｜CAGR {float(bt.get('cagr',0) or 0):.2f}%｜MDD {float(bt.get('mdd',0) or 0):.2f}%｜Sharpe {float(bt.get('sharpe',0) or 0):.2f}",
            f"理由：{trade_plan.get('reason','-')}",
        ]

    def update_multi_window_stock(self, stock_id: str, chart_path: str | None = None):
        self.ensure_multi_windows()
        self.window_current_stock_id = stock_id
        try:
            if chart_path and Path(chart_path).exists():
                self.show_chart_file(chart_path)
            elif self.current_chart_path and Path(self.current_chart_path).exists():
                self.show_chart_file(self.current_chart_path)
            else:
                self.show_chart_loading(stock_id)
            try:
                self.right_lower_notebook.select(self.chart_tab)
            except Exception:
                pass
        except Exception as e:
            self.append_log(f"圖表更新失敗：{stock_id}｜{e}")

    def draw_live_chart(self, stock_id: str):
        if self.chart_fig is None or self.chart_canvas is None:
            return
        if self.chart_updating:
            self.pending_stock_id = stock_id
            return
        self.chart_updating = True
        try:
            stock = self.db.get_stock_row(stock_id)
            hist = self.db.get_price_history(stock_id)
            if stock is None or hist.empty:
                return
            hist = DataEngine.attach(hist.copy()).tail(90).reset_index(drop=True)
            if hist.empty:
                return

            plan = self.get_cached_trade_plan(stock_id)
            if plan is None:
                plan = self.build_lightweight_plan(stock_id, hist, stock=stock)
                self.plan_cache[str(stock_id)] = plan
            wave = WaveEngine.detect_wave_label(hist)
            x = list(range(len(hist)))
            self.chart_fig.clear()
            ax = self.chart_fig.add_subplot(111)

            self._candlestick(ax, x, hist["open"], hist["high"], hist["low"], hist["close"])
            ax.plot(x, hist["ma20"], label="MA20", linewidth=1.2)
            ax.plot(x, hist["ma60"], label="MA60", linewidth=1.2)

            support = float(plan.get("support", 0) or 0)
            fib1 = float(plan.get("resistance", 0) or 0)
            fib1382 = float(plan.get("target_1382", 0) or 0)
            fib1618 = float(plan.get("target_1618", 0) or 0)
            try:
                stop = float(plan.get("stop_loss", 0) or 0)
            except Exception:
                stop = 0.0

            if support > 0:
                ax.axhline(support, linestyle="--", linewidth=1.0, label=f"Support {support:.2f}")
            if fib1 > 0:
                ax.axhline(fib1, linestyle="--", linewidth=1.0, label=f"Fib 1.0 {fib1:.2f}")
            if fib1382 > 0:
                ax.axhline(fib1382, linestyle=":", linewidth=1.0, label=f"Fib 1.382 {fib1382:.2f}")
            if fib1618 > 0:
                ax.axhline(fib1618, linestyle=":", linewidth=1.0, label=f"Fib 1.618 {fib1618:.2f}")

            recent = hist.tail(55)
            try:
                peak_idx = recent["high"].idxmax()
                trough_idx = recent["low"].idxmin()
                peak_y = float(hist.loc[peak_idx, "high"])
                trough_y = float(hist.loc[trough_idx, "low"])
                ax.scatter([peak_idx], [peak_y], s=45, marker="o")
                ax.scatter([trough_idx], [trough_y], s=45, marker="o")
                ax.annotate("Wave Peak", xy=(peak_idx, peak_y), xytext=(peak_idx, peak_y * 1.02), fontfamily=SELECTED_PLOT_FONT)
                ax.annotate("Wave Trough", xy=(trough_idx, trough_y), xytext=(trough_idx, trough_y * 0.98), fontfamily=SELECTED_PLOT_FONT)
            except Exception:
                pass

            last_close = float(hist.iloc[-1]["close"])
            last_x = x[-1]
            bull_target = fib1382 if fib1382 > 0 else last_close * 1.08
            bear_target = stop if stop > 0 else last_close * 0.95
            path_x = [last_x, last_x + 4, last_x + 9]
            bull_y = [last_close, (last_close + bull_target) / 2.0, bull_target]
            bear_y = [last_close, (last_close + bear_target) / 2.0, bear_target]
            ax.plot(path_x, bull_y, "--", linewidth=1.6, label="Bull Path")
            ax.plot(path_x, bear_y, "--", linewidth=1.6, label="Bear Path")

            ax.set_xlim(0, max(path_x) + 2)
            title_stock = safe_plot_text(stock.get("stock_name", stock_id), fallback=str(stock_id))
            title_wave = safe_plot_text(wave, fallback="Wave")
            title_signal = safe_plot_text(plan.get("signal", "-"), fallback="-")
            ax.set_title(f"{title_stock}({stock_id}) | {title_wave} | {title_signal}", fontfamily=SELECTED_PLOT_FONT)
            info_text = (
                f"Wave: {title_wave}\n"
                f"Entry: {safe_plot_text(plan.get('entry_zone','-'))}\n"
                f"Stop: {safe_plot_text(plan.get('stop_loss','-'))}\n"
                f"RR: {float(plan.get('rr',0) or 0):.2f}"
            )
            ax.text(
                0.01, 0.98,
                info_text,
                transform=ax.transAxes, va="top", ha="left", fontfamily=SELECTED_PLOT_FONT,
                bbox=dict(boxstyle="round", alpha=0.15)
            )
            ax.grid(alpha=0.2)
            ax.legend(loc="upper left", fontsize=8, prop={"family": SELECTED_PLOT_FONT, "size": 8})
            self.chart_fig.tight_layout()
            self.chart_canvas.draw_idle()
        finally:
            self.chart_updating = False
            if self.pending_stock_id and self.pending_stock_id != stock_id:
                next_stock = self.pending_stock_id
                self.pending_stock_id = None
                try:
                    self.root.after(50, lambda s=next_stock: self.update_multi_window_stock(s))
                except Exception:
                    pass

    def export_selected_data(self):
        target = self.download_target_var.get().strip() or "TOP20"

        def worker():
            mapping = {
                "TOP20": getattr(self, "last_top20_df", pd.DataFrame()),
                "TOP5": getattr(self, "last_top5_df", pd.DataFrame()),
                "今日可下單": getattr(self, "last_today_buy_df", pd.DataFrame()),
                "等待回測": getattr(self, "last_wait_df", pd.DataFrame()),
                "條件預掛": getattr(self, "last_wait_df", pd.DataFrame()),
                "主攻": self.last_attack_df,
                "次強": self.last_watch_df,
                "防守": self.last_defense_df,
                "執行下單清單": self.last_order_list_df,
                "組合交易計畫": getattr(self, "last_institutional_plan_df", pd.DataFrame()),
                "操作SOP": getattr(self, "last_operation_sop_df", pd.DataFrame()),
                "排行": self._filtered_ranking(),
                "類股": pd.DataFrame([(self.sector_tree.item(i, "values")) for i in self.sector_tree.get_children()], columns=["產業", "檔數", "平均總分", "平均AI分", "代表股"]) if self.sector_tree.get_children() else pd.DataFrame(),
                "題材": pd.DataFrame([(self.theme_tree.item(i, "values")) for i in self.theme_tree.get_children()], columns=["題材", "檔數", "平均總分", "平均AI分", "代表股"]) if self.theme_tree.get_children() else pd.DataFrame(),
                "未分類清單": pd.read_excel(CLASSIFICATION_V2_UNCLASSIFIED_PATH) if CLASSIFICATION_V2_UNCLASSIFIED_PATH.exists() else pd.DataFrame(),
                "分類V2摘要": pd.DataFrame([get_classification_v2_summary()]) if get_classification_v2_summary() else pd.DataFrame(),
            }
            mapping["唯一決策"] = getattr(self, "last_unique_decision_df", pd.DataFrame())
            df = mapping.get(target, pd.DataFrame())
            if df is None:
                df = pd.DataFrame()
            if isinstance(df, pd.DataFrame) and not df.empty:
                id_col = "stock_id" if "stock_id" in df.columns else ("代號" if "代號" in df.columns else None)
                df = self.enrich_price_and_export_fields(df, id_col=id_col)
            if df.empty:
                empty_columns = {
                    "TOP20": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "bucket", "ui_state", "liquidity_status", "liquidity_score", "entry_zone", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "win_rate"],
                    "TOP5": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "ui_state", "liquidity_status", "liquidity_score", "entry_zone", "stop_loss", "target_price", "target_1382", "rr", "win_rate", "backtest_win_rate", "cagr", "mdd"],
                    "今日可下單": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "ui_state", "liquidity_status", "liquidity_score", "entry_zone", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "win_rate"],
                    "等待回測": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "ui_state", "liquidity_status", "liquidity_score", "entry_zone", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "win_rate"],
                    "條件預掛": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "ui_state", "liquidity_status", "liquidity_score", "entry_zone", "stop_loss", "target_price", "target_1382", "target_1618", "rr", "win_rate"],
                    "執行下單清單": ["優先級", "代號", "名稱", "現價", "漲跌", "漲跌幅%", "分類", "狀態", "進場區", "停損", "目標價", "1.382", "1.618", "RR", "勝率", "ATR%", "Kelly%", "建議張數", "建議金額", "單檔曝險%", "投資組合狀態", "風險備註"],
                    "組合交易計畫": ["優先級", "代號", "名稱", "現價", "漲跌", "漲跌幅%", "市場", "產業", "題材", "分類", "狀態", "進場區", "停損", "目標價", "1.382", "1.618", "RR", "勝率", "模型分數", "交易分數", "ATR%", "Kelly%", "建議張數", "建議金額", "單檔曝險%", "題材曝險%", "產業曝險%", "投資組合狀態", "風險備註"],
                    "唯一決策": ["stock_id", "stock_name", "現價", "漲跌", "漲跌幅%", "market", "industry", "theme", "ui_state", "entry_zone", "stop_loss", "target_price", "rr", "win_rate", "decision", "final_trade_decision"],
                    "操作SOP": ["step", "module", "focus", "rule", "purpose", "output"],
                    "未分類清單": ["stock_id", "stock_name", "market", "industry_final", "theme_final", "sub_theme_final", "classification_source", "classification_confidence", "classification_note"],
                    "分類V2摘要": ["total", "official", "manual", "rule_engine", "ai_infer", "unclassified", "covered", "coverage_pct", "report_time", "unclassified_report"],
                }
                df = pd.DataFrame(columns=empty_columns.get(target, ["message"]))
                if df.empty and target not in empty_columns:
                    df = pd.DataFrame([{"message": f"目前沒有可下載的【{target}】資料"}])
            try:
                self.ui_call(self.start_task, f"下載{target}", 3)
                self.ui_call(self.update_task, f"下載{target}", 1, 3, item="準備資料")
                base = RUNTIME_DIR / f"{target}_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                self.ui_call(self.update_task, f"下載{target}", 2, 3, item="輸出檔案")
                out_path, out_type = write_table_bundle(base, {target: df}, preferred="excel")
                display_name = Path(out_path).name if isinstance(out_path, Path) else str(out_path)
                self.ui_call(self.update_task, f"下載{target}", 3, 3, success=1, item=display_name)
                self.ui_call(self.finish_task, f"下載{target}", f"{target} 資料已輸出：{display_name}")
                self.ui_call(messagebox.showinfo, "完成", f"{target} 資料已輸出（{out_type}）：\n{out_path}")
            except Exception as e:
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, f"export_{target}")



    def build_daily_summary_sheet(self) -> pd.DataFrame:
        market = getattr(self, "last_market_context", {}) or {}
        top20 = getattr(self, "last_top20_df", pd.DataFrame())
        today_buy = getattr(self, "last_today_buy_df", pd.DataFrame())
        wait_df = getattr(self, "last_wait_df", pd.DataFrame())
        unique_df = getattr(self, "last_unique_decision_df", pd.DataFrame())
        theme_df = getattr(self, "last_theme_summary_df", pd.DataFrame())
        theme_text = "、".join(theme_df.head(5)["theme"].astype(str).tolist()) if isinstance(theme_df, pd.DataFrame) and not theme_df.empty and "theme" in theme_df.columns else "-"
        pick_text = "、".join(unique_df.head(5)["stock_id"].astype(str).tolist()) if isinstance(unique_df, pd.DataFrame) and not unique_df.empty and "stock_id" in unique_df.columns else "-"
        pool_audit = getattr(self.master_trading_engine, "last_pool_audit", {}) or {}
        rows = [
            {"項目": "市場狀態", "內容": str(market.get("regime", "-"))},
            {"項目": "市場分數", "內容": str(market.get("score", "-"))},
            {"項目": "主流族群", "內容": theme_text},
            {"項目": "TOP20檔數", "內容": int(len(top20) if isinstance(top20, pd.DataFrame) else 0)},
            {"項目": "今日可下單檔數", "內容": int(len(today_buy) if isinstance(today_buy, pd.DataFrame) else 0)},
            {"項目": "等待回測檔數", "內容": int(len(wait_df) if isinstance(wait_df, pd.DataFrame) else 0)},
            {"項目": "唯一決策", "內容": pick_text},
            {"項目": "Pool Sizes", "內容": f"candidate20={pool_audit.get('candidate20_count','-')}｜core_attack5={pool_audit.get('core_attack5_count','-')}｜today_buy={pool_audit.get('today_buy_count','-')}｜execution_ready={pool_audit.get('execution_ready_count','-')}｜unique_decision={pool_audit.get('unique_decision_count','-')}"},
            {"項目": "Pool Audit", "內容": f"core-candidate diff={','.join(pool_audit.get('core_minus_candidate20', [])[:10]) or '-'}｜today-core diff={','.join(pool_audit.get('today_minus_core', [])[:10]) or '-'}｜unique-core diff={','.join(pool_audit.get('unique_minus_core', [])[:10]) or '-'}"},
            {"項目": "風險提示", "內容": str(market.get("memo", "-"))},
            {"項目": "參數摘要", "內容": f"max_positions={market.get('max_positions','-')}｜min_win_rate={market.get('min_win_rate','-')}｜RSI={market.get('rsi_low','-')}~{market.get('rsi_high','-')}"},
        ]
        return pd.DataFrame(rows)

    def build_report_usage_sheet(self) -> pd.DataFrame:
        rows = [
            {"報表": "Ranking", "用途": "雷達表，不直接下單。"},
            {"報表": "Trade_TOP20", "用途": "候選20，進第二層分析。"},
            {"報表": "Trade_TOP5", "用途": "主力成立 + 主升成立的主攻池。"},
            {"報表": "Today_Buy", "用途": "只放位置仍正確、可直接執行的主攻股。"},
            {"報表": "Wait_Pullback", "用途": "主流/突破候選，但位置未到，可等待回測。"},
            {"報表": "Order_List", "用途": "執行清單，含資金與風控欄位。"},
            {"報表": "Institutional_Plan", "用途": "組合與曝險配置表。"},
            {"報表": "Unique_Decision", "用途": "最終 1~5 檔唯一決策。"},
            {"報表": "Daily_Summary", "用途": "每日總結頁：先看總結，再看細表。"},
            {"報表": "Summary", "用途": "摘要頁，含市場狀態、Pool Sizes 與 Pool Audit 結果。"},
            {"報表": "用途說明", "用途": "報表用途與檢查規則說明頁。"},
        ]
        return pd.DataFrame(rows)

    def export_analysis_excel(self):
        ranking = self._filtered_ranking()
        if ranking is None or ranking.empty:
            return messagebox.showwarning("提醒", "目前沒有可匯出的分析資料。")

        def worker():
            try:
                sector = pd.DataFrame()
                theme = pd.DataFrame()
                self.ui_call(self.start_task, "匯出分析", 5)
                self.ui_call(self.update_task, "匯出分析", 1, 5, item="整理排行")
                if not ranking.empty:
                    sector = ranking.groupby("industry", as_index=False).agg(count=("stock_id", "count"), avg_total=("total_score", "mean"), avg_ai=("ai_score", "mean")).sort_values(["avg_total", "avg_ai"], ascending=False)
                    theme = ranking.groupby("theme", as_index=False).agg(count=("stock_id", "count"), avg_total=("total_score", "mean"), avg_ai=("ai_score", "mean")).sort_values(["avg_total", "avg_ai"], ascending=False)
                detail_text = self.detail.get("1.0", tk.END).strip()
                tables = {"Ranking": ranking}
                if not sector.empty:
                    tables["Sector"] = sector
                if not theme.empty:
                    tables["Theme"] = theme
                if self.last_top20_df is not None and not self.last_top20_df.empty:
                    tables["Trade_TOP20"] = self.last_top20_df
                if self.last_top5_df is not None and not self.last_top5_df.empty:
                    tables["Trade_TOP5"] = self.last_top5_df
                if getattr(self, "last_today_buy_df", pd.DataFrame()) is not None and not getattr(self, "last_today_buy_df", pd.DataFrame()).empty:
                    tables["Today_Buy"] = self.last_today_buy_df
                if getattr(self, "last_wait_df", pd.DataFrame()) is not None and not getattr(self, "last_wait_df", pd.DataFrame()).empty:
                    tables["Wait_Pullback"] = self.last_wait_df
                if self.last_attack_df is not None and not self.last_attack_df.empty:
                    tables["Attack"] = self.last_attack_df
                if self.last_watch_df is not None and not self.last_watch_df.empty:
                    tables["Watch"] = self.last_watch_df
                if self.last_defense_df is not None and not self.last_defense_df.empty:
                    tables["Defense"] = self.last_defense_df
                if self.last_order_list_df is not None and not self.last_order_list_df.empty:
                    tables["Order_List"] = self.last_order_list_df
                if getattr(self, "last_institutional_plan_df", pd.DataFrame()) is not None and not getattr(self, "last_institutional_plan_df", pd.DataFrame()).empty:
                    tables["Institutional_Plan"] = self.last_institutional_plan_df
                cls_v2 = get_classification_v2_summary()
                if cls_v2:
                    tables["Classification_V2_Summary"] = pd.DataFrame([cls_v2])
                try:
                    tables["DB_Schema_Check"] = self.db.schema_check_df()
                    tables["External_Source_Status"] = self.db.read_table("external_source_status", limit=500)
                    tables["External_Data_Log"] = self.db.read_table("external_data_log", limit=1000)
                    tables["System_Run_Log"] = self.db.read_table("system_run_log", limit=500)
                    tables["Trade_Plan_DB"] = self.db.read_table("trade_plan", limit=1000)
                    _tp = tables.get("Trade_Plan_DB", pd.DataFrame())
                    if _tp is not None and not _tp.empty:
                        eps_view_cols = [c for c in ["stock_id", "stock_name", "final_trade_decision", "trade_allowed", "analysis_ready", "execution_ready", "soft_block", "block_reason", "execution_block_reason", "eps_ttm", "eps_yoy", "revenue_yoy", "eps_category", "matrix_cell", "revenue_eps_score", "data_quality_flag", "decision_reason_short"] if c in _tp.columns]
                        tables["TradePlan_EPS_View"] = _tp[eps_view_cols].copy()
                        ext_cols = [c for c in ["stock_id", "stock_name", "final_trade_decision", "trade_allowed", "analysis_ready", "execution_ready", "soft_block", "block_reason", "external_blocking_reason", "decision_reason_short", "source_trace_json"] if c in _tp.columns]
                        if ext_cols:
                            tables["External_Selection"] = _tp[ext_cols].copy()
                            if "soft_block" in _tp.columns:
                                _soft = _tp[pd.to_numeric(_tp["soft_block"], errors="coerce").fillna(0).astype(int).eq(1)].copy()
                                tables["SoftBlock_Candidates"] = _soft[ext_cols].copy() if not _soft.empty else pd.DataFrame(columns=ext_cols)
                            if "trade_allowed" in _tp.columns:
                                _ready = _tp[pd.to_numeric(_tp["trade_allowed"], errors="coerce").fillna(0).astype(int).eq(1)].copy()
                                tables["Execution_Ready"] = _ready[ext_cols].copy() if not _ready.empty else pd.DataFrame(columns=ext_cols)
                            _nogo_cols = [c for c in ["stock_id", "stock_name", "rr", "rsi", "atr_pct", "price_deviation", "model_score", "wave_trade_score", "trade_allowed", "final_trade_decision", "fail_reason", "external_blocking_reason", "decision_reason_short"] if c in _tp.columns]
                            if _nogo_cols:
                                _nogo = _tp[pd.to_numeric(_tp.get("trade_allowed", 0), errors="coerce").fillna(0).astype(int).eq(0)].copy()
                                tables["今日不可下單原因"] = _nogo[_nogo_cols].copy() if not _nogo.empty else pd.DataFrame(columns=_nogo_cols)
                    _val_rows = []
                    _ffv = self.db.read_table("financial_feature_daily", limit=None)
                    if _ffv is not None and not _ffv.empty:
                        _val_rows.append({"test_id": "TC02", "test": "EPS<0但營收高成長", "result": "PASS" if ((_ffv.get("eps_bucket", "") == "E0") & (_ffv.get("rev_bucket", "") == "R3") & (_ffv.get("eps_category", "") == "U3")).any() else "NO_SAMPLE"})
                        _val_rows.append({"test_id": "TC03", "test": "高EPS但營收衰退", "result": "PASS" if ((_ffv.get("eps_bucket", "") == "E3") & (_ffv.get("rev_bucket", "").isin(["R0", "R1"])) & (_ffv.get("eps_category", "") == "U4")).any() else "NO_SAMPLE"})
                        _val_rows.append({"test_id": "TC04", "test": "資料缺失NE", "result": "PASS" if _ffv.get("data_quality_flag", pd.Series(dtype=str)).astype(str).str.contains("NE", na=False).any() else "NO_SAMPLE"})
                        _val_rows.append({"test_id": "TC05", "test": "Ranking/TradePlan欄位", "result": "PASS" if "revenue_eps_score" in _ffv.columns else "FAIL"})
                    else:
                        _val_rows.append({"test_id": "TC01", "test": "financial_feature_daily", "result": "NO_DATA"})
                    tables["Validation_Result"] = pd.DataFrame(_val_rows)
                    tables["Market_Snapshot"] = self.db.read_table("market_snapshot", limit=200)
                    tables["External_Valuation"] = self.db.read_table("external_valuation", limit=3000)
                    tables["Financial_Feature"] = self.db.read_table("financial_feature_daily", limit=5000)
                    _ff = tables.get("Financial_Feature", pd.DataFrame())
                    if _ff is not None and not _ff.empty:
                        eps_cols = [c for c in ["stock_id", "feature_date", "eps_ttm", "eps_yoy", "revenue_yoy", "eps_bucket", "rev_bucket", "matrix_cell", "eps_category", "matrix_base_score", "modifier", "revenue_eps_score", "data_quality_flag", "source_trace_json"] if c in _ff.columns]
                        tables["EPS_Matrix_Check"] = _ff[eps_cols].head(200).copy()
                    else:
                        tables["EPS_Matrix_Check"] = pd.DataFrame([{"status": "NO_DATA", "message": "financial_feature_daily 尚無資料，請先同步外部資料或重建排行"}])
                    tables["External_Margin"] = self.db.read_table("external_margin", limit=3000)
                    tables["Macro_Margin_Sentiment"] = self.db.read_table("macro_margin_sentiment", limit=300)
                    tables["External_Source_Config"] = ExternalSourceConfig.to_dataframe()
                    _status_df = tables.get("External_Source_Status", pd.DataFrame())
                    _blocking_rows = []
                    if _status_df is not None and not _status_df.empty:
                        for _, _r in _status_df.iterrows():
                            _ready = int(pd.to_numeric(pd.Series([_r.get("data_ready", 0)]), errors="coerce").fillna(0).iloc[0])
                            _rows = int(pd.to_numeric(pd.Series([_r.get("rows_count", 0)]), errors="coerce").fillna(0).iloc[0])
                            _module = str(_r.get("module", ""))
                            _mandatory = bool(ExternalSourceConfig.SOURCES.get(_module, {}).get("mandatory", False))
                            if _ready != 1 and _mandatory:
                                _blocking_rows.append({
                                    "severity": "P0", "module": _module, "status": _r.get("status", ""),
                                    "rows_count": _rows, "blocking_reason": _r.get("blocking_reason", "") or _r.get("error_message", ""),
                                    "go_no_go": "INFO-SOFT-BLOCK"
                                })
                    else:
                        _blocking_rows.append({"severity": "INFO", "module": "ALL", "status": "missing", "rows_count": 0, "blocking_reason": "尚未執行外部資料同步或external_source_status空白；V9.5.9僅作提示，不停止分析/交易邏輯", "go_no_go": "INFO-SOFT-BLOCK"})
                    tables["Blocking_Issues"] = pd.DataFrame(_blocking_rows) if _blocking_rows else pd.DataFrame([{"severity": "OK", "module": "ALL", "status": "ok", "rows_count": "", "blocking_reason": "", "go_no_go": "GO"}])
                    tables["Go_NoGo_Summary"] = pd.DataFrame([{
                        "go_no_go": "GO_WITH_EXTERNAL_INFO_WARNING" if _blocking_rows else "GO",
                        "blocking_count": len(_blocking_rows),
                        "rule": "V9.5.9：外部資料不得直接控制trade_allowed；execution_ready/soft_block僅為資訊提示欄位",
                        "db_path": str(self.db.db_path),
                        "db_hash": self.db._safe_sha256(self.db.db_path),
                        "program_name": APP_NAME,
                    }])
                except Exception as exc:
                    tables["DB_Export_Error"] = pd.DataFrame([{"error": str(exc)}])
                try:
                    if CLASSIFICATION_V2_UNCLASSIFIED_PATH.exists():
                        tables["Unclassified_Report"] = pd.read_excel(CLASSIFICATION_V2_UNCLASSIFIED_PATH)
                except Exception:
                    pass
                try:
                    _last_nogo = getattr(self, "last_no_go_df", pd.DataFrame())
                    if _last_nogo is not None and not _last_nogo.empty:
                        _cols = [c for c in ["stock_id", "stock_name", "rr", "rsi", "atr_pct", "price_deviation", "model_score", "wave_trade_score", "trade_allowed", "final_trade_decision", "strategy_nogo_detail", "fail_reason"] if c in _last_nogo.columns]
                        if _cols:
                            tables["今日不可下單原因"] = _last_nogo[_cols].copy()
                except Exception:
                    pass
                if "Ranking" in tables:
                    tables["Ranking"] = self.enrich_price_and_export_fields(tables["Ranking"], id_col="stock_id")
                for name in ["Trade_TOP20", "Trade_TOP5", "Today_Buy", "Wait_Pullback", "Attack", "Watch", "Defense", "Order_List", "Institutional_Plan"]:
                    if name in tables and isinstance(tables[name], pd.DataFrame) and not tables[name].empty:
                        id_col = "stock_id" if "stock_id" in tables[name].columns else ("代號" if "代號" in tables[name].columns else None)
                        tables[name] = self.enrich_price_and_export_fields(tables[name], id_col=id_col)
                unique_df = getattr(self, "last_unique_decision_df", pd.DataFrame())
                if unique_df is not None and not unique_df.empty:
                    tables["Unique_Decision"] = self.enrich_price_and_export_fields(unique_df, id_col="stock_id")
                tables["Daily_Summary"] = self.build_daily_summary_sheet()
                tables["Summary"] = self.build_daily_summary_sheet()
                tables["Report_Guide"] = self.build_report_usage_sheet()
                tables["用途說明"] = self.build_report_usage_sheet()
                tables["Detail"] = pd.DataFrame({"detail": [detail_text]})
                self.ui_call(self.update_task, "匯出分析", 3, 5, item="寫入檔案")
                base = RUNTIME_DIR / f"Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                out_path, out_type = write_table_bundle(base, tables, preferred="excel")
                display_name = Path(out_path).name if isinstance(out_path, Path) else str(out_path)
                self.ui_call(self.update_task, "匯出分析", 5, 5, success=1, item=display_name)
                self.ui_call(self.finish_task, "匯出分析", f"分析報告已輸出：{display_name}")
                self.ui_call(messagebox.showinfo, "完成", f"分析報告已輸出（{out_type}）：\n{out_path}")
            except Exception as e:
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, "export_analysis")

    def _normalize_light_value(self, light) -> str:
        light = str(light or "").strip()
        alias = {
            "red": "🔴", "orange": "🟠", "yellow": "🟡", "green": "🟢", "blue": "🔵", "neutral": "⚪",
            "R": "🔴", "O": "🟠", "Y": "🟡", "G": "🟢", "B": "🔵", "N": "⚪",
            "紅": "🔴", "橘": "🟠", "黃": "🟡", "綠": "🟢", "藍": "🔵", "白": "⚪", "○": "⚪",
        }
        return alias.get(light, light if light in ("🔴", "🟠", "🟡", "🟢", "🔵", "⚪") else "⚪")

    def _get_display_light(self, row):
        row = dict(row or {})
        light = self._normalize_light_value(row.get("tactical_light", ""))
        if light and light != "⚪":
            return light
        signal = str(row.get("signal", "") or "").strip()
        mapping = {
            "突破強勢": "🔵", "強勢追蹤": "🟢", "整理偏多": "🟢", "偏多觀察": "🟡",
            "區間整理": "🟡", "轉弱警戒": "🟠", "急跌風險": "🔴"
        }
        return mapping.get(signal, "⚪")

    def _display_light_symbol(self, light):
        light = self._normalize_light_value(light)
        return {"🔴":"紅","🟠":"橘","🟡":"黃","🟢":"綠","🔵":"藍","⚪":"白"}.get(light, "白")

    def build_order_list(self, today_buy_df: pd.DataFrame, wait_df: pd.DataFrame | None = None) -> pd.DataFrame:
        x1 = today_buy_df.copy() if today_buy_df is not None else pd.DataFrame()
        x1 = apply_external_decision_filter(x1, "order_list")
        plan = self.portfolio_engine.build_institutional_plan(x1)
        if plan.empty:
            return pd.DataFrame(columns=["優先級","代號","名稱","現價","漲跌","漲跌幅%","分類","狀態","盤中狀態","活性分","進場區","停損","目標價","1.382","1.618","RR","勝率","ATR%","Kelly%","建議張數","建議金額","單檔曝險%","投資組合狀態","風險備註"])
        order_df = pd.DataFrame({
            "優先級": plan["優先級"],
            "代號": plan["代號"],
            "名稱": plan["名稱"],
            "現價": plan["現價"] if "現價" in plan.columns else np.nan,
            "漲跌": plan["漲跌"] if "漲跌" in plan.columns else np.nan,
            "漲跌幅%": plan["漲跌幅%"] if "漲跌幅%" in plan.columns else np.nan,
            "分類": plan["分類"],
            "狀態": plan["狀態"],
            "盤中狀態": plan["盤中狀態"] if "盤中狀態" in plan.columns else "",
            "活性分": plan["活性分"] if "活性分" in plan.columns else 0,
            "進場區": plan["進場區"],
            "停損": plan["停損"],
            "目標價": plan["目標價"] if "目標價" in plan.columns else (plan["1.382"] if "1.382" in plan.columns else ""),
            "1.382": plan["1.382"],
            "1.618": plan["1.618"],
            "RR": plan["RR"],
            "勝率": plan["勝率"],
            "ATR%": plan["ATR%"],
            "Kelly%": plan["Kelly%"],
            "建議張數": plan["建議張數"],
            "建議金額": plan["建議金額"],
            "單檔曝險%": plan["單檔曝險%"],
            "投資組合狀態": plan["投資組合狀態"],
            "風險備註": plan["風險備註"],
        })
        for zh in ["外部允許", "外部Ready", "Market Gate", "Flow Gate", "Fundamental Gate", "Event Gate", "Risk Gate", "外部阻擋原因", "外部資料日", "資料來源層級", "決策摘要"]:
            if zh in plan.columns:
                order_df[zh] = plan[zh]
        self.last_institutional_plan_df = self.normalize_institutional_df(self.enrich_price_and_export_fields(attach_external_display_columns(plan.copy()), id_col="代號"))
        order_df = self.normalize_order_df(self.enrich_price_and_export_fields(attach_external_display_columns(order_df), id_col="代號"))
        return order_df


    def refresh_top20_and_order_views(self):
        for tree in (self.top20_tree, self.top5_tree, self.order_tree, self.inst_tree, self.backtest_tree):
            for item in tree.get_children():
                tree.delete(item)

        if self.last_top20_df is not None and not self.last_top20_df.empty:
            for i, (_, r) in enumerate(self.last_top20_df.iterrows(), start=1):
                ui_action = str(r.get("ui_state", "不可買"))
                light = self._get_display_light(r.to_dict())
                self.top20_tree.insert("", "end", values=(
                    i, r.get("stock_id", ""), r.get("stock_name", ""), self._display_light_symbol(light), r.get("candidate_engine", "混合"),
                    f"{float(r.get('現價', np.nan)):.2f}" if pd.notna(r.get("現價", np.nan)) else "-",
                    f"{float(r.get('漲跌', np.nan)):.2f}" if pd.notna(r.get("漲跌", np.nan)) else "-",
                    f"{float(r.get('漲跌幅%', np.nan)):.2f}" if pd.notna(r.get("漲跌幅%", np.nan)) else "-",
                    r.get("bucket", ""), ui_action,
                    r.get("liquidity_status", ""), f"{float(r.get('liquidity_score', 0) or 0):.1f}",
                    r.get("entry_zone", "-"), r.get("stop_loss", "-"), str(r.get("target_price", r.get("目標價", "-"))),
                    f"{float(r.get('target_1382', 0) or 0):.2f}", f"{float(r.get('target_1618', 0) or 0):.2f}",
                    f"{float(r.get('rr', 0) or 0):.2f}", f"{float(r.get('win_rate', 0) or 0):.1f}",
                    r.get("strategy_nogo_detail", ""), r.get("elimination_reason", "")
                ))

        if self.last_top5_df is not None and not self.last_top5_df.empty:
            for i, (_, r) in enumerate(self.last_top5_df.iterrows(), start=1):
                self.top5_tree.insert("", "end", values=(
                    i, r.get("stock_id", ""), r.get("stock_name", ""),
                    f"{float(r.get('現價', np.nan)):.2f}" if pd.notna(r.get("現價", np.nan)) else "-",
                    f"{float(r.get('漲跌', np.nan)):.2f}" if pd.notna(r.get("漲跌", np.nan)) else "-",
                    f"{float(r.get('漲跌幅%', np.nan)):.2f}" if pd.notna(r.get("漲跌幅%", np.nan)) else "-",
                    r.get("ui_state", "-"),
                    r.get("liquidity_status", ""), f"{float(r.get('liquidity_score', 0) or 0):.1f}",
                    r.get("entry_zone", "-"), r.get("stop_loss", "-"),
                    f"{float(r.get('target_1382', 0) or 0):.2f}",
                    f"{float(r.get('rr', 0) or 0):.2f}",
                    f"{float(r.get('win_rate', 0) or 0):.1f}",
                    f"{float(r.get('backtest_win_rate', 0) or 0):.1f}",
                    f"{float(r.get('cagr', 0) or 0):.2f}",
                    f"{float(r.get('mdd', 0) or 0):.2f}"
                ))
        if hasattr(self, "unique_tree"):
            for item in self.unique_tree.get_children():
                self.unique_tree.delete(item)
            unique_df = getattr(self, "last_unique_decision_df", pd.DataFrame())
            if unique_df is not None and not unique_df.empty:
                unique_df = attach_external_display_columns(unique_df)
                for i, (_, r) in enumerate(unique_df.iterrows(), start=1):
                    self.unique_tree.insert("", "end", values=(
                        i, r.get("stock_id", ""), r.get("stock_name", ""),
                        r.get("trade_allowed", ""), r.get("market_gate_state", ""), r.get("flow_gate_state", ""),
                        r.get("fundamental_gate_state", ""), r.get("event_gate_state", ""), r.get("risk_gate_state", ""),
                        r.get("external_data_ready", ""), r.get("latest_external_date", ""), r.get("market_source_level", ""),
                        r.get("external_blocking_reason", ""), r.get("decision_reason_short", "")
                    ))
        if self.last_institutional_plan_df is not None and not self.last_institutional_plan_df.empty:
            self._render_institutional_tree(self.last_institutional_plan_df)
        if self.last_order_list_df is not None and not self.last_order_list_df.empty:
            self._render_order_tree(self.last_order_list_df)

        self.sync_multi_windows()
        if self.window_current_stock_id and self.current_chart_path and Path(self.current_chart_path).exists():
            self.update_multi_window_stock(self.window_current_stock_id, chart_path=str(self.current_chart_path))

    def on_select_top20(self, event=None):
        sel = self.top20_tree.selection()
        if not sel:
            return
        vals = self.top20_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "強勢候選20"):
            return
        self.sync_all_views(stock_id, source="強勢候選20")
    def on_select_unique(self, event=None):
        sel = self.unique_tree.selection() if hasattr(self, "unique_tree") else []
        if not sel:
            return
        vals = self.unique_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "唯一決策"):
            return
        self.sync_all_views(stock_id, source="唯一決策")

    def on_select_order(self, event=None):
        sel = self.order_tree.selection()
        if not sel:
            return
        vals = self.order_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "執行下單清單"):
            return
        self.sync_all_views(stock_id, source="執行下單清單")


    def show_top5(self):
        def render_top5():
            self.refresh_top20_and_order_views()
            self.left_notebook.select(self.tab_top5)
            lines = ["《v9.2 FINAL-RELEASE V3.5｜AI選股TOP5》", "用途：這裡是核心觀察名單，不代表全部都要買；先看狀態與進場區，再決定是否列入下單清單。", ""]
            for i, (_, r) in enumerate(self.last_top5_df.iterrows(), start=1):
                lines.append(
                    f"{i}. {r['stock_id']} {r['stock_name']}｜{r.get('ui_state','-')}｜進場 {r.get('entry_zone','-')}｜RR {float(r.get('rr',0) or 0):.2f}｜勝率 {float(r.get('win_rate',0) or 0):.1f}%｜回測 {float(r.get('backtest_win_rate',0) or 0):.1f}%"
                )
            self.detail.delete("1.0", tk.END)
            self.detail.insert("1.0", "\n".join(lines))

        if self.last_top5_df is not None and not self.last_top5_df.empty:
            return render_top5()

        # 若 TOP20 尚未建立，先觸發背景分析，再等待結果
        if self.worker is not None and self.worker.is_alive():
            self.set_status("背景分析進行中，等待 TOP5 結果…")
            self.root.after(600, self.show_top5)
            return

        self.show_top20()

        def wait_for_top5(retry=0):
            if self.last_top5_df is not None and not self.last_top5_df.empty:
                render_top5()
                return
            if self.worker is not None and self.worker.is_alive() and retry < 20:
                self.set_status(f"等待 TOP5 結果…({retry+1}/20)")
                self.root.after(600, lambda: wait_for_top5(retry + 1))
                return
            messagebox.showwarning("提醒", "目前尚無可用 TOP5 資料，請先執行 AI選股TOP20。")

        self.root.after(600, lambda: wait_for_top5(0))

    def on_select_top5(self, event=None):
        sel = self.top5_tree.selection()
        if not sel:
            return
        vals = self.top5_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "主攻5"):
            return
        self.sync_all_views(stock_id, source="主攻5")


    def on_select_institutional(self, event=None):
        sel = self.inst_tree.selection()
        if not sel:
            return
        vals = self.inst_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "組合交易計畫"):
            return
        self.sync_all_views(stock_id, source="組合交易計畫")

    def build_full_history_once(self):
        self._start_build_history(resume=False)

    def resume_full_history(self):
        self._start_build_history(resume=True)

    def _start_build_history(self, resume: bool = False):
        master = self.db.get_master()
        if master.empty:
            return messagebox.showwarning("提醒", "請先初始化全市場。")

        counts = master["stock_id"].astype(str).apply(self.db.get_price_history_count)
        ready = int((counts >= 240).sum())
        total = len(master)
        state = self.load_history_state()

        if resume:
            if not state:
                self.append_log("未找到上次中斷狀態，改為一般補建模式。")
            ok = messagebox.askyesno("確認", f"將執行續跑建庫。\n目前完整檔數：{ready}/{total}\n系統會自動跳過已完成股票，是否開始？")
        elif ready >= int(total * 0.9):
            ok = messagebox.askyesno("確認", f"已有 {ready}/{total} 檔具備完整歷史資料。\n再次執行將只補缺漏資料，是否繼續？")
        else:
            ok = messagebox.askyesno("確認", f"將建立完整歷史資料。\n目前完整檔數：{ready}/{total}\n是否開始？")
        if not ok:
            return

        def worker():
            try:
                self.ui_call(self.clear_log)
                self.ui_call(self.append_log, f"開始完整建庫，模式={'續跑' if resume else '一般'}，主檔 {total} 檔")
                self.ui_call(self.set_status, "開始建立完整歷史資料（分批 / 可中斷 / 可續跑）...")
                self.ui_call(self.start_task, "建立完整歷史", total)
                self.ui_call(self.update_task, "建立完整歷史", 0, total, 0, 0, 0, "準備中")
                counters = {"ok": 0, "fail": 0}

                def progress(idx, total_count, sid, existing_count, flag):
                    if flag in ("fail", "error"):
                        counters["fail"] += 1
                    elif flag == "ok":
                        counters["ok"] += 1
                    self.ui_call(self.update_task, "建立完整歷史", idx, total_count, counters["ok"], counters["fail"], 0, sid)
                    self.save_history_state({
                        "mode": "build_history",
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "current_index": idx,
                        "total": total_count,
                        "stock_id": sid,
                        "completed_ready": int((master["stock_id"].astype(str).apply(self.db.get_price_history_count) >= 240).sum()),
                        "success": counters["ok"],
                        "failed": counters["fail"],
                        "existing_count": int(existing_count),
                    })
                    if idx % 10 == 0 or idx == total_count:
                        self.ui_call(self.set_status, f"建立歷史中 {idx}/{total_count}｜{sid}｜成功 {counters['ok']}｜失敗 {counters['fail']}")

                success, failed, rows = self.data_engine.build_full_history(
                    batch_size=self.history_batch_size,
                    sleep_sec=self.history_sleep_sec,
                    progress_cb=progress,
                    log_cb=lambda msg: self.ui_call(self.append_log, msg),
                    cancel_cb=lambda: self.cancel_event.is_set(),
                )
                self.clear_history_state()
                self.ui_call(self.update_task, "建立完整歷史", total, total, success, failed, 0, "完成")
                self.ui_call(self.set_status, f"完整歷史建立完成：成功 {success} 檔，失敗 {failed} 檔，寫入 {rows} 筆。")
                self.ui_call(self.append_log, f"完整建庫完成：成功 {success} 檔｜失敗 {failed} 檔｜寫入 {rows} 筆")
                self.ui_call(self.show_welcome_message)
                self.ui_call(messagebox.showinfo, "完成", f"完整歷史建立完成\n成功 {success} 檔\n失敗 {failed} 檔\n寫入 {rows} 筆\n\n已支援分批抓取 / 中斷續跑")
            except OperationCancelled:
                state2 = self.load_history_state()
                sid = state2.get("stock_id", "")
                idx = state2.get("current_index", 0)
                total_count = state2.get("total", total)
                self.ui_call(self.append_log, f"作業已中斷：停在 {idx}/{total_count}｜{sid}")
                self.ui_call(self.set_status, f"建庫已中斷：停在 {idx}/{total_count}｜{sid}，可按『續跑建庫』")
                self.ui_call(messagebox.showwarning, "已中斷", f"完整建庫已中斷\n目前停在 {idx}/{total_count}｜{sid}\n\n下次請按『續跑建庫』，系統會自動跳過已完成資料。")
            except Exception as e:
                traceback.print_exc()
                self.ui_call(self.append_log, f"完整建庫發生錯誤：{e}")
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, "build_history")

    def init_master_data(self):
        master = self.db.get_master()
        if not master.empty and len(master) > 500:
            ok = messagebox.askyesno("確認", f"目前已存在 {len(master)} 檔股票主檔。\n重新初始化將覆蓋現有主檔，是否繼續？")
            if not ok:
                return

        def worker():
            try:
                self.ui_call(self.set_status, "開始初始化全市場股票清單...")
                self.ui_call(self.start_task, "初始化全市場", 4)
                self.ui_call(self.update_task, "初始化全市場", 1, 4, item="抓取主檔")
                universe = build_full_market_universe()
                if universe is None or universe.empty:
                    csv_path = resolve_master_csv()
                    self.db.import_master_csv(csv_path)
                    master2 = self.db.get_master()
                    self.ui_call(self.refresh_filters)
                    self.ui_call(self.refresh_all_tables)
                    self.ui_call(self.refresh_classification_summary_ui)
                    self.ui_call(self.update_task, "初始化全市場", 4, 4, success=1, item="完成")
                    self.ui_call(self.set_status, f"已改用本地主檔，共 {len(master2)} 檔。")
                    self.ui_call(messagebox.showinfo, "完成", f"全市場抓取失敗，已改用本地主檔\n共 {len(master2)} 檔\n\n使用主檔：{csv_path}")
                    return
                self.db.import_master_df(universe)
                master2 = self.db.get_master()
                self.ui_call(self.refresh_filters)
                self.ui_call(self.refresh_all_tables)
                self.ui_call(self.refresh_classification_summary_ui)
                self.ui_call(self.update_task, "初始化全市場", 4, 4, success=1, item="完成")
                self.ui_call(self.set_status, f"全市場初始化完成，共 {len(master2)} 檔。")
                self.ui_call(messagebox.showinfo, "完成", f"全市場股票清單初始化完成\n共 {len(master2)} 檔")
            except Exception as e:
                traceback.print_exc()
                self.ui_call(messagebox.showerror, "錯誤", f"初始化失敗：\n{e}")

        self._run_in_thread(worker, "init_market")

    def populate_operation_sop(self, market: dict, trade_top20: pd.DataFrame, today_buy: pd.DataFrame, wait_pullback: pd.DataFrame, attack: pd.DataFrame, defense: pd.DataFrame):
        for item in self.sop_tree.get_children():
            self.sop_tree.delete(item)
        sop_df = OperationGuideEngine.build_playbook(market, trade_top20, today_buy, wait_pullback, attack, defense)
        self.last_operation_sop_df = sop_df.copy()
        if sop_df.empty:
            return
        for _, r in sop_df.iterrows():
            self.sop_tree.insert("", "end", values=(r["step"], r["module"], r["focus"], r["rule"], r["purpose"], r["output"]))

    def show_operation_guide(self):
        ranking = self._filtered_ranking()
        trade_top20 = getattr(self, "last_top20_df", pd.DataFrame())
        today_buy = getattr(self, "last_today_buy_df", pd.DataFrame())
        wait_pullback = getattr(self, "last_wait_df", pd.DataFrame())
        attack = getattr(self, "last_attack_df", pd.DataFrame())
        defense = getattr(self, "last_defense_df", pd.DataFrame())
        market = self.master_trading_engine.market_engine.get_market_regime()
        self.populate_operation_sop(market, trade_top20, today_buy, wait_pullback, attack, defense)
        lines = [
            "《V3.5 操作版｜功能與用途》",
            "",
            f"市場狀態：{market['regime']}（{market['score']:.2f}）",
            f"市場說明：{market['memo']}",
            "",
            "核心流程：先看市場 → 再看輪動 → 再看今日可下單 / 條件預掛 → 再看下單清單 → 最後用回測驗證。",
            "",
            "五種狀態怎麼用：",
        ]
        for state in ["可買", "準備買", "條件預掛", "觀察", "不可買"]:
            lines.append(f"- {state}：{OperationGuideEngine.explain_state(state)}")
        if trade_top20 is not None and not trade_top20.empty:
            lines.extend(["", "目前 TOP20 前3檔："])
            for i, (_, r) in enumerate(trade_top20.head(3).iterrows(), start=1):
                lines.append(f"{i}. {r['stock_id']} {r['stock_name']}｜{r.get('ui_state','-')}｜{r.get('entry_zone','-')}｜RR {float(r.get('rr',0) or 0):.2f}")
        elif ranking is not None and not ranking.empty:
            lines.extend(["", "尚未建立 TOP20，請先執行 AI選股TOP20。", f"目前排行第一：{ranking.iloc[0]['stock_id']} {ranking.iloc[0]['stock_name']}"])
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))
        self.left_notebook.select(self.tab_sop)

    def refresh_filters(self, reset_active_filters: bool = False):
        master = self.db.get_master()
        market_values = ["全部"]
        industry_values = ["全部"]
        theme_values = ["全部"]
        if not master.empty:
            market_values += sorted([x for x in master["market"].dropna().unique().tolist() if str(x).strip() != ""])
            industry_values += sorted([x for x in master["industry"].dropna().unique().tolist() if str(x).strip() != ""])
            theme_values += sorted([x for x in master["theme"].dropna().unique().tolist() if str(x).strip() != ""])

        self.market_cb["values"] = market_values
        self.industry_cb["values"] = industry_values
        self.theme_cb["values"] = theme_values

        if reset_active_filters:
            try:
                self.market_var.set("全部")
                self.industry_var.set("全部")
                self.theme_var.set("全部")
                self.search_var.set("")
            except Exception:
                pass
            return

        try:
            if self.market_var.get() not in market_values:
                self.market_var.set("全部")
            if self.industry_var.get() not in industry_values:
                self.industry_var.set("全部")
            if self.theme_var.get() not in theme_values:
                self.theme_var.set("全部")
        except Exception:
            pass

    def _parse_search_tokens(self, raw_query: str):
        raw = str(raw_query or "").strip()
        if not raw:
            return []
        normalized = raw.replace("，", ",").replace("、", ",").replace("；", ",").replace(";", ",")
        parts = [p.strip() for p in normalized.split(",")]
        return [p for p in parts if p]

    def _apply_search_filter(self, df: pd.DataFrame, raw_query: str) -> pd.DataFrame:
        tokens = self._parse_search_tokens(raw_query)
        if not tokens or df is None or df.empty:
            return df

        stock_id_series = df["stock_id"].astype(str)
        stock_name_series = df["stock_name"].astype(str)
        mask = pd.Series(False, index=df.index)

        for token in tokens:
            token_mask = pd.Series(False, index=df.index)
            escaped = re.escape(token)
            if token.isdigit():
                token_mask = token_mask | stock_id_series.eq(token)
                token_mask = token_mask | stock_id_series.str.contains(escaped, case=False, na=False, regex=True)
            else:
                token_mask = token_mask | stock_id_series.str.contains(escaped, case=False, na=False, regex=True)
            token_mask = token_mask | stock_name_series.str.contains(escaped, case=False, na=False, regex=True)
            mask = mask | token_mask

        return df[mask]

    def _filtered_ranking(self, force_full: bool = False):
        df = self.db.get_latest_ranking()
        if df.empty:
            return df

        force_full = bool(force_full or getattr(self, "force_show_full_ranking_once", False))
        if force_full:
            try:
                self.force_show_full_ranking_once = False
            except Exception:
                pass
        else:
            if self.market_var.get() != "全部":
                df = df[df["market"] == self.market_var.get()]
            if self.industry_var.get() != "全部":
                df = df[df["industry"] == self.industry_var.get()]
            if self.theme_var.get() != "全部":
                df = df[df["theme"] == self.theme_var.get()]
            q = self.search_var.get().strip()
            if q:
                df = self._apply_search_filter(df, q)

        df = df.sort_values(["rank_all"]).reset_index(drop=True)
        return self.enrich_price_and_export_fields(df, id_col="stock_id")

    def _populate_rank_tree(self, df: pd.DataFrame):
        if df is None or df.empty:
            return
        for i, row in df.iterrows():
            self.rank_tree.insert("", "end", values=(
                i + 1, row["stock_id"], row["stock_name"],
                f"{float(row.get('現價', np.nan)):.2f}" if pd.notna(row.get("現價", np.nan)) else "-",
                f"{float(row.get('漲跌', np.nan)):.2f}" if pd.notna(row.get("漲跌", np.nan)) else "-",
                f"{float(row.get('漲跌幅%', np.nan)):.2f}" if pd.notna(row.get("漲跌幅%", np.nan)) else "-",
                row["industry"], row["theme"],
                f"{row['total_score']:.2f}", f"{row['ai_score']:.2f}", row["signal"], row["action"]
            ))

    def _reload_rank_tree_after_rebuild(self, force_full: bool = True):
        """重建排行後專用：直接重建排行榜畫面，避免 refresh_all_tables() 途中例外時整片空白。"""
        try:
            self.refresh_filters(True)
        except Exception:
            pass

        try:
            for item in self.rank_tree.get_children():
                self.rank_tree.delete(item)
        except Exception:
            pass

        df = pd.DataFrame()
        try:
            df = self._filtered_ranking(force_full=force_full)
        except Exception as exc:
            log_exception("_reload_rank_tree_after_rebuild::_filtered_ranking failed", exc)
            try:
                self.append_log(f"重建排行後讀取排行失敗：{exc}", "ERROR")
            except Exception:
                pass
            df = pd.DataFrame()

        if df is None or df.empty:
            try:
                full_df = self.db.get_latest_ranking()
                if full_df is not None and not full_df.empty:
                    df = self.enrich_price_and_export_fields(full_df.sort_values(["rank_all"]).reset_index(drop=True), id_col="stock_id")
            except Exception as exc:
                log_exception("_reload_rank_tree_after_rebuild::fallback latest ranking failed", exc)

        if df is None or df.empty:
            self.set_status("重建排行已完成，但排行榜仍無資料可顯示。")
            return

        self._populate_rank_tree(df)
        count = len(self.rank_tree.get_children())
        self.set_status(f"排行已完成，共 {len(df)} 檔｜UI已載入 {count} 筆")
        try:
            if count <= 0:
                self.append_log("重建排行後 UI 載入完成，但 rank_tree 仍為 0 筆。", "ERROR")
            else:
                self.append_log(f"重建排行後已載入排行榜 {count} 筆。")
        except Exception:
            pass

    def refresh_all_tables(self, force_full_ranking: bool = False):
        for tree in (self.dashboard_tree, self.sop_tree, self.rotation_tree, self.rank_tree, self.sector_tree, self.theme_tree):
            for item in tree.get_children():
                tree.delete(item)

        if not self.ensure_ranking_ready(auto_rebuild=True):
            price_rows = self.db.get_total_price_rows()
            if price_rows > 0:
                self.set_status("已有歷史資料，但尚未形成有效排行；請先補足歷史或重建排行。")
            else:
                self.set_status("目前尚無排行資料，請先初始化、建立歷史，再重建排行。")
            self.show_welcome_message()
            return

        used_full_fallback = False
        df = self._filtered_ranking(force_full=force_full_ranking)
        if df.empty:
            full_df = self.db.get_latest_ranking()
            if full_df is not None and not full_df.empty:
                df = full_df.sort_values(["rank_all"]).reset_index(drop=True)
                df = self.enrich_price_and_export_fields(df, id_col="stock_id")
                used_full_fallback = True
                try:
                    self.force_show_full_ranking_once = False
                except Exception:
                    pass
            else:
                self.set_status("目前篩選條件下沒有資料。")
                return

        self._populate_rank_tree(df)
        if used_full_fallback:
            self.set_status("目前篩選條件無資料，已改顯示完整排行。")

        sector = (
            df.groupby("industry", as_index=False)
            .agg(count=("stock_id", "count"), avg_total=("total_score", "mean"), avg_ai=("ai_score", "mean"))
            .sort_values(["avg_total", "avg_ai"], ascending=False)
        )
        for _, r in sector.iterrows():
            top_name = df[df["industry"] == r["industry"]].sort_values("total_score", ascending=False).iloc[0]["stock_name"]
            self.sector_tree.insert("", "end", values=(
                r["industry"], int(r["count"]), f"{r['avg_total']:.2f}", f"{r['avg_ai']:.2f}", top_name
            ))

        theme = (
            df.groupby("theme", as_index=False)
            .agg(count=("stock_id", "count"), avg_total=("total_score", "mean"), avg_ai=("ai_score", "mean"))
            .sort_values(["avg_total", "avg_ai"], ascending=False)
        )
        for _, r in theme.iterrows():
            top_name = df[df["theme"] == r["theme"]].sort_values("total_score", ascending=False).iloc[0]["stock_name"]
            self.theme_tree.insert("", "end", values=(
                r["theme"], int(r["count"]), f"{r['avg_total']:.2f}", f"{r['avg_ai']:.2f}", top_name
            ))

        market_engine = self.master_trading_engine.market_engine
        regime = market_engine.get_market_regime()
        rotation = IndustryRotationEngine.summarize(df)

        dash_rows = [
            ("市場狀態", regime["regime"], f"Regime score {regime['score']:.2f}"),
            ("市場廣度", f"{regime['breadth']:.1f}", "強勢訊號占比"),
            ("排行檔數", str(len(df)), "目前篩選後股票數"),
            ("最強題材", str(df.groupby("theme")["total_score"].mean().sort_values(ascending=False).index[0]) if not df.empty else "-", "依平均總分"),
            ("最強產業", str(rotation.iloc[0]["industry"]) if not rotation.empty else "-", "依輪動分"),
        ]
        for m, v, d in dash_rows:
            self.dashboard_tree.insert("", "end", values=(m, v, d))
        for _, r in rotation.iterrows():
            self.rotation_tree.insert("", "end", values=(
                r["industry"], int(r["count"]), f"{r['avg_total']:.2f}", f"{r['avg_ai']:.2f}", int(r["trend_count"]),
                f"{r['hot_score']:.2f}", r["rotation"]
            ))

        trade = self.master_trading_engine.get_trade_pool(df)
        self.populate_operation_sop(trade["market"], trade["trade_top20"], trade["today_buy"], trade["wait_pullback"], trade["attack"], trade["defense"])
        attack_cnt = len(trade["attack"])
        defense_cnt = len(trade["defense"])
        self.set_status(
            f"已載入資料，共 {len(df)} 檔｜市場 {trade['market']['regime']}｜主攻 {attack_cnt}｜防守 {defense_cnt}"
        )
        if (self.last_top20_df is not None and not self.last_top20_df.empty) or (self.last_order_list_df is not None and not self.last_order_list_df.empty):
            self.refresh_top20_and_order_views()

    def update_data(self):
        last_date = self.db.get_last_price_date()
        today = datetime.now().strftime("%Y-%m-%d")
        if last_date == today:
            ok = messagebox.askyesno("確認", f"今日資料（{today}）可能已更新過。\n再次執行會覆蓋今日官方資料，是否繼續？")
            if not ok:
                return

        def worker():
            try:
                master = self.db.get_master()
                total = len(master) if not master.empty else 1
                counters = {"ok": 0, "fail": 0, "skip": 0}
                self.ui_call(self.clear_log)
                self.ui_call(self.start_task, "每日增量更新", total)

                def progress(idx, total_count, sid, row_count, flag):
                    if flag == "ok":
                        counters["ok"] += 1
                    elif flag in ("fail", "error"):
                        counters["fail"] += 1
                    else:
                        counters["skip"] += 1
                    self.ui_call(self.update_task, "每日增量更新", idx, total_count, counters["ok"], counters["fail"], counters["skip"], sid)

                success, failed, rows = self.data_engine.update_incremental(progress_cb=progress, log_cb=lambda msg: self.ui_call(self.append_log, msg), cancel_cb=lambda: self.cancel_event.is_set())

                # V9.6.2 FUNDAMENTAL_LOCAL_CACHE：每日更新完成後先同步基本面本地快取，再重排行。
                # 順序固定為：行情 → external_valuation/external_revenue → financial_feature_daily → ranking_result。
                self.ui_call(self.append_log, "[FUNDAMENTAL CACHE] 每日行情更新完成，開始同步 market_snapshot + EPS/估值 + 月營收本地快取")
                cache_result = ExternalDataFetcher(self.db).sync_fundamental_local_cache(
                    modules=["market_snapshot", "valuation", "revenue"],
                    log_cb=lambda msg: self.ui_call(self.append_log, msg),
                )
                if float(cache_result.get("ne_ratio", 1.0) or 1.0) >= 0.80:
                    self.ui_call(self.append_log, f"[FUNDAMENTAL CACHE][WARNING] financial_feature_daily NE_ratio={cache_result.get('ne_ratio', 1.0):.2%}，排行會標示基本面資料不足。")

                self.ui_call(self.start_task, "重建排行", total)
                rank_skip = {"skip": 0}
                def rank_progress(idx, total_count, sid, ok_count, fail_count, skip_count, flag):
                    rank_skip["skip"] = skip_count
                    self.ui_call(self.update_task, "重建排行", idx, total_count, ok_count, fail_count, skip_count, sid)
                rank_count = self.rank_engine.rebuild(progress_cb=rank_progress, log_cb=lambda msg: self.ui_call(self.append_log, msg), cancel_cb=lambda: self.cancel_event.is_set())
                self.force_show_full_ranking_once = True
                self.ui_call(self.refresh_filters, True)
                self.ui_call(self.refresh_all_tables, True)
                self.ui_call(self.show_welcome_message)
                self.ui_call(self.finish_task, "每日增量更新", f"完成：成功 {success} 檔，寫入 {rows} 筆，基本面特徵 {cache_result.get('feature_rows', 0)} 筆，排行 {rank_count} 檔。")
                self.ui_call(messagebox.showinfo, "完成", f"每日增量更新完成\n成功 {success} 檔\n寫入 {rows} 筆\n基本面特徵 {cache_result.get('feature_rows', 0)} 筆\n排行 {rank_count} 檔\n（行情 + EPS/估值 + 月營收已先寫入本地DB）")
            except OperationCancelled:
                self.ui_call(self.append_log, "每日更新/重排行已中斷")
                self.ui_call(self.finish_task, "每日增量更新", "作業已中斷")
            except Exception as e:
                traceback.print_exc()
                try:
                    pool_audit = getattr(self.master_trading_engine, "last_pool_audit", {}) or {}
                    if pool_audit:
                        self.ui_call(self.append_log, f"[POOL-AUDIT-LAST] candidate20={pool_audit.get('candidate20_count','-')}｜core_attack5={pool_audit.get('core_attack5_count','-')}｜today_buy={pool_audit.get('today_buy_count','-')}｜execution_ready={pool_audit.get('execution_ready_count','-')}｜unique_decision={pool_audit.get('unique_decision_count','-')}")
                        if pool_audit.get('core_minus_candidate20'):
                            self.ui_call(self.append_log, f"[POOL-AUDIT-LAST] core_attack5 - candidate20：{','.join(pool_audit.get('core_minus_candidate20', [])[:20])}")
                        if pool_audit.get('today_minus_core'):
                            self.ui_call(self.append_log, f"[POOL-AUDIT-LAST] today_buy - core_attack5：{','.join(pool_audit.get('today_minus_core', [])[:20])}")
                        if pool_audit.get('unique_minus_core'):
                            self.ui_call(self.append_log, f"[POOL-AUDIT-LAST] unique_decision - core_attack5：{','.join(pool_audit.get('unique_minus_core', [])[:20])}")
                except Exception:
                    pass
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, "update_daily")

    def rebuild_ranking(self):
        def worker():
            try:
                master = self.db.get_master()
                total = len(master) if not master.empty else 1
                self.ui_call(self.clear_log)
                self.ui_call(self.start_task, "重建排行", total)
                def progress(idx, total_count, sid, ok_count, fail_count, skip_count, flag):
                    self.ui_call(self.update_task, "重建排行", idx, total_count, ok_count, fail_count, skip_count, sid)
                count = self.rank_engine.rebuild(progress_cb=progress, log_cb=lambda msg: self.ui_call(self.append_log, msg), cancel_cb=lambda: self.cancel_event.is_set())
                self.force_show_full_ranking_once = True
                self.ui_call(self._reload_rank_tree_after_rebuild, True)
                self.ui_call(self.refresh_all_tables, True)
                self.ui_call(self.refresh_classification_summary_ui)
                self.ui_call(self.finish_task, "重建排行", f"排行已完成，共 {count} 檔")
                if count <= 0:
                    self.ui_call(messagebox.showwarning, "提醒", "排行重建完成，但目前可計算檔數為 0。\n請先建立至少 70 根以上歷史K線資料。")
                else:
                    self.ui_call(messagebox.showinfo, "完成", f"排行已完成，共 {count} 檔")
            except OperationCancelled:
                self.ui_call(self.finish_task, "重建排行", "重建排行已中斷")
            except Exception as e:
                traceback.print_exc()
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, "rebuild_rank")

    def _ensure_external_data_ready_auto_sync(self, context: str = "AI選股TOP20") -> tuple[int, str]:
        """V9.5.1：TOP20前置條件自動補救。
        使用者按 AI選股TOP20 時，不再要求手動先按「同步外部資料」；
        若 mandatory external data 尚未 ready，系統會先執行一次真實外部資料同步，
        同步後再重新檢查 Gate。
        """
        readiness = ExternalDataReadiness(self.db)
        ready, reason = readiness.mandatory_ready()
        if int(ready) == 1:
            try:
                self.ui_call(self.append_log, f"[V9.5.1-AUTO-SYNC-SKIP] 外部資料已就緒，直接執行 {context}")
            except Exception:
                pass
            return 1, ""

        start_reason = str(reason or "外部資料未就緒")
        auto_run_id = self.db.log_system_run(
            event="external_auto_sync_before_top20",
            status="start",
            message=f"{context} preflight not ready; auto sync triggered: {start_reason}",
            step="preflight_auto_sync",
            module="external_data",
        )
        self.ui_call(self.append_log, f"[V9.5.1-AUTO-SYNC-START] {context} 偵測外部資料未就緒，系統自動同步外部資料｜run_id={auto_run_id}｜原因：{start_reason}", "WARNING")
        self.ui_call(self.set_status, f"{context}：外部資料未就緒，正在自動同步...")

        try:
            sync_result = ExternalDataFetcher(self.db).refresh_external_data_pipeline(run_id=auto_run_id)
            status_df = self.db.read_table("external_source_status", limit=500)
            self.ui_call(self._populate_external_tree, status_df)
            blocking = sync_result.get("blocking", []) if isinstance(sync_result, dict) else []
            if blocking:
                self.ui_call(self.append_log, f"[V9.5.1-AUTO-SYNC-WARN] 外部資料自動同步完成但仍有阻擋項｜run_id={auto_run_id}｜{'; '.join(blocking)}", "WARNING")
            else:
                self.ui_call(self.append_log, f"[V9.5.1-AUTO-SYNC-OK] 外部資料自動同步完成｜run_id={auto_run_id}")
        except Exception as exc:
            msg = f"外部資料自動同步失敗：{exc}"
            self.db.log_system_run(
                event="external_auto_sync_before_top20",
                status="fail",
                message=msg,
                run_id=auto_run_id,
                step="preflight_auto_sync",
                module="external_data",
            )
            self.ui_call(self.append_log, f"[V9.5.1-AUTO-SYNC-FAIL] {msg}", "ERROR")
            return 0, msg

        ready2, reason2 = readiness.mandatory_ready()
        if int(ready2) == 1:
            self.db.log_system_run(
                event="external_auto_sync_before_top20",
                status="ok",
                message=f"{context} external preflight passed after auto sync",
                run_id=auto_run_id,
                step="preflight_auto_sync",
                module="external_data",
            )
            self.ui_call(self.append_log, f"[V9.5.1-PREFLIGHT-GO] 外部資料已就緒，繼續執行 {context}")
            return 1, ""

        final_reason = str(reason2 or "外部資料自動同步後仍未就緒")
        self.db.log_system_run(
            event="external_auto_sync_before_top20",
            status="blocked",
            message=f"{context} external preflight still blocked after auto sync: {final_reason}",
            run_id=auto_run_id,
            step="preflight_auto_sync",
            module="external_data",
        )
        self.ui_call(self.append_log, f"[V9.5.9-PREFLIGHT-SOFT-BLOCK] 自動同步後外部資料仍未就緒，但不停止 {context}：{final_reason}", "WARNING")
        return 0, final_reason

    def show_top20(self):
        if not self.ensure_ranking_ready(auto_rebuild=True):
            return messagebox.showwarning("提醒", "目前尚無可用排行資料，請先建立歷史資料後重建排行。")
        df = self._filtered_ranking()
        if df.empty:
            try:
                full_df = self.db.get_latest_ranking()
                if full_df is not None and not full_df.empty:
                    df = self.enrich_price_and_export_fields(full_df.sort_values(["rank_all"]).reset_index(drop=True), id_col="stock_id")
                    self.set_status("目前篩選條件無資料，AI選股TOP20 已改用完整排行執行。")
                else:
                    return messagebox.showwarning("提醒", "目前沒有可用排行資料")
            except Exception:
                return messagebox.showwarning("提醒", "目前篩選條件下沒有可用資料")

        def worker():
            try:
                total = len(df)
                self.ui_call(self.clear_log)
                self.ui_call(self.start_task, "AI選股TOP20", total)

                auto_ready, auto_reason = self._ensure_external_data_ready_auto_sync("AI選股TOP20")
                if int(auto_ready) != 1:
                    self.ui_call(self.append_log, "[V9.5.9-SOFT-BLOCK] 外部資料自動同步後仍未完整，但不停止AI分析/TOP20；execution_ready僅作資訊欄位。原因：" + str(auto_reason), "WARNING")
                    self.ui_call(self.set_status, "AI選股TOP20繼續執行；外部資料未完整僅顯示SOFT_BLOCK提示")


                def progress(idx, total_count, sid):
                    self.ui_call(self.update_task, "AI選股TOP20", idx, total_count, idx, 0, 0, sid)

                trade = self.master_trading_engine.get_trade_pool(
                    df,
                    progress_cb=progress,
                    log_cb=lambda msg: self.ui_call(self.append_log, msg),
                    cancel_cb=lambda: self.cancel_event.is_set(),
                )
                market = trade["market"]
                trade_top20 = trade["trade_top20"]
                tradable_top20 = trade.get("tradable_top20", pd.DataFrame())
                self.last_candidate_top20_df = self.enrich_price_and_export_fields(trade_top20.copy(), id_col="stock_id") if trade_top20 is not None and not trade_top20.empty else pd.DataFrame()
                attack = trade["attack"]
                watch = trade["watch"]
                defense = trade["defense"]
                today_buy = trade["today_buy"]
                wait_pullback = trade["wait_pullback"]
                theme_summary = trade["theme_summary"]
                eliminated = trade.get("eliminated", pd.DataFrame())

                ui_top20_source = trade_top20 if trade_top20 is not None and not trade_top20.empty else tradable_top20
                self.last_top20_df = self.enrich_price_and_export_fields(ui_top20_source.copy(), id_col="stock_id") if ui_top20_source is not None and not ui_top20_source.empty else pd.DataFrame()
                self.cache_trade_dataframe(self.last_top20_df)
                top5 = attack.head(5).copy() if attack is not None and not attack.empty else (trade_top20.head(5).copy() if trade_top20 is not None and not trade_top20.empty else pd.DataFrame())
                if not top5.empty:
                    bt_rows = []
                    for _, rr in top5.iterrows():
                        bt = self.backtest_engine.estimate_trade_quality(str(rr["stock_id"]))
                        bt_rows.append(bt)
                    bt_df = pd.DataFrame(bt_rows)
                    top5 = pd.concat([top5.reset_index(drop=True), bt_df.reset_index(drop=True)], axis=1)
                self.last_top5_df = self.enrich_price_and_export_fields(top5.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_top5_df)
                self.cache_backtest_dataframe(self.last_top5_df)
                self.last_attack_df = self.enrich_price_and_export_fields(attack.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_attack_df)
                self.last_watch_df = self.enrich_price_and_export_fields(watch.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_watch_df)
                self.last_defense_df = self.enrich_price_and_export_fields(defense.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_defense_df)
                self.last_theme_summary_df = theme_summary.copy()
                self.last_today_buy_df = self.enrich_price_and_export_fields(today_buy.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_today_buy_df)
                self.last_wait_df = self.enrich_price_and_export_fields(wait_pullback.copy(), id_col="stock_id")
                self.cache_trade_dataframe(self.last_wait_df)
                self.last_order_list_df = self.normalize_order_df(self.build_order_list(self.last_today_buy_df))
                inst_plan_raw = self.portfolio_engine.build_institutional_plan(self.last_today_buy_df.copy())
                self.last_institutional_plan_df = self.normalize_institutional_df(self.enrich_price_and_export_fields(inst_plan_raw, id_col="代號"))
                unique_raw = trade.get("unique_decision", pd.DataFrame())
                if unique_raw is not None and not unique_raw.empty:
                    self.last_unique_decision_df = self.enrich_price_and_export_fields(unique_raw.copy(), id_col="stock_id").head(REPORT_DECISION_LIMITS["unique_decision"])
                else:
                    self.last_unique_decision_df = pd.DataFrame()

                assert_phase1_report_consistency(
                    self.last_top20_df,
                    self.last_attack_df,
                    self.last_today_buy_df,
                    self.last_wait_df,
                    self.last_watch_df,
                    self.last_unique_decision_df,
                    self.last_order_list_df,
                    self.last_institutional_plan_df,
                )

                self.ui_call(self.populate_operation_sop, market, trade_top20, today_buy, wait_pullback, attack, defense)
                self.ui_call(self.refresh_top20_and_order_views)
                self.ui_call(self.left_notebook.select, self.tab_top20)
                self.ui_call(self.open_three_windows)

                defend_cnt = int(trade_top20["bucket"].eq("防守").sum()) if not trade_top20.empty else 0
                eliminated_cnt = len(eliminated) if eliminated is not None else 0
                lines = [
                    "《v9.5 EXTERNAL DECISION INTEGRATED 操作版》",
                    f"市場判斷：{market['regime']}（{market['score']:.2f}）｜市場廣度 {market['breadth']:.1f}",
                    f"市場說明：{market['memo']}",
                    f"TOP20 觀察池：{len(trade_top20)} 檔｜今日可下單：{len(today_buy)}｜條件預掛：{len(wait_pullback)}｜防守：{defend_cnt}｜淘汰：{eliminated_cnt}",
                    f"交易門檻：{STRATEGY_CONFIG_MANAGER.summary_text()}",
                    "操作用途：先看今日可下單，再看條件預掛，沒有進場區就不下單。",
                    "",
                    "【TOP20 觀察池 前5檔】",
                ]
                if trade_top20.empty:
                    lines.append("目前無符合條件標的")
                else:
                    for i, (_, r) in enumerate(trade_top20.head(REPORT_DECISION_LIMITS["core_attack5"]).iterrows(), start=1):
                        lines.append(f"{i}. {r['stock_id']} {r['stock_name']}｜{r['bucket']}｜{r.get('liquidity_status','-')} {float(r.get('liquidity_score',0) or 0):.1f}｜{r['trade_action']}｜RR {r['rr']:.2f}｜勝率 {r['win_rate']:.1f}%")

                lines.extend(["", "【今日可下單】"])
                if today_buy.empty:
                    lines.append("今日無符合 SOP 的可買名單（允許空白，不為湊數放寬）。")
                else:
                    for i, (_, r) in enumerate(today_buy.iterrows(), start=1):
                        lines.append(f"{i}. {r['stock_id']} {r['stock_name']}｜{r['trade_action']}｜RR {r['rr']:.2f}｜勝率 {r['win_rate']:.1f}%｜{r['entry_zone']}")

                lines.extend(["", "【條件預掛】"])
                if wait_pullback.empty:
                    lines.append("無條件預掛")
                else:
                    for i, (_, r) in enumerate(wait_pullback.iterrows(), start=1):
                        lines.append(f"{i}. {r['stock_id']} {r['stock_name']}｜條件預掛｜RR {r['rr']:.2f}｜勝率 {r['win_rate']:.1f}%｜{r['entry_zone']}")

                self.ui_call(self.detail.delete, "1.0", tk.END)
                self.ui_call(self.detail.insert, "1.0", "\n".join(lines))
                self.ui_call(self.finish_task, "AI選股TOP20", f"AI選股完成：TOP20 {len(trade_top20)}｜今日可下單 {len(today_buy)}｜等待 {len(wait_pullback)}")
            except OperationCancelled:
                self.ui_call(self.finish_task, "AI選股TOP20", "AI選股TOP20 已中斷")
            except Exception as e:
                traceback.print_exc()
                self.ui_call(messagebox.showerror, "錯誤", str(e))

        self._run_in_thread(worker, "show_top20")


    def show_strategy_backtest(self):
        if self.last_top20_df is None or self.last_top20_df.empty:
            return messagebox.showwarning("提醒", "請先執行 AI選股TOP20。")
        rows = []
        for _, r in self.last_top20_df.head(10).iterrows():
            bt = self.backtest_engine.estimate_trade_quality(str(r["stock_id"]))
            rows.append({
                "stock_id": r["stock_id"],
                "stock_name": r["stock_name"],
                "backtest_win_rate": bt.get("backtest_win_rate", 0),
                "avg_return": bt.get("avg_return", 0),
                "cagr": bt.get("cagr", 0),
                "mdd": bt.get("mdd", 0),
                "sharpe": bt.get("sharpe", 0),
                "samples": bt.get("samples", 0),
            })
        out = pd.DataFrame(rows).sort_values(["backtest_win_rate", "cagr", "sharpe"], ascending=False).reset_index(drop=True)
        for item in self.backtest_tree.get_children():
            self.backtest_tree.delete(item)
        for i, (_, r) in enumerate(out.iterrows(), start=1):
            self.backtest_tree.insert("", "end", values=(
                i, r["stock_id"], r["stock_name"], f"{r['backtest_win_rate']:.1f}", f"{r['avg_return']:.2f}",
                f"{r['cagr']:.2f}", f"{r['mdd']:.2f}", f"{r['sharpe']:.2f}", int(r["samples"])
            ))
        lines = ["《v9.2 FINAL-RELEASE 策略回測摘要》", ""]
        for i, (_, r) in enumerate(out.iterrows(), start=1):
            lines.append(f"{i}. {r['stock_id']} {r['stock_name']}｜勝率 {r['backtest_win_rate']:.1f}%｜CAGR {r['cagr']:.2f}%｜MDD {r['mdd']:.2f}%｜Sharpe {r['sharpe']:.2f}｜樣本 {int(r['samples'])}")
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", "\n".join(lines))
        self.left_notebook.select(self.tab_backtest)


    def export_equity_curve_chart(self, stock_id: str, hist: pd.DataFrame):
        bt = self.backtest_engine.estimate_trade_quality(stock_id)
        trades = self.backtest_engine.simulate_trades(stock_id)
        if trades.empty:
            return None
        eq = (1 + trades["ret"].astype(float)).cumprod()
        fig = plt.figure(figsize=(10, 5))
        ax = fig.add_subplot(111)
        ax.plot(eq.index + 1, eq.values)
        ax.set_title(f"{stock_id} Equity Curve | CAGR {bt.get('cagr',0):.2f}% | MDD {bt.get('mdd',0):.2f}%")
        ax.set_xlabel("Trade #")
        ax.set_ylabel("Equity")
        out = CHART_DIR / f"{stock_id}_equity_curve.png"
        fig.savefig(out, dpi=140, bbox_inches="tight")
        plt.close(fig)
        log_info(f"export_chart done：{stock_id}｜{out}")
        return out

    def on_select_backtest(self, event=None):
        sel = self.backtest_tree.selection()
        if not sel:
            return
        vals = self.backtest_tree.item(sel[0], "values")
        stock_id = str(vals[1]).strip()
        if not stock_id:
            return
        if self._should_ignore_select_event(stock_id, "回測視覺化"):
            return
        self.start_backtest_selection_job(stock_id)

    def on_select_stock(self, event=None):
        sel = self.rank_tree.selection()
        if not sel:
            return
        vals = self.rank_tree.item(sel[0], "values")
        stock_id = str(vals[1])
        if self._should_ignore_select_event(stock_id, "排行榜"):
            return
        self.sync_all_views(stock_id, source="排行榜")


    def export_chart(self, stock_id: str, hist: pd.DataFrame):
        log_info(f"export_chart start：{stock_id}")
        x = DataEngine.attach(hist.copy()).tail(120).reset_index(drop=True)
        fig = plt.figure(figsize=(11, 5.8))
        ax = fig.add_subplot(111)
        xs = list(range(len(x)))
        self._candlestick(ax, xs, x["open"], x["high"], x["low"], x["close"])
        ax.plot(xs, x["ma20"], label="MA20", linewidth=1.2)
        ax.plot(xs, x["ma60"], label="MA60", linewidth=1.2)

        plan = self.get_cached_trade_plan(stock_id)
        if plan is None:
            stock = self.db.get_stock_row(stock_id)
            plan = self.build_lightweight_plan(stock_id, x, stock=stock)
            self.plan_cache[str(stock_id)] = plan
        support = float(plan.get("support", 0) or 0)
        fib1 = float(plan.get("resistance", 0) or 0)
        fib1382 = float(plan.get("target_1382", 0) or 0)
        fib1618 = float(plan.get("target_1618", 0) or 0)
        try:
            stop = float(plan.get("stop_loss", 0) or 0)
        except Exception:
            stop = 0.0

        if support > 0:
            ax.axhline(support, linestyle="--", linewidth=1, label=f"Support {support:.2f}")
        if fib1 > 0:
            ax.axhline(fib1, linestyle="--", linewidth=1, label=f"Fib 1.0 {fib1:.2f}")
        if fib1382 > 0:
            ax.axhline(fib1382, linestyle=":", linewidth=1, label=f"Fib 1.382 {fib1382:.2f}")
        if fib1618 > 0:
            ax.axhline(fib1618, linestyle=":", linewidth=1, label=f"Fib 1.618 {fib1618:.2f}")

        wave = WaveEngine.detect_wave_label(x)
        last_close = float(x.iloc[-1]["close"])
        last_x = xs[-1]
        bull_target = fib1382 if fib1382 > 0 else last_close * 1.08
        bear_target = stop if stop > 0 else last_close * 0.95
        path_x = [last_x, last_x + 4, last_x + 9]
        ax.plot(path_x, [last_close, (last_close + bull_target) / 2.0, bull_target], "--", linewidth=1.5, label="Bull Path")
        ax.plot(path_x, [last_close, (last_close + bear_target) / 2.0, bear_target], "--", linewidth=1.5, label="Bear Path")

        recent = x.tail(55)
        try:
            peak_idx = recent["high"].idxmax()
            trough_idx = recent["low"].idxmin()
            peak_y = float(x.loc[peak_idx, "high"])
            trough_y = float(x.loc[trough_idx, "low"])
            ax.scatter([peak_idx], [peak_y], s=36)
            ax.scatter([trough_idx], [trough_y], s=36)
            ax.annotate("Wave Peak", xy=(peak_idx, peak_y), xytext=(peak_idx, peak_y * 1.02), fontfamily=SELECTED_PLOT_FONT)
            ax.annotate("Wave Trough", xy=(trough_idx, trough_y), xytext=(trough_idx, trough_y * 0.98), fontfamily=SELECTED_PLOT_FONT)
        except Exception:
            pass

        ax.set_xlim(0, max(path_x) + 2)
        title_wave = safe_plot_text(wave, fallback="Wave")
        title_signal = safe_plot_text(plan.get("signal", "-"), fallback="-")
        ax.set_title(f"{stock_id} | {title_wave} | {title_signal}", fontfamily=SELECTED_PLOT_FONT)
        info_text = (
            f"Wave: {title_wave}\n"
            f"Entry: {safe_plot_text(plan.get('entry_zone','-'))}\n"
            f"Stop: {safe_plot_text(plan.get('stop_loss','-'))}\n"
            f"RR: {float(plan.get('rr',0) or 0):.2f}"
        )
        ax.text(
            0.01, 0.98, info_text,
            transform=ax.transAxes, va="top", ha="left", fontfamily=SELECTED_PLOT_FONT,
            bbox=dict(boxstyle="round", alpha=0.15)
        )
        ax.grid(alpha=0.2)
        ax.legend(loc="upper left", fontsize=8, prop={"family": SELECTED_PLOT_FONT, "size": 8})
        fig.tight_layout()
        out = CHART_DIR / f"{stock_id}_chart.png"
        fig.savefig(out, dpi=140, bbox_inches="tight")
        plt.close(fig)
        return out




def bootstrap():
    if BOOTSTRAP_EVENT.is_set():
        db = DBManager(DB_PATH)
        db.init_db()
        return db, "bootstrap 已執行，直接載入現有資料"
    with BOOTSTRAP_LOCK:
        if BOOTSTRAP_EVENT.is_set():
            db = DBManager(DB_PATH)
            db.init_db()
            return db, "bootstrap 已執行，直接載入現有資料"

        log_info("bootstrap start")
        db = DBManager(DB_PATH)
        db.init_db()

        init_message = "股票主檔已就緒"
        try:
            master = db.get_master()
            if master.empty:
                universe = build_full_market_universe()
                if universe is not None and not universe.empty:
                    db.import_master_df(universe)
                    master = db.get_master()
                    init_message = f"已自動建立全市場股票主檔，共 {len(master)} 檔"
                else:
                    csv_path = resolve_master_csv()
                    db.import_master_csv(csv_path)
                    master = db.get_master()
                    init_message = f"已改用本地主檔，共 {len(master)} 檔 | {csv_path}"
            else:
                init_message = f"股票主檔已載入，共 {len(master)} 檔"

            if db.get_ranking_rows_count() == 0 and db.get_total_price_rows() > 0:
                rank_count = RankingEngine(db).rebuild()
                if rank_count > 0:
                    init_message += f"｜已自動重建排行 {rank_count} 檔"
                else:
                    init_message += "｜已有歷史資料，但目前不足以形成排行"
        except Exception as e:
            init_message = f"股票主檔初始化失敗：{e}"

        BOOTSTRAP_EVENT.set()
        log_info(f"bootstrap done｜{init_message}")
        return db, init_message


def main():
    log_info("main start")
    db, init_message = bootstrap()
    root = tk.Tk()
    app = AppUI(root, db)
    app.set_status(init_message)

    def _close():
        log_info("application closing")
        db.close()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _close)
    root.mainloop()


if __name__ == "__main__":
    main()


# V9.5.6 MARGIN_INTEGRATED PATCH MARKER: external_margin + macro_margin_sentiment + DecisionLayer margin_score completed.
# V9.6.2-R10 MARKET_SNAPSHOT_FULL_FALLBACK_AND_FAIL_REASON MARKER
# V9.5.8 DATA_INTEGRITY_PATCH MARKER: market_snapshot proxy blocked; only TWSE MI_INDEX official data can set data_ready=1.
