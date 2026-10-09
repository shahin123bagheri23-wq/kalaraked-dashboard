import base64
import hmac
import html as _htmllib
import io
import os
import re
import shutil
import sqlite3
from contextlib import closing
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_js_eval import streamlit_js_eval
from personel_loader import load_personel

try:
    from zoneinfo import ZoneInfo
    _TZ = ZoneInfo("Asia/Tehran")
except Exception:
    _TZ = None


def now_tehran():
    return datetime.now(_TZ) if _TZ else datetime.now()


st.set_page_config(
    page_title="داشبورد کالای راکد | افق کوروش",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ================== تشخیص دستگاه ==================
if "screen_width" not in st.session_state:
    st.session_state["screen_width"] = streamlit_js_eval(
        js_expressions="window.innerWidth", key="WIDTH"
    )
screen_width = st.session_state.get("screen_width") or 400
is_mobile = screen_width < 768
is_desktop = not is_mobile

# ================== رنگ برند ==================
OK_RED = "#E6003E"
OK_RED_DARK = "#A30029"
OK_RED_LIGHT = "#FF1F5A"


# ================== پوشه داده ==================
def _resolve_data_dir():
    # اولویت با پوشه خود برنامه است تا فایل‌های Excel کنار app.py
    # به‌صورت خودکار پیدا و ذخیره شوند. در سرویس‌های محدود، مسیرهای جایگزین استفاده می‌شوند.
    candidates = []
    env_dir = os.environ.get("DATA_DIR")
    if env_dir:
        candidates.append(Path(env_dir).expanduser())
    try:
        candidates.append(Path(__file__).resolve().parent)
    except Exception:
        pass
    candidates += [Path.cwd(), Path.home() / ".kalaraked", Path("/tmp")]
    seen = set()
    for c in candidates:
        try:
            c = c.resolve()
            if str(c) in seen:
                continue
            seen.add(str(c))
            c.mkdir(parents=True, exist_ok=True)
            test = c / ".write_test"
            test.write_text("ok", encoding="utf-8")
            test.unlink(missing_ok=True)
            return c
        except Exception:
            continue
    return Path("/tmp")


DATA_DIR = _resolve_data_dir()

try:
    BACKUP_DIR = DATA_DIR / "backups"
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    BACKUP_DIR = DATA_DIR

RAAKED_NAME = "1405-07-15 projraked.xlsx"
SALES_NAME = "گزارش فروش راکد.xlsx"
TARGET_NAME = "فایل تارگت.xlsx"

# مسیر خواندن: اگر فایل در DATA_DIR نبود، از پوشه برنامه می‌خوانیم
RAAKED_FILE = DATA_DIR / RAAKED_NAME
SALES_FILE = DATA_DIR / SALES_NAME
TARGET_FILE = DATA_DIR / TARGET_NAME
# مسیر نوشتن (آپلود): همیشه DATA_DIR
RAAKED_DEST = DATA_DIR / RAAKED_NAME
SALES_DEST = DATA_DIR / SALES_NAME

TARGET_SHEET = "روند و تارگت"
HISTORY_DB = DATA_DIR / "history.db"

def _excel_candidates():
    roots = [DATA_DIR]
    try:
        roots.append(Path(__file__).resolve().parent)
    except Exception:
        pass
    if Path.cwd() not in roots:
        roots.append(Path.cwd())
    out = []
    seen = set()
    for root in roots:
        try:
            for f in root.glob("*.xlsx"):
                key = str(f.resolve())
                if key not in seen and f.name != "history.xlsx":
                    seen.add(key)
                    out.append(f)
        except Exception:
            pass
    return out

def _has_sheets(path, required):
    try:
        names = pd.ExcelFile(path).sheet_names
        norm = {str(x).translate(FA_DIGITS).strip().lower() for x in names}
        return all(str(x).translate(FA_DIGITS).strip().lower() in norm for x in required)
    except Exception:
        return False

def _has_columns(path, required):
    try:
        d = pd.read_excel(path, sheet_name=0, nrows=3)
        cols = {str(x).strip() for x in d.columns}
        return all(any(r in c or c in r for c in cols) for r in required)
    except Exception:
        return False

def _discover_input_files():
    global RAAKED_FILE, SALES_FILE, TARGET_FILE
    roots = [DATA_DIR]
    try:
        roots += [Path(__file__).resolve().parent]
    except Exception:
        pass
    roots += [Path.cwd()]

    def first_existing(name, fallback):
        for root in roots:
            try:
                candidate = root / name
                if candidate.exists():
                    return candidate
            except Exception:
                pass
        return fallback

    # هر فایل باید مستقل کشف شود؛ قبلاً هر سه مسیر به فایل راکد اشاره می‌کردند.
    RAAKED_FILE = first_existing(RAAKED_NAME, DATA_DIR / RAAKED_NAME)
    SALES_FILE = first_existing(SALES_NAME, DATA_DIR / SALES_NAME)
    TARGET_FILE = first_existing(TARGET_NAME, DATA_DIR / TARGET_NAME)
    return RAAKED_FILE, SALES_FILE, TARGET_FILE


RAAKED_FILE, SALES_FILE, TARGET_FILE = _discover_input_files()

# ================== ستون‌ها ==================
BR, BC, NM, QTY, VAL, SUP = ("نام شعبه", "بارکد", "نام کالا", "موجودی سیستمی",
                             "موجودی ریالی اقلام راکد", "سوپروایزر")
BCODE = "کد شعبه"
SHIFT = "شیفت"
ROW_NUM = "ردیف"
S_QTY, S_AMT, S_REL = "فروش (تعداد)", "فروش (ریال)", "ارزش آزادشده"

SALES_BARCODE = "بارکد کالا"
SALES_QTY = "فروش تعدادی"
SALES_AMOUNT = "فروش خالص"
SALES_STORE = "نام انبار/فروشگاه"
SALES_STORE_CODE = "کد انبار/فروشگاه"

SHEET_TBL60 = "tbl60"
SHEET_TBL45 = "tbl45"
SHEET_SUMMARY60 = "60 روزه"
SHEET_PQ45 = "pq45"
SHEET_PQ60 = "pq60"
SHEET_SALES = "tblSales"

S60_PCT_OLD = "درصد راکد قبل"
S60_PCT_NEW = "درصد راکد فعلی"
S60_DIFF_PCT = "افت/رشد درصدی"
S60_DIFF_TXT = "وضعیت"
S60_RATIO = "نسبت راکد"
S60_INV_AMOUNT = "ارزش کل موجودی"

REQUIRED_RAAKED = [BR, BC, NM, QTY, VAL]
REQUIRED_SALES = [SALES_BARCODE, SALES_QTY, SALES_AMOUNT, SALES_STORE]

ACH_OK, ACH_WARN = 100, 80
HOLD_RATE_DEFAULT = 2.0
VALID_PAGES = {"home", "district", "store", "supervisor", "target",
               "analytics", "upload", "shift", "trend", "report", "presentation",
               "my_store"}

FA_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
BARCODE_LIKE_RE = r"^\d{8,}$"

BG, BG2 = "#0f0f12", "#1a1a1f"
TXT, TXT2 = "#e5e7eb", "#9ca3af"
CARD_BG, INPUT_BG = "#1a1a1f", "#26262c"
LIGHT_CARD, LIGHT_CARD_TXT, LIGHT_CARD_TXT2 = "#1a1a1f", OK_RED_LIGHT, "#9ca3af"

CHART_BG, CHART_GRID = "#1a1a1f", "#2e2e38"
CHART_TEXT, CHART_TEXT2 = "#e5e7eb", "#b0b3b8"
CHART_MAIN = "#FF1F5A"
CHART_GRADIENT = ["#3b0a1a", "#7a0e2e", "#a30029", "#E6003E", "#FF1F5A"]
CHART_PIE_GREEN, CHART_PIE_YELLOW, CHART_PIE_RED = "#10b981", "#f59e0b", "#E6003E"
CHART_NEUTRAL = "#6b7280"


def _html(s):
    # خطوط خالی حذف می‌شوند تا مارک‌داون بلاک HTML را وسط کارت نشکند
    return "\n".join(line.lstrip() for line in str(s).strip().splitlines()
                     if line.strip())


def esc(v):
    """escape کردن متن‌های داده‌ای قبل از قرار گرفتن داخل HTML"""
    return _htmllib.escape(str(v))


def st_md(s, **kwargs):
    kwargs.setdefault("unsafe_allow_html", True)
    return st.markdown(_html(s), **kwargs)


st.markdown(f"""
<style>
    :root {{
        --ok-red: {OK_RED};
        --ok-red-dark: {OK_RED_DARK};
        --ok-red-light: {OK_RED_LIGHT};
    }}
    html, body, [class*="css"] {{
        font-family: Tahoma, sans-serif;
        background-color: {BG};
        color: {TXT};
        -webkit-text-size-adjust: 100%;
    }}
    .stApp {{ background-color: {BG}; }}
    .block-container {{
        padding-top: 0.4rem; padding-bottom: 90px;
        padding-left: 0.5rem; padding-right: 0.5rem;
        max-width: 100% !important;
    }}
    input, textarea, select {{
        background-color: {INPUT_BG} !important;
        color: {TXT} !important;
        font-size: 16px !important;
    }}
    div[data-baseweb="select"] > div {{
        background-color: {INPUT_BG} !important;
        color: {TXT} !important;
    }}

    /* ============ هدر ============ */
    .ok-header {{
        background: linear-gradient(135deg, {OK_RED_DARK} 0%, {OK_RED} 55%, {OK_RED_LIGHT} 100%);
        padding: 14px 16px; border-radius: 14px;
        display: flex; align-items: center; justify-content: space-between;
        color: white; margin-bottom: 14px;
        box-shadow: 0 4px 14px rgba(230,0,62,0.35);
    }}
    .ok-header-title {{ font-size: 1rem; font-weight: bold; color: white; margin: 0; line-height: 1.3; }}
    .ok-header-sub {{ font-size: 0.72rem; color: #ffe0e8; margin: 2px 0 0 0; }}
    .ok-logo-wrap {{
        background: white; padding: 4px 8px; border-radius: 10px;
        display: inline-flex; align-items: center; justify-content: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        flex-shrink: 0;
    }}

    /* ============ هیرو ============ */
    .app-hero {{
        background: linear-gradient(135deg, {OK_RED_DARK} 0%, {OK_RED} 55%, {OK_RED_LIGHT} 100%);
        border-radius: 18px;
        padding: 22px 20px;
        margin-bottom: 18px;
        color: white;
        box-shadow: 0 8px 24px rgba(230,0,62,0.35);
        position: relative;
        overflow: hidden;
    }}
    .app-hero::before {{
        content: '';
        position: absolute;
        top: -50%; right: -50%;
        width: 200%; height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
        pointer-events: none;
    }}
    .app-hero h2 {{
        margin: 0 0 6px 0;
        font-size: 1.15rem;
        color: white !important;
        font-weight: bold;
        position: relative;
    }}
    .app-hero p {{
        margin: 0;
        font-size: 0.82rem;
        opacity: 0.95;
        position: relative;
    }}

    /* ============ عنوان بخش ============ */
    .app-section-title {{
        font-size: 0.92rem;
        font-weight: bold;
        color: {TXT};
        margin: 22px 0 12px 0;
        padding-right: 12px;
        border-right: 4px solid {OK_RED};
        position: relative;
    }}

    /* ============ کارت‌های اصلی (secondary buttons) ============ */
    div[data-testid="stButton"] > button[kind="secondary"] {{
        background: linear-gradient(180deg, #1e1e26 0%, #15151c 100%) !important;
        border: 1px solid #2a2a30 !important;
        border-radius: 18px !important;
        padding: 22px 12px !important;
        min-height: 130px !important;
        color: {TXT} !important;
        font-size: 0.85rem !important;
        line-height: 1.4 !important;
        white-space: pre-line !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
        text-align: center !important;
        transition: all 0.2s;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }}
    div[data-testid="stButton"] > button[kind="secondary"]:hover {{
        background: linear-gradient(180deg, #2a1a20 0%, #1e1218 100%) !important;
        border-color: {OK_RED} !important;
        transform: translateY(-3px);
        box-shadow: 0 8px 20px rgba(230,0,62,0.35);
    }}
    div[data-testid="stButton"] > button[kind="secondary"]:active {{
        transform: scale(0.97);
    }}
    div[data-testid="stButton"] > button[kind="secondary"] p {{
        color: {TXT} !important;
        font-size: 0.85rem !important;
        margin: 0 !important;
        line-height: 1.5 !important;
        font-weight: 500 !important;
        white-space: pre-line !important;
        text-align: center !important;
    }}
    div[data-testid="stButton"] > button[kind="secondary"] p strong {{
        font-size: 1.6rem !important;
        display: block !important;
        margin-bottom: 8px !important;
        color: {OK_RED_LIGHT} !important;
    }}

    /* ============ Bottom Nav (primary buttons) ============ */
    div[data-testid="stButton"] > button[kind="primary"] {{
        background: #14141a !important;
        border: 1px solid #2a2a30 !important;
        border-radius: 12px !important;
        padding: 10px 6px !important;
        min-height: 52px !important;
        color: #8b8b95 !important;
        font-size: 0.78rem !important;
        font-weight: 500 !important;
        transition: all 0.15s;
    }}
    div[data-testid="stButton"] > button[kind="primary"]:hover {{
        border-color: {OK_RED_LIGHT} !important;
        color: {OK_RED_LIGHT} !important;
        background: #1a0f15 !important;
    }}

    /* دکمه بازگشت */
    div[data-testid="stButton"] > button[kind="secondary"][title*="بازگشت"] {{
        min-height: 36px !important;
        padding: 6px 14px !important;
        border-radius: 8px !important;
        width: auto !important;
        background: {OK_RED} !important;
        color: white !important;
        border: none !important;
        font-weight: bold !important;
    }}

    /* ============ کارت‌های KPI ============ */
    div[data-testid="stMetric"] {{
        background: {CARD_BG}; padding: 10px; border-radius: 10px;
        margin-bottom: 6px; border-right: 4px solid {OK_RED};
        box-shadow: 0 2px 6px rgba(0,0,0,0.3);
    }}
    div[data-testid="stMetric"] label {{ color: {TXT2} !important; font-size: 0.72rem !important; }}
    div[data-testid="stMetricValue"] {{ color: {OK_RED_LIGHT} !important; font-size: 1rem !important; line-height: 1.2; }}

    /* ============ هدینگ‌ها ============ */
    h1 {{ font-size: 1.1rem !important; }}
    h2 {{ font-size: 1rem !important; }}
    h3 {{ font-size: 0.92rem !important; }}

    /* ============ فوتر ============ */
    .footer-text {{
        text-align: center; color: {TXT2}; font-size: 0.72rem;
        padding: 20px 0 6px 0; border-top: 2px solid {OK_RED}; margin-top: 20px;
    }}
    .footer-text b {{ color: {OK_RED_LIGHT}; }}

    /* ============ تب‌ها ============ */
    button[data-baseweb="tab"] {{
        font-size: 0.82rem !important;
        padding-left: 10px !important; padding-right: 10px !important;
    }}
    button[data-baseweb="tab"][aria-selected="true"] {{
        color: {OK_RED_LIGHT} !important;
        border-bottom-color: {OK_RED} !important;
    }}

    /* ============ اسلایدر ============ */
    div[data-testid="stSlider"] div[role="slider"] {{ background-color: {OK_RED} !important; }}

    /* ============ کارت‌های هشدار ============ */
    .alert-card {{
        background: {CARD_BG}; padding: 12px 14px;
        border-radius: 10px; border-right: 4px solid {OK_RED};
        margin-bottom: 8px; color: {TXT};
        font-size: 0.82rem; line-height: 1.4;
    }}
    .alert-card.warn {{ border-right-color: #f59e0b; }}
    .alert-card.ok {{ border-right-color: #10b981; }}

    /* ============ کارت‌های موبایل ============ */
    .mcard {{
        background: {CARD_BG}; border-right: 4px solid {OK_RED};
        border-radius: 12px; padding: 12px 14px; margin-bottom: 10px;
        color: {TXT}; box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    }}
    .mcard-badge {{
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        color: white;
        font-size: 0.72rem;
        font-weight: bold;
        margin-bottom: 6px;
    }}
    .mcard-title {{
        font-size: 0.9rem; font-weight: bold;
        color: {OK_RED_LIGHT}; margin-bottom: 6px;
        padding-bottom: 6px; border-bottom: 1px solid {TXT2}30; line-height: 1.3;
    }}
    .mcard-sup {{ font-size: 0.72rem; color: {TXT2}; margin-bottom: 8px; }}
    .mcard-row {{
        display: flex; justify-content: space-between;
        align-items: center; padding: 3px 0; font-size: 0.78rem; gap: 10px;
    }}
    .mcard-row span {{ color: {TXT2}; flex-shrink: 0; }}
    .mcard-row b {{ color: {TXT}; font-weight: 600; text-align: left; word-break: break-word; }}

    /* ============ دکمه دانلود ============ */
    .stDownloadButton button {{ font-size: 0.78rem !important; padding: 6px 12px !important; }}

    div[data-testid="stSelectbox"] {{ margin-bottom: 4px; }}
    div[data-testid="stSlider"] {{ margin-bottom: 4px; }}

    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{ display: none; }}

    .bnav-marker {{ height: 8px; }}

    /* ============ فشرده‌سازی متریک‌ها ============ */
    div[data-testid="stMetric"] {{
        background: #1a1a1f;
        padding: 8px 10px;
        border-radius: 10px;
        margin-bottom: 4px;
        border-right: 3px solid #E6003E;
        box-shadow: 0 1px 4px rgba(0,0,0,0.25);
    }}
    div[data-testid="stMetric"] label {{
        color: #9ca3af !important;
        font-size: 0.66rem !important;
        line-height: 1.2 !important;
    }}
    div[data-testid="stMetricValue"] {{
        color: #FF1F5A !important;
        font-size: 0.88rem !important;
        line-height: 1.15 !important;
    }}
    div[data-testid="stMetricDelta"] {{
        font-size: 0.65rem !important;
        line-height: 1.1 !important;
    }}
    div[data-testid="stMetric"] > div {{
        gap: 2px !important;
    }}
    /* کاهش فاصله بین متریک‌ها */
    div[data-testid="stHorizontalBlock"] {{
        gap: 4px !important;
    }}


    /* ============ فشرده‌سازی متریک‌ها ============ */
    div[data-testid="stMetric"] {{
        background: #1a1a1f;
        padding: 8px 10px;
        border-radius: 10px;
        margin-bottom: 4px;
        border-right: 3px solid #E6003E;
        box-shadow: 0 1px 4px rgba(0,0,0,0.25);
    }}
    div[data-testid="stMetric"] label {{
        color: #9ca3af !important;
        font-size: 0.66rem !important;
        line-height: 1.2 !important;
    }}
    div[data-testid="stMetricValue"] {{
        color: #FF1F5A !important;
        font-size: 0.88rem !important;
        line-height: 1.15 !important;
    }}
    div[data-testid="stMetricDelta"] {{
        font-size: 0.65rem !important;
        line-height: 1.1 !important;
    }}
    div[data-testid="stMetric"] > div {{
        gap: 2px !important;
    }}
    div[data-testid="stHorizontalBlock"] {{
        gap: 4px !important;
    }}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
/* UI-OVERRIDE-V1 */
@font-face { font-family: 'Vazirmatn'; font-weight: 100 500; font-display: swap;
  src: url('app/static/Vazirmatn-Regular.woff2') format('woff2'); }
@font-face { font-family: 'Vazirmatn'; font-weight: 600 900; font-display: swap;
  src: url('app/static/Vazirmatn-Bold.woff2') format('woff2'); }

html, body, [class*="css"], .stApp, button, input, textarea, select, label,
p, h1, h2, h3, h4, td, th, .mcard, .ok-header, .alert-card, .app-hero, .kpi-card,
div[data-testid="stMetric"], div[data-testid="stMarkdownContainer"] {
    font-family: 'Vazirmatn', Tahoma, sans-serif !important;
}

/* ---- کارت KPI رنگی ---- */
.kpi-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-bottom: 12px; }
@media (min-width: 769px) { .kpi-grid { grid-template-columns: repeat(4, 1fr); } }
.kpi-card { background: #1a1a1f; border-radius: 14px; padding: 14px 12px; text-align: center;
            border-right: 5px solid #6b7280; box-shadow: 0 2px 6px rgba(0,0,0,.25); }
.kpi-card .kpi-label { color: #9ca3af; font-size: .8rem; margin-bottom: 4px; }
.kpi-card .kpi-value { color: #e5e7eb; font-size: 1.45rem; font-weight: 700; line-height: 1.25; }
.kpi-card.green  { border-right-color: #10b981; } .kpi-card.green  .kpi-value { color: #10b981; }
.kpi-card.yellow { border-right-color: #f59e0b; } .kpi-card.yellow .kpi-value { color: #f59e0b; }
.kpi-card.red    { border-right-color: #E6003E; } .kpi-card.red    .kpi-value { color: #FF1F5A; }

/* ---- موبایل: متن بزرگ‌تر و خواناتر ---- */
@media (max-width: 768px) {
  h1 { font-size: 1.25rem !important; }
  h2 { font-size: 1.12rem !important; }
  h3 { font-size: 1.02rem !important; }
  div[data-testid="stMetric"] { padding: 12px 12px !important; border-radius: 14px !important; }
  div[data-testid="stMetric"] label, div[data-testid="stMetric"] label p,
  div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] p { font-size: .82rem !important; }
  div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] div { font-size: 1.25rem !important; font-weight: 700 !important; }
  div[data-testid="stMetricDelta"] { font-size: .76rem !important; }
  .mcard { padding: 14px 16px !important; border-radius: 14px !important; }
  .mcard-title { font-size: 1.02rem !important; }
  .mcard-sup { font-size: .84rem !important; }
  .mcard-row { font-size: .92rem !important; padding: 5px 0 !important; }
  .mcard-badge { font-size: .8rem !important; }
  .alert-card { font-size: .92rem !important; }
  .ok-header-title { font-size: 1.1rem !important; }
  .ok-header-sub { font-size: .8rem !important; }
  .app-section-title { font-size: 1rem !important; }
  button[data-baseweb="tab"] { font-size: .92rem !important; }
  div[data-testid="stButton"] > button[kind="primary"] { font-size: .85rem !important; }
  div[data-testid="stButton"] > button[kind="secondary"] p { font-size: .92rem !important; }
}
/* روی لمس، افکت hover (پرش کارت) خاموش */
@media (hover: none) {
  div[data-testid="stButton"] > button:hover { transform: none !important; }
}
</style>
""", unsafe_allow_html=True)


# ================== توابع کمکی ==================
def clean_barcode(series):
    return (series.astype(str).str.translate(FA_DIGITS)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip().replace({"nan": "", "None": ""}))


def normalize_name(series):
    return (series.astype(str).str.strip()
            .str.replace("\u200c", " ", regex=False)
            .str.replace("ي", "ی", regex=False)
            .str.replace("ك", "ک", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .replace({"nan": "", "None": ""}))


def normalize_query(q):
    return (q.strip().lower().translate(FA_DIGITS)
            .replace("ي", "ی").replace("ك", "ک"))


def to_number(series):
    if series is None or len(series) == 0:
        return pd.Series([], dtype=float)
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0).astype(float)
    s = (series.astype(str).str.translate(FA_DIGITS)
         .str.replace(",", "", regex=False)
         .str.replace("،", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("\u200c", "", regex=False)
         .str.replace("%", "", regex=False)
         .replace({"nan": None, "": None, "-": None, "None": None, "NaN": None}))
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(float)


def parse_percent(series):
    """تشخیص خودکار نسبت اعشاری از درصد کامل — per-value"""
    raw = series.astype(str)
    has_sign = raw.str.contains("%").any()
    s = (raw.str.translate(FA_DIGITS)
         .str.replace(",", "", regex=False).str.replace("،", "", regex=False)
         .str.replace(" ", "", regex=False).str.replace("\u200c", "", regex=False)
         .str.replace("%", "", regex=False)
         .replace({"nan": None, "": None, "-": None, "None": None}))
    s = pd.to_numeric(s, errors="coerce").fillna(0)
    if has_sign:
        return s

    col_name = str(getattr(series, "name", "") or "")

    # ستون‌های «تحقق»: مقادیر 0-2 هستن، همیشه × ۱۰۰
    if "تحقق" in col_name:
        return s * 100

    # ستون‌های دیگه: per-value — مقادیر زیر ۰.۵ اعشاری، بالای ۰.۵ کامل
    if len(s) > 0:
        s = s.apply(lambda v: v * 100 if 0 < abs(v) < 0.5 else v)
    return s

TOTAL_RE = r"^\s*(?:جمع|مجموع|[کک]ل|total|grand\s*total|sum)(?:\s|$)"


def remove_totals(df, branch_col=BR):
    if branch_col not in df.columns:
        return df
    name = df[branch_col].astype(str).str.strip()
    is_empty = df[branch_col].isna() | (name == "") | (name.str.lower() == "nan")
    looks_total = name.str.contains(TOTAL_RE, case=False, regex=True)
    if BC in df.columns:
        code = df[BC].astype(str).str.strip()
        looks_total &= df[BC].isna() | code.isin(["", "nan"])
    return df[~(is_empty | looks_total)].copy()


def remove_invalid_rows(df):
    if BR not in df.columns:
        return df
    name = df[BR].astype(str).str.strip()
    looks_barcode = name.str.match(BARCODE_LIKE_RE, na=False)
    if VAL in df.columns and QTY in df.columns:
        v_num = to_number(df[VAL])
        q_num = to_number(df[QTY])
        both_zero = (v_num == 0) & (q_num == 0)
    else:
        both_zero = pd.Series(False, index=df.index)
    mask = looks_barcode | both_zero
    return df[~mask].copy()


def smart_int(s):
    s = s.fillna(0)
    try:
        if (s == s.round()).all():
            return s.astype("int64")
    except Exception:
        pass
    return s.round(2)


def money(v, short=False, with_unit=False):
    try:
        v = float(v)
    except (TypeError, ValueError):
        v = 0.0
    if v != v:  # NaN
        v = 0.0
    if not short:
        return f"{v:,.0f} ریال"
    abs_v = abs(v)
    unit = " ریال" if with_unit else ""
    if abs_v >= 1_000_000_000_000:
        return f"{v/1_000_000_000_000:,.1f} همت{unit}"
    if abs_v >= 1_000_000_000:
        return f"{v/1_000_000_000:,.1f} میلیارد{unit}"
    if abs_v >= 1_000_000:
        return f"{v/1_000_000:,.1f} میلیون{unit}"
    if abs_v >= 1_000:
        return f"{v/1_000:,.1f} هزار{unit}"
    return f"{v:,.0f}{unit}"


def valid_date(s):
    try:
        datetime.strptime(s.strip(), "%Y-%m-%d")
        return True
    except (ValueError, AttributeError):
        return False


def find_sheet_exact(xls, name):
    target = name.translate(FA_DIGITS).strip().lower()
    for s in xls.sheet_names:
        if s.translate(FA_DIGITS).strip().lower() == target:
            return s
    return None


def find_sheet_contains(xls, name):
    target = name.translate(FA_DIGITS).strip().lower()
    for s in xls.sheet_names:
        if target in s.translate(FA_DIGITS).strip().lower():
            return s
    return None


# ================== پارسرها ==================
def _clean_code(series):
    return (series.astype(str).str.strip().str.upper()
            .replace({"NAN": "", "NONE": ""}))


def _prep_raaked(d, with_code=True):
    d = d.copy()
    d.columns = d.columns.astype(str).str.strip()
    required = [BR, BC, NM, QTY, VAL]
    if with_code:
        required.append(BCODE)
    missing = [c for c in required if c not in d.columns]
    if missing:
        raise ValueError(
            f"ستون‌های لازم پیدا نشد: {'، '.join(missing)}\n"
            f"ستون‌های موجود: {'، '.join(d.columns.astype(str).tolist())}"
        )
    for drop_col in ("کد شعبه .1", "کد شعبه.1"):
        if drop_col in d.columns:
            d = d.drop(columns=[drop_col])
    d = remove_totals(d)
    d = remove_invalid_rows(d)
    d[BC] = clean_barcode(d[BC])
    d[BR] = normalize_name(d[BR])
    if BCODE in d.columns:
        d[BCODE] = _clean_code(d[BCODE])
    else:
        d[BCODE] = ""
    d[NM] = d[NM].fillna("").astype(str).str.strip()
    if SUP in d.columns:
        d[SUP] = normalize_name(d[SUP]).replace("", "نامشخص")
    else:
        d[SUP] = "نامشخص"
    d[QTY] = to_number(d[QTY])
    d[VAL] = to_number(d[VAL])
    # اگه سوپروایزر خالی بود، از mapping پر کن
    _fix_supervisor_from_map(d)
    return d.reset_index(drop=True)


def _fix_supervisor_from_map(d):
    """پر کردن سوپروایزر خالی از شیت تارگت — با کد شعبه و در صورت نیاز نام شعبه"""
    if d.empty or SUP not in d.columns:
        return d
    try:
        bad = d[SUP].isin(["نامشخص", "", "nan", "None"]) | d[SUP].isna()
        if bad.sum() == 0:
            return d

        xls = pd.ExcelFile(RAAKED_FILE)
        sh = find_sheet_exact(xls, "تارگت") or find_sheet_exact(xls, TARGET_SHEET)
        if not sh:
            return d
        t = pd.read_excel(xls, sheet_name=sh)
        t.columns = t.columns.astype(str).str.strip()

        code_col = next((c for c in t.columns if "کد" in c and "شعبه" in c), None)
        name_col = next((c for c in t.columns if "نام شعبه" in c), None)
        sup_col = next((c for c in t.columns if c.strip() == "سوپروایزر"), None)
        if not sup_col:
            return d

        sup_vals = normalize_name(t[sup_col])

        # mapping با کد
        if code_col and BCODE in d.columns:
            by_code = dict(zip(_clean_code(t[code_col]), sup_vals))
            d.loc[bad, SUP] = d.loc[bad, BCODE].map(by_code).fillna(d.loc[bad, SUP])

        # fallback با نام شعبه
        still = d[SUP].isin(["نامشخص", "", "nan", "None"]) | d[SUP].isna()
        if still.sum() > 0 and name_col and BR in d.columns:
            by_name = dict(zip(normalize_name(t[name_col]), sup_vals))
            d.loc[still, SUP] = d.loc[still, BR].map(by_name).fillna("نامشخص")
    except Exception:
        pass
    return d


def parse_raaked(source):
    xls = pd.ExcelFile(source)
    sheet60 = find_sheet_exact(xls, SHEET_TBL60)
    sheet45 = find_sheet_exact(xls, SHEET_TBL45)
    if not sheet60 or not sheet45:
        raise ValueError(
            f"شیت‌های tbl60 و tbl45 پیدا نشد.\n"
            f"شیت‌های موجود: {'، '.join(xls.sheet_names)}"
        )
    d60 = _prep_raaked(pd.read_excel(xls, sheet_name=sheet60))
    d45 = _prep_raaked(pd.read_excel(xls, sheet_name=sheet45))
    return d60, d45


def _prep_sales(d):
    d = d.copy()
    d.columns = d.columns.astype(str).str.strip()
    rename = {}
    for c in d.columns:
        base = c.replace(".1", "").strip()
        if "بارکد" in base and "کالا" in base and SALES_BARCODE not in d.columns:
            rename[c] = SALES_BARCODE
        elif "نام کالا" in base and NM not in d.columns and NM not in rename.values():
            rename[c] = NM
        elif "فروش تعدادی" in base and SALES_QTY not in d.columns:
            rename[c] = SALES_QTY
        elif ("فروش خالص" in base) and SALES_AMOUNT not in d.columns:
            rename[c] = SALES_AMOUNT
        elif "نام انبار" in base and SALES_STORE not in d.columns:
            rename[c] = SALES_STORE
        elif "کد انبار" in base and SALES_STORE_CODE not in d.columns:
            rename[c] = SALES_STORE_CODE
    if rename:
        d = d.rename(columns=rename)
    missing = [c for c in (SALES_BARCODE, SALES_QTY, SALES_AMOUNT, SALES_STORE)
               if c not in d.columns]
    if missing:
        raise ValueError(f"ستون‌های فروش پیدا نشد: {'، '.join(missing)}")
    d[SALES_BARCODE] = clean_barcode(d[SALES_BARCODE])
    d[SALES_STORE] = normalize_name(d[SALES_STORE])
    d[SALES_QTY] = to_number(d[SALES_QTY])
    d[SALES_AMOUNT] = to_number(d[SALES_AMOUNT])
    if NM in d.columns:
        d[NM] = d[NM].fillna("").astype(str).str.strip()
    if SALES_STORE_CODE in d.columns:
        d[SALES_STORE_CODE] = _clean_code(d[SALES_STORE_CODE])
    return d.reset_index(drop=True)


def _prep_pq(d, default_shift="صبح"):
    d = d.copy()
    d.columns = d.columns.astype(str).str.strip()
    for drop_col in ("کد شعبه .1", "کد شعبه.1", "Index", "Index.1"):
        if drop_col in d.columns:
            d = d.drop(columns=[drop_col])
    required = [BR, BC, NM, QTY, VAL]
    missing = [c for c in required if c not in d.columns]
    if missing:
        raise ValueError(f"ستون‌های لازم pq پیدا نشد: {'، '.join(missing)}")
    d = remove_totals(d)
    d = remove_invalid_rows(d)
    d[BC] = clean_barcode(d[BC])
    d[BR] = normalize_name(d[BR])
    if BCODE in d.columns:
        d[BCODE] = _clean_code(d[BCODE])
    else:
        d[BCODE] = ""
    d[NM] = d[NM].fillna("").astype(str).str.strip()
    if SUP in d.columns:
        d[SUP] = normalize_name(d[SUP]).replace("", "نامشخص")
    else:
        d[SUP] = "نامشخص"
    d[QTY] = to_number(d[QTY])
    d[VAL] = to_number(d[VAL])
    if SHIFT not in d.columns:
        d[SHIFT] = default_shift
    else:
        d[SHIFT] = d[SHIFT].astype(str).str.strip().replace({"nan": default_shift})
    if ROW_NUM in d.columns:
        try:
            d[ROW_NUM] = to_number(d[ROW_NUM]).astype("int64")
        except Exception:
            pass
    return d.reset_index(drop=True)


def _detect_pct_columns(columns):
    """ستون‌های «درصد راکد <عدد>» را به‌ترتیب حضور در شیت پیدا می‌کند
    (اولی = قبل، آخری = فعلی) تا به تاریخ ثابت وابسته نباشد."""
    found = []
    seen = set()
    for c in columns:
        base = re.sub(r"\.\d+$", "", str(c)).translate(FA_DIGITS).strip()
        base = base.replace("\u200c", " ")
        if re.fullmatch(r"درصد\s*راکد\s*\d+", base) and base not in seen:
            seen.add(base)
            found.append(c)
    return found


def _prep_summary60(d):
    d = d.copy()
    d.columns = d.columns.astype(str).str.strip()
    col = {}
    pct_cols = _detect_pct_columns(d.columns)
    if len(pct_cols) >= 2:
        col["pct_old"], col["pct_new"] = pct_cols[0], pct_cols[-1]
    elif len(pct_cols) == 1:
        col["pct_new"] = pct_cols[0]
    for c in d.columns:
        base = c.replace(".1", "").replace(".2", "").strip()
        low = base.lower()
        if "کد شعبه" in base and "code" not in col:
            col["code"] = c
        elif "نام شعبه" in base and "branch" not in col:
            col["branch"] = c
        elif base == "سوپروایزر" and "sup" not in col:
            col["sup"] = c
        elif ("ریالی اقلام راکد" in base or "موجودی ریالی" in base) and "val" not in col:
            col["val"] = c
        elif "تعدادی اقلام راکد" in base and "qty" not in col:
            col["qty"] = c
        elif ("افت/ رشد درصدی" in base or "افت / رشد درصدی" in base
              or "افت/رشد درصدی" in base):
            col["diff_pct"] = c
        elif low in ("افت / رشد", "افت/ رشد", "افت /رشد", "افت/رشد"):
            col["diff_txt"] = c
        elif "نسبت موجودی اقلام راکد" in base and "ratio" not in col:
            col["ratio"] = c
        elif "Inventory Amount" in base and "inv" not in col:
            col["inv"] = c
    out = pd.DataFrame()
    if col.get("branch"):
        out[BR] = normalize_name(d[col["branch"]])
    if col.get("code"):
        out[BCODE] = _clean_code(d[col["code"]])
    if col.get("sup"):
        out[SUP] = normalize_name(d[col["sup"]]).replace("", "نامشخص")
    if col.get("val"):
        out[VAL] = to_number(d[col["val"]])
    if col.get("qty"):
        out[QTY] = to_number(d[col["qty"]])
    if col.get("pct_old"):
        out[S60_PCT_OLD] = parse_percent(d[col["pct_old"]])
    if col.get("pct_new"):
        out[S60_PCT_NEW] = parse_percent(d[col["pct_new"]])
    if col.get("diff_pct"):
        out[S60_DIFF_PCT] = parse_percent(d[col["diff_pct"]])
    if col.get("diff_txt"):
        out[S60_DIFF_TXT] = d[col["diff_txt"]].astype(str).str.strip()
    if col.get("ratio"):
        out[S60_RATIO] = parse_percent(d[col["ratio"]])
    if col.get("inv"):
        out[S60_INV_AMOUNT] = to_number(d[col["inv"]])
    out = remove_totals(out, BR)
    return out.reset_index(drop=True)


def _mtime(path):
    try:
        return Path(path).stat().st_mtime
    except OSError:
        return 0


@st.cache_data(show_spinner="در حال بارگذاری داده‌ها…", max_entries=4)
def load_all(file_path, mtime):
    out = {"d60": None, "d45": None, "pq45": None, "pq60": None,
           "summary60": None, "sales": None, "sheets": []}
    if not Path(file_path).exists():
        return out

    cache_dir = Path(str(file_path)).parent / ".excel_cache"
    try:
        cache_dir.mkdir(exist_ok=True)
    except Exception:
        pass

    mtime_int = int(float(mtime or 0))
    final_cache = cache_dir / f"_processed_v2_{mtime_int}.pkl"

    # اگر نتیجه نهایی قبلاً پردازش و ذخیره شده، مستقیم از Pickle بخوان
    if final_cache.exists():
        try:
            return pd.read_pickle(final_cache)
        except Exception:
            pass

    # اولین بار: اکسل را بخوان و پردازش کن
    try:
        xls = pd.ExcelFile(file_path)
    except Exception:
        return out
    out["sheets"] = list(xls.sheet_names)

    def _read(sheet_name):
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", str(sheet_name))[:60]
        cache_file = cache_dir / f"raw_{safe}.pkl"
        try:
            if cache_file.exists() and cache_file.stat().st_mtime >= mtime_int:
                return pd.read_pickle(cache_file)
        except Exception:
            pass
        df = pd.read_excel(xls, sheet_name=sheet_name)
        try:
            df.to_pickle(cache_file)
        except Exception:
            pass
        return df

    try:
        s60 = find_sheet_exact(xls, SHEET_TBL60)
        s45 = find_sheet_exact(xls, SHEET_TBL45)
        if s60 and s45:
            out["d60"] = _prep_raaked(_read(s60))
            out["d45"] = _prep_raaked(_read(s45))
    except Exception as e:
        st.warning(f"⚠️ خطا در tbl60/tbl45: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_PQ45)
        if s:
            out["pq45"] = _prep_pq(_read(s), "صبح")
    except Exception as e:
        st.warning(f"⚠️ خطا در pq45: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_PQ60)
        if s:
            out["pq60"] = _prep_pq(_read(s), "عصر")
    except Exception as e:
        st.warning(f"⚠️ خطا در pq60: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_SUMMARY60)
        if s:
            out["summary60"] = _prep_summary60(_read(s))
    except Exception as e:
        st.warning(f"⚠️ خطا در خلاصه ۶۰ روزه: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_SALES)
        if s:
            out["sales"] = _prep_sales(_read(s))
    except Exception:
        pass

    # ذخیره نتیجه نهایی پردازش‌شده
    try:
        pd.to_pickle(out, final_cache)
        for old in cache_dir.glob("_processed_v2_*.pkl"):
            if old != final_cache:
                try:
                    old.unlink()
                except Exception:
                    pass
    except Exception:
        pass

    return out





@st.cache_data(show_spinner=False)
def load_sales(mtime):
    if not RAAKED_FILE.exists():
        return None
    try:
        xls = pd.ExcelFile(RAAKED_FILE)
        sh = find_sheet_exact(xls, SHEET_SALES)
        if sh:
            return _prep_sales(pd.read_excel(xls, sheet_name=sh))
    except Exception:
        pass
    return None



@st.cache_data(show_spinner=False)
def load_target(mtime):
    if not RAAKED_FILE.exists():
        return None
    try:
        xls = pd.ExcelFile(RAAKED_FILE)
        sh = find_sheet_exact(xls, "تارگت") or find_sheet_exact(xls, TARGET_SHEET)
        if not sh:
            return None
        t = pd.read_excel(xls, sheet_name=sh)
        t.columns = t.columns.astype(str).str.strip()
        for c in t.columns:
            if "نام شعبه" in c or "سرپرست" in c or "سوپروایزر" in c:
                t[c] = normalize_name(t[c])
            if "ریالی راکد" in c:
                t[c] = to_number(t[c])
            if "درصد" in c or c.startswith("تغییرات") or c.startswith("تارگت") or "تحقق" in c:
                t[c] = parse_percent(t[c])
        return t
    except Exception:
        return None



def detect_target_columns(t):
    if t is None or t.empty:
        return {"branch": None, "code": None, "supervisor": None,
                "zone_supervisor": None, "raked_value": None,
                "trend": [], "target": None, "achievement": None, "change": None}
    cols = list(t.columns)
    shop_sup = next((c for c in cols if "سرپرست فروشگاه" in c or "سرپرست شعبه" in c), None)
    zone_sup = next((c for c in cols if c.strip() == "سوپروایزر"), None)
    if not shop_sup:
        shop_sup = next((c for c in cols if "سرپرست" in c), None)
    if not zone_sup:
        zone_sup = next((c for c in cols if "سوپروایزر" in c), None)
    return {
        "branch": next((c for c in cols if "نام شعبه" in c), None),
        "code": next((c for c in cols if "کد" in c and "شعبه" in c), None),
        "supervisor": shop_sup,
        "zone_supervisor": zone_sup,
        "raked_value": next((c for c in cols if "ریالی راکد" in c), None),
        "trend": [c for c in cols if "درصد" in c and "راکد" in c],
        "target": next((c for c in cols if c.startswith("تارگت") and "تغییرات" not in c), None),
        "achievement": next((c for c in cols if "تحقق" in c), None),
        "change": next((c for c in cols if c.startswith("تغییرات") and "تارگت" not in c), None),
    }


def _connect():
    return closing(sqlite3.connect(HISTORY_DB))


def init_db():
    with _connect() as conn, conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS raaked_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL, branch TEXT NOT NULL, barcode TEXT NOT NULL,
                product_name TEXT, qty REAL, value REAL, supervisor TEXT,
                UNIQUE(date, branch, barcode)
            )""")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sales_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL, store TEXT NOT NULL, barcode TEXT NOT NULL,
                product_name TEXT, qty REAL, amount REAL,
                UNIQUE(date, store, barcode)
            )""")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS updates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL, kind TEXT NOT NULL, rows INTEGER, note TEXT
            )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_raaked_date ON raaked_history(date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sales_date ON sales_history(date)")


def save_raaked_history(df60, df45, date_str, note=""):
    parts = []
    for df, kind in ((df60, "60"), (df45, "45")):
        if df.empty:
            continue
        g = (df.groupby([BR, BC], as_index=False)
             .agg(product_name=(NM, "first"), qty=(QTY, "sum"),
                  value=(VAL, "sum"), supervisor=(SUP, "first")))
        g["date"] = f"{date_str}_{kind}"
        parts.append(g)
    if not parts:
        raise ValueError("فایل راکد هیچ ردیف معتبری ندارد.")
    allp = pd.concat(parts, ignore_index=True)
    records = list(allp[["date", BR, BC, "product_name", "qty", "value", "supervisor"]]
                   .itertuples(index=False, name=None))
    with _connect() as conn, conn:
        conn.execute("DELETE FROM raaked_history WHERE date IN (?, ?)",
                     (f"{date_str}_60", f"{date_str}_45"))
        conn.executemany(
            "INSERT INTO raaked_history (date, branch, barcode, product_name, qty, value, supervisor) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)", records)
        conn.execute("INSERT INTO updates (date, kind, rows, note) VALUES (?, ?, ?, ?)",
                     (date_str, "raaked", len(records), note))
    return len(records)


def save_sales_history(df_sales, date_str, note=""):
    d = df_sales.copy()
    d["_name"] = d[NM].fillna("").astype(str) if NM in d.columns else ""
    g = (d.groupby([SALES_STORE, SALES_BARCODE], as_index=False)
         .agg(product_name=("_name", "first"), qty=(SALES_QTY, "sum"),
              amount=(SALES_AMOUNT, "sum")))
    if g.empty:
        raise ValueError("فایل فروش هیچ ردیف معتبری ندارد.")
    g["date"] = date_str
    records = list(g[["date", SALES_STORE, SALES_BARCODE, "product_name", "qty", "amount"]]
                   .itertuples(index=False, name=None))
    with _connect() as conn, conn:
        conn.execute("DELETE FROM sales_history WHERE date = ?", (date_str,))
        conn.executemany(
            "INSERT INTO sales_history (date, store, barcode, product_name, qty, amount) "
            "VALUES (?, ?, ?, ?, ?, ?)", records)
        conn.execute("INSERT INTO updates (date, kind, rows, note) VALUES (?, ?, ?, ?)",
                     (date_str, "sales", len(records), note))
    return len(records)


def get_last_update():
    try:
        with _connect() as conn:
            return conn.execute(
                "SELECT date, kind, rows FROM updates ORDER BY id DESC LIMIT 1").fetchone()
    except Exception:
        return None


def get_history_dates(kind="raaked"):
    try:
        with _connect() as conn:
            df = pd.read_sql("SELECT DISTINCT date FROM updates WHERE kind = ? ORDER BY date DESC",
                             conn, params=(kind,))
        return df["date"].tolist()
    except Exception:
        return []


def get_history_comparison(date1, date2, kind="60"):
    try:
        q = ("SELECT branch, SUM(value) AS total FROM raaked_history "
             "WHERE date = ? GROUP BY branch")
        with _connect() as conn:
            d1 = pd.read_sql(q, conn, params=(f"{date1}_{kind}",))
            d2 = pd.read_sql(q, conn, params=(f"{date2}_{kind}",))
        m = d1.merge(d2, on="branch", how="outer", suffixes=("_قبل", "_الان")).fillna(0)
        m["تغییر"] = m["total_الان"] - m["total_قبل"]
        m["درصد تغییر"] = np.where(m["total_قبل"] > 0, m["تغییر"] / m["total_قبل"] * 100, 0)
        return m.sort_values("تغییر", ascending=False).reset_index(drop=True)
    except Exception:
        return None


init_db()


def _atomic_write(path, data):
    tmp = Path(str(path) + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def _backup(path, keep=10):
    p = Path(path)
    if not p.exists():
        return
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(p, BACKUP_DIR / f"{now_tehran():%Y%m%d_%H%M%S}__{p.name}")
    files = sorted(BACKUP_DIR.glob(f"*__{p.name}"))
    if len(files) > keep:
        for old in files[:-keep]:
            try:
                old.unlink(missing_ok=True)
            except Exception:
                pass


def data_quality_report(df60, df45, df_sales=None):
    issues = []
    for label, df, bcol, ncol, vcol in [
        ("راکد ۶۰ روزه", df60, BC, NM, VAL),
        ("راکد ۴۵ روزه", df45, BC, NM, VAL),
    ]:
        if df is None:  # این بخش در این آپلود بررسی نمی‌شود
            continue
        if df.empty:
            issues.append(f"⚠️ **{label}**: خالی است")
            continue
        n = len(df)
        no_bc = int((df[bcol] == "").sum())
        no_name = int((df[ncol] == "").sum())
        zero_val = int((df[vcol] == 0).sum())
        dup = int(df.duplicated(subset=[BR, bcol]).sum())
        if no_bc or no_name or zero_val or dup:
            msg = f"📋 **{label}** ({n:,} ردیف):"
            if no_bc: msg += f" بارکد خالی: {no_bc}"
            if no_name: msg += f" | نام خالی: {no_name}"
            if zero_val: msg += f" | ارزش صفر: {zero_val}"
            if dup: msg += f" | تکراری: {dup}"
            issues.append(msg)
    if df_sales is not None and not df_sales.empty:
        n = len(df_sales)
        no_bc = int((df_sales[SALES_BARCODE] == "").sum())
        zero_amt = int((df_sales[SALES_AMOUNT] == 0).sum())
        if no_bc or zero_amt:
            msg = f"📋 **فروش** ({n:,} ردیف):"
            if no_bc: msg += f" بارکد خالی: {no_bc}"
            if zero_amt: msg += f" | مبلغ صفر: {zero_amt}"
            issues.append(msg)
    return issues


def process_raaked_upload(data, date_str, note):
    d60, d45 = parse_raaked(io.BytesIO(data))
    rows = save_raaked_history(d60, d45, date_str, note)
    _backup(RAAKED_DEST)
    _atomic_write(RAAKED_DEST, data)
    return rows, data_quality_report(d60, d45)


def process_sales_upload(data, date_str, note):
    s = _prep_sales(pd.read_excel(io.BytesIO(data), sheet_name=0))
    rows = save_sales_history(s, date_str, note)
    _backup(SALES_DEST)
    _atomic_write(SALES_DEST, data)
    return rows, data_quality_report(None, None, s)


def get_logo_html(size=38):
    logo_path = Path("logo.png")
    if logo_path.exists():
        data = base64.b64encode(logo_path.read_bytes()).decode()
        return f'<img src="data:image/png;base64,{data}" style="height:{size}px;">'
    return f'''
    <svg height="{size}" viewBox="0 0 200 80" xmlns="http://www.w3.org/2000/svg">
        <circle cx="40" cy="40" r="32" fill="none" stroke="{OK_RED}" stroke-width="6"/>
        <text x="40" y="55" text-anchor="middle" font-family="Arial Black" font-size="38" font-weight="900" fill="{OK_RED}">OK</text>
        <text x="90" y="38" font-family="Tahoma" font-size="20" font-weight="bold" fill="{OK_RED}">افق</text>
        <text x="90" y="62" font-family="Tahoma" font-size="20" font-weight="bold" fill="{OK_RED}">کوروش</text>
    </svg>
    '''


def render_header(subtitle=""):
    st_md(f"""
    <div class="ok-header">
        <div>
            <p class="ok-header-title">داشبورد مدیریت کالای راکد</p>
            <p class="ok-header-sub">{esc(subtitle or "فروشگاه‌های زنجیره‌ای افق کوروش")}</p>
        </div>
        <div class="ok-logo-wrap">{get_logo_html(size=32 if is_mobile else 45)}</div>
    </div>
    """)


def back_link():
    if st.button("⬅️ بازگشت", key=f"back_btn_{st.session_state.get('page', 'home')}",
                 type="primary"):
        st.session_state["page"] = "home"
        st.rerun()



def metric_row(items):
    # روی موبایل: ۲ ستون، روی دسکتاپ: همه در یک ردیف
    if is_mobile:
        per_row = 2
    else:
        per_row = max(len(items), 1)
    for i in range(0, len(items), per_row):
        chunk = items[i:i + per_row]
        cols = st.columns(len(chunk), gap="small")
        for col, it in zip(cols, chunk):
            col.metric(it[0], it[1], **(it[2] if len(it) > 2 else {}))


def plotly_style(fig, title_size=14, show_legend_bg=False, mobile_height=None):
    height = mobile_height if (is_mobile and mobile_height) else None
    fig.update_layout(
        paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG,
        font=dict(color=CHART_TEXT, size=11, family="Tahoma"),
        title_font=dict(color=CHART_MAIN, size=title_size, family="Tahoma"),
        legend=dict(font=dict(color=CHART_TEXT, size=10),
                    bgcolor="rgba(0,0,0,0)" if not show_legend_bg else CHART_BG),
        xaxis=dict(tickfont=dict(color=CHART_TEXT, size=10),
                   title_font=dict(color=CHART_TEXT2, size=11),
                   gridcolor=CHART_GRID, linecolor=CHART_GRID, zerolinecolor=CHART_GRID),
        yaxis=dict(tickfont=dict(color=CHART_TEXT, size=10),
                   title_font=dict(color=CHART_TEXT2, size=11),
                   gridcolor=CHART_GRID, linecolor=CHART_GRID, zerolinecolor=CHART_GRID),
        margin=dict(l=5, r=15, t=45, b=10),
    )
    if height:
        fig.update_layout(height=height)
    return fig


def _fmt_cell(v, col_name=""):
    if pd.isna(v):
        return "—"
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        if "٪" in col_name or "درصد" in col_name or "تحقق" in col_name:
            return f"{v:.1f}%"
        if abs(v) >= 1_000:
            return f"{v:,.0f}"
        if v == int(v):
            return f"{int(v):,}"
        return f"{v:.2f}"
    return str(v)


def _badge_html(text, bg_color, icon=""):
    label = f"{icon} {text}".strip()
    return f'<div class="mcard-badge" style="background:{bg_color};">{label}</div>'


def render_mobile_cards(df, max_rows=30, title_col=BR, sub_col=SUP,
                        priority_cols=None, badge_col=None):
    if df.empty:
        st.info("ردیفی برای نمایش وجود ندارد.")
        return
    if priority_cols is None:
        priority_cols = [c for c in df.columns if c not in (BR, SUP)]
    available_cols = [c for c in priority_cols if c in df.columns and c != badge_col]
    display = df.head(max_rows)
    for _, row in display.iterrows():
        title = _fmt_cell(row.get(title_col, "")) if title_col in df.columns else ""
        sub = _fmt_cell(row.get(sub_col, "")) if (sub_col and sub_col in df.columns) else ""
        if not title:
            title = sub
            sub = ""
        badge_html = ""
        if badge_col and badge_col in df.columns:
            badge_val = str(row.get(badge_col, ""))
            if "فوری" in badge_val:
                badge_html = _badge_html("فوری", OK_RED, "🔴")
            elif "پیگیری" in badge_val:
                badge_html = _badge_html("پیگیری", "#f59e0b", "🟡")
            elif "عادی" in badge_val:
                badge_html = _badge_html("عادی", CHART_PIE_GREEN, "🟢")
        rows_html = ""
        for c in available_cols:
            v = row.get(c)
            if pd.isna(v):
                continue
            val_str = esc(_fmt_cell(v, c))
            rows_html += f'<div class="mcard-row"><span>{esc(c)}</span><b>{val_str}</b></div>'
        sub_html = f'<div class="mcard-sup">👤 {esc(sub)}</div>' if sub else ""
        icon = "🏪"
        if "کالا" in title_col:
            icon = "📦"
        elif "سوپروایزر" in title_col or "سرپرست" in title_col:
            icon = "👤"
        st_md(f"""
        <div class="mcard">
            {badge_html}
            <div class="mcard-title">{icon} {esc(title)}</div>
            {sub_html}
            {rows_html}
        </div>
        """)
    if len(df) > max_rows:
        st.caption(f"🔸 نمایش {max_rows} ردیف اول از {len(df):,} — برای دیدن همه، CSV را دانلود کن.")


def filter_by_search(view, key, placeholder):
    q = st.text_input("🔍 " + placeholder, key=f"search_{key}", placeholder=placeholder)
    if q and q.strip():
        terms = [normalize_query(t) for t in q.strip().split() if t.strip()]
        if not terms:
            return view
        mask = pd.Series(True, index=view.index)
        for term in terms:
            term_mask = pd.Series(False, index=view.index)
            for col in view.columns:
                txt = (view[col].astype(str).str.lower().str.translate(FA_DIGITS)
                       .str.replace("ي", "ی", regex=False).str.replace("ك", "ک", regex=False))
                term_mask |= txt.str.contains(term, na=False, regex=False)
            mask &= term_mask
        filtered = view[mask]
        st.caption(f"🔸 {len(filtered):,} ردیف پیدا شد (از {len(view):,})")
        return filtered
    return view


def calc_width(df, col, is_numeric=False, is_percent=False):
    try:
        max_content = df[col].astype(str).str.len().max()
        if pd.isna(max_content):
            max_content = 5
    except Exception:
        max_content = 5
    max_len = max(int(max_content), len(str(col)))
    if is_percent:
        return int(min(max(max_len * 11 + 25, 105), 450))
    if is_numeric:
        return int(min(max(max_len * 10 + 25, 105), 450))
    return int(min(max(max_len * 9 + 25, 90), 480))


def column_config(df):
    cfg = {}
    for col in df.columns:
        if col in (QTY, S_QTY):
            v = df[col].fillna(0)
            fmt = "%d" if bool((v == v.round()).all()) else "%.2f"
            cfg[col] = st.column_config.NumberColumn(col, format=fmt,
                                                     width=calc_width(df, col, is_numeric=True))
        elif col in (VAL, S_AMT, S_REL):
            cfg[col] = st.column_config.NumberColumn(col, format="%,d",
                                                     width=calc_width(df, col, is_numeric=True))
        else:
            cfg[col] = st.column_config.Column(col, width=calc_width(df, col))
    return cfg


def target_column_config(df):
    cfg = {}
    for c in df.columns:
        if "ریالی راکد" in c:
            cfg[c] = st.column_config.NumberColumn(c, format="%,d",
                                                   width=calc_width(df, c, is_numeric=True))
        elif "تحقق" in c or "درصد" in c or c.startswith("تغییرات") or c.startswith("تارگت"):
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f%%",
                                                   width=calc_width(df, c, is_percent=True))
        else:
            cfg[c] = st.column_config.Column(c, width=calc_width(df, c))
    return cfg


def summary_config(df):
    cfg = {}
    for c in df.columns:
        if c in (BR, SUP):
            cfg[c] = st.column_config.Column(c, width=calc_width(df, c))
        elif "٪" in c:
            cfg[c] = st.column_config.NumberColumn(c, format="%.1f")
        elif c in ("تعداد شعب", "تعداد قلم"):
            cfg[c] = st.column_config.NumberColumn(c, format="%d")
        else:
            cfg[c] = st.column_config.NumberColumn(c, format="%,d",
                                                   width=calc_width(df, c, is_numeric=True))
    return cfg


def show_table(df, key, hide_supervisor=False, search=True, placeholder="جستجو...",
               priority_cols=None, title_col=NM, badge_col=None):
    view = df.copy()
    if hide_supervisor and SUP in view.columns:
        view = view.drop(columns=[SUP])
    if search and not view.empty:
        view = filter_by_search(view, key, placeholder)
    if view.empty:
        st.info("ردیفی برای نمایش وجود ندارد.")
        return
    if is_mobile:
        render_mobile_cards(view, max_rows=50, title_col=title_col,
                            sub_col=SUP if SUP in view.columns else None,
                            priority_cols=priority_cols, badge_col=badge_col)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True,
                     column_config=column_config(view))
    st.download_button("⬇️ دانلود CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       f"{key}.csv", "text/csv", key=f"dl_{key}")


def show_summary(df, key):
    if df.empty:
        st.info("داده‌ای برای نمایش وجود ندارد.")
        return
    if is_mobile:
        priority = [BR, SUP, "راکد ۶۰ روزه", "راکد ۴۵ روزه",
                    "تعداد قلم", "فروش (ریال)"]
        render_mobile_cards(df, max_rows=30, title_col=BR, sub_col=SUP,
                            priority_cols=priority)
    else:
        st.dataframe(df, use_container_width=True, hide_index=True,
                     column_config=summary_config(df))
    st.download_button("⬇️ دانلود CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       f"{key}.csv", "text/csv", key=f"dl_{key}")


def export_to_excel(dfs: dict, filename="report.xlsx"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wrote = False
        for name, df in dfs.items():
            if df is None or df.empty:
                continue
            safe = re.sub(r'[\\/*?:\[\]]', '', str(name))[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            wrote = True
        if not wrote:
            pd.DataFrame({"پیام": ["داده‌ای برای خروجی وجود ندارد"]}).to_excel(
                writer, sheet_name="خالی", index=False)
    buf.seek(0)
    return buf.getvalue()


def build_sales_map(sales_df):
    if sales_df is None or sales_df.empty:
        return None
    return (sales_df.groupby([SALES_STORE, SALES_BARCODE])
            .agg({SALES_QTY: "sum", SALES_AMOUNT: "sum"})
            .reset_index()
            .rename(columns={SALES_STORE: BR, SALES_BARCODE: BC,
                             SALES_QTY: S_QTY, SALES_AMOUNT: S_AMT}))


def _fast_df_hash(df):
    if df is None:
        return "none"
    try:
        return (id(df), df.shape, tuple(df.columns), len(df))
    except Exception:
        return "err"


@st.cache_data(show_spinner=False, max_entries=32,
               hash_funcs={pd.DataFrame: _fast_df_hash})
def attach_sales(df_raaked, sales_df):
    if df_raaked is None or df_raaked.empty:
        return pd.DataFrame(columns=[BR, BC, NM, QTY, VAL, S_QTY, S_AMT, S_REL, SUP])
    base = (df_raaked.groupby([BR, BC], as_index=False)
            .agg(**{NM: (NM, "first"), QTY: (QTY, "sum"),
                    VAL: (VAL, "sum"), SUP: (SUP, "first")}))
    sm = build_sales_map(sales_df)
    if sm is not None:
        out = base.merge(sm, on=[BR, BC], how="left")
    else:
        out = base.copy()
        out[S_QTY] = 0.0
        out[S_AMT] = 0.0
    out[S_QTY] = out[S_QTY].fillna(0)
    out[S_AMT] = out[S_AMT].fillna(0)
    unit = (out[VAL] / out[QTY]).where(out[QTY] > 0, 0)
    sold_cap = np.minimum(out[S_QTY].clip(lower=0), out[QTY].clip(lower=0))
    out[S_REL] = (sold_cap * unit).round(0).astype("int64")
    out[S_QTY] = smart_int(out[S_QTY])
    out[S_AMT] = out[S_AMT].round(0).astype("int64")
    cols = [BR, BC, NM, QTY, VAL, S_QTY, S_AMT, S_REL, SUP]
    return out[cols].sort_values(VAL, ascending=False).reset_index(drop=True)


def unsold_replacement_list(df_raaked, sales_df, n=6):
    """لیست نهایی اقلام راکدِ بدون فروش؛ فروش‌رفته‌ها حذف و از ردیف بعدی جایگزین می‌شوند."""
    if df_raaked is None or df_raaked.empty or n <= 0:
        return pd.DataFrame(columns=[BR, BC, NM, QTY, VAL, S_QTY, S_AMT, S_REL, SUP])
    x = attach_sales(df_raaked, sales_df)
    # ابتدا بر اساس ارزش راکد مرتب است؛ سپس هر قلم فروخته‌شده کنار گذاشته می‌شود.
    x = x[x[S_QTY] <= 0].sort_values(VAL, ascending=False).reset_index(drop=True)
    return x.head(int(n))


def participation_summary(df_raaked, sales_df, n=6):
    """وضعیت مشارکت فروش برای ۶ قلم اول راکد هر شعبه را گزارش می‌کند."""
    if df_raaked is None or df_raaked.empty:
        return pd.DataFrame()
    x = attach_sales(df_raaked, sales_df)
    x = x.sort_values([BR, VAL], ascending=[True, False]).copy()
    x["رتبه راکد"] = x.groupby(BR).cumcount() + 1
    first = x[x["رتبه راکد"] <= int(n)].copy()
    if first.empty:
        return pd.DataFrame()
    g = first.groupby(BR).agg(
        **{
            "تعداد اقلام بررسی": (BC, "count"),
            "تعداد اقلام فروخته‌شده": (S_QTY, lambda s: int((s > 0).sum())),
            "تعداد اقلام بدون فروش": (S_QTY, lambda s: int((s <= 0).sum())),
            "فروش تعدادی": (S_QTY, "sum"),
            "فروش ریالی": (S_AMT, "sum"),
        }
    ).reset_index()
    g["وضعیت مشارکت"] = np.where(
        g["تعداد اقلام فروخته‌شده"] <= 0, "بدون مشارکت",
        np.where(g["تعداد اقلام بدون فروش"] <= 0, "مشارکت کامل", "مشارکت ناقص")
    )
    return g.sort_values(["تعداد اقلام فروخته‌شده", "فروش ریالی"], ascending=[True, False]).reset_index(drop=True)




def collapse_group(labels, key, key_suffix=""):
    """دکمه‌های بازشو به جای تب — پیش‌فرض بسته"""
    state_key = f"collapse_{key}_{key_suffix}"
    if state_key not in st.session_state:
        st.session_state[state_key] = None

    cols = st.columns(len(labels))
    for i, (col, label) in enumerate(zip(cols, labels)):
        with col:
            is_active = st.session_state[state_key] == i
            btn_type = "primary" if is_active else "secondary"
            if st.button(label, key=f"cbtn_{key}_{i}_{key_suffix}",
                         use_container_width=True, type=btn_type):
                st.session_state[state_key] = None if is_active else i
                st.rerun()

    return st.session_state[state_key]


# ================== KPI ها ==================
def render_kpis(f60, f45):
    if is_mobile:
        c1, c2 = st.columns(2, gap="small")
        c1.metric("💰 ارزش راکد ۴۵", money(f45[VAL].sum(), short=True, with_unit=True))
        c2.metric("📦 اقلام ۶۰", f"{len(f60):,}")
        st.metric("📦 اقلام ۴۵", f"{len(f45):,}")
    else:
        metric_row([
            ("ارزش راکد ۴۵ روزه", money(f45[VAL].sum(), short=False)),
            ("تعداد اقلام ۶۰ روزه", f"{len(f60):,}"),
            ("تعداد اقلام ۴۵ روزه", f"{len(f45):,}"),
        ])


def render_management_kpis(df60, df45, sales_df=None):
    v60, v45 = df60[VAL].sum(), df45[VAL].sum()
    released = sold = 0
    if sales_df is not None and not sales_df.empty and not df60.empty:
        m = attach_sales(df60, sales_df)
        released, sold = m[S_REL].sum(), m[S_QTY].sum()
    share = (v60 / v45 * 100) if v45 else 0

    if is_mobile:
        r1c1, r1c2 = st.columns(2, gap="small")
        r1c1.metric("💰 ارزش راکد ۶۰ روزه", money(v60, short=True, with_unit=True))
        r1c2.metric("📊 اختلاف ۴۵ و ۶۰", money(v45 - v60, short=True, with_unit=True),
                    delta=f"سهم: {share:.1f}%", delta_color="off")
        r2c1, r2c2 = st.columns(2, gap="small")
        r2c1.metric("🛒 آزادشده", money(released, short=True, with_unit=True))
        r2c2.metric("📦 فروش (تعداد)", f"{sold:,.0f}")
    else:
        metric_row([
            ("💰 ارزش راکد ۶۰ روزه", money(v60)),
            ("📊 اختلاف ارزش ۴۵ و ۶۰ روزه", money(v45 - v60),
             {"delta": f"سهم ۶۰ روزه از ۴۵ روزه: {share:.1f}%", "delta_color": "off"}),
            ("🛒 ارزش راکدِ آزادشده", money(released),
             {"help": "برای هر قلم: min(فروش، موجودی) × قیمت واحد"}),
            ("📦 فروش تعدادی از اقلام راکد", f"{sold:,.0f}"),
        ])


def render_top_products(f60, f45, sales_df, key_suffix="", hide_sup=False):
    top_n = st.slider("چند قلم نمایش داده شود؟", 5, 50, 20, step=5,
                      key=f"slider_{key_suffix}")

    # دکمه‌های کارتی ۶۰ و ۴۵
    col1, col2 = st.columns(2, gap="small")
    with col1:
        show_60 = st.session_state.get(f"show60_{key_suffix}", False)
        if st.button("📅  ۶۰ روزه" + ("  ✓" if show_60 else ""),
                     key=f"btn60_{key_suffix}",
                     use_container_width=True,
                     type="primary" if show_60 else "secondary"):
            st.session_state[f"show60_{key_suffix}"] = not show_60
            st.rerun()

    with col2:
        show_45 = st.session_state.get(f"show45_{key_suffix}", False)
        if st.button("📅  ۴۵ روزه" + ("  ✓" if show_45 else ""),
                     key=f"btn45_{key_suffix}",
                     use_container_width=True,
                     type="primary" if show_45 else "secondary"):
            st.session_state[f"show45_{key_suffix}"] = not show_45
            st.rerun()

    # نمایش محتوا
    if st.session_state.get(f"show60_{key_suffix}"):
        unsold = unsold_replacement_list(f60, sales_df, top_n)
        show_table(unsold, f"top60_{key_suffix}",
                   hide_supervisor=hide_sup, placeholder="نام کالا، بارکد...",
                   title_col=NM,
                   priority_cols=[BR, QTY, VAL, S_QTY, S_AMT, S_REL])

    if st.session_state.get(f"show45_{key_suffix}"):
        unsold = unsold_replacement_list(f45, sales_df, top_n)
        show_table(unsold, f"top45_{key_suffix}",
                   hide_supervisor=hide_sup, placeholder="نام کالا، بارکد...",
                   title_col=NM,
                   priority_cols=[BR, QTY, VAL, S_QTY, S_AMT, S_REL])

    if not st.session_state.get(f"show60_{key_suffix}") and not st.session_state.get(f"show45_{key_suffix}"):
        st.info("👆 روی یکی از دکمه‌ها بزن تا لیست کالاها نمایش داده بشه.")


def render_matched(f60, f45, sales_df, key_suffix="", hide_sup=False):
    if sales_df is None or sales_df.empty:
        st.info("برای مشاهده این بخش، فایل فروش لازم است.")
        return

    def matched(df):
        m = attach_sales(df, sales_df)
        return m[m[S_QTY] > 0].sort_values(S_AMT, ascending=False).reset_index(drop=True)

    m60, m45 = matched(f60), matched(f45)
    metric_row([("کالای راکد ۶۰ روزه فروش رفته", f"{len(m60):,}"),
                ("کالای راکد ۴۵ روزه فروش رفته", f"{len(m45):,}")])
    active = collapse_group(["📅 ۶۰ روزه فروش‌رفته", "📅 ۴۵ روزه فروش‌رفته"],
                            "matched", key_suffix)
    if active == 0:
        if m60.empty:
            st.info("از اقلام راکد ۶۰ روزه، چیزی فروش نرفت.")
        else:
            show_table(m60, f"m60_{key_suffix}", hide_supervisor=hide_sup,
                       placeholder="نام کالا، بارکد...", title_col=NM,
                       priority_cols=[BR, QTY, VAL, S_QTY, S_AMT, S_REL])
    elif active == 1:
        if m45.empty:
            st.info("از اقلام راکد ۴۵ روزه، چیزی فروش نرفت.")
        else:
            show_table(m45, f"m45_{key_suffix}", hide_supervisor=hide_sup,
                       placeholder="نام کالا، بارکد...", title_col=NM,
                       priority_cols=[BR, QTY, VAL, S_QTY, S_AMT, S_REL])
    else:
        st.info("👆 روی یکی از دکمه‌های بالا بزن.")


def render_action_center(df60, sales_df=None, limit=15, key="ac"):
    st.markdown("### 🚨 مرکز اقدام فوری")
    if df60.empty:
        st.info("داده‌ای وجود ندارد.")
        return
    x = attach_sales(df60, sales_df)
    no_sale = x[S_QTY] <= 0
    thr = x[VAL].quantile(0.75)
    x["_p"] = 2
    x.loc[no_sale, "_p"] = 1
    x.loc[no_sale & (x[VAL] >= thr), "_p"] = 0
    x["اولویت"] = x["_p"].map({0: "فوری", 1: "پیگیری", 2: "عادی"})
    x = x.sort_values(["_p", VAL], ascending=[True, False])

    st.caption("🔴 فوری = بدون فروش و جزو ۲۵٪ بالای ارزش | 🟡 پیگیری = بدون فروش | 🟢 عادی = فروش رفته")

    limit = st.slider("تعداد ردیف", 10, 100, 15, 5, key=f"ac_lim_{key}")
    csv_bytes = (x.drop(columns=["_p"]).to_csv(index=False).encode("utf-8-sig"))

    if is_mobile:
        full = x[["اولویت", BR, NM, BC, QTY, VAL, S_QTY, S_AMT]].head(limit)
        for _, r in full.iterrows():
            prio = r["اولویت"]
            if prio == "فوری":
                badge_bg = OK_RED
                badge_icon = "🔴"
            elif prio == "پیگیری":
                badge_bg = "#f59e0b"
                badge_icon = "🟡"
            else:
                badge_bg = CHART_PIE_GREEN
                badge_icon = "🟢"

            st_md(f"""
            <div class="mcard" style="border-right-color: {badge_bg};">
                <div class="mcard-badge" style="background:{badge_bg};">{badge_icon} {prio}</div>
                <div class="mcard-title">📦 {esc(r[NM])}</div>
                <div class="mcard-sup">🏪 {esc(r[BR])}</div>
                <div class="mcard-row"><span>بارکد</span><b>{esc(r[BC])}</b></div>
                <div class="mcard-row"><span>موجودی</span><b>{esc(_fmt_cell(r[QTY]))}</b></div>
                <div class="mcard-row"><span>ارزش راکد</span><b>{esc(_fmt_cell(r[VAL]))} ریال</b></div>
                <div class="mcard-row"><span>فروش</span><b>{esc(_fmt_cell(r[S_QTY]))} عدد</b></div>
            </div>
            """)
    else:
        full = x[["اولویت", BR, NM, BC, QTY, VAL, S_QTY, S_AMT, SUP]].head(limit)
        st.dataframe(full, use_container_width=True, hide_index=True,
                     column_config=column_config(full))

    st.download_button("⬇️ دانلود فهرست کامل (CSV)", csv_bytes,
                       f"action_center_{key}.csv", "text/csv", key=f"dl_ac_{key}")


def render_participation(df60, sales_df, key="participation"):
    if sales_df is None or sales_df.empty:
        st.info("برای محاسبه مشارکت فروش، فایل فروش لازم است.")
        return
    st.subheader("📊 مشارکت شعب در فروش ۶ قلم اول")
    p = participation_summary(df60, sales_df, 6)
    if p.empty:
        st.info("داده‌ای برای محاسبه مشارکت وجود ندارد.")
        return
    c1, c2, c3 = st.columns(3)
    c1.metric("بدون مشارکت", f"{(p['وضعیت مشارکت'] == 'بدون مشارکت').sum():,} شعبه")
    c2.metric("مشارکت ناقص", f"{(p['وضعیت مشارکت'] == 'مشارکت ناقص').sum():,} شعبه")
    c3.metric("مشارکت کامل", f"{(p['وضعیت مشارکت'] == 'مشارکت کامل').sum():,} شعبه")
    show_table(p, key, hide_supervisor=True, placeholder="جستجوی شعبه...", title_col=BR,
               priority_cols=[BR, "تعداد اقلام بررسی", "تعداد اقلام فروخته‌شده",
                              "تعداد اقلام بدون فروش", "فروش تعدادی", "فروش ریالی", "وضعیت مشارکت"])



def _finish_summary(out, int_cols):
    for c in int_cols:
        if c in out.columns:
            out[c] = out[c].fillna(0).round(0).astype("int64")
    out["سهم ۶۰ از ۴۵ (٪)"] = np.where(
        out["راکد ۴۵ روزه"] > 0, out["راکد ۶۰ روزه"] / out["راکد ۴۵ روزه"] * 100, 0).round(1)
    return out


@st.cache_data(show_spinner=False, max_entries=16,
               hash_funcs={pd.DataFrame: _fast_df_hash})
def branch_summary(df60, df45, sales_df=None):
    if df60.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        SUP: df60.groupby(BR)[SUP].first(),
        "راکد ۶۰ روزه": df60.groupby(BR)[VAL].sum(),
        "راکد ۴۵ روزه": df45.groupby(BR)[VAL].sum(),
        "تعداد قلم": df60.groupby(BR)[BC].count(),
    })
    out[SUP] = out[SUP].fillna("نامشخص")
    out["راکد ۴۵ روزه"] = out["راکد ۴۵ روزه"].fillna(0)
    out["تعداد قلم"] = out["تعداد قلم"].fillna(0).astype("int64")
    out["اختلاف ۴۵ و ۶۰"] = out["راکد ۴۵ روزه"] - out["راکد ۶۰ روزه"]
    out["فروش (تعداد)"] = 0.0
    out["فروش (ریال)"] = 0.0
    if sales_df is not None and not sales_df.empty:
        matched = attach_sales(df60, sales_df)
        s = matched.groupby(BR)[[S_QTY, S_AMT]].sum()
        out["فروش (تعداد)"] = s[S_QTY].reindex(out.index).fillna(0)
        out["فروش (ریال)"] = s[S_AMT].reindex(out.index).fillna(0)
    out.index.name = BR
    out = out.reset_index()
    out = _finish_summary(out, ["راکد ۶۰ روزه", "راکد ۴۵ روزه", "اختلاف ۴۵ و ۶۰",
                                "فروش (تعداد)", "فروش (ریال)"])
    return out.sort_values("راکد ۶۰ روزه", ascending=False).reset_index(drop=True)


@st.cache_data(show_spinner=False, max_entries=16,
               hash_funcs={pd.DataFrame: _fast_df_hash})
def supervisor_summary(df60, df45, sales_df=None):
    if df60.empty:
        return pd.DataFrame()
    bmap = df60.groupby(BR)[SUP].first()
    d45 = df45.assign(_s=df45[BR].map(bmap))
    out = pd.DataFrame({
        "تعداد شعب": df60.groupby(SUP)[BR].nunique(),
        "راکد ۶۰ روزه": df60.groupby(SUP)[VAL].sum(),
        "راکد ۴۵ روزه": d45.groupby("_s")[VAL].sum(),
    }).fillna(0)
    out["اختلاف ۴۵ و ۶۰"] = out["راکد ۴۵ روزه"] - out["راکد ۶۰ روزه"]
    out["فروش (تعداد)"] = 0.0
    out["فروش (ریال)"] = 0.0
    if sales_df is not None and not sales_df.empty:
        matched = attach_sales(df60, sales_df)
        s = matched.groupby(SUP)[[S_QTY, S_AMT]].sum()
        out["فروش (تعداد)"] = s[S_QTY].reindex(out.index).fillna(0)
        out["فروش (ریال)"] = s[S_AMT].reindex(out.index).fillna(0)
    out.index.name = SUP
    out = out.reset_index()
    out = _finish_summary(out, ["تعداد شعب", "راکد ۶۰ روزه", "راکد ۴۵ روزه",
                                "اختلاف ۴۵ و ۶۰", "فروش (تعداد)", "فروش (ریال)"])
    return out.sort_values("راکد ۶۰ روزه", ascending=False).reset_index(drop=True)


def render_cash_and_abc(df60, sales_df=None):
    st.markdown("### 💰 پول خوابیده در انبار")
    rate = st.slider("نرخ هزینه نگهداری ماهانه (٪)", 0.0, 10.0, HOLD_RATE_DEFAULT, 0.5,
                     key="hold_rate",
                     help="هزینه سرمایه، انبارداری، فساد و ...")
    total = df60[VAL].sum()
    monthly = total * rate / 100

    if is_mobile:
        st.metric("ارزش پول خوابیده", money(total, short=True, with_unit=True))
        st.metric(f"هزینه ماهانه ({rate:g}٪)", money(monthly, short=True, with_unit=True))
        st.metric("هزینه سالانه", money(monthly * 12, short=True, with_unit=True))
        st.metric("تعداد کالاهای راکد", f"{len(df60):,}")
    else:
        metric_row([
            ("ارزش پول خوابیده", money(total)),
            (f"هزینه نگهداری ماهانه ({rate:g}٪)", money(monthly)),
            ("هزینه نگهداری سالانه", money(monthly * 12)),
            ("تعداد کالاهای راکد", f"{len(df60):,}"),
        ])

    st.divider()
    st.markdown("### ⏳ تخمین DIO")
    if sales_df is not None and not sales_df.empty:
        days = st.number_input("تعداد روزهای گزارش فروش", min_value=1, max_value=365,
                               value=30, step=1, key="dio_days",
                               help="فایل فروش مربوط به چند روز است؟ فروش روزانه = کل فروش ÷ این عدد")
        matched_sales = attach_sales(df60, sales_df)
        period_sales = matched_sales[S_AMT].sum()
        daily_sales_value = period_sales / days
        if daily_sales_value > 0:
            dio = total / daily_sales_value
            st.metric("DIO تخمینی", f"{dio:.1f} روز",
                      delta="🟢 مناسب" if dio < 90 else "🔴 بحرانی",
                      delta_color="normal" if dio < 90 else "inverse",
                      help="ارزش راکد ÷ فروش روزانه (فقط فروش همین اقلام راکد)")
            st.caption(f"💡 ارزش راکد: {money(total, short=True, with_unit=True)} | فروش روزانه: {money(daily_sales_value, short=True, with_unit=True)}")
        else:
            st.info("فروشی برای محاسبه DIO وجود ندارد.")
    else:
        st.info("برای محاسبه DIO، فایل فروش لازم است.")

    st.divider()
    st.markdown("### 🔤 تحلیل ABC")
    abc = (df60.groupby([BC, NM], as_index=False)[VAL].sum()
           .sort_values(VAL, ascending=False).reset_index(drop=True))
    total_val = abc[VAL].sum()
    if total_val <= 0:
        st.info("ارزش راکدی برای تحلیل وجود ندارد.")
        return

    share = abc[VAL] / total_val * 100
    abc["درصد تجمعی"] = share.cumsum()
    prev = abc["درصد تجمعی"] - share
    abc["دسته"] = np.select([prev < 80, prev < 95],
                            ["A — بحرانی", "B — متوسط"], default="C — کم‌ارزش")
    summary = (abc.groupby("دسته").agg(تعداد=(BC, "count"), ارزش=(VAL, "sum"))
               .reset_index().sort_values("دسته"))
    summary["برچسب"] = summary["تعداد"].apply(lambda v: f"{v} قلم")
    colors = {"A — بحرانی": CHART_PIE_RED, "B — متوسط": CHART_PIE_YELLOW,
              "C — کم‌ارزش": CHART_PIE_GREEN}

    for _, r in summary.iterrows():
        c = colors.get(r["دسته"], "#666")
        st_md(f"""
        <div class="mcard" style="border-right-color:{c};">
            <div class="mcard-badge" style="background:{c};">{r["دسته"]}</div>
            <div class="mcard-row"><span>تعداد</span><b>{r["تعداد"]:,} قلم</b></div>
            <div class="mcard-row"><span>ارزش</span><b>{money(r["ارزش"], short=True, with_unit=True)}</b></div>
        </div>
        """)

    # text از ستون px داده می‌شود تا برای هر دسته (هر trace) برچسب درست بیفتد
    fig = px.bar(summary, x="دسته", y="ارزش", color="دسته", color_discrete_map=colors,
                 text="برچسب", title="توزیع ارزش راکد بر اساس ABC")
    fig.update_layout(showlegend=False, height=320 if is_mobile else 350)
    fig = plotly_style(fig)
    fig.update_traces(textfont=dict(color=CHART_TEXT, size=10), textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📋 جدول کامل ABC"):
        if is_mobile:
            render_mobile_cards(abc, max_rows=30, title_col=NM,
                                sub_col=None, priority_cols=[BC, VAL])
        else:
            st.dataframe(abc, use_container_width=True, hide_index=True)
        st.download_button("⬇️ دانلود ABC", abc.to_csv(index=False).encode("utf-8-sig"),
                           "abc.csv", "text/csv", key="dl_abc")


def render_charts(df60, df_target):
    st.markdown("### 📊 نمودارهای تحلیلی")
    branch_val = df60.groupby(BR)[VAL].sum().reset_index().sort_values(VAL, ascending=False)
    top10 = branch_val.head(10).sort_values(VAL)
    if top10.empty:
        st.info("داده‌ای برای نمودار وجود ندارد.")
        return

    fig = px.bar(top10, x=VAL, y=BR, orientation="h",
                 title="🚨 ۱۰ شعبه با بیشترین ارزش راکد",
                 color=VAL, color_continuous_scale=CHART_GRADIENT)
    fig.update_layout(height=400 if is_mobile else 430, showlegend=False,
                      coloraxis_showscale=False)
    fig = plotly_style(fig)
    fig.update_traces(textfont=dict(color=CHART_TEXT, size=9), textposition="outside",
                      text=top10[VAL].apply(lambda v: money(v, short=True, with_unit=True)))
    st.plotly_chart(fig, use_container_width=True)

    if df_target is not None and not df_target.empty:
        tc = detect_target_columns(df_target)
        ach_col, branch_col = tc.get("achievement"), tc.get("branch")
        if ach_col and branch_col:
            def status(v):
                if v >= ACH_OK: return "🟢 موفق"
                if v >= ACH_WARN: return "🟡 در حال پیشرفت"
                return "🔴 بحرانی"

            tmp = df_target[[branch_col, ach_col]].copy()
            tmp["وضعیت"] = tmp[ach_col].apply(status)
            dist = tmp["وضعیت"].value_counts().reset_index()
            dist.columns = ["وضعیت", "تعداد"]
            fig2 = px.pie(dist, values="تعداد", names="وضعیت",
                          title="🥧 توزیع وضعیت شعب",
                          color="وضعیت", hole=0.45,
                          color_discrete_map={"🟢 موفق": CHART_PIE_GREEN,
                                              "🟡 در حال پیشرفت": CHART_PIE_YELLOW,
                                              "🔴 بحرانی": CHART_PIE_RED})
            fig2.update_layout(height=350 if is_mobile else 400)
            fig2 = plotly_style(fig2)
            fig2.update_traces(textfont=dict(color="#ffffff", size=11),
                               textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig2, use_container_width=True)


def render_history_trend():
    st.markdown("### 🕰️ روند در آپدیت‌ها")
    try:
        with _connect() as conn:
            r = pd.read_sql("SELECT date, SUM(value) AS total FROM raaked_history GROUP BY date", conn)
            s = pd.read_sql("SELECT date, SUM(qty) AS qty, SUM(amount) AS amount "
                            "FROM sales_history GROUP BY date", conn)
    except Exception:
        st.info("تاریخچه‌ای موجود نیست.")
        return
    if r.empty:
        st.info("هنوز تاریخچه‌ای ذخیره نشده.")
        return
    parts = r["date"].str.rsplit("_", n=1, expand=True)
    r["تاریخ"], r["نوع"] = parts[0], parts[1].map({"60": "۶۰ روزه", "45": "۴۵ روزه"})
    r = r.sort_values("تاریخ")
    fig = px.line(r, x="تاریخ", y="total", color="نوع", markers=True,
                  title="روند ارزش راکد",
                  color_discrete_sequence=[CHART_MAIN, CHART_NEUTRAL])
    fig.update_layout(height=320 if is_mobile else 380)
    st.plotly_chart(plotly_style(fig, show_legend_bg=True), use_container_width=True)
    if not s.empty:
        s = s.sort_values("date")
        fig2 = px.bar(s, x="date", y="amount", title="فروش اقلام راکد در هر آپدیت",
                      labels={"date": "تاریخ", "amount": "فروش"})
        fig2.update_traces(marker_color=CHART_MAIN)
        fig2.update_layout(height=300 if is_mobile else 340)
        st.plotly_chart(plotly_style(fig2), use_container_width=True)


def render_comparison():
    st.markdown("### 📈 مقایسه با آپدیت قبلی")
    dates = get_history_dates("raaked")
    if len(dates) < 2:
        st.info("برای مقایسه، حداقل ۲ بار آپدیت لازم است.")
        if len(dates) == 1:
            st.caption(f"📌 آخرین ذخیره: {dates[0]}")
        return

    kind = st.radio("نوع", ["60", "45"], horizontal=True, key="cmp_kind",
                    format_func=lambda k: "۶۰ روزه" if k == "60" else "۴۵ روزه")

    if is_mobile:
        date1 = st.selectbox("از تاریخ", dates[1:], index=0, key="cmp_date1")
        date2 = st.selectbox("به تاریخ", dates, index=0, key="cmp_date2")
    else:
        cc = st.columns([1, 1, 2])
        date1 = cc[0].selectbox("از تاریخ", dates[1:], index=0, key="cmp_date1")
        date2 = cc[1].selectbox("به تاریخ", dates, index=0, key="cmp_date2")

    if date1 == date2:
        st.warning("دو تاریخ یکسان انتخاب شده.")
        return
    if date1 > date2:
        st.warning("تاریخ «از» باید قبل از تاریخ «به» باشد.")
        return

    cmp_df = get_history_comparison(date1, date2, kind)
    if cmp_df is None or cmp_df.empty:
        st.info("داده‌ای برای مقایسه وجود ندارد.")
        return

    before, now = cmp_df["total_قبل"].sum(), cmp_df["total_الان"].sum()
    diff = now - before
    pct = (diff / before * 100) if before > 0 else 0

    if is_mobile:
        st.metric("ارزش قبلی", money(before, short=True, with_unit=True))
        st.metric("ارزش فعلی", money(now, short=True, with_unit=True),
                  delta=f"{diff:+,.0f} ریال", delta_color="inverse")
        st.metric("درصد تغییر", f"{pct:+.2f}%")
        st.metric("تعداد شعب", f"{len(cmp_df):,}")
    else:
        metric_row([
            ("ارزش قبلی", money(before)),
            ("ارزش فعلی", money(now), {"delta": f"{diff:+,.0f} ریال", "delta_color": "inverse"}),
            ("درصد تغییر", f"{pct:+.2f}%"),
            ("تعداد شعب", f"{len(cmp_df):,}"),
        ])

    st.markdown("#### 🎯 بیشترین افزایش و کاهش")
    gainers = (cmp_df[cmp_df["تغییر"] > 0]
               .sort_values("تغییر", ascending=False).head(5).copy())
    losers = (cmp_df[cmp_df["تغییر"] < 0]
              .sort_values("تغییر", ascending=True).head(5).copy())

    if is_mobile:
        st.markdown("**🔴 بیشترین افزایش راکد**")
        if gainers.empty:
            st.caption("شعبه‌ای با افزایش راکد نبود.")
        for _, r in gainers.iterrows():
            st_md(f"""
            <div class="mcard" style="border-right-color: {OK_RED};">
                <div class="mcard-badge" style="background:{OK_RED};">🔴 افزایش</div>
                <div class="mcard-title">🏪 {esc(r['branch'])}</div>
                <div class="mcard-row"><span>افزایش</span><b>{money(r['تغییر'], short=True, with_unit=True)}</b></div>
                <div class="mcard-row"><span>درصد</span><b>{r['درصد تغییر']:+.1f}%</b></div>
            </div>
            """)
        st.markdown("**🟢 بیشترین کاهش راکد**")
        if losers.empty:
            st.caption("شعبه‌ای با کاهش راکد نبود.")
        for _, r in losers.iterrows():
            st_md(f"""
            <div class="mcard" style="border-right-color: {CHART_PIE_GREEN};">
                <div class="mcard-badge" style="background:{CHART_PIE_GREEN};">🟢 کاهش</div>
                <div class="mcard-title">🏪 {esc(r['branch'])}</div>
                <div class="mcard-row"><span>کاهش</span><b>{money(r['تغییر'], short=True, with_unit=True)}</b></div>
                <div class="mcard-row"><span>درصد</span><b>{r['درصد تغییر']:+.1f}%</b></div>
            </div>
            """)
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**🔴 بیشترین افزایش راکد**")
            st.dataframe(gainers.rename(columns={"branch": "شعبه", "تغییر": "افزایش",
                                                  "درصد تغییر": "درصد"}),
                         use_container_width=True, hide_index=True)
        with c2:
            st.markdown("**🟢 بیشترین کاهش راکد**")
            st.dataframe(losers.rename(columns={"branch": "شعبه", "تغییر": "کاهش",
                                                 "درصد تغییر": "درصد"}),
                         use_container_width=True, hide_index=True)

    top = cmp_df.loc[cmp_df["تغییر"].abs().sort_values(ascending=False).index].head(15)
    fig = go.Figure()
    for name, col, color in (("قبل", "total_قبل", CHART_NEUTRAL),
                              ("فعلی", "total_الان", CHART_MAIN)):
        fig.add_trace(go.Bar(name=name, x=top["branch"], y=top[col], marker_color=color,
                             text=top[col].apply(lambda v: money(v, short=True, with_unit=True)),
                             textfont=dict(color=CHART_TEXT, size=9),
                             textposition="outside"))
    fig.update_layout(barmode="group", height=420 if is_mobile else 470,
                      xaxis_tickangle=-45, title="مقایسه ارزش راکد")
    st.plotly_chart(plotly_style(fig, show_legend_bg=True), use_container_width=True)

    with st.expander("📋 جدول کامل مقایسه"):
        disp = cmp_df.rename(columns={"branch": "شعبه", "total_قبل": "قبل",
                                       "total_الان": "الان"})
        if is_mobile:
            render_mobile_cards(disp, max_rows=30, title_col="شعبه",
                                sub_col=None, priority_cols=["قبل", "الان", "تغییر", "درصد تغییر"])
        else:
            st.dataframe(disp, use_container_width=True, hide_index=True)


# ================== چک‌لیست شیفت ==================
def render_shift_checklist(pq45, pq60, key_suffix=""):
    st.markdown("### 📋 چک‌لیست پیگیری شیفت")
    st.caption("۱۰ قلم راکد با بالاترین ارزش ریالی هر شعبه — مخصوص پیگیری سرپرست")
    if (pq45 is None or pq45.empty) and (pq60 is None or pq60.empty):
        st.info("داده‌های شیفت (pq45 / pq60) در فایل موجود نیست.")
        return

    with st.expander("🎛 فیلترها", expanded=False):
        sup_list = []
        for df in (pq45, pq60):
            if df is not None and not df.empty and SUP in df.columns:
                sup_list += df[SUP].dropna().unique().tolist()
        sup_list = sorted(set(sup_list))
        selected_sup = st.selectbox("سرپرست", ["همه"] + sup_list,
                                    key=f"sh_sup_{key_suffix}")
        search_q = st.text_input("🔍 جستجو (نام کالا / بارکد / شعبه)",
                                 key=f"sh_search_{key_suffix}",
                                 placeholder="مثلاً: برنج یا 6260...")

    def apply_filters(df):
        if df is None or df.empty:
            return df
        x = df.copy()
        if SUP in x.columns:
            x[SUP] = normalize_name(x[SUP])
        if selected_sup != "همه" and SUP in x.columns:
            sel_norm = normalize_name(pd.Series([selected_sup])).iloc[0]
            x = x[x[SUP] == sel_norm]
        if search_q and search_q.strip():
            terms = [normalize_query(t) for t in search_q.split() if t.strip()]
            mask = pd.Series(True, index=x.index)
            for term in terms:
                m = pd.Series(False, index=x.index)
                for c in x.columns:
                    txt = (x[c].astype(str).str.lower()
                           .str.translate(FA_DIGITS)
                           .str.replace("ي", "ی", regex=False)
                           .str.replace("ك", "ک", regex=False))
                    m |= txt.str.contains(term, na=False, regex=False)
                mask &= m
            x = x[mask]
        return x.reset_index(drop=True)

    f45 = apply_filters(pq45)
    f60 = apply_filters(pq60)
    n45 = 0 if f45 is None else len(f45)
    n60 = 0 if f60 is None else len(f60)

    active_shift = collapse_group(
        [f"🌅 صبح ({n45})", f"🌙 عصر ({n60})"],
        "shift_checklist", key_suffix
    )

    for idx_sh, df, shift_name, k in ((0, f45, "صبح", "morning"),
                                        (1, f60, "عصر", "evening")):
        if active_shift != idx_sh:
            continue
        if True:
            if df is None or df.empty:
                st.info(f"داده‌ای برای شیفت {shift_name} وجود ندارد.")
                continue

            kpi_cols = st.columns(3)
            kpi_cols[0].metric("تعداد اقلام", f"{len(df):,}")
            kpi_cols[1].metric("ارزش راکد",
                                money(df[VAL].sum(), short=is_mobile, with_unit=is_mobile))
            kpi_cols[2].metric("تعداد شعبه", f"{df[BR].nunique():,}")
            st.divider()

            if is_mobile:
                render_mobile_cards(df, max_rows=100,
                                    title_col=NM, sub_col=BR,
                                    priority_cols=[SUP, BC, QTY, VAL])
            else:
                view = df[[c for c in [BR, BCODE, BC, NM, QTY, VAL, SUP]
                           if c in df.columns]]
                st.dataframe(view, use_container_width=True, hide_index=True,
                             column_config=column_config(view))

            st.download_button(
                f"⬇️ دانلود چک‌لیست {shift_name}",
                df.to_csv(index=False).encode("utf-8-sig"),
                f"shift_{k}_{key_suffix}.csv", "text/csv",
                key=f"dl_shift_{k}_{key_suffix}"
            )



# ================== روند و مقایسه ==================
def render_trend_comparison(summary60, key_suffix=""):
    st.markdown("### 📈 روند و مقایسه شعب")
    if summary60 is None or summary60.empty:
        st.info("شیت «۶۰ روزه» در فایل موجود نیست یا خالی است.")
        return
    df = summary60.copy()
    has_old = S60_PCT_OLD in df.columns
    has_new = S60_PCT_NEW in df.columns
    has_diff = S60_DIFF_PCT in df.columns
    if not (has_old and has_new):
        st.warning("ستون‌های «درصد راکد <عدد>» (دو ستون یا بیشتر) در شیت خلاصه پیدا نشد.")
        return
    mean_old = df[S60_PCT_OLD].mean()
    mean_new = df[S60_PCT_NEW].mean()
    diff_mean = mean_new - mean_old
    improved = int((df[S60_PCT_NEW] < df[S60_PCT_OLD]).sum())
    worsened = int((df[S60_PCT_NEW] > df[S60_PCT_OLD]).sum())
    unchanged = len(df) - improved - worsened
    if is_mobile:
        st.metric("میانگین درصد قبل", f"{mean_old:.2f}%")
        st.metric("میانگین درصد فعلی", f"{mean_new:.2f}%",
                  delta=f"{diff_mean:+.2f}%",
                  delta_color="inverse" if diff_mean > 0 else "normal")
        st.metric("🟢 بهبود", f"{improved:,} شعبه")
        st.metric("🔴 افت", f"{worsened:,} شعبه")
        st.metric("⚪ بدون تغییر", f"{unchanged:,} شعبه")
    else:
        metric_row([
            ("میانگین قبل", f"{mean_old:.2f}%"),
            ("میانگین فعلی", f"{mean_new:.2f}%",
             {"delta": f"{diff_mean:+.2f}%",
              "delta_color": "inverse" if diff_mean > 0 else "normal"}),
            ("🟢 بهبود", f"{improved:,}"),
            ("🔴 افت", f"{worsened:,}"),
            ("⚪ بدون تغییر", f"{unchanged:,}"),
        ])
    st.divider()
    st.markdown("#### 📊 مقایسه درصد راکد قبل و بعد")
    chart_df = df.dropna(subset=[S60_PCT_OLD, S60_PCT_NEW]).copy()
    chart_df["_change"] = chart_df[S60_PCT_NEW] - chart_df[S60_PCT_OLD]
    chart_df = chart_df.reindex(
        chart_df["_change"].abs().sort_values(ascending=False).index
    ).head(15)
    if not chart_df.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="قبل", x=chart_df[BR], y=chart_df[S60_PCT_OLD],
            marker_color=CHART_NEUTRAL,
            text=chart_df[S60_PCT_OLD].apply(lambda v: f"{v:.2f}%"),
            textposition="outside",
            textfont=dict(color=CHART_TEXT, size=9)
        ))
        fig.add_trace(go.Bar(
            name="فعلی", x=chart_df[BR], y=chart_df[S60_PCT_NEW],
            marker_color=CHART_MAIN,
            text=chart_df[S60_PCT_NEW].apply(lambda v: f"{v:.2f}%"),
            textposition="outside",
            textfont=dict(color=CHART_TEXT, size=9)
        ))
        fig.update_layout(
            barmode="group",
            height=450 if is_mobile else 500,
            xaxis_tickangle=-45,
            title="۱۵ شعبه با بیشترین تغییر"
        )
        st.plotly_chart(plotly_style(fig, show_legend_bg=True),
                        use_container_width=True)
    st.divider()
    st.markdown("#### 🏆 رتبه‌بندی افت / رشد")
    if has_diff:
        df = df.sort_values(S60_DIFF_PCT, ascending=True).reset_index(drop=True)
    else:
        df["_diff"] = df[S60_PCT_NEW] - df[S60_PCT_OLD]
        df = df.sort_values("_diff", ascending=True).reset_index(drop=True)
        df[S60_DIFF_PCT] = df["_diff"]
        df = df.drop(columns=["_diff"])
    q = st.text_input("🔍 جستجو", key=f"trend_search_{key_suffix}",
                      placeholder="نام شعبه یا سرپرست")
    view = df.copy()
    if q and q.strip():
        qn = normalize_query(q)
        mask = view.astype(str).apply(
            lambda col: (col.str.lower().str.translate(FA_DIGITS)
                         .str.replace("ي", "ی", regex=False)
                         .str.replace("ك", "ک", regex=False)
                         .str.contains(qn, na=False, regex=False))
        ).any(axis=1)
        view = view[mask]
        st.caption(f"🔸 {len(view):,} ردیف")
    if is_mobile:
        for _, r in view.head(50).iterrows():
            diff = r.get(S60_DIFF_PCT, 0)
            if pd.isna(diff):
                diff = 0
            if diff < 0:
                bg, icon, label = CHART_PIE_GREEN, "🟢", "بهبود"
            elif diff > 0:
                bg, icon, label = OK_RED, "🔴", "افت"
            else:
                bg, icon, label = CHART_NEUTRAL, "⚪", "بدون تغییر"
            old_v = r.get(S60_PCT_OLD, 0)
            new_v = r.get(S60_PCT_NEW, 0)
            old_v = 0 if pd.isna(old_v) else old_v
            new_v = 0 if pd.isna(new_v) else new_v
            st_md(f"""
            <div class="mcard" style="border-right-color: {bg};">
                <div class="mcard-badge" style="background:{bg};">{icon} {label} {diff:+.2f}%</div>
                <div class="mcard-title">🏪 {esc(r.get(BR, ''))}</div>
                <div class="mcard-sup">👤 {esc(r.get(SUP, '—'))}</div>
                <div class="mcard-row"><span>قبل</span><b>{old_v:.2f}%</b></div>
                <div class="mcard-row"><span>فعلی</span><b>{new_v:.2f}%</b></div>
                <div class="mcard-row"><span>ارزش راکد</span><b>{money(r.get(VAL, 0), short=True)}</b></div>
            </div>
            """)
    else:
        display_cols = [c for c in [BR, BCODE, SUP, VAL,
                                     S60_PCT_OLD, S60_PCT_NEW,
                                     S60_DIFF_PCT, S60_DIFF_TXT]
                        if c in view.columns]
        st.dataframe(view[display_cols], use_container_width=True,
                     hide_index=True)
    st.download_button(
        "⬇️ دانلود روند",
        view.to_csv(index=False).encode("utf-8-sig"),
        f"trend_{key_suffix}.csv", "text/csv",
        key=f"dl_trend_{key_suffix}"
    )


# ================== ارسال ایمیل ==================
def render_email_sender():
    """بخش ارسال گزارش مدیریتی از داخل داشبورد"""
    import configparser
    import smtplib
    import ssl
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    st.markdown("### 📧 ارسال گزارش مدیریتی به مدیران")
    st.caption("گزارش تحقق شعب و عملکرد سرپرست‌ها را به ایمیل‌های مشخص ارسال کنید")

    cfg = configparser.ConfigParser()
    cfg_path = Path("config.ini")
    default_recipients = ""
    default_subject = "گزارش مدیریتی هفتگی — کالای راکد افق کوروش"
    if cfg_path.exists():
        try:
            cfg.read(cfg_path, encoding="utf-8")
            default_recipients = cfg["REPORT"].get("recipients", "")
            default_subject = cfg["REPORT"].get("subject", default_subject)
        except Exception:
            pass

    if "email_recipients_list" not in st.session_state:
        st.session_state["email_recipients_list"] = [
            e.strip() for e in default_recipients.split(",") if e.strip()
        ]

    col1, col2 = st.columns([1, 1])
    with col1:
        subject_input = st.text_input(
            "📝 موضوع ایمیل",
            value=default_subject,
            key="email_subject_input",
        )
    with col2:
        recipients_text = st.text_area(
            "📬 ایمیل گیرندگان (با کاما جدا کنید)",
            value=", ".join(st.session_state["email_recipients_list"]),
            height=80,
            key="email_recipients_text",
        )

    recipients = [r.strip() for r in recipients_text.split(",") if r.strip() and "@" in r]

    if recipients:
        st.caption(f"📧 {len(recipients)} گیرنده: {', '.join(recipients)}")
    else:
        st.caption("⚠️ هنوز ایمیلی وارد نشده")

    st.divider()

    col_a, col_b = st.columns([1, 1])
    with col_a:
        preview_btn = st.button("👁 پیش‌نمایش گزارش", use_container_width=True, key="preview_email_btn")
    with col_b:
        send_btn = st.button("📧 ارسال ایمیل", type="primary", use_container_width=True, key="send_email_btn")

    if preview_btn:
        with st.spinner("در حال ساخت پیش‌نمایش..."):
            try:
                from send_report import build_report
                html, stats = build_report()
                st.success("✅ پیش‌نمایش آماده شد")
                st.components.v1.html(html, height=700, scrolling=True)
            except Exception as e:
                st.error(f"❌ خطا در ساخت پیش‌نمایش: {e}")

    if send_btn:
        if not recipients:
            st.error("❌ حداقل یک ایمیل معتبر وارد کنید")
        else:
            with st.spinner("در حال ساخت و ارسال گزارش..."):
                try:
                    from send_report import build_report, send_email
                    html, stats = build_report()
                    sent = send_email(html, subject=subject_input)
                    try:
                        if not cfg.has_section("REPORT"):
                            cfg.add_section("REPORT")
                        cfg["REPORT"]["recipients"] = ",".join(recipients)
                        cfg["REPORT"]["subject"] = subject_input
                        with open(cfg_path, "w", encoding="utf-8") as f:
                            cfg.write(f)
                    except Exception:
                        pass
                    st.success(f"✅ ایمیل با موفقیت به {len(sent)} گیرنده ارسال شد!")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ خطا در ارسال ایمیل: {e}")
                    st.caption("💡 مطمئن شو send_report.py و config.ini در همان پوشه هستند")

    st.divider()
    st.markdown("#### 📋 مدیریت لیست ایمیل‌ها")

    col_x, col_y = st.columns([4, 1])
    with col_x:
        new_email = st.text_input(
            "افزودن ایمیل به لیست",
            key="add_email_input",
            placeholder="manager@okco.ir",
        )
    with col_y:
        st.write("")
        st.write("")
        if st.button("➕ افزودن", key="add_email_btn"):
            if new_email and "@" in new_email:
                if new_email not in st.session_state["email_recipients_list"]:
                    st.session_state["email_recipients_list"].append(new_email)
                    st.rerun()
                else:
                    st.warning("این ایمیل قبلاً اضافه شده")
            else:
                st.error("ایمیل نامعتبر")

    if st.session_state["email_recipients_list"]:
        st.markdown("**ایمیل‌های ذخیره‌شده:**")
        for i, em in enumerate(st.session_state["email_recipients_list"]):
            c1, c2 = st.columns([6, 1])
            c1.write(f"📧 {em}")
            if c2.button("🗑", key=f"del_email_{i}"):
                st.session_state["email_recipients_list"].pop(i)
                st.rerun()


# ================== آپلود ==================
def get_upload_password():
    try:
        pw = st.secrets.get("UPLOAD_PASSWORD")
    except Exception:
        pw = None
    if pw:
        return pw
    pw = os.environ.get("UPLOAD_PASSWORD")
    if pw:
        return pw
    try:
        import configparser
        _cfg = configparser.ConfigParser()
        _cfg.read(Path(__file__).parent / "config.ini", encoding="utf-8")
        return (_cfg.get("ADMIN", "upload_password", fallback="")
                or _cfg.get("ADMIN", "password", fallback="") or None)
    except Exception:
        return None


def check_upload_access():
    if _passwords_disabled():
        return True
    pw = get_upload_password()
    if not pw:
        # بدون رمز، هر کسی که لینک را داشته باشد می‌تواند داده را جایگزین کند؛ پس دسترسی بسته می‌ماند
        st.error("🔒 رمز آپلود تنظیم نشده است. برای فعال‌سازی آپلود، "
                 "`UPLOAD_PASSWORD` را در Secrets یا متغیر محیطی تنظیم کنید.")
        return False
    if st.session_state.get("upload_ok"):
        return True
    entered = st.text_input("🔒 رمز آپلود", type="password", key="upload_pw")
    if entered:
        if hmac.compare_digest(entered.encode(), str(pw).encode()):
            st.session_state["upload_ok"] = True
            st.rerun()
        else:
            st.error("رمز اشتباه است.")
    return False


def render_upload():
    if not check_upload_access():
        return
    flash = st.session_state.pop("flash", None)
    flash_issues = st.session_state.pop("flash_issues", None)
    if flash:
        st.success(flash)
    if flash_issues:
        with st.expander("🔍 گزارش کیفیت داده", expanded=True):
            for it in flash_issues:
                st.markdown(it)

    st.markdown("### 📤 آپلود فایل‌های جدید")
    st.warning("⚠️ **قبل از آپلود**، در اکسل `Ctrl+Alt+F5` بزنید (Refresh All) "
               "و فایل را ذخیره کنید. در غیر این‌صورت داده قدیمی آپلود می‌شود.")
    st.info("فایل قبل از جایگزینی اعتبارسنجی می‌شود و از نسخه قبلی بک‌آپ گرفته می‌شود.")

    picked = st.date_input("📅 تاریخ این آپدیت",
                           value=now_tehran().date(), key="upload_date_pick")
    upd_date = picked.strftime("%Y-%m-%d")
    try:
        import jdatetime
        st.caption(f"🗓️ شمسی: {jdatetime.date.fromgregorian(date=picked).strftime('%Y/%m/%d')}")
    except ImportError:
        st.caption(f"📅 {upd_date}")

    note = st.text_input("📝 یادداشت", placeholder="مثلاً: آپدیت هفتگی", key="upload_note")
    st.divider()

    if is_mobile:
        col1 = st.container()
        col2 = st.container()
    else:
        col1, col2 = st.columns(2)

    jobs = [
        (col1, "📊 فایل راکد", "فایل راکد جدید", "up_raaked", "save_raaked",
         "💾 ذخیره راکد", process_raaked_upload, "فایل راکد"),
        (col2, "🛒 فایل فروش", "فایل فروش جدید", "up_sales", "save_sales",
         "💾 ذخیره فروش", process_sales_upload, "فایل فروش")
    ]
    result = None
    issues = None
    for col, title, label, fkey, bkey, btn, fn, name in jobs:
        with col:
            st.markdown(f"**{title}**")
            f = st.file_uploader(label, type=["xlsx"], key=fkey)
            if f and st.button(btn, use_container_width=True, key=bkey, type="primary"):
                try:
                    rows, dq = fn(f.getvalue(), upd_date, note)
                    result = f"✅ {name} ذخیره شد ({rows:,} ردیف)"
                    if dq:
                        issues = dq
                except Exception as e:
                    st.error(f"❌ خطا: {e}")
                    st.caption("💡 فایل قبلی دست‌نخورده ماند.")

    if result:
        st.cache_data.clear()
        st.session_state["flash"] = result
        if issues:
            st.session_state["flash_issues"] = issues
        st.rerun()

    st.divider()
    st.markdown("### 🕐 آخرین آپدیت‌ها")
    try:
        with _connect() as conn:
            updates = pd.read_sql(
                "SELECT date, kind, rows FROM updates ORDER BY id DESC LIMIT 15", conn)
        if updates.empty:
            st.info("هنوز آپدیتی ذخیره نشده.")
        else:
            updates = updates.rename(columns={"date": "تاریخ", "kind": "نوع", "rows": "ردیف"})
            updates["نوع"] = updates["نوع"].replace({"raaked": "📊 راکد", "sales": "🛒 فروش"})
            st.dataframe(updates, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"خطا: {e}")

    if HISTORY_DB.exists():
        st.download_button("⬇️ دانلود بک‌آپ دیتابیس",
                           HISTORY_DB.read_bytes(),
                           f"history_{now_tehran():%Y%m%d}.db",
                           "application/octet-stream", key="dl_db")

    # ══════════════════════════════════════════════════
    # آپلود فایل پرسنلی
    # ══════════════════════════════════════════════════
    st.divider()
    st.markdown("### 👥 آپلود فایل پرسنلی")
    st.caption("فایل اکسل با ستون‌های «کد پرسنلي، Location، سریال سمت، سریال نام واحد»")

    # چک وجود فایل
    try:
        from personel_loader import load_personel, _find_personel_file, PERSONEL_FILE
        _existing = _find_personel_file()
        if _existing.exists():
            _count = len(load_personel())
            st.success(f"✅ فایل پرسنلی موجود — {_count} کاربر لود شده")
            st.caption(f"📁 {_existing}")
        else:
            st.warning("⚠️ فایل پرسنلی آپلود نشده — کاربران با نقش «other» وارد می‌شن")
    except Exception as e:
        st.error(f"❌ خطا در بارگذاری: {e}")

    # آپلودر
    personel_file = st.file_uploader(
        "📁 فایل پرسنلی (xlsx)",
        type=["xlsx"],
        key="upload_personel_file",
    )

    if personel_file is not None:
        if st.button("💾 ذخیره فایل پرسنلی",
                     use_container_width=True,
                     key="save_personel_btn",
                     type="primary"):
            try:
                # ذخیره در DATA_DIR
                from personel_loader import _CANDIDATE_NAMES
                target = DATA_DIR / _CANDIDATE_NAMES[0]

                # بک‌آپ از فایل قبلی
                if target.exists():
                    _backup(target)

                # نوشتن فایل جدید
                _atomic_write(target, personel_file.getvalue())

                # کش رو پاک کن
                st.cache_data.clear()

                st.success(f"✅ فایل پرسنلی ذخیره شد ({len(personel_file.getvalue())/1024:.0f} KB)")
                st.info("🔄 حالا از منوی بالا یک بار خروج بزن و دوباره با کد پرسنلی وارد شو.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ خطا: {e}")

    # پیش‌نمایش کاربران
    try:
        from personel_loader import load_personel
        _personel = load_personel()
        if _personel:
            with st.expander(f"👥 لیست کاربران ({len(_personel)})", expanded=False):
                # تبدیل به DataFrame
                _rows = []
                for code, u in _personel.items():
                    _rows.append({
                        "کد پرسنلی": code,
                        "نام": u.get("name", ""),
                        "نقش": u.get("role", ""),
                        "کد شعبه": u.get("branch_code", ""),
                        "نام واحد": u.get("branch_name", ""),
                    })
                _p_df = pd.DataFrame(_rows)
                # توزیع نقش‌ها
                _role_counts = _p_df["نقش"].value_counts().to_dict()
                st.markdown("**توزیع نقش‌ها:**")
                for r, c in _role_counts.items():
                    st.markdown(f"- `{r}`: **{c}** نفر")
                st.dataframe(_p_df.head(30), use_container_width=True, hide_index=True)
                st.caption(f"👆 نمایش ۳۰ ردیف اول از {len(_p_df)}")
    except Exception:
        pass
    

# ================== تارگت ==================
def _kpi_cards(items):
    """items: [(label, value, tone)] — tone: green / yellow / red / ''"""
    cards = ""
    for it in items:
        label, value = it[0], it[1]
        tone = it[2] if len(it) > 2 else ""
        cards += (f'<div class="kpi-card {tone}"><div class="kpi-label">{esc(label)}</div>'
                  f'<div class="kpi-value">{esc(value)}</div></div>')
    st_md(f'<div class="kpi-grid">{cards}</div>')


def render_target_kpis(t_df):
    if t_df is None or t_df.empty:
        st.info("فایل تارگت بارگذاری نشده است.")
        return
    ach_col = detect_target_columns(t_df).get("achievement")
    if not ach_col:
        return
    vals = t_df[ach_col].dropna()
    vals = vals[vals != 0]
    if len(vals) == 0:
        return
    _m = vals.mean()
    _tone = "green" if _m >= ACH_OK else ("yellow" if _m >= ACH_WARN else "red")
    _kpi_cards([
        ("میانگین تحقق", f"{_m:.1f}%", _tone),
        ("موفق (≥100%)", f"{(vals >= ACH_OK).sum():,}", "green"),
        ("بحرانی (<80%)", f"{(vals < ACH_WARN).sum():,}", "red"),
        ("کل شعب", f"{len(t_df):,}", ""),
    ])


def render_target_table(t_df, key_suffix=""):
    cm = detect_target_columns(t_df)
    display_cols = [cm[k] for k in ("branch", "supervisor", "zone_supervisor",
                                     "raked_value", "target", "achievement", "change")
                    if cm.get(k)]
    if not display_cols:
        st.warning("ستون‌های تارگت قابل شناسایی نیست.")
        return
    view = t_df[display_cols].copy()
    if cm.get("achievement"):
        view = view.sort_values(cm["achievement"], ascending=False).reset_index(drop=True)
        view.insert(0, "رتبه", range(1, len(view) + 1))

    st.markdown("**📋 جدول تارگت و تحقق**")
    q = st.text_input("🔍 جستجو", key=f"search_target_{key_suffix}",
                      placeholder="نام شعبه یا سرپرست")
    if q and q.strip():
        qn = normalize_query(q)
        mask = view.astype(str).apply(
            lambda col: (col.str.lower().str.translate(FA_DIGITS)
                         .str.replace("ي", "ی", regex=False)
                         .str.replace("ك", "ک", regex=False)
                         .str.contains(qn, na=False, regex=False))).any(axis=1)
        view = view[mask]
        st.caption(f"🔸 {len(view):,} ردیف")

    if view.empty:
        st.info("ردیفی پیدا نشد.")
        return

    if is_mobile:
        branch_c = cm.get("branch")
        sup_c = cm.get("supervisor")
        priority = [c for c in [sup_c] if c and c in view.columns]
        priority += [c for c in view.columns if "تارگت" in c or "تحقق" in c or "تغییرات" in c]
        render_mobile_cards(view, max_rows=30,
                            title_col=branch_c if branch_c in view.columns else display_cols[0],
                            sub_col=sup_c if sup_c in view.columns else None,
                            priority_cols=priority)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True,
                     column_config=target_column_config(view))

    st.download_button("⬇️ دانلود CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       f"target_{key_suffix}.csv", "text/csv",
                       key=f"dl_target_{key_suffix}")


def render_target_trend(t_df, key_suffix=""):
    cm = detect_target_columns(t_df)
    trend_cols, branch_col = cm.get("trend") or [], cm.get("branch")
    if not trend_cols or not branch_col:
        st.info("ستون‌های روند پیدا نشدند.")
        return

    # ترتیب ماه‌های شمسی
    MONTH_ORDER = {"فروردین": 1, "اردیبهشت": 2, "خرداد": 3, "تیر": 4,
                   "مرداد": 5, "شهریور": 6, "مهر": 7, "آبان": 8,
                   "آذر": 9, "دی": 10, "بهمن": 11, "اسفند": 12}

    def _date_key(c):
        """مرتب‌سازی بر اساس (ماه, روز)"""
        c_clean = str(c).translate(FA_DIGITS)
        # پیدا کردن روز و ماه
        m = re.search(r"([0-9]+)\s*(شهریور|مهر|آبان|آذر|دی|بهمن|اسفند|فروردین|اردیبهشت|خرداد|تیر|مرداد)", c_clean)
        if not m:
            # اگه نبود، فقط عدد
            m2 = re.search(r"([0-9]+)", c_clean)
            return (0, int(m2.group(1))) if m2 else (99, 99)
        day = int(m.group(1))
        month = MONTH_ORDER.get(m.group(2), 99)
        return (month, day)

    trend_sorted = sorted(trend_cols, key=_date_key)
    labels = [str(c).replace("درصد راکد", "").replace("درصد  راکد", "").strip()
              for c in trend_sorted]

    df = t_df[[branch_col] + trend_sorted].copy()
    df = df.set_index(branch_col)
    for c in df.columns:
        df[c] = parse_percent(df[c])

    all_branches = df.index.tolist()
    default_sel = all_branches[:3] if len(all_branches) >= 3 else all_branches

    selected = st.multiselect(
        "🔎 انتخاب شعب",
        options=all_branches,
        default=default_sel,
        key=f"chart_branches_{key_suffix}"
    )

    if not selected:
        st.info("شعبه‌ای انتخاب نشده.")
        return

    colors = ["#E6003E", "#FF1F5A", "#f59e0b", "#10b981", "#3b82f6",
              "#8b5cf6", "#ec4899", "#06b6d4", "#f97316", "#84cc16"]

    fig = go.Figure()
    for i, bname in enumerate(selected):
        vals = df.loc[bname, trend_sorted].values
        fig.add_trace(go.Scatter(
            x=labels,
            y=vals,
            mode="lines+markers",
            name=str(bname),
            line=dict(width=2.5, color=colors[i % len(colors)]),
            marker=dict(size=9),
        ))

    # محاسبه min/max برای padding
    all_vals = []
    for bname in selected:
        all_vals.extend([v for v in df.loc[bname, trend_sorted].values if v == v])
    if all_vals:
        y_min = min(all_vals) * 0.9
        y_max = max(all_vals) * 1.1
    else:
        y_min, y_max = 0, 10

    fig.update_layout(
        title="📈 روند درصد کالای راکد",
        height=520 if is_mobile else 600,
        xaxis_title="تاریخ",
        yaxis_title="درصد راکد (%)",
        xaxis_tickangle=-40,
        yaxis=dict(range=[y_min, y_max]),
        hovermode="x unified",
    )
    fig = plotly_style(fig, show_legend_bg=True)
    fig.update_layout(
        legend=dict(orientation="h", yanchor="bottom", y=-0.7, xanchor="center", x=0.5),
        margin=dict(l=10, r=20, t=60, b=120),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📊 جدول روند"):
        disp = df.loc[selected, trend_sorted].copy()
        disp.columns = labels
        disp = disp.round(2)
        disp.index.name = "شعبه"
        st.dataframe(disp, use_container_width=True)
        st.download_button(
            "⬇️ دانلود CSV",
            disp.to_csv().encode("utf-8-sig"),
            f"trend_{key_suffix}.csv", "text/csv",
            key=f"dl_trend_{key_suffix}"
        )



def render_ranking(t_df, key_suffix="", all_df=None):
    """رتبه‌بندی — اگه تعداد کم بود، فقط یک لیست نشون بده"""
    cm = detect_target_columns(t_df)
    branch_col, sup_col, ach_col = cm.get("branch"), cm.get("supervisor"), cm.get("achievement")
    change_col, raked_col = cm.get("change"), cm.get("raked_value")
    if not branch_col or not ach_col:
        st.info("برای رتبه‌بندی، نام شعبه و درصد تحقق لازم است.")
        return

    def status_badge(v):
        if v >= ACH_OK:
            return ("موفق", CHART_PIE_GREEN, "🟢")
        if v >= ACH_WARN:
            return ("در حال پیشرفت", "#f59e0b", "🟡")
        return ("بحرانی", OK_RED, "🔴")

    # منبع رتبه: اگه all_df داده شده، از اون استفاده کن (همه شعب)
    source_df = all_df if all_df is not None and not all_df.empty else t_df

    # فیلتر ردیف‌های خالی از منبع
    _bcol = next((c for c in source_df.columns if "نام شعبه" in c), None)
    if _bcol:
        _nm = source_df[_bcol].astype(str).str.strip()
        _empty = source_df[_bcol].isna() | _nm.isin(["", "nan", "None", "NaN"])
        _total = _nm.str.contains(r"^\s*(?:جمع|مجموع|کل|total|grand|sum)",
                                   case=False, regex=True, na=False)
        source_df = source_df[~(_empty | _total)].copy().reset_index(drop=True)

    # رتبه‌بندی کامل روی منبع
    full_rank = source_df.sort_values(ach_col, ascending=False).reset_index(drop=True)
    full_rank["رتبه"] = range(1, len(full_rank) + 1)

    # اگه t_df زیرمجموعه‌ست
    if all_df is not None and not all_df.empty and len(t_df) < len(source_df):
        _rank_map = dict(zip(full_rank[branch_col], full_rank["رتبه"]))
        rank_df = t_df.copy()
        rank_df["رتبه"] = rank_df[branch_col].map(_rank_map).fillna(0).astype(int)
        rank_df["وضعیت"] = rank_df[ach_col].apply(lambda v: status_badge(v)[0])
        rank_df = rank_df.sort_values("رتبه").reset_index(drop=True)
    else:
        rank_df = full_rank[[c for c in [branch_col, sup_col, raked_col, ach_col, change_col] if c]].copy()
        rank_df["رتبه"] = full_rank["رتبه"]
        rank_df["وضعیت"] = rank_df[ach_col].apply(lambda v: status_badge(v)[0])

    total_n = len(source_df)
    n = len(rank_df)
    if n == 0:
        st.info("داده‌ای وجود ندارد.")
        return

    def show_rank(df, title):
        st.markdown(f"**{title}**")
        if is_mobile:
            for _, r in df.iterrows():
                _, bg, icon = status_badge(r[ach_col])
                sup_txt = r.get(sup_col, '—') if sup_col else '—'
                st_md(f"""
                <div class="mcard">
                    <div class="mcard-badge" style="background:{bg};">{icon} رتبه {int(r['رتبه'])} از {total_n}</div>
                    <div class="mcard-title">🏪 {esc(r[branch_col])}</div>
                    <div class="mcard-sup">👤 {esc(sup_txt)}</div>
                    <div class="mcard-row"><span>تحقق</span><b>{r[ach_col]:.1f}%</b></div>
                    <div class="mcard-row"><span>وضعیت</span><b>{esc(r['وضعیت'])}</b></div>
                </div>
                """)
        else:
            st.dataframe(df, use_container_width=True, hide_index=True,
                         column_config=target_column_config(df))

    # ═════════════════════════════════════════
    # اگه تعداد ردیف کمه، فقط یک لیست
    # ═════════════════════════════════════════
    if n == 1:
        r = rank_df.iloc[0]
        st.info(f"📌 این شعبه رتبه **{int(r['رتبه'])}** از **{total_n}** شعبه دیستریکت را دارد.")
        show_rank(rank_df, "📌 این شعبه")
        return

    if n <= 15:
        # یک لیست کامل، بدون بهترین/بدترین
        show_rank(rank_df, f"📋 لیست شعب ({n} شعبه از {total_n})")
    else:
        # حالت عادی: بهترین ۱۰ و بدترین ۱۰
        top10 = rank_df.head(10)
        show_rank(top10, "🥇 بهترین ۱۰ شعبه")
        bottom10 = rank_df.sort_values("رتبه", ascending=False).head(10).reset_index(drop=True)
        # اگه با top10 همپوشانی داشت، فقط بخش پایین
        if n > 20:
            show_rank(bottom10, "⚠️ بدترین ۱۰ شعبه")
        else:
            show_rank(bottom10, "📋 پایین‌ترین رتبه‌ها")

    # رتبه‌بندی سرپرست‌ها (فقط برای دیستریکت)
    if sup_col and all_df is None and n > 15:
        st.markdown("**👤 رتبه‌بندی سرپرست‌ها**")
        agg = {ach_col: "mean"}
        if raked_col:
            agg[raked_col] = "sum"
        agg[branch_col] = "count"
        sup_rank = (t_df.groupby(sup_col).agg(agg).reset_index()
                    .rename(columns={branch_col: "تعداد شعبه", ach_col: "میانگین تحقق"})
                    .sort_values("میانگین تحقق", ascending=False).reset_index(drop=True))
        sup_rank.insert(0, "رتبه", range(1, len(sup_rank) + 1))

        if is_mobile:
            for _, r in sup_rank.head(20).iterrows():
                _, bg, icon = status_badge(r["میانگین تحقق"])
                st_md(f"""
                <div class="mcard">
                    <div class="mcard-badge" style="background:{bg};">{icon} رتبه {r['رتبه']}</div>
                    <div class="mcard-title">👤 {esc(r[sup_col])}</div>
                    <div class="mcard-row"><span>شعب</span><b>{int(r['تعداد شعبه'])}</b></div>
                    <div class="mcard-row"><span>میانگین تحقق</span><b>{r['میانگین تحقق']:.1f}%</b></div>
                </div>
                """)
        else:
            st.dataframe(sup_rank, use_container_width=True, hide_index=True)



def render_weekly_target(t_df):
    if t_df is None or t_df.empty:
        return
    cm = detect_target_columns(t_df)
    target_col, ach_col = cm.get("target"), cm.get("achievement")
    if not target_col or not ach_col:
        return
    d = t_df[[c for c in [cm.get("branch"), target_col, ach_col] if c]].copy()
    d[target_col] = to_number(d[target_col])
    d[ach_col] = to_number(d[ach_col])
    d = d[(d[target_col] > 0) | (d[ach_col] > 0)]
    if d.empty:
        return
    target = d[target_col].mean()
    achieved = d[ach_col].mean()
    gap = achieved - target
    # هفته ایران از شنبه شروع می‌شود (شنبه=۱ ... جمعه=۷) و بر اساس ساعت تهران
    elapsed = (now_tehran().weekday() + 2) % 7 + 1
    forecast = achieved / elapsed * 7

    st.markdown("### 🎯 تارگت هفتگی")
    if is_mobile:
        st.metric("تارگت", f"{target:.1f}%")
        st.metric("تحقق فعلی", f"{achieved:.1f}%",
                  delta=f"{gap:+.1f}%", delta_color="normal")
        st.metric("پیش‌بینی پایان هفته", f"{forecast:.1f}%")
    else:
        metric_row([
            ("🎯 تارگت", f"{target:.1f}%"),
            ("✅ تحقق", f"{achieved:.1f}%", {"delta": f"{gap:+.1f}%"}),
            ("📉 فاصله", f"{max(target - achieved, 0):.1f}%"),
            ("🔮 پیش‌بینی", f"{forecast:.1f}%"),
        ])

    status = ("🟢 روی تارگت" if achieved >= target
              else ("🟡 قابل جبران" if forecast >= target else "🔴 عقب از برنامه"))
    cls = "ok" if achieved >= target else "warn"
    st_md(f'<div class="alert-card {cls}"><b>وضعیت:</b> {status}</div>')


def render_target(t_df, key_suffix="", all_df=None):
    if t_df is None:
        st.warning("⚠️ فایل تارگت موجود نیست.")
        return
    if t_df.empty:
        st.info("داده‌ای برای این فیلتر نیست.")
        return
    render_target_kpis(t_df)
    st.divider()
    active = collapse_group(["📋 جدول", "📈 روند", "🏆 رتبه‌بندی"],
                            "target", key_suffix)
    if active == 0:
        render_target_table(t_df, key_suffix)
    elif active == 1:
        render_target_trend(t_df, key_suffix)
    elif active == 2:
        render_ranking(t_df, key_suffix, all_df=all_df)
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")



# ================== Session State Helper ==================
def _set_page(target):
    st.session_state["page"] = target
    st.rerun()


def _nav_card(target, icon, label, sub="", key=""):
    """کارت با دکمه — آیکون بالا، label پایین"""
    # ساخت متن: آیکون (bold بزرگ) + label + sub
    if sub:
        md_text = f"**{icon}**\n{label}\n{sub}"
    else:
        md_text = f"**{icon}**\n{label}"

    if st.button(md_text, key=f"nav_{target}_{key}", use_container_width=True):
        st.session_state["page"] = target
        st.rerun()






# ================== Login + Splash ==================
import sqlite3 as _sqlite3
from datetime import datetime as _dt
import time as _time


def _log_activity(personnel_code, page_name, action="view"):
    try:
        conn = _sqlite3.connect(HISTORY_DB)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS user_activity (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                personnel_code TEXT,
                page TEXT,
                action TEXT,
                ts TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute(
            "INSERT INTO user_activity (personnel_code, page, action) VALUES (?, ?, ?)",
            (str(personnel_code or ""), str(page_name or ""), str(action or ""))
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def _splash_screen(personnel_code):
    st_md(f"""
    <div style="display:flex; flex-direction:column; align-items:center;
                justify-content:center; min-height:70vh; text-align:center;
                color:#e5e7eb;">
        <div style="font-size:3rem; margin-bottom:16px;">⏳</div>
        <h2 style="color:#FF1F5A; margin-bottom:8px;">لطفاً صبر کنید</h2>
        <p style="color:#9ca3af; font-size:0.9rem;">
            در حال آماده‌سازی داشبورد برای کد <b>{personnel_code}</b>
        </p>
    </div>
    """)
    try:
        _ = load_all(RAAKED_FILE, _mtime(RAAKED_FILE))
    except Exception:
        pass
    _time.sleep(0.5)


def _render_login_gate():
    st.markdown("## 🔐 ورود به داشبورد")
    st_md("""
    <div class="app-hero" style="text-align:center; padding:26px 18px;">
        <h2>🛒 داشبورد کالای راکد</h2>
        <p>فروشگاه‌های زنجیره‌ای افق کوروش</p>
    </div>
    """)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        code = st.text_input("🔑 کد پرسنلی", key="login_code_input",
                              placeholder="کد پرسنلی خود را وارد کنید")
        if st.button("ورود", type="primary", use_container_width=True, key="login_btn"):
            code_clean = (code or "").strip()
            if not code_clean:
                st.error("❌ لطفاً کد پرسنلی را وارد کنید")
            else:
                # جستجو در فایل پرسنلی
                personel = load_personel()
                user_info = personel.get(code_clean, {})
                st.session_state["current_user"] = {
                    "personnel_code": code_clean,
                    "login_time": _dt.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "name": user_info.get("name", ""),
                    "role": user_info.get("role", "other"),
                    "branch_code": user_info.get("branch_code", ""),
                    "branch_name": user_info.get("branch_name", ""),
                    "mobile": user_info.get("mobile", ""),
                    "known": bool(user_info),
                }
                _log_activity(code_clean, "login", "login")
                st.session_state["show_splash"] = True
                st.rerun()
        st.markdown("")
        st.caption("🔒 تمام ورودها در سیستم ثبت می‌شوند.")


def _logout():
    u = st.session_state.get("current_user", {})
    if u:
        _log_activity(u.get("personnel_code"), "logout", "logout")
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()


def _check_locked_access(page_label, key_suffix="main"):
    """رمز مشترک برای سه صفحه — یک بار وارد کردن کافیه"""
    if _passwords_disabled():
        return True
    if st.session_state.get("unlock_locked_pages"):
        return True

    st.warning(f"🔒 صفحه «{page_label}» محافظت‌شده است.")
    st.caption("یک بار رمز رو وارد کن، برای هر سه صفحه (گزارش، تحلیل، آپلود) کافیه.")

    col1, col2 = st.columns([2, 3])
    with col1:
        pw = st.text_input("🔑 رمز مدیریت", type="password",
                            key=f"pw_locked_{key_suffix}")
    with col2:
        st.write("")
        st.write("")
        if st.button("ورود", type="primary", key=f"unlock_{key_suffix}",
                     use_container_width=True):
            import hmac
            try:
                from report_sender import _get_admin_password
                expected = _get_admin_password()
            except Exception:
                expected = ""
            if pw and expected and hmac.compare_digest(pw.encode(), expected.encode()):
                st.session_state["unlock_locked_pages"] = True
                st.rerun()
            else:
                st.error("❌ رمز اشتباه")
    st.stop()




def _passwords_disabled():
    """بررسی می‌کنه که رمزها خاموش باشن یا نه"""
    import configparser
    if os.environ.get("DISABLE_ALL_PASSWORDS", "").strip().lower() in ("1", "true", "yes"):
        return True
    try:
        cfg = configparser.ConfigParser()
        cfg.read(Path(__file__).parent / "config.ini", encoding="utf-8")
        return cfg.getboolean("SECURITY", "disable_all_passwords", fallback=False)
    except Exception:
        return False


# ================== Login Check ==================
if _passwords_disabled():
    if "current_user" not in st.session_state:
        st.session_state["current_user"] = {
            "personnel_code": "---",
            "login_time": _dt.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
else:
    if "current_user" not in st.session_state:
        _render_login_gate()
        st.stop()

    if st.session_state.get("show_splash"):
        _splash_screen(st.session_state["current_user"].get("personnel_code", ""))
        st.session_state["show_splash"] = False
        st.rerun()




# ================== پروژه پرزنتی راکد ==================
def _render_presentation_data(df, title, sales_df=None, key_suffix=""):
    """نمایش داده راکد با فیلترهای هوشمند"""
    if df is None or df.empty:
        st.info("داده‌ای برای نمایش نیست.")
        return

    # نرمال‌سازی نام‌ها برای فیلتر
    df = df.copy()
    if BR in df.columns:
        df[BR] = normalize_name(df[BR])
    if SUP in df.columns:
        df[SUP] = normalize_name(df[SUP])

    # KPI
    total_val = df[VAL].sum()
    total_qty = df[QTY].sum()
    branch_count = df[BR].nunique()
    item_count = len(df)

    if is_mobile:
        c1, c2 = st.columns(2, gap="small")
        c1.metric("💰 ارزش کل", money(total_val, short=True, with_unit=True))
        c2.metric("📦 تعداد اقلام", f"{item_count:,}")
        c3, c4 = st.columns(2, gap="small")
        c3.metric("🏪 تعداد شعب", f"{branch_count:,}")
        c4.metric("📊 موجودی کل", f"{total_qty:,.0f}")
    else:
        metric_row([
            ("💰 ارزش کل راکد", money(total_val)),
            ("📦 تعداد اقلام", f"{item_count:,}"),
            ("🏪 تعداد شعب", f"{branch_count:,}"),
            ("📊 موجودی کل", f"{total_qty:,.0f}"),
        ])

    st.divider()

    # فیلترها
    with st.expander("🎛 فیلترها", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            sup_options = ["همه"]
            if SUP in df.columns:
                sup_options += sorted(df[SUP].dropna().unique().tolist())
            sel_sup = st.selectbox("سرپرست", sup_options,
                                    key=f"pres_sup_{key_suffix}")

        # شعب رو بر اساس سرپرست فیلتر کن
        if sel_sup != "همه" and SUP in df.columns:
            branch_options = ["همه"] + sorted(
                df[df[SUP] == sel_sup][BR].dropna().unique().tolist()
            )
        else:
            branch_options = ["همه"] + sorted(df[BR].dropna().unique().tolist())

        with col2:
            sel_branch = st.selectbox("شعبه", branch_options,
                                       key=f"pres_branch_{key_suffix}")

        search_q = st.text_input("🔍 جستجو (نام کالا / بارکد / شعبه)",
                                  key=f"pres_search_{key_suffix}",
                                  placeholder="مثلاً: برنج یا 6260...")

    # اعمال فیلترها
    view = df.copy()
    if sel_sup != "همه" and SUP in view.columns:
        view = view[view[SUP] == sel_sup]
    if sel_branch != "همه":
        view = view[view[BR] == sel_branch]
    if search_q and search_q.strip():
        terms = [normalize_query(t) for t in search_q.split() if t.strip()]
        mask = pd.Series(True, index=view.index)
        for term in terms:
            m = pd.Series(False, index=view.index)
            for c in view.columns:
                txt = (view[c].astype(str).str.lower().str.translate(FA_DIGITS)
                       .str.replace("ي", "ی", regex=False)
                       .str.replace("ك", "ک", regex=False))
                m |= txt.str.contains(term, na=False, regex=False)
            mask &= m
        view = view[mask]

    # پیام‌های کمکی
    if view.empty:
        st.warning(f"⚠️ هیچ ردیفی با این فیلترها پیدا نشد.")
        st.info(f"💡 مجموع: {len(df)} ردیف | "
                f"بعد از فیلتر سرپرست: {len(df[df[SUP] == sel_sup]) if sel_sup != 'همه' and SUP in df.columns else len(df)} ردیف")
        if sel_branch != "همه" and sel_sup != "همه" and SUP in df.columns:
            b_sups = df[df[BR] == sel_branch][SUP].dropna().unique().tolist()
            st.caption(f"📌 سرپرست‌های «{sel_branch}»: {', '.join(b_sups) if b_sups else 'هیچ‌کدام'}")
        return

    st.caption(f"🔸 {len(view):,} ردیف | ارزش: {money(view[VAL].sum(), short=is_mobile)}")

    # نمایش
    if is_mobile:
        render_mobile_cards(view, max_rows=100, title_col=NM, sub_col=BR,
                            priority_cols=[BC, QTY, VAL, SUP])
    else:
        display_cols = [c for c in [BR, BCODE, BC, NM, QTY, VAL, SUP] if c in view.columns]
        st.dataframe(view[display_cols].head(500),
                     use_container_width=True, hide_index=True,
                     column_config=column_config(view[display_cols]))

    st.download_button(
        f"⬇️ دانلود {title} (CSV)",
        view.to_csv(index=False).encode("utf-8-sig"),
        f"presentation_{key_suffix}.csv", "text/csv",
        key=f"dl_pres_{key_suffix}"
    )



def render_presentation_page(pq45, pq60, df60, df45, key_suffix="pres"):
    st.markdown("### 📦 پروژه پرزنتی راکد")
    st.caption("نمایش کامل ۱۰ قلم برتر هر شعبه و همه اقلام راکد")

    active = collapse_group([
        "🎯 ۱۰ قلم ۶۰ روزه",
        "🎯 ۱۰ قلم ۴۵ روزه",
        "📊 همه ۶۰ روزه",
        "📊 همه ۴۵ روزه",
    ], "pres", key_suffix)

    if active == 0:
        if pq60 is None or pq60.empty:
            st.info("داده‌های pq60 در فایل موجود نیست.")
        else:
            _render_presentation_data(pq60, "۱۰ قلم ۶۰ روزه",
                                       key_suffix=f"{key_suffix}_pq60")
    elif active == 1:
        if pq45 is None or pq45.empty:
            st.info("داده‌های pq45 در فایل موجود نیست.")
        else:
            _render_presentation_data(pq45, "۱۰ قلم ۴۵ روزه",
                                       key_suffix=f"{key_suffix}_pq45")
    elif active == 2:
        _render_presentation_data(df60, "همه اقلام ۶۰ روزه",
                                   key_suffix=f"{key_suffix}_tbl60")
    elif active == 3:
        _render_presentation_data(df45, "همه اقلام ۴۵ روزه",
                                   key_suffix=f"{key_suffix}_tbl45")
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")




# ================== صفحه «شعبه من» ==================
def render_my_store(user_info, df60, df45, pq45, pq60, df_sales, df_target):
    """صفحه اختصاصی کارمند/مسئول فروشگاه"""
    branch_code = user_info.get("branch_code", "")
    branch_name = user_info.get("branch_name", "")
    user_name = user_info.get("name", "")

    def _find_branch(df):
        if df is None or df.empty:
            return df
        bc = df.get(BCODE)
        if bc is None:
            return df.iloc[0:0]
        return df[df[BCODE].astype(str).str.upper() == branch_code.upper()]

    b60 = _find_branch(df60)
    b45 = _find_branch(df45)
    b_target = None
    if df_target is not None:
        tc = detect_target_columns(df_target)
        code_c = tc.get("code")
        if code_c:
            b_target = df_target[
                df_target[code_c].astype(str).str.upper() == branch_code.upper()
            ].copy()

    actual_name = b60[BR].iloc[0] if not b60.empty else (branch_name or "شعبه شما")

    # ── هدر ──
    st_md(f"""
    <div class="app-hero">
        <h2>🏪 {actual_name}</h2>
        <p>👤 {user_name} — کد پرسنلی: {user_info.get('personnel_code', '')}</p>
    </div>
    """)

    if b60.empty and b45.empty:
        st.warning(f"⚠️ اطلاعات شعبه‌ای با کد {branch_code} پیدا نشد.")
        return

    # ══════════════════════════════════════════════════
    # ۱. تارگت و رتبه من (اول)
    # ══════════════════════════════════════════════════
    if b_target is not None and not b_target.empty and df_target is not None:
        st.markdown("### 🎯 تارگت و رتبه من")
        render_target(b_target, key_suffix=f"my_{branch_code}", all_df=df_target)
    else:
        st.info("ℹ️ تارگت برای شعبه شما تنظیم نشده.")

    st.divider()

    # ══════════════════════════════════════════════════
    # ۲. KPIها
    # ══════════════════════════════════════════════════
    st.markdown("### 📊 وضعیت کالاهای راکد")
    total_60 = b60[VAL].sum() if not b60.empty else 0
    total_45 = b45[VAL].sum() if not b45.empty else 0
    count_60 = len(b60)
    count_45 = len(b45)

    c1, c2 = st.columns(2, gap="small")
    c1.metric("💰 راکد ۶۰ روزه", money(total_60, short=True, with_unit=True))
    c2.metric("📦 تعداد اقلام", f"{count_60:,}")

    c3, c4 = st.columns(2, gap="small")
    c3.metric("📊 راکد ۴۵ روزه", money(total_45, short=True, with_unit=True))
    c4.metric("📋 اقلام ۴۵", f"{count_45:,}")

    st.divider()

    # ══════════════════════════════════════════════════
    # ۳. چک‌لیست شیفت من
    # ══════════════════════════════════════════════════
    st.markdown("### 📋 چک‌لیست شیفت من")
    _bc = branch_code.upper()
    my_pq60 = pd.DataFrame()
    my_pq45 = pd.DataFrame()
    if pq60 is not None and not pq60.empty and BCODE in pq60.columns:
        my_pq60 = pq60[pq60[BCODE].astype(str).str.upper() == _bc].copy()
    if pq45 is not None and not pq45.empty and BCODE in pq45.columns:
        my_pq45 = pq45[pq45[BCODE].astype(str).str.upper() == _bc].copy()

    active_my = collapse_group(
        [f"🌅 صبح ({len(my_pq45)})", f"🌙 عصر ({len(my_pq60)})"],
        "my_store_shift", branch_code
    )

    for idx_my, df, shift in ((0, my_pq45, "صبح"), (1, my_pq60, "عصر")):
        if active_my != idx_my:
            continue
        if df.empty:
            st.info(f"داده‌ای برای شیفت {shift} وجود ندارد.")
            continue

        st.markdown(f"**{len(df)} قلم برای شیفت {shift}**")
        if is_mobile:
            for _, r in df.head(30).iterrows():
                st_md(f"""
                <div class="mcard">
                    <div class="mcard-title">📦 {esc(r.get(NM, ''))}</div>
                    <div class="mcard-row"><span>بارکد</span><b>{esc(r.get(BC, ''))}</b></div>
                    <div class="mcard-row"><span>موجودی</span><b>{_fmt_cell(r.get(QTY, 0))}</b></div>
                    <div class="mcard-row"><span>ارزش راکد</span><b>{money(r.get(VAL, 0), short=True)}</b></div>
                </div>
                """)
        else:
            view = df[[c for c in [NM, BC, QTY, VAL] if c in df.columns]].copy()
            view.columns = ["نام کالا", "بارکد", "موجودی", "ارزش راکد"][:len(view.columns)]
            st.dataframe(view.head(50), use_container_width=True, hide_index=True)

    st.divider()

    # ══════════════════════════════════════════════════
    # ۴. کالاهای راکدی که فروش رفتند
    # ══════════════════════════════════════════════════
    st.markdown("### ✅ کالاهای راکدی که فروش رفتند")
    if df_sales is not None and not df_sales.empty and not b60.empty:
        matched = attach_sales(b60, df_sales)
        sold = matched[matched[S_QTY] > 0].copy()
        if sold.empty:
            st.info("هنوز چیزی از اقلام راکد شما فروش نرفته.")
        else:
            st.metric("تعداد فروش‌رفته", f"{len(sold):,}")
            show_table(sold, f"my_sold_{branch_code}", hide_supervisor=True,
                       placeholder="جستجو...", title_col=NM,
                       priority_cols=[QTY, VAL, S_QTY, S_AMT, S_REL])
    else:
        st.info("برای مشاهده این بخش، فایل فروش لازم است.")


# ================== Login Check ==================
if "current_user" not in st.session_state:
    _render_login_gate()
    st.stop()

if st.session_state.get("show_splash"):
    _splash_screen(st.session_state["current_user"].get("personnel_code", ""))
    st.session_state["show_splash"] = False
    _u = st.session_state.get("current_user", {})
    _role = _u.get("role", "other")
    if _role in ("store_manager", "store_staff", "store_deputy", "store_deputy2"):
        if _u.get("branch_code"):
            st.session_state["page"] = "my_store"
    st.rerun()


# ================== مسیریابی ==================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

page = st.session_state.get("page", "home")
if page not in VALID_PAGES:
    page = "home"

# ── محافظت: کارمندان فقط صفحه «شعبه من» ──
_u_check = st.session_state.get("current_user", {})
_role_check = _u_check.get("role", "other")
if _role_check in ("store_manager", "store_staff", "store_deputy", "store_deputy2"):
    if _u_check.get("branch_code") and page not in ("my_store", "home"):
        page = "my_store"
        st.session_state["page"] = "my_store"

# لاگ بازدید صفحه
_page_log_key = f"_logged_page_{page}"
if not st.session_state.get(_page_log_key):
    _u = st.session_state.get("current_user", {})
    _log_activity(_u.get("personnel_code"), page, "view")
    st.session_state[_page_log_key] = True

# ================== خواندن داده ==================
data_error = None
all_data = {"d60": None, "d45": None, "pq45": None, "pq60": None,
            "summary60": None, "sales": None, "sheets": []}
try:
    all_data = load_all(RAAKED_FILE, _mtime(RAAKED_FILE))
except Exception as e:
    data_error = str(e)

df60 = all_data.get("d60")
df45 = all_data.get("d45")
pq45 = all_data.get("pq45")
pq60 = all_data.get("pq60")
summary60 = all_data.get("summary60")

if df60 is None:
    df60 = pd.DataFrame(columns=[BR, BCODE, BC, NM, QTY, VAL, SUP])
if df45 is None:
    df45 = pd.DataFrame(columns=[BR, BCODE, BC, NM, QTY, VAL, SUP])

df_sales = None
if SALES_FILE.exists():
    df_sales = load_sales(_mtime(SALES_FILE))
if (df_sales is None or df_sales.empty) and all_data.get("sales") is not None:
    df_sales = all_data.get("sales")

df_target = load_target(_mtime(TARGET_FILE))

if data_error and page != "upload":
    render_header("خطا در بارگذاری داده")
    st.error(f"فایل راکد خوانده نشد: {data_error}")
    st_md('<a class="back-link" href="?page=upload" target="_self">📤 آپلود</a>')
    st.stop()

if page != "upload" and all_data.get("d60") is None:
    render_header("داده‌ای برای نمایش نیست")
    if not Path(RAAKED_FILE).exists():
        st.warning("فایل راکد پیدا نشد. از بخش آپلود، فایل راکد را بارگذاری کنید.")
    else:
        sheets = "، ".join(all_data.get("sheets") or []) or "—"
        st.warning(f"شیت‌های tbl60 و tbl45 در فایل راکد پیدا نشد. شیت‌های موجود: {sheets}")
    st_md('<a class="back-link" href="?page=upload" target="_self">📤 آپلود</a>')
    st.stop()


def show_last_update_badge():
    last = get_last_update()
    if last and last[0]:
        date, kind, rows = last
        label = "راکد" if kind == "raaked" else "فروش"
        st_md(f'<div class="alert-card ok">📅 <b>آخرین آپدیت:</b> {date} | '
              f'{label} | {(rows or 0):,} ردیف</div>')




if page == "home":
    render_header("فروشگاه‌های زنجیره‌ای افق کوروش")

    user = st.session_state.get("current_user", {})
    user_name = user.get("personnel_code", "کاربر")

    st_md(f"""
    <div class="app-hero">
        <h2>سلام {user_name} 👋</h2>
        <p>داشبورد مدیریت کالای راکد — افق کوروش</p>
    </div>
    """)

    show_last_update_badge()

    # ── اگه کارمند/مسئول فروشگاهه، بفرست به «شعبه من» ──
    _u = st.session_state.get("current_user", {})
    _role = _u.get("role", "other")
    if _role in ("store_manager", "store_staff", "store_deputy", "store_deputy2"):
        if _u.get("branch_code"):
            st.info(f"👋 خوش آمدی {_u.get('name', '')}! در حال رفتن به صفحه شعبه...")
            if st.button("🏪 برو به شعبه من", type="primary", use_container_width=True):
                st.session_state["page"] = "my_store"
                st.rerun()
            st.stop()

    st_md('<div class="app-section-title">📌 بخش‌های اصلی</div>')

    c1, c2 = st.columns(2)
    with c1:
        _nav_card("store", "🏪", "عملکرد فروشگاه‌ها", "تحلیل هر شعبه", "h1")
    with c2:
        _nav_card("supervisor", "👤", "عملکرد سوپروایزرها", "عملکرد هر سرپرست", "h2")

    c3, c4 = st.columns(2)
    with c3:
        _nav_card("target", "🎯", "تارگت و روند", "اهداف و رتبه‌بندی", "h3")
    with c4:
        _nav_card("district", "📊", "عملکرد دیستریکت", "KPI کلی", "h4")

    st_md('<div class="app-section-title">📦 پروژه ویژه</div>')

    _nav_card("presentation", "📦", "پروژه پرزنتی راکد",
              "۱۰ قلم برتر + همه اقلام", "h_pres")

    st_md('<div class="app-section-title">🎯 پیگیری</div>')

    c5, c6 = st.columns(2)
    with c5:
        _nav_card("shift", "📋", "چک‌لیست شیفت", "صبح و عصر", "h5")
    with c6:
        _nav_card("trend", "📈", "روند و مقایسه", "افت و رشد", "h6")

    st_md('<div class="footer-text">ساخته شده توسط <b>شاهین باقری</b></div>')

elif page == "district":
    back_link()
    render_header("عملکرد کلی دیستریکت")
    show_last_update_badge()
    st.divider()
    render_management_kpis(df60, df45, df_sales)
    render_kpis(df60, df45)
    st.divider()
    st.subheader("🎯 تارگت هفتگی")
    render_weekly_target(df_target)
    st.divider()
    render_action_center(df60, df_sales, key="district")
    st.divider()
    active = collapse_group(["🏪 شعب", "👤 سرپرست‌ها"], "district", "main")
    if active == 0:
        show_summary(branch_summary(df60, df45, df_sales), "branches_summary")
    elif active == 1:
        show_summary(supervisor_summary(df60, df45, df_sales), "supervisors_summary")
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")
    st.divider()
    if st.button("📥 خروجی اکسل کامل", key="export_district"):
        with st.spinner("در حال ساخت..."):
            xls_bytes = export_to_excel({
                "خلاصه شعب": branch_summary(df60, df45, df_sales),
                "خلاصه سرپرست‌ها": supervisor_summary(df60, df45, df_sales),
                "راکد ۶۰ روزه": attach_sales(df60, df_sales),
                "راکد ۴۵ روزه": attach_sales(df45, df_sales),
                "تارگت": df_target,
            })
        st.download_button("⬇️ دانلود", xls_bytes,
                           f"district_{now_tehran():%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_district_xls")
    st.divider()
    st.subheader("🚨 کالاهای راکدِ بدون فروش")
    render_top_products(df60, df45, df_sales, key_suffix="district")
    st.divider()
    st.subheader("✅ کالاهای راکدی که فروش رفتند")
    render_matched(df60, df45, df_sales, key_suffix="district")

elif page == "store":
    back_link()
    render_header("عملکرد فروشگاه‌ها")
    branch_list = sorted(df60[BR].dropna().unique().tolist())
    if not branch_list:
        st.info("شعبه‌ای وجود ندارد.")
        st.stop()
    selected_branch = st.selectbox("انتخاب فروشگاه", branch_list, key="sb_branch")
    b60 = df60[df60[BR] == selected_branch].copy()
    b45 = df45[df45[BR] == selected_branch].copy()
    branch_sales = (df_sales[df_sales[SALES_STORE] == selected_branch].copy()
                    if df_sales is not None and not df_sales.empty else None)
    branch_target = None
    if df_target is not None:
        tc = detect_target_columns(df_target)
        if tc["branch"]:
            branch_target = df_target[df_target[tc["branch"]] == selected_branch].copy()
    sup_names = b60[SUP].dropna().unique().tolist()
    st.markdown(f"**سوپروایزر:** {sup_names[0] if sup_names else '—'}")
    st.divider()
    render_management_kpis(b60, b45, branch_sales)
    render_kpis(b60, b45)
    st.divider()
    active = collapse_group(["📦 اقلام راکد", "🎯 تارگت"],
                            "store", f"br_{selected_branch}")
    if active == 0:
        st.subheader("🚨 کالاهای راکدِ بدون فروش")
        render_top_products(b60, b45, branch_sales,
                            key_suffix=f"branch_{selected_branch}", hide_sup=True)
        st.divider()
        st.subheader("✅ کالاهای راکدی که فروش رفتند")
        render_matched(b60, b45, branch_sales,
                       key_suffix=f"branch_m_{selected_branch}", hide_sup=True)
    elif active == 1:
        render_target(branch_target, key_suffix=f"branch_{selected_branch}",
                      all_df=df_target)
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")

elif page == "supervisor":
    back_link()
    render_header("عملکرد سوپروایزرها")
    supervisor_list = sorted(df60[SUP].dropna().unique().tolist())
    if not supervisor_list:
        st.info("سوپروایزری وجود ندارد.")
        st.stop()
    selected_sup = st.selectbox("انتخاب سوپروایزر", supervisor_list, key="sb_sup")
    s60 = df60[df60[SUP] == selected_sup].copy()
    sup_branches = sorted(s60[BR].dropna().unique().tolist())
    s45 = df45[df45[BR].isin(sup_branches)].copy()
    sup_sales = (df_sales[df_sales[SALES_STORE].isin(sup_branches)].copy()
                 if df_sales is not None and not df_sales.empty else None)
    sup_target = None
    if df_target is not None:
        tc = detect_target_columns(df_target)
        if tc["branch"]:
            sup_target = df_target[df_target[tc["branch"]].isin(sup_branches)].copy()
    st.divider()
    render_management_kpis(s60, s45, sup_sales)
    render_kpis(s60, s45)
    st.markdown(f"**شعبه‌های تحت پوشش:** {len(sup_branches)}")
    with st.expander("📋 وضعیت شعبه‌ها"):
        show_summary(branch_summary(s60, s45, sup_sales).drop(columns=[SUP], errors="ignore"),
                     f"sup_branches_{selected_sup}")
    active = collapse_group(["📦 اقلام راکد", "🎯 تارگت"],
                            "supervisor", f"sup_{selected_sup}")
    if active == 0:
        st.subheader("🚨 کالاهای راکدِ بدون فروش")
        render_top_products(s60, s45, sup_sales, key_suffix=f"sup_{selected_sup}", hide_sup=True)
        st.divider()
        st.subheader("✅ کالاهای راکدی که فروش رفتند")
        render_matched(s60, s45, sup_sales, key_suffix=f"sup_m_{selected_sup}", hide_sup=True)
    elif active == 1:
        render_target(sup_target, key_suffix=f"sup_{selected_sup}",
                      all_df=df_target)
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")

elif page == "target":
    back_link()
    render_header("تارگت و روند کل دیستریکت")
    render_target(df_target, key_suffix="district")

elif page == "analytics":
    back_link()
    render_header("تحلیل پیشرفته")
    _check_locked_access("تحلیل پیشرفته", "analytics")
    active = collapse_group(["💰 پول خوابیده", "📊 نمودارها", "📈 مقایسه"],
                            "analytics", "main")
    if active == 0:
        render_cash_and_abc(df60, df_sales)
    elif active == 1:
        render_charts(df60, df_target)
    elif active == 2:
        render_comparison()
        st.divider()
        render_history_trend()
    else:
        st.info("👆 یکی از بخش‌ها رو انتخاب کن.")
    st.divider()
    render_action_center(df60, df_sales, key="analytics")

elif page == "shift":
    back_link()
    render_header("چک‌لیست شیفت صبح و عصر")
    show_last_update_badge()
    render_shift_checklist(pq45, pq60, key_suffix="main")

elif page == "trend":
    back_link()
    render_header("روند و مقایسه شعب")
    show_last_update_badge()
    render_trend_comparison(summary60, key_suffix="main")

elif page == "upload":
    back_link()
    render_header("آپلود فایل و تاریخچه")
    _check_locked_access("آپلود و تاریخچه", "upload")
    render_upload()
elif page == "presentation":
    back_link()
    render_header("پروژه پرزنتی راکد")
    show_last_update_badge()
    render_presentation_page(pq45, pq60, df60, df45, key_suffix="main")

elif page == "my_store":
    user = st.session_state.get("current_user", {})
    render_header(f"شعبه من — {user.get('name', '')}")
    show_last_update_badge()
    render_my_store(user, df60, df45, pq45, pq60, df_sales, df_target)

elif page == "report":
    back_link()
    render_header("ارسال گزارش")
    _check_locked_access("ارسال گزارش", "report")
    from report_sender import render_report_sender
    render_report_sender(key_suffix="main")

# ================== Bottom Navigation ==================
def _bottom_nav():
    if "current_user" not in st.session_state:
        return
    st.markdown('<div class="bnav-marker"></div>', unsafe_allow_html=True)
    cols = st.columns(3)
    items = [("report", "📨", "گزارش"),
             ("analytics", "💰", "تحلیل"),
             ("upload", "📤", "آپلود")]
    for col, (target, icon, label) in zip(cols, items):
        with col:
            active = "● " if page == target else ""
            if st.button(f"{active}{icon} {label}", key=f"bnav_{target}",
                         use_container_width=True, type="primary"):
                st.session_state["page"] = target
                st.rerun()


if "current_user" in st.session_state and page == "home":
    _bottom_nav()
