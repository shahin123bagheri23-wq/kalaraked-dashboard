import base64
import hmac
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

st.set_page_config(
    page_title="داشبورد کالای راکد | افق کوروش",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="collapsed",
)

screen_width = streamlit_js_eval(js_expressions="window.innerWidth", key="WIDTH")
is_mobile = (screen_width or 1200) < 768

OK_RED = "#E6003E"
OK_RED_DARK = "#A30029"
OK_RED_LIGHT = "#FF1F5A"

DATA_DIR = Path(os.environ.get("DATA_DIR", "."))
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR = DATA_DIR / "backups"

RAAKED_FILE = DATA_DIR / "اقلام راکد 60  45   دیتا (1).xlsx"
SALES_FILE = DATA_DIR / "گزارش فروش راکد.xlsx"
TARGET_FILE = DATA_DIR / "فایل تارگت.xlsx"
TARGET_SHEET = "روند و تارگت"
HISTORY_DB = DATA_DIR / "history.db"

SALES_BARCODE = "بارکد کالا"
SALES_QTY = "فروش تعدادی"
SALES_AMOUNT = "فروش خالص با مالیات"
SALES_STORE = "نام انبار/فروشگاه"

BR, BC, NM, QTY, VAL, SUP = ("نام شعبه", "بارکد", "نام کالا", "موجودی سیستمی",
                             "موجودی ریالی اقلام راکد", "سوپروایزر")
S_QTY, S_AMT, S_REL = "فروش (تعداد)", "فروش (ریال)", "ارزش آزادشده"

REQUIRED_RAAKED = [BR, BC, NM, QTY, VAL]
REQUIRED_SALES = [SALES_BARCODE, SALES_QTY, SALES_AMOUNT, SALES_STORE]

ACH_OK, ACH_WARN = 100, 80
HOLD_RATE_DEFAULT = 2.0
VALID_PAGES = {"home", "district", "store", "supervisor", "target", "analytics", "upload"}

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
    }}
    .stApp {{ background-color: {BG}; }}
    .block-container {{
        padding-top: 0.5rem; padding-bottom: 1rem;
        padding-left: 0.6rem; padding-right: 0.6rem;
    }}
    input, textarea, select {{
        background-color: {INPUT_BG} !important;
        color: {TXT} !important;
    }}
    div[data-baseweb="select"] > div {{
        background-color: {INPUT_BG} !important;
        color: {TXT} !important;
    }}
    button[kind="secondary"] {{
        background-color: {BG2} !important;
        color: {TXT} !important;
    }}
    .ok-header {{
        background: linear-gradient(135deg, {OK_RED_DARK} 0%, {OK_RED} 55%, {OK_RED_LIGHT} 100%);
        padding: 18px 25px; border-radius: 15px;
        display: flex; align-items: center; justify-content: space-between;
        color: white; margin-bottom: 20px;
        box-shadow: 0 6px 20px rgba(230,0,62,0.35);
    }}
    .ok-header-title {{ font-size: 1.35rem; font-weight: bold; color: white; margin: 0; }}
    .ok-header-sub {{ font-size: 0.85rem; color: #ffe0e8; margin: 3px 0 0 0; }}
    .ok-logo-wrap {{
        background: white; padding: 6px 12px; border-radius: 12px;
        display: inline-flex; align-items: center; justify-content: center;
        box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }}
    a.card-link {{ text-decoration: none !important; display: block; margin-bottom: 15px; }}
    a.card-link:hover {{ text-decoration: none !important; }}
    .home-card {{
        background: linear-gradient(135deg, {OK_RED_DARK} 0%, {OK_RED} 100%);
        padding: 35px 20px; border-radius: 20px; text-align: center;
        color: white; box-shadow: 0 4px 15px rgba(230,0,62,0.3);
        transition: transform 0.2s, box-shadow 0.2s; cursor: pointer;
        min-height: 220px; border: 2px solid transparent;
    }}
    a.card-link:hover .home-card {{
        transform: translateY(-4px);
        box-shadow: 0 10px 30px rgba(230,0,62,0.5);
        border-color: {OK_RED_LIGHT};
    }}
    .home-card.light {{
        background: {LIGHT_CARD};
        color: {LIGHT_CARD_TXT} !important;
        border: 2px solid {OK_RED};
    }}
    .home-card.light h2 {{ color: {LIGHT_CARD_TXT} !important; }}
    .home-card.light p {{ color: {LIGHT_CARD_TXT2} !important; }}
    .home-card h2 {{ color: white !important; margin: 10px 0 8px 0; font-size: 1.25rem; }}
    .home-card p {{ color: #ffe0e8; font-size: 0.85rem; margin: 0; }}
    .home-card .icon {{ font-size: 3rem; }}
    a.back-link {{
        display: inline-block; padding: 8px 18px;
        background: {OK_RED}; color: white !important;
        border-radius: 10px; text-decoration: none !important;
        font-weight: bold; font-size: 0.9rem; margin-bottom: 15px;
    }}
    a.back-link:hover {{ background: {OK_RED_DARK}; }}
    div[data-testid="stMetric"] {{
        background: {CARD_BG}; padding: 12px; border-radius: 12px;
        margin-bottom: 8px; border-right: 5px solid {OK_RED};
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }}
    div[data-testid="stMetric"] label {{ color: {TXT2} !important; }}
    div[data-testid="stMetricValue"] {{ color: {OK_RED_LIGHT} !important; }}
    h1, h2, h3 {{ color: {OK_RED_LIGHT} !important; }}
    .footer-text {{
        text-align: center; color: {TXT2}; font-size: 0.85rem;
        padding: 30px 0 10px 0; border-top: 2px solid {OK_RED}; margin-top: 30px;
    }}
    .footer-text b {{ color: {OK_RED_LIGHT}; }}
    button[data-baseweb="tab"][aria-selected="true"] {{
        color: {OK_RED_LIGHT} !important;
        border-bottom-color: {OK_RED} !important;
    }}
    div[data-testid="stSlider"] div[role="slider"] {{
        background-color: {OK_RED} !important;
    }}
    .alert-card {{
        background: {CARD_BG}; padding: 15px 20px;
        border-radius: 12px; border-right: 5px solid {OK_RED};
        margin-bottom: 10px; color: {TXT};
    }}
    .alert-card.warn {{ border-right-color: #f59e0b; }}
    .alert-card.ok {{ border-right-color: #10b981; }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
</style>
""", unsafe_allow_html=True)

if is_mobile:
    st.markdown("""
    <style>
        h1 { font-size: 1.4rem !important; }
        h2 { font-size: 1.2rem !important; }
        h3 { font-size: 1.05rem !important; }
        .home-card { padding: 25px 15px; min-height: 170px; }
        .home-card h2 { font-size: 1.05rem; }
        .home-card .icon { font-size: 2.5rem; }
        .ok-header-title { font-size: 1rem; }
        .ok-header-sub { font-size: 0.75rem; }
        .ok-header { padding: 12px 15px; }
    </style>
    """, unsafe_allow_html=True)


# ================== توابع کمکی داده ==================
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
    raw = series.astype(str)
    has_sign = raw.str.contains("%").any()
    s = (raw.str.translate(FA_DIGITS)
         .str.replace(",", "", regex=False).str.replace("،", "", regex=False)
         .str.replace(" ", "", regex=False).str.replace("\u200c", "", regex=False)
         .str.replace("%", "", regex=False)
         .replace({"nan": None, "": None, "-": None, "None": None}))
    s = pd.to_numeric(s, errors="coerce").fillna(0)
    col_name = str(getattr(series, "name", "") or "")
    if (not has_sign and "درصد" in col_name
            and len(s) > 0 and s.abs().quantile(0.95) <= 1.5):
        s = s * 100
    return s


TOTAL_RE = r"^\s*(?:جمع|مجموع|[کك]ل|total|grand\s*total|sum)(?:\s|$)"


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
    """حذف ردیف‌هایی که نام شعبه در واقع بارکد است یا داده کاملاً خالی دارند."""
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


def money(v):
    return f"{float(v or 0):,.0f} ریال"


def valid_date(s):
    try:
        datetime.strptime(s.strip(), "%Y-%m-%d")
        return True
    except (ValueError, AttributeError):
        return False


# ================== خواندن اکسل ==================
def find_sheet(xls, key):
    for name in xls.sheet_names:
        if key in name.translate(FA_DIGITS):
            return name
    raise ValueError(f"شیت «{key} روزه» در فایل پیدا نشد. "
                     f"شیت‌های موجود: {'، '.join(xls.sheet_names)}")


def _prep_raaked(d):
    d = d.copy()
    d.columns = d.columns.astype(str).str.strip()
    missing = [c for c in REQUIRED_RAAKED if c not in d.columns]
    if missing:
        raise ValueError(
            f"ستون‌های لازم پیدا نشد: {'، '.join(missing)}\n"
            f"ستون‌های موجود در فایل: {'، '.join(d.columns.astype(str).tolist())}\n"
            f"💡 مطمئن شو فایل با ستون‌های استاندارد ذخیره شده."
        )
    d = remove_totals(d)
    d = remove_invalid_rows(d)
    d[BC] = clean_barcode(d[BC])
    d[BR] = normalize_name(d[BR])
    d[NM] = d[NM].fillna("").astype(str).str.strip()
    if SUP in d.columns:
        d[SUP] = normalize_name(d[SUP]).replace("", "نامشخص")
    else:
        d[SUP] = "نامشخص"
    d[QTY] = to_number(d[QTY])
    d[VAL] = to_number(d[VAL])
    return d.reset_index(drop=True)


def parse_raaked(source):
    xls = pd.ExcelFile(source)
    d60 = _prep_raaked(pd.read_excel(xls, sheet_name=find_sheet(xls, "60")))
    d45 = _prep_raaked(pd.read_excel(xls, sheet_name=find_sheet(xls, "45")))
    return d60, d45


def parse_sales(source):
    s = pd.read_excel(source, sheet_name=0)
    s.columns = s.columns.astype(str).str.strip()
    missing = [c for c in REQUIRED_SALES if c not in s.columns]
    if missing:
        raise ValueError(
            f"ستون‌های لازم پیدا نشد: {'، '.join(missing)}\n"
            f"ستون‌های موجود: {'، '.join(s.columns.astype(str).tolist())}\n"
            f"💡 مطمئن شو فایل فروش با ستون‌های استاندارد ذخیره شده."
        )
    s[SALES_BARCODE] = clean_barcode(s[SALES_BARCODE])
    s[SALES_STORE] = normalize_name(s[SALES_STORE])
    s[SALES_QTY] = to_number(s[SALES_QTY])
    s[SALES_AMOUNT] = to_number(s[SALES_AMOUNT])
    return s


def _mtime(path):
    try:
        return Path(path).stat().st_mtime
    except OSError:
        return 0


@st.cache_data(show_spinner="در حال بارگذاری داده‌ها…")
def load_raaked(mtime):
    return parse_raaked(RAAKED_FILE)


@st.cache_data(show_spinner=False)
def load_sales(mtime):
    if not SALES_FILE.exists():
        return None
    try:
        return parse_sales(SALES_FILE)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def load_target(mtime):
    if not TARGET_FILE.exists():
        return None
    try:
        t = pd.read_excel(TARGET_FILE, sheet_name=TARGET_SHEET)
        t.columns = t.columns.astype(str).str.strip()
        for c in t.columns:
            if "نام شعبه" in c or "سرپرست" in c:
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
        return {"branch": None, "code": None, "supervisor": None, "raked_value": None,
                "trend": [], "target": None, "achievement": None, "change": None}
    cols = list(t.columns)
    return {
        "branch": next((c for c in cols if "نام شعبه" in c), None),
        "code": next((c for c in cols if "کد" in c and "شعبه" in c), None),
        "supervisor": next((c for c in cols if "سرپرست" in c or "سوپروایزر" in c), None),
        "raked_value": next((c for c in cols if "ریالی راکد" in c), None),
        "trend": [c for c in cols if "درصد" in c and "راکد" in c],
        "target": next((c for c in cols if c.startswith("تارگت") and "تغییرات" not in c), None),
        "achievement": next((c for c in cols if "تحقق" in c), None),
        "change": next((c for c in cols if c.startswith("تغییرات") and "تارگت" not in c), None),
    }


# ================== SQLite ==================
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


# ================== فایل‌ها: بک‌آپ و نوشتن امن ==================
def _atomic_write(path, data):
    tmp = Path(str(path) + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def _backup(path, keep=10):
    p = Path(path)
    if not p.exists():
        return
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(p, BACKUP_DIR / f"{datetime.now():%Y%m%d_%H%M%S}__{p.name}")
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
        if df is None or df.empty:
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
            if dup: msg += f" | تکراری (شعبه+بارکد): {dup}"
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
    _backup(RAAKED_FILE)
    _atomic_write(RAAKED_FILE, data)
    return rows, data_quality_report(d60, d45)


def process_sales_upload(data, date_str, note):
    s = parse_sales(io.BytesIO(data))
    rows = save_sales_history(s, date_str, note)
    _backup(SALES_FILE)
    _atomic_write(SALES_FILE, data)
    return rows, data_quality_report(pd.DataFrame(), pd.DataFrame(), s)


# ================== UI کمکی ==================
def get_logo_html(size=50):
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
    st.markdown(f"""
    <div class="ok-header">
        <div>
            <p class="ok-header-title">داشبورد مدیریت کالای راکد</p>
            <p class="ok-header-sub">{subtitle or "فروشگاه‌های زنجیره‌ای افق کوروش"}</p>
        </div>
        <div class="ok-logo-wrap">{get_logo_html(size=50 if not is_mobile else 38)}</div>
    </div>
    """, unsafe_allow_html=True)


def back_link():
    st.markdown('<a class="back-link" href="?page=home" target="_self">⬅️ بازگشت</a>',
                unsafe_allow_html=True)


def metric_row(items):
    per_row = 2 if is_mobile else max(len(items), 1)
    for i in range(0, len(items), per_row):
        chunk = items[i:i + per_row]
        for col, it in zip(st.columns(len(chunk)), chunk):
            col.metric(it[0], it[1], **(it[2] if len(it) > 2 else {}))


def plotly_style(fig, title_size=15, show_legend_bg=False):
    fig.update_layout(
        paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG,
        font=dict(color=CHART_TEXT, size=12, family="Tahoma"),
        title_font=dict(color=CHART_MAIN, size=title_size, family="Tahoma"),
        legend=dict(font=dict(color=CHART_TEXT, size=11),
                    bgcolor="rgba(0,0,0,0)" if not show_legend_bg else CHART_BG),
        xaxis=dict(tickfont=dict(color=CHART_TEXT, size=11),
                   title_font=dict(color=CHART_TEXT2, size=12),
                   gridcolor=CHART_GRID, linecolor=CHART_GRID, zerolinecolor=CHART_GRID),
        yaxis=dict(tickfont=dict(color=CHART_TEXT, size=11),
                   title_font=dict(color=CHART_TEXT2, size=12),
                   gridcolor=CHART_GRID, linecolor=CHART_GRID, zerolinecolor=CHART_GRID),
        margin=dict(l=10, r=20, t=50, b=10),
    )
    return fig


# ================== جداول ==================
def build_sales_map(sales_df):
    if sales_df is None or sales_df.empty:
        return None
    return (sales_df.groupby([SALES_STORE, SALES_BARCODE])
            .agg({SALES_QTY: "sum", SALES_AMOUNT: "sum"})
            .reset_index()
            .rename(columns={SALES_STORE: BR, SALES_BARCODE: BC,
                             SALES_QTY: S_QTY, SALES_AMOUNT: S_AMT}))


@st.cache_data(show_spinner=False, max_entries=32)
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
    out[S_REL] = (np.minimum(out[S_QTY].clip(lower=0), out[QTY]) * unit).round(0).astype("int64")
    out[S_QTY] = smart_int(out[S_QTY])
    out[S_AMT] = out[S_AMT].round(0).astype("int64")
    cols = [BR, BC, NM, QTY, VAL, S_QTY, S_AMT, S_REL, SUP]
    return out[cols].sort_values(VAL, ascending=False).reset_index(drop=True)


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
        st.caption(f"🔸 {len(filtered):,} ردیف پیدا شد (از {len(view):,}) — {len(terms)} کلمه")
        return filtered
    return view


def show_table(df, key, hide_supervisor=False, search=True, placeholder="جستجو در جدول..."):
    view = df.copy()
    if hide_supervisor and SUP in view.columns:
        view = view.drop(columns=[SUP])
    if search and not view.empty:
        view = filter_by_search(view, key, placeholder)
    if view.empty:
        st.info("ردیفی برای نمایش وجود ندارد.")
        return
    st.dataframe(view, use_container_width=True, hide_index=True,
                 column_config=column_config(view))
    st.download_button("⬇️ دانلود این جدول (CSV)", view.to_csv(index=False).encode("utf-8-sig"),
                       f"{key}.csv", "text/csv", key=f"dl_{key}")


def show_summary(df, key):
    if df.empty:
        st.info("داده‌ای برای نمایش وجود ندارد.")
        return
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=summary_config(df))
    st.download_button("⬇️ دانلود (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                       f"{key}.csv", "text/csv", key=f"dl_{key}")


def export_to_excel(dfs: dict, filename="report.xlsx"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in dfs.items():
            if df is None or df.empty:
                continue
            safe = re.sub(r'[\\/*?:\[\]]', '', str(name))[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
    buf.seek(0)
    return buf.getvalue()
# ================== KPI ها ==================
def render_kpis(f60, f45):
    metric_row([
        ("ارزش راکد ۴۵ روزه", money(f45[VAL].sum())),
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
    metric_row([
        ("💰 ارزش راکد ۶۰ روزه", money(v60)),
        ("📊 اختلاف ارزش ۴۵ و ۶۰ روزه", money(v45 - v60),
         {"delta": f"سهم ۶۰ روزه از ۴۵ روزه: {share:.1f}%", "delta_color": "off"}),
        ("🛒 ارزش راکدِ آزادشده", money(released),
         {"help": "برای هر قلم: min(فروش، موجودی) × قیمت واحد"}),
        ("📦 فروش تعدادی از اقلام راکد", f"{sold:,.0f}"),
    ])


def render_top_products(f60, f45, sales_df, key_suffix="", hide_sup=False):
    top_n = st.slider("چند قلم نمایش داده شود؟", 5, 50, 20, step=5, key=f"slider_{key_suffix}")
    tab1, tab2 = st.tabs(["۶۰ روزه", "۴۵ روزه"])
    with tab1:
        show_table(attach_sales(f60, sales_df).head(top_n), f"top60_{key_suffix}",
                   hide_supervisor=hide_sup, placeholder="نام کالا، بارکد...")
    with tab2:
        show_table(attach_sales(f45, sales_df).head(top_n), f"top45_{key_suffix}",
                   hide_supervisor=hide_sup, placeholder="نام کالا، بارکد...")


def render_matched(f60, f45, sales_df, key_suffix="", hide_sup=False):
    if sales_df is None or sales_df.empty:
        st.info("برای مشاهده این بخش، فایل فروش لازم است.")
        return

    def matched(df):
        m = attach_sales(df, sales_df)
        return m[m[S_QTY] > 0].sort_values(S_AMT, ascending=False).reset_index(drop=True)

    m60, m45 = matched(f60), matched(f45)
    metric_row([("کالای راکد ۶۰ روزه که فروش رفت", f"{len(m60):,}"),
                ("کالای راکد ۴۵ روزه که فروش رفت", f"{len(m45):,}")])
    t1, t2 = st.tabs(["۶۰ روزه — فروش‌رفته", "۴۵ روزه — فروش‌رفته"])
    for tab, m, k, label in ((t1, m60, "m60", "۶۰"), (t2, m45, "m45", "۴۵")):
        with tab:
            if m.empty:
                st.info(f"از اقلام راکد {label} روزه، چیزی فروش نرفت.")
            else:
                show_table(m, f"{k}_{key_suffix}", hide_supervisor=hide_sup,
                           placeholder="نام کالا، بارکد...")


# ================== مرکز اقدام و خلاصه‌ها ==================
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
    x["اولویت اقدام"] = x["_p"].map({0: "🔴 فوری", 1: "🟡 پیگیری", 2: "🟢 عادی"})
    x = x.sort_values(["_p", VAL], ascending=[True, False])
    show_cols = ["اولویت اقدام", BR, NM, BC, QTY, VAL, S_QTY, S_AMT, SUP]
    full = x[show_cols]
    st.caption("🔴 فوری = بدون فروش و جزو ۲۵٪ بالای ارزش راکد | 🟡 پیگیری = بدون فروش | 🟢 عادی = فروش رفته")
    limit = st.slider("تعداد ردیف نمایش", 10, 100, limit, 5, key=f"ac_lim_{key}")
    view = full.head(limit)
    st.dataframe(view, use_container_width=True, hide_index=True, column_config=column_config(view))
    st.download_button("⬇️ دانلود فهرست کامل اولویت‌ها (CSV)",
                       full.to_csv(index=False).encode("utf-8-sig"),
                       f"action_center_{key}.csv", "text/csv", key=f"dl_ac_{key}")


def _finish_summary(out, int_cols):
    for c in int_cols:
        if c in out.columns:
            out[c] = out[c].fillna(0).round(0).astype("int64")
    out["سهم ۶۰ روزه از ۴۵ روزه (٪)"] = np.where(
        out["راکد ۴۵ روزه"] > 0, out["راکد ۶۰ روزه"] / out["راکد ۴۵ روزه"] * 100, 0).round(1)
    return out


@st.cache_data(show_spinner=False, max_entries=16)
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
    out["تعداد قلم"] = out["تعداد قلم"].fillna(0)
    out["اختلاف ۴۵ و ۶۰ روزه"] = out["راکد ۴۵ روزه"] - out["راکد ۶۰ روزه"]
    out["فروش اقلام راکد (تعداد)"] = 0.0
    out["فروش اقلام راکد (ریال)"] = 0.0
    if sales_df is not None and not sales_df.empty:
        matched = attach_sales(df60, sales_df)
        s = matched.groupby(BR)[[S_QTY, S_AMT]].sum()
        out["فروش اقلام راکد (تعداد)"] = s[S_QTY].reindex(out.index).fillna(0)
        out["فروش اقلام راکد (ریال)"] = s[S_AMT].reindex(out.index).fillna(0)
    out.index.name = BR
    out = out.reset_index()
    out = _finish_summary(out, ["راکد ۶۰ روزه", "راکد ۴۵ روزه", "اختلاف ۴۵ و ۶۰ روزه",
                                "فروش اقلام راکد (تعداد)", "فروش اقلام راکد (ریال)"])
    return out.sort_values("راکد ۶۰ روزه", ascending=False).reset_index(drop=True)


@st.cache_data(show_spinner=False, max_entries=16)
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
    out["اختلاف ۴۵ و ۶۰ روزه"] = out["راکد ۴۵ روزه"] - out["راکد ۶۰ روزه"]
    out["فروش اقلام راکد (تعداد)"] = 0.0
    out["فروش اقلام راکد (ریال)"] = 0.0
    if sales_df is not None and not sales_df.empty:
        matched = attach_sales(df60, sales_df)
        s = matched.groupby(SUP)[[S_QTY, S_AMT]].sum()
        out["فروش اقلام راکد (تعداد)"] = s[S_QTY].reindex(out.index).fillna(0)
        out["فروش اقلام راکد (ریال)"] = s[S_AMT].reindex(out.index).fillna(0)
    out.index.name = SUP
    out = out.reset_index()
    out = _finish_summary(out, ["تعداد شعب", "راکد ۶۰ روزه", "راکد ۴۵ روزه",
                                "اختلاف ۴۵ و ۶۰ روزه", "فروش اقلام راکد (تعداد)", "فروش اقلام راکد (ریال)"])
    return out.sort_values("راکد ۶۰ روزه", ascending=False).reset_index(drop=True)


# ================== Cash + ABC + DIO ==================
def render_cash_and_abc(df60, sales_df=None):
    st.markdown("### 💰 پول خوابیده در انبار (Cash Tied Up)")
    rate = st.slider("نرخ هزینه نگهداری ماهانه (٪)", 0.0, 10.0, HOLD_RATE_DEFAULT, 0.5,
                     key="hold_rate",
                     help="هزینه سرمایه، انبارداری، فساد و ... را بر اساس سیاست شرکت تنظیم کن.")
    total = df60[VAL].sum()
    monthly = total * rate / 100
    metric_row([
        ("ارزش پول خوابیده", money(total)),
        (f"هزینه نگهداری ماهانه ({rate:g}٪)", money(monthly)),
        ("هزینه نگهداری سالانه", money(monthly * 12)),
        ("تعداد کالاهای راکد", f"{len(df60):,}"),
    ])

    st.divider()
    st.markdown("### ⏳ تخمین DIO (روزهای ماندن در انبار)")
    if sales_df is not None and not sales_df.empty:
        matched_sales = attach_sales(df60, sales_df)
        daily_sales_value = matched_sales[S_AMT].sum()
        if daily_sales_value > 0:
            dio = total / daily_sales_value
            metric_row([
                ("ارزش راکد", money(total)),
                ("فروش روزانه اقلام راکد", money(daily_sales_value)),
                ("DIO تخمینی", f"{dio:.1f} روز",
                 {"help": "چند روز طول می‌کشد تا ارزش راکد فعلی فروش برود"}),
                ("وضعیت", "🟢 مناسب" if dio < 90 else "🔴 بحرانی"),
            ])
            st.caption("💡 DIO = ارزش راکد ÷ فروش روزانه. هرچه کمتر، بهتر.")
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
    colors = {"A — بحرانی": CHART_PIE_RED, "B — متوسط": CHART_PIE_YELLOW,
              "C — کم‌ارزش": CHART_PIE_GREEN}

    cols = st.columns(2 if is_mobile else 3)
    for i, (_, r) in enumerate(summary.iterrows()):
        with cols[i % len(cols)]:
            st.markdown(
                f'<div class="alert-card" style="border-right-color:{colors.get(r["دسته"], "#666")}">'
                f'<b>{r["دسته"]}</b><br>تعداد: {r["تعداد"]:,} قلم<br>'
                f'ارزش: {r["ارزش"]:,.0f} ریال</div>', unsafe_allow_html=True)

    fig = px.bar(summary, x="دسته", y="ارزش", color="دسته", color_discrete_map=colors,
                 title="توزیع ارزش راکد بر اساس دسته ABC")
    fig.update_layout(showlegend=False, height=350)
    fig = plotly_style(fig)
    fig.update_traces(textfont=dict(color=CHART_TEXT, size=11), textposition="outside",
                      text=summary["تعداد"].apply(lambda v: f"{v} قلم"))
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📋 جدول کامل ABC"):
        st.dataframe(abc, use_container_width=True, hide_index=True)
        st.download_button("⬇️ دانلود ABC (CSV)", abc.to_csv(index=False).encode("utf-8-sig"),
                           "abc.csv", "text/csv", key="dl_abc")


# ================== نمودارها ==================
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
    fig.update_layout(height=430, showlegend=False, coloraxis_showscale=False)
    fig = plotly_style(fig)
    fig.update_traces(textfont=dict(color=CHART_TEXT, size=11), textposition="outside",
                      text=top10[VAL].apply(lambda v: f"{v:,.0f}"))
    st.plotly_chart(fig, use_container_width=True)

    if df_target is not None and not df_target.empty:
        tc = detect_target_columns(df_target)
        ach_col, branch_col = tc.get("achievement"), tc.get("branch")
        if ach_col and branch_col:
            def status(v):
                if v >= ACH_OK:
                    return "🟢 موفق (≥100%)"
                if v >= ACH_WARN:
                    return "🟡 در حال پیشرفت (80-100%)"
                return "🔴 بحرانی (<80%)"

            tmp = df_target[[branch_col, ach_col]].copy()
            tmp["وضعیت"] = tmp[ach_col].apply(status)
            dist = tmp["وضعیت"].value_counts().reset_index()
            dist.columns = ["وضعیت", "تعداد"]
            fig2 = px.pie(dist, values="تعداد", names="وضعیت", title="🥧 توزیع وضعیت شعب",
                          color="وضعیت", hole=0.45,
                          color_discrete_map={"🟢 موفق (≥100%)": CHART_PIE_GREEN,
                                              "🟡 در حال پیشرفت (80-100%)": CHART_PIE_YELLOW,
                                              "🔴 بحرانی (<80%)": CHART_PIE_RED})
            fig2.update_layout(height=400)
            fig2 = plotly_style(fig2)
            fig2.update_traces(textfont=dict(color="#ffffff", size=12),
                               textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig2, use_container_width=True)


# ================== مقایسه و روند تاریخچه ==================
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
                  title="روند ارزش راکد", labels={"total": "ارزش راکد (ریال)"},
                  color_discrete_sequence=[CHART_MAIN, CHART_NEUTRAL])
    fig.update_layout(height=380)
    st.plotly_chart(plotly_style(fig, show_legend_bg=True), use_container_width=True)
    if not s.empty:
        s = s.sort_values("date")
        fig2 = px.bar(s, x="date", y="amount", title="فروش اقلام راکد در هر آپدیت (ریال)",
                      labels={"date": "تاریخ", "amount": "فروش ریالی"})
        fig2.update_traces(marker_color=CHART_MAIN)
        fig2.update_layout(height=340)
        st.plotly_chart(plotly_style(fig2), use_container_width=True)


def render_comparison():
    st.markdown("### 📈 مقایسه با آپدیت قبلی")
    dates = get_history_dates("raaked")
    if len(dates) < 2:
        st.info("برای مقایسه، حداقل ۲ بار باید آپدیت کرده باشی. از این به بعد خودکار ذخیره می‌شود.")
        if len(dates) == 1:
            st.caption(f"📌 آخرین ذخیره: {dates[0]}")
        return

    kind = st.radio("نوع", ["60", "45"], horizontal=True, key="cmp_kind",
                    format_func=lambda k: "۶۰ روزه" if k == "60" else "۴۵ روزه")
    cc = st.columns(2) if is_mobile else st.columns([1, 1, 2])
    date1 = cc[0].selectbox("از تاریخ", dates[1:], index=0, key="cmp_date1")
    date2 = cc[1].selectbox("به تاریخ", dates, index=0, key="cmp_date2")
    if date1 == date2:
        st.warning("دو تاریخ یکسان انتخاب شده.")
        return

    cmp_df = get_history_comparison(date1, date2, kind)
    if cmp_df is None or cmp_df.empty:
        st.info("داده‌ای برای مقایسه وجود ندارد.")
        return

    before, now = cmp_df["total_قبل"].sum(), cmp_df["total_الان"].sum()
    diff = now - before
    pct = (diff / before * 100) if before > 0 else 0
    metric_row([
        ("ارزش قبلی", money(before)),
        ("ارزش فعلی", money(now), {"delta": f"{diff:+,.0f} ریال", "delta_color": "inverse"}),
        ("درصد تغییر", f"{pct:+.2f}%"),
        ("تعداد شعب", f"{len(cmp_df):,}"),
    ])

    st.markdown("#### 🎯 بیشترین افزایش و کاهش")
    gainers = cmp_df.sort_values("تغییر", ascending=False).head(5).copy()
    losers = cmp_df.sort_values("تغییر", ascending=True).head(5).copy()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**🔴 بیشترین افزایش راکد**")
        st.dataframe(
            gainers.rename(columns={"branch": "شعبه", "total_قبل": "قبل",
                                    "total_الان": "الان", "تغییر": "افزایش (ریال)",
                                    "درصد تغییر": "درصد"}),
            use_container_width=True, hide_index=True
        )
    with c2:
        st.markdown("**🟢 بیشترین کاهش راکد**")
        st.dataframe(
            losers.rename(columns={"branch": "شعبه", "total_قبل": "قبل",
                                   "total_الان": "الان", "تغییر": "کاهش (ریال)",
                                   "درصد تغییر": "درصد"}),
            use_container_width=True, hide_index=True
        )

    top = cmp_df.loc[cmp_df["تغییر"].abs().sort_values(ascending=False).index].head(15)
    fig = go.Figure()
    for name, col, color in (("قبل", "total_قبل", CHART_NEUTRAL), ("فعلی", "total_الان", CHART_MAIN)):
        fig.add_trace(go.Bar(name=name, x=top["branch"], y=top[col], marker_color=color,
                             text=top[col].apply(lambda v: f"{v / 1e6:,.0f}M"),
                             textfont=dict(color=CHART_TEXT, size=10), textposition="outside"))
    fig.update_layout(barmode="group", height=470, xaxis_tickangle=-45,
                      title="مقایسه ارزش راکد (۱۵ شعبه با بیشترین تغییر)")
    st.plotly_chart(plotly_style(fig, show_legend_bg=True), use_container_width=True)

    with st.expander("📋 جدول کامل مقایسه"):
        disp = cmp_df.rename(columns={"branch": "شعبه", "total_قبل": "ارزش قبل",
                                      "total_الان": "ارزش فعلی"})
        st.dataframe(disp, use_container_width=True, hide_index=True)


# ================== آپلود ==================
def get_upload_password():
    try:
        pw = st.secrets.get("UPLOAD_PASSWORD")
    except Exception:
        pw = None
    return pw or os.environ.get("UPLOAD_PASSWORD")


def check_upload_access():
    pw = get_upload_password()
    if not pw:
        st.warning("⚠️ رمز آپلود تنظیم نشده و صفحه برای همه باز است. "
                   "کلید UPLOAD_PASSWORD را در `.streamlit/secrets.toml` یا متغیر محیطی بگذار.")
        return True
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
    if flash:
        st.success(flash)

    st.markdown("### 📤 آپلود فایل‌های جدید")
    st.info("فایل قبل از جایگزینی اعتبارسنجی می‌شود، از نسخه قبلی بک‌آپ گرفته می‌شود و "
            "داده در تاریخچه ذخیره می‌شود.")

    picked = st.date_input("📅 تاریخ این آپدیت",
                           value=datetime.now().date(), key="upload_date_pick")
    upd_date = picked.strftime("%Y-%m-%d")
    try:
        import jdatetime
        st.caption(f"🗓️ معادل شمسی: {jdatetime.date.fromgregorian(date=picked).strftime('%Y/%m/%d')}")
    except ImportError:
        st.caption(f"📅 تاریخ میلادی: {upd_date}")

    note = st.text_input("📝 یادداشت (اختیاری)", placeholder="مثلاً: آپدیت هفتگی", key="upload_note")
    st.divider()

    col1, col2 = st.columns(2)
    jobs = [(col1, "📊 فایل راکد", "فایل راکد جدید", "up_raaked", "save_raaked",
             "💾 ذخیره راکد جدید", process_raaked_upload, "فایل راکد"),
            (col2, "🛒 فایل فروش", "فایل فروش جدید", "up_sales", "save_sales",
             "💾 ذخیره فروش جدید", process_sales_upload, "فایل فروش")]
    result = None
    issues = None
    for col, title, label, fkey, bkey, btn, fn, name in jobs:
        with col:
            st.markdown(f"**{title}**")
            f = st.file_uploader(label, type=["xlsx"], key=fkey)
            if f and st.button(btn, use_container_width=True, key=bkey):
                try:
                    out = fn(f.getvalue(), upd_date, note)
                    if isinstance(out, tuple):
                        rows, dq = out
                    else:
                        rows, dq = out, []
                    result = f"✅ {name} ذخیره شد ({rows:,} ردیف)"
                    if dq:
                        issues = dq
                except Exception as e:
                    st.error(f"❌ خطا: {e}")
                    st.caption("💡 فایل قبلی دست‌نخورده ماند. فایل را اصلاح و دوباره آپلود کن.")

    if issues:
        with st.expander("🔍 گزارش کیفیت داده (بررسی کن)", expanded=True):
            for it in issues:
                st.markdown(it)

    if result:
        st.cache_data.clear()
        st.session_state["flash"] = result
        st.rerun()

    st.divider()
    st.markdown("### 🕐 آخرین آپدیت‌ها")
    try:
        with _connect() as conn:
            updates = pd.read_sql(
                "SELECT date, kind, rows, note FROM updates ORDER BY id DESC LIMIT 20", conn)
        if updates.empty:
            st.info("هنوز آپدیتی ذخیره نشده.")
        else:
            updates = updates.rename(columns={"date": "تاریخ", "kind": "نوع",
                                              "rows": "ردیف", "note": "یادداشت"})
            updates["نوع"] = updates["نوع"].replace({"raaked": "📊 راکد", "sales": "🛒 فروش"})
            st.dataframe(updates, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"خطا در خواندن تاریخچه: {e}")

    if HISTORY_DB.exists():
        st.download_button("⬇️ دانلود بک‌آپ دیتابیس تاریخچه", HISTORY_DB.read_bytes(),
                           f"history_{datetime.now():%Y%m%d}.db", "application/octet-stream",
                           key="dl_db")


# ================== تارگت ==================
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
    metric_row([
        ("میانگین درصد تحقق", f"{vals.mean():.2f}%"),
        ("شعب موفق (≥100%)", f"{(vals >= ACH_OK).sum():,}"),
        ("شعب بحرانی (<80%)", f"{(vals < ACH_WARN).sum():,}"),
        ("کل شعب", f"{len(t_df):,}"),
    ])


def render_target_table(t_df, key_suffix=""):
    cm = detect_target_columns(t_df)
    display_cols = [cm[k] for k in ("branch", "supervisor", "raked_value", "target",
                                    "achievement", "change") if cm.get(k)]
    if not display_cols:
        st.warning("ستون‌های فایل تارگت قابل شناسایی نبودند.")
        return
    view = t_df[display_cols].copy()
    if cm.get("achievement"):
        view = view.sort_values(cm["achievement"], ascending=False).reset_index(drop=True)
        view.insert(0, "رتبه", range(1, len(view) + 1))
    st.markdown("**📋 جدول تارگت و درصد تحقق**")
    q = st.text_input("🔍 جستجو در جدول تارگت", key=f"search_target_{key_suffix}",
                      placeholder="نام شعبه یا سرپرست")
    if q and q.strip():
        qn = normalize_query(q)
        mask = view.astype(str).apply(
            lambda col: col.str.lower().str.contains(qn, na=False, regex=False)).any(axis=1)
        view = view[mask]
        st.caption(f"🔸 {len(view):,} ردیف پیدا شد")
    if view.empty:
        st.info("ردیفی پیدا نشد.")
        return
    st.dataframe(view, use_container_width=True, hide_index=True,
                 column_config=target_column_config(view))
    st.download_button("⬇️ دانلود جدول تارگت (CSV)", view.to_csv(index=False).encode("utf-8-sig"),
                       f"target_{key_suffix}.csv", "text/csv", key=f"dl_target_{key_suffix}")


def render_target_trend(t_df, key_suffix=""):
    cm = detect_target_columns(t_df)
    trend_cols, branch_col = cm.get("trend") or [], cm.get("branch")
    if not trend_cols or not branch_col:
        st.info("ستون‌های روند پیدا نشدند.")
        return
    chart_data = t_df[[branch_col] + trend_cols].copy()
    names = [c.replace("درصد راکد", "").replace("درصد  راکد", "").strip() for c in trend_cols]
    chart_data = chart_data.rename(columns=dict(zip(trend_cols, names))).set_index(branch_col)
    selected = st.multiselect("🔎 انتخاب شعب برای نمودار (خالی = همه)",
                              options=chart_data.index.tolist(), default=[],
                              key=f"chart_branches_{key_suffix}")
    if selected:
        chart_data = chart_data.loc[selected]
    if chart_data.empty:
        st.info("شعبه‌ای برای نمایش انتخاب نشده.")
        return
    fig = go.Figure()
    for col in chart_data.columns:
        fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data[col],
                                 mode="lines+markers", name=str(col)))
    fig.update_layout(title="روند درصد راکد شعب", height=500,
                      xaxis_title="شعبه", yaxis_title="درصد راکد (%)")
    fig = plotly_style(fig, show_legend_bg=True)
    fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=-0.4))
    st.plotly_chart(fig, use_container_width=True)
    with st.expander("📊 جدول روند"):
        st.dataframe(chart_data, use_container_width=True)


def render_ranking(t_df, key_suffix=""):
    cm = detect_target_columns(t_df)
    branch_col, sup_col, ach_col = cm.get("branch"), cm.get("supervisor"), cm.get("achievement")
    change_col, raked_col = cm.get("change"), cm.get("raked_value")
    if not branch_col or not ach_col:
        st.info("برای رتبه‌بندی، نام شعبه و درصد تحقق لازم است.")
        return

    def status_emoji(v):
        if v >= ACH_OK:
            return "🟢 موفق"
        if v >= ACH_WARN:
            return "🟡 در حال پیشرفت"
        return "🔴 بحرانی"

    rank_df = t_df[[c for c in [branch_col, sup_col, raked_col, ach_col, change_col] if c]].copy()
    rank_df["وضعیت"] = rank_df[ach_col].apply(status_emoji)
    rank_df = rank_df.sort_values(ach_col, ascending=False).reset_index(drop=True)
    rank_df.insert(0, "رتبه", range(1, len(rank_df) + 1))
    n = len(rank_df)
    if n == 0:
        st.info("داده‌ای برای رتبه‌بندی وجود ندارد.")
        return
    if n == 1:
        st.markdown("**📌 رتبه‌بندی (تنها یک شعبه)**")
        st.dataframe(rank_df, use_container_width=True, hide_index=True,
                     column_config=target_column_config(rank_df))
        return

    st.markdown("**🥇 بهترین ۱۰ شعبه**")
    best = rank_df.head(min(10, n)).copy()
    st.dataframe(best, use_container_width=True, hide_index=True,
                 column_config=target_column_config(best))
    st.markdown("**⚠️ بدترین ۱۰ شعبه (نیازمند اقدام فوری)**")
    worst = rank_df.tail(min(10, n)).sort_values(ach_col).reset_index(drop=True)
    worst["رتبه"] = range(n - len(worst) + 1, n + 1)
    st.dataframe(worst, use_container_width=True, hide_index=True,
                 column_config=target_column_config(worst))

    if sup_col:
        st.markdown("**👤 رتبه‌بندی سرپرست‌ها**")
        agg = {ach_col: "mean"}
        if raked_col:
            agg[raked_col] = "sum"
        agg[branch_col] = "count"
        sup_rank = (t_df.groupby(sup_col).agg(agg).reset_index()
                    .rename(columns={branch_col: "تعداد شعبه", ach_col: "میانگین تحقق"})
                    .sort_values("میانگین تحقق", ascending=False).reset_index(drop=True))
        sup_rank.insert(0, "رتبه", range(1, len(sup_rank) + 1))
        cfg = {}
        for c in sup_rank.columns:
            if "تحقق" in c:
                cfg[c] = st.column_config.NumberColumn(c, format="%.2f%%",
                                                       width=calc_width(sup_rank, c, is_percent=True))
            elif raked_col and c == raked_col:
                cfg[c] = st.column_config.NumberColumn(c, format="%,d",
                                                       width=calc_width(sup_rank, c, is_numeric=True))
            else:
                cfg[c] = st.column_config.Column(c, width=calc_width(sup_rank, c))
        st.dataframe(sup_rank, use_container_width=True, hide_index=True, column_config=cfg)


def render_weekly_target(t_df):
    """داشبورد مدیریتی تارگت هفتگی بر پایه ستون‌های تارگت و تحقق فایل هدف."""
    if t_df is None or t_df.empty:
        st.info("فایل تارگت برای محاسبه تارگت هفتگی موجود نیست.")
        return
    cm = detect_target_columns(t_df)
    target_col, ach_col = cm.get("target"), cm.get("achievement")
    if not target_col or not ach_col:
        st.warning("برای تارگت هفتگی، ستون‌های «تارگت» و «تحقق» باید در فایل تارگت وجود داشته باشند.")
        return
    d = t_df[[c for c in [cm.get("branch"), target_col, ach_col] if c]].copy()
    d[target_col] = to_number(d[target_col])
    d[ach_col] = to_number(d[ach_col])
    d = d[(d[target_col] > 0) | (d[ach_col] > 0)]
    if d.empty:
        st.info("مقدار قابل محاسبه‌ای برای تارگت هفتگی وجود ندارد.")
        return
    target = d[target_col].mean()
    achieved = d[ach_col].mean()
    gap = achieved - target
    remaining = max(target - achieved, 0)
    elapsed = max(datetime.now().weekday() + 1, 1)
    forecast = achieved / elapsed * 7
    metric_row([
        ("🎯 تارگت هفتگی", f"{target:.1f}%"),
        ("✅ تحقق فعلی", f"{achieved:.1f}%",
         {"delta": f"{gap:+.1f}% نسبت به تارگت", "delta_color": "normal"}),
        ("📉 فاصله تا تارگت", f"{remaining:.1f}%"),
        ("🔮 پیش‌بینی پایان هفته", f"{forecast:.1f}%",
         {"delta": "بر اساس سرعت فعلی", "delta_color": "off"}),
    ])
    status = ("🟢 روی تارگت" if achieved >= target
              else ("🟡 قابل جبران" if forecast >= target else "🔴 عقب از برنامه"))
    cls = "ok" if achieved >= target else "warn"
    st.markdown(f'<div class="alert-card {cls}"><b>وضعیت تارگت هفتگی:</b> {status}</div>',
                unsafe_allow_html=True)


def render_target(t_df, key_suffix=""):
    if t_df is None:
        st.warning("⚠️ فایل تارگت بارگذاری نشده یا خالی است.")
        return
    if t_df.empty:
        st.info("داده‌ای برای این فیلتر وجود ندارد.")
        return
    render_target_kpis(t_df)
    st.divider()
    tab1, tab2, tab3 = st.tabs(["📋 جدول تارگت", "📈 روند", "🏆 رتبه‌بندی"])
    with tab1:
        render_target_table(t_df, key_suffix)
    with tab2:
        render_target_trend(t_df, key_suffix)
    with tab3:
        render_ranking(t_df, key_suffix)


# ================== مسیریابی ==================
try:
    page = st.query_params.get("page", "home")
except Exception:
    page = st.experimental_get_query_params().get("page", ["home"])[0]
if page not in VALID_PAGES:
    page = "home"

# ================== خواندن داده ==================
data_error = None
try:
    df60, df45 = load_raaked(_mtime(RAAKED_FILE))
except Exception as e:
    data_error = str(e)
    empty = pd.DataFrame(columns=[BR, BC, NM, QTY, VAL, SUP])
    df60, df45 = empty.copy(), empty.copy()
df_sales = load_sales(_mtime(SALES_FILE))
df_target = load_target(_mtime(TARGET_FILE))

if data_error and page != "upload":
    render_header("خطا در بارگذاری داده")
    st.error(f"فایل راکد خوانده نشد: {data_error}")
    st.markdown('<a class="back-link" href="?page=upload" target="_self">📤 رفتن به صفحه آپلود</a>',
                unsafe_allow_html=True)
    st.stop()


def show_last_update_badge():
    last = get_last_update()
    if last and last[0]:
        date, kind, rows = last
        label = "راکد" if kind == "raaked" else "فروش"
        st.markdown(f'<div class="alert-card ok">📅 <b>آخرین آپدیت:</b> {date} | نوع: {label} | '
                    f'{(rows or 0):,} ردیف</div>', unsafe_allow_html=True)


def card_html(target, icon, title, sub, light):
    cls = "home-card light" if light else "home-card"
    return (f'<a class="card-link" href="?page={target}" target="_self" style="flex:1;">'
            f'<div class="{cls}"><div class="icon">{icon}</div><h2>{title}</h2><p>{sub}</p></div></a>')


CARDS = [
    ("district", "📊", "عملکرد کلی دیستریکت", "نمای کلی و KPIهای کل", False),
    ("store", "🏪", "عملکرد فروشگاه‌ها", "تحلیل هر شعبه جداگانه", True),
    ("supervisor", "👤", "عملکرد سوپروایزرها", "عملکرد هر سرپرست", True),
    ("target", "🎯", "تارگت و روند", "اهداف و رتبه‌بندی شعب", False),
    ("analytics", "💰", "تحلیل پیشرفته", "ABC، پول خوابیده، نمودارها", False),
    ("upload", "📤", "آپلود و تاریخچه", "فایل جدید + مقایسه با گذشته", True),
]

# ================== صفحه ورودی ==================
if page == "home":
    render_header("فروشگاه‌های زنجیره‌ای افق کوروش")
    show_last_update_badge()
    st.markdown("### 📌 خلاصه مدیریتی امروز")
    render_management_kpis(df60, df45, df_sales)
    render_weekly_target(df_target)
    st.divider()
    st.markdown("### لطفاً یکی از گزینه‌ها را انتخاب کنید:")
    st.divider()
    if is_mobile:
        for i in range(0, len(CARDS), 2):
            for col, c in zip(st.columns(2), CARDS[i:i + 2]):
                col.markdown(card_html(*c).replace(' style="flex:1;"', ""), unsafe_allow_html=True)
    else:
        for i in range(0, len(CARDS), 3):
            row = "".join(card_html(*c) for c in CARDS[i:i + 3])
            st.markdown(f'<div style="display:flex; gap:12px; align-items:stretch; '
                        f'margin-bottom:12px;">{row}</div>', unsafe_allow_html=True)
    st.markdown('<div class="footer-text">ساخته شده توسط <b>شاهین باقری</b></div>',
                unsafe_allow_html=True)

# ================== دیستریکت ==================
elif page == "district":
    back_link()
    render_header("عملکرد کلی دیستریکت")
    show_last_update_badge()
    st.divider()
    render_management_kpis(df60, df45, df_sales)
    render_kpis(df60, df45)
    st.divider()
    st.subheader("🎯 تارگت هفتگی دیستریکت")
    render_weekly_target(df_target)
    st.divider()
    render_action_center(df60, df_sales, key="district")
    st.divider()
    t1, t2 = st.tabs(["🏪 وضعیت شعب", "👤 وضعیت سوپروایزرها"])
    with t1:
        show_summary(branch_summary(df60, df45, df_sales), "branches_summary")
    with t2:
        show_summary(supervisor_summary(df60, df45, df_sales), "supervisors_summary")
    st.divider()
    if st.button("📥 خروجی اکسل کامل (همه شیت‌ها)", key="export_district"):
        with st.spinner("در حال ساخت فایل..."):
            xls_bytes = export_to_excel({
                "خلاصه شعب": branch_summary(df60, df45, df_sales),
                "خلاصه سرپرست‌ها": supervisor_summary(df60, df45, df_sales),
                "راکد ۶۰ روزه": attach_sales(df60, df_sales),
                "راکد ۴۵ روزه": attach_sales(df45, df_sales),
                "مرکز اقدام": attach_sales(df60, df_sales).sort_values(VAL, ascending=False),
                "تارگت": df_target,
            }, f"district_{datetime.now():%Y%m%d}.xlsx")
        st.download_button(
            "⬇️ دانلود فایل اکسل",
            xls_bytes,
            f"district_{datetime.now():%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_district_xls"
        )
    st.divider()
    st.subheader("🚨 کالاهای راکد برتر دیستریکت")
    render_top_products(df60, df45, df_sales, key_suffix="district")
    st.divider()
    st.subheader("✅ کالاهای راکدی که فروش رفتند (کل دیستریکت)")
    render_matched(df60, df45, df_sales, key_suffix="district")

# ================== فروشگاه ==================
elif page == "store":
    back_link()
    render_header("عملکرد فروشگاه‌ها")
    branch_list = sorted(df60[BR].dropna().unique().tolist())
    if not branch_list:
        st.info("شعبه‌ای در داده وجود ندارد.")
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
    tabs = st.tabs(["📦 اقلام راکد", "🎯 تارگت و روند"])
    with tabs[0]:
        st.subheader("🚨 کالاهای راکد برتر این فروشگاه")
        render_top_products(b60, b45, branch_sales,
                            key_suffix=f"branch_{selected_branch}", hide_sup=True)
        st.divider()
        st.subheader("✅ کالاهای راکدی که فروش رفتند (این فروشگاه)")
        render_matched(b60, b45, branch_sales,
                       key_suffix=f"branch_m_{selected_branch}", hide_sup=True)
    with tabs[1]:
        render_target(branch_target, key_suffix=f"branch_{selected_branch}")

# ================== سوپروایزر ==================
elif page == "supervisor":
    back_link()
    render_header("عملکرد سوپروایزرها")
    supervisor_list = sorted(df60[SUP].dropna().unique().tolist())
    if not supervisor_list:
        st.info("سوپروایزری در داده وجود ندارد.")
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
    tabs = st.tabs(["📦 اقلام راکد", "🎯 تارگت و روند"])
    with tabs[0]:
        st.subheader("🚨 کالاهای راکد برتر این سوپروایزر")
        render_top_products(s60, s45, sup_sales, key_suffix=f"sup_{selected_sup}", hide_sup=True)
        st.divider()
        st.subheader("✅ کالاهای راکدی که فروش رفتند (این سوپروایزر)")
        render_matched(s60, s45, sup_sales, key_suffix=f"sup_m_{selected_sup}", hide_sup=True)
    with tabs[1]:
        render_target(sup_target, key_suffix=f"sup_{selected_sup}")

# ================== تارگت ==================
elif page == "target":
    back_link()
    render_header("تارگت و روند کل دیستریکت")
    render_target(df_target, key_suffix="district")

# ================== تحلیل پیشرفته ==================
elif page == "analytics":
    back_link()
    render_header("تحلیل پیشرفته")
    tab1, tab2, tab3 = st.tabs(["💰 پول خوابیده و ABC", "📊 نمودارها", "📈 مقایسه با گذشته"])
    with tab1:
        render_cash_and_abc(df60, df_sales)
    with tab2:
        render_charts(df60, df_target)
    with tab3:
        render_comparison()
        st.divider()
        render_history_trend()
    st.divider()
    render_action_center(df60, df_sales, key="analytics")

# ================== آپلود ==================
elif page == "upload":
    back_link()
    render_header("آپلود فایل و تاریخچه")
    render_upload()