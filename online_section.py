"""
بخش «فروش اینترنتی و مغایرت‌گیری»
- لود دو فایل اکسل: 01_internet_sales.xlsx  و  مغایرت گیری - ادجاست.xlsx
- جدول‌های آماده برای گزارش تصویری (از طریق report_sender.py → _load_source_data)
- صفحه‌ی جدا در داشبورد: آپلود، پیش‌نمایش، رفتن به ارسال تصویری
"""
import os
import re
from functools import lru_cache
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent

SALES_NAME = "01_internet_sales.xlsx"
ADJUST_NAME = "مغایرت گیری - ادجاست.xlsx"

_SALES_CANDIDATES = [SALES_NAME]
_ADJUST_CANDIDATES = [ADJUST_NAME, "مغایرت_گیری_-_ادجاست.xlsx", "مغایرت گیری ادجاست.xlsx"]

# کلید منبع → عنوان (برای منوی ارسال تصویری)
ONLINE_SOURCES = {
    "🌐 فروش اینترنتی — عملکرد شعب": "online_perf",
    "🌐 فروش اینترنتی — مقایسه هفتگی": "online_compare",
    "📋 مغایرت‌گیری انبار — درصد انجام": "adjust_count",
    "⚖️ ادجاست — خلاصه شعب (میلیون ریال)": "adjust_summary",
    "⚖️ ادجاست — ۳۰ قلم برتر منفی (میلیون ریال)": "adjust_top",
}


# ═══════════════════════ پیدا کردن فایل‌ها ═══════════════════════
def _roots():
    roots = []
    env_dir = os.environ.get("DATA_DIR")
    if env_dir:
        roots.append(Path(env_dir))
    roots += [BASE_DIR, Path.cwd(), Path.home() / ".kalaraked", Path("/tmp")]
    return roots


def find_file(kind):
    """kind: 'sales' یا 'adjust' — مسیر فایل یا None"""
    names = _SALES_CANDIDATES if kind == "sales" else _ADJUST_CANDIDATES
    pattern = "*internet_sales*.xlsx" if kind == "sales" else "*ادجاست*.xlsx"
    for root in _roots():
        for n in names:
            f = root / n
            if f.exists():
                return f
        try:
            for f in sorted(root.glob(pattern)):
                if not f.name.startswith("~$"):
                    return f
        except Exception:
            continue
    return None


# ═══════════════════════ ابزارهای کمکی ═══════════════════════
def _norm(s):
    s = str(s if s is not None else "")
    s = s.replace("ي", "ی").replace("ك", "ک").replace("\u200c", " ")
    return re.sub(r"\s+", " ", s).strip()


def _norm_sup(s):
    """نام سوپروایزر: حذف فاصله‌ی اضافه و رقم انتهایی (مثل «جواد رستمی2»)"""
    s = _norm(s)
    if s.lower() == "nan":
        return ""
    return re.sub(r"\d+$", "", s).strip()


def _sheet(xls, key):
    for s in xls.sheet_names:
        if key in _norm(s):
            return s
    return None


def _num(series):
    return pd.to_numeric(series, errors="coerce")


def _pct(v):
    try:
        n = float(v)
        if n != n:
            return "—"
        return f"{n * 100:.2f}%"
    except (TypeError, ValueError):
        return "—"


def _int(v):
    try:
        n = float(v)
        if n != n:
            return "—"
        return f"{int(round(n)):,}"
    except (TypeError, ValueError):
        return "—"


def _mil(v):
    """ریال → میلیون ریال با یک رقم اعشار"""
    try:
        n = float(v)
        if n != n:
            return "—"
        return f"{n / 1e6:,.1f}"
    except (TypeError, ValueError):
        return "—"


_TREND = {"افزایشی": "▲ افزایش", "کاهشی": "▼ کاهش", "بدون تغییر": "بدون تغییر"}


def _mtime(p):
    try:
        return Path(p).stat().st_mtime
    except Exception:
        return 0


# ═══════════════════════ فروش اینترنتی ═══════════════════════
def _read_perf_raw(path):
    xls = pd.ExcelFile(path)
    sh = _sheet(xls, "عملکرد")
    if not sh:
        return None, None
    raw = pd.read_excel(xls, sheet_name=sh, header=None)
    hr = None
    for i in range(len(raw)):
        if raw.iloc[i].astype(str).str.strip().eq("TarikhShamsi").any():
            hr = i
            break
    if hr is None:
        return None, None
    heads = [str(c).strip() for c in raw.iloc[hr]]
    c0 = heads.index("TarikhShamsi")
    d = raw.iloc[hr + 1:, c0:].copy()
    d.columns = [_norm(c) for c in heads[c0:]]
    d = d.dropna(how="all").reset_index(drop=True)
    total = d[d["TarikhShamsi"].astype(str).str.contains("Total", case=False, na=False)]
    d = d[~d["TarikhShamsi"].astype(str).str.contains("Total", case=False, na=False)]
    d = d[d["نام شعبه"].notna()].reset_index(drop=True)
    return d, (total.iloc[0] if len(total) else None)


@lru_cache(maxsize=8)
def _perf_cached(path, mt):
    d, total = _read_perf_raw(path)
    return d, total


def load_online_perf(path=None):
    path = path or find_file("sales")
    if not path:
        return None
    d, _ = _perf_cached(str(path), _mtime(path))
    if d is None or d.empty:
        return None
    d = d.copy()
    for c in ("ثبتی", "تحویلی", "تاخیر", "درصد تاخیر", "ناموجودی", "درصد ناموجودی", "مرجوعی"):
        if c in d.columns:
            d[c] = _num(d[c])
    d["سوپروایزر"] = d["سوپروایزر"].map(_norm_sup)
    d["نام شعبه"] = d["نام شعبه"].map(_norm)
    d = d.sort_values("درصد تاخیر", ascending=False).reset_index(drop=True)
    out = pd.DataFrame({
        "نام شعبه": d["نام شعبه"],
        "سوپروایزر": d["سوپروایزر"],
        "ثبتی": d["ثبتی"].map(_int),
        "تحویلی": d["تحویلی"].map(_int),
        "تاخیر": d["تاخیر"].map(_int),
        "درصد تاخیر": d["درصد تاخیر"].map(_pct),
        "ناموجودی": d["ناموجودی"].map(_int),
        "درصد ناموجودی": d["درصد ناموجودی"].map(_pct),
        "مرجوعی": d["مرجوعی"].map(_int),
    })
    return out


def online_perf_summary(path=None):
    """جمع کل و تاریخ گزارش برای نمایش در صفحه"""
    path = path or find_file("sales")
    if not path:
        return None
    d, total = _perf_cached(str(path), _mtime(path))
    if d is None or d.empty:
        return None
    dates = d["TarikhShamsi"].astype(str).str.replace(r"\.0$", "", regex=True).unique().tolist()
    s = {"date": dates[0] if len(dates) == 1 else f"{min(dates)} تا {max(dates)}"}
    if total is not None:
        for k in ("ثبتی", "تحویلی", "تاخیر", "درصد تاخیر", "ناموجودی", "مرجوعی"):
            s[k] = total.get(k)
    return s


@lru_cache(maxsize=8)
def _compare_cached(path, mt):
    xls = pd.ExcelFile(path)
    sh = _sheet(xls, "مقایسه")
    if not sh:
        return None, ("", "")
    raw = pd.read_excel(xls, sheet_name=sh, header=None)
    if raw.shape[1] < 14 or raw.shape[0] < 3:
        return None, ("", "")
    titles = (_norm(raw.iloc[0, 0]), _norm(raw.iloc[0, 6]))
    d = raw.iloc[2:, :14].copy().reset_index(drop=True)
    d = d[d[2].notna()]
    return d, titles


def load_online_compare(path=None):
    path = path or find_file("sales")
    if not path:
        return None
    d, _ = _compare_cached(str(path), _mtime(path))
    if d is None or d.empty:
        return None
    late_new = _num(d[8])
    order = late_new.fillna(-1).sort_values(ascending=False).index
    d = d.loc[order]

    def pc(col):
        return _num(d[col]).map(_pct)

    def tr(col):
        return d[col].map(lambda v: _TREND.get(_norm(v), _norm(v) or "—"))

    out = pd.DataFrame({
        "نام شعبه": d[2].map(_norm),
        "سوپروایزر": d[0].map(_norm_sup),
        "تاخیر قبل": pc(3),
        "تاخیر جدید": pc(8),
        "روند تاخیر": tr(11),
        "ناموجودی قبل": pc(4),
        "ناموجودی جدید": pc(9),
        "روند ناموجودی": tr(12),
        "استرداد قبل": pc(5),
        "استرداد جدید": pc(10),
        "روند استرداد": tr(13),
    })
    return out.reset_index(drop=True)


def online_compare_titles(path=None):
    path = path or find_file("sales")
    if not path:
        return ("", "")
    return _compare_cached(str(path), _mtime(path))[1]


# ═══════════════════════ مغایرت‌گیری و ادجاست ═══════════════════════
def load_adjust_count(path=None):
    path = path or find_file("adjust")
    if not path:
        return None
    xls = pd.ExcelFile(path)
    sh = _sheet(xls, "مغایرت")
    if not sh:
        return None
    d = pd.read_excel(xls, sheet_name=sh)
    d.columns = [_norm(c) for c in d.columns]
    need = ["نام فروشگاه", "Discrepancy ID", "Fact Diff Inventory Count", "Is Done", "درصد مغایرت گیری"]
    if any(c not in d.columns for c in need):
        return None
    d = d[d["نام فروشگاه"].notna()].copy()
    d["_p"] = _num(d["درصد مغایرت گیری"])
    d = d.sort_values("_p", ascending=True).reset_index(drop=True)
    sup = d["سوپروایزر"].map(_norm_sup) if "سوپروایزر" in d.columns else ""
    return pd.DataFrame({
        "نام شعبه": d["نام فروشگاه"].map(_norm),
        "سوپروایزر": sup,
        "تعداد مغایرت": d["Discrepancy ID"].map(_int),
        "کل شمارش": d["Fact Diff Inventory Count"].map(_int),
        "انجام‌شده": d["Is Done"].map(_int),
        "درصد انجام": d["_p"].map(_pct),
    })


@lru_cache(maxsize=8)
def _journal_cached(path, mt):
    xls = pd.ExcelFile(path)
    sh = _sheet(xls, "ادجاست")
    if not sh:
        return None
    d = pd.read_excel(xls, sheet_name=sh)
    if d.shape[1] < 11:
        return None
    d = d.iloc[:, :11].copy()
    d.columns = ["journal", "date", "loc", "store", "item_id", "item",
                 "qty_pos", "qty_neg", "rial_pos", "rial_neg", "sup"]
    for c in ("qty_pos", "qty_neg", "rial_pos", "rial_neg"):
        d[c] = _num(d[c]).fillna(0)
    d["store"] = d["store"].map(_norm)
    d["item"] = d["item"].map(_norm)
    d["sup"] = d["sup"].map(_norm_sup)
    d["date"] = d["date"].astype(str).str.strip()
    return d[d["store"] != ""].reset_index(drop=True)


def _journal(path=None):
    path = path or find_file("adjust")
    if not path:
        return None
    return _journal_cached(str(path), _mtime(path))


def adjust_period(path=None):
    d = _journal(path)
    if d is None or d.empty:
        return ""
    return f"{d['date'].min()} تا {d['date'].max()}"


def load_adjust_summary(path=None):
    d = _journal(path)
    if d is None or d.empty:
        return None
    g = d.groupby("store", as_index=False).agg(
        sup=("sup", "first"), items=("item_id", "count"),
        qp=("qty_pos", "sum"), qn=("qty_neg", "sum"),
        rp=("rial_pos", "sum"), rn=("rial_neg", "sum"))
    g["net"] = g["rp"] - g["rn"]
    g = g.sort_values("rn", ascending=False).reset_index(drop=True)
    return pd.DataFrame({
        "نام شعبه": g["store"],
        "سوپروایزر": g["sup"],
        "تعداد اقلام": g["items"].map(_int),
        "تعداد مثبت": g["qp"].map(_int),
        "تعداد منفی": g["qn"].map(_int),
        "مبلغ مثبت (میلیون)": g["rp"].map(_mil),
        "مبلغ منفی (میلیون)": g["rn"].map(_mil),
        "خالص (میلیون)": g["net"].map(_mil),
    })


def load_adjust_top(path=None, n=30):
    d = _journal(path)
    if d is None or d.empty:
        return None
    t = d.sort_values("rial_neg", ascending=False).head(n).reset_index(drop=True)
    return pd.DataFrame({
        "نام شعبه": t["store"],
        "سوپروایزر": t["sup"],
        "نام کالا": t["item"],
        "تاریخ": t["date"],
        "تعداد منفی": t["qty_neg"].map(_int),
        "مبلغ منفی (میلیون)": t["rial_neg"].map(_mil),
    })


# ═══════════════════════ اتصال به ارسال تصویری ═══════════════════════
_LOADERS = {
    "online_perf": load_online_perf,
    "online_compare": load_online_compare,
    "adjust_count": load_adjust_count,
    "adjust_summary": load_adjust_summary,
    "adjust_top": load_adjust_top,
}


def load_online_source(source_key):
    """برای report_sender._load_source_data — خطا نمی‌دهد، در شکست None برمی‌گرداند"""
    fn = _LOADERS.get(source_key)
    if not fn:
        return None
    try:
        return fn()
    except Exception:
        return None


# ═══════════════════════ صفحه‌ی جدا در داشبورد ═══════════════════════
def _save_upload(data_dir, name, data):
    data_dir = Path(data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    target = data_dir / name
    tmp = Path(str(target) + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, target)
    for fn in (_perf_cached, _compare_cached, _journal_cached):
        fn.cache_clear()
    return target


def render_online_page(data_dir):
    import streamlit as st

    st.markdown("### 🌐 فروش اینترنتی و مغایرت‌گیری")
    st.caption("دو فایل اکسل را اینجا بارگذاری کنید؛ جدول‌ها خودکار ساخته می‌شوند "
               "و در «ارسال گزارش ← 📸 گزارش تصویری» قابل ارسال به بله هستند.")

    sales_f, adjust_f = find_file("sales"), find_file("adjust")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**📦 فایل فروش اینترنتی**")
        st.success(f"✅ {sales_f.name}") if sales_f else st.warning("⚠️ هنوز آپلود نشده")
        up1 = st.file_uploader("01_internet_sales.xlsx", type=["xlsx"], key="online_up_sales")
        if up1 is not None and st.button("💾 ذخیره فایل فروش اینترنتی", key="online_save_sales",
                                         use_container_width=True, type="primary"):
            try:
                _save_upload(data_dir, SALES_NAME, up1.getvalue())
                st.success("✅ ذخیره شد")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        st.markdown("**⚖️ فایل مغایرت‌گیری و ادجاست**")
        st.success(f"✅ {adjust_f.name}") if adjust_f else st.warning("⚠️ هنوز آپلود نشده")
        up2 = st.file_uploader("مغایرت گیری - ادجاست.xlsx", type=["xlsx"], key="online_up_adjust")
        if up2 is not None and st.button("💾 ذخیره فایل مغایرت‌گیری", key="online_save_adjust",
                                         use_container_width=True, type="primary"):
            try:
                _save_upload(data_dir, ADJUST_NAME, up2.getvalue())
                st.success("✅ ذخیره شد")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    if not (sales_f or adjust_f):
        return

    st.divider()
    if st.button("📸 رفتن به ارسال گزارش تصویری", key="online_goto_report",
                 type="primary", use_container_width=True):
        st.session_state["page"] = "report"
        try:
            st.query_params["p"] = "report"
        except Exception:
            pass
        st.rerun()

    tabs = st.tabs(["🌐 عملکرد اینترنتی", "🔄 مقایسه هفتگی", "📋 مغایرت‌گیری", "⚖️ ادجاست"])

    with tabs[0]:
        s = online_perf_summary()
        if s:
            st.caption(f"📅 تاریخ گزارش: {s['date']}")
            m = st.columns(4)
            m[0].metric("ثبتی", _int(s.get("ثبتی")))
            m[1].metric("تحویلی", _int(s.get("تحویلی")))
            m[2].metric("درصد تاخیر", _pct(s.get("درصد تاخیر")))
            m[3].metric("مرجوعی", _int(s.get("مرجوعی")))
        df = load_online_source("online_perf")
        if df is None:
            st.info("برای این بخش فایل فروش اینترنتی لازم است.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tabs[1]:
        t = online_compare_titles()
        if any(t):
            st.caption(f"هفته قبل: {t[0]}  |  هفته جدید: {t[1]}")
        df = load_online_source("online_compare")
        if df is None:
            st.info("شیت «مقایسه‌ای» پیدا نشد.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tabs[2]:
        df = load_online_source("adjust_count")
        if df is None:
            st.info("برای این بخش فایل مغایرت‌گیری لازم است.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tabs[3]:
        p = adjust_period()
        if p:
            st.caption(f"📅 بازه‌ی ادجاست: {p}")
        df = load_online_source("adjust_summary")
        if df is None:
            st.info("شیت «ادجاست» پیدا نشد.")
        else:
            st.markdown("**خلاصه‌ی شعب**")
            st.dataframe(df, use_container_width=True, hide_index=True)
            top = load_online_source("adjust_top")
            if top is not None:
                st.markdown("**۳۰ قلم برتر منفی**")
                st.dataframe(top, use_container_width=True, hide_index=True)
