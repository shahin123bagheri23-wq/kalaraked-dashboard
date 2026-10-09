from pathlib import Path
import threading
import time as _time_mod


# ================== خروجی فروش ==================
# تبدیل ارقام فارسی/عربی به انگلیسی
FA_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

def load_sold_report(mtime=0):
    """خوندن شیت خروجی — گزارش فروش پروژه"""
    import pandas as pd
    from send_report import EXCEL_FILE, find_sheet, normalize_name

    if not Path(EXCEL_FILE).exists():
        return None
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sh = find_sheet(xls, "خروجی")
        if not sh:
            for n in xls.sheet_names:
                if "خروجی" in str(n):
                    sh = n
                    break
        if not sh:
            return None
        d = pd.read_excel(xls, sheet_name=sh)
        d.columns = d.columns.astype(str).str.strip()
        for c in d.columns:
            if "نام شعبه" in c or "نام کالا" in c or "سوپروایزر" in c:
                d[c] = normalize_name(d[c])
            if "شیفت" in c or "وضعیت فروش" in c:
                d[c] = d[c].astype(str).str.strip()
        needed = ["نام شعبه", "بارکد", "نام کالا", "شیفت", "وضعیت فروش"]
        available = []
        for n in needed:
            for c in d.columns:
                if n == c or n in c:
                    available.append(c)
                    break
        if len(available) < 3:
            return None
        return d[available].copy()
    except Exception:
        return None


"""صفحه ارسال گزارش — همه کانال‌ها در یک جا"""
import hmac
import os
from datetime import datetime

import pandas as pd
import streamlit as st

from email_manager import load_contacts, save_contacts, parse_contacts_file
from email_section import (
    _build_manager_report,
    _build_regular_report,
    _build_supervisor_report,
    _send_emails,
    TYPE_LABELS,
    LABEL_TO_TYPE,
)


# ================== رمز مدیریت ==================
def _get_admin_password():
    try:
        pw = st.secrets.get("ADMIN_PASSWORD")
        if pw:
            return pw
    except Exception:
        pass
    pw = os.environ.get("ADMIN_PASSWORD")
    if pw:
        return pw
    import configparser
    cfg = configparser.ConfigParser()
    cfg.read(Path(__file__).parent / "config.ini", encoding="utf-8")
    return cfg.get("ADMIN", "password", fallback="")


# ================== ارسال‌کننده‌های بله ==================
def _send_bale(chat_id, text):
    from bale_sender import send_message
    return send_message(chat_id, text)


# ================== سازنده متن‌های بله ==================
def _norm_one(s):
    """نرمال‌سازی یک نام مثل ستون‌های اکسل (ی/ي، نیم‌فاصله، فاصله‌ها)"""
    from send_report import normalize_name
    return normalize_name(pd.Series([s])).iloc[0]


def _build_supervisor_bale_full(supervisor_name):
    """متن کامل بله برای سوپروایزر — شخصی‌سازی شده"""
    from send_report import (
        EXCEL_FILE, find_sheet, parse_percent, money, normalize_name, to_number
    )

    xls = pd.ExcelFile(EXCEL_FILE)
    sh = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    if not sh:
        raise ValueError("شیت تارگت پیدا نشد")

    t = pd.read_excel(xls, sheet_name=sh)
    t.columns = t.columns.astype(str).str.strip()

    branch_col = next((c for c in t.columns if "نام شعبه" in c), None)
    ach_col = next((c for c in t.columns if "تحقق" in c), None)
    sup_col = next((c for c in t.columns if c.strip() == "سوپروایزر"), None)
    raked_col = next((c for c in t.columns if "ریالی راکد" in c), None)

    if not (sup_col and branch_col and ach_col):
        raise ValueError("ستون‌های لازم پیدا نشد")

    t[sup_col] = normalize_name(t[sup_col])
    t[branch_col] = normalize_name(t[branch_col])
    t[ach_col] = parse_percent(t[ach_col])
    if raked_col:
        t[raked_col] = to_number(t[raked_col])

    # حذف ردیف‌های خالی، جمع و صفر
    _nm = t[branch_col].astype(str).str.strip()
    _is_total = _nm.str.contains(r"^\s*(?:جمع|مجموع|کل|total|grand|sum)", case=False, regex=True)
    _is_empty = t[branch_col].isna() | _nm.isin(["", "nan", "None", "NaN"])
    _is_zero = to_number(t[ach_col]) == 0
    t = t[~(_is_total | _is_empty | _is_zero)].copy()

    sub = t[t[sup_col].astype(str).str.strip() == _norm_one(supervisor_name)].copy()
    if sub.empty:
        raise ValueError(f"شعبه‌ای برای «{supervisor_name}» پیدا نشد")

    sub = sub.sort_values(ach_col, ascending=False).reset_index(drop=True)
    mean_ach = sub[ach_col].mean()
    total_raked = sub[raked_col].sum() if raked_col else 0

    # رتبه سوپروایزر
    all_sup = t.groupby(sup_col)[ach_col].mean().sort_values(ascending=False)
    rank = "?"
    _sn = _norm_one(supervisor_name)
    if _sn in all_sup.index:
        rank = list(all_sup.index).index(_sn) + 1
    total_sup = len(all_sup)

    if mean_ach >= 100:
        status = "🟢 بالاتر از تارگت"
    elif mean_ach >= 80:
        status = "🟡 در حال پیشرفت"
    else:
        status = "🔴 نیازمند توجه"

    best = sub.head(3)
    worst = sub.tail(3).iloc[::-1]

    lines = [
        "📊 <b>گزارش عملکرد هفتگی</b>",
        f"👤 <b>{supervisor_name}</b>",
        "━━━━━━━━━━━━━━━",
        "",
        "<b>🎯 عملکرد شما:</b>",
        f"   میانگین تحقق: <b>{mean_ach:.1f}%</b>",
        f"   رتبه شما: <b>{rank}</b> از {total_sup} سوپروایزر",
        f"   وضعیت: {status}",
    ]

    if raked_col and total_raked:
        lines.append(f"   مجموع راکد: {money(total_raked)}")

    lines.append("")
    lines.append(f"<b>🏪 شعبه‌های تحت پوشش ({len(sub)}):</b>")
    lines.append("")
    lines.append("<b>🥇 بهترین شعبه‌ها:</b>")
    for _, r in best.iterrows():
        lines.append(f"   • {r[branch_col]} — {r[ach_col]:.1f}%")

    lines.append("")
    lines.append("<b>⚠️ شعبه‌های نیازمند توجه:</b>")
    for _, r in worst.iterrows():
        lines.append(f"   • {r[branch_col]} — {r[ach_col]:.1f}%")

    lines.append("")
    lines.append("━━━━━━━━━━━━━━━")
    lines.append(f"📅 {datetime.now().strftime('%Y-%m-%d')}")

    return "\n".join(lines)


def _build_manager_bale_text():
    """متن بله برای مدیران — استراتژیک"""
    from send_report import (
        EXCEL_FILE, find_sheet, parse_percent, money, normalize_name, to_number
    )

    xls = pd.ExcelFile(EXCEL_FILE)
    sh = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    if not sh:
        raise ValueError("شیت تارگت پیدا نشد")

    t = pd.read_excel(xls, sheet_name=sh)
    t.columns = t.columns.astype(str).str.strip()

    branch_col = next((c for c in t.columns if "نام شعبه" in c), None)
    ach_col = next((c for c in t.columns if "تحقق" in c), None)
    sup_col = next((c for c in t.columns if c.strip() == "سوپروایزر"), None)
    raked_col = next((c for c in t.columns if "ریالی راکد" in c), None)

    if not (ach_col and branch_col):
        raise ValueError("ستون‌های لازم پیدا نشد")

    t[ach_col] = parse_percent(t[ach_col])
    t[branch_col] = normalize_name(t[branch_col])
    if raked_col:
        t[raked_col] = to_number(t[raked_col])
    if sup_col:
        t[sup_col] = normalize_name(t[sup_col])

    # حذف ردیف‌های خالی، جمع و صفر
    name = t[branch_col].astype(str).str.strip()
    is_total = name.str.contains(r"^\s*(?:جمع|مجموع|کل|total|grand|sum)", case=False, regex=True)
    is_empty = t[branch_col].isna() | name.isin(["", "nan", "None", "NaN"])
    is_zero = to_number(t[ach_col]) == 0
    t = t[~(is_total | is_empty | is_zero)].copy()

    vals = t[ach_col].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success = int((vals >= 100).sum())
    warning = int(((vals >= 80) & (vals < 100)).sum())
    danger = int((vals < 80).sum())
    total = len(vals)

    pct_success = success / total * 100 if total else 0
    pct_warning = warning / total * 100 if total else 0
    pct_danger = danger / total * 100 if total else 0

    t_sorted = t.sort_values(ach_col, ascending=False)
    best5 = t_sorted.head(5)
    worst5 = t_sorted.tail(5).iloc[::-1]

    lines = [
        "📊 <b>گزارش مدیریتی هفتگی</b>",
        "🏢 فروشگاه‌های زنجیره‌ای افق کوروش",
        "━━━━━━━━━━━━━━━",
        "",
        "<b>📈 عملکرد کلی دیستریکت:</b>",
        f"   میانگین تحقق: <b>{mean_ach:.1f}%</b>",
        f"   🟢 موفق: {success} ({pct_success:.0f}%)",
        f"   🟡 پیشرفت: {warning} ({pct_warning:.0f}%)",
        f"   🔴 بحرانی: {danger} ({pct_danger:.0f}%)",
        "",
        "<b>🥇 ۵ شعبه برتر:</b>",
    ]

    for _, r in best5.iterrows():
        lines.append(f"   • {r[branch_col]} — {r[ach_col]:.1f}%")

    lines.append("")
    lines.append("<b>⚠️ ۵ شعبه بحرانی:</b>")
    for _, r in worst5.iterrows():
        lines.append(f"   • {r[branch_col]} — {r[ach_col]:.1f}%")

    if sup_col:
        sup_avg = t.groupby(sup_col)[ach_col].mean().sort_values()
        worst_sups = sup_avg.head(3)
        lines.append("")
        lines.append("<b>📉 سوپروایزرهای نیازمند توجه:</b>")
        for nm, val in worst_sups.items():
            lines.append(f"   • {nm} — {val:.1f}%")

    if raked_col:
        total_raked = t[raked_col].sum()
        lines.append("")
        lines.append(f"<b>💰 مجموع راکد دیستریکت:</b> {money(total_raked)}")

    lines.append("")
    lines.append("━━━━━━━━━━━━━━━")
    lines.append(f"📅 {datetime.now().strftime('%Y-%m-%d')}")

    return "\n".join(lines)


def _build_regular_bale_text():
    """متن بله برای افراد عادی — کوتاه"""
    from send_report import EXCEL_FILE, find_sheet, parse_percent

    xls = pd.ExcelFile(EXCEL_FILE)
    sh = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    if not sh:
        raise ValueError("شیت تارگت پیدا نشد")

    t = pd.read_excel(xls, sheet_name=sh)
    t.columns = t.columns.astype(str).str.strip()

    ach_col = next((c for c in t.columns if "تحقق" in c), None)
    if not ach_col:
        raise ValueError("ستون تحقق پیدا نشد")

    t[ach_col] = parse_percent(t[ach_col])
    _branch = next((c for c in t.columns if "نام شعبه" in c), None)
    if _branch:
        _nm = t[_branch].astype(str).str.strip()
        _empty = t[_branch].isna() | _nm.isin(["", "nan", "None", "NaN"])
        t = t[~_empty].copy()
    vals = t[ach_col].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success = int((vals >= 100).sum())
    total = len(vals)

    status = "🟢" if mean_ach >= 100 else ("🟡" if mean_ach >= 80 else "🔴")

    return (
        f"📢 <b>گزارش هفتگی کالای راکد</b>\n"
        f"━━━━━━━━━━━━━━━\n\n"
        f"{status} میانگین تحقق: <b>{mean_ach:.1f}%</b>\n"
        f"🟢 شعب موفق: <b>{success}</b> از {total}\n\n"
        f"برای جزئیات بیشتر به داشبورد مراجعه کنید.\n"
        f"━━━━━━━━━━━━━━━\n"
        f"📅 {datetime.now().strftime('%Y-%m-%d')}"
    )


# ================== صفحه اصلی ==================


# ================== لود منابع گزارش تصویری ==================
def _load_source_data(source_key):
    """لود داده بر اساس منبع انتخابی — بدون نیاز به app.py"""
    import pandas as pd
    from pathlib import Path

    # ── بخش «فروش اینترنتی و مغایرت‌گیری» (فایل‌های جدا) ──
    if str(source_key).startswith(("online_", "adjust_")):
        from online_section import load_online_source
        return load_online_source(source_key)

    from send_report import (
        EXCEL_FILE, find_sheet, normalize_name, to_number, parse_percent
    )

    if not Path(EXCEL_FILE).exists():
        return None

    xls = pd.ExcelFile(EXCEL_FILE)

    def _read(sheet_name):
        if not sheet_name:
            return None
        try:
            d = pd.read_excel(xls, sheet_name=sheet_name)
            d.columns = d.columns.astype(str).str.strip()
            return d
        except Exception:
            return None

    # ── خروجی فروش ──
    if source_key == "sold":
        sh = find_sheet(xls, "خروجی")
        if not sh:
            for n in xls.sheet_names:
                if "خروجی" in str(n):
                    sh = n
                    break
        d = _read(sh)
        if d is None:
            return None
        for c in d.columns:
            if "نام شعبه" in c or "نام کالا" in c or "سوپروایزر" in c:
                d[c] = normalize_name(d[c])
            if "شیفت" in c or "وضعیت فروش" in c:
                d[c] = d[c].astype(str).str.strip()
        return d

    # ── tbl60 / tbl45 ──
    if source_key == "tbl60":
        return _read(find_sheet(xls, "tbl60"))
    if source_key == "tbl45":
        return _read(find_sheet(xls, "tbl45"))

    # ── pq60 / pq45 ──
    if source_key == "pq60":
        return _read(find_sheet(xls, "pq60"))
    if source_key == "pq45":
        return _read(find_sheet(xls, "pq45"))

    if source_key == "project":
        d60 = _read(find_sheet(xls, "pq60"))
        d45 = _read(find_sheet(xls, "pq45"))
        frames = []
        if d60 is not None and not d60.empty:
            d60 = d60.copy()
            d60.insert(0, "نوع پروژه", "۶۰ روزه")
            frames.append(d60)
        if d45 is not None and not d45.empty:
            d45 = d45.copy()
            d45.insert(0, "نوع پروژه", "۴۵ روزه")
            frames.append(d45)
        if not frames:
            return None
        import pandas as _pd
        out = _pd.concat(frames, ignore_index=True)
        for c in out.columns:
            if "نام شعبه" in c or "نام کالا" in c or "سوپروایزر" in c:
                out[c] = normalize_name(out[c])
            if "شیفت" in c:
                out[c] = out[c].astype(str).str.strip()
        return out

    # ── خلاصه ۶۰ روزه ──
    if source_key == "summary60":
        return _read(find_sheet(xls, "60 روزه"))

    # ── تارگت ──
    if source_key == "target":
        sh = find_sheet(xls, "تارگت")
        if not sh:
            sh = find_sheet(xls, "روند و تارگت")
        return _read(sh)

    # ── خلاصه شعب ──
    if source_key == "branches":
        d = _read(find_sheet(xls, "tbl60"))
        if d is None or d.empty:
            return None
        # پیدا کردن ستون‌ها
        br = next((c for c in d.columns if "نام شعبه" in c), None)
        val = next((c for c in d.columns if "ریالی اقلام راکد" in c), None)
        sup = next((c for c in d.columns if "سوپروایزر" in c), None)
        bc = next((c for c in d.columns if "بارکد" in c), None)
        if not br or not val:
            return None
        d[val] = to_number(d[val])
        g = d.groupby(br, as_index=False).agg({
            val: "sum",
            bc: "count",
            sup: "first"
        })
        g = g.rename(columns={val: "راکد ۶۰ روزه (ریال)", bc: "تعداد قلم", sup: "سوپروایزر"})
        g = g.sort_values("راکد ۶۰ روزه (ریال)", ascending=False).reset_index(drop=True)
        return g

    # ── خلاصه سوپروایزرها ──
    if source_key == "supervisors":
        d = _read(find_sheet(xls, "tbl60"))
        if d is None or d.empty:
            return None
        br = next((c for c in d.columns if "نام شعبه" in c), None)
        val = next((c for c in d.columns if "ریالی اقلام راکد" in c), None)
        sup = next((c for c in d.columns if "سوپروایزر" in c), None)
        if not sup or not val:
            return None
        d[val] = to_number(d[val])
        g = d.groupby(sup, as_index=False).agg({
            val: "sum",
            br: "nunique"
        })
        g = g.rename(columns={val: "مجموع راکد (ریال)", br: "تعداد شعب"})
        g = g.sort_values("مجموع راکد (ریال)", ascending=False).reset_index(drop=True)
        return g

    return None


def render_report_sender(key_suffix="rs"):
    """صفحه ارسال گزارش — با رمز محافظت‌شده"""
    st.markdown("## 📨 ارسال گزارش")
    st.caption("ارسال گزارش به سه گروه از طریق کانال‌های مختلف")

    # ================== رمز ورود ==================
    # اگه رمزها خاموشه، مستقیم رد شو
    try:
        import configparser as _cfg_pw
        _c_pw = _cfg_pw.ConfigParser()
        _c_pw.read("config.ini", encoding="utf-8")
        _pw_disabled = _c_pw.getboolean("SECURITY", "disable_all_passwords", fallback=False)
    except Exception:
        _pw_disabled = False

    state_key = "unlock_locked_pages"
    if _pw_disabled:
        st.session_state[state_key] = True

    if not st.session_state.get(state_key):
        st.warning("🔒 این صفحه محافظت‌شده است. برای ورود رمز را وارد کنید.")

        c1, c2 = st.columns([2, 3])
        with c1:
            entered = st.text_input("🔑 رمز ارسال گزارش", type="password",
                                     key=f"report_pw_{key_suffix}")
        with c2:
            st.write("")
            st.write("")
            if st.button("ورود", key=f"report_login_{key_suffix}",
                         type="primary", use_container_width=True):
                pw = _get_admin_password()
                if entered and hmac.compare_digest(entered.encode(), str(pw).encode()):
                    st.session_state[state_key] = True
                    st.rerun()
                else:
                    st.error("❌ رمز اشتباه است")

        st.stop()
        return

    # ================== خروج ==================
    col_out, col_empty = st.columns([1, 4])
    with col_out:
        if st.button("🚪 خروج", key=f"report_logout_{key_suffix}"):
            st.session_state.pop(state_key, None)
            st.rerun()

    st.divider()

    # ================== ادامه صفحه ==================
    contacts = load_contacts()
    if not contacts:
        st.warning("⚠️ هنوز مخاطبی نداری.")
        st.divider()
        st.markdown("### 👥 مدیریت مخاطبین")
        _render_contacts_management(contacts, key_suffix)
        return

    supervisors = [c for c in contacts if c["type"] == "supervisor"]
    managers = [c for c in contacts if c["type"] == "manager"]
    regulars = [c for c in contacts if c["type"] == "regular"]

    # آمار بالا
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👤 سوپروایزر", len(supervisors))
    c2.metric("🏆 مدیر", len(managers))
    c3.metric("📬 عادی", len(regulars))
    with_phone = sum(1 for c in contacts if c.get("bale_chat_id"))
    c4.metric("📱 دارای chat_id", with_phone)

    st.divider()

    tab_bale, tab_email, tab_image, tab_contacts = st.tabs([
         "📱 بله",
        "📧 ایمیل",
        "📸 گزارش تصویری",

        "👥 مدیریت مخاطبین 🔒",
    ])

    with tab_bale:
        _render_bale_channel(supervisors, managers, regulars, key_suffix)

    with tab_email:
        _render_email_channel(supervisors, managers, regulars, key_suffix)

    with tab_image:
        import configparser as _cfg
        _c = _cfg.ConfigParser()
        _c.read("config.ini", encoding="utf-8")
        if _c.getboolean("SECURITY", "disable_image_report", fallback=False):
            st.warning("⚠️ این قابلیت موقتاً غیرفعال شده است.")
            st.stop()

        st.markdown("### 📸 گزارش تصویری")
        st.caption("یک منبع انتخاب کن، فیلتر بزن، و به بله بفرست")

        # ══════════════════════════════════════════════
        # مرحله ۱: منبع داده
        # ══════════════════════════════════════════════
        SOURCES = {
            "📦 پروژه راکد (۲۰ قلم — ۶۰ و ۴۵ روزه)": "project",
            "🛒 خروجی فروش (کالاهای فروش‌رفته)": "sold",
            "📦 همه اقلام راکد ۶۰ روزه": "tbl60",
            "📦 همه اقلام راکد ۴۵ روزه": "tbl45",
            "🏪 خلاصه شعب": "branches",
            "👤 خلاصه سوپروایزرها": "supervisors",
            "🎯 تارگت و رتبه‌بندی": "target",
            "📈 روند ۶۰ روزه": "summary60",
        }
        try:
            from online_section import ONLINE_SOURCES
            SOURCES.update(ONLINE_SOURCES)
        except Exception:
            pass

        col_src, col_rec = st.columns([2, 2])
        with col_src:
            source_label = st.selectbox(
                "📁 منبع داده",
                list(SOURCES.keys()),
                key=f"img_source_{key_suffix}",
            )
        source_key = SOURCES[source_label]

        # گیرنده
        recipient_options = {"👤 خودم (شاهین)": 1343334968}
        for c in contacts:
            cid = c.get("bale_chat_id")
            if cid and cid not in recipient_options.values():
                recipient_options[f"{c.get('name')} — {c.get('type')}"] = cid

        with col_rec:
            selected_recip = st.selectbox(
                "📬 گیرنده",
                list(recipient_options.keys()),
                key=f"img_recip_{key_suffix}",
            )
        target_chat_id = recipient_options[selected_recip]

        # لود داده
        try:
            sold_df = _load_source_data(source_key)
        except Exception as e:
            st.error(f"خطا در بارگذاری: {e}")
            sold_df = None

        # ══════════════════════════════════════════════
        # فیلتر ستون‌های روند (برای summary60 و target)
        # ══════════════════════════════════════════════
        if source_key in ("summary60", "target") and sold_df is not None and not sold_df.empty:
            import re as _re_tr
            _MONTHS = r"(?:فروردین|اردیبهشت|خرداد|تیر|مرداد|شهریور|مهر|آبان|آذر|دی|بهمن|اسفند)"

            def _is_trend_col(c):
                cl = str(c).translate(FA_DIGITS).replace("\u200c", " ")
                cl = _re_tr.sub(r"\s+", " ", cl).strip()
                return bool(_re_tr.match(rf"^درصد\s+راکد\s*\d+\s*(?:{_MONTHS})?\s*\d*$", cl))

            _trend_cols = [c for c in sold_df.columns if _is_trend_col(c)]

            if len(_trend_cols) > 2:
                _MO = {"فروردین": 1, "اردیبهشت": 2, "خرداد": 3, "تیر": 4,
                       "مرداد": 5, "شهریور": 6, "مهر": 7, "آبان": 8,
                       "آذر": 9, "دی": 10, "بهمن": 11, "اسفند": 12}

                def _date_key(c):
                    cl = str(c).translate(FA_DIGITS).replace("\u200c", " ")
                    cl = _re_tr.sub(r"\s+", " ", cl)
                    m = _re_tr.search(rf"(\d+)\s*({_MONTHS})", cl)
                    if m:
                        return (_MO.get(m.group(2), 99), int(m.group(1)))
                    m2 = _re_tr.search(r"(\d+)", cl)
                    if m2:
                        return (99, int(m2.group(1)))
                    return (999, 999)

                _trend_sorted = sorted(_trend_cols, key=_date_key)

                st.markdown("**📊 انتخاب ستون‌های روند:**")
                _cs, _co = st.columns([3, 1])
                with _cs:
                    _show_n = st.slider(
                        "چند ستون آخر نمایش داده بشه؟",
                        min_value=2, max_value=len(_trend_sorted),
                        value=min(2, len(_trend_sorted)), step=1,
                        key=f"tr_show_n_{key_suffix}",
                    )
                with _co:
                    st.write(""); st.write("")
                    _show_all = st.checkbox("نمایش همه", value=False,
                                             key=f"tr_show_all_{key_suffix}")
                    if _show_all:
                        _show_n = len(_trend_sorted)

                _keep_trend = _trend_sorted[-_show_n:] if not _show_all else _trend_sorted
                _other_cols = [c for c in sold_df.columns if c not in _trend_cols]
                sold_df = sold_df[_other_cols + _keep_trend].copy()

        # ══════════════════════════════════════════════
        # فیلتر فروش‌رفته (فقط برای sold)
        # ══════════════════════════════════════════════
        if source_key == "sold" and sold_df is not None and not sold_df.empty:
            _status_col = next((c for c in sold_df.columns
                                if "وضعیت فروش" in c or "وضعیت" in c), None)
            if _status_col:
                if st.checkbox("✅ فقط کالاهای فروش‌رفته", value=True,
                               key=f"only_sold_{key_suffix}"):
                    sold_df = sold_df[
                        sold_df[_status_col].astype(str).str.contains("فروش شد", na=False)
                    ].copy().reset_index(drop=True)

            # ترتیب ستون‌ها
            _desired = ["نام شعبه", "شیفت", "بارکد", "نام کالا", "وضعیت فروش"]
            _existing = []
            for d in _desired:
                for c in sold_df.columns:
                    if d == c or d in c:
                        _existing.append(c)
                        break
            for c in sold_df.columns:
                if c not in _existing:
                    _existing.append(c)
            sold_df = sold_df[_existing].copy()

        # ── بازچینی و فرمت ستون‌های target ──
        if source_key == "target" and sold_df is not None and not sold_df.empty:
            import re as _re_tc
            _tc_fa = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

            _sup_c = next((c for c in sold_df.columns
                          if str(c).strip() == "سوپروایزر"), None)
            _name_c = next((c for c in sold_df.columns if "نام شعبه" in str(c)), None)
            _shop_sup_c = next((c for c in sold_df.columns
                               if "سرپرست فروشگاه" in str(c) or "سرپرست شعبه" in str(c)), None)
            _raked_c = next((c for c in sold_df.columns
                            if "ریالی راکد" in str(c) or "ریالی اقلام راکد" in str(c)), None)
            _target_c = next((c for c in sold_df.columns
                             if str(c).startswith("تارگت") and "تغییرات" not in str(c)), None)
            _ach_c = next((c for c in sold_df.columns if "تحقق" in str(c)), None)

            _trend_now = []
            for c in sold_df.columns:
                cs = str(c).translate(_tc_fa).replace("\u200c", " ")
                cs = _re_tc.sub(r"\s+", " ", cs).strip()
                if _re_tc.match(r"^درصد\s+راکد\s*\d+", cs):
                    _trend_now.append(c)

            _order = []
            for _c in ([_sup_c, _name_c, _shop_sup_c, _raked_c] + _trend_now
                       + [_target_c, _ach_c]):
                if _c is not None and _c not in _order and _c in sold_df.columns:
                    _order.append(_c)

            if len(_order) >= 5:
                sold_df = sold_df[_order].copy()

                def _fmt_pct(v):
                    try:
                        n = float(v)
                        if n != n:
                            return "—"
                        if abs(n) <= 1.0 and n != 0:
                            return f"{n * 100:.2f}%"
                        return f"{n:.2f}%"
                    except (ValueError, TypeError):
                        return str(v) if v is not None else "—"

                for _c in sold_df.columns:
                    cs = str(_c)
                    if "درصد" in cs or "تحقق" in cs or ("تارگت" in cs and "تغییرات" not in cs):
                        if sold_df[_c].dtype.kind in "fiu":
                            sold_df[_c] = sold_df[_c].apply(_fmt_pct)

            try:
                from image_sender import set_keep_cols
                set_keep_cols(list(sold_df.columns))
            except (ImportError, AttributeError):
                pass

        # قالب اختصاصی عکس برای منبع‌های فروش اینترنتی/مغایرت‌گیری (بقیه: ظاهر قبلی)
        try:
            from image_sender import set_theme
            set_theme(source_key)
        except (ImportError, AttributeError):
            pass

        if sold_df is None or sold_df.empty:
            st.warning(f"⚠️ داده‌ای برای «{source_label}» پیدا نشد.")
            if str(source_key).startswith(("online_", "adjust_")):
                st.info("فایل‌ها را از صفحه «🌐 فروش اینترنتی و مغایرت‌گیری» آپلود کنید.")
            st.stop()

        st.success(f"✅ {len(sold_df):,} ردیف بارگذاری شد")
        st.divider()

        # ══════════════════════════════════════════════
        # مرحله ۲: فیلترها (expandable)
        # ══════════════════════════════════════════════
        # چک کن ستون سرپرست وجود داره
        _sup_col = next((c for c in sold_df.columns
                         if "سوپروایزر" in c or "سرپرست" in c), None)
        _branch_col = next((c for c in sold_df.columns if "نام شعبه" in c), None)

        with st.expander("🎛 فیلترها (سرپرست، شعبه، سورت)", expanded=False):
            fc1, fc2 = st.columns(2)

            # فیلتر سرپرست
            selected_sup = "همه"
            if _sup_col:
                sup_list = ["همه"] + sorted(
                    sold_df[_sup_col].dropna().astype(str).unique().tolist()
                )
                with fc1:
                    selected_sup = st.selectbox(
                        "👤 سوپروایزر",
                        sup_list,
                        key=f"filter_sup_{key_suffix}",
                    )

            # فیلتر شعبه (بر اساس سرپرست انتخاب‌شده)
            selected_branch = "همه"
            if _branch_col:
                if selected_sup != "همه" and _sup_col:
                    _sub = sold_df[sold_df[_sup_col].astype(str) == selected_sup]
                    branch_list = ["همه"] + sorted(_sub[_branch_col].dropna().astype(str).unique().tolist())
                else:
                    branch_list = ["همه"] + sorted(sold_df[_branch_col].dropna().astype(str).unique().tolist())
                with fc2:
                    selected_branch = st.selectbox(
                        "🏪 شعبه",
                        branch_list,
                        key=f"filter_branch_{key_suffix}",
                    )

            # سورت
            fc3, fc4 = st.columns(2)
            _sortable = [c for c in sold_df.columns if sold_df[c].dtype == object][:6]
            _ORIG = "— ترتیب اصلی —"
            _is_online_src = str(source_key).startswith(("online_", "adjust_"))
            if _is_online_src:
                _sortable = [_ORIG] + _sortable
            with fc3:
                _default_sort = (_ORIG if _is_online_src else
                                 ("نام شعبه" if "نام شعبه" in _sortable else (_sortable[0] if _sortable else None)))
                sort_choice = st.selectbox(
                    "↕️ ترتیب",
                    _sortable,
                    index=_sortable.index(_default_sort) if _default_sort in _sortable else 0,
                    key=f"sort_{key_suffix}",
                )
            with fc4:
                sort_dir = st.radio("جهت", ["صعودی", "نزولی"],
                                    horizontal=True, key=f"sort_dir_{key_suffix}")

            # اعمال فیلترها
            if selected_sup != "همه" and _sup_col:
                sold_df = sold_df[sold_df[_sup_col].astype(str) == selected_sup]
            if selected_branch != "همه" and _branch_col:
                sold_df = sold_df[sold_df[_branch_col].astype(str) == selected_branch]

            # سورت
            ascending = (sort_dir == "صعودی")
            if sort_choice and sort_choice != _ORIG:
                try:
                    sold_df = sold_df.sort_values(
                        sort_choice, ascending=ascending,
                        key=lambda x: x.astype(str)
                    ).reset_index(drop=True)
                except Exception:
                    pass

        # نمایش تعداد نهایی
        st.info(f"📊 **{len(sold_df):,} ردیف** پس از فیلتر — "
                f"**{sold_df[_branch_col].nunique() if _branch_col else '—'} شعبه**")

        # پیش‌نمایش
        with st.expander("👁 پیش‌نمایش داده (۱۵ ردیف اول)", expanded=False):
            st.dataframe(sold_df.head(15), use_container_width=True, hide_index=True)

        st.divider()

        # ══════════════════════════════════════════════
        # مرحله ۳: تنظیمات ارسال
        # ══════════════════════════════════════════════
        st.markdown("#### ⚙️ تنظیمات ارسال")

        # حالت ارسال
        send_mode = st.radio(
            "حالت ارسال",
            ["📄 یک عکس کلی", "🏪 جدا برای هر شعبه"],
            horizontal=True,
            key=f"send_mode_{key_suffix}",
        )
        per_branch_mode = (send_mode == "🏪 جدا برای هر شعبه")

        per_branch_msg = "لطفاً اقدام به پالت چینی نمایید"
        if per_branch_mode:
            per_branch_msg = st.text_input(
                "پیام روی هر شعبه:",
                value=per_branch_msg,
                key=f"per_branch_msg_{key_suffix}",
            )

        compact_mode = st.checkbox(
            "📱 حالت فشرده (مثل خروجی اکسل — بله عکس را کوچک نمی‌کند و کیفیت حفظ می‌شود)",
            value=True, key=f"img_compact_{key_suffix}",
            help="عکس با اندازه‌ی نهایی ساخته می‌شود. برای حفظ کیفیت، حداکثر ۶۰ ردیف در هر عکس.")

        sc1, sc2, sc3 = st.columns(3)
        with sc1:
            chunk_size = st.number_input("ردیف در هر عکس",
                                          min_value=5, max_value=100,
                                          value=25, step=5,
                                          key=f"chunk_{key_suffix}")
            if compact_mode and int(chunk_size) > 60:
                chunk_size = 60
                st.caption("ℹ️ در حالت فشرده حداکثر ۶۰ ردیف در هر عکس.")
        with sc2:
            delay = st.number_input("فاصله (ثانیه)",
                                     min_value=2, max_value=120,
                                     value=5, step=1,
                                     key=f"delay_{key_suffix}")
        with sc3:
            if not per_branch_mode:
                n_total = (len(sold_df) + chunk_size - 1) // chunk_size
                from_chunk = st.number_input("از عکس", min_value=1, max_value=n_total,
                                              value=1, step=1, key=f"from_chunk_{key_suffix}")
            else:
                from_chunk = 1

        if not per_branch_mode:
            n_to_send = (len(sold_df) + chunk_size - 1) // chunk_size
            st.info(f"📊 **{n_to_send} عکس** ارسال می‌شه — زمان: {(n_to_send-1)*int(delay)} ثانیه")

        # پیش‌نمایش عکس اول
        if st.button("👁 پیش‌نمایش عکس", key=f"img_preview_{key_suffix}",
                     use_container_width=True):
            try:
                from image_sender import df_to_image
                Path(".bale_images").mkdir(exist_ok=True)
                tmp_path = f".bale_images/preview_{key_suffix}.png"
                preview_chunk = sold_df.head(min(int(chunk_size), 60) if compact_mode else int(chunk_size)).copy()
                df_to_image(preview_chunk, tmp_path,
                            title=f"{source_label} — پیش‌نمایش",
                            font_size=11, compact=compact_mode,
                            per_branch_msg=per_branch_msg if per_branch_mode else None)
                st.image(tmp_path, use_container_width=True)
            except Exception as e:
                st.error(f"❌ {e}")

        st.divider()

        # ══════════════════════════════════════════════
        # مرحله ۴: ارسال
        # ══════════════════════════════════════════════
        state_key = f"sending_{key_suffix}"
        state = st.session_state.get(state_key)

        if state and state.get("active"):
            sent = state.get("sent", 0)
            total = state.get("total", 0)
            _is_per_branch = state.get("per_branch_mode", False)
            _mode_label = "شعبه" if _is_per_branch else "عکس"
            st.warning(f"⏳ در حال ارسال... {sent} از {total} {_mode_label}")
            st.progress(sent / max(total, 1))

            # اگه قبلاً خطا داشته، به کاربر نمایش بده
            if state.get("error"):
                st.info(f"ℹ️ آخرین خطا: {state['error']}")
            if st.button("🛑 توقف", type="primary",
                         key=f"cancel_{key_suffix}",
                         use_container_width=True):
                state["cancel"] = True
                st.rerun()

            if sent < total and not state.get("cancel"):
                if _is_per_branch:
                    # ── حالت جدا برای هر شعبه ──
                    try:
                        _branch_names = state["branch_names"]
                        _branch_name = _branch_names[sent]
                        _chunk_size = state["chunk_size"]
                        _msg = state["per_branch_msg"]
                        _title = state["source_label"]

                        # ردیف‌های این شعبه
                        _branch_df = sold_df[sold_df[_branch_col] == _branch_name].reset_index(drop=True)
                        _n_rows = len(_branch_df)
                        _n_chunks = (_n_rows + _chunk_size - 1) // _chunk_size

                        from image_sender import df_to_image, send_photo
                        Path(".bale_images").mkdir(exist_ok=True)

                        for _ci in range(_n_chunks):
                            _start = _ci * _chunk_size
                            _end = min(_start + _chunk_size, _n_rows)
                            _chunk = _branch_df.iloc[_start:_end].copy()
                            _part = f" (بخش {_ci+1}/{_n_chunks})" if _n_chunks > 1 else ""
                            _full_title = f"{_title} — {_branch_name}{_part}"

                            _img_path = Path(".bale_images") / f"br_{sent+1}_{_ci+1}.png"
                            df_to_image(_chunk, str(_img_path), title=_full_title,
                                        font_size=11, per_branch_msg=_msg,
                                        compact=st.session_state.get(f"img_compact_{key_suffix}", False))
                            _cap = f"🏪 {_branch_name} | {_msg}"
                            send_photo(target_chat_id, str(_img_path), caption=_cap)

                        state["sent"] = sent + 1
                    except Exception as e:
                        # نگه‌داشتن sent برای ادامه
                        state["active"] = False
                        state["error"] = str(e)
                        # sent قبلاً آپدیت شده، پس ادامه از اینجا
                        st.rerun()
                else:
                    # ── حالت یک عکس کلی ──
                    i = state["from_chunk"] + sent - 1
                    start_row = i * state["chunk_size"]
                    end_row = min(start_row + state["chunk_size"], len(sold_df))
                    chunk = sold_df.iloc[start_row:end_row].copy()
                    title = f"{state['source_label']} — بخش {i+1}"

                    try:
                        from image_sender import df_to_image, send_photo
                        Path(".bale_images").mkdir(exist_ok=True)
                        img_path = Path(".bale_images") / f"r_{i+1}.png"
                        df_to_image(chunk, str(img_path), title=title, font_size=11,
                                    compact=st.session_state.get(f"img_compact_{key_suffix}", False))
                        send_photo(target_chat_id, str(img_path), caption=title)
                        state["sent"] = sent + 1
                    except Exception as e:
                        # نگه‌داشتن sent برای ادامه
                        state["active"] = False
                        state["error"] = str(e)
                        # sent قبلاً آپدیت شده، پس ادامه از اینجا
                        st.rerun()

                if state["sent"] < total:
                    _time_mod.sleep(state["delay"])
                    st.rerun()
                else:
                    state["active"] = False
                    state["done"] = True
                    st.rerun()

        elif state and state.get("done"):
            _sent = state.get("sent", 0)
            _total = state.get("total", 0)
            _is_pb = state.get("per_branch_mode", False)
            _lbl = "شعبه" if _is_pb else "عکس"
            st.success(f"✅ **تمام شد** — {_sent} {_lbl} ارسال شد")
            st.balloons()
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("🗑 پاک کردن", key=f"clear_{key_suffix}",
                             use_container_width=True):
                    del st.session_state[state_key]
                    st.rerun()
            with col_b:
                st.metric("ارسال شده", f"{_sent}/{_total}")

        elif state and state.get("cancel"):
            _sent = state.get("sent", 0)
            _total = state.get("total", 0)
            _is_pb = state.get("per_branch_mode", False)
            _lbl = "شعبه" if _is_pb else "عکس"

            st.warning(f"🛑 متوقف شد — **{_sent} از {_total}** {_lbl} ارسال شد")

            c1, c2 = st.columns(2)
            with c1:
                if st.button("▶️ ادامه",
                             type="primary",
                             key=f"resume_c_{key_suffix}",
                             use_container_width=True):
                    state["cancel"] = False
                    state["error"] = None
                    state["active"] = True
                    st.rerun()
            with c2:
                if st.button("🗑 پاک کردن",
                             key=f"clear2_{key_suffix}",
                             use_container_width=True):
                    del st.session_state[state_key]
                    st.rerun()

        elif state and state.get("error"):
            _sent = state.get("sent", 0)
            _total = state.get("total", 0)
            _is_pb = state.get("per_branch_mode", False)
            _lbl = "شعبه" if _is_pb else "عکس"

            st.error(f"❌ {state['error']}")
            st.warning(f"⚠️ ارسال متوقف شد — **{_sent} از {_total}** {_lbl} ارسال شد")

            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("▶️ ادامه از جایی که قطع شد",
                             type="primary",
                             key=f"resume_{key_suffix}",
                             use_container_width=True):
                    # از sent ادامه بده
                    state["error"] = None
                    state["active"] = True
                    state["cancel"] = False
                    st.rerun()
            with c2:
                if st.button("🔄 شروع مجدد از اول",
                             key=f"restart_{key_suffix}",
                             use_container_width=True):
                    state["sent"] = 0
                    state["error"] = None
                    state["active"] = True
                    state["cancel"] = False
                    st.rerun()
            with c3:
                if st.button("🗑 پاک کردن",
                             key=f"clear3_{key_suffix}",
                             use_container_width=True):
                    del st.session_state[state_key]
                    st.rerun()

        else:
            # دکمه‌های اصلی
            btn1, btn2 = st.columns(2)

            with btn1:
                if st.button("📄 ارسال PDF", type="secondary",
                             key=f"pdf_send_{key_suffix}",
                             use_container_width=True):
                    try:
                        from image_sender import df_to_pdf, send_pdf
                        Path(".bale_docs").mkdir(exist_ok=True)
                        pdf_path = f".bale_docs/report_{key_suffix}.pdf"
                        title_pdf = f"{source_label}"
                        df_to_pdf(sold_df, pdf_path, title=title_pdf, font_size=11)
                        send_pdf(target_chat_id, pdf_path, caption=title_pdf)
                        st.success(f"✅ PDF ارسال شد ({len(sold_df)} ردیف)")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ {e}")

            with btn2:
                btn_label = "📤 ارسال جدا برای هر شعبه" if per_branch_mode else "📤 شروع ارسال"
                if st.button(btn_label, type="primary",
                             key=f"img_send_{key_suffix}",
                             use_container_width=True):
                    if per_branch_mode:
                        # ── حالت تدریجی: توی هر rerun یک عکس ──
                        _branches = sold_df.groupby(_branch_col, sort=False).indices
                        _branch_names = list(_branches.keys())
                        st.session_state[state_key] = {
                            "active": True,
                            "done": False,
                            "cancel": False,
                            "sent": 0,
                            "total": len(_branch_names),
                            "chunk_size": int(chunk_size),
                            "delay": int(delay),
                            "per_branch_mode": True,
                            "branch_names": _branch_names,
                            "per_branch_msg": per_branch_msg,
                            "source_label": source_label,
                            "error": None,
                        }
                        st.rerun()
                    else:
                        n_total = (len(sold_df) + chunk_size - 1) // chunk_size
                        st.session_state[state_key] = {
                            "active": True,
                            "done": False,
                            "cancel": False,
                            "sent": 0,
                            "total": n_total,
                            "chunk_size": int(chunk_size),
                            "delay": int(delay),
                            "from_chunk": 1,
                            "source_label": source_label,
                            "error": None,
                        }
                        st.rerun()

    with tab_contacts:
        _render_contacts_management(contacts, key_suffix)


# ================== کانال بله ==================
def _render_bale_channel(supervisors, managers, regulars, key_suffix):
    st.markdown("### 📱 ارسال از طریق بله")

    tab_sup, tab_mgr, tab_reg = st.tabs([
        f"👤 سوپروایزرها ({len(supervisors)})",
        f"🏆 مدیران ({len(managers)})",
        f"📬 عادی ({len(regulars)})",
    ])

    with tab_sup:
        _bale_supervisors(supervisors, key_suffix)
    with tab_mgr:
        _bale_managers(managers, key_suffix)
    with tab_reg:
        _bale_regulars(regulars, key_suffix)


def _bale_supervisors(supervisors, key_suffix):
    if not supervisors:
        st.info("هیچ سوپروایزری نیست.")
        return

    st.markdown("#### 📤 ارسال تکی")
    sup_names = [(c.get("name") or c.get("email") or f"#{i}", c)
                 for i, c in enumerate(supervisors)]
    selected_label = st.selectbox("🔍 انتخاب سوپروایزر",
                                   ["(هیچ‌کدام)"] + [n for n, _ in sup_names],
                                   key=f"bale_sup_sel_{key_suffix}")

    if selected_label != "(هیچ‌کدام)":
        contact = next(c for n, c in sup_names if n == selected_label)
        actual_name = contact.get("name") or contact.get("email") or ""
        cid = contact.get("bale_chat_id")

        st.markdown(f"**👤 نام:** {actual_name} | **📱 chat_id:** `{cid or '— ثبت نشده'}`")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("👁 پیش‌نمایش", key=f"bale_sup_prev_{key_suffix}",
                         use_container_width=True):
                try:
                    txt = _build_supervisor_bale_full(actual_name)
                    st.session_state[f"bale_sup_txt_{key_suffix}"] = txt
                except Exception as e:
                    st.error(f"❌ {e}")
        with c2:
            if st.button("📱 ارسال بله", key=f"bale_sup_send_{key_suffix}",
                         type="primary", use_container_width=True):
                if not cid:
                    st.error("❌ chat_id ثبت نشده")
                else:
                    try:
                        txt = _build_supervisor_bale_full(actual_name)
                        _send_bale(cid, txt)
                        st.success("✅ ارسال شد")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ {e}")

        if st.session_state.get(f"bale_sup_txt_{key_suffix}"):
            with st.expander("👁 پیش‌نمایش پیام", expanded=True):
                st.code(st.session_state[f"bale_sup_txt_{key_suffix}"], language=None)

    st.divider()

    st.markdown("#### 📨 ارسال گروهی به همه سوپروایزرها")

    with_phone = [c for c in supervisors if c.get("bale_chat_id")]
    without_phone = [c for c in supervisors if not c.get("bale_chat_id")]

    st.caption(f"✅ {len(with_phone)} نفر آماده ارسال | ⚠️ {len(without_phone)} نفر بدون chat_id")

    if st.button("📱 ارسال به همه", key=f"bale_sup_all_{key_suffix}",
                 type="primary", use_container_width=True):
        if not with_phone:
            st.error("❌ هیچ سوپروایزری chat_id نداره")
        else:
            progress = st.progress(0)
            ok, fail = [], []
            for i, c in enumerate(with_phone):
                name = c.get("name") or c.get("email")
                progress.progress((i + 1) / len(with_phone), text=f"ارسال به {name}...")
                try:
                    txt = _build_supervisor_bale_full(name)
                    _send_bale(c["bale_chat_id"], txt)
                    ok.append(name)
                except Exception as e:
                    fail.append(f"{name}: {e}")
            progress.empty()
            if ok:
                st.success(f"✅ موفق: {len(ok)}")
                st.balloons()
            if fail:
                with st.expander(f"❌ {len(fail)} خطا"):
                    for f in fail:
                        st.error(f)


def _bale_managers(managers, key_suffix):
    st.markdown("#### 🏆 ارسال به مدیران در بله")

    if not managers:
        st.info("هیچ مدیری در لیست نیست.")
        return

    with_phone = [c for c in managers if c.get("bale_chat_id")]
    if not with_phone:
        st.warning("⚠️ هیچ مدیری chat_id بله نداره")
        return

    st.markdown(f"**گیرندگان ({len(with_phone)}):** " +
                "، ".join(c.get("name") or c.get("email") for c in with_phone))

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"bale_mgr_prev_{key_suffix}",
                     use_container_width=True):
            try:
                st.session_state[f"bale_mgr_txt_{key_suffix}"] = _build_manager_bale_text()
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📱 ارسال به همه مدیران", key=f"bale_mgr_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                txt = _build_manager_bale_text()
                progress = st.progress(0)
                ok, fail = [], []
                for i, c in enumerate(with_phone):
                    progress.progress((i + 1) / len(with_phone),
                                      text=f"ارسال به {c.get('name')}...")
                    try:
                        _send_bale(c["bale_chat_id"], txt)
                        ok.append(c.get("name"))
                    except Exception as e:
                        fail.append(f"{c.get('name')}: {e}")
                progress.empty()
                if ok:
                    st.success(f"✅ موفق: {len(ok)}")
                    st.balloons()
                if fail:
                    with st.expander(f"❌ {len(fail)} خطا"):
                        for f in fail:
                            st.error(f)
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"bale_mgr_txt_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            st.code(st.session_state[f"bale_mgr_txt_{key_suffix}"], language=None)


def _bale_regulars(regulars, key_suffix):
    st.markdown("#### 📬 ارسال به افراد عادی در بله")

    if not regulars:
        st.info("هیچ فرد عادی در لیست نیست.")
        return

    with_phone = [c for c in regulars if c.get("bale_chat_id")]
    if not with_phone:
        st.warning("⚠️ هیچ‌کدام chat_id بله ندارن")
        return

    st.markdown(f"**گیرندگان ({len(with_phone)}):** " +
                "، ".join(c.get("name") or c.get("email") for c in with_phone))

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"bale_reg_prev_{key_suffix}",
                     use_container_width=True):
            try:
                st.session_state[f"bale_reg_txt_{key_suffix}"] = _build_regular_bale_text()
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📱 ارسال به همه", key=f"bale_reg_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                txt = _build_regular_bale_text()
                progress = st.progress(0)
                ok, fail = [], []
                for i, c in enumerate(with_phone):
                    progress.progress((i + 1) / len(with_phone),
                                      text=f"ارسال به {c.get('name')}...")
                    try:
                        _send_bale(c["bale_chat_id"], txt)
                        ok.append(c.get("name"))
                    except Exception as e:
                        fail.append(f"{c.get('name')}: {e}")
                progress.empty()
                if ok:
                    st.success(f"✅ موفق: {len(ok)}")
                    st.balloons()
                if fail:
                    with st.expander(f"❌ {len(fail)} خطا"):
                        for f in fail:
                            st.error(f)
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"bale_reg_txt_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            st.code(st.session_state[f"bale_reg_txt_{key_suffix}"], language=None)


# ================== کانال ایمیل ==================
def _render_email_channel(supervisors, managers, regulars, key_suffix):
    st.markdown("### 📧 ارسال از طریق ایمیل")

    tab_sup, tab_mgr, tab_reg = st.tabs([
        f"👤 سوپروایزرها ({len(supervisors)})",
        f"🏆 مدیران ({len(managers)})",
        f"📬 عادی ({len(regulars)})",
    ])

    with tab_sup:
        _email_supervisors(supervisors, key_suffix)
    with tab_mgr:
        _email_managers(managers, key_suffix)
    with tab_reg:
        _email_regulars(regulars, key_suffix)


def _email_supervisors(supervisors, key_suffix):
    if not supervisors:
        st.info("هیچ سوپروایزری نیست.")
        return

    st.markdown("#### 📤 ارسال تکی")
    sup_names = [(c.get("name") or c.get("email") or f"#{i}", c)
                 for i, c in enumerate(supervisors)]
    selected_label = st.selectbox("🔍 انتخاب سوپروایزر",
                                   ["(هیچ‌کدام)"] + [n for n, _ in sup_names],
                                   key=f"em_sup_sel_{key_suffix}")

    if selected_label != "(هیچ‌کدام)":
        contact = next(c for n, c in sup_names if n == selected_label)
        actual_name = contact.get("name") or contact.get("email") or ""
        email = contact.get("email")

        st.markdown(f"**👤 نام:** {actual_name} | **📧 ایمیل:** `{email or '— ثبت نشده'}`")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("👁 پیش‌نمایش", key=f"em_sup_prev_{key_suffix}",
                         use_container_width=True):
                try:
                    html, _ = _build_supervisor_report(actual_name)
                    st.session_state[f"em_sup_html_{key_suffix}"] = html
                except Exception as e:
                    st.error(f"❌ {e}")
        with c2:
            if st.button("📤 ارسال ایمیل", key=f"em_sup_send_{key_suffix}",
                         type="primary", use_container_width=True):
                if not email:
                    st.error("❌ ایمیل ثبت نشده")
                else:
                    try:
                        html, _ = _build_supervisor_report(actual_name)
                        subject = f"گزارش عملکرد — {actual_name}"
                        _send_emails(html, subject, [email])
                        st.success("✅ ارسال شد")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ {e}")

        if st.session_state.get(f"em_sup_html_{key_suffix}"):
            with st.expander("👁 پیش‌نمایش", expanded=True):
                st.components.v1.html(st.session_state[f"em_sup_html_{key_suffix}"],
                                       height=500, scrolling=True)

    st.divider()
    st.markdown("#### 📨 ارسال گروهی")

    with_email = [c for c in supervisors if c.get("email")]
    st.caption(f"✅ {len(with_email)} نفر آماده ارسال")

    if st.button("📤 ارسال به همه", key=f"em_sup_all_{key_suffix}",
                 type="primary", use_container_width=True):
        if not with_email:
            st.error("❌ هیچ سوپروایزری ایمیل نداره")
        else:
            progress = st.progress(0)
            ok, fail = [], []
            for i, c in enumerate(with_email):
                name = c.get("name") or c.get("email")
                progress.progress((i + 1) / len(with_email), text=f"ارسال به {name}...")
                try:
                    html, _ = _build_supervisor_report(name)
                    _send_emails(html, f"گزارش عملکرد — {name}", [c["email"]])
                    ok.append(name)
                except Exception as e:
                    fail.append(f"{name}: {e}")
            progress.empty()
            if ok:
                st.success(f"✅ موفق: {len(ok)}")
                st.balloons()
            if fail:
                with st.expander(f"❌ {len(fail)} خطا"):
                    for f in fail:
                        st.error(f)


def _email_managers(managers, key_suffix):
    st.markdown("#### 🏆 ارسال گزارش مدیریتی")
    emails = [c["email"] for c in managers if c.get("email")]
    if not emails:
        st.info("هیچ مدیری با ایمیل ثبت نشده.")
        return

    st.markdown(f"**گیرندگان ({len(emails)}):** " + "، ".join(emails))

    subject = st.text_input("موضوع",
                             value="گزارش مدیریتی هفتگی — کالای راکد افق کوروش",
                             key=f"em_mgr_subj_{key_suffix}")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"em_mgr_prev_{key_suffix}", use_container_width=True):
            try:
                html, _ = _build_manager_report()
                st.session_state[f"em_mgr_html_{key_suffix}"] = html
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📤 ارسال به همه مدیران", key=f"em_mgr_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                html, _ = _build_manager_report()
                sent = _send_emails(html, subject, emails)
                st.success(f"✅ ایمیل به {', '.join(sent)} ارسال شد")
                st.balloons()
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"em_mgr_html_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            st.components.v1.html(st.session_state[f"em_mgr_html_{key_suffix}"],
                                   height=700, scrolling=True)


def _email_regulars(regulars, key_suffix):
    st.markdown("#### 📬 ارسال گزارش خلاصه")
    emails = [c["email"] for c in regulars if c.get("email")]
    if not emails:
        st.info("هیچ فرد عادی با ایمیل ثبت نشده.")
        return

    st.markdown(f"**گیرندگان ({len(emails)}):** " + "، ".join(emails))

    subject = st.text_input("موضوع",
                             value="گزارش خلاصه کالای راکد — افق کوروش",
                             key=f"em_reg_subj_{key_suffix}")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"em_reg_prev_{key_suffix}", use_container_width=True):
            try:
                html, _ = _build_regular_report()
                st.session_state[f"em_reg_html_{key_suffix}"] = html
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📤 ارسال به همه", key=f"em_reg_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                html, _ = _build_regular_report()
                sent = _send_emails(html, subject, emails)
                st.success(f"✅ ایمیل به {', '.join(sent)} ارسال شد")
                st.balloons()
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"em_reg_html_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            st.components.v1.html(st.session_state[f"em_reg_html_{key_suffix}"],
                                   height=400, scrolling=True)


# ================== مدیریت مخاطبین ==================
def _render_contacts_management(contacts, key_suffix):
    tab_import, tab_manual, tab_auto, tab_list = st.tabs([
        "📁 ایمپورت از اکسل",
        "➕ افزودن دستی",
        "📥 دریافت chat_id از بله",
        "📋 لیست مخاطبین",
    ])

    with tab_import:
        st.markdown("""
        **فرمت فایل اکسل/CSV:**
        | نام | ایمیل | نوع | chat_id |
        |---|---|---|---|
        | علی رضایی | ali@okco.ir | supervisor | 123456789 |
        | مدیر عامل | ceo@okco.ir | manager | 987654321 |
        """)
        up_file = st.file_uploader("انتخاب فایل", type=["xlsx", "csv"],
                                    key=f"rs_contacts_up_{key_suffix}")
        if up_file and st.button("📥 ایمپورت کن", key=f"rs_import_{key_suffix}", type="primary"):
            try:
                new_c = parse_contacts_file(up_file.getvalue(), up_file.name)
                existing_emails = {c.get("email") for c in contacts if c.get("email")}
                existing_chats = {c.get("bale_chat_id") for c in contacts if c.get("bale_chat_id")}
                added = 0
                for c in new_c:
                    e = c.get("email")
                    cid = c.get("bale_chat_id")
                    if e and e in existing_emails:
                        continue
                    if cid and cid in existing_chats:
                        continue
                    contacts.append(c)
                    if e:
                        existing_emails.add(e)
                    if cid:
                        existing_chats.add(cid)
                    added += 1
                save_contacts(contacts)
                st.success(f"✅ {added} مخاطب اضافه شد")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    with tab_manual:
        with st.form(f"rs_add_form_{key_suffix}", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                n = st.text_input("نام")
            with c2:
                e = st.text_input("ایمیل")
            c3, c4 = st.columns(2)
            with c3:
                ch = st.text_input("chat_id بله")
            with c4:
                tl = st.selectbox("نوع", list(TYPE_LABELS.values()), index=0)
            sub = st.form_submit_button("➕ افزودن", type="primary", use_container_width=True)
            if sub:
                n = (n or "").strip()
                e = (e or "").strip()
                ch = (ch or "").strip()
                t = LABEL_TO_TYPE.get(tl, "regular")
                cid = None
                if ch:
                    try:
                        cid = int(float(ch))
                    except (ValueError, TypeError):
                        st.error("chat_id باید عدد باشه")
                if not n:
                    st.error("نام رو وارد کن")
                elif not e and not cid:
                    st.error("حداقل ایمیل یا chat_id لازمه")
                else:
                    entry = {"name": n, "email": e, "type": t}
                    if cid:
                        entry["bale_chat_id"] = cid
                    contacts.append(entry)
                    save_contacts(contacts)
                    st.success("✅ اضافه شد")
                    st.rerun()

    with tab_auto:
        try:
            from bale_chat_id_helper import render_chat_id_fetcher
            render_chat_id_fetcher(key_suffix)
        except Exception as e:
            st.error(f"❌ {e}")

    with tab_list:
        if not contacts:
            st.info("خالی")
        else:
            st.markdown(f"**{len(contacts)} مخاطب**")
            fo = ["همه"] + list(TYPE_LABELS.values())
            sf = st.selectbox("فیلتر", fo, key=f"rs_filter_{key_suffix}")
            if sf == "همه":
                dc = list(enumerate(contacts))
            else:
                ft = LABEL_TO_TYPE[sf]
                dc = [(i, c) for i, c in enumerate(contacts) if c["type"] == ft]

            edit_key = f"rs_edit_idx_{key_suffix}"

            for i, c in dc:
                a, b, cc, d, e = st.columns([3, 3, 2, 1, 1])
                with a:
                    st.markdown(f"**{c.get('name') or '—'}**")
                    if c.get("bale_chat_id"):
                        st.caption(f"📱 {c['bale_chat_id']}")
                with b:
                    st.markdown(f"`{c.get('email') or '—'}`")
                with cc:
                    st.markdown(TYPE_LABELS.get(c["type"], c["type"]))
                if d.button("✏️", key=f"rs_edit_{i}_{key_suffix}", use_container_width=True):
                    st.session_state[edit_key] = i
                    st.rerun()
                if e.button("🗑", key=f"rs_del_{i}_{key_suffix}", use_container_width=True):
                    contacts.pop(i)
                    save_contacts(contacts)
                    st.session_state.pop(edit_key, None)
                    st.rerun()

            if st.session_state.get(edit_key) is not None:
                idx = st.session_state[edit_key]
                if 0 <= idx < len(contacts):
                    c = contacts[idx]
                    st.divider()
                    st.markdown(f"### ✏️ ویرایش: {c.get('name') or c.get('email')}")

                    with st.form(f"rs_edit_form_{key_suffix}"):
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            en = st.text_input("نام", value=c.get("name", ""))
                        with ec2:
                            ee = st.text_input("ایمیل", value=c.get("email", ""))

                        ec3, ec4 = st.columns(2)
                        with ec3:
                            ech = st.text_input("chat_id بله",
                                                 value=str(c.get("bale_chat_id", "") or ""))
                        with ec4:
                            current_label = TYPE_LABELS.get(c.get("type"), TYPE_LABELS["regular"])
                            etl = st.selectbox("نوع", list(TYPE_LABELS.values()),
                                                index=list(TYPE_LABELS.values()).index(current_label))

                        btn1, btn2 = st.columns(2)
                        with btn1:
                            save = st.form_submit_button("💾 ذخیره", type="primary",
                                                          use_container_width=True)
                        with btn2:
                            cancel = st.form_submit_button("❌ انصراف", use_container_width=True)

                        if save:
                            new_entry = {
                                "name": (en or "").strip(),
                                "email": (ee or "").strip(),
                                "type": LABEL_TO_TYPE.get(etl, "regular"),
                            }
                            ch_raw = (ech or "").strip()
                            if ch_raw:
                                try:
                                    new_entry["bale_chat_id"] = int(float(ch_raw))
                                except (ValueError, TypeError):
                                    st.error("chat_id باید عدد باشه")
                                    st.stop()
                            if not new_entry["name"]:
                                st.error("نام نمی‌تونه خالی باشه")
                            else:
                                contacts[idx] = new_entry
                                save_contacts(contacts)
                                st.session_state.pop(edit_key, None)
                                st.success("✅ ویرایش شد")
                                st.rerun()

                        if cancel:
                            st.session_state.pop(edit_key, None)
                            st.rerun()

            st.divider()
            if st.button("🗑 پاک کردن همه", key=f"rs_clear_{key_suffix}"):
                save_contacts([])
                st.session_state.pop(edit_key, None)
                st.rerun()
