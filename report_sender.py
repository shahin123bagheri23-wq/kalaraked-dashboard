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
    cfg.read("config.ini", encoding="utf-8")
    return cfg.get("ADMIN", "password", fallback="admin1234")


# ================== ارسال‌کننده‌های بله ==================
def _send_bale(chat_id, text):
    from bale_sender import send_message
    return send_message(chat_id, text)


# ================== سازنده متن‌های بله ==================
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

    sub = t[t[sup_col].astype(str).str.strip() == supervisor_name.strip()].copy()
    if sub.empty:
        raise ValueError(f"شعبه‌ای برای «{supervisor_name}» پیدا نشد")

    sub = sub.sort_values(ach_col, ascending=False).reset_index(drop=True)
    mean_ach = sub[ach_col].mean()
    total_raked = sub[raked_col].sum() if raked_col else 0

    # رتبه سوپروایزر
    all_sup = t.groupby(sup_col)[ach_col].mean().sort_values(ascending=False)
    rank = "?"
    if supervisor_name in all_sup.index:
        rank = list(all_sup.index).index(supervisor_name) + 1
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
def render_report_sender(key_suffix="rs"):
    """صفحه ارسال گزارش — با رمز محافظت‌شده"""
    st.markdown("## 📨 ارسال گزارش")
    st.caption("ارسال گزارش به سه گروه از طریق کانال‌های مختلف")

    # ================== رمز ورود ==================
    state_key = f"report_auth_{key_suffix}"
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

    tab_bale, tab_email, tab_contacts = st.tabs([
        "📱 بله",
        "📧 ایمیل",
        "👥 مدیریت مخاطبین",
    ])

    with tab_bale:
        _render_bale_channel(supervisors, managers, regulars, key_suffix)

    with tab_email:
        _render_email_channel(supervisors, managers, regulars, key_suffix)

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