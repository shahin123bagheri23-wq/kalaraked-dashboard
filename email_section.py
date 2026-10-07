"""بخش ارسال ایمیل — داخل صفحه آپلود"""
from datetime import datetime

import pandas as pd
import streamlit as st

from email_manager import (
    load_contacts, save_contacts, parse_contacts_file,
)


TYPE_LABELS = {
    "supervisor": "👤 سوپروایزر",
    "manager": "🏆 مدیر",
    "regular": "📬 عادی",
}
LABEL_TO_TYPE = {v: k for k, v in TYPE_LABELS.items()}


def _send_emails(html, subject, recipients):
    from send_report import send_email
    if not recipients:
        raise ValueError("لیست گیرندگان خالیه")
    return send_email(html, subject=subject)


def _build_manager_report():
    from send_report import build_report
    return build_report()


def _build_regular_report():
    from send_report import EXCEL_FILE, find_sheet, parse_percent, normalize_name

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"فایل اکسل پیدا نشد: {EXCEL_FILE}")

    xls = pd.ExcelFile(EXCEL_FILE)
    sh = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    if not sh:
        raise ValueError("شیت تارگت پیدا نشد")

    t = pd.read_excel(xls, sheet_name=sh)
    t.columns = t.columns.astype(str).str.strip()

    ach_col = next((c for c in t.columns if "تحقق" in c), None)
    branch_col = next((c for c in t.columns if "نام شعبه" in c), None)
    if not ach_col or not branch_col:
        raise ValueError("ستون‌های لازم پیدا نشد")

    t[branch_col] = normalize_name(t[branch_col])
    t[ach_col] = parse_percent(t[ach_col])
    vals = t[ach_col].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success = int((vals >= 100).sum())
    warning = int(((vals >= 80) & (vals < 100)).sum())
    danger = int((vals < 80).sum())

    html = f"""
    <html dir="rtl"><body style="font-family:Tahoma;background:#f4f4f4;padding:20px;">
      <div style="max-width:600px;margin:auto;background:#fff;border-radius:12px;overflow:hidden;">
        <div style="background:#E6003E;color:white;padding:20px;text-align:center;">
          <h1 style="margin:0;">گزارش خلاصه کالای راکد</h1>
          <p style="margin:6px 0 0 0;font-size:13px;">افق کوروش — {datetime.now().strftime('%Y-%m-%d')}</p>
        </div>
        <div style="padding:20px;">
          <p><b>میانگین تحقق:</b> {mean_ach:.1f}%</p>
          <p>🟢 موفق: {success} شعبه</p>
          <p>🟡 در حال پیشرفت: {warning} شعبه</p>
          <p>🔴 بحرانی: {danger} شعبه</p>
          <p style="margin-top:20px;color:#888;font-size:12px;">
            برای جزئیات بیشتر به داشبورد مراجعه کنید.
          </p>
        </div>
      </div>
    </body></html>
    """
    return html, {"mean_ach": mean_ach, "success": success,
                  "warning": warning, "danger": danger}


def _build_supervisor_report(supervisor_name):
    from send_report import (
        EXCEL_FILE, find_sheet, parse_percent, money, normalize_name, to_number
    )

    if not EXCEL_FILE.exists():
        raise FileNotFoundError("فایل اکسل پیدا نشد")

    xls = pd.ExcelFile(EXCEL_FILE)
    sh = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    t = pd.read_excel(xls, sheet_name=sh)
    t.columns = t.columns.astype(str).str.strip()

    branch_col = next((c for c in t.columns if "نام شعبه" in c), None)
    ach_col = next((c for c in t.columns if "تحقق" in c), None)
    sup_col = next((c for c in t.columns if c.strip() == "سوپروایزر"), None)
    raked_col = next((c for c in t.columns if "ریالی راکد" in c), None)
    target_col = next((c for c in t.columns if c.startswith("تارگت") and "تغییرات" not in c), None)
    change_col = next((c for c in t.columns if c.startswith("تغییرات") and "تارگت" not in c), None)

    if not sup_col or not branch_col:
        raise ValueError("ستون‌های لازم پیدا نشد")

    t[sup_col] = normalize_name(t[sup_col])
    t[branch_col] = normalize_name(t[branch_col])
    t[ach_col] = parse_percent(t[ach_col])
    if raked_col:
        t[raked_col] = to_number(t[raked_col])
    if target_col:
        t[target_col] = parse_percent(t[target_col])
    if change_col:
        t[change_col] = parse_percent(t[change_col])

    sub = t[t[sup_col].astype(str).str.strip() == supervisor_name.strip()].copy()
    if sub.empty:
        raise ValueError(f"برای سوپروایزر «{supervisor_name}» شعبه‌ای پیدا نشد")

    sub = sub.sort_values(ach_col, ascending=False).reset_index(drop=True)
    sub.insert(0, "رتبه", range(1, len(sub) + 1))

    display_cols = [c for c in [branch_col, raked_col, target_col, ach_col, change_col] if c]
    display = sub[display_cols].copy()

    def _fmt_p(v):
        if pd.isna(v):
            return ""
        try:
            x = float(v)
            return "0.0%" if abs(x) < 0.05 else f"{x:.1f}%"
        except Exception:
            return str(v)

    col_map = {
        branch_col: "شعبه",
        raked_col: "ارزش راکد",
        target_col: "تارگت",
        ach_col: "تحقق",
        change_col: "تغییرات",
    }
    col_map = {k: v for k, v in col_map.items() if k}

    if raked_col:
        display[raked_col] = display[raked_col].apply(lambda v: money(v))
    for c in display.columns:
        if any(k in c for k in ["تحقق", "تارگت", "تغییرات"]):
            display[c] = display[c].apply(_fmt_p)

    display = display.rename(columns=col_map)
    table_html = display.to_html(index=False, classes="table", border=0, escape=False)

    mean_ach = sub[ach_col].mean()
    total_raked = sub[raked_col].sum() if raked_col else 0

    html = f"""
    <html dir="rtl"><head><meta charset="UTF-8">
    <style>
      body {{ font-family:Tahoma;background:#f4f4f4;padding:20px; }}
      .container {{ max-width:800px;margin:auto;background:#fff;border-radius:12px;overflow:hidden; }}
      .header {{ background:#E6003E;color:white;padding:20px;text-align:center; }}
      .content {{ padding:20px; }}
      .table {{ width:100%;border-collapse:collapse;font-size:12px;margin-top:12px; }}
      .table th {{ background:#E6003E;color:white;padding:8px;text-align:right; }}
      .table td {{ border-bottom:1px solid #eee;padding:8px;text-align:right; }}
      .kpi {{ background:#fafafa;border-right:4px solid #E6003E;padding:12px;border-radius:8px;margin-bottom:12px; }}
    </style></head>
    <body>
      <div class="container">
        <div class="header">
          <h1 style="margin:0;">گزارش سوپروایزر</h1>
          <p style="margin:6px 0 0 0;font-size:13px;">{supervisor_name} — {datetime.now().strftime('%Y-%m-%d')}</p>
        </div>
        <div class="content">
          <div class="kpi"><b>میانگین تحقق شما:</b> {mean_ach:.1f}%</div>
          <div class="kpi"><b>مجموع راکد شعبه‌ها:</b> {money(total_raked)}</div>
          <div class="kpi"><b>تعداد شعبه:</b> {len(sub)}</div>
          <h2>شعبه‌های تحت پوشش</h2>
          <div style="overflow-x:auto;">{table_html}</div>
        </div>
      </div>
    </body></html>
    """
    return html, {"mean_ach": mean_ach, "branches": len(sub), "total_raked": total_raked}


# ================== صفحه اصلی ==================
def render_email_section_inside_upload(key_suffix="up"):
    st.markdown("### 📧 ارسال ایمیل گزارش")
    st.caption("مدیریت مخاطبین — دستی یا ایمپورت از اکسل — و ارسال گزارش به سه گروه")

    contacts = load_contacts()

    # ================== مدیریت مخاطبین ==================
    with st.expander("👥 مدیریت مخاطبین", expanded=not contacts):
        tab_import, tab_manual, tab_list = st.tabs([
            "📁 ایمپورت از اکسل",
            "➕ افزودن دستی",
            "📋 لیست مخاطبین",
        ])

        # ---------- تب ایمپورت ----------
        with tab_import:
            st.markdown("""
            **فرمت فایل اکسل/CSV:**
            | نام | ایمیل | نوع |
            |---|---|---|
            | علی رضایی | ali@okco.ir | supervisor |
            | مدیر عامل | ceo@okco.ir | manager |
            | سارا احمدی | sara@okco.ir | regular |
            """)

            up_file = st.file_uploader(
                "انتخاب فایل",
                type=["xlsx", "csv"],
                key=f"contacts_up_{key_suffix}",
            )

            if up_file and st.button("📥 ایمپورت کن", key=f"import_btn_{key_suffix}", type="primary"):
                try:
                    new_c = parse_contacts_file(up_file.getvalue(), up_file.name)
                    existing = {c["email"] for c in contacts}
                    added = 0
                    for c in new_c:
                        if c["email"] not in existing:
                            contacts.append(c)
                            existing.add(c["email"])
                            added += 1
                    save_contacts(contacts)
                    st.success(f"✅ {added} مخاطب اضافه شد — کل: {len(contacts)}")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")

        # ---------- تب افزودن دستی ----------
        with tab_manual:
            st.markdown("**افزودن یک مخاطب جدید**")

            with st.form(f"add_contact_form_{key_suffix}", clear_on_submit=True):
                col_a, col_b = st.columns(2)
                with col_a:
                    m_name = st.text_input("نام", placeholder="مثلاً علی ابراهیمی نیا")
                with col_b:
                    m_email = st.text_input("ایمیل", placeholder="ali@okco.ir")

                m_type_label = st.selectbox(
                    "نوع",
                    list(TYPE_LABELS.values()),
                    index=0,
                )

                submitted = st.form_submit_button("➕ افزودن", type="primary", use_container_width=True)

                if submitted:
                    e = (m_email or "").strip()
                    n = (m_name or "").strip()
                    t = LABEL_TO_TYPE.get(m_type_label, "regular")

                    if not e:
                        st.error("ایمیل رو وارد کن")
                    elif "@" not in e or "." not in e:
                        st.error("فرمت ایمیل درست نیست")
                    elif any(c["email"] == e for c in contacts):
                        st.error("این ایمیل قبلاً اضافه شده")
                    else:
                        contacts.append({"name": n, "email": e, "type": t})
                        save_contacts(contacts)
                        st.success(f"✅ {e} اضافه شد")
                        st.rerun()

        # ---------- تب لیست ----------
        with tab_list:
            if not contacts:
                st.info("هنوز مخاطبی نداری.")
            else:
                st.markdown(f"**{len(contacts)} مخاطب**")

                # فیلتر بر اساس نوع
                filter_options = ["همه"] + list(TYPE_LABELS.values())
                selected_filter = st.selectbox(
                    "🔍 فیلتر بر اساس نوع",
                    filter_options,
                    key=f"contact_filter_{key_suffix}",
                )

                if selected_filter == "همه":
                    display_contacts = list(enumerate(contacts))
                else:
                    ft = LABEL_TO_TYPE[selected_filter]
                    display_contacts = [
                        (i, c) for i, c in enumerate(contacts) if c["type"] == ft
                    ]

                if not display_contacts:
                    st.info("مخاطبی با این فیلتر پیدا نشد.")
                else:
                    for i, c in display_contacts:
                        col1, col2, col3, col4 = st.columns([3, 3, 2, 1])
                        col1.markdown(f"**{c.get('name') or '—'}**")
                        col2.markdown(f"`{c['email']}`")
                        col3.markdown(TYPE_LABELS.get(c["type"], c["type"]))
                        if col4.button("🗑", key=f"del_contact_{i}_{key_suffix}",
                                       use_container_width=True):
                            contacts.pop(i)
                            save_contacts(contacts)
                            st.rerun()

                if st.button("🗑 پاک کردن همه مخاطبین",
                             key=f"clear_all_{key_suffix}",
                             type="secondary"):
                    save_contacts([])
                    st.rerun()

    # ================== ادامه ==================
    if not contacts:
        st.info("📭 برای ارسال ایمیل، اول یک مخاطب اضافه کن.")
        return

    supervisors = [c for c in contacts if c["type"] == "supervisor"]
    managers = [c for c in contacts if c["type"] == "manager"]
    regulars = [c for c in contacts if c["type"] == "regular"]

    st.markdown(
        f"**📊 مخاطبین:** 👤 {len(supervisors)} سوپروایزر | "
        f"🏆 {len(managers)} مدیر | 📬 {len(regulars)} عادی"
    )

    st.divider()

    tab_sup, tab_mgr, tab_reg = st.tabs([
        f"👤 سوپروایزرها ({len(supervisors)})",
        f"🏆 مدیران ({len(managers)})",
        f"📬 افراد عادی ({len(regulars)})",
    ])

    with tab_sup:
        _tab_supervisors(supervisors, key_suffix)
    with tab_mgr:
        _tab_managers(managers, key_suffix)
    with tab_reg:
        _tab_regulars(regulars, key_suffix)


# ================== تب سوپروایزر ==================
def _tab_supervisors(supervisors, key_suffix):
    st.markdown("#### 👤 ارسال به سوپروایزرها")
    st.caption("هر سوپروایزر فقط گزارش شعبه‌های خودش رو دریافت می‌کنه")

    if not supervisors:
        st.info("هیچ سوپروایزری در لیست نیست.")
        return

    sup_names = [(c["name"] or c["email"], c) for c in supervisors]
    selected_label = st.selectbox(
        "🔍 پیش‌نمایش کدام سوپروایزر؟",
        ["(هیچ‌کدام)"] + [n for n, _ in sup_names],
        key=f"sup_sel_{key_suffix}",
    )

    if selected_label != "(هیچ‌کدام)":
        contact = next(c for n, c in sup_names if n == selected_label)
        actual_name = contact["name"] or contact["email"]

        c1, c2 = st.columns(2)
        with c1:
            if st.button("👁 پیش‌نمایش", key=f"sup_prev_{key_suffix}", use_container_width=True):
                try:
                    html, stats = _build_supervisor_report(actual_name)
                    st.session_state[f"sup_prev_html_{key_suffix}"] = html
                    st.session_state[f"sup_prev_stats_{key_suffix}"] = stats
                except Exception as e:
                    st.error(f"❌ {e}")
        with c2:
            if st.button("📤 ارسال", key=f"sup_send_{key_suffix}",
                         type="primary", use_container_width=True):
                try:
                    html, _ = _build_supervisor_report(actual_name)
                    subject = f"گزارش عملکرد — {actual_name}"
                    sent = _send_emails(html, subject, [contact["email"]])
                    st.success(f"✅ ارسال به {', '.join(sent)}")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ {e}")

        if st.session_state.get(f"sup_prev_html_{key_suffix}"):
            with st.expander("👁 پیش‌نمایش", expanded=True):
                stats = st.session_state.get(f"sup_prev_stats_{key_suffix}", {})
                if stats:
                    st.caption(
                        f"میانگین تحقق: {stats.get('mean_ach', 0):.1f}% | "
                        f"تعداد شعبه: {stats.get('branches', 0)}"
                    )
                st.components.v1.html(
                    st.session_state[f"sup_prev_html_{key_suffix}"],
                    height=500, scrolling=True,
                )

    st.divider()
    st.markdown("#### 📤 ارسال گروهی به همه سوپروایزرها")

    subject_all = st.text_input(
        "موضوع ایمیل",
        value="گزارش عملکرد هفتگی — کالای راکد",
        key=f"sup_subj_{key_suffix}",
    )

    if st.button("📨 ارسال به همه", key=f"sup_all_{key_suffix}", type="primary"):
        progress = st.progress(0)
        ok, fail = [], []
        for i, c in enumerate(supervisors):
            progress.progress((i + 1) / len(supervisors),
                              text=f"ارسال به {c['name'] or c['email']}...")
            try:
                name = c["name"] or c["email"]
                html, _ = _build_supervisor_report(name)
                _send_emails(html, subject_all, [c["email"]])
                ok.append(c["email"])
            except Exception as e:
                fail.append(f"{c['email']}: {e}")
        progress.empty()
        if ok:
            st.success(f"✅ ارسال موفق به {len(ok)} سوپروایزر")
        if fail:
            with st.expander(f"❌ {len(fail)} خطا"):
                for f in fail:
                    st.error(f)


# ================== تب مدیر ==================
def _tab_managers(managers, key_suffix):
    st.markdown("#### 🏆 ارسال گزارش مدیریتی")
    st.caption("گزارش کامل با KPI، رتبه‌بندی شعب، سرپرست‌ها و سوپروایزرهای منطقه")

    if not managers:
        st.info("هیچ مدیری در لیست نیست.")
        return

    emails = [c["email"] for c in managers]
    st.markdown(f"**گیرندگان ({len(managers)}):** " + "، ".join(emails))

    subject = st.text_input(
        "موضوع ایمیل",
        value="گزارش مدیریتی هفتگی — کالای راکد افق کوروش",
        key=f"mgr_subj_{key_suffix}",
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"mgr_prev_{key_suffix}", use_container_width=True):
            try:
                html, stats = _build_manager_report()
                st.session_state[f"mgr_prev_html_{key_suffix}"] = html
                st.session_state[f"mgr_prev_stats_{key_suffix}"] = stats
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📤 ارسال به همه مدیران", key=f"mgr_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                html, _ = _build_manager_report()
                sent = _send_emails(html, subject, emails)
                st.success(f"✅ ارسال به {', '.join(sent)}")
                st.balloons()
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"mgr_prev_html_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            stats = st.session_state.get(f"mgr_prev_stats_{key_suffix}", {})
            if stats:
                st.caption(
                    f"میانگین تحقق: {stats.get('mean_ach', 0):.1f}% | "
                    f"موفق: {stats.get('success', 0)} | "
                    f"بحرانی: {stats.get('danger', 0)} | "
                    f"کل: {stats.get('branches', 0)}"
                )
            st.components.v1.html(
                st.session_state[f"mgr_prev_html_{key_suffix}"],
                height=700, scrolling=True,
            )


# ================== تب عادی ==================
def _tab_regulars(regulars, key_suffix):
    st.markdown("#### 📬 ارسال گزارش خلاصه")
    st.caption("گزارش کوتاه شامل KPI کلی")

    if not regulars:
        st.info("هیچ فرد عادی در لیست نیست.")
        return

    emails = [c["email"] for c in regulars]
    st.markdown(f"**گیرندگان ({len(regulars)}):** " + "، ".join(emails))

    subject = st.text_input(
        "موضوع ایمیل",
        value="گزارش خلاصه کالای راکد — افق کوروش",
        key=f"reg_subj_{key_suffix}",
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁 پیش‌نمایش", key=f"reg_prev_{key_suffix}", use_container_width=True):
            try:
                html, stats = _build_regular_report()
                st.session_state[f"reg_prev_html_{key_suffix}"] = html
            except Exception as e:
                st.error(f"❌ {e}")
    with c2:
        if st.button("📤 ارسال به همه", key=f"reg_send_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                html, _ = _build_regular_report()
                sent = _send_emails(html, subject, emails)
                st.success(f"✅ ارسال به {', '.join(sent)}")
                st.balloons()
            except Exception as e:
                st.error(f"❌ {e}")

    if st.session_state.get(f"reg_prev_html_{key_suffix}"):
        with st.expander("👁 پیش‌نمایش", expanded=True):
            st.components.v1.html(
                st.session_state[f"reg_prev_html_{key_suffix}"],
                height=400, scrolling=True,
            )