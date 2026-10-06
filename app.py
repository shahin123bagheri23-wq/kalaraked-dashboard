import streamlit as st
import pandas as pd
from streamlit_js_eval import streamlit_js_eval

st.set_page_config(
    page_title="داشبورد کالای راکد",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ================== تشخیص دستگاه ==================
screen_width = streamlit_js_eval(js_expressions='window.innerWidth', key='WIDTH')
is_mobile = (screen_width or 1200) < 768

# ================== استایل ==================
st.markdown("""
<style>
    html, body, [class*="css"] { font-family: Tahoma, sans-serif; }
    .block-container {
        padding-top: 1rem; padding-bottom: 1rem;
        padding-left: 0.6rem; padding-right: 0.6rem;
    }
    div[data-testid="stMetric"] {
        background-color: #1e1e1e; padding: 10px;
        border-radius: 10px; margin-bottom: 8px;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

if is_mobile:
    st.markdown("""
    <style>
        h1 { font-size: 1.4rem !important; }
        h2 { font-size: 1.2rem !important; }
        h3 { font-size: 1.05rem !important; }
    </style>
    """, unsafe_allow_html=True)

# ================== تنظیمات فایل‌ها ==================
RAAKED_FILE = "اقلام راکد 60  45   دیتا (1).xlsx"
SALES_FILE = "گزارش فروش راکد.xlsx"

SALES_BARCODE = "بارکد کالا"
SALES_QTY = "فروش تعدادی"
SALES_AMOUNT = "فروش خالص با مالیات"
SALES_STORE = "نام انبار/فروشگاه"


# ================== توابع کمکی ==================
def clean_barcode(series):
    return (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .replace("nan", "")
    )


def normalize_name(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace("\u200c", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )


def remove_totals(df):
    if "نام شعبه" not in df.columns:
        return df
    mask = ~df["نام شعبه"].astype(str).str.contains("جمع|مجموع|Total|کل", case=False, na=False)
    mask &= df["نام شعبه"].notna()
    return df[mask].copy()


@st.cache_data
def load_raaked():
    d60 = pd.read_excel(RAAKED_FILE, sheet_name="60 روزه", header=9)
    d45 = pd.read_excel(RAAKED_FILE, sheet_name="45 روزه ", header=9)
    d60.columns = d60.columns.str.strip()
    d45.columns = d45.columns.str.strip()
    d60 = remove_totals(d60)
    d45 = remove_totals(d45)
    d60["بارکد"] = clean_barcode(d60["بارکد"])
    d45["بارکد"] = clean_barcode(d45["بارکد"])
    d60["نام شعبه"] = normalize_name(d60["نام شعبه"])
    d45["نام شعبه"] = normalize_name(d45["نام شعبه"])
    return d60, d45


@st.cache_data
def load_sales():
    try:
        s = pd.read_excel(SALES_FILE, sheet_name="Sheet1")
        s.columns = s.columns.str.strip()
        s[SALES_BARCODE] = clean_barcode(s[SALES_BARCODE])
        s[SALES_STORE] = normalize_name(s[SALES_STORE])
        return s
    except Exception:
        return None


# ================== خواندن داده ==================
df60, df45 = load_raaked()
df_sales = load_sales()

st.title("📊 داشبورد مدیریت کالای راکد")
st.caption(f"📱 حالت: {'موبایل' if is_mobile else 'کامپیوتر'}  |  عرض صفحه: {screen_width}px")

if df_sales is None:
    st.warning(f"⚠️ فایل فروش «{SALES_FILE}» پیدا نشد یا خوانده نشد.")
else:
    st.success(f"✅ فایل فروش بارگذاری شد ({len(df_sales):,} ردیف)")

st.divider()

# ================== فیلترها ==================
st.subheader("🔎 فیلترها")

supervisors = ["همه"] + sorted(df60["سوپروایزر"].dropna().unique().tolist())

if is_mobile:
    selected_supervisor = st.selectbox("انتخاب سوپروایزر", supervisors)
    if selected_supervisor == "همه":
        branch_options = sorted(df60["نام شعبه"].dropna().unique().tolist())
    else:
        branch_options = sorted(
            df60.loc[df60["سوپروایزر"] == selected_supervisor, "نام شعبه"]
            .dropna().unique().tolist()
        )
    branches = ["همه"] + branch_options
    selected_branch = st.selectbox("انتخاب شعبه", branches)
else:
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        selected_supervisor = st.selectbox("انتخاب سوپروایزر", supervisors)
    with col_f2:
        if selected_supervisor == "همه":
            branch_options = sorted(df60["نام شعبه"].dropna().unique().tolist())
        else:
            branch_options = sorted(
                df60.loc[df60["سوپروایزر"] == selected_supervisor, "نام شعبه"]
                .dropna().unique().tolist()
            )
        branches = ["همه"] + branch_options
        selected_branch = st.selectbox("انتخاب شعبه", branches)


def apply_filter(df):
    d = df.copy()
    if selected_branch != "همه":
        d = d[d["نام شعبه"] == selected_branch]
    if selected_supervisor != "همه":
        d = d[d["سوپروایزر"] == selected_supervisor]
    return d


def get_branch_sales():
    if df_sales is None or df_sales.empty:
        return None
    sm = df_sales.copy()
    if selected_branch != "همه":
        sm = sm[sm[SALES_STORE] == selected_branch]
    return sm


f60 = apply_filter(df60)
f45 = apply_filter(df45)
branch_sales = get_branch_sales()

st.divider()

# ================== KPIها ==================
total60 = f60["موجودی ریالی اقلام راکد"].sum()
total45 = f45["موجودی ریالی اقلام راکد"].sum()

if is_mobile:
    col1, col2 = st.columns(2)
    col1.metric("ارزش راکد ۶۰ روزه", f"{total60:,.0f} ریال")
    col2.metric("ارزش راکد ۴۵ روزه", f"{total45:,.0f} ریال")
    col3, col4 = st.columns(2)
    col3.metric("تعداد اقلام ۶۰ روزه", f"{len(f60):,}")
    col4.metric("تعداد اقلام ۴۵ روزه", f"{len(f45):,}")
else:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("ارزش راکد ۶۰ روزه", f"{total60:,.0f} ریال")
    col2.metric("ارزش راکد ۴۵ روزه", f"{total45:,.0f} ریال")
    col3.metric("تعداد اقلام ۶۰ روزه", f"{len(f60):,}")
    col4.metric("تعداد اقلام ۴۵ روزه", f"{len(f45):,}")

st.divider()


# ================== توابع جدول ==================
def build_sales_map(sales_df):
    if sales_df is None or sales_df.empty:
        return None
    return (
        sales_df.groupby([SALES_STORE, SALES_BARCODE])
        .agg({SALES_QTY: "sum", SALES_AMOUNT: "sum"})
        .reset_index()
        .rename(columns={
            SALES_STORE: "نام شعبه",
            SALES_BARCODE: "بارکد",
            SALES_QTY: "فروش (تعداد)",
            SALES_AMOUNT: "فروش (ریال)"
        })
    )


def attach_sales(df_raaked):
    cols = ["نام شعبه", "بارکد", "نام کالا", "موجودی سیستمی",
            "موجودی ریالی اقلام راکد", "سوپروایزر"]
    out = df_raaked[cols].copy()

    sales_map = build_sales_map(branch_sales)

    if sales_map is not None:
        out = out.merge(sales_map, on=["نام شعبه", "بارکد"], how="left")
        out["فروش (تعداد)"] = out["فروش (تعداد)"].fillna(0).astype(int)
        out["فروش (ریال)"] = out["فروش (ریال)"].fillna(0).astype(int)
    else:
        out["فروش (تعداد)"] = 0
        out["فروش (ریال)"] = 0

    out = out.sort_values("موجودی ریالی اقلام راکد", ascending=False).reset_index(drop=True)
    return out


def column_config(df):
    """فقط فرمت اعداد؛ عرض را Streamlit خودش تنظیم می‌کند."""
    cfg = {}
    for col in df.columns:
        if col == "موجودی سیستمی":
            cfg[col] = st.column_config.NumberColumn(col, format="%d")
        elif col == "موجودی ریالی اقلام راکد":
            cfg[col] = st.column_config.NumberColumn(col, format="%,d")
        elif col == "فروش (تعداد)":
            cfg[col] = st.column_config.NumberColumn(col, format="%d")
        elif col == "فروش (ریال)":
            cfg[col] = st.column_config.NumberColumn(col, format="%,d")
        else:
            cfg[col] = st.column_config.Column(col)
    return cfg


def show_table(df, key):
    """نمایش جدول با کپی و فیت خودکار."""
    view = df.copy()

    if selected_supervisor != "همه" and "سوپروایزر" in view.columns:
        view = view.drop(columns=["سوپروایزر"])

    cfg = column_config(view)

    st.dataframe(
        view,
        use_container_width=True,
        hide_index=True,
        column_config=cfg,
        on_select="rerun",
        selection_mode="multi-row",
        key=key
    )

    csv = view.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ دانلود این جدول (CSV)",
        csv,
        f"{key}.csv",
        "text/csv",
        key=f"dl_{key}"
    )


# ================== ۲۰ کالای راکد برتر ==================
st.subheader("🚨 ۲۰ کالای راکد برتر")
top_n = st.slider("چند قلم نمایش داده شود؟", 5, 50, 20, step=5)

tab1, tab2 = st.tabs(["۶۰ روزه", "۴۵ روزه"])

with tab1:
    top60 = attach_sales(f60).head(top_n)
    show_table(top60, "top60")

with tab2:
    top45 = attach_sales(f45).head(top_n)
    show_table(top45, "top45")

st.divider()

# ================== کالاهای راکدی که فروش رفتند ==================
st.subheader("✅ کالاهای راکدی که فروش رفتند")

if branch_sales is None or branch_sales.empty:
    st.info("برای مشاهده این بخش، فایل فروش لازم است یا برای این شعبه فروشی ثبت نشده.")
else:
    def matched_with_sales(df_raaked):
        cols = ["نام شعبه", "بارکد", "نام کالا", "موجودی سیستمی",
                "موجودی ریالی اقلام راکد", "سوپروایزر"]
        out = df_raaked[cols].copy()

        sales_map = build_sales_map(branch_sales)
        merged = out.merge(sales_map, on=["نام شعبه", "بارکد"], how="inner")
        merged = merged.sort_values("فروش (ریال)", ascending=False).reset_index(drop=True)
        return merged

    matched60 = matched_with_sales(f60)
    matched45 = matched_with_sales(f45)

    col_m1, col_m2 = st.columns(2)
    col_m1.metric("کالای راکد ۶۰ روزه که فروش رفت", f"{len(matched60):,}")
    col_m2.metric("کالای راکد ۴۵ روزه که فروش رفت", f"{len(matched45):,}")

    tabm1, tabm2 = st.tabs(["۶۰ روزه — فروش‌رفته", "۴۵ روزه — فروش‌رفته"])

    with tabm1:
        if matched60.empty:
            st.info("از اقلام راکد ۶۰ روزه، دیروز چیزی فروش نرفت.")
        else:
            show_table(matched60, "matched60")

    with tabm2:
        if matched45.empty:
            st.info("از اقلام راکد ۴۵ روزه، دیروز چیزی فروش نرفت.")
        else:
            show_table(matched45, "matched45")