"""
گزارش مدیریتی هفتگی — تحقق شعب، سرپرست فروشگاه و سوپروایزر منطقه
"""
import os
import os
import smtplib
import ssl
import configparser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from datetime import datetime

import pandas as pd
from jinja2 import Template


BASE_DIR = Path(__file__).parent
_EXCEL_NAME = "1405-07-15 projraked.xlsx"
_DATA_DIR = Path(os.environ.get("DATA_DIR") or BASE_DIR)
EXCEL_FILE = _DATA_DIR / _EXCEL_NAME
if not EXCEL_FILE.exists() and (BASE_DIR / _EXCEL_NAME).exists():
    EXCEL_FILE = BASE_DIR / _EXCEL_NAME
CONFIG_FILE = BASE_DIR / "config.ini"

FA_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
ACH_OK, ACH_WARN = 100, 80


# ================== توابع کمکی ==================
def normalize_name(s):
    return (s.astype(str).str.strip()
            .str.replace("\u200c", " ", regex=False)
            .str.replace("ي", "ی", regex=False)
            .str.replace("ك", "ک", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .replace({"nan": "", "None": ""}))


def to_number(series):
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0).astype(float)
    s = (series.astype(str).str.translate(FA_DIGITS)
         .str.replace(",", "", regex=False)
         .str.replace("،", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("\u200c", "", regex=False)
         .str.replace("%", "", regex=False)
         .replace({"nan": None, "": None, "-": None, "None": None}))
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(float)


def parse_percent(series):
    """تشخیص خودکار نسبت اعشاری (0.05) از درصد واقعی (5)"""
    raw = series.astype(str)
    has_sign = raw.str.contains("%").any()
    s = (raw.str.translate(FA_DIGITS)
         .str.replace(",", "", regex=False).str.replace("،", "", regex=False)
         .str.replace(" ", "", regex=False).str.replace("\u200c", "", regex=False)
         .str.replace("%", "", regex=False)
         .replace({"nan": None, "": None, "-": None, "None": None}))
    s = pd.to_numeric(s, errors="coerce").fillna(0)
    if (not has_sign and len(s) > 0
            and s.abs().quantile(0.95) <= 2.0):
        s = s * 100
    return s


def money(v, short=True):
    v = float(v or 0)
    if not short:
        return f"{v:,.0f} ریال"
    abs_v = abs(v)
    if abs_v >= 1_000_000_000:
        return f"{v / 1_000_000_000:,.1f} میلیارد"
    if abs_v >= 1_000_000:
        return f"{v / 1_000_000:,.1f} میلیون"
    if abs_v >= 1_000:
        return f"{v / 1_000:,.1f} هزار"
    return f"{v:,.0f}"


def find_sheet(xls, name):
    target = name.translate(FA_DIGITS).strip().lower()
    for s in xls.sheet_names:
        if s.translate(FA_DIGITS).strip().lower() == target:
            return s
    return None


def detect_target_cols(df):
    cols = list(df.columns)
    return {
        "branch": next((c for c in cols if "نام شعبه" in c), None),
        "code": next((c for c in cols if "کد" in c and "شعبه" in c), None),
        "supervisor": next((c for c in cols if "سرپرست فروشگاه" in c), None),
        "zone_supervisor": next((c for c in cols if c.strip() == "سوپروایزر"), None),
        "raked_value": next((c for c in cols if "ریالی راکد" in c), None),
        "target": next((c for c in cols if c.startswith("تارگت") and "تغییرات" not in c), None),
        "achievement": next((c for c in cols if "تحقق" in c), None),
        "change": next((c for c in cols if c.startswith("تغییرات") and "تارگت" not in c), None),
        "trend": [c for c in cols if "درصد" in c and "راکد" in c],
    }


# ================== ساخت گزارش مدیریتی ==================
def build_report():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"فایل اکسل پیدا نشد: {EXCEL_FILE}")

    xls = pd.ExcelFile(EXCEL_FILE)

    sh_target = find_sheet(xls, "تارگت") or find_sheet(xls, "روند و تارگت")
    if not sh_target:
        raise ValueError("شیت تارگت پیدا نشد")

    t = pd.read_excel(xls, sheet_name=sh_target)
    t.columns = t.columns.astype(str).str.strip()

    cm = detect_target_cols(t)
    if not cm["branch"] or not cm["achievement"]:
        raise ValueError("ستون‌های کلیدی تارگت پیدا نشد")

    # پاکسازی
    t[cm["branch"]] = normalize_name(t[cm["branch"]])
    if cm["supervisor"]:
        t[cm["supervisor"]] = normalize_name(t[cm["supervisor"]])
    if cm.get("zone_supervisor"):
        t[cm["zone_supervisor"]] = normalize_name(t[cm["zone_supervisor"]])
    t[cm["achievement"]] = parse_percent(t[cm["achievement"]])
    if cm["raked_value"]:
        t[cm["raked_value"]] = to_number(t[cm["raked_value"]])
    if cm["target"]:
        t[cm["target"]] = parse_percent(t[cm["target"]])
    if cm["change"]:
        t[cm["change"]] = parse_percent(t[cm["change"]])

    # حذف جمع‌ها، ردیف‌های خالی و NaN
    name = t[cm["branch"]].astype(str).str.strip()
    is_total = name.str.contains(r"^\s*(?:جمع|مجموع|کل|total|grand|sum)", case=False, regex=True)
    is_empty = t[cm["branch"]].isna() | name.isin(["", "nan", "None", "NaN"])
    ach_num = to_number(t[cm["achievement"]])
    if cm["raked_value"]:
        val_num = to_number(t[cm["raked_value"]])
        is_zero = (ach_num == 0) & (val_num == 0)
    else:
        is_zero = (ach_num == 0)
    t = t[~(is_total | is_empty | is_zero)].copy()
    t = t.reset_index(drop=True)

    if t.empty:
        raise ValueError("بعد از فیلتر کردن، هیچ شعبه معتبری باقی نماند")

    # ============ KPI مدیریتی ============
    vals = t[cm["achievement"]].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success_count = int((vals >= ACH_OK).sum())
    warning_count = int(((vals >= ACH_WARN) & (vals < ACH_OK)).sum())
    danger_count = int((vals < ACH_WARN).sum())
    total_branches = len(t)

    # ============ بهترین و بدترین ۱۰ شعبه ============
    rank = t.sort_values(cm["achievement"], ascending=False).reset_index(drop=True)
    rank.insert(0, "رتبه", range(1, len(rank) + 1))

    display_cols = [c for c in [cm["branch"], cm["supervisor"],
                                 cm.get("zone_supervisor"), cm["raked_value"],
                                 cm["target"], cm["achievement"], cm["change"]]
                    if c]

    top10 = rank.head(10)[display_cols].copy()
    worst10 = rank.tail(10).iloc[::-1][display_cols].copy()
    worst10.insert(0, "رتبه", range(len(rank), len(rank) - len(worst10), -1))
    worst10 = worst10.reset_index(drop=True)

    # ============ رتبه‌بندی سرپرست فروشگاه ============
    sup_rank = None
    if cm["supervisor"]:
        agg = {cm["achievement"]: "mean"}
        if cm["raked_value"]:
            agg[cm["raked_value"]] = "sum"
        agg[cm["branch"]] = lambda s: "، ".join(
            sorted(set(str(x).strip() for x in s.dropna() if str(x).strip()))
        )
        sup_rank = (t.groupby(cm["supervisor"], as_index=False).agg(agg)
                    .sort_values(cm["achievement"], ascending=False)
                    .reset_index(drop=True))
        sup_rank = sup_rank.rename(columns={cm["branch"]: "شعبه"})
        sup_rank.insert(0, "رتبه", range(1, len(sup_rank) + 1))

    # ============ رتبه‌بندی سوپروایزر منطقه ============
    zone_rank = None
    if cm.get("zone_supervisor"):
        agg = {cm["achievement"]: "mean"}
        if cm["raked_value"]:
            agg[cm["raked_value"]] = "sum"
        agg[cm["branch"]] = "count"
        zone_rank = (t.groupby(cm["zone_supervisor"], as_index=False).agg(agg)
                     .sort_values(cm["achievement"], ascending=False)
                     .reset_index(drop=True))
        zone_rank = zone_rank.rename(columns={cm["branch"]: "تعداد شعبه"})
        zone_rank.insert(0, "رتبه", range(1, len(zone_rank) + 1))

    # ============ روند تغییرات (شیت 60 روزه) ============
    trend_info = None
    sh_summary = find_sheet(xls, "60 روزه")
    if sh_summary:
        s = pd.read_excel(xls, sheet_name=sh_summary)
        s.columns = s.columns.astype(str).str.strip()
        _pct_cols = [c for c in s.columns if "درصد راکد" in c]
        pct_old_col = next((c for c in _pct_cols if "قبل" in c), None)
        pct_new_col = next((c for c in _pct_cols if "فعلی" in c), None)
        if not (pct_old_col and pct_new_col) and len(_pct_cols) >= 2:
            pct_old_col, pct_new_col = _pct_cols[0], _pct_cols[-1]
        if pct_old_col and pct_new_col:
            old = parse_percent(s[pct_old_col]).mean()
            new = parse_percent(s[pct_new_col]).mean()
            trend_info = {
                "old": old, "new": new, "diff": new - old,
                "improved": int((parse_percent(s[pct_new_col]) < parse_percent(s[pct_old_col])).sum()),
                "worsened": int((parse_percent(s[pct_new_col]) > parse_percent(s[pct_old_col])).sum()),
            }

    # ============ ساخت HTML جداول ============
    def _fmt_percent(v):
        if pd.isna(v):
            return ""
        try:
            x = float(v)
            if abs(x) < 0.05:
                return "0.0%"
            return f"{x:.1f}%"
        except (ValueError, TypeError):
            return str(v)

    def make_table(df, cols_map, money_cols=None):
        if df is None or df.empty:
            return "<p style='color:#888;'>داده‌ای موجود نیست</p>"
        d = df.copy()
        if money_cols:
            for orig in money_cols:
                if orig in d.columns:
                    d[orig] = d[orig].apply(lambda v: money(v))
        d = d.rename(columns=cols_map)
        for c in d.columns:
            if any(k in c for k in ["تحقق", "تارگت", "تغییرات", "درصد"]):
                d[c] = d[c].apply(_fmt_percent)
        return d.to_html(index=False, classes="table", border=0, escape=False)

    cols_map = {
        cm["branch"]: "شعبه",
        cm["supervisor"]: "سرپرست فروشگاه",
        cm.get("zone_supervisor"): "سوپروایزر منطقه",
        cm["raked_value"]: "ارزش راکد",
        cm["target"]: "تارگت",
        cm["achievement"]: "تحقق",
        cm["change"]: "تغییرات",
    }
    cols_map = {k: v for k, v in cols_map.items() if k}

    money_cols = [cm["raked_value"]] if cm["raked_value"] else []
    top10_html = make_table(top10, cols_map, money_cols)
    worst10_html = make_table(worst10, cols_map, money_cols)

    sup_html = None
    if sup_rank is not None and not sup_rank.empty:
        sup_map = {
            cm["supervisor"]: "سرپرست فروشگاه",
            "شعبه": "شعبه",
            cm["achievement"]: "میانگین تحقق",
            cm["raked_value"]: "مجموع راکد",
        }
        sup_map = {k: v for k, v in sup_map.items() if k}
        sup_html = make_table(sup_rank, sup_map, money_cols)

    zone_html = None
    if zone_rank is not None and not zone_rank.empty:
        zone_map = {
            cm["zone_supervisor"]: "سوپروایزر منطقه",
            cm["achievement"]: "میانگین تحقق",
            cm["raked_value"]: "مجموع راکد",
            "تعداد شعبه": "تعداد شعبه",
        }
        zone_map = {k: v for k, v in zone_map.items() if k}
        zone_html = make_table(zone_rank, zone_map, money_cols)

    # ============ قالب ایمیل (ریسپانسیو موبایل) ============
    template = Template("""
<!DOCTYPE html>
<html dir="rtl" lang="fa">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  body { font-family: Tahoma, Arial, sans-serif; background: #f4f4f4; color: #333; margin: 0; padding: 0; -webkit-text-size-adjust: 100%; }
  .wrapper { padding: 10px; }
  .container { max-width: 900px; margin: 0 auto; background: #fff; border-radius: 12px; overflow: hidden; }
  .header { background: #E6003E; color: white; padding: 22px 16px; text-align: center; }
  .header h1 { margin: 0; font-size: 20px; }
  .header p { margin: 6px 0 0 0; font-size: 12px; opacity: 0.9; }
  .content { padding: 16px; }

  .kpi-table { width: 100%; border-collapse: separate; border-spacing: 8px 0; margin-bottom: 20px; }
  .kpi-cell { background: #fafafa; border-right: 4px solid #E6003E; padding: 12px 8px; border-radius: 8px; text-align: center; vertical-align: middle; }
  .kpi-cell.green { border-right-color: #10b981; }
  .kpi-cell.yellow { border-right-color: #f59e0b; }
  .kpi-label { font-size: 11px; color: #888; margin-bottom: 4px; display: block; }
  .kpi-value { font-size: 16px; font-weight: bold; color: #E6003E; display: block; }
  .kpi-cell.green .kpi-value { color: #10b981; }
  .kpi-cell.yellow .kpi-value { color: #f59e0b; }

  h2 { font-size: 15px; color: #333; border-bottom: 2px solid #E6003E; padding-bottom: 6px; margin-top: 26px; }

  .table-wrap { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; margin-top: 10px; }
  .table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .table th { background: #E6003E; color: white; padding: 8px 6px; text-align: right; font-weight: bold; }
  .table td { border-bottom: 1px solid #eee; padding: 8px 6px; text-align: right; }
  .table tr:nth-child(even) td { background: #fafafa; }

  .footer { text-align: center; padding: 14px; font-size: 11px; color: #999; background: #fafafa; }
  .trend-up { color: #E6003E; font-weight: bold; }
  .trend-down { color: #10b981; font-weight: bold; }
  .summary-box { background: #fff5f7; border-right: 4px solid #E6003E; padding: 12px; border-radius: 8px; margin-top: 10px; font-size: 13px; line-height: 1.7; }

  @media only screen and (max-width: 600px) {
    .content { padding: 12px !important; }
    .header h1 { font-size: 17px !important; }
    .header { padding: 16px 12px !important; }
    .kpi-table { border-spacing: 4px 0 !important; }
    .kpi-cell { padding: 8px 4px !important; }
    .kpi-value { font-size: 13px !important; }
    .kpi-label { font-size: 10px !important; }
    .table { font-size: 10px !important; }
    .table th, .table td { padding: 5px 3px !important; }
    h2 { font-size: 13px !important; margin-top: 18px !important; }
    .summary-box { font-size: 12px !important; }
  }
</style>
</head>
<body>
<div class="wrapper">
  <div class="container">
    <div class="header">
      <h1>گزارش مدیریتی هفتگی</h1>
      <p>فروشگاه‌های زنجیره‌ای افق کوروش — {{ date }}</p>
    </div>

    <div class="content">
      <table class="kpi-table" role="presentation" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td class="kpi-cell" width="25%">
            <span class="kpi-label">میانگین تحقق</span>
            <span class="kpi-value">{{ "%.1f"|format(kpi.mean) }}%</span>
          </td>
          <td class="kpi-cell green" width="25%">
            <span class="kpi-label">🟢 موفق</span>
            <span class="kpi-value">{{ kpi.success }} شعبه</span>
          </td>
          <td class="kpi-cell yellow" width="25%">
            <span class="kpi-label">🟡 در حال پیشرفت</span>
            <span class="kpi-value">{{ kpi.warning }} شعبه</span>
          </td>
          <td class="kpi-cell" width="25%">
            <span class="kpi-label">🔴 بحرانی</span>
            <span class="kpi-value">{{ kpi.danger }} شعبه</span>
          </td>
        </tr>
      </table>

      {% if trend %}
      <h2>📈 روند هفته</h2>
      <div class="summary-box">
        میانگین درصد راکد از <b>{{ "%.2f"|format(trend.old) }}%</b>
        به <b>{{ "%.2f"|format(trend.new) }}%</b>
        {% if trend.diff < 0 %}
          <span class="trend-down">(کاهش {{ "%.2f"|format(-trend.diff) }}% ✅)</span>
        {% elif trend.diff > 0 %}
          <span class="trend-up">(افزایش {{ "%.2f"|format(trend.diff) }}% ⚠️)</span>
        {% else %}
          <span>(بدون تغییر)</span>
        {% endif %}
        <br>
        🟢 <b>{{ trend.improved }}</b> شعبه بهبود &nbsp;|&nbsp;
        🔴 <b>{{ trend.worsened }}</b> شعبه افت
      </div>
      {% endif %}

      <h2>🥇 بهترین ۱۰ شعبه</h2>
      <div class="table-wrap">{{ top10 }}</div>

      <h2>⚠️ بدترین ۱۰ شعبه</h2>
      <div class="table-wrap">{{ worst10 }}</div>

      {% if zone_html %}
      <h2>🏆 رتبه‌بندی سوپروایزرهای منطقه</h2>
      <div class="table-wrap">{{ zone_html }}</div>
      {% endif %}

      {% if sup_html %}
      <h2>👤 رتبه‌بندی سرپرست فروشگاه‌ها</h2>
      <div class="table-wrap">{{ sup_html }}</div>
      {% endif %}
    </div>

    <div class="footer">
      این ایمیل به صورت خودکار توسط داشبورد کالای راکد افق کوروش ارسال شده است.
    </div>
  </div>
</div>
</body>
</html>
    """)

    html = template.render(
        date=datetime.now().strftime("%Y-%m-%d"),
        kpi={
            "mean": mean_ach,
            "success": success_count,
            "warning": warning_count,
            "danger": danger_count,
        },
        total_branches=total_branches,
        trend=trend_info,
        top10=top10_html,
        worst10=worst10_html,
        sup_html=sup_html,
        zone_html=zone_html,
    )

    return html, {
        "mean_ach": mean_ach,
        "success": success_count,
        "warning": warning_count,
        "danger": danger_count,
        "branches": total_branches,
    }


# ================== ارسال ایمیل ==================


def _load_smtp_config():
    """خوندن تنظیمات SMTP از secrets، env یا config.ini"""
    import os
    try:
        import streamlit as st
        try:
            return {
                "server": st.secrets["SMTP"]["server"],
                "port": int(st.secrets["SMTP"]["port"]),
                "sender_email": st.secrets["SMTP"]["sender_email"],
                "sender_password": st.secrets["SMTP"]["sender_password"],
            }
        except Exception:
            pass
    except Exception:
        pass

    # env
    if os.environ.get("SMTP_SERVER"):
        return {
            "server": os.environ.get("SMTP_SERVER"),
            "port": int(os.environ.get("SMTP_PORT", 465)),
            "sender_email": os.environ.get("SMTP_EMAIL"),
            "sender_password": os.environ.get("SMTP_PASSWORD"),
        }

    # config.ini
    import configparser
    cfg = configparser.ConfigParser()
    cfg.read("config.ini", encoding="utf-8")
    return {
        "server": cfg.get("SMTP", "server", fallback="smtp.gmail.com"),
        "port": cfg.getint("SMTP", "port", fallback=465),
        "sender_email": cfg.get("SMTP", "sender_email", fallback=""),
        "sender_password": cfg.get("SMTP", "sender_password", fallback=""),
    }


def send_email(html, subject=None, recipients=None):
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding="utf-8")

    server = config["SMTP"]["server"]
    port = int(config["SMTP"]["port"])
    sender = os.environ.get("SMTP_SENDER_EMAIL") or config["SMTP"]["sender_email"]
    password = os.environ.get("SMTP_SENDER_PASSWORD") or config["SMTP"]["sender_password"]
    if isinstance(recipients, str):
        recipients = [recipients]
    recipients = [r.strip() for r in (recipients or []) if r and r.strip()]
    if not recipients:
        recipients = [r.strip() for r in config["REPORT"]["recipients"].split(",") if r.strip()]
    subject = subject or config["REPORT"]["subject"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(html, "html", "utf-8"))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(server, port, context=context, timeout=30) as s:
        s.set_debuglevel(0)
        s.login(sender, password)
        s.sendmail(sender, recipients, msg.as_string())

    return recipients


# ================== اجرا ==================
if __name__ == "__main__":
    print("📊 در حال ساخت گزارش مدیریتی...")
    html, stats = build_report()
    print("✅ گزارش ساخته شد:")
    print(f"   - میانگین تحقق: {stats['mean_ach']:.1f}%")
    print(f"   - موفق: {stats['success']} شعبه")
    print(f"   - در حال پیشرفت: {stats['warning']} شعبه")
    print(f"   - بحرانی: {stats['danger']} شعبه")
    print(f"   - کل شعب: {stats['branches']}")

    print("📧 در حال ارسال ایمیل...")
    recipients = send_email(html)
    print(f"✅ ایمیل با موفقیت به {', '.join(recipients)} ارسال شد!")