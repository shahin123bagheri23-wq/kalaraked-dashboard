"""تبدیل DataFrame به عکس و ارسال به بله"""
import time
from pathlib import Path
from datetime import datetime

import pandas as pd

from bale_sender import _load_token, _url
import requests






def send_pdf(chat_id, file_path, caption="", token=None):
    """ارسال PDF به بله"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(file_path)
    with open(p, "rb") as f:
        files = {"document": (p.name, f, "application/pdf")}
        data = {"chat_id": chat_id, "caption": caption}
        r = requests.post(_url(token, "sendDocument"),
                          data=data, files=files, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()


def send_photo(chat_id, file_path, caption="", token=None):
    """ارسال عکس به بله"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(file_path)
    with open(p, "rb") as f:
        files = {"photo": (p.name, f, "image/png")}
        data = {"chat_id": chat_id, "caption": caption}
        r = _req("post", _url(token, "sendPhoto"), data=data, files=files, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()





def _find_chrome():
    """پیدا کردن Chrome/Chromium روی ویندوز و لینوکس"""
    from pathlib import Path
    import os

    local = os.environ.get("LOCALAPPDATA", "")
    pf = os.environ.get("PROGRAMFILES", "C:\\Program Files")
    pfx86 = os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)")

    candidates = [
        # ── لینوکس (Streamlit Cloud / Render) ──
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chrome",
        "/snap/bin/chromium",
        # ── ویندوز — Chrome ──
        os.path.join(local, r"Google\\Chrome\\Application\\chrome.exe"),
        os.path.join(pf, r"Google\\Chrome\\Application\\chrome.exe"),
        os.path.join(pfx86, r"Google\\Chrome\\Application\\chrome.exe"),
        # ── ویندوز — Edge ──
        os.path.join(pfx86, r"Microsoft\\Edge\\Application\\msedge.exe"),
        os.path.join(pf, r"Microsoft\\Edge\\Application\\msedge.exe"),
        os.path.join(local, r"Microsoft\\Edge\\Application\\msedge.exe"),
    ]
    for c in candidates:
        if c and Path(c).exists():
            return c
    return None

def _fmt_number(v):
    """فرمت اعداد با جداکننده هزارگان"""
    try:
        if pd.isna(v):
            return "—"
        n = float(v)
        if n == int(n):
            return f"{int(n):,}"
        return f"{n:,.2f}"
    except (ValueError, TypeError):
        return str(v)


def _clean_df_for_image(df):
    """پاکسازی دیتافریم برای نمایش در عکس"""
    df = df.copy()

    # ── حذف ستون‌های غیرضروری ──
    drop_cols = []
    for c in df.columns:
        cl = str(c)
        # کد شعبه
        if "کد" in cl and "شعبه" in cl:
            drop_cols.append(c)
            continue
        # نوع پروژه
        if "نوع پروژه" in cl:
            drop_cols.append(c)
            continue
        # سوپروایزر / سرپرست
        if "سوپروایزر" in cl or "سرپرست" in cl:
            drop_cols.append(c)
            continue
        # ستون‌های ریالی
        if "ریال" in cl:
            drop_cols.append(c)
            continue
        # Index و ردیف
        if cl.strip() in ("Index", "Index.1", "ردیف", "Unnamed: 0"):
            drop_cols.append(c)
            continue
        # موجودی ریالی
        if "موجودی ریالی" in cl:
            drop_cols.append(c)
            continue

    for c in drop_cols:
        df = df.drop(columns=[c])

    # فرمت اعداد (موجودی سیستمی، فروش تعدادی)
    for c in df.columns:
        cl = str(c)
        if "موجودی" in cl or "فروش" in cl:
            try:
                df[c] = df[c].apply(_fmt_number)
            except Exception:
                pass

    return df




def df_to_pdf(df, output_path, title="", font_size=12):
    """DataFrame → PDF وکتور با Plotly"""
    import plotly.graph_objects as go

    if len(df) == 0:
        raise ValueError("داده‌ای برای نمایش نیست")

    df = _clean_df_for_image(df)
    n_rows = len(df)

    char_width = font_size * 0.75
    col_widths = []
    for c in df.columns:
        max_len = len(str(c))
        for v in df[c].astype(str):
            max_len = max(max_len, len(str(v)))
        col_widths.append(max(100, int(max_len * char_width) + 50))

    total_width_px = max(1400, min(4000, sum(col_widths)))
    if sum(col_widths) > 4000:
        ratio = 4000 / sum(col_widths)
        col_widths = [max(90, int(w * ratio)) for w in col_widths]
        total_width_px = sum(col_widths)

    row_h = 38
    height = max(400, n_rows * row_h + 160)

    fig = go.Figure(data=[go.Table(
        columnwidth=col_widths,
        header=dict(
            values=[f"<b>{c}</b>" for c in df.columns],
            fill_color="#E6003E",
            align="center",
            font=dict(color="white", size=font_size + 2, family="Tahoma"),
            height=50,
        ),
        cells=dict(
            values=[df[c].astype(str).tolist() for c in df.columns],
            fill_color=[["#f9f9f9" if i % 2 == 0 else "#ffffff" for i in range(n_rows)]],
            align="center",
            font=dict(color="#111111", size=font_size, family="Tahoma"),
            height=row_h - 4,
        ),
    )])

    fig.update_layout(
        title=dict(text=title, font=dict(size=font_size + 5, color="#E6003E", family="Tahoma"),
                   x=0.5, xanchor="center"),
        margin=dict(l=15, r=15, t=70 if title else 20, b=20),
        height=height,
        width=total_width_px,
        paper_bgcolor="white",
    )

    fig.write_image(output_path, format="pdf", width=total_width_px, height=height)
    return output_path

# ═══════════════════════════════════════════════════════════════════════
#  قالب‌های اختصاصی گزارش‌های «فروش اینترنتی و مغایرت‌گیری»
#  هر گزارش رنگ، چیدمان و نوع نمایش مخصوص خودش را دارد:
#    online_perf     → آبی    | کارت‌های جمع کل + نوار برای درصدها
#    online_compare  → بنفش   | هدر دو طبقه (قبل/جدید/روند) + رنگ بهتر/بدتر
#    adjust_count    → سبز    | نوار پیشرفت «درصد انجام»
#    adjust_summary  → کهربایی | رتبه + نوار مبلغ منفی + برچسب خالص
#    adjust_top      → نارنجی | لیست رتبه‌بندی‌شده کالا/شعبه با نوار مبلغ
#  گزارش‌های دیگر (کالای راکد و ...) همان ظاهر قبلی را دارند.
# ═══════════════════════════════════════════════════════════════════════
import html as _html
import re as _re

_THEMES = {
    "online_perf":    {"kind": "perf",    "accent": "#38bdf8"},
    "online_compare": {"kind": "compare", "accent": "#a78bfa"},
    "adjust_count":   {"kind": "count",   "accent": "#34d399"},
    "adjust_summary": {"kind": "summary", "accent": "#fbbf24"},
    "adjust_top":     {"kind": "top",     "accent": "#fb923c"},
}
_CURRENT_THEME = None

_C_BAD, _C_GOOD, _C_WARN = "#fb7185", "#34d399", "#fbbf24"


def set_theme(source_key):
    """report_sender قبل از ساخت عکس صدا می‌زند؛ منبع غیرِ فروش اینترنتی → ظاهر قبلی"""
    global _CURRENT_THEME
    _CURRENT_THEME = source_key if source_key in _THEMES else None


def _rgba(hex_color, a):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{a})"


def _tf(v):
    """عدد داخل یک متن (مثل «12.34%» یا «-1,234.5») — در نبود عدد None"""
    m = _re.search(r"-?\d+(?:\.\d+)?", str(v).replace(",", "").replace("٬", ""))
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def _col(df, *keys):
    for c in df.columns:
        if all(k in str(c) for k in keys):
            return c
    return None


def _sum(df, col):
    if col is None:
        return None
    vals = [x for x in (_tf(v) for v in df[col]) if x is not None]
    return sum(vals) if vals else None


def _fi(x, d=0):
    return "—" if x is None else f"{x:,.{d}f}"


_NUMLIKE = _re.compile(r"^[\-+]?[\d,]+(\.\d+)?%?$")


def _plain(v):
    """اعداد را LTR ایزوله می‌کند تا «-123» در متن راست‌به‌چپ وارونه نشود"""
    s = str(v).strip()
    e = _html.escape(s)
    if s == "—" or _NUMLIKE.match(s):
        return f'<span class="num">{e}</span>'
    return e


def _bar(text, frac, color):
    w = max(0.0, min(1.0, frac)) * 100
    return (f'<td><div class="b"><i style="width:{w:.1f}%;background:{color}"></i>'
            f'<span class="num">{_html.escape(str(text))}</span></div></td>')


def _kpis_html(items):
    if not items:
        return ""
    out = "".join(
        f'<div class="k {tone}"><div class="v"><span class="num">{_html.escape(v)}</span></div>'
        f'<div class="l">{_html.escape(l)}</div></div>'
        for l, v, tone in items)
    return f'<div class="kpis">{out}</div>'


def _thead_simple(cols, rank=False):
    ths = []
    if rank:
        ths.append("<th>#</th>")
    for i, c in enumerate(cols):
        cls = ' class="name-h"' if (i == 0 and not rank) else ""
        ths.append(f"<th{cls}>{_html.escape(str(c))}</th>")
    return "<tr>" + "".join(ths) + "</tr>"


# ───────────────────────── سازنده‌ها ─────────────────────────
# هر سازنده: (kpi_html, thead_html, tbody_html, عرض_اضافه، تعداد_سطر_هدر)

def _b_generic(df, t, compact):
    cols = list(df.columns)
    rows = []
    for row in df.values:
        tds = [f'<td class="name">{_html.escape(str(v))}</td>' if i == 0 else f"<td>{_plain(v)}</td>"
               for i, v in enumerate(row)]
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return "", _thead_simple(cols), "".join(rows), [], [], 0, 1


def _b_perf(df, t, compact):
    cols = list(df.columns)
    name_c = _col(df, "نام شعبه") or cols[0]
    pct_cols = [c for c in cols if str(c).startswith("درصد")]
    mx = {c: max([_tf(v) or 0 for v in df[c]] + [0]) for c in pct_cols}

    kp = []
    for lab, tone in (("ثبتی", ""), ("تحویلی", "acc"), ("تاخیر", "hot"), ("مرجوعی", "")):
        c = next((c for c in cols if str(c).strip() == lab), None)
        s = _sum(df, c)
        if s is not None:
            kp.append((f"مجموع {lab}", _fi(s), tone))

    rows = []
    for row in df.values:
        d = dict(zip(cols, row))
        tds = []
        for c in cols:
            v = d[c]
            if c == name_c:
                tds.append(f'<td class="name">{_html.escape(str(v))}</td>')
            elif c in pct_cols:
                x, m = _tf(v), mx[c]
                frac = (x / m) if (x is not None and m > 0) else 0
                tds.append(_bar(v, frac, _rgba(_C_BAD, .55) if frac >= .66 else _rgba(t["accent"], .35)))
            else:
                tds.append(f"<td>{_plain(v)}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return _kpis_html(kp), _thead_simple(cols), "".join(rows), pct_cols, [], 0, 1


def _b_compare(df, t, compact):
    cols = list(df.columns)
    name_c = _col(df, "نام شعبه") or cols[0]
    rest = [c for c in cols if c != name_c]
    info, prev_of = {}, {}
    for c in rest:
        s = str(c)
        sub = "قبل" if "قبل" in s else ("جدید" if "جدید" in s else ("روند" if "روند" in s else None))
        if sub is None:
            return _b_generic(df, t, compact)
        g = s.replace("قبل", "").replace("جدید", "").replace("روند", "").strip()
        info[c] = (g, sub)
        if sub == "قبل":
            prev_of[g] = c
    groups = []
    for c in rest:
        g = info[c][0]
        if groups and groups[-1][0] == g:
            groups[-1][1].append(c)
        else:
            groups.append((g, [c]))
    first_of_group = {gc[0] for _, gc in groups}

    h1 = (f'<tr><th rowspan="2" class="name-h">{_html.escape(str(name_c))}</th>' +
          "".join(f'<th colspan="{len(gc)}" class="grp gs">{_html.escape(g)}</th>' for g, gc in groups) +
          "</tr>")
    h2 = "<tr>" + "".join(
        f'<th class="sub{" gs" if c in first_of_group else ""}">{_html.escape(info[c][1])}</th>'
        for c in rest) + "</tr>"

    worse = better = 0
    late_new = next((c for c in rest if info[c][1] == "جدید" and "تاخیر" in info[c][0]), None)
    late_prev = prev_of.get(info[late_new][0]) if late_new else None

    rows = []
    for row in df.values:
        d = dict(zip(cols, row))
        tds = [f'<td class="name">{_html.escape(str(d[name_c]))}</td>']
        for c in rest:
            g, sub = info[c]
            v = d[c]
            gs = " gs" if c in first_of_group else ""
            if sub == "قبل":
                tds.append(f'<td class="dim{gs}">{_plain(v)}</td>')
            elif sub == "جدید":
                x = _tf(v)
                p = _tf(d[prev_of[g]]) if g in prev_of else None
                state = ""
                if x is not None and p is not None:
                    state = "worse" if x > p else ("better" if x < p else "")
                if c == late_new and p is not None and x is not None:
                    worse += state == "worse"
                    better += state == "better"
                tds.append(f'<td class="newv {state}{gs}">{_plain(v)}</td>')
            else:  # روند
                s = str(v)
                if "▲" in s:
                    cls, txt = "bad", ("▲" if compact else s)
                elif "▼" in s:
                    cls, txt = "good", ("▼" if compact else s)
                else:
                    cls, txt = "flat", "—"
                tds.append(f'<td class="{gs.strip()}"><span class="pill {cls}">{_html.escape(txt)}</span></td>')
        rows.append("<tr>" + "".join(tds) + "</tr>")

    kp = []
    if late_new is not None:
        kp = [("شعبه با تاخیر بدتر", _fi(worse), "hot" if worse else ""),
              ("شعبه با تاخیر بهتر", _fi(better), "good")]
    return _kpis_html(kp), h1 + h2, "".join(rows), [], [], 0, 2


def _b_count(df, t, compact):
    cols = list(df.columns)
    name_c = _col(df, "نام شعبه") or cols[0]
    pc = _col(df, "درصد انجام")
    if pc is None:
        return _b_generic(df, t, compact)
    ps = [_tf(v) for v in df[pc]]
    low = sum(1 for p in ps if p is not None and p < 70)
    full = sum(1 for p in ps if p is not None and p >= 100)
    kp = [("تعداد شعب", _fi(len(df)), ""),
          ("کمتر از ۷۰ درصد", _fi(low), "hot" if low else ""),
          ("به ۱۰۰ درصد رسیده", _fi(full), "good")]
    rows = []
    for row in df.values:
        d = dict(zip(cols, row))
        tds = []
        for c in cols:
            v = d[c]
            if c == name_c:
                tds.append(f'<td class="name">{_html.escape(str(v))}</td>')
            elif c == pc:
                p = _tf(v)
                color = _C_GOOD if (p or 0) >= 100 else (_C_WARN if (p or 0) >= 70 else _C_BAD)
                tds.append(_bar(v, (p or 0) / 100.0, _rgba(color, .5)))
            else:
                tds.append(f"<td>{_plain(v)}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return _kpis_html(kp), _thead_simple(cols), "".join(rows), [pc], [], 0, 1


def _b_summary(df, t, compact):
    cols = list(df.columns)
    name_c = _col(df, "نام شعبه") or cols[0]
    neg_c, pos_c, net_c = _col(df, "مبلغ", "منفی"), _col(df, "مبلغ", "مثبت"), _col(df, "خالص")
    if neg_c is None or net_c is None:
        return _b_generic(df, t, compact)
    mxn = max([abs(_tf(v) or 0) for v in df[neg_c]] + [0])
    net_sum = _sum(df, net_c)
    kp = [("جمع مثبت (میلیون)", _fi(_sum(df, pos_c), 1), "good"),
          ("جمع منفی (میلیون)", _fi(_sum(df, neg_c), 1), "hot"),
          ("خالص (میلیون)", _fi(net_sum, 1), "good" if (net_sum or 0) >= 0 else "hot")]
    rows = []
    for i, row in enumerate(df.values, 1):
        d = dict(zip(cols, row))
        tds = [f'<td class="rk">{i}</td>']
        for c in cols:
            v = d[c]
            if c == name_c:
                tds.append(f'<td class="name">{_html.escape(str(v))}</td>')
            elif c == neg_c:
                x = abs(_tf(v) or 0)
                tds.append(_bar(v, (x / mxn) if mxn > 0 else 0, _rgba(_C_BAD, .5)))
            elif c == net_c:
                x = _tf(v)
                cls = "flat" if x is None else ("bad" if x < 0 else "good")
                tds.append(f'<td><span class="pill {cls}">{_plain(v)}</span></td>')
            elif c == pos_c:
                tds.append(f'<td style="color:{_C_GOOD}">{_plain(v)}</td>')
            else:
                tds.append(f"<td>{_plain(v)}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return _kpis_html(kp), _thead_simple(cols, rank=True), "".join(rows), [neg_c], [], 1, 1


def _b_top(df, t, compact):
    cols = list(df.columns)
    br_c, item_c, amt_c = _col(df, "نام شعبه"), _col(df, "نام کالا"), _col(df, "مبلغ")
    if None in (br_c, item_c, amt_c):
        return _b_generic(df, t, compact)
    date_c, qty_c = _col(df, "تاریخ"), _col(df, "تعداد")
    vals = [abs(_tf(v) or 0) for v in df[amt_c]]
    mx = max(vals + [0])
    kp = [("مجموع مبلغ منفی (میلیون)", _fi(sum(vals), 1), "hot"),
          ("بزرگ‌ترین قلم (میلیون)", _fi(mx, 1), "acc"),
          ("تعداد قلم", _fi(len(df)), "")]
    head = ["#", "کالا / شعبه"] + [str(c) for c in (date_c, qty_c, amt_c) if c is not None]
    thead = "<tr>" + "".join(f"<th>{_html.escape(h)}</th>" for h in head) + "</tr>"
    rows = []
    for i, row in enumerate(df.values, 1):
        d = dict(zip(cols, row))
        tds = [f'<td class="rk">{i}</td>',
               f'<td class="item"><div class="it">{_html.escape(str(d[item_c]))}</div>'
               f'<div class="br">{_html.escape(str(d[br_c]))}</div></td>']
        if date_c is not None:
            tds.append(f'<td class="dim">{_plain(d[date_c])}</td>')
        if qty_c is not None:
            tds.append(f"<td>{_plain(d[qty_c])}</td>")
        x = abs(_tf(d[amt_c]) or 0)
        tds.append(_bar(d[amt_c], (x / mx) if mx > 0 else 0, _rgba(t["accent"], .5)))
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return _kpis_html(kp), thead, "".join(rows), [amt_c], [br_c], 1, 1


_BUILDERS = {"perf": _b_perf, "compare": _b_compare, "count": _b_count,
             "summary": _b_summary, "top": _b_top}

_THEMED_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
html,body{font-family:Vazirmatn,Vazir,'Segoe UI',Tahoma,Arial,sans-serif;-webkit-font-smoothing:antialiased;
  background:#0e1015;direction:rtl;width:%%W%%px;overflow:hidden}
.wrap{padding:%%PAD%%px}
.card{background:#161a21;border:1px solid #262c37;border-radius:12px;overflow:hidden}
.head{display:flex;align-items:stretch;gap:12px;padding:%%HEADPAD%%}
.stripe{width:5px;border-radius:3px;background:%%ACC%%;flex:none}
.t{font-size:%%TFS%%px;font-weight:800;color:#fff;line-height:1.5}
.s{font-size:%%SFS%%px;color:#8d96a6;margin-top:2px}
.kpis{display:flex;gap:8px;padding:0 %%PAD%%px %%PAD%%px}
.k{flex:1;background:#1c212a;border:1px solid #2a313d;border-radius:8px;padding:%%KPAD%%;text-align:center}
.k .v{font-size:%%KFS%%px;font-weight:800;color:#fff;line-height:1.3}
.k .l{font-size:%%SFS%%px;color:#8d96a6}
.k.hot .v{color:#fb7185}.k.good .v{color:#34d399}.k.acc .v{color:%%ACC%%}
table{border-collapse:collapse;width:100%;font-size:%%FS%%px}
th{background:#1c212a;color:%%ACC%%;font-weight:800;padding:%%THPAD%%;text-align:center;white-space:normal;line-height:1.35;
  border-bottom:2px solid %%ACC%%}
th.grp{color:#fff;background:%%ACCBG%%;border-bottom:1px solid #2f3745}
th.sub{font-size:%%SFS%%px}
th.name-h{text-align:right;padding-right:%%NPAD%%px}
td{padding:%%TDPAD%%;text-align:center;color:#e6e9ef;white-space:nowrap;font-weight:600;
  border-bottom:1px solid #222833;vertical-align:middle}
td.name{text-align:right;font-weight:800;color:#fff;padding-right:%%NPAD%%px}
td.dim{color:#8d96a6;font-weight:500}
td.newv{font-weight:800}td.newv.worse{color:#fb7185}td.newv.better{color:#34d399}
td.rk{color:%%ACC%%;font-weight:800;width:%%RKW%%px}
td.item{white-space:normal;text-align:right;max-width:%%IW%%px;line-height:1.45}
td.item .it{font-weight:800;color:#fff}
td.item .br{color:#8d96a6;font-size:%%SFS%%px;font-weight:500}
.gs{border-right:1px solid #2f3745}
.num{direction:ltr;unicode-bidi:isolate;display:inline-block}
.b{position:relative;width:%%BW%%px;height:%%BH%%px;margin:0 auto;background:#222833;border-radius:5px;overflow:hidden}
.b i{position:absolute;top:0;bottom:0;right:0;border-radius:5px}
.b span{position:relative;z-index:1;display:block;line-height:%%BH%%px;font-weight:800;color:#fff;
  font-size:%%FS%%px;text-align:center}
.pill{display:inline-block;padding:1px 9px;border-radius:99px;font-weight:800;font-size:%%SFS%%px}
.pill.bad{background:rgba(251,113,133,.16);color:#fb7185}
.pill.good{background:rgba(52,211,153,.16);color:#34d399}
.pill.flat{background:#222833;color:#8d96a6}
.foot{text-align:center;color:#6f7889;font-size:%%SFS%%px;padding:%%FOOTPAD%%px;border-top:1px solid #262c37}
"""


def _sizes(compact):
    if compact:   # اندازه‌ی نهایی، scale=1 (بله فشرده‌اش نمی‌کند)
        return dict(fs=13, sfs=11, tfs=15, kfs=17, pad=8, headpad="10px 12px", kpad="6px 8px",
                    thpad="6px 7px", tdpad="4px 7px", npad=10, rkw=26, iw=250, bw=108, bh=18,
                    footpad=6, scale=1, head_h=52, kpi_h=58, th_h=46, row_h=29, row_h_top=42,
                    foot_h=26, cellx=16,
                    minw=dict(perf=720, compare=780, count=620, summary=740, top=700))
    return dict(fs=17, sfs=14, tfs=21, kfs=26, pad=14, headpad="16px 18px", kpad="10px 12px",
                thpad="10px 10px", tdpad="8px 10px", npad=16, rkw=40, iw=480, bw=190, bh=30,
                footpad=10, scale=3, head_h=84, kpi_h=92, th_h=70, row_h=52, row_h_top=76,
                foot_h=44, cellx=26,
                minw=dict(perf=1100, compare=1200, count=950, summary=1150, top=1100))


def _themed_html(df, theme, title, subtitle, compact):
    t = _THEMES[theme]
    kind = t["kind"]
    S = _sizes(compact)
    kp, thead, tbody, bar_cols, skip_cols, n_rank, n_head = _BUILDERS[kind](df, t, compact)

    # عرض عکس: از محتوا تخمین می‌زنیم (هدر ستون‌ها شکسته می‌شود، پس طولانی‌ترین کلمه‌ی هدر)
    char_w = S["fs"] * 0.68
    est = n_rank * (S["rkw"] + S["cellx"])
    for c in df.columns:
        if c in skip_cols:
            continue
        if c in bar_cols:
            est += S["bw"] + S["cellx"] + 4
            continue
        if kind == "compare":
            word = 5
        else:
            word = max([len(w) for w in str(c).split()] + [1])
        ml = max([word] + [len(str(v)) for v in df[c]])
        cap = S["iw"] if "نام کالا" in str(c) else 260
        est += min(ml * char_w + S["cellx"], cap)
    width = int(max(S["minw"][kind], min(est * 1.04, 1500 if compact else 3000)))

    css = _THEMED_CSS
    for k, v in {"W": width, "PAD": S["pad"], "HEADPAD": S["headpad"], "ACC": t["accent"],
                 "ACCBG": _rgba(t["accent"], .16), "TFS": S["tfs"], "SFS": S["sfs"], "KPAD": S["kpad"],
                 "KFS": S["kfs"], "FS": S["fs"], "THPAD": S["thpad"], "TDPAD": S["tdpad"],
                 "NPAD": S["npad"], "RKW": S["rkw"], "IW": S["iw"], "BW": S["bw"], "BH": S["bh"],
                 "FOOTPAD": S["footpad"]}.items():
        css = css.replace(f"%%{k}%%", str(v))

    sub_html = f'<div class="s">{_html.escape(subtitle)}</div>' if subtitle else ""
    n = len(df)
    html = (f'<!DOCTYPE html><html dir="rtl"><head><meta charset="UTF-8"><style>{css}</style></head>'
            f'<body><div class="wrap"><div class="card">'
            f'<div class="head"><div class="stripe"></div><div><div class="t">{_html.escape(title or "")}</div>'
            f'{sub_html}</div></div>'
            f'{kp}<table><thead>{thead}</thead><tbody>{tbody}</tbody></table>'
            f'<div class="foot">افق کوروش • {n} ردیف</div></div></div></body></html>')

    rh = S["row_h_top"] if kind == "top" else S["row_h"]
    height = (S["head_h"] + (S["kpi_h"] if kp else 0) + S["th_h"] * n_head +
              n * rh + S["foot_h"] + S["pad"] * 2)
    height = int(height * 1.15 + 40)       # با سخاوت؛ بعد از عکس‌برداری بالا/پایین اضافه بریده می‌شود
    return html, width, height, S


def _autocrop(path, width_px, pad_px):
    """فضای خالی پایین (و حاشیه‌ی اضافه‌ی راست) را می‌برد؛ بدون Pillow کاری نمی‌کند"""
    try:
        from PIL import Image, ImageChops
        im = Image.open(path).convert("RGB")
        bg = im.getpixel((im.width - 1, im.height - 1))
        bbox = ImageChops.difference(im, Image.new("RGB", im.size, bg)).getbbox()
        if not bbox:
            return
        bottom = bbox[3]
        # اگر محتوا تا ته عکس رسیده، احتمالاً بریده شده؛ ارتفاع را دست نمی‌زنیم
        new_h = im.height if bottom >= im.height - 2 else min(im.height, bottom + pad_px)
        # صفحه راست‌به‌چپ است و به لبه‌ی راست پنجره می‌چسبد؛ حاشیه‌ی اضافه از سمت چپ بریده شود
        left = max(0, im.width - width_px)
        im.crop((left, 0, im.width, new_h)).save(path, optimize=True)
    except Exception:
        pass


def _shoot(html, output_path, width, height, scale, pad):
    import os
    import subprocess
    import time as _tm
    from pathlib import Path

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp_html = out.parent / f"_tmp_report_{os.getpid()}_{int(_tm.time() * 1000)}.html"
    tmp_html.write_text(html, encoding="utf-8")
    try:
        out.unlink(missing_ok=True)       # عکس قدیمی با عکس تازه اشتباه گرفته نشود
    except Exception:
        pass

    chrome = _find_chrome()
    if not chrome:
        try:
            tmp_html.unlink(missing_ok=True)
        except Exception:
            pass
        raise RuntimeError("Chrome پیدا نشد")

    cmd = [
        chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-sandbox",
        f"--force-device-scale-factor={scale}",
        f"--window-size={width + 20},{height + 20}",
        f"--screenshot={out.resolve()}",
        "--default-background-color=FF0E1015",
        "--virtual-time-budget=1500",
        tmp_html.resolve().as_uri(),
    ]
    try:
        subprocess.run(cmd, capture_output=True, timeout=90)
    finally:
        try:
            tmp_html.unlink(missing_ok=True)
        except Exception:
            pass

    if not out.exists():
        raise RuntimeError("عکس ساخته نشد")
    _autocrop(str(out), width * scale, pad * scale)
    return output_path


def _df_to_image_themed(df, output_path, title, compact, theme, subtitle=""):
    html, width, height, S = _themed_html(df, theme, title, subtitle, compact)
    return _shoot(html, output_path, width, height, S["scale"], S["pad"])



def df_to_image(df, output_path, title="", font_size=None, cell_padding=0.15,
                 per_branch_msg=None, compact=False, theme=None):
    """DataFrame → PNG با تم دارک افق کوروش + خط قرمز بین شعبه‌ها"""
    import html as _html
    import subprocess
    from pathlib import Path

    if len(df) == 0:
        raise ValueError("داده‌ای برای نمایش نیست")

    # پاکسازی + فرمت اعداد
    df = _clean_df_for_image(df)

    # گزارش‌های «فروش اینترنتی و مغایرت‌گیری» قالب اختصاصی خودشان را دارند
    _th = theme or _CURRENT_THEME
    if _th in _THEMES:
        return _df_to_image_themed(df, output_path, title, compact, _th,
                                   subtitle=per_branch_msg or "")

    n_rows = len(df)
    # کیفیت: فونت کمتر از ۱۶ باعث محو شدن حروف فارسی می‌شود
    if compact:
        # حالت فشرده: عکس از اول در اندازه‌ی نهایی ساخته می‌شود (مثل خروجی اکسل)
        # تا بله آن را کوچک و فشرده نکند
        font_size = 12
    else:
        font_size = max(font_size or 16, 16)

    # ── محاسبه عرض هر ستون ──
    char_w = font_size * 0.68
    col_widths = []
    for c in df.columns:
        max_len = len(str(c))
        for v in df[c].astype(str):
            max_len = max(max_len, len(str(v)))
        col_widths.append(max(80, int(max_len * char_w) + 30))

    total_width = max(680 if compact else 1000, min(5000, sum(col_widths)))

    if compact:
        header_h, row_h = 28, 20
        title_h = 40 if title else 0
        total_height = header_h + (n_rows * row_h) + title_h + 60
        wrap_pad, banner_pad, th_pad, td_pad, foot_pad = 8, "8px 10px", "4px 8px", "2px 8px", 5
        banner_fs, scale = font_size + 4, 1
    else:
        header_h, row_h = 56, 48
        title_h = 70 if title else 0
        total_height = header_h + (n_rows * row_h) + title_h + 110
        wrap_pad, banner_pad, th_pad, td_pad, foot_pad = 14, "16px 12px", "10px 10px", "8px 10px", 10
        banner_fs, scale = font_size + 7, 3

    # ── ساخت ردیف‌های HTML با خط قرمز بین شعبه‌ها ──
    _branch_col_name = next((c for c in df.columns if "نام شعبه" in c), df.columns[0])
    _status_col_name = next((c for c in df.columns if "وضعیت" in c), None)

    rows_html_parts = []
    _prev_branch = None

    for row in df.values:
        row_dict = dict(zip(df.columns, row))
        _branch = str(row_dict.get(_branch_col_name, ""))
        _is_new_branch = (_prev_branch is not None) and (_branch != _prev_branch)
        _cls = "new-branch" if _is_new_branch else ""

        # اگه شعبه جدید → یک ردیف هدر برای شعبه
        if _is_new_branch and per_branch_msg:
            _total_cols = len(df.columns)
            rows_html_parts.append(
                f'<tr class="branch-header"><td colspan="{_total_cols}">'
                f'🏪 <b>{_html.escape(_branch)}</b> — {_html.escape(per_branch_msg)}'
                f'</td></tr>'
            )

        cells = []
        for c in df.columns:
            v = row_dict.get(c, "")
            _cell_cls = ""
            if c == _status_col_name:
                sv = str(v)
                if "فروش شد" in sv:
                    _cell_cls = "sold"
                elif "بدون" in sv or "نشد" in sv:
                    _cell_cls = "not-sold"
            elif "تحقق" in str(c) or "درصد انجام" in str(c):
                try:
                    _p = float(str(v).replace("%", "").replace(",", "").strip())
                    _cell_cls = "pct-good" if _p >= 100 else ("pct-mid" if _p >= 70 else "pct-bad")
                except Exception:
                    pass
            cells.append(f'<td class="{_cell_cls}">{_html.escape(str(v))}</td>')

        rows_html_parts.append(f'<tr class="{_cls}">' + ''.join(cells) + '</tr>')
        _prev_branch = _branch

    rows_html = ''.join(rows_html_parts)
    cols_html = ''.join(f'<th>{_html.escape(str(c))}</th>' for c in df.columns)
    title_html = f'<div class="banner">{_html.escape(title)}</div>' if title else ''

    footer_html = (f'<div class="footer">افق کوروش • {n_rows} ردیف</div>')
    html = f"""<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="UTF-8">
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html, body {{
        font-family: Tahoma, 'Segoe UI', Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        background: #0f0f14;
        direction: rtl;
        width: {total_width}px;
        overflow: hidden;
    }}
    .wrap {{ padding: {wrap_pad}px; }}
    .card {{
        background: #17171e;
        border: 1px solid #2a2a35;
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 8px 28px rgba(0,0,0,0.45);
    }}
    .banner {{
        background: linear-gradient(135deg, #7a001d 0%, #E6003E 55%, #ff4d79 100%);
        color: #ffffff;
        text-align: center;
        font-size: {banner_fs}px;
        font-weight: 900;
        padding: {banner_pad};
        text-shadow: 0 2px 6px rgba(0,0,0,0.35);
        border-bottom: 3px solid #ff1f5a;
    }}
    table {{ border-collapse: collapse; width: 100%; font-size: {font_size}px; }}
    thead th {{
        background: linear-gradient(180deg, #2b2b38 0%, #20202a 100%);
        color: #ffd6df;
        padding: {th_pad};
        text-align: center;
        font-weight: 900;
        white-space: nowrap;
        height: {header_h}px;
        font-size: {font_size + 1}px;
        border-bottom: 2px solid #E6003E;
    }}
    tbody td {{
        padding: {td_pad};
        text-align: center;
        color: #ffffff;
        white-space: nowrap;
        height: {row_h}px;
        font-weight: 600;
        font-size: {font_size}px;
        border-bottom: 1px solid #262631;
    }}
    tbody tr:nth-child(odd) td {{ background: #18181f; }}
    tbody tr:nth-child(even) td {{ background: #1f1f28; }}
    tbody td:first-child {{ text-align: right; padding-right: 16px; font-weight: 800; color: #f4f4f8; }}
    tbody tr.new-branch td {{ border-top: 2px solid #E6003E !important; }}
    tbody td.sold {{ color: #34d399; font-weight: 900; }}
    tbody td.not-sold {{ color: #fbbf24; font-weight: 900; }}
    tbody td.pct-good {{ color: #34d399; font-weight: 900; }}
    tbody td.pct-mid {{ color: #fbbf24; font-weight: 900; }}
    tbody td.pct-bad {{ color: #fb7185; font-weight: 900; }}
    tbody tr.branch-header td {{
        background: linear-gradient(90deg, #E6003E 0%, #7a001d 100%) !important;
        color: #ffffff !important;
        text-align: right !important;
        padding: 10px 16px !important;
        font-size: {font_size + 1}px !important;
        font-weight: 900 !important;
        border-top: 3px solid #ff1f5a !important;
    }}
    .footer {{
        text-align: center;
        color: #8b8b9a;
        font-size: {max(font_size - 3, 12)}px;
        padding: {foot_pad}px;
        background: #13131a;
        border-top: 1px solid #2a2a35;
    }}
</style>
</head>
<body><div class="wrap"><div class="card">
    {title_html}
    <table>
        <thead><tr>{cols_html}</tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
    {footer_html}
</div></div></body>
</html>"""

    out_dir = Path(output_path).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_html = out_dir / "_tmp_report.html"
    tmp_html.write_text(html, encoding="utf-8")

    chrome = _find_chrome()
    if not chrome:
        raise RuntimeError("Chrome پیدا نشد")

    abs_html = tmp_html.resolve().as_uri()
    abs_out = str(Path(output_path).resolve())

    cmd = [
        chrome,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--no-sandbox",
        f"--force-device-scale-factor={scale}",
        f"--window-size={total_width + 20},{total_height + 20}",
        f"--screenshot={abs_out}",
        "--default-background-color=FF15151A",
        "--virtual-time-budget=1500",
        abs_html,
    ]
    subprocess.run(cmd, capture_output=True, timeout=90)

    try:
        tmp_html.unlink(missing_ok=True)
    except Exception:
        pass

    if not Path(output_path).exists():
        raise RuntimeError("عکس ساخته نشد")

    return output_path


def send_dataframe_as_images(chat_id, df, chunk_size=50, delay_seconds=30,
                              title_prefix="گزارش فروش", key_suffix="", compact=False):
    """DataFrame رو به صورت چند عکس ۵۰ ردیفی با فاصله ۳۰ ثانیه می‌فرسته"""
    total_rows = len(df)
    if total_rows == 0:
        return {"sent": 0, "errors": [], "total_chunks": 0}

    n_chunks = (total_rows + chunk_size - 1) // chunk_size
    out_dir = Path(".bale_images")
    out_dir.mkdir(exist_ok=True)

    sent = 0
    errors = []

    for i in range(n_chunks):
        start = i * chunk_size
        end = min(start + chunk_size, total_rows)
        chunk = df.iloc[start:end].copy()

        title = f"{title_prefix} — بخش {i+1} از {n_chunks}"

        img_path = out_dir / f"report_{key_suffix}_{i+1}_{datetime.now():%Y%m%d_%H%M%S}.png"
        try:
            df_to_image(chunk, img_path, title=title, font_size=15, compact=compact)
            caption = f"{title}\n({len(chunk)} ردیف)"
            send_photo(chat_id, str(img_path), caption=caption)
            sent += 1
        except Exception as e:
            errors.append(f"بخش {i+1}: {e}")

        # بین عکس‌ها صبر کن (نه بعد از آخری)
        if i < n_chunks - 1:
            time.sleep(delay_seconds)

    return {"sent": sent, "errors": errors, "total_chunks": n_chunks}

def _req(method, url, *args, **kwargs):
    """wrapper برای requests — method: post/get/..."""
    fn = getattr(requests, method.lower(), None)
    if fn is None:
        raise ValueError(f"متد نامعتبر: {method}")
    return fn(url, *args, **kwargs)


def send_per_branch_images(chat_id, df, branch_col_name="نام شعبه",
                            chunk_size=25, delay_seconds=5,
                            per_branch_msg="لطفاً اقدام به پالت چینی نمایید",
                            title_prefix="گزارش فروش پروژه", key_suffix="", compact=False):
    """ارسال جدا برای هر شعبه — یک یا چند عکس با هدر شعبه"""
    import time as _t
    from pathlib import Path
    from datetime import datetime

    if df is None or df.empty:
        return {"sent": 0, "errors": ["داده خالی"], "total_branches": 0}

    # اطمینان از وجود ستون شعبه
    if branch_col_name not in df.columns:
        for c in df.columns:
            if "نام شعبه" in str(c):
                branch_col_name = c
                break
        else:
            return {"sent": 0, "errors": ["ستون نام شعبه پیدا نشد"], "total_branches": 0}

    # گروه‌بندی بر اساس شعبه
    branches = df.groupby(branch_col_name, sort=False)
    n_branches = len(branches)

    out_dir = Path(".bale_images")
    out_dir.mkdir(exist_ok=True)

    sent = 0
    errors = []

    for idx, (branch_name, branch_df) in enumerate(branches):
        branch_df = branch_df.reset_index(drop=True)
        n_rows = len(branch_df)
        n_chunks = (n_rows + chunk_size - 1) // chunk_size

        for ci in range(n_chunks):
            start = ci * chunk_size
            end = min(start + chunk_size, n_rows)
            chunk = branch_df.iloc[start:end].copy()

            part_info = f" (بخش {ci+1}/{n_chunks})" if n_chunks > 1 else ""
            title = f"{title_prefix} — {branch_name}{part_info}"

            img_path = out_dir / f"br_{idx+1}_{ci+1}_{datetime.now():%H%M%S}.png"
            try:
                df_to_image(chunk, str(img_path), title=title,
                            font_size=11, per_branch_msg=per_branch_msg, compact=compact)
                caption = f"🏪 {branch_name} | {per_branch_msg}"
                send_photo(chat_id, str(img_path), caption=caption)
                sent += 1
            except Exception as e:
                errors.append(f"{branch_name} - بخش {ci+1}: {e}")

            # فاصله بین عکس‌ها (نه آخری)
            if not (idx == n_branches - 1 and ci == n_chunks - 1):
                _t.sleep(delay_seconds)

    return {"sent": sent, "errors": errors, "total_branches": n_branches}

