"""تبدیل DataFrame به عکس و ارسال به بله"""
import time
from pathlib import Path
from datetime import datetime

import pandas as pd

from bale_sender import _load_token, _url, _req
import requests




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
    """پیدا کردن Chrome روی سیستم — همه مسیرهای ممکن"""
    from pathlib import Path
    import os

    local = os.environ.get("LOCALAPPDATA", "")
    pf = os.environ.get("PROGRAMFILES", "C:\\Program Files")
    pfx86 = os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)")

    candidates = [
        # Chrome
        os.path.join(local, r"Google\Chrome\Application\chrome.exe"),
        os.path.join(pf, r"Google\Chrome\Application\chrome.exe"),
        os.path.join(pfx86, r"Google\Chrome\Application\chrome.exe"),
        # Edge (به عنوان fallback)
        os.path.join(pfx86, r"Microsoft\Edge\Application\msedge.exe"),
        os.path.join(pf, r"Microsoft\Edge\Application\msedge.exe"),
        os.path.join(local, r"Microsoft\Edge\Application\msedge.exe"),
    ]
    for c in candidates:
        if c and Path(c).exists():
            return c
    return None


def df_to_image(df, output_path, title="", font_size=None, cell_padding=0.15):
    """DataFrame → PNG با تم دارک افق کوروش + خط قرمز بین شعبه‌ها"""
    import html as _html
    import subprocess
    from pathlib import Path

    if len(df) == 0:
        raise ValueError("داده‌ای برای نمایش نیست")

    n_rows = len(df)
    if font_size is None:
        font_size = 14

    # ── محاسبه عرض هر ستون ──
    char_w = font_size * 0.68
    col_widths = []
    for c in df.columns:
        max_len = len(str(c))
        for v in df[c].astype(str):
            max_len = max(max_len, len(str(v)))
        col_widths.append(max(80, int(max_len * char_w) + 30))

    total_width = max(760, min(5000, sum(col_widths)))

    header_h = 52
    row_h = 42
    title_h = 60 if title else 0
    total_height = header_h + (n_rows * row_h) + title_h + 40

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
            cells.append(f'<td class="{_cell_cls}">{_html.escape(str(v))}</td>')

        rows_html_parts.append(f'<tr class="{_cls}">' + ''.join(cells) + '</tr>')
        _prev_branch = _branch

    rows_html = ''.join(rows_html_parts)
    cols_html = ''.join(f'<th>{_html.escape(str(c))}</th>' for c in df.columns)
    title_html = f'<div class="banner">{_html.escape(title)}</div>' if title else ''

    html = f"""<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="UTF-8">
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html, body {{
        font-family: Tahoma, Arial, sans-serif;
        background: #15151a;
        direction: rtl;
        margin: 0;
        padding: 0;
        width: {total_width}px;
        overflow: hidden;
    }}
    .wrap {{ padding: 10px; }}
    .banner {{
        background: linear-gradient(135deg, #A30029 0%, #E6003E 100%);
        color: #ffffff;
        text-align: center;
        font-size: {font_size + 6}px;
        font-weight: 900;
        padding: 12px 10px;
        border-radius: 10px;
        margin-bottom: 8px;
        letter-spacing: 0.5px;
        box-shadow: 0 4px 12px rgba(230,0,62,0.3);
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        font-size: {font_size}px;
        border-radius: 8px;
        overflow: hidden;
    }}
    thead th {{
        background: #A30029;
        color: #ffffff;
        padding: 10px 8px;
        text-align: center;
        font-weight: 900;
        border: none;
        white-space: nowrap;
        height: {header_h}px;
        font-size: {font_size + 2}px;
    }}
    tbody td {{
        padding: 8px 8px;
        text-align: center;
        border: none;
        color: #e5e7eb;
        white-space: nowrap;
        height: {row_h}px;
        font-weight: 600;
        font-size: {font_size}px;
    }}
    tbody tr:nth-child(odd) td {{
        background: #1a1a1f;
    }}
    tbody tr:nth-child(even) td {{
        background: #21212a;
    }}
    tbody tr.new-branch td {{
        border-top: 2px solid #E6003E !important;
    }}
    tbody td.sold {{
        color: #10b981;
        font-weight: 900;
    }}
    tbody td.not-sold {{
        color: #f59e0b;
        font-weight: 900;
    }}
</style>
</head>
<body><div class="wrap">
    {title_html}
    <table>
        <thead><tr>{cols_html}</tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
</div></body>
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
        "--force-device-scale-factor=2",
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
                              title_prefix="گزارش فروش", key_suffix=""):
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
            df_to_image(chunk, img_path, title=title, font_size=15)
            caption = f"{title}\n({len(chunk)} ردیف)"
            send_photo(chat_id, str(img_path), caption=caption)
            sent += 1
        except Exception as e:
            errors.append(f"بخش {i+1}: {e}")

        # بین عکس‌ها صبر کن (نه بعد از آخری)
        if i < n_chunks - 1:
            time.sleep(delay_seconds)

    return {"sent": sent, "errors": errors, "total_chunks": n_chunks}