"""
ارسال پیام و فایل به بله — بدون کتابخانه سنگین، فقط requests
"""
import mimetypes
import requests
from pathlib import Path

BALE_API = "https://tapi.bale.ai/bot{token}/{method}"


def _req(method, url, **kw):
    """درخواست HTTP — در خطا، آدرس (که توکن داخلشه) نشت نمی‌کنه"""
    try:
        return requests.request(method, url, **kw)
    except requests.RequestException as e:
        raise RuntimeError(f"Bale connection error ({type(e).__name__})") from None


def _check(r):
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}")


def _req(method, url, **kw):
    """درخواست HTTP — در خطا، آدرس (که توکن داخلشه) نشت نمی‌کنه"""
    try:
        return requests.request(method, url, **kw)
    except requests.RequestException as e:
        raise RuntimeError(f"Bale connection error ({type(e).__name__})") from None


def _check(r):
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}")


def _url(token, method):
    return BALE_API.format(token=token, method=method)


def _load_token():
    """توکن: اول متغیر محیطی BALE_TOKEN، بعد config.ini کنار همین فایل"""
    import os
    import configparser
    v = os.environ.get("BALE_TOKEN", "").strip()
    if v:
        return v
    cfg = configparser.ConfigParser()
    cfg.read(Path(__file__).parent / "config.ini", encoding="utf-8")
    return cfg.get("BALE", "token", fallback="").strip()


def get_me(token=None):
    """اطلاعات ربات"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده. در config.ini بخش [BALE] اضافه کن.")
    r = _req("get", _url(token, "getMe"), timeout=15)
    _check(r)
    return r.json()


def get_updates(token=None, offset=None):
    """آخرین پیام‌های دریافتی ربات (برای گرفتن chat_id)"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    params = {"timeout": 5}
    if offset is not None:
        params["offset"] = offset
    r = _req("get", _url(token, "getUpdates"), params=params, timeout=20)
    _check(r)
    return r.json().get("result", [])


def send_message(chat_id, text, token=None, parse_mode="HTML"):
    """ارسال پیام متنی"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": parse_mode,
    }
    r = _req("post", _url(token, "sendMessage"), json=payload, timeout=20)
    if r.status_code != 200 and parse_mode:
        # احتمالا HTML خراب بوده (مثلا < یا & در نام) — بدون فرمت دوباره تلاش کن
        import re as _re
        payload["text"] = _re.sub(r"<[^>]+>", "", str(text))
        payload.pop("parse_mode", None)
        r = _req("post", _url(token, "sendMessage"), json=payload, timeout=20)
    if r.status_code != 200 and parse_mode:
        # احتمالا HTML خراب بوده (مثلا < یا & در نام) — بدون فرمت دوباره تلاش کن
        import re as _re
        payload["text"] = _re.sub(r"<[^>]+>", "", str(text))
        payload.pop("parse_mode", None)
        r = _req("post", _url(token, "sendMessage"), json=payload, timeout=20)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()


def send_document(chat_id, file_path, caption="", token=None):
    """ارسال فایل (Excel / PDF / ...)"""
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(file_path)
    mime, _ = mimetypes.guess_type(p.name)
    with open(p, "rb") as f:
        files = {"document": (p.name, f, mime or "application/octet-stream")}
        data = {"chat_id": chat_id, "caption": caption}
        r = _req("post", _url(token, "sendDocument"), data=data, files=files, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()


def collect_chat_ids(token=None):
    """
    همه chat_idهایی که به ربات پیام دادن رو برمی‌گردونه.
    خروجی: لیستی از دیکشنری با کلیدهای chat_id، name، username
    """
    updates = get_updates(token=token)
    seen = {}
    for u in updates:
        msg = u.get("message") or u.get("edited_message") or u.get("channel_post")
        if not msg:
            continue
        chat = msg.get("chat") or {}
        cid = chat.get("id")
        if not cid or cid in seen:
            continue
        name = chat.get("first_name") or chat.get("title") or ""
        last = chat.get("last_name") or ""
        full = (name + " " + last).strip()
        seen[cid] = {
            "chat_id": cid,
            "name": full,
            "username": chat.get("username") or "",
        }
    return list(seen.values())