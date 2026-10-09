"""
ارسال پیام و فایل به بله — با requests
توکن: به صورت هاردکد + قابل تنظیم از secrets/env/config.ini
"""
import mimetypes
import requests
from pathlib import Path

BALE_API = "https://tapi.bale.ai/bot{token}/{method}"

# ⚠️ توکن هاردکد شده (پیش‌فرض)
_HARDCODED_TOKEN = "1448976780:DhxKe0ttzEUJPHbmtElQDIYtvZNiorx-c3k"


def _url(token, method):
    return BALE_API.format(token=token, method=method)


def _load_token():
    """توکن بله — اولویت: secrets > env > config.ini > هاردکد"""
    import os
    from pathlib import Path as _P

    # ۱. Streamlit Secrets
    try:
        import streamlit as st
        token = st.secrets.get("BALE_TOKEN")
        if token:
            return token
        try:
            token = st.secrets["BALE"]["token"]
            if token:
                return token
        except Exception:
            pass
    except Exception:
        pass

    # ۲. Environment
    token = os.environ.get("BALE_TOKEN")
    if token:
        return token

    # ۳. config.ini (کنار خود فایل + cwd)
    try:
        import configparser
        _base = _P(__file__).resolve().parent
        for cfg_path in [_base / "config.ini", _P.cwd() / "config.ini"]:
            if cfg_path.exists():
                cfg = configparser.ConfigParser()
                cfg.read(str(cfg_path), encoding="utf-8")
                token = cfg.get("BALE", "token", fallback="")
                if token:
                    return token
    except Exception:
        pass

    # ۴. هاردکد (پشتیبان)
    return _HARDCODED_TOKEN


def get_me(token=None):
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    r = requests.get(_url(token, "getMe"), timeout=15)
    r.raise_for_status()
    return r.json()


def get_updates(token=None, offset=None):
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    params = {"timeout": 5}
    if offset is not None:
        params["offset"] = offset
    r = requests.get(_url(token, "getUpdates"), params=params, timeout=20)
    r.raise_for_status()
    return r.json().get("result", [])


def send_message(chat_id, text, token=None, parse_mode="HTML"):
    token = token or _load_token()
    if not token:
        raise ValueError("توکن بله تنظیم نشده.")
    payload = {"chat_id": chat_id, "text": text, "parse_mode": parse_mode}
    r = requests.post(_url(token, "sendMessage"), json=payload, timeout=20)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()


def send_document(chat_id, file_path, caption="", token=None):
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
        r = requests.post(_url(token, "sendDocument"), data=data, files=files, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Bale API error {r.status_code}: {r.text}")
    return r.json()


def collect_chat_ids(token=None):
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