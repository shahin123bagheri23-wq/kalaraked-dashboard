"""
لایه تنظیمات هوشمند
اولویت: Environment Variable > config.ini > مقدار پیش‌فرض
"""
import os
import configparser
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.ini"

def _load_ini():
    cp = configparser.ConfigParser()
    if CONFIG_PATH.exists():
        cp.read(CONFIG_PATH, encoding="utf-8")
    return cp

_ini = _load_ini()

def get(section: str, key: str, default: str = "", env_name: str = None) -> str:
    """
    مقدار رو با اولویت env می‌خونه.
    env_name: اگه بدی، از متغیر محیطی با این اسم می‌خونه.
              اگه ندی، از SECTION_KEY می‌سازه.
    """
    if env_name is None:
        env_name = f"{section}_{key}".upper()

    val = os.environ.get(env_name)
    if val:
        return val.strip()

    if _ini.has_section(section) and _ini.has_option(section, key):
        return _ini.get(section, key).strip()

    return default

def get_int(section: str, key: str, default: int = 0, env_name: str = None) -> int:
    v = get(section, key, "", env_name)
    try:
        return int(v)
    except (ValueError, TypeError):
        return default

def get_list(section: str, key: str, default: list = None, sep: str = ",") -> list:
    v = get(section, key, "")
    if not v:
        return default or []
    return [x.strip() for x in v.split(sep) if x.strip()]

# ---------------- SMTP ----------------
SMTP_SERVER   = get("SMTP", "server", "smtp.gmail.com")
SMTP_PORT     = get_int("SMTP", "port", 465)
SMTP_USER     = get("SMTP", "sender_email", "")
SMTP_PASS     = get("SMTP", "sender_password", "")

# ---------------- REPORT ----------------
REPORT_RECIPIENTS = get_list("REPORT", "recipients")
REPORT_SUBJECT    = get("REPORT", "subject", "گزارش هفتگی")
REPORT_DAY        = get_int("REPORT", "day_of_week", 5)   # 0=دوشنبه ... 5=شنبه
REPORT_HOUR       = get_int("REPORT", "send_hour", 8)
REPORT_MINUTE     = get_int("REPORT", "send_minute", 0)

# ---------------- BALE ----------------
BALE_TOKEN = get("BALE", "token", "")
BALE_API   = f"https://tapi.bale.ai/bot{BALE_TOKEN}"

# ---------------- ADMIN ----------------
ADMIN_PASSWORD  = get("ADMIN", "password", "")
UPLOAD_PASSWORD = get("ADMIN", "upload_password", ADMIN_PASSWORD)

# ---------------- دیباگ ----------------
def debug_status() -> dict:
    """وضعیت لود شدن تنظیمات رو نشون می‌ده (بدون افشای رمز)"""
    return {
        "config.ini exists": CONFIG_PATH.exists(),
        "SMTP_USER": SMTP_USER or "(خالی)",
        "SMTP_PASS": "✅" if SMTP_PASS else "❌",
        "BALE_TOKEN": "✅" if BALE_TOKEN else "❌",
        "ADMIN_PASSWORD": "✅" if ADMIN_PASSWORD else "❌",
        "REPORT_RECIPIENTS": len(REPORT_RECIPIENTS),
    }

if __name__ == "__main__":
    import json
    print(json.dumps(debug_status(), ensure_ascii=False, indent=2))
