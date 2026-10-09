"""لود کردن داده پرسنلی از فایل اکسل"""
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).parent
# چند مسیر ممکن
_CANDIDATE_NAMES = [
    "گزارش عادی و فرانچایز 14050715.xlsx",
    "personel.xlsx",
    "personnel.xlsx",
]

def _find_personel_file():
    """پیدا کردن فایل پرسنلی در چند مسیر"""
    import os
    roots = [BASE_DIR, Path.cwd()]
    env_dir = os.environ.get("DATA_DIR")
    if env_dir:
        roots.insert(0, Path(env_dir))
    for root in roots:
        for name in _CANDIDATE_NAMES:
            f = root / name
            if f.exists():
                return f
    return BASE_DIR / _CANDIDATE_NAMES[0]

PERSONEL_FILE = _find_personel_file()


def _role_from_serial(serial):
    """تشخیص نقش بر اساس کد سریال سمت"""
    s = str(serial or "")
    if s.startswith("787"):
        return "supervisor"   # سرپرست فروش (سوپروایزر منطقه)
    if s.startswith("362"):
        return "store_manager"  # مسئول فروشگاه
    if s.startswith("3777"):
        return "district_manager"
    if s.startswith("4200"):
        return "supply_chain"
    if s.startswith("357"):
        return "store_staff"
    if s.startswith("359"):
        return "store_deputy"
    if s.startswith("3767"):
        return "store_deputy2"
    return "other"


def _extract_branch_code(location):
    """از Location (کد شعبه) → کد"""
    return str(location or "").strip().upper()


def _extract_branch_name(unit_name):
    """از «سریال نام واحد» اسم تمیز شعبه استخراج کن"""
    s = str(unit_name or "").strip()
    # حذف کد داخل پرانتز
    s = re.sub(r"\s*\([^)]*\)\s*$", "", s)
    return s.strip()


import re


def load_personel():
    """خروجی: دیکشنری {کد پرسنلی: {...}}"""
    if not PERSONEL_FILE.exists():
        return {}

    try:
        d = pd.read_excel(PERSONEL_FILE, sheet_name=0)
        d.columns = d.columns.astype(str).str.strip()

        # پیدا کردن ستون‌ها
        def find_col(keys):
            for c in d.columns:
                cl = str(c).lower()
                if any(k in cl or k in c for k in keys):
                    return c
            return None

        code_col = find_col(["کد پرسنلي", "کد پرسنلی", "کد پرسنل", "personnel"])
        loc_col = find_col(["Location", "location", "لوکیشن"])
        role_col = find_col(["سریال سمت", "سمت"])
        unit_col = find_col(["سریال نام واحد", "نام واحد", "واحد"])
        name_col = find_col(["نام", "name"])
        family_col = find_col(["نام خانوادگي", "نام خانوادگی", "خانوادگی"])
        mobile_col = find_col(["تلفن همراه", "موبایل", "همراه"])

        if not code_col:
            return {}

        out = {}
        for _, row in d.iterrows():
            code = str(row.get(code_col, "")).strip()
            if code.endswith(".0"):
                code = code[:-2]
            if not code or code == "nan":
                continue
            role = _role_from_serial(row.get(role_col, ""))
            branch_code = _extract_branch_code(row.get(loc_col, ""))
            branch_name = _extract_branch_name(row.get(unit_col, ""))
            name = str(row.get(name_col, "")).strip()
            family = str(row.get(family_col, "")).strip()
            full_name = f"{name} {family}".strip()
            mobile = str(row.get(mobile_col, "")).strip()

            out[code] = {
                "personnel_code": code,
                "name": full_name,
                "role": role,
                "branch_code": branch_code,
                "branch_name": branch_name,
                "mobile": mobile,
            }

        return out
    except Exception:
        return {}
