p = 'report_sender.py'
s = open(p, encoding='utf-8').read()

# پچ ۱: تابع سوپروایزر
old1 = '''    t[sup_col] = normalize_name(t[sup_col])
    t[branch_col] = normalize_name(t[branch_col])
    t[ach_col] = parse_percent(t[ach_col])
    if raked_col:
        t[raked_col] = to_number(t[raked_col])

    sub = t[t[sup_col].astype(str).str.strip() == supervisor_name.strip()].copy()'''

new1 = '''    t[sup_col] = normalize_name(t[sup_col])
    t[branch_col] = normalize_name(t[branch_col])
    t[ach_col] = parse_percent(t[ach_col])
    if raked_col:
        t[raked_col] = to_number(t[raked_col])

    # حذف ردیف‌های خالی، جمع و صفر
    _nm = t[branch_col].astype(str).str.strip()
    _is_total = _nm.str.contains(r"^\\s*(?:جمع|مجموع|کل|total|grand|sum)", case=False, regex=True)
    _is_empty = t[branch_col].isna() | _nm.isin(["", "nan", "None", "NaN"])
    _is_zero = to_number(t[ach_col]) == 0
    t = t[~(_is_total | _is_empty | _is_zero)].copy()

    sub = t[t[sup_col].astype(str).str.strip() == supervisor_name.strip()].copy()'''

s = s.replace(old1, new1)

# پچ ۲: تابع مدیر
old2 = '''    # حذف ردیف‌های جمع
    name = t[branch_col].astype(str)
    is_total = name.str.contains(r"^\\s*(?:جمع|مجموع|کل|total|grand)",
                                  case=False, regex=True)
    t = t[~is_total].copy()'''

new2 = '''    # حذف ردیف‌های خالی، جمع و صفر
    name = t[branch_col].astype(str).str.strip()
    is_total = name.str.contains(r"^\\s*(?:جمع|مجموع|کل|total|grand|sum)", case=False, regex=True)
    is_empty = t[branch_col].isna() | name.isin(["", "nan", "None", "NaN"])
    is_zero = to_number(t[ach_col]) == 0
    t = t[~(is_total | is_empty | is_zero)].copy()'''

s = s.replace(old2, new2)

# پچ ۳: تابع عادی
old3 = '''    t[ach_col] = parse_percent(t[ach_col])
    vals = t[ach_col].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success = int((vals >= 100).sum())
    total = len(vals)'''

new3 = '''    t[ach_col] = parse_percent(t[ach_col])
    _branch = next((c for c in t.columns if "نام شعبه" in c), None)
    if _branch:
        _nm = t[_branch].astype(str).str.strip()
        _empty = t[_branch].isna() | _nm.isin(["", "nan", "None", "NaN"])
        t = t[~_empty].copy()
    vals = t[ach_col].dropna()
    vals = vals[vals != 0]

    mean_ach = vals.mean() if len(vals) else 0
    success = int((vals >= 100).sum())
    total = len(vals)'''

s = s.replace(old3, new3)

open(p, 'w', encoding='utf-8').write(s)
print('patched OK')