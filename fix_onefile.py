import re

NEW_DISCOVER = '''def _discover_input_files():
    global RAAKED_FILE, SALES_FILE, TARGET_FILE
    candidates = [
        DATA_DIR / "1405-07-15 projraked.xlsx",
        Path("1405-07-15 projraked.xlsx"),
        Path.cwd() / "1405-07-15 projraked.xlsx",
    ]
    one = next((c for c in candidates if c.exists()), candidates[0])
    RAAKED_FILE = one
    SALES_FILE = one
    TARGET_FILE = one
    return RAAKED_FILE, SALES_FILE, TARGET_FILE
'''

NEW_LOAD_SALES = '''@st.cache_data(show_spinner=False)
def load_sales(mtime):
    if not RAAKED_FILE.exists():
        return None
    try:
        xls = pd.ExcelFile(RAAKED_FILE)
        sh = find_sheet_exact(xls, SHEET_SALES)
        if sh:
            return _prep_sales(pd.read_excel(xls, sheet_name=sh))
    except Exception:
        pass
    return None
'''

NEW_LOAD_TARGET = '''@st.cache_data(show_spinner=False)
def load_target(mtime):
    if not RAAKED_FILE.exists():
        return None
    try:
        xls = pd.ExcelFile(RAAKED_FILE)
        sh = find_sheet_exact(xls, "تارگت") or find_sheet_exact(xls, TARGET_SHEET)
        if not sh:
            return None
        t = pd.read_excel(xls, sheet_name=sh)
        t.columns = t.columns.astype(str).str.strip()
        for c in t.columns:
            if "نام شعبه" in c or "سرپرست" in c:
                t[c] = normalize_name(t[c])
            if "ریالی راکد" in c:
                t[c] = to_number(t[c])
            if "درصد" in c or c.startswith("تغییرات") or c.startswith("تارگت") or "تحقق" in c:
                t[c] = parse_percent(t[c])
        return t
    except Exception:
        return None
'''


def find_func_range(lines, func_name):
    start = None
    header_idx = None
    for i, line in enumerate(lines):
        if re.match(rf"^def {func_name}\(", line):
            header_idx = i
            start = i
            j = i - 1
            while j >= 0 and lines[j].lstrip().startswith('@'):
                start = j
                j -= 1
            break
    if start is None:
        return None, None
    end = None
    for k in range(header_idx + 1, len(lines)):
        line = lines[k]
        if line.strip() == '':
            continue
        if not line.startswith((' ', '\t')):
            end = k - 1
            break
    if end is None:
        end = len(lines) - 1
    while end > start and lines[end].strip() == '':
        end -= 1
    return start, end


p = 'app.py'
lines = open(p, encoding='utf-8').readlines()

for name, new_code in [
    ('_discover_input_files', NEW_DISCOVER),
    ('load_sales', NEW_LOAD_SALES),
    ('load_target', NEW_LOAD_TARGET),
]:
    s, e = find_func_range(lines, name)
    if s is None:
        print(f"NOT FOUND: {name}")
        continue
    lines[s:e + 1] = [new_code.rstrip() + '\n', '\n']
    print(f"replaced {name}: lines {s+1}-{e+1}")

open(p, 'w', encoding='utf-8').writelines(lines)
print("done")