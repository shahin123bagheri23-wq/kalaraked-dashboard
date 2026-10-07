import re

NEW_LOAD_ALL = '''@st.cache_data(show_spinner="در حال بارگذاری داده‌ها…", max_entries=4)
def load_all(file_path, mtime):
    out = {"d60": None, "d45": None, "pq45": None, "pq60": None,
           "summary60": None, "sales": None, "sheets": []}
    if not Path(file_path).exists():
        return out
    try:
        xls = pd.ExcelFile(file_path)
    except Exception:
        return out
    out["sheets"] = list(xls.sheet_names)

    cache_dir = Path(str(file_path)).parent / ".excel_cache"
    try:
        cache_dir.mkdir(exist_ok=True)
    except Exception:
        pass
    try:
        excel_mtime = Path(file_path).stat().st_mtime
    except Exception:
        excel_mtime = 0

    def _read_cached(sheet_name):
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", str(sheet_name))[:60]
        cache_file = cache_dir / f"{safe}.pkl"
        try:
            if cache_file.exists() and cache_file.stat().st_mtime >= excel_mtime:
                return pd.read_pickle(cache_file)
        except Exception:
            pass
        df = pd.read_excel(xls, sheet_name=sheet_name)
        try:
            df.to_pickle(cache_file)
        except Exception:
            pass
        return df

    try:
        s60 = find_sheet_exact(xls, SHEET_TBL60)
        s45 = find_sheet_exact(xls, SHEET_TBL45)
        if s60 and s45:
            out["d60"] = _prep_raaked(_read_cached(s60))
            out["d45"] = _prep_raaked(_read_cached(s45))
    except Exception as e:
        st.warning(f"⚠️ خطا در tbl60/tbl45: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_PQ45)
        if s:
            out["pq45"] = _prep_pq(_read_cached(s), "صبح")
    except Exception as e:
        st.warning(f"⚠️ خطا در pq45: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_PQ60)
        if s:
            out["pq60"] = _prep_pq(_read_cached(s), "عصر")
    except Exception as e:
        st.warning(f"⚠️ خطا در pq60: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_SUMMARY60)
        if s:
            out["summary60"] = _prep_summary60(_read_cached(s))
    except Exception as e:
        st.warning(f"⚠️ خطا در خلاصه ۶۰ روزه: {e}")
    try:
        s = find_sheet_exact(xls, SHEET_SALES)
        if s:
            out["sales"] = _prep_sales(_read_cached(s))
    except Exception:
        pass
    return out
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

s, e = find_func_range(lines, 'load_all')
if s is None:
    print("NOT FOUND: load_all")
else:
    lines[s:e + 1] = [NEW_LOAD_ALL.rstrip() + '\n', '\n']
    print(f"replaced load_all: lines {s+1}-{e+1}")

open(p, 'w', encoding='utf-8').writelines(lines)
print("done")