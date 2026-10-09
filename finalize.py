# finalize.py — پاکسازی نهایی برای استقرار
import shutil
from pathlib import Path

print("=" * 60)
print("🧹 پاکسازی نهایی برای سرور")
print("=" * 60)

# ── فایل‌های حذفی ──
files_to_delete = [
    'add_branch_ui.py',
    'add_help_page.py',
    'add_personel_upload.py',
    'cleanup.py',
    'New Text Document.txt',
    'fix_all_rial.py',
    'fix_bale_path.py',
    'fix_bale_token.py',
    'fix_branch_image.py',
    'fix_browser_back.py',
    'fix_caption.py',
    'fix_dual_sort.py',
    'fix_home_cards_pos.py',
    'fix_pdf_back.py',
    'fix_per_branch_incremental.py',
    'fix_project_direct.py',
    'fix_project_everywhere.py',
    'fix_project_sold_only.py',
    'fix_remove_cols.py',
    'fix_report_pw.py',
    'fix_resume.py',
    'fix_send_pdf.py',
    'fix_static.py',
    'fix_trend_all_formats.py',
    'fix_target_columns.py',
    'fix_target_order.py',
    'fix_target_order_v2.py',
    'fix_req_auto.py',
    'fix_req_wrapper_v2.py',
    'fix_fa_digits_simple.py',
    'fix_restore_report.py',
    'fix_bale_req.py',
    'fix_trend_last2_ui.py',
    'fix_trend_last2_in_report.py',
    'fix_trend_last2_target_too.py',
    'fix_all_trend.py',
    'fix_trend_selectable_all.py',
    'check_order.py',
    'check2.py',
    'check3.py',
    'check_img_tool.py',
    'show_req.py',
]

# ── پوشه‌های cache ──
dirs_to_delete = [
    '.bale_docs',
    '.bale_images',
    '.excel_cache',
    '__pycache__',
    '.venv',
    'backups',
    '.pytest_cache',
]

# ── فایل‌های حیاتی (نگه‌داری) ──
critical = [
    'app.py', 'report_sender.py', 'send_report.py', 'image_sender.py',
    'online_section.py', 'email_manager.py', 'email_section.py',
    'bale_sender.py', 'bale_chat_id_helper.py', 'personel_loader.py',
    'config_loader.py',
    'config.ini', 'contacts.json', 'recipients.json', 'logo.png',
    'requirements.txt', 'packages.txt',
    '.gitignore', '.streamlit',
    '1405-07-15 projraked.xlsx', 'فایل تارگت.xlsx',
    'گزارش عادی و فرانچایز 14050715.xlsx', 'گزارش فروش راکد.xlsx',
    '01_internet_sales.xlsx', 'مغایرت گیری - ادجاست.xlsx',
]

# ── نمایش ──
print(f"\n📦 فایل‌های قابل حذف:")
found_files = []
for name in files_to_delete:
    p = Path(name)
    if p.exists():
        print(f"  • {name}")
        found_files.append(p)

print(f"\n📁 پوشه‌های قابل حذف:")
found_dirs = []
for name in dirs_to_delete:
    p = Path(name)
    if p.exists():
        print(f"  • {name}/")
        found_dirs.append(p)

print(f"\n✅ فایل‌های حیاتی:")
missing = []
for name in critical:
    exists = Path(name).exists()
    icon = "✅" if exists else "❌"
    print(f"  {icon} {name}")
    if not exists:
        missing.append(name)

if missing:
    print(f"\n⚠️  این‌ها گم شدن: {missing}")
    print("   قبل از حذف بررسی کن!")

print()
ans = input("⚠️  حذف بشه؟ (yes/no): ").strip().lower()
if ans not in ('yes', 'y'):
    print("❌ لغو شد")
    raise SystemExit(0)

# ── حذف ──
count = 0
for f in found_files:
    try:
        f.unlink()
        count += 1
    except Exception as e:
        print(f"  ⚠️ {f.name}: {e}")

for d in found_dirs:
    try:
        shutil.rmtree(d)
        print(f"  ✅ حذف شد: {d}/")
    except Exception as e:
        print(f"  ⚠️ {d}: {e}")

print()
print("=" * 60)
print(f"🎉 {count} فایل + {len(found_dirs)} پوشه حذف شد")
print("=" * 60)

# ── ساختار نهایی ──
print("\n📂 ساختار نهایی پروژه:")
for f in sorted(Path('.').iterdir()):
    if f.name.startswith('.') and f.name not in ('.gitignore', '.streamlit'):
        continue
    if f.name in ('__pycache__', '.venv'):
        continue
    if f.is_file():
        size = f.stat().st_size
        if size > 1024 * 1024:
            size_str = f"{size / 1024 / 1024:.1f} MB"
        elif size > 1024:
            size_str = f"{size / 1024:.0f} KB"
        else:
            size_str = f"{size} B"
        print(f"  📄 {f.name:55} {size_str:>10}")
    elif f.is_dir():
        print(f"  📁 {f.name}/")