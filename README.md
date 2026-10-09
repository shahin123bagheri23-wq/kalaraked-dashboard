# 🛒 داشبورد مدیریت کالای راکد — افق کوروش

داشبورد مدیریت و تحلیل کالاهای راکد فروشگاه‌های زنجیره‌ای افق کوروش.

---

## 🎯 امکانات

- 📊 نمایش داده‌های راکد ۶۰ و ۴۵ روزه
- 🏪 تحلیل عملکرد فروشگاه‌ها
- 👤 تحلیل عملکرد سوپروایزرها
- 🎯 تارگت و رتبه‌بندی شعب
- 📋 چک‌لیست شیفت صبح و عصر
- 📈 روند و مقایسه درصد راکد
- 📦 پروژه پرزنتی راکد
- 📨 ارسال گزارش به بله و ایمیل
- 📸 گزارش تصویری با فرمت اکسل
- 🔐 ورود با کد پرسنلی (نقش‌محور)
- 📝 لاگ فعالیت کاربران

---

## 🏗 معماری

```
app.py                  ← فایل اصلی Streamlit
├── report_sender.py    ← صفحه ارسال گزارش
├── image_sender.py     ← تولید عکس از دیتافریم
├── send_report.py      ← ساخت گزارش HTML + ارسال SMTP
├── email_manager.py    ← مدیریت مخاطبین
├── bale_sender.py      ← ارسال پیام بله
├── personel_loader.py  ← لود فایل پرسنلی
└── config_loader.py    ← تنظیمات
```

---

## 🚀 نصب

```bash
git clone https://github.com/shahin123bagheri23-wq/kalaraked-dashboard
cd kalaraked-dashboard
pip install -r requirements.txt
python -m streamlit run app.py
```

---

## 🌐 استقرار

### Streamlit Cloud (نیاز به VPN)
1. share.streamlit.io
2. New app → انتخاب مخزن
3. Main file: app.py
4. Settings → Secrets → محتوای config.ini

### Render (بدون VPN)
1. dashboard.render.com
2. New Web Service → GitHub
3. Build: pip install -r requirements.txt
4. Start: python -m streamlit run app.py --server.port $PORT

---

## 🔐 دسترسی‌ها

| نقش | صفحه |
|---|---|
| مسئول فروشگاه | 🏪 شعبه من |
| کارمند فروشگاه | 🏪 شعبه من |
| سرپرست فروش | 🏠 داشبورد کامل |
| مدیر دیستریکت | 🏠 داشبورد کامل |
| کارشناس | 🏠 داشبورد کامل |

---

## 📞 پشتیبانی

**شاهین باقری فرد** — کد پرسنلی: 905939
**واحد:** پشتیبانی زنجیره تامین دیستریکت خرم آباد

**نسخه:** 1.0
