# 🏥 نظام تتبع التغذية — Nutrition Tracker

تطبيق Windows لتسجيل ومتابعة زيارات التغذية، محوّل من VBA Excel.

---

## ✅ المتطلبات

- Python 3.10 أو أعلى
- Windows 10/11

---

## ⚙️ خطوات التشغيل

### 1. تثبيت Python
- حمّل من: https://python.org/downloads
- **مهم جداً:** ضع علامة ✅ على **Add Python to PATH** أثناء التثبيت

### 2. تثبيت المكتبات
افتح **Command Prompt** واكتب:
```
pip install customtkinter openpyxl
```

### 3. تشغيل التطبيق
```
python main.py
```

---

## 📁 هيكل الملفات

```
nutrition-tracker/
├── main.py           ← نقطة التشغيل
├── app.py            ← الواجهة الرئيسية
├── data_manager.py   ← إدارة ملف Excel
├── requirements.txt  ← المكتبات المطلوبة
└── README.md
```

---

## 📋 الميزات

- ✅ بحث عن مريض وتحديث عمره تلقائياً
- ✅ منطق MUAC → ZScore → Diagnosis → Action تلقائي
- ✅ قيود إدخال رقمي
- ✅ تسجيل مرضى جدد
- ✅ بحث في LocalData
- ✅ تصدير Excel إلى سطح المكتب
- ✅ يعمل مع ملف Nutrition_Data.xlsx الحالي مباشرة
