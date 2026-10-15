# Pipecal Pro 2.0 🚀
### سیستم جامع صدور پیش‌فاکتور و مدیریت قیمت لوله و اتصالات پلی‌اتیلن
**Polyethylene Pipes & Fittings Invoicing & Pricing Desktop System**

[![React](https://img.shields.io/badge/React-19-61DAFB?logo=react&logoColor=white)](https://react.dev/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Electron](https://img.shields.io/badge/Electron-44-47848F?logo=electron&logoColor=white)](https://www.electronjs.org/)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-4-38B2AC?logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)
[![Vite](https://img.shields.io/badge/Vite-8-646CFF?logo=vite&logoColor=white)](https://vitejs.dev/)
[![Platform](https://img.shields.io/badge/Platform-macOS%20|%20Windows-lightgrey)](#-دانلود-نرم‌افزار-download-latest-release)

---

## 📥 دانلود نرم‌افزار (Download Latest Release)

آخرین نسخه‌های آماده نصب برای مک و ویندوز از طریق تب [GitHub Releases](https://github.com/ryanyv/invoice_app/releases/latest) در دسترس هستند:

| سیستم‌عامل (OS) | نوع فایل | لینک دانلود مستقیم |
| :--- | :--- | :--- |
| 🍏 **macOS** (Apple Silicon M1/M2/M3/M4) | DMG Installer | [دانلود Pipecal Pro-2.0.0-arm64.dmg](https://github.com/ryanyv/invoice_app/releases/latest/download/Pipecal.Pro-2.0.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer (.exe) | [دانلود Pipecal Pro Setup 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/latest/download/Pipecal.Pro.Setup.2.0.0.exe) |
| 🪟 **Windows** | Portable (.exe بدون نیاز به نصب) | [دانلود Pipecal Pro 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/latest/download/Pipecal.Pro.2.0.0.exe) |

> 💡 همچنین می‌توانید تمامی نسخه‌ها و تغییرات را در [صفحه Releases پروژه](https://github.com/ryanyv/invoice_app/releases) مشاهده کنید.

---

## 🌟 امکانات و قابلیت‌ها (Features)

- **مدیریت هوشمند قیمت‌ها:**
  - ویرایش سریع نرخ پایه هر کیلوگرم پلی‌اتیلن (PE100, PE80, PE63).
  - جدول داده‌ای پیشرفته اتصالات با امکان **ویرایش درون‌خطی (Inline Edit)** برای تمام ۲۵۰+ اتصال استاندارد.
  - اعمال گروهی درصد افزایش یا کاهش قیمت بر کل کاتالوگ یا یک دسته‌بندی با یک کلیک.
  - خروجی و ورودی اکسل (CSV) برای همگام‌سازی آسان قیمت‌ها.
- **صدور پیش‌فاکتور بومی راست‌به‌چپ (RTL):**
  - فرم تعاملی محاسبه لوله با هماهنگی آنی بین متراژ و وزن بر مبنای جدول استانداردهای وزنی DIN.
  - انتخاب سریع اتصالات بر اساس دسته، محصول، فشار کاری و سایز.
  - اعمال پله‌های تخفیف درصدی یا تخفیف مبلغی دلخواه.
  - محاسبه خودکار مالیات بر ارزش افزوده و تبدیل مبلغ به حروف فارسی.
  - ثبت مشخصات مشتری، تاریخ شمسی، شرایط پرداخت و توضیحات فاکتور.
- **چاپ و خروجی PDF با سربرگ رسمی:**
  - طراحی استاندارد افقی (Landscape) منطبق با کاغذ A4 رسمی شرکت.
  - فونت استاندارد فارسی (وزیرمتن) و سازگار با انواع پرینترها.
- **آرشیو و سوابق فاکتورها:**
  - جستجو و فیلتر پیش‌فاکتورهای صادر شده بر اساس نام مشتری، تاریخ یا شماره فاکتور.
  - پشتیبان‌گیری کامل از داده‌ها به صورت فایل JSON و بازیابی سریع.

---

## 🛠️ راهنمای ساخت و اجرای پروژه (Build From Source)

اگر می‌خواهید پروژه را شخصاً کامپایل کرده یا در محیط توسعه اجرا نمایید، مراحل زیر را دنبال کنید:

### ۱. پیش‌نیازها (Prerequisites)
- **Node.js** (نسخه 18 یا بالاتر) – [دانلود Node.js](https://nodejs.org/)
- **npm** (همراه با Node.js نصب می‌شود)
- **Git**

### ۲. کلون کردن مخزن (Clone Repository)
```bash
git clone https://github.com/ryanyv/invoice_app.git
cd invoice_app
```

### ۳. نصب وابستگی‌ها (Install Dependencies)
```bash
npm install
```

### ۴. اجرای نسخه توسعه (Development Mode)

- **اجرای نسخه تحت وب در مرورگر (Vite):**
  ```bash
  npm run dev
  ```
  سپس آدرس `http://localhost:5173` را در مرورگر باز کنید.

- **اجرای کامل برنامه دسکتاپ با Electron:**
  ```bash
  npm run electron:dev
  ```

---

## 📦 کامپایل و ساخت فایل‌های نصبی (Packaging & Distribution)

برای ساخت بسته‌های نصبی نهایی سیستم‌عامل خود، دستورات زیر را استفاده کنید:

### 🍏 ساخت برای مک (macOS):
```bash
# ساخت فایل DMG و نسخه نصبی مک:
npm run electron:build:mac

# یا ساخت دایرکتوری برنامه آماده اجرا (بدون پکیج کردن در DMG):
npm run electron:pack
```
> خروجی در پوشه `release/` قرار می‌گیرد:
> - `release/Pipecal Pro-2.0.0-arm64.dmg`
> - `release/mac-arm64/Pipecal Pro.app`

### 🪟 ساخت برای ویندوز (Windows):
```bash
# ساخت فایل نصبی ویندوز (.exe Setup) و نسخه پرتابل:
npm run electron:build:win
```
> خروجی در پوشه `release/` قرار می‌گیرد:
> - `release/Pipecal Pro Setup 2.0.0.exe` (فایل نصبی استاندارد)
> - `release/Pipecal Pro 2.0.0.exe` (نسخه بدون نیاز به نصب / Portable)

---

## 📁 ساختار سورس‌کد (Project Structure)

```text
├── electron/                 # کدهای اصلی Electron (main & preload process)
├── public/                   # دارایی‌های عمومی و نشان‌واره (logo.png)
├── src/
│   ├── components/
│   │   ├── InvoiceBuilder/   # فرم ثبت و صدور پیش‌فاکتور
│   │   ├── InvoiceHistory/   # سوابق و تاریخچه پیش‌فاکتورها
│   │   ├── InvoicePrint/     # تمپلیت و پیش‌نمایش چاپ A4 رسمی
│   │   ├── PriceManager/     # مدیریت قیمت‌های پایه، کاتالوگ و درصدها
│   │   └── Settings/         # پنجره تنظیمات شرکت، مالیات و پشتیبان‌گیری
│   ├── data/                 # داده‌های پیش‌فرض و مقادیر پایه
│   ├── services/
│   │   ├── calculator.ts     # موتور محاسبات وزن DIN، متراژ، قیمت و تخفیف
│   │   ├── persianUtils.ts   # توابع تاریخ جلالی و تبدیل اعداد و حروف فارسی
│   │   └── storage.ts        # مدیریت پایگاه‌داده محلی و واردات/صادرات داده
│   └── types/                # تعاریف تایپ‌های TypeScript
├── package.json              # وابستگی‌ها و اسکریپت‌های ساخت
└── vite.config.ts            # تنظیمات باندلر Vite
```

---

## 📄 لایسنس (License)

توسعه‌یافته برای شرکت **پلی غرب** — کلیه حقوق محفوظ است.
All rights reserved © Polygharb.
