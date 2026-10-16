<p align="center">
  <img src="public/logo.png" alt="Pipecal Pro Logo" width="160" />
</p>

# Pipecal Pro 2.0 🚀
### سیستم جامع صدور پیش‌فاکتور و مدیریت قیمت لوله و اتصالات پلی‌اتیلن
**Enterprise Desktop System for Polyethylene Pipes & Fittings Pricing, Quotation & Logistics**

<p align="center">
  <a href="https://github.com/ryanyv/invoice_app/releases/latest"><img src="https://img.shields.io/badge/Release-v2.0.0-0284c7?style=flat-square&logo=github" alt="Latest Release" /></a>
  <a href="#-تضمین-کیفیت-و-تست‌ها-testing--qa"><img src="https://img.shields.io/badge/Tests-208%20Passing-success?style=flat-square" alt="Tests" /></a>
  <a href="https://react.dev/"><img src="https://img.shields.io/badge/React-19.3-61DAFB?style=flat-square&logo=react&logoColor=black" alt="React 19" /></a>
  <a href="https://www.typescriptlang.org/"><img src="https://img.shields.io/badge/TypeScript-5.0-3178C6?style=flat-square&logo=typescript&logoColor=white" alt="TypeScript" /></a>
  <a href="https://www.electronjs.org/"><img src="https://img.shields.io/badge/Electron-44.5-47848F?style=flat-square&logo=electron&logoColor=white" alt="Electron 44" /></a>
  <a href="https://tailwindcss.com/"><img src="https://img.shields.io/badge/Tailwind_CSS-v4-38B2AC?style=flat-square&logo=tailwind-css&logoColor=white" alt="Tailwind CSS 4" /></a>
  <a href="https://vitejs.dev/"><img src="https://img.shields.io/badge/Vite-8.3-646CFF?style=flat-square&logo=vite&logoColor=white" alt="Vite 8" /></a>
  <img src="https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-lightgrey?style=flat-square" alt="Platform" />
  <img src="https://img.shields.io/badge/License-Proprietary%20%28Polygharb%29-amber?style=flat-square" alt="License" />
</p>

---

## 📑 فهرست مطالب (Table of Contents)
- [📥 دانلود نرم‌افزار (Downloads)](#-دانلود-نرم‌افزار-downloads)
- [🌟 امکانات و قابلیت‌ها (Key Features)](#-امکانات-و-قابلیت‌ها-key-features)
- [⌨️ کلیدهای میانبر دسکتاپ (Keyboard Shortcuts)](#️-کلیدهای-میانبر-دسکتاپ-keyboard-shortcuts)
- [💻 راهنمای نصب و راه‌اندازی (Installation Guide)](#-راهنمای-نصب-و-راه‌اندازی-installation-guide)
- [🔒 اعتبارسنجی و هش فایل‌ها (Checksums)](#-اعتبارسنجی-و-هش-فایل‌ها-checksums)
- [🛠️ راهنمای ساخت و توسعه (Development & Build)](#️-راهنمای-ساخت-و-توسعه-development--build)
- [🧪 تضمین کیفیت و تست‌ها (Testing & QA)](#-تضمین-کیفیت-و-تست‌ها-testing--qa)
- [📁 ساختار پروژه (Project Architecture)](#-ساختار-پروژه-project-architecture)
- [📄 لایسنس و حقوق تجاری (License)](#-لایسنس-و-حقوق-تجاری-license)

---

## 📥 دانلود نرم‌افزار (Downloads)

آخرین نسخه‌های آماده نصب برای سیستم‌عامل‌های مکینتاش و ویندوز مستقیماً از طریق تب [GitHub Releases](https://github.com/ryanyv/invoice_app/releases/latest) در دسترس هستند:

| سیستم‌عامل (OS) | نوع فایل / معماری (Package) | حجم (Size) | لینک مستقیم دانلود (Direct Download) |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [دانلود Pipecal Pro 2.0.0 macOS DMG](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro-2.0.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer x64 `.exe` (نصبی) | ~108 MB | [دانلود Pipecal Pro Setup 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.Setup.2.0.0.exe) |
| 🪟 **Windows** | Portable x64 `.exe` (بدون نیاز به نصب) | ~108 MB | [دانلود Pipecal Pro 2.0.0.exe (Portable)](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.2.0.0.exe) |

> 💡 **یادداشت:** برای مشاهده تاریخچه تمامی نسخه‌ها، یادداشت‌های تغییرات و نسخه‌های پیشین، به [صفحه انتشارات پروژه (Releases Page)](https://github.com/ryanyv/invoice_app/releases) مراجعه کنید.

---

## 🌟 امکانات و قابلیت‌ها (Key Features)

### ۱. 🧮 محاسبات تخصصی مهندسی پلی‌اتیلن (PE Domain Engine)
- **انطباق کامل با استانداردهای DIN 8074 / DIN 8075 و INSO 14427:** جدول وزنی استاندارد یکپارچه برای تمام سایزها (از ۱۶ تا ۱۲۰۰ میلی‌متر)، گریدهای PE100, PE80, PE63 و تمامی رده‌های فشاری و SDR (از SDR 7.4 تا SDR 41).
- **محاسبات دوطرفه (مستقیم و معکوس):**
  - تبدیل لحظه‌ای متراژ به وزن کل و قیمت نهایی.
  - محاسبه معکوس متراژ از روی تناژ یا بودجه کل مشتری (مثلاً تعیین طول دقیق لوله بر حسب ۲ تن ظرفیت بارگیری).
- **سیستم هوشمند بسته‌بندی (کلاف در برابر شاخه):** تشخیص خودکار سایز لوله؛ کلاف ۱۰۰ یا ۲۰۰ متری برای سایزهای $\le 125\text{ mm}$، و شاخه‌های صلب ۶ متری یا ۱۲ متری برای سایزهای $> 125\text{ mm}$ به همراه اعلام متراژ مازاد و پرتی.
- **تخمین‌گر نوع ناوگان حمل‌ونقل بار:** ارزیابی تناژ وزنی و طول شاخه جهت پیشنهاد هوشمند نوع خودرو (وانت نیسان، خاور، کامیون تک ۶ چرخ، کامیون جفت ۱۰ چرخ، و تریلی کفی).
- **فرمول قیمت‌گذاری مواد اولیه بورس کالا (IME):** محاسبه قیمت پایه بر مبنای نرخ پلیمر بورس، درصد مستربچ، درصد ضایعات، هزینه‌های سربار و حاشیه سود کارخانه.

### ۲. 🧾 موتور پیشرفته صدور پیش‌فاکتور و گردش کار (Invoicing Workflow)
- **ویرایش درون‌خطی و زنده (Inline Editing):** ویرایش مستقیم سایز، طول، قیمت لوله‌ها و تعداد و نرخ اتصالات در خود جدول پیش‌فاکتور بدون نیاز به باز کردن دیالوگ یا فرم مجدد.
- **آرشیو و بازبینی سوابق فاکتورها:**
  - امکان ویرایش درجا (In-Place Edit) فاکتورهای قدیمی و ذخیره با حفظ شناسه.
  - کپی و تکثیر سریع (Duplicate) فاکتور با تولید شماره جدید جهت سرعت عمل بالا.
- **دفترچه مخاطبان و مشتریان (CRM Lite):** ثبت مشخصات خریداران، کد ملی/شناسه اقتصادی، کد پستی، شماره تماس و درصد تخفیف پیش‌فرض با قابلیت تکمیل خودکار (Autofill).
- **تضمین ثبات مالی ارزی:** امکان صدور و نگهداری مستقل به دو واحد **تومان** و **ریال** بدون ایجاد خطای گردکردن در پیش‌فاکتورهای گذشته.
- **تبدیل هوشمند به حروف فارسی:** درج خودکار مبلغ کل پرداختی به حروف شیوای فارسی در تمام پیش‌نمایش‌ها و برگه‌های چاپی.

### ۳. 🖨️ چاپ رسمی و خروجی‌های چندگانه (Official Printing & Export)
- **فاکتور رسمی ماده ۱۶۹ در برابر پیش‌فاکتور ساده:** سوئیچ با یک کلیک بین پیش‌فاکتور ساده تجاری و فاکتور رسمی مالیاتی استاندارد دارایی (شامل کد اقتصادی، شناسه ملی، کد پستی و شماره شبا بانکی).
- **تغییر جهت برگه (افقی / عمودی):** امکان انتخاب حالت چاپ A4 Landscape (افقی) یا A4 Portrait (عمودی) متناسب با حجم اقلام، با استایل‌بندی دقیق `@page` برای عدم شکست جدول.
- **خروجی PDF مستقیم:** تولید مستقیم فایل PDF استاندارد از محیط برنامه دسکتاپ.
- **خروجی اکسل (CSV) با استاندارد UTF-8 BOM:** همگام با نرم‌افزار Microsoft Excel در ویندوز و مک بدون به‌هم‌ریختگی حروف فارسی.
- **اشتراک‌گذاری سریع واتساپ و پیامک:** کپی متن خلاصه پیش‌فاکتور به صورت مرتب و آماده ارسال در پیام‌رسان‌ها با یک کلیک.

### ۴. 🛡️ پایداری آفلاین و امنیت اطلاعات (Offline Reliability)
- **استقلال کامل از اینترنت:** پکیج کامل فونت فارسی وزیرمتن (Vazirmatn) بدون نیاز به اینترنت یا CDN.
- **پشتیبان‌گیری چرخشی خودکار روی دیسک (Rolling Backups):** تهیه خودکار نسخه‌های پشتیبان در مسیر دایرکتوری سیستم‌عامل (`userData/backups/`) با نگهداری ۲۰ نسخه اخیر و امکان بازیابی آسان.
- **ذخیره خودکار پیش‌نویس (Draft Auto-Save):** جلوگیری کامل از اتلاف زمان و داده‌ها در صورت قطعی ناگهانی برق یا بسته‌شدن پنجره.

---

## ⌨️ کلیدهای میانبر دسکتاپ (Keyboard Shortcuts)

| کلید میانبر (macOS) | کلید میانبر (Windows) | عملکرد (Action) |
| :---: | :---: | :--- |
| <kbd>Cmd</kbd> + <kbd>S</kbd> | <kbd>Ctrl</kbd> + <kbd>S</kbd> | ذخیره پیش‌فاکتور جاری (Save Invoice) |
| <kbd>Cmd</kbd> + <kbd>P</kbd> | <kbd>Ctrl</kbd> + <kbd>P</kbd> | ورود به پیش‌نمایش و چاپ / خروجی PDF |
| <kbd>Cmd</kbd> + <kbd>N</kbd> | <kbd>Ctrl</kbd> + <kbd>N</kbd> | ایجاد پیش‌فاکتور جدید و پاکسازی فرم |
| <kbd>Esc</kbd> | <kbd>Esc</kbd> | بستن پنجره‌های باز و مودال‌ها |

---

## 💻 راهنمای نصب و راه‌اندازی (Installation Guide)

### 🍏 سیستم‌عامل مکینتاش (macOS)
1. فایل `Pipecal Pro-2.0.0-arm64.dmg` را از جدول دانلود دریافت کنید.
2. روی فایل دوبار کلیک کرده و آیکون **Pipecal Pro** را به پوشه **Applications** بکشید.
3. **نکته امنیتی Gatekeeper مک:** از آنجایی که برنامه به صورت سازمانی و بدون امضای توسعه‌دهنده تجاری اپل توزیع شده است، در اولین اجرا ممکن است پیام امنیتی دریافت کنید:
   - به منوی **System Settings** > **Privacy & Security** بروید.
   - به بخش **Security** اسکرول کنید و دکمه **Open Anyway** را کلیک نمایید.
   - یا در ترمینال دستور زیر را اجرا کنید:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

### 🪟 سیستم‌عامل ویندوز (Windows)
- **نسخه نصبی (Setup):** فایل `Pipecal Pro Setup 2.0.0.exe` را اجرا کنید تا مراحل نصب استاندارد طی شده و میانبر روی دسکتاپ قرار گیرد.
- **نسخه پرتابل (Portable):** فایل `Pipecal Pro 2.0.0.exe` را دانلود کرده و بدون نیاز به نصب یا دسترسی Administrator، مستقیماً اجرا نمایید.
- **نکته امنیتی Windows SmartScreen:** در صورت نمایش پنجره آبی‌رنگ محافظت ویندوز:
  - روی عبارت **More info** کلیک کنید.
  - دکمه **Run anyway** را بزنید.

---

## 🔒 اعتبارسنجی و هش فایل‌ها (Checksums)

برای اطمینان از سلامت و عدم تغییر فایل‌های دریافتی، می‌توانید هش‌های SHA-256 زیر را بررسی کنید:

| نام فایل (File) | الگوریتم SHA-256 |
| :--- | :--- |
| `Pipecal.Pro-2.0.0-arm64.dmg` | `342a4ba06a08894d0ccd1283496efd92880ebc1797fb820315111d63f353466d` |
| `Pipecal.Pro.Setup.2.0.0.exe` | `6e741f98fc55322d7ebdef7307a9b23ae7b4600b93ecded04e8eccd4424fb667` |
| `Pipecal.Pro.2.0.0.exe` | `707acdd96df944509ff832762a6e881f059a29c02b16e2b6eb6f0106cea6f7df` |

---

## 🛠️ راهنمای ساخت و توسعه (Development & Build)

### ۱. پیش‌نیازها
- **Node.js** نسخه 18 یا بالاتر
- **npm** یا **yarn**
- **Git**

### ۲. کلون و نصب وابستگی‌ها
```bash
git clone https://github.com/ryanyv/invoice_app.git
cd invoice_app
npm install
```

### ۳. اجرای محیط توسعه
```bash
# اجرای نسخه وب با سرور Vite:
npm run dev

# اجرای نسخه کامل دسکتاپ با Electron:
npm run electron:dev
```

### ۴. بیلد بسته‌های نصبی (Packaging)
```bash
# بیلد فایل DMG برای مک:
npm run electron:build:mac

# بیلد فایل‌های نصبی و پرتابل برای ویندوز:
npm run electron:build:win

# پکیج سریع بدون ایمیج نصبی:
npm run electron:pack
```

---

## 🧪 تضمین کیفیت و تست‌ها (Testing & QA)

پروژه مجهز به یک مجموعه جامع آزمون‌های خودکار و چالش‌های Adversarial با فریم‌ورک **Vitest** است:
- بیش از **۲۰۸ آزمون واحد و یکپارچه** با درصد پوشش بالا.
- صحه‌گذاری کامل فرمول‌های استاندارد DIN، محاسبات معکوس، حمل‌ونقل و الگوریتم‌های مالی.

برای اجرای آزمون‌ها:
```bash
npm test
```

---

## 📁 ساختار پروژه (Project Architecture)

```text
├── electron/
│   ├── main.cjs               # پروسه اصلی الکترون، مدیریت IPC، پنجره‌ها و چاپ
│   └── preload.cjs            # پل امن دسترسی پروسه رندر به دیسک و سیستم‌عامل
├── public/
│   ├── fonts/                 # فونت‌های محلی آفلاین وزیرمتن (Vazirmatn)
│   └── logo.png               # نشان‌واره رسمی با کیفیت عالی
├── src/
│   ├── components/
│   │   ├── CustomerManager/   # مدیریت مشتریان و دفترچه تلفن (CRM Lite)
│   │   ├── InvoiceBuilder/    # ویرایشگر فرم صدور پیش‌فاکتور و جدول اقلام
│   │   ├── InvoiceHistory/    # سوابق و آرشیو پیش‌فاکتورها، ویرایش و تکثیر
│   │   ├── InvoicePrint/      # پیش‌نمایش چاپ رسمی A4 افقی/عمودی و خروجی‌ها
│   │   ├── PriceManager/      # جدول مدیریت قیمت ۲۵۰+ اتصال و نرخ لوله‌ها
│   │   └── Settings/          # اطلاعات رسمی شرکت، مالیات و پشتیبان‌گیری
│   ├── services/
│   │   ├── calculator.ts      # موتور محاسبات DIN، بسته‌بندی، لجستیک و مواد
│   │   ├── persianUtils.ts    # توابع تاریخ جلالی و تبدیل اعداد و حروف
│   │   └── storage.ts         # مدیریت ذخیره‌سازی، کش و نسخه پشتیبان دیسک
│   ├── types/                 # تعاریف مدل‌های داده‌ای TypeScript
│   └── index.css              # استایل‌های سراسری با Tailwind CSS v4
├── release/                   # بسته‌های کامپایل‌شده نصبی مک و ویندوز
├── package.json               # تنظیمات پروژه و وابستگی‌ها
├── vite.config.ts             # تنظیمات Vite
└── vitest.config.ts           # تنظیمات Vitest
```

---

## 📄 لایسنس و حقوق تجاری (License)

توسعه‌یافته به سفارش و برای مجموعه **پلی غرب (Polygharb)**.  
کلیه حقوق مادی و معنوی محفوظ است © Polygharb.
