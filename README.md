<p align="center">
  <img src="public/logo.png" alt="Pipecal Pro Logo" width="160" />
</p>

# Pipecal Pro 2.2 🚀
### Enterprise Desktop System for Polyethylene Pipes & Fittings Pricing, Quotation, Accounting & Logistics
### سیستم جامع صدور پیش‌فاکتور، مدیریت قیمت، حسابداری و اسناد مالی لوله و اتصالات پلی‌اتیلن

<p align="center">
  <a href="https://github.com/ryanyv/invoice_app/releases/latest"><img src="https://img.shields.io/badge/Release-v2.2.0-0284c7?style=flat-square&logo=github" alt="Latest Release" /></a>
  <a href="https://react.dev/"><img src="https://img.shields.io/badge/React-19.3-61DAFB?style=flat-square&logo=react&logoColor=black" alt="React 19" /></a>
  <a href="https://www.typescriptlang.org/"><img src="https://img.shields.io/badge/TypeScript-5.0-3178C6?style=flat-square&logo=typescript&logoColor=white" alt="TypeScript" /></a>
  <a href="https://www.electronjs.org/"><img src="https://img.shields.io/badge/Electron-44.5-47848F?style=flat-square&logo=electron&logoColor=white" alt="Electron 44" /></a>
  <a href="https://tailwindcss.com/"><img src="https://img.shields.io/badge/Tailwind_CSS-v4-38B2AC?style=flat-square&logo=tailwind-css&logoColor=white" alt="Tailwind CSS 4" /></a>
  <a href="https://vitejs.dev/"><img src="https://img.shields.io/badge/Vite-8.3-646CFF?style=flat-square&logo=vite&logoColor=white" alt="Vite 8" /></a>
  <img src="https://img.shields.io/badge/Tests-224%20Passing-success?style=flat-square" alt="Tests" />
  <img src="https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-lightgrey?style=flat-square" alt="Platform" />
  <img src="https://img.shields.io/badge/License-Proprietary%20%28Polygharb%29-amber?style=flat-square" alt="License" />
</p>

<p align="center">
  <b>🌐 Language / زبان:</b>
  <a href="#-english"><b>English</b></a> &nbsp;•&nbsp;
  <a href="#-فارسی"><b>فارسی</b></a>
</p>

---

## 🇬🇧 English

### Table of Contents
- [Downloads](#-downloads)
- [Key Features](#-key-features)
- [Keyboard Shortcuts](#-keyboard-shortcuts)
- [Installation Guide](#-installation-guide)
- [File Integrity & Checksums](#-file-integrity--checksums)
- [Development & Building from Source](#-development--building-from-source)
- [Quality Assurance & Testing](#-quality-assurance--testing)
- [Project Architecture](#-project-architecture)
- [License & Commercial Rights](#-license--commercial-rights)

---

### 📥 Downloads

Ready-to-use desktop binaries for macOS and Windows (**Latest: v2.2.0**) are published and available on the [GitHub Releases page](https://github.com/ryanyv/invoice_app/releases/latest):

| Operating System | Package Type / Architecture | File Size | Direct Download (v2.2.0) |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [Download Pipecal Pro 2.2.0 macOS DMG](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro-2.2.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer x64 & ARM64 `.exe` (NSIS) | ~220 MB | [Download Pipecal Pro Setup 2.2.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro.Setup.2.2.0.exe) |
| 🪟 **Windows** | Standalone Portable x64 & ARM64 `.exe` | ~220 MB | [Download Pipecal Pro 2.2.0.exe (Portable)](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro.2.2.0.exe) |

> 💡 **Previous Releases Archive**: Older versions remain permanently available on the [GitHub Releases Archive](https://github.com/ryanyv/invoice_app/releases), including [**v2.1.0**](https://github.com/ryanyv/invoice_app/releases/tag/v2.1.0) and [**v2.0.0**](https://github.com/ryanyv/invoice_app/releases/tag/v2.0.0).

---

### 🌟 Key Features

#### 1. 🧮 Polyethylene Engineering Engine (PE Domain)
- **DIN 8074 / 8075 & INSO 14427 Standard Compliance**: Built-in standard weight lookup tables for nominal pipe diameters from 16 mm to 1200 mm across all standard SDRs (SDR 7.4 through SDR 41) and working pressures (PN 2.5 to PN 25) for PE100, PE80, and PE63 (with quick default configuration for PE100, SDR 26, 5m branch).
- **Bidirectional Calculations & Price per Meter**:
  - Forward: Compute pipe weights, live price per meter (`قیمت هر متر`), unit rates per kg, and total sums dynamically from length.
  - Reverse: Determine required length and price per meter from client budget or target tonnage (e.g. calculating exact meters for a 2-tonne transport threshold).
- **Intelligent Packaging Detector**:
  - Automatically identifies packaging constraints: flexible coils (کلاف) for diameters $\le 125\text{ mm}$, and rigid branches (شاخه 5m, 6m, or 12m) for all diameters with automated surplus/remainder length tracking.
- **Freight Fleet Estimator**: Recommends appropriate commercial transport vehicles (Nissan pickup, Khavar truck, 6-wheel rigid truck, 10-wheel heavy truck, and flatbed trailer) based on cumulative mass and branch length.
- **IME Polymer Pricing Formula (بورس کالا)**: Calculates raw material unit cost from Iran Mercantile Exchange polymer index, masterbatch ratios, production overhead, waste percentage, and manufacturer target margins.

#### 2. 🧾 Invoicing & Quotations Workflow
- **Live In-line Item Editing**: Update pipe diameters, SDRs, lengths, and fitting quantities/prices directly inside the invoice table with immediate real-time recalculation.
- **Historical Invoices Archive**:
  - In-place editing of historical invoices with revision integrity.
  - 1-click invoice duplication into a new draft with fresh sequential numbering.
- **CRM Lite Customer Directory**: Searchable client directory storing contact phone numbers, national economic IDs, postal codes, and client-specific discount profiles with instant form autofill.
- **Dual Currency Consistency**: Native support for both Toman (تومان) and Rial (ریال) without historical rounding drift.
- **Persian Number-to-Words**: Live conversion of total payable figures into fluent Persian words on all invoices and print templates.

#### 3. 🖨️ Official Tax Printing & Versatile Exporting
- **Official VAT Invoice vs Standard Quotation**: 1-click toggle between informal commercial quotations and official Iranian VAT tax invoices conforming to **Article 169 of the Direct Tax Code (ماده ۱۶۹ م.م)** with full seller/buyer economic IDs, postal codes, and bank IBAN.
- **Orientation Toggle (Landscape / Portrait)**: Switch smoothly between A4 Landscape (افقی) and A4 Portrait (عمودی) print layouts with print-safe CSS `@page` page-break rules.
- **Native PDF Export**: 1-click desktop PDF document generation via Electron IPC printing bridge.
- **Excel / CSV Export with UTF-8 BOM**: Instant spreadsheet export compatible with Microsoft Excel on Windows and macOS with Persian character preservation.
- **WhatsApp / SMS Instant Share**: Formatted plain-text summary generator ready for 1-click clipboard copy and sharing via messaging apps.

#### 4. 🛡️ Offline-First Reliability & Data Safety
- **100% Offline Font Bundle**: Packaged with local Vazirmatn variable font assets; completely independent of external CDNs or active internet connectivity.
- **Rolling Disk Backups**: Automated rolling JSON snapshots stored in `userData/backups/` with a 20-version retention policy and one-click restore.
- **Draft Auto-Save**: Prevents data loss during power disruptions by continuously persisting builder state.

---

### ⌨️ Keyboard Shortcuts

| macOS Shortcut | Windows Shortcut | Action |
| :---: | :---: | :--- |
| <kbd>Cmd</kbd> + <kbd>S</kbd> | <kbd>Ctrl</kbd> + <kbd>S</kbd> | Save current invoice |
| <kbd>Cmd</kbd> + <kbd>P</kbd> | <kbd>Ctrl</kbd> + <kbd>P</kbd> | Open print preview & PDF export |
| <kbd>Cmd</kbd> + <kbd>N</kbd> | <kbd>Ctrl</kbd> + <kbd>N</kbd> | Create new blank draft |
| <kbd>Esc</kbd> | <kbd>Esc</kbd> | Close active modals & dialogs |

---

### 💻 Installation Guide

#### 🍏 macOS
1. Download `Pipecal.Pro-2.2.0-arm64.dmg`.
2. Double-click the `.dmg` file and drag **Pipecal Pro** into your **Applications** folder.
3. *If prompted by macOS Gatekeeper ("App is from an unidentified developer"):*
   - Open **System Settings** > **Privacy & Security**.
   - Scroll down to the **Security** section and click **Open Anyway**.
   - Alternatively, execute this in Terminal:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

#### 🪟 Windows
- **Setup Installer**: Run `Pipecal.Pro.Setup.2.2.0.exe` and follow the guided installer. A desktop shortcut will be generated.
- **Portable Version**: Download `Pipecal.Pro.2.2.0.exe` and run it directly without installation or administrator privileges.
- *If prompted by Windows SmartScreen ("Windows protected your PC"):*
  - Click **More info**.
  - Click **Run anyway**.

---

### 🔒 File Integrity & Checksums

| File Name (v2.2.0) | SHA-256 Digest |
| :--- | :--- |
| `Pipecal.Pro-2.2.0-arm64.dmg` | `04cc133c835a13279435de2886459bbd19a2635d1f53e048017573f40fdc3d2f` |
| `Pipecal.Pro.Setup.2.2.0.exe` | `f9e206369cb1a398d6b93dee0dc1b7135b5f0ddc551bedf010fed6cd45c6d1da` |
| `Pipecal.Pro.2.2.0.exe` | `a685d0880b20fe4503532b1cf21a8f4e487080a912f2703c11badfc5effa196e` |

---

### 🛠️ Development & Building from Source

#### 1. Prerequisites
- **Node.js** (v18 or higher)
- **npm** or **yarn**
- **Git**

#### 2. Clone & Install
```bash
git clone https://github.com/ryanyv/invoice_app.git
cd invoice_app
npm install
```

#### 3. Development Mode
```bash
# Web preview mode via Vite:
npm run dev

# Full Electron desktop app development:
npm run electron:dev
```

#### 4. Packaging Binaries
```bash
# Package macOS DMG (Apple Silicon):
npm run electron:build:mac

# Package Windows NSIS Setup & Portable executable:
npm run electron:build:win

# Fast unpack directory build:
npm run electron:pack
```

---

### 🧪 Quality Assurance & Testing

The project includes an automated test suite powered by **Vitest**:
- Over **208 automated unit and adversarial stress tests**.
- Comprehensive verification of DIN weight tables, logistics estimations, and financial calculations.

Run the test suite:
```bash
npm test
```

---

### 📁 Project Architecture

```text
├── electron/
│   ├── main.cjs               # Main process: IPC bridge, window lifecycle, PDF print
│   └── preload.cjs            # Secure preload bridge exposing desktop APIs
├── public/
│   ├── fonts/                 # Offline bundled Vazirmatn variable fonts
│   └── logo.png               # High-resolution company brand asset
├── src/
│   ├── components/
│   │   ├── CustomerManager/   # Client CRM directory & auto-fill profile
│   │   ├── InvoiceBuilder/    # Invoice editor form & dynamic item table
│   │   ├── InvoiceHistory/    # Invoices archive, in-place edit, and duplicate
│   │   ├── InvoicePrint/      # Official A4 landscape/portrait print preview
│   │   ├── PriceManager/      # 250+ standard PE fittings & pipe price catalog
│   │   └── Settings/          # Company tax registration & backup manager
│   ├── services/
│   │   ├── calculator.ts      # DIN standards, packaging, logistics, & IME formula
│   │   ├── persianUtils.ts    # Jalali calendar, Persian digits & number-to-words
│   │   └── storage.ts         # Dual-layer storage (localStorage + disk backups)
│   ├── types/                 # TypeScript domain and entity types
│   └── index.css              # Global styles & Tailwind CSS v4 setup
├── release/                   # Packaged macOS DMG and Windows EXEs
├── package.json               # Dependencies and build scripts
├── vite.config.ts             # Vite configuration
└── vitest.config.ts           # Vitest configuration
```

---

### 📄 License & Commercial Rights

Developed for **Polygharb (پلی غرب)**.  
All rights reserved © Polygharb.

---
---

## 🇮🇷 فارسی

### فهرست مطالب
- [📥 دانلود نرم‌افزار](#-دانلود-نرم‌افزار)
- [🌟 امکانات و قابلیت‌ها](#-امکانات-و-قابلیت‌ها)
- [⌨️ کلیدهای میانبر دسکتاپ](#️-کلیدهای-میانبر-دسکتاپ)
- [💻 راهنمای نصب و راه‌اندازی](#-راهنمای-نصب-و-راه‌اندازی)
- [🔒 اعتبارسنجی و هش فایل‌ها](#-اعتبارسنجی-و-هش-فایل‌ها)
- [🛠️ راهنمای ساخت و توسعه](#️-راهنمای-ساخت-و-توسعه)
- [🧪 تضمین کیفیت و تست‌ها](#-تضمین-کیفیت-و-تست‌ها)
- [📁 ساختار پروژه](#-ساختار-پروژه)
- [📄 لایسنس و حقوق تجاری](#-لایسنس-و-حقوق-تجاری)

---

### 📥 دانلود نرم‌افزار

آخرین نسخه‌های آماده نصب (**نسخه ۲.۲.۰**) برای سیستم‌عامل‌های مکینتاش و ویندوز مستقیماً از طریق تب [GitHub Releases](https://github.com/ryanyv/invoice_app/releases/latest) در دسترس هستند:

| سیستم‌عامل (OS) | نوع فایل / معماری (Package) | حجم (Size) | لینک مستقیم دانلود (نسخه ۲.۲.۰) |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [دانلود Pipecal Pro 2.2.0 macOS DMG](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro-2.2.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer x64 & ARM64 `.exe` (نصبی) | ~220 MB | [دانلود Pipecal Pro Setup 2.2.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro.Setup.2.2.0.exe) |
| 🪟 **Windows** | Portable x64 & ARM64 `.exe` (بدون نیاز به نصب) | ~220 MB | [دانلود Pipecal Pro 2.2.0.exe (Portable)](https://github.com/ryanyv/invoice_app/releases/download/v2.2.0/Pipecal.Pro.2.2.0.exe) |

> 💡 **آرشیو نسخه‌های پیشین:** تمامی نسخه‌های قبلی نرم‌افزار از جمله [**نسخه ۲.۱.۰ (v2.1.0)**](https://github.com/ryanyv/invoice_app/releases/tag/v2.1.0) و [**نسخه ۲.۰.۰ (v2.0.0)**](https://github.com/ryanyv/invoice_app/releases/tag/v2.0.0) در [صفحه انتشارات پروژه (Releases Page)](https://github.com/ryanyv/invoice_app/releases) نگهداری می‌شوند و همواره قابل دریافت هستند.

---

### 🌟 امکانات و قابلیت‌ها

#### ۱. 🧮 محاسبات تخصصی مهندسی پلی‌اتیلن (PE Domain Engine)
- **انطباق کامل با استانداردهای DIN 8074 / DIN 8075 و INSO 14427:** جدول وزنی استاندارد یکپارچه برای تمام سایزها (از ۱۶ تا ۱۲۰۰ میلی‌متر)، گریدهای PE100, PE80, PE63 و تمامی رده‌های فشاری و SDR (با تنظیمات پیش‌فرض سریع PE100، رده SDR 26 و شاخه ۵ متری).
- **محاسبات دوطرفه (مستقیم و معکوس) و نمایش قیمت هر متر:**
  - تبدیل لحظه‌ای متراژ به وزن کل، **قیمت هر متر (`قیمت/متر`)** و قیمت نهایی در فرم، جدول و چاپ.
  - محاسبه معکوس متراژ و قیمت هر متر از روی تناژ یا بودجه کل مشتری (مثلاً تعیین طول دقیق لوله بر حسب ۲ تن ظرفیت بارگیری).
- **سیستم هوشمند بسته‌بندی (کلاف در برابر شاخه ۵، ۶ و ۱۲ متری):** تشخیص خودکار سایز لوله؛ کلاف ۵۰، ۱۰۰ یا ۲۰۰ متری برای سایزهای $\le 125\text{ mm}$، و شاخه‌های صلب **۵ متری، ۶ متری یا ۱۲ متری** به همراه اعلام متراژ مازاد و پرتی.
- **تخمین‌گر نوع ناوگان حمل‌ونقل بار:** ارزیابی تناژ وزنی و طول شاخه جهت پیشنهاد هوشمند نوع خودرو (وانت نیسان، خاور، کامیون تک ۶ چرخ، کامیون جفت ۱۰ چرخ، و تریلی کفی).
- **فرمول قیمت‌گذاری مواد اولیه بورس کالا (IME):** محاسبه قیمت پایه بر مبنای نرخ پلیمر بورس، درصد مستربچ، درصد ضایعات، هزینه‌های سربار و حاشیه سود کارخانه.

#### ۲. 🧾 موتور پیشرفته صدور پیش‌فاکتور و گردش کار (Invoicing Workflow)
- **ویرایش درون‌خطی و زنده (Inline Editing):** ویرایش مستقیم سایز، طول، قیمت لوله‌ها و تعداد و نرخ اتصالات در خود جدول پیش‌فاکتور بدون نیاز به باز کردن دیالوگ یا فرم مجدد.
- **آرشیو و بازبینی سوابق فاکتورها:**
  - امکان ویرایش درجا (In-Place Edit) فاکتورهای قدیمی و ذخیره با حفظ شناسه.
  - کپی و تکثیر سریع (Duplicate) فاکتور با تولید شماره جدید جهت سرعت عمل بالا.
- **دفترچه مخاطبان و مشتریان (CRM Lite):** ثبت مشخصات خریداران، کد ملی/شناسه اقتصادی، کد پستی، شماره تماس و درصد تخفیف پیش‌فرض با قابلیت تکمیل خودکار (Autofill).
- **تضمین ثبات مالی ارزی:** امکان صدور و نگهداری مستقل به دو واحد **تومان** و **ریال** بدون ایجاد خطای گردکردن در پیش‌فاکتورهای گذشته.
- **تبدیل هوشمند به حروف فارسی:** درج خودکار مبلغ کل پرداختی به حروف شیوای فارسی در تمام پیش‌نمایش‌ها و برگه‌های چاپی.

#### ۳. 🖨️ چاپ رسمی و خروجی‌های چندگانه (Official Printing & Export)
- **فاکتور رسمی ماده ۱۶۹ در برابر پیش‌فاکتور ساده:** سوئیچ با یک کلیک بین پیش‌فاکتور ساده تجاری و فاکتور رسمی مالیاتی استاندارد دارایی (شامل کد اقتصادی، شناسه ملی، کد پستی و شماره شبا بانکی).
- **تغییر جهت برگه (افقی / عمودی):** امکان انتخاب حالت چاپ A4 Landscape (افقی) یا A4 Portrait (عمودی) متناسب با حجم اقلام، با استایل‌بندی دقیق `@page` برای عدم شکست جدول.
- **خروجی PDF مستقیم:** تولید مستقیم فایل PDF استاندارد از محیط برنامه دسکتاپ.
- **خروجی اکسل (CSV) با استاندارد UTF-8 BOM:** همگام با نرم‌افزار Microsoft Excel در ویندوز و مک بدون به‌هم‌ریختگی حروف فارسی.
- **اشتراک‌گذاری سریع واتساپ و پیامک:** کپی متن خلاصه پیش‌فاکتور به صورت مرتب و آماده ارسال در پیام‌رسان‌ها با یک کلیک.

#### ۴. 🛡️ پایداری آفلاین و امنیت اطلاعات (Offline Reliability)
- **استقلال کامل از اینترنت:** پکیج کامل فونت فارسی وزیرمتن (Vazirmatn) بدون نیاز به اینترنت یا CDN.
- **پشتیبان‌گیری چرخشی خودکار روی دیسک (Rolling Backups):** تهیه خودکار نسخه‌های پشتیبان در مسیر دایرکتوری سیستم‌عامل (`userData/backups/`) با نگهداری ۲۰ نسخه اخیر و امکان بازیابی آسان.
- **ذخیره خودکار پیش‌نویس (Draft Auto-Save):** جلوگیری کامل از اتلاف زمان و داده‌ها در صورت قطعی ناگهانی برق یا بسته‌شدن پنجره.

---

### ⌨️ کلیدهای میانبر دسکتاپ

| کلید میانبر (macOS) | کلید میانبر (Windows) | عملکرد (Action) |
| :---: | :---: | :--- |
| <kbd>Cmd</kbd> + <kbd>S</kbd> | <kbd>Ctrl</kbd> + <kbd>S</kbd> | ذخیره پیش‌فاکتور جاری (Save Invoice) |
| <kbd>Cmd</kbd> + <kbd>P</kbd> | <kbd>Ctrl</kbd> + <kbd>P</kbd> | ورود به پیش‌نمایش و چاپ / خروجی PDF |
| <kbd>Cmd</kbd> + <kbd>N</kbd> | <kbd>Ctrl</kbd> + <kbd>N</kbd> | ایجاد پیش‌فاکتور جدید و پاکسازی فرم |
| <kbd>Esc</kbd> | <kbd>Esc</kbd> | بستن پنجره‌های باز و مودال‌ها |

---

### 💻 راهنمای نصب و راه‌اندازی

#### 🍏 سیستم‌عامل مکینتاش (macOS)
1. فایل `Pipecal.Pro-2.2.0-arm64.dmg` را از جدول دانلود دریافت کنید.
2. روی فایل دوبار کلیک کرده و آیکون **Pipecal Pro** را به پوشه **Applications** بکشید.
3. **نکته امنیتی Gatekeeper مک:** از آنجایی که برنامه به صورت سازمانی و بدون امضای توسعه‌دهنده تجاری اپل توزیع شده است، در اولین اجرا ممکن است پیام امنیتی دریافت کنید:
   - به منوی **System Settings** > **Privacy & Security** بروید.
   - به بخش **Security** اسکرول کنید و دکمه **Open Anyway** را کلیک نمایید.
   - یا در ترمینال دستور زیر را اجرا کنید:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

#### 🪟 سیستم‌عامل ویندوز (Windows)
- **نسخه نصبی (Setup):** فایل `Pipecal.Pro.Setup.2.2.0.exe` را اجرا کنید تا مراحل نصب استاندارد طی شده و میانبر روی دسکتاپ قرار گیرد.
- **نسخه پرتابل (Portable):** فایل `Pipecal.Pro.2.2.0.exe` را دانلود کرده و بدون نیاز به نصب یا دسترسی Administrator، مستقیماً اجرا نمایید.
- **نکته امنیتی Windows SmartScreen:** در صورت نمایش پنجره آبی‌رنگ محافظت ویندوز:
  - روی عبارت **More info** کلیک کنید.
  - دکمه **Run anyway** را بزنید.

---

### 🔒 اعتبارسنجی و هش فایل‌ها

برای اطمینان از سلامت و عدم تغییر فایل‌های دریافتی، می‌توانید هش‌های SHA-256 زیر را بررسی کنید:

| نام فایل (نسخه ۲.۲.۰) | الگوریتم SHA-256 |
| :--- | :--- |
| `Pipecal.Pro-2.2.0-arm64.dmg` | `04cc133c835a13279435de2886459bbd19a2635d1f53e048017573f40fdc3d2f` |
| `Pipecal.Pro.Setup.2.2.0.exe` | `f9e206369cb1a398d6b93dee0dc1b7135b5f0ddc551bedf010fed6cd45c6d1da` |
| `Pipecal.Pro.2.2.0.exe` | `a685d0880b20fe4503532b1cf21a8f4e487080a912f2703c11badfc5effa196e` |

---

### 🛠️ راهنمای ساخت و توسعه

#### ۱. پیش‌نیازها
- **Node.js** نسخه 18 یا بالاتر
- **npm** یا **yarn**
- **Git**

#### ۲. کلون و نصب وابستگی‌ها
```bash
git clone https://github.com/ryanyv/invoice_app.git
cd invoice_app
npm install
```

#### ۳. اجرای محیط توسعه
```bash
# اجرای نسخه وب با سرور Vite:
npm run dev

# اجرای نسخه کامل دسکتاپ با Electron:
npm run electron:dev
```

#### ۴. بیلد بسته‌های نصبی (Packaging)
```bash
# بیلد فایل DMG برای مک:
npm run electron:build:mac

# بیلد فایل‌های نصبی و پرتابل برای ویندوز:
npm run electron:build:win

# پکیج سریع بدون ایمیج نصبی:
npm run electron:pack
```

---

### 🧪 تضمین کیفیت و تست‌ها

پروژه مجهز به یک مجموعه جامع آزمون‌های خودکار و چالش‌های Adversarial با فریم‌ورک **Vitest** است:
- بیش از **۲۰۸ آزمون واحد و یکپارچه** با درصد پوشش بالا.
- صحه‌گذاری کامل فرمول‌های استاندارد DIN، محاسبات معکوس، حمل‌ونقل و الگوریتم‌های مالی.

برای اجرای آزمون‌ها:
```bash
npm test
```

---

### 📁 ساختار پروژه

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

### 📄 لایسنس و حقوق تجاری

توسعه‌یافته به سفارش و برای مجموعه **پلی غرب (Polygharb)**.  
کلیه حقوق مادی و معنوی محفوظ است © Polygharb.
