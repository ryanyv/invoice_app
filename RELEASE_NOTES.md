# Pipecal Pro v2.0.0 — Official Desktop Release 🚀
### Enterprise Desktop Solution for Polyethylene Pipes & Fittings Pricing, Quotation & Logistics
### سیستم جامع صدور پیش‌فاکتور و مدیریت قیمت لوله و اتصالات پلی‌اتیلن

<p align="center">
  <b>🌐 Language / زبان:</b>
  <a href="#-english"><b>English</b></a> &nbsp;•&nbsp;
  <a href="#-فارسی"><b>فارسی</b></a>
</p>

---

## 🇬🇧 English

We are proud to announce the official release of **Pipecal Pro v2.0.0**, an enterprise-grade desktop application engineered specifically for polyethylene (PE) pipe and fitting manufacturers, distributors, and engineering contractors.

Pipecal Pro 2.0 delivers precision DIN-standard calculations, fluid Persian right-to-left invoicing with Jalali calendar integration, official tax invoice generation, automated transport logistics estimation, and bulletproof offline reliability.

---

### 📥 Download Packages

| Operating System | Package Type / Architecture | File Size | Direct Download Link |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [Pipecal.Pro-2.0.0-arm64.dmg](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro-2.0.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer x64 `.exe` (NSIS) | ~108 MB | [Pipecal.Pro.Setup.2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.Setup.2.0.0.exe) |
| 🪟 **Windows** | Standalone Portable x64 `.exe` | ~108 MB | [Pipecal.Pro.2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.2.0.0.exe) |

> **Note**: Both a full Windows installer (`Setup.exe`) and a standalone portable single-file executable are provided. The portable edition requires no installation or administrator rights.

---

### 🔒 File Integrity & Checksums (SHA-256)

For security and file verification, match the SHA-256 checksums below:

```text
342a4ba06a08894d0ccd1283496efd92880ebc1797fb820315111d63f353466d  Pipecal.Pro-2.0.0-arm64.dmg
6e741f98fc55322d7ebdef7307a9b23ae7b4600b93ecded04e8eccd4424fb667  Pipecal.Pro.Setup.2.0.0.exe
707acdd96df944509ff832762a6e881f059a29c02b16e2b6eb6f0106cea6f7df  Pipecal.Pro.2.0.0.exe
```

---

### 🌟 What's New in Version 2.0.0

#### 1. 🧮 Polyethylene Engineering Engine (PE Domain)
- **DIN 8074 / 8075 & INSO 14427 Standard Compliance**: Built-in standard weight lookup tables for nominal pipe diameters from 16 mm to 1200 mm across all standard SDRs (SDR 7.4 through SDR 41) and working pressures (PN 2.5 to PN 25) for PE100, PE80, and PE63.
- **Bidirectional Calculations**:
  - Forward: Compute pipe weights, unit rates, and total sums dynamically from length.
  - Reverse: Determine required length from client budget or target tonnage (e.g. calculating exact meters for a 2-tonne transport threshold).
- **Intelligent Packaging Detector**:
  - Automatically identifies packaging constraints: flexible coils (کلاف) for diameters $\le 125\text{ mm}$, and rigid branches (شاخه 6m or 12m) for diameters $> 125\text{ mm}$ with automated surplus/remainder length tracking.
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
- **Desktop Keyboard Shortcuts**:
  - `Cmd/Ctrl` + `S`: Save current invoice
  - `Cmd/Ctrl` + `P`: Open print preview & PDF export
  - `Cmd/Ctrl` + `N`: Create new blank draft
  - `Esc`: Close active modals & dialogs

---

### 💻 Installation Instructions

#### 🍏 macOS
1. Download `Pipecal.Pro-2.0.0-arm64.dmg`.
2. Double-click the `.dmg` file and drag **Pipecal Pro** into your **Applications** folder.
3. *If prompted by macOS Gatekeeper ("App is from an unidentified developer"):*
   - Open **System Settings** > **Privacy & Security**.
   - Scroll down to the **Security** section and click **Open Anyway**.
   - Alternatively, execute this in Terminal:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

#### 🪟 Windows
- **Setup Installer**: Run `Pipecal.Pro.Setup.2.0.0.exe` and follow the guided installer. A desktop shortcut will be generated.
- **Portable Version**: Download `Pipecal.Pro.2.0.0.exe` and run it directly without installation or administrator privileges.
- *If prompted by Windows SmartScreen ("Windows protected your PC"):*
  - Click **More info**.
  - Click **Run anyway**.

---

### 🛠️ System Requirements

- **macOS**: macOS 11.0 (Big Sur) or higher (Apple Silicon M1/M2/M3/M4 or Rosetta 2).
- **Windows**: Windows 10 or Windows 11 (64-bit).
- **RAM**: 2 GB minimum (4 GB recommended).
- **Disk Space**: ~350 MB free space.

---

### 🤝 Support & Rights
- **Developed for**: Polygharb (پلی غرب)
- **Technical Stack**: React 19, TypeScript, Electron 44, Vite, Tailwind CSS v4, Vitest.
- **All Rights Reserved © Polygharb**.

---
---

## 🇮🇷 فارسی

با افتخار انتشار رسمی **Pipecal Pro نسخه ۲.۰.۰** را اعلام می‌نماییم؛ نرم‌افزاری تخصصی و پیشرفته برای تولیدکنندگان، توزیع‌کنندگان، پیمانکاران و دفاتر فنی لوله و اتصالات پلی‌اتیلن.

پایپ‌کال پرو ۲.۰ با بهره‌گیری از استانداردهای مهندسی DIN و INSO، فرم‌های صدور پیش‌فاکتور بومی راست‌به‌چپ (RTL)، تاریخ‌های جلالی (شمسی)، قالب‌های رسمی مالیاتی دارایی (ماده ۱۶۹ م.م)، برآورد هوشمند ناوگان بارگیری و پایداری کامل بدون نیاز به اینترنت، بهره‌وری و دقت واحدهای فروش را به بالاترین سطح ممکن ارتقا می‌دهد.

---

### 📥 دریافت فایل‌های نصبی (Download Packages)

| سیستم‌عامل (OS) | نوع بسته / معماری (Package) | حجم (Size) | لینک مستقیم دانلود (Direct Download) |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [دانلود Pipecal Pro 2.0.0 macOS DMG](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro-2.0.0-arm64.dmg) |
| 🪟 **Windows** | فایل نصبی x64 `.exe` (NSIS Installer) | ~108 MB | [دانلود Pipecal Pro Setup 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.Setup.2.0.0.exe) |
| 🪟 **Windows** | نسخه پرتابل x64 `.exe` (بدون نیاز به نصب) | ~108 MB | [دانلود Pipecal Pro 2.0.0.exe (Portable)](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.2.0.0.exe) |

> **نکته برای کاربران ویندوز**: هر دو نسخه نصبی با میانبر دسکتاپ و نسخه پرتابل آماده اجرا (بدون نیاز به نصب یا دسترسی ادمین) در دسترس هستند.

---

### 🔒 اعتبارسنجی و کد هش فایل‌ها (SHA-256)

جهت بررسی اصالت و سلامت بسته‌های دانلودی، کدهای هش SHA-256 زیر را مقایسه فرمایید:

```text
342a4ba06a08894d0ccd1283496efd92880ebc1797fb820315111d63f353466d  Pipecal.Pro-2.0.0-arm64.dmg
6e741f98fc55322d7ebdef7307a9b23ae7b4600b93ecded04e8eccd4424fb667  Pipecal.Pro.Setup.2.0.0.exe
707acdd96df944509ff832762a6e881f059a29c02b16e2b6eb6f0106cea6f7df  Pipecal.Pro.2.0.0.exe
```

---

### 🌟 ویژگی‌ها و امکانات جدید نسخه ۲.۰.۰

#### ۱. 🧮 محاسبات تخصصی مهندسی پلی‌اتیلن (PE Domain Engine)
- **انطباق با استانداردهای DIN 8074 / DIN 8075 و INSO 14427:** جدول وزنی استاندارد یکپارچه برای تمام سایزها (از ۱۶ تا ۱۲۰۰ میلی‌متر)، گریدهای PE100, PE80, PE63 و تمامی رده‌های فشاری و SDR (از SDR 7.4 تا SDR 41).
- **محاسبات دوطرفه (مستقیم و معکوس):**
  - تبدیل لحظه‌ای متراژ به وزن کل و قیمت نهایی.
  - محاسبه معکوس متراژ از روی تناژ یا بودجه کل مشتری (مثلاً تعیین طول دقیق لوله بر حسب ۲ تن ظرفیت بارگیری).
- **سیستم هوشمند بسته‌بندی (کلاف در برابر شاخه):** تشخیص خودکار سایز لوله؛ کلاف ۱۰۰ یا ۲۰۰ متری برای سایزهای $\le 125\text{ mm}$، و شاخه‌های صلب ۶ متری یا ۱۲ متری برای سایزهای $> 125\text{ mm}$ به همراه اعلام متراژ مازاد و پرتی.
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
- **کلیدهای میانبر دسکتاپ:**
  - `Cmd/Ctrl` + `S`: ذخیره سریع پیش‌فاکتور
  - `Cmd/Ctrl` + `P`: پیش‌نمایش و چاپ / صدور PDF
  - `Cmd/Ctrl` + `N`: ایجاد فرم پیش‌فاکتور جدید
  - `Esc`: بستن پنجره‌های باز و مودال‌ها

---

### 💻 راهنمای نصب و راه‌اندازی

#### 🍏 سیستم‌عامل مکینتاش (macOS)
1. فایل `Pipecal.Pro-2.0.0-arm64.dmg` را دانلود نمایید.
2. روی فایل دوبار کلیک کرده و آیکون **Pipecal Pro** را به پوشه **Applications** منتقل کنید.
3. **پیام امنیتی Gatekeeper مک:** از آنجایی که برنامه به صورت سازمانی توزیع شده است، در اولین اجرا ممکن است پیام امنیتی دریافت کنید:
   - به منوی **System Settings** > **Privacy & Security** بروید.
   - در بخش **Security** دکمه **Open Anyway** را کلیک نمایید.
   - یا در ترمینال دستور زیر را اجرا کنید:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

#### 🪟 سیستم‌عامل ویندوز (Windows)
- **نسخه نصبی (Setup):** فایل `Pipecal.Pro.Setup.2.0.0.exe` را اجرا کنید تا مراحل نصب استاندارد طی شده و میانبر روی دسکتاپ قرار گیرد.
- **نسخه پرتابل (Portable):** فایل `Pipecal.Pro.2.0.0.exe` را دانلود کرده و بدون نیاز به نصب یا دسترسی Administrator، مستقیماً اجرا نمایید.
- **پیام امنیتی Windows SmartScreen:** در صورت نمایش پنجره محافظت ویندوز:
  - روی عبارت **More info** کلیک کنید.
  - دکمه **Run anyway** را بزنید.

---

### 🛠️ حداقل مشخصات سیستم

- **مکینتاش**: macOS 11.0 یا بالاتر (پردازنده‌های Apple Silicon M1/M2/M3/M4).
- **ویندوز**: ویندوز 10 یا ویندوز 11 (۶۴ بیتی).
- **حافظه رم (RAM)**: حداقل ۲ گیگابایت (۴ گیگابایت پیشنهادی).
- **فضای دیسک**: حدود ۳۵۰ مگابایت فضای خالی.

---

### 📄 پشتیبانی و حقوق تجاری

- **توسعه‌یافته به سفارش و برای مجموعه**: پلی غرب (Polygharb)
- **فناوری‌های اصلی**: React 19, TypeScript, Electron 44, Vite, Tailwind CSS v4, Vitest.
- **کلیه حقوق مادی و معنوی محفوظ است © Polygharb**.
