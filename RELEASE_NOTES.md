# Pipecal Pro v2.0.0 — Official Release 🚀
### سیستم جامع صدور پیش‌فاکتور و مدیریت قیمت لوله و اتصالات پلی‌اتیلن
**Polyethylene Pipes & Fittings Invoicing & Pricing Desktop System**

---

We are thrilled to announce the official release of **Pipecal Pro v2.0.0**, an enterprise-grade desktop application engineered specifically for polyethylene pipe and fitting manufacturers, distributors, and engineering contractors.

Pipecal Pro 2.0 delivers precision engineering calculations, seamless Jalali (شمسی) invoicing, official tax quotation generation, real-time logistics estimation, and bulletproof offline reliability.

---

## 📥 Download Packages (فایل‌های نصبی آماده)

| سیستم‌عامل (OS) | معماری / نوع (Package) | حجم (Size) | لینک دانلود مستقیم (Direct Download) |
| :--- | :--- | :--- | :--- |
| 🍏 **macOS** | Apple Silicon (M1/M2/M3/M4) `.dmg` | ~126 MB | [Pipecal Pro-2.0.0-arm64.dmg](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro-2.0.0-arm64.dmg) |
| 🪟 **Windows** | Setup Installer x64 `.exe` (NSIS) | ~108 MB | [Pipecal Pro Setup 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.Setup.2.0.0.exe) |
| 🪟 **Windows** | Standalone Portable x64 `.exe` | ~108 MB | [Pipecal Pro 2.0.0.exe](https://github.com/ryanyv/invoice_app/releases/download/v2.0.0/Pipecal.Pro.2.0.0.exe) |

> **Note on Windows Downloads**: Both installer (`Setup.exe`) and portable single-file executable options are provided. The portable edition requires no installation or administrative permissions.

---

## 🔒 Verification & Checksums (SHA-256)

For security and file integrity verification, compare the SHA-256 checksums below:

```text
342a4ba06a08894d0ccd1283496efd92880ebc1797fb820315111d63f353466d  Pipecal Pro-2.0.0-arm64.dmg
6e741f98fc55322d7ebdef7307a9b23ae7b4600b93ecded04e8eccd4424fb667  Pipecal Pro Setup 2.0.0.exe
707acdd96df944509ff832762a6e881f059a29c02b16e2b6eb6f0106cea6f7df  Pipecal Pro 2.0.0.exe
```

---

## 🌟 What's New in Version 2.0.0 (ویژگی‌های جدید)

### 1. 🧮 PE Engineering Calculations (محاسبات تخصصی پلی‌اتیلن)
- **DIN 8074 / 8075 & INSO 14427 Standard Compliance**: Built-in standard weight lookup tables across all standard SDRs (SDR 7.4 through SDR 41) and nominal pressures (PN 2.5 to PN 25) for PE100, PE80, and PE63.
- **Bidirectional Calculations**:
  - Calculate total weight and pricing instantly from pipe length and diameter.
  - Reverse calculation: compute required pipe length from target batch mass/tonnage or target budget.
- **Intelligent Packaging Detector**:
  - Automatically identifies packaging constraints: flexible coils (کلاف) for diameters $\le 125\text{ mm}$, and rigid branches (شاخه ۶ یا ۱۲ متری) for diameters $> 125\text{ mm}$ with automated surplus/remainder tracking.
- **Freight & Logistics Estimator**:
  - Evaluates cumulative invoice weight and maximum pipe length to recommend appropriate freight transport (وانت نیسان، خاور، کامیون تک ۶ چرخ، جفت ۱۰ چرخ، تریلی کفی).
- **IME Polymer Base Pricing Formula (بورس کالا)**:
  - Calculate unit pricing derived from Iran Mercantile Exchange raw material base prices, masterbatch ratio, operational overhead, production waste, and profit margin.

### 2. 🧾 Invoicing & In-Line Workflow (صدور پیش‌فاکتور و گردش کار)
- **Interactive In-line Item Editing**: Modify pipe dimensions, lengths, fitting quantities, and unit prices directly within the line items table with zero latency recalculation.
- **Historical Invoices Archive**:
  - In-place editing of historical invoices with customer version preservation.
  - 1-click invoice duplication into a fresh draft with updated invoice numbering.
- **CRM Lite Customer Directory**:
  - Searchable client repository storing contact numbers, national IDs, addresses, and customer-specific discount profiles.
  - Auto-fills client information immediately upon selection.
- **Currency Consistency**: Preserves financial amounts in both Toman (تومان) and Rial (ریال) without historical rounding drift.
- **Persian Number-to-Words**: Live conversion of total payable amounts into Persian words on all invoices and printouts.

### 3. 🖨️ Official Printing & Multi-Channel Exporting (چاپ و خروجی‌ها)
- **Official Tax Invoice vs Simple Quotation**:
  - Conforms to **Article 169 of the Iranian Direct Tax Law (ماده ۱۶۹ م.م)** with full seller and buyer national ID, economic code, registration number, postal code, and bank IBAN.
- **Orientation Toggle (Landscape / Portrait)**:
  - Dynamically switches between A4 Landscape (افقی) and A4 Portrait (عمودی) print layouts with print-safe CSS `@page` styling.
- **Native PDF Export**: 1-click desktop PDF generation via Electron IPC printing bridge.
- **Excel / CSV Export with UTF-8 BOM**: Instant spreadsheet export compatible with Microsoft Excel on Windows and macOS with Persian character preservation.
- **WhatsApp / SMS Instant Share**: Formatted plain-text quotation generator ready for 1-click clipboard copy and sharing via messaging apps.

### 4. ⚡ Offline Reliability & Enterprise Performance (پایداری و اجرای آفلاین)
- **100% Offline Font Bundle**: Packaged with local Vazirmatn variable font assets; completely independent of Google Fonts or internet access.
- **Rolling Disk Backups**: Automated rolling JSON snapshots stored in `userData/backups/` with a 20-version retention policy and one-click restore.
- **Draft Auto-Save**: Prevents data loss during power disruptions by continuously persisting builder state.
- **Desktop Keyboard Shortcuts**:
  - `Ctrl` / `Cmd` + `S`: Save invoice
  - `Ctrl` / `Cmd` + `P`: Print / Export PDF
  - `Ctrl` / `Cmd` + `N`: New blank draft
  - `Esc`: Dismiss active modal dialogs

---

## 💻 Installation Instructions (راهنمای نصب و راه‌اندازی)

### 🍏 macOS Installation
1. Download `Pipecal Pro-2.0.0-arm64.dmg`.
2. Double-click the `.dmg` file and drag **Pipecal Pro** into your **Applications** folder.
3. *If prompted by macOS Gatekeeper ("App is from an unidentified developer"):*
   - Open **System Settings** > **Privacy & Security**.
   - Scroll down to the **Security** section and click **Open Anyway**.
   - Alternatively, open Terminal and run:
     ```bash
     xattr -cr /Applications/Pipecal\ Pro.app
     ```

### 🪟 Windows Installation
- **Setup Installer**: Download `Pipecal Pro Setup 2.0.0.exe` and follow the on-screen installation wizard. A desktop shortcut will be created automatically.
- **Portable Version**: Download `Pipecal Pro 2.0.0.exe` and place it anywhere (e.g., Desktop or USB drive). Double-click to run immediately without installation.
- *If prompted by Windows SmartScreen ("Windows protected your PC"):*
  - Click **More info**.
  - Click **Run anyway**.

---

## 🛠️ System Requirements (حداقل سیستم مورد نیاز)

- **macOS**: macOS 11.0 (Big Sur) or higher (Apple Silicon M1/M2/M3/M4 or Rosetta 2).
- **Windows**: Windows 10 or Windows 11 (64-bit).
- **RAM**: 2 GB minimum (4 GB recommended).
- **Disk Space**: ~350 MB free space.

---

## 🤝 Support & Acknowledgments

- **Developed for**: Polygharb (پلی غرب)
- **Technical Stack**: React 19, TypeScript, Electron 44, Vite, Tailwind CSS v4, Vitest.
- **All Rights Reserved © Polygharb**.

