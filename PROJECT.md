# Project: Pipecal Pro 2.0 Upgrades

## Architecture
- **Framework**: React 19 + TypeScript + Vite + Tailwind CSS v4.
- **Desktop Runtime**: Electron 44 (Main process in `electron/main.cjs`, Preload in `electron/preload.cjs`).
- **Data & State Management**: Lightweight synchronous state synchronized with `localStorage` via `StorageService` and dual-layer IPC disk backups in `app.getPath('userData')/backups/`.
- **Domain Services**:
  - `src/services/calculator.ts`: DIN 8074/8075 weight lookups, forward/reverse pipe calculations, packaging (coil vs branch), freight estimator, raw material IME pricing formula, invoice financial summary.
  - `src/services/persianUtils.ts`: Persian/English digit sanitization, Jalali date formatting, comma separation, Persian number-to-words.
  - `src/services/storage.ts`: Persistence layer for settings, catalogs, customers, draft invoices, and history.
- **Git & Branching Architecture**:
  - Base: `main` (untouched, read-only).
  - Feature branches:
    1. `feat/invoicing-workflow` (R1)
    2. `feat/pe-domain-calculator` (R2)
    3. `feat/print-export-sharing` (R3)
    4. `feat/offline-reliability` (R4)
    5. `feat/test-suite` (R5)
  - Integration branch: `staging` (receives merges from all feature branches).
  - Packaging: Triggered on `staging` for macOS DMG and Windows NSIS/portable x64.

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Edit Past Invoices | Load past invoice from history into builder for editing and re-saving | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 2 | Duplicate Past Invoices | Load past invoice as new draft with new invoice number | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 3 | Inline Items Table Editing | Edit pipe length/price and fitting quantity/price directly in table | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 4 | Customer Autofill | Autofill phone, address, and default discount on customer datalist selection | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 5 | Draft Auto-Save | Persist active builder draft in real-time to localStorage | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 6 | Currency Consistency | Store active currency (`تومان`/`ریال`) with invoice; preserve past amounts | M1 (`feat/invoicing-workflow`) | ORIGINAL_REQUEST §R1 |
| 7 | Reverse Pipe Calculation | Calculate required length from mass/tonnage, and unit price per kg from total | M2 (`feat/pe-domain-calculator`) | ORIGINAL_REQUEST §R2 |
| 8 | Freight Vehicle Estimator | Recommend Iranian freight vehicles (Nissan, Khavar, 6W, 10W, Flatbed) | M2 (`feat/pe-domain-calculator`) | ORIGINAL_REQUEST §R2 |
| 9 | Packaging Mode Indicator | Coil (کلاف) for ≤125mm vs 6m/12m branch (شاخه) for >125mm with count | M2 (`feat/pe-domain-calculator`) | ORIGINAL_REQUEST §R2 |
| 10 | Raw Material Cost Calculator | Calculate pricing from IME polymer base, masterbatch, waste, overhead, margin | M2 (`feat/pe-domain-calculator`) | ORIGINAL_REQUEST §R2 |
| 11 | Portrait/Landscape Print Toggle | Toggle A4 Landscape vs A4 Portrait in print preview with dynamic @page | M3 (`feat/print-export-sharing`) | ORIGINAL_REQUEST §R3 |
| 12 | Excel/CSV Invoice Export | Export invoice table and financial summary to CSV with UTF-8 BOM | M3 (`feat/print-export-sharing`) | ORIGINAL_REQUEST §R3 |
| 13 | WhatsApp/SMS Summary Copy | 1-click clipboard summary copy formatted for Iranian messaging apps | M3 (`feat/print-export-sharing`) | ORIGINAL_REQUEST §R3 |
| 14 | Official vs Simple Invoice | Toggle Iranian official tax invoice details (ماده ۱۶۹ م.م) and bank IBAN | M3 (`feat/print-export-sharing`) | ORIGINAL_REQUEST §R3 |
| 15 | Local Vazirmatn Font Bundle | 100% offline Persian typography; eliminate Google Fonts CDN links | M4 (`feat/offline-reliability`) | ORIGINAL_REQUEST §R4 |
| 16 | Desktop Keyboard Shortcuts | Support Ctrl/Cmd+S (save), Ctrl/Cmd+P (print), Ctrl/Cmd+N (new), Esc (exit) | M4 (`feat/offline-reliability`) | ORIGINAL_REQUEST §R4 |
| 17 | Rolling Backups & Safe Storage | Electron IPC disk backup snapshots to userData/backups with auto-retention | M4 (`feat/offline-reliability`) | ORIGINAL_REQUEST §R4 |
| 18 | Vitest Automation Suite | Install Vitest, configure `npm test`, test DIN math, Persian utils, summaries | M5 (`feat/test-suite`) | ORIGINAL_REQUEST §R5 |
| 19 | Git Integration to Staging | Merge all 5 feature branches cleanly into `staging` branch (do not touch main) | M6 (`feat/integration-packaging`) | ORIGINAL_REQUEST §R6 |
| 20 | Cross-Platform Desktop Packaging | Build macOS DMG and Windows x64 binaries into `release/` | M6 (`feat/integration-packaging`) | ORIGINAL_REQUEST §R6 |

## Milestones
| # | Name | Branch | Scope | Dependencies | Status |
|---|------|--------|-------|--------------|--------|
| M1 | Invoicing Workflow Enhancements | `feat/invoicing-workflow` | R1: Edit/duplicate past invoices, inline item editing, customer autofill, draft auto-save, currency consistency | none | DONE (commit: 19e949d, 42 tests passing) |
| M2 | PE Domain Capabilities | `feat/pe-domain-calculator` | R2: Reverse calculations, freight vehicle estimator, packaging rules, raw material cost formula | none | DONE (commits: dc2362c, 9566e3a, 79317ac, 914191b, 24bb7b0; 36k+ tests passed) |
| M3 | Print, Export & Sharing | `feat/print-export-sharing` | R3: Landscape/portrait print toggle, CSV export, WhatsApp copy, official tax invoice layout | none | DONE (commits: edd6c58, 2fdd0ab, 8e2962f, df97299; 6/6 test suites passed) |
| M4 | Offline & Desktop Reliability | `feat/offline-reliability` | R4: Local Vazirmatn font bundle, keyboard shortcuts, electron rolling disk backups | none | DONE (commits: abd10d8, 2e13afd, 1d7fb5d, 2852d4f, 43014e8, dcf0454; 73+ tests passed) |
| M5 | Test Automation Suite | `feat/test-suite` | R5: Vitest setup, 34+ test cases for DIN weights, Persian utils, calculations | none | DONE (commits: 948f270, 6f776cd, 42d89b4, d764825; 80 tests passing in ~200ms) |
| M6 | Integration & Packaging | `staging` | R6: Merge all branches into `staging`, verify `npm test`, build macOS & Windows in `release/` | M1, M2, M3, M4, M5 | DONE (all 5 branches merged into staging, 198 tests passing, macOS DMG & Windows EXE generated in release/) |

## Interface Contracts
### Invoice & Storage Contract (`src/types/index.ts`, `src/services/storage.ts`)
- `Invoice.currencyUnit`: `'تومان' | 'ریال'` (mandatory on new invoices; default to `'تومان'` on legacy).
- `Invoice.isOfficial`: `boolean` (optional).
- `StorageService.getDraft()`: returns `Partial<Invoice> | null`.
- `StorageService.saveDraft(draft: Partial<Invoice>)`: void.
- `StorageService.clearDraft()`: void.

### Domain Calculations Contract (`src/services/calculator.ts`)
- `isCoilAllowed(diameter: number): boolean`
- `computePackaging(diameter: number, length: number, options?: { preferredType?: 'coil' | 'branch'; branchLength?: 6 | 12; coilLength?: number; }): PackagingResult`
- `estimateFreightVehicle(totalWeightKg: number, maxBranchLengthMeters?: number): FreightRecommendation`
- `calculateRawMaterialCost(params: RawMaterialCostParams): RawMaterialCostResult`
- `calculateLengthFromMass(totalMass: number, diameter: number, sdr: number, weightTable: WeightTable): number | null`
- `calculatePricePerKgFromTotal(totalPrice: number, totalWeight: number): number | null`

### Persian Utils Contract (`src/services/persianUtils.ts`)
- `toEnglishDigits(input: string | number): string`
- `parseFormattedNumber(input: string | number): number`
- `formatNumber(num: number | string, usePersianDigits?: boolean, decimals?: number): string`

### Electron IPC Contract (`electron/main.cjs`, `electron/preload.cjs`)
- `save-pdf`: `(event, { defaultName, landscape }) => Promise<{ success: boolean; filePath?: string }>`
- `save-backup`: `(event, { backupJson }) => Promise<{ success: boolean; filePath?: string }>`
- `list-backups`: `() => Promise<{ files: string[] }>`
- `load-backup`: `(event, { filename }) => Promise<{ content: string }>`

## Code Layout
- `src/types/index.ts`: Shared domain & entity types.
- `src/services/storage.ts`: Persistence and backup layer.
- `src/services/calculator.ts`: Math and domain rules.
- `src/services/persianUtils.ts`: Persian text/digits/dates.
- `src/components/InvoiceBuilder/`: Invoice creation, inline editing, domain calculators.
- `src/components/InvoiceHistory/`: Archive, edit, duplicate, delete.
- `src/components/InvoicePrint/`: Print preview, orientation toggle, official layout, CSV & WhatsApp export.
- `src/components/Settings/`: Settings modal, official company details, backup manager.
- `electron/main.cjs` & `preload.cjs`: Electron processes and IPC handlers.
- `public/fonts/`: Local Vazirmatn font assets.
- `src/services/__tests__/`: Vitest test suites.
- `release/`: Electron builder packaged binaries.
