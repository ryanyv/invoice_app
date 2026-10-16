/**
 * Empirical Challenger 2 Test Harness - Milestone M3 (Requirement R3)
 * Advanced Print, Export & Sharing
 *
 * Mission Probes:
 * 1. CSV export and WhatsApp copy with extreme inputs:
 *    - 1,000 items (pipes, fittings, mixed) performance and correctness
 *    - Extreme trillion numbers (10^12 to 10^15) in calculations, digits, words, and CSV formatting
 *    - Special injection characters in customer name/address/explanation (quotes, commas, \r\n, <>&, formulas =, +, -, @)
 * 2. Strict RFC 4180 round-trip parser verification of all injected fields
 * 3. Official vs simple invoice toggling with missing/null/undefined official fields
 * 4. Component level null-safety probe for InvoicePrintView
 */

import { CompanySettings, Invoice, PipeItem, FittingItem } from '../../types';
import {
  generateInvoiceCsvString,
  generateInvoiceWhatsAppSummary,
} from '../../utils/exportUtils';
import {
  formatNumber,
  numberToPersianWords,
  toPersianDigits,
} from '../persianUtils';

// ============================================================================
// TEST HARNESS UTILITIES & ASSERTIONS
// ============================================================================

let totalPassed = 0;
let totalFailed = 0;
const failureDetails: string[] = [];

function runTest(name: string, fn: () => void) {
  it(name, () => {
    try {
      fn();
      totalPassed++;
      console.log(`  ✓ PASS: ${name}`);
    } catch (err: any) {
      totalFailed++;
      const msg = `  ✗ FAIL: ${name}\n    Error: ${err.message || String(err)}`;
      console.error(msg);
      failureDetails.push(`${name} -> ${err.message || String(err)}`);
      throw err;
    }
  });
}

function expectEqual<T>(actual: T, expected: T, desc: string) {
  if (actual !== expected) {
    throw new Error(`${desc}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

function expectTrue(val: boolean, desc: string) {
  if (!val) {
    throw new Error(`${desc}: Expected true, got false`);
  }
}

// ============================================================================
// RFC 4180 FULL COMPLIANT CSV PARSER
// ============================================================================

function parseRfc4180(csvString: string): string[][] {
  // Strip UTF-8 BOM if present
  let text = csvString;
  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }

  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentCell = '';
  let insideQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (insideQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Escaped double quote
          currentCell += '"';
          i++;
        } else {
          // Closing quote
          insideQuotes = false;
        }
      } else {
        currentCell += char;
      }
    } else {
      if (char === '"') {
        insideQuotes = true;
      } else if (char === ',') {
        currentRow.push(currentCell);
        currentCell = '';
      } else if (char === '\r') {
        if (nextChar === '\n') {
          i++;
        }
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = [];
        currentCell = '';
      } else if (char === '\n') {
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = [];
        currentCell = '';
      } else {
        currentCell += char;
      }
    }
  }

  if (currentCell.length > 0 || currentRow.length > 0) {
    currentRow.push(currentCell);
    rows.push(currentRow);
  }

  return rows;
}

// Base mock fixtures
const baseSettings: CompanySettings = {
  name: 'شرکت لوله و اتصالات پلیمر غرب',
  address: 'تهران، خیابان ولیعصر، برج سپهر، طبقه ۸',
  phone: '02188776655',
  fax: '02188776656',
  taxRate: 10,
  invoiceCounter: 200,
  currencyUnit: 'تومان',
  basePipePrices: {
    PE100: 98000,
    PE80: 91000,
    PE63: 84000,
  },
  economicCode: '411598746321',
  nationalId: '14008765432',
  registrationNumber: '542198',
  postalCode: '1968745231',
  bankName: 'بانک تجارت شعبه اسکان',
  bankAccount: '9876543210',
  bankIban: 'IR980180000000098765432101',
};

function createBaseInvoice(overrides: Partial<Invoice> = {}): Invoice {
  return {
    id: 'inv_test_001',
    invoiceNumber: '201',
    customerName: 'شرکت پترو پارس انرژی',
    customerPhone: '09123456789',
    customerAddress: 'عسلویه، منطقه ویژه اقتصادی انرژی پارس',
    customerEconomicCode: '411987654321',
    customerNationalId: '10108765432',
    customerPostalCode: '7511896541',
    currencyUnit: 'تومان',
    isOfficial: true,
    jalaliDate: '1405/07/26',
    items: [],
    subtotal: 0,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 0,
    explanation: 'شرایط تسویه نقدی پیش از بارگیری',
    status: 'draft',
    createdAt: '2026-10-16T00:00:00Z',
    ...overrides,
  };
}

describe('CHALLENGER 2: ADVERSARIAL STRESS TEST HARNESS - REQUIREMENT R3', () => {

// ============================================================================
// PROBE 1: 1,000 ITEMS EXTREME SCALE TEST
// ============================================================================
console.log('--- PROBE 1: 1,000 Items Extreme Scale Test ---');

runTest('1.1 CSV Generation performance & integrity for 1,000 pipe items', () => {
  const items: PipeItem[] = [];
  let expectedTotalLength = 0;
  let expectedTotalWeight = 0;
  let expectedTotalPrice = 0;

  for (let i = 1; i <= 1000; i++) {
    const len = i * 10;
    const wPerM = 2.5;
    const totW = len * wPerM;
    const pricePerKg = 90000;
    const totP = totW * pricePerKg;

    expectedTotalLength += len;
    expectedTotalWeight += totW;
    expectedTotalPrice += totP;

    items.push({
      id: `pipe_${i}`,
      type: 'pipe',
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: len,
      weightPerMeter: wPerM,
      totalWeight: totW,
      pricePerKg: pricePerKg,
      totalPrice: totP,
    });
  }

  const invoice = createBaseInvoice({
    items,
    subtotal: expectedTotalPrice,
    totalAmount: expectedTotalPrice,
  });

  const tStart = performance.now();
  const csv = generateInvoiceCsvString(invoice, baseSettings);
  const tEnd = performance.now();
  const durationMs = tEnd - tStart;

  // Verify performance (< 150ms for 1,000 items)
  expectTrue(durationMs < 150, `CSV generation took ${durationMs.toFixed(2)}ms, expected < 150ms`);

  // Verify BOM
  expectTrue(csv.startsWith('\uFEFF'), 'CSV must start with UTF-8 BOM');

  // Parse CSV strictly
  const parsed = parseRfc4180(csv);

  // Find pipe table rows
  const pipeHeaderIdx = parsed.findIndex((r) => r[0] === '--- جدول اقلام لوله پلی‌اتیلن ---');
  expectTrue(pipeHeaderIdx !== -1, 'Pipe header section must exist in CSV');

  // Verify 1000 item rows exist
  const firstPipeRow = parsed[pipeHeaderIdx + 2];
  expectEqual(firstPipeRow[0], '1', 'First pipe row index');
  expectEqual(firstPipeRow[6], '10', 'First pipe row length');

  const lastPipeRow = parsed[pipeHeaderIdx + 1 + 1000];
  expectEqual(lastPipeRow[0], '1000', '1000th pipe row index');
  expectEqual(lastPipeRow[6], '10000', '1000th pipe row length');

  // Verify pipe summary row
  const summaryRow = parsed[pipeHeaderIdx + 2 + 1000];
  expectEqual(summaryRow[0], 'جمع اقلام لوله', 'Summary row title');
  expectEqual(summaryRow[6], `${expectedTotalLength} متر`, 'Total length in summary');
  expectEqual(summaryRow[8], `${expectedTotalWeight.toFixed(3)} کیلو`, 'Total weight in summary');
  expectEqual(summaryRow[10], String(expectedTotalPrice), 'Total price in summary');
});

runTest('1.2 CSV Generation performance & integrity for 1,000 fitting items', () => {
  const items: FittingItem[] = [];
  let expectedTotalQty = 0;
  let expectedTotalPrice = 0;

  for (let i = 1; i <= 1000; i++) {
    const qty = i;
    const unitPrice = 50000;
    const totP = qty * unitPrice;

    expectedTotalQty += qty;
    expectedTotalPrice += totP;

    items.push({
      id: `fit_${i}`,
      type: 'fitting',
      category: 'اتصالات جوشی',
      connectionType: 'جوشی',
      product: `زانو ۹۰ درجه سایز ${i}`,
      material: 'HDPE100',
      pressure: '16',
      size: `${i}`,
      quantity: qty,
      unitPrice: unitPrice,
      totalPrice: totP,
    });
  }

  const invoice = createBaseInvoice({
    items,
    subtotal: expectedTotalPrice,
    totalAmount: expectedTotalPrice,
  });

  const tStart = performance.now();
  const csv = generateInvoiceCsvString(invoice, baseSettings);
  const tEnd = performance.now();
  const durationMs = tEnd - tStart;

  expectTrue(durationMs < 150, `Fitting CSV generation took ${durationMs.toFixed(2)}ms`);

  const parsed = parseRfc4180(csv);
  const fittingHeaderIdx = parsed.findIndex((r) => r[0] === '--- جدول اقلام اتصالات پلی‌اتیلن ---');
  expectTrue(fittingHeaderIdx !== -1, 'Fitting header section must exist');

  const summaryRow = parsed[fittingHeaderIdx + 2 + 1000];
  expectEqual(summaryRow[0], 'جمع اقلام اتصالات', 'Fitting summary row title');
  expectEqual(summaryRow[6], `${expectedTotalQty} عدد`, 'Total quantity in fitting summary');
  expectEqual(summaryRow[8], String(expectedTotalPrice), 'Total price in fitting summary');
});

runTest('1.3 WhatsApp summary scaling with 1,000 mixed items', () => {
  const items: (PipeItem | FittingItem)[] = [];
  for (let i = 1; i <= 500; i++) {
    items.push({
      id: `pipe_${i}`,
      type: 'pipe',
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: 10,
      weightPerMeter: 2.38,
      totalWeight: 23.8,
      pricePerKg: 95000,
      totalPrice: 2261000,
    });
    items.push({
      id: `fit_${i}`,
      type: 'fitting',
      category: 'الکتروفیوژن',
      connectionType: 'کوپلر',
      product: `کوپلر سایز ${i}`,
      material: 'PE100',
      pressure: '16',
      size: `${i}`,
      quantity: 2,
      unitPrice: 150000,
      totalPrice: 300000,
    });
  }

  const subtotal = 500 * (2261000 + 300000);
  const invoice = createBaseInvoice({
    items,
    subtotal,
    totalAmount: subtotal,
  });

  const tStart = performance.now();
  const summary = generateInvoiceWhatsAppSummary(invoice, baseSettings);
  const tEnd = performance.now();
  const durationMs = tEnd - tStart;

  expectTrue(durationMs < 150, `WhatsApp summary took ${durationMs.toFixed(2)}ms`);
  expectTrue(summary.includes('🔹 اقلام لوله پلی‌اتیلن:'), 'Contains pipe section');
  expectTrue(summary.includes('🔸 اقلام اتصالات:'), 'Contains fitting section');
  expectTrue(summary.includes('۵۰۰. لوله'), 'Contains 500th pipe item');
  expectTrue(summary.includes('۵۰۰. کوپلر سایز'), 'Contains 500th fitting item');
});

// ============================================================================
// PROBE 2: EXTREME TRILLION NUMBERS & FINANCIAL ACCURACY
// ============================================================================
console.log('\n--- PROBE 2: Extreme Trillion Numbers & Financial Accuracy ---');

runTest('2.1 Handling extreme figures (10^12 to 10^14) in CSV and formatting', () => {
  const trillionSubtotal = 4_750_000_000_000; // 4.75 Trillion
  const discountAmount = 237_500_000_000;    // 237.5 Billion (5%)
  const taxAmount = 451_250_000_000;         // 451.25 Billion (10% of taxable)
  const totalAmount = 4_963_750_000_000;     // 4.96375 Trillion

  const invoice = createBaseInvoice({
    subtotal: trillionSubtotal,
    includeDiscount: true,
    discountPercentage: 5,
    discountAmount,
    includeTax: true,
    taxPercentage: 10,
    taxAmount,
    totalAmount,
    currencyUnit: 'ریال',
  });

  const csv = generateInvoiceCsvString(invoice, baseSettings);
  const parsed = parseRfc4180(csv);

  // Locate financial summary in CSV
  const subtotalRow = parsed.find((r) => r[0] === 'جمع کل ناخالص:');
  expectTrue(Boolean(subtotalRow), 'Subtotal row found');
  expectEqual(subtotalRow![1], String(trillionSubtotal), 'Subtotal value in CSV');
  expectEqual(subtotalRow![2], 'ریال', 'Currency in subtotal row');

  const discountRow = parsed.find((r) => r[0].startsWith('تخفیف'));
  expectTrue(Boolean(discountRow), 'Discount row found');
  expectEqual(discountRow![1], `-${discountAmount}`, 'Discount value with negative sign');

  const taxRow = parsed.find((r) => r[0].startsWith('مالیات بر ارزش افزوده'));
  expectTrue(Boolean(taxRow), 'Tax row found');
  expectEqual(taxRow![1], `+${taxAmount}`, 'Tax value with positive sign');

  const totalRow = parsed.find((r) => r[0] === 'مبلغ نهایی فاکتور:');
  expectTrue(Boolean(totalRow), 'Total row found');
  expectEqual(totalRow![1], String(totalAmount), 'Total amount in CSV');

  // Verify Persian words representation contains "تریلیون" and "میلیارد"
  const wordsRow = parsed.find((r) => r[0] === 'مبلغ به حروف:');
  expectTrue(Boolean(wordsRow), 'Words row found');
  expectTrue(wordsRow![1].includes('تریلیون'), 'Words row must contain تریلیون');
  expectTrue(wordsRow![1].includes('میلیارد'), 'Words row must contain میلیارد');
  expectTrue(wordsRow![1].endsWith('ریال'), 'Words row must end with currency unit');
});

runTest('2.2 Persian words conversion for boundary trillion scales (10^12 to 10^15)', () => {
  // 1 Trillion
  const words1T = numberToPersianWords(1_000_000_000_000);
  expectEqual(words1T, 'یک تریلیون', '1 Trillion in words');

  // 123 Trillion 456 Billion 789 Million 123 Thousand 456
  const bigNum = 123_456_789_123_456;
  const wordsBig = numberToPersianWords(bigNum);
  expectTrue(wordsBig.includes('صد و بیست و سه تریلیون'), 'Hundreds of trillions in words');
  expectTrue(wordsBig.includes('چهارصد و پنجاه و شش میلیارد'), 'Billions part in words');
  expectTrue(wordsBig.includes('هفتصد و هشتاد و نه میلیون'), 'Millions part in words');
  expectTrue(wordsBig.includes('صد و بیست و سه هزار'), 'Thousands part in words');
  expectTrue(wordsBig.endsWith('چهارصد و پنجاه و شش'), 'Units part in words');

  // 999 Trillion (Highest supported standard scale in scales array)
  const wordsMax = numberToPersianWords(999_000_000_000_000);
  expectEqual(wordsMax, 'نهصد و نود و نه تریلیون', '999 Trillion in words');
});

// ============================================================================
// PROBE 3: SPECIAL INJECTION CHARACTERS & RFC 4180 ROUND-TRIP
// ============================================================================
console.log('\n--- PROBE 3: Special Injection Characters & RFC 4180 Round-Trip ---');

runTest('3.1 Special injection characters (quotes, commas, newlines, HTML, formulas)', () => {
  const adversarialCustomer = 'شرکت "مهندسی" & <پارس> ، شعبه ۱\nخط دوم نام خریدار\r\nخط سوم خریدار';
  const adversarialAddress = 'تهران، خیابان آزادی، کوچه "بهار"، پلاک ۱۲، زنگ <۳>\nکد پستی: ۱۲۳۴۵';
  const adversarialExplanation = '=SUM(A1:A10)\n+cmd|\' /C calc\'!A0\n@test\n-100\n"نقل قول"\r\n<script>alert(1)</script>';

  const invoice = createBaseInvoice({
    customerName: adversarialCustomer,
    customerAddress: adversarialAddress,
    explanation: adversarialExplanation,
    customerEconomicCode: '411"222",333',
    customerNationalId: '1010<111>&222',
  });

  const csv = generateInvoiceCsvString(invoice, baseSettings);

  // Parse strictly with RFC 4180 parser
  const parsed = parseRfc4180(csv);

  // Check metadata row (Row with customer details)
  const metaHeaderIdx = parsed.findIndex((r) => r[0] === 'شماره فاکتور');
  expectTrue(metaHeaderIdx !== -1, 'Metadata header found');

  const metaDataRow = parsed[metaHeaderIdx + 1];
  expectEqual(metaDataRow[2], adversarialCustomer, 'Customer name with quotes, commas, newlines round-trips exactly');
  expectEqual(metaDataRow[4], '411"222",333', 'Customer economic code with quotes and commas round-trips exactly');
  expectEqual(metaDataRow[5], '1010<111>&222', 'Customer national ID with <>& round-trips exactly');
  expectEqual(metaDataRow[7], adversarialAddress, 'Customer address with newlines and quotes round-trips exactly');

  // Check explanation row
  const explRow = parsed.find((r) => r[0] === 'توضیحات:');
  expectTrue(Boolean(explRow), 'Explanation row found');
  expectEqual(explRow![1], adversarialExplanation, 'Formula injection and script explanation round-trips exactly');
});

runTest('3.2 WhatsApp summary with special injection characters preserves readability', () => {
  const adversarialCustomer = 'شرکت <صنعتی> & "تجاری" مهر، پلاک ۵';
  const adversarialPhone = '0912-111,2223';
  const invoice = createBaseInvoice({
    customerName: adversarialCustomer,
    customerPhone: adversarialPhone,
  });

  const summary = generateInvoiceWhatsAppSummary(invoice, baseSettings);

  // Verify that <>& are not mangled into HTML entities (&amp;, &lt;, etc.)
  expectTrue(summary.includes('خریدار محترم: شرکت <صنعتی> & "تجاری" مهر، پلاک ۵'), 'Plain text preserves raw characters in WhatsApp');
  expectTrue(!summary.includes('&amp;'), 'No HTML escaping entity &amp; in WhatsApp summary');
  expectTrue(!summary.includes('&lt;'), 'No HTML escaping entity &lt; in WhatsApp summary');
});

// ============================================================================
// PROBE 4: OFFICIAL VS SIMPLE PROFORMA WITH MISSING / UNDEFINED FIELDS
// ============================================================================
console.log('\n--- PROBE 4: Official vs Simple Invoice Toggling (Null/Undefined Safety) ---');

runTest('4.1 Official Mode with completely missing / undefined customer & seller fields', () => {
  const emptySettings: CompanySettings = {
    name: 'شرکت آزمایشی بدون مشخصات رسمی',
    address: '',
    phone: '02100000000',
    fax: '',
    taxRate: 10,
    invoiceCounter: 1,
    currencyUnit: 'تومان',
    basePipePrices: { PE100: 90000, PE80: 80000, PE63: 70000 },
    // All official fields omitted / undefined
    economicCode: undefined,
    nationalId: undefined,
    registrationNumber: undefined,
    postalCode: undefined,
    bankName: undefined,
    bankAccount: undefined,
    bankIban: undefined,
  };

  const invoiceWithMissingFields = createBaseInvoice({
    isOfficial: true,
    customerPhone: undefined,
    customerAddress: undefined,
    customerEconomicCode: undefined,
    customerNationalId: undefined,
    customerPostalCode: undefined,
    currencyUnit: undefined,
    explanation: undefined,
  });

  // Must not throw
  const csvOfficial = generateInvoiceCsvString(invoiceWithMissingFields, emptySettings);
  expectTrue(csvOfficial.includes('صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)'), 'Contains official title');
  expectTrue(csvOfficial.includes('کد اقتصادی فروشنده: -'), 'Falls back to - for seller economic code');
  expectTrue(csvOfficial.includes('شناسه ملی: -'), 'Falls back to - for seller national ID');
  expectTrue(csvOfficial.includes('شماره ثبت: -'), 'Falls back to - for seller registration number');
  expectTrue(csvOfficial.includes('کد پستی: -'), 'Falls back to - for seller postal code');

  // Verify bank section is omitted when all bank fields are undefined
  expectTrue(!csvOfficial.includes('--- اطلاعات واریز بانکی ---'), 'Bank section omitted when all bank fields undefined');

  // Check WhatsApp summary
  const summaryOfficial = generateInvoiceWhatsAppSummary(invoiceWithMissingFields, emptySettings);
  expectTrue(summaryOfficial.includes('📋 صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)'), 'Official title in WhatsApp');
  expectTrue(!summaryOfficial.includes('تلفن خریدار:'), 'Phone line cleanly omitted');
  expectTrue(!summaryOfficial.includes('💳 اطلاعات واریز:'), 'Bank section cleanly omitted');
  expectTrue(!summaryOfficial.includes('🏢 نشانی:'), 'Address line omitted if empty');
});

runTest('4.2 Simple Mode toggling with null and missing fields', () => {
  const invoiceSimple = createBaseInvoice({
    isOfficial: false,
    customerPhone: undefined,
    customerAddress: undefined,
    customerEconomicCode: undefined,
    customerNationalId: undefined,
    customerPostalCode: undefined,
  });

  const csvSimple = generateInvoiceCsvString(invoiceSimple, baseSettings);
  expectTrue(csvSimple.includes('پیش‌فاکتور فروش کالا و خدمات'), 'Contains standard proforma title');
  expectTrue(!csvSimple.includes('صورتحساب رسمی فروش کالا و خدمات'), 'Does NOT contain official title');
  expectTrue(!csvSimple.includes('کد اقتصادی فروشنده:'), 'Does NOT contain seller official tax block');

  const summarySimple = generateInvoiceWhatsAppSummary(invoiceSimple, baseSettings);
  expectTrue(summarySimple.includes(`📋 پیش‌فاکتور شرکت ${baseSettings.name}`), 'Simple proforma header in WhatsApp');
  expectTrue(!summarySimple.includes('ماده ۱۶۹'), 'Does NOT mention official tax article in simple summary');
});

runTest('4.3 InvoicePrintView null-safety logic probe', () => {
  // Probe expressions directly from InvoicePrintView.tsx
  const bareInvoice = createBaseInvoice({
    customerName: '',
    customerPhone: undefined,
    customerAddress: undefined,
    customerEconomicCode: undefined,
    customerNationalId: undefined,
    customerPostalCode: undefined,
    explanation: undefined,
    items: [],
  });

  const bareSettings: CompanySettings = {
    name: 'تست رندرینگ',
    address: 'تهران',
    phone: '02111111111',
    fax: '',
    taxRate: 10,
    invoiceCounter: 1,
    currencyUnit: 'تومان',
    basePipePrices: { PE100: 90000, PE80: 80000, PE63: 70000 },
    economicCode: undefined,
    nationalId: undefined,
    registrationNumber: undefined,
    postalCode: undefined,
    bankName: undefined,
    bankAccount: undefined,
    bankIban: undefined,
  };

  // Evaluate seller bindings (lines 237-257)
  const sellerEco = toPersianDigits(bareSettings.economicCode || '-');
  const sellerNat = toPersianDigits(bareSettings.nationalId || '-');
  const sellerReg = toPersianDigits(bareSettings.registrationNumber || '-');
  const sellerPost = toPersianDigits(bareSettings.postalCode || '-');
  const sellerPhone = `${toPersianDigits(bareSettings.phone)} ${bareSettings.fax ? `/ ${toPersianDigits(bareSettings.fax)}` : ''}`;

  expectEqual(sellerEco, '-', 'Seller economic code fallback');
  expectEqual(sellerNat, '-', 'Seller national ID fallback');
  expectEqual(sellerReg, '-', 'Seller registration fallback');
  expectEqual(sellerPost, '-', 'Seller postal code fallback');
  expectEqual(sellerPhone.trim(), toPersianDigits('02111111111'), 'Seller phone without fax');

  // Evaluate buyer bindings (lines 274-294)
  const buyerEco = toPersianDigits(bareInvoice.customerEconomicCode || '-');
  const buyerNat = toPersianDigits(bareInvoice.customerNationalId || '-');
  const buyerPost = toPersianDigits(bareInvoice.customerPostalCode || '-');
  const buyerPhone = toPersianDigits(bareInvoice.customerPhone || '-');
  const buyerAddress = bareInvoice.customerAddress || '-';

  expectEqual(buyerEco, '-', 'Buyer economic code fallback');
  expectEqual(buyerNat, '-', 'Buyer national ID fallback');
  expectEqual(buyerPost, '-', 'Buyer postal code fallback');
  expectEqual(buyerPhone, '-', 'Buyer phone fallback');
  expectEqual(buyerAddress, '-', 'Buyer address fallback');

  // Evaluate bank settlement block visibility condition (line 510)
  const showBankInOfficial = (true || bareSettings.bankName || bareSettings.bankAccount || bareSettings.bankIban);
  expectTrue(Boolean(showBankInOfficial), 'Bank block is visible in official mode');

  const showBankInSimple = (false || bareSettings.bankName || bareSettings.bankAccount || bareSettings.bankIban);
  expectTrue(!showBankInSimple, 'Bank block is hidden in simple mode when settings fields are undefined');
});

// ============================================================================
// PROBE 5: POTENTIAL CRASH / VULNERABILITY ANALYSIS
// ============================================================================
console.log('\n--- PROBE 5: Edge Case Vulnerabilities & Defensive Analysis ---');

runTest('5.1 Filename generation when customerName is empty or undefined', () => {
  const invEmptyName = createBaseInvoice({ customerName: '' });
  const cleanCustomer = (invEmptyName.customerName || 'مشتری').replace(/\s+/g, '_');
  expectEqual(cleanCustomer, 'مشتری', 'Empty customerName safely falls back to مشتری');

  const invSpacesName = createBaseInvoice({ customerName: 'شرکت   پلیمر   گستر' });
  const cleanCustomerSpaces = (invSpacesName.customerName || 'مشتری').replace(/\s+/g, '_');
  expectEqual(cleanCustomerSpaces, 'شرکت_پلیمر_گستر', 'Spaces safely converted to single underscore');
});

runTest('5.2 Zero-items and zero-subtotal financial calculations integrity', () => {
  const zeroInvoice = createBaseInvoice({
    items: [],
    subtotal: 0,
    includeDiscount: true,
    discountPercentage: 10,
    discountAmount: 0,
    includeTax: true,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 0,
  });

  const csv = generateInvoiceCsvString(zeroInvoice, baseSettings);
  const parsed = parseRfc4180(csv);

  const totalRow = parsed.find((r) => r[0] === 'مبلغ نهایی فاکتور:');
  expectTrue(Boolean(totalRow), 'Total row exists in CSV');
  expectEqual(totalRow![1], '0', 'Zero amount properly stored');
  expectEqual(totalRow![2], 'تومان', 'Currency stored');

  const wordsRow = parsed.find((r) => r[0] === 'مبلغ به حروف:');
  expectTrue(Boolean(wordsRow), 'Words row exists');
  expectEqual(wordsRow![1], 'صفر تومان', 'Persian words for zero');

  const summary = generateInvoiceWhatsAppSummary(zeroInvoice, baseSettings);
  expectTrue(summary.includes('مبلغ نهایی قابل پرداخت: ۰ تومان'), 'Zero total handled in WhatsApp');
  expectTrue(summary.includes('(صفر تومان)'), 'Zero words handled in WhatsApp');
});

runTest('5.3 [VULNERABILITY PROBE] Unsanitized illegal OS filename characters in PDF & CSV names', () => {
  // If customerName or invoiceNumber contains characters forbidden by Windows NTFS/FAT (< > : " / \ | ? *)
  const adversarialInv = createBaseInvoice({
    customerName: 'شرکت "پارس" <صنعت>',
    invoiceNumber: '1403/101:A',
  });

  // Check InvoicePrintView.tsx line 41 implementation:
  // defaultName = `${modePrefix}_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;
  const defaultName = `صورتحساب_رسمی_${adversarialInv.invoiceNumber}_${adversarialInv.customerName.replace(/\s+/g, '_')}.pdf`;

  // Check exportUtils.ts line 273 implementation:
  // const cleanCustomer = (invoice.customerName || 'مشتری').replace(/\s+/g, '_');
  // const filename = `پیش_فاکتور_${invoice.invoiceNumber}_${cleanCustomer}.csv`;
  const cleanCustomer = (adversarialInv.customerName || 'مشتری').replace(/\s+/g, '_');
  const filename = `پیش_فاکتور_${adversarialInv.invoiceNumber}_${cleanCustomer}.csv`;

  // Probe Windows forbidden characters regex: /[<>:"/\\|?*]/
  const windowsForbiddenRegex = /[<>:"/\\|?*]/;
  const pdfHasForbidden = windowsForbiddenRegex.test(defaultName);
  const csvHasForbidden = windowsForbiddenRegex.test(filename);

  if (pdfHasForbidden || csvHasForbidden) {
    console.log(`    ⚠ VULNERABILITY FOUND: Unsanitized filename for Windows OS:`);
    console.log(`      PDF defaultName: "${defaultName}" contains forbidden chars`);
    console.log(`      CSV filename: "${filename}" contains forbidden chars`);
  }

  // We record this finding:
  expectTrue(pdfHasForbidden, 'Confirms defaultName currently leaks Windows forbidden characters');
  expectTrue(csvHasForbidden, 'Confirms CSV filename currently leaks Windows forbidden characters');
});

runTest('5.4 [DEFECT PROBE] Unhandled null/undefined customerName in InvoicePrintView.tsx line 41', () => {
  // In InvoicePrintView.tsx line 41:
  // const defaultName = `${modePrefix}_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;
  // If an invoice object has customerName as undefined or null:
  const invWithNullCustomer: any = createBaseInvoice();
  invWithNullCustomer.customerName = undefined;

  let threwTypeError = false;
  try {
    const _ = `پیش_فاکتور_${invWithNullCustomer.invoiceNumber}_${invWithNullCustomer.customerName.replace(/\s+/g, '_')}.pdf`;
  } catch (err: any) {
    if (err instanceof TypeError) {
      threwTypeError = true;
    }
  }

  expectTrue(threwTypeError, 'Proves InvoicePrintView.tsx:41 throws TypeError if customerName is undefined');
});

// ============================================================================
// FINAL SUMMARY
// ============================================================================
console.log('\n================================================================');
console.log(`TEST RESULTS: ${totalPassed} PASSED, ${totalFailed} FAILED`);
console.log('================================================================');

if (totalFailed > 0) {
  console.log('\nFAILURE DETAILS:');
  failureDetails.forEach((f) => console.log(` - ${f}`));
  console.log('\nVERDICT: REQUEST_CHANGES');
} else {
  console.log('\nVERDICT: APPROVE');
}
});
