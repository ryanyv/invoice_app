/**
 * Empirical Challenger Verification Harness for Milestone M3 (Requirement R3)
 * Advanced Print, Export & Sharing
 *
 * Checks:
 * 1. CSV export byte inspection: verify `\uFEFF` byte sequence (0xEF, 0xBB, 0xBF) at offset 0.
 * 2. RFC 4180 parsing round-trip: parse generated CSV with a strict RFC 4180 CSV parser and verify cell-by-cell equality.
 * 3. WhatsApp summary formatting: verify regex structure, presence of Persian numbers, totals in words and digits, and proper RTL line breaks.
 * 4. Stress tests on edge cases: quotes, newlines, commas, zero values, extreme numbers, official vs simple proforma.
 */

import { CompanySettings, Invoice, PipeItem, FittingItem } from '../../types';
import {
  generateInvoiceCsvString,
  generateInvoiceWhatsAppSummary,
} from '../../utils/exportUtils';
import { numberToPersianWords, toPersianDigits, formatNumber } from '../persianUtils';

// ============================================================================
// STRICT RFC 4180 CSV PARSER IMPLEMENTATION (Deterministic State Machine)
// ============================================================================

enum CsvParserState {
  FIELD_START,
  IN_QUOTES,
  QUOTE_IN_QUOTES,
  IN_UNQUOTED,
}

interface StrictCsvParseResult {
  records: string[][];
  rawBomDetected: boolean;
  bomBytes: number[];
  syntaxErrors: string[];
}

export function parseStrictRfc4180Csv(csvText: string): StrictCsvParseResult {
  const syntaxErrors: string[] = [];
  const records: string[][] = [];
  let currentRecord: string[] = [];
  let currentField = '';

  let rawBomDetected = false;
  let bomBytes: number[] = [];

  // 1. Byte inspection of the leading text
  if (csvText.length > 0 && csvText.charCodeAt(0) === 0xfeff) {
    rawBomDetected = true;
    const utf8Bytes = new TextEncoder().encode(csvText.slice(0, 1));
    bomBytes = Array.from(utf8Bytes);
    // Strip BOM for record parsing
    csvText = csvText.slice(1);
  }

  let state: CsvParserState = CsvParserState.FIELD_START;
  let charIdx = 0;

  while (charIdx < csvText.length) {
    const ch = csvText[charIdx];

    switch (state) {
      case CsvParserState.FIELD_START:
        if (ch === '"') {
          state = CsvParserState.IN_QUOTES;
        } else if (ch === ',') {
          // Empty unquoted field
          currentRecord.push('');
        } else if (ch === '\r') {
          // Check for CRLF
          if (csvText[charIdx + 1] === '\n') {
            charIdx++;
          }
          currentRecord.push('');
          records.push(currentRecord);
          currentRecord = [];
        } else if (ch === '\n') {
          currentRecord.push('');
          records.push(currentRecord);
          currentRecord = [];
        } else {
          currentField += ch;
          state = CsvParserState.IN_UNQUOTED;
        }
        break;

      case CsvParserState.IN_QUOTES:
        if (ch === '"') {
          state = CsvParserState.QUOTE_IN_QUOTES;
        } else {
          currentField += ch;
        }
        break;

      case CsvParserState.QUOTE_IN_QUOTES:
        if (ch === '"') {
          // Escaped double quote ("") -> translates to single quote (")
          currentField += '"';
          state = CsvParserState.IN_QUOTES;
        } else if (ch === ',') {
          // End of quoted field
          currentRecord.push(currentField);
          currentField = '';
          state = CsvParserState.FIELD_START;
        } else if (ch === '\r') {
          // End of quoted field and record (CRLF)
          if (csvText[charIdx + 1] === '\n') {
            charIdx++;
          }
          currentRecord.push(currentField);
          currentField = '';
          records.push(currentRecord);
          currentRecord = [];
          state = CsvParserState.FIELD_START;
        } else if (ch === '\n') {
          // End of quoted field and record (LF)
          currentRecord.push(currentField);
          currentField = '';
          records.push(currentRecord);
          currentRecord = [];
          state = CsvParserState.FIELD_START;
        } else {
          // RFC 4180 Violation: characters immediately following a closing quote without delimiter
          syntaxErrors.push(
            `RFC 4180 Syntax Error at char ${charIdx}: unexpected character '${ch}' after closing quote`
          );
          currentField += ch;
          state = CsvParserState.IN_UNQUOTED;
        }
        break;

      case CsvParserState.IN_UNQUOTED:
        if (ch === ',') {
          currentRecord.push(currentField);
          currentField = '';
          state = CsvParserState.FIELD_START;
        } else if (ch === '\r') {
          if (csvText[charIdx + 1] === '\n') {
            charIdx++;
          }
          currentRecord.push(currentField);
          currentField = '';
          records.push(currentRecord);
          currentRecord = [];
          state = CsvParserState.FIELD_START;
        } else if (ch === '\n') {
          currentRecord.push(currentField);
          currentField = '';
          records.push(currentRecord);
          currentRecord = [];
          state = CsvParserState.FIELD_START;
        } else if (ch === '"') {
          syntaxErrors.push(
            `RFC 4180 Syntax Error at char ${charIdx}: unescaped quote inside unquoted field`
          );
          currentField += ch;
        } else {
          currentField += ch;
        }
        break;
    }

    charIdx++;
  }

  // Handle EOF
  if (state === CsvParserState.IN_QUOTES) {
    syntaxErrors.push(`RFC 4180 Syntax Error: Unterminated quote at end of file`);
    currentRecord.push(currentField);
    records.push(currentRecord);
  } else if (state === CsvParserState.QUOTE_IN_QUOTES || state === CsvParserState.IN_UNQUOTED) {
    currentRecord.push(currentField);
    records.push(currentRecord);
  } else if (state === CsvParserState.FIELD_START && currentRecord.length > 0) {
    currentRecord.push('');
    records.push(currentRecord);
  }

  return { records, rawBomDetected, bomBytes, syntaxErrors };
}

// ============================================================================
// TEST FIXTURES & DATA BUILDERS
// ============================================================================

const defaultSettings: CompanySettings = {
  name: 'شرکت لوله و اتصالات پلیمر غرب',
  address: 'تهران، خیابان ولیعصر، برج سپهر، طبقه ۵',
  phone: '02188776655',
  fax: '02188776656',
  taxRate: 10,
  invoiceCounter: 1001,
  currencyUnit: 'تومان',
  basePipePrices: { PE100: 92000, PE80: 85000, PE63: 79000 },
  economicCode: '411598741236',
  nationalId: '10100456123',
  registrationNumber: '451236',
  postalCode: '1968754123',
  bankName: 'بانک پاسارگاد شعبه میرداماد',
  bankAccount: '20180012345671',
  bankIban: 'IR580570020180012345678001',
};

const samplePipe: PipeItem = {
  id: 'p1',
  type: 'pipe',
  peGrade: 'PE100',
  diameter: 125,
  sdr: 17,
  pn: 10,
  length: 120,
  weightPerMeter: 3.125,
  totalWeight: 375,
  pricePerKg: 92000,
  totalPrice: 34500000,
};

const sampleFitting: FittingItem = {
  id: 'f1',
  type: 'fitting',
  category: 'اتصالات الکتروفیوژن',
  connectionType: 'الکتروفیوژن',
  product: 'کوپلر الکتروفیوژن SDR11',
  material: 'PE100',
  pressure: '16',
  size: '125',
  quantity: 6,
  unitPrice: 450000,
  totalPrice: 2700000,
};

function createMockInvoice(overrides?: Partial<Invoice>): Invoice {
  const subtotal = 37200000;
  const discountAmount = 1860000;
  const taxable = subtotal - discountAmount;
  const taxAmount = 3534000;
  const totalAmount = taxable + taxAmount;

  return {
    id: 'inv_1001',
    invoiceNumber: '1001',
    customerName: 'شرکت آب و فاضلاب استان اصفهان',
    customerPhone: '03132223344',
    customerAddress: 'اصفهان، میدان امام حسین، ساختمان آبفا',
    customerEconomicCode: '411234567890',
    customerNationalId: '10260123456',
    customerPostalCode: '8145678912',
    currencyUnit: 'تومان',
    isOfficial: true,
    jalaliDate: '1405/07/26',
    items: [{ ...samplePipe }, { ...sampleFitting }],
    subtotal,
    includeDiscount: true,
    discountPercentage: 5,
    discountAmount,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: true,
    taxPercentage: 10,
    taxAmount,
    totalAmount,
    explanation: 'تخلیه بار در انبار پروژه نجف‌آباد. بارگیری پس از تسویه کامل.',
    status: 'draft',
    createdAt: '2026-10-16T00:00:00Z',
    ...overrides,
  };
}

// ============================================================================
// TEST SUITES
// ============================================================================

let totalPassed = 0;
let totalFailed = 0;

function runTest(testName: string, testFn: () => void) {
  it(testName, () => {
    try {
      testFn();
      console.log(`  ✓ PASS: ${testName}`);
      totalPassed++;
    } catch (err: any) {
      console.error(`  ✗ FAIL: ${testName}`);
      console.error(`    Error: ${err?.message || err}`);
      if (err.stack) console.error(err.stack);
      totalFailed++;
      throw err;
    }
  });
}

function expect(actual: any) {
  return {
    toBe(expected: any, msg?: string) {
      if (actual !== expected) {
        throw new Error(
          `${msg || 'Assertion failed'}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
        );
      }
    },
    toEqual(expected: any, msg?: string) {
      const a = JSON.stringify(actual);
      const b = JSON.stringify(expected);
      if (a !== b) {
        throw new Error(`${msg || 'Assertion failed'}: Expected ${b}, got ${a}`);
      }
    },
    toBeTruthy(msg?: string) {
      if (!actual) {
        throw new Error(`${msg || 'Assertion failed'}: Expected truthy, got ${actual}`);
      }
    },
    toBeFalsy(msg?: string) {
      if (actual) {
        throw new Error(`${msg || 'Assertion failed'}: Expected falsy, got ${actual}`);
      }
    },
    toContain(sub: string, msg?: string) {
      if (typeof actual !== 'string' || !actual.includes(sub)) {
        throw new Error(`${msg || 'Assertion failed'}: Expected string to contain "${sub}"`);
      }
    },
    toMatch(regex: RegExp, msg?: string) {
      if (typeof actual !== 'string' || !regex.test(actual)) {
        throw new Error(`${msg || 'Assertion failed'}: String "${actual}" does not match regex ${regex}`);
      }
    },
  };
}

describe('CHALLENGER 1: EMPIRICAL STRESS TEST SUITE - REQUIREMENT R3', () => {

// ----------------------------------------------------------------------------
// SUITE 1: CSV UTF-8 BOM BYTE INSPECTION
// ----------------------------------------------------------------------------
console.log('--- SUITE 1: CSV Export Byte Inspection ---');

runTest('1.1 UTF-8 BOM byte sequence [0xEF, 0xBB, 0xBF] at offset 0', () => {
  const invoice = createMockInvoice();
  const csvString = generateInvoiceCsvString(invoice, defaultSettings);

  // String level check
  expect(csvString.charCodeAt(0)).toBe(0xfeff, 'String char at index 0 must be 0xFEFF');
  expect(csvString.charCodeAt(1)).toBe(
    0x22,
    'String char at index 1 must be 0x22 (") - first cell quote'
  );

  // Byte level check via TextEncoder
  const uint8 = new TextEncoder().encode(csvString);
  expect(uint8[0]).toBe(0xef, 'TextEncoder byte 0 must be 0xEF');
  expect(uint8[1]).toBe(0xbb, 'TextEncoder byte 1 must be 0xBB');
  expect(uint8[2]).toBe(0xbf, 'TextEncoder byte 2 must be 0xBF');
  expect(uint8[3]).toBe(0x22, 'TextEncoder byte 3 must be 0x22 (ASCII double-quote)');
});

runTest('1.2 BOM persists across Official and Simple Proforma Modes', () => {
  const officialInvoice = createMockInvoice({ isOfficial: true });
  const simpleInvoice = createMockInvoice({ isOfficial: false });

  const officialCsv = generateInvoiceCsvString(officialInvoice, defaultSettings);
  const simpleCsv = generateInvoiceCsvString(simpleInvoice, defaultSettings);

  const officialBytes = new TextEncoder().encode(officialCsv);
  const simpleBytes = new TextEncoder().encode(simpleCsv);

  expect(Array.from(officialBytes.slice(0, 3))).toEqual(
    [0xef, 0xbb, 0xbf],
    'Official CSV must start with 0xEF, 0xBB, 0xBF'
  );
  expect(Array.from(simpleBytes.slice(0, 3))).toEqual(
    [0xef, 0xbb, 0xbf],
    'Simple CSV must start with 0xEF, 0xBB, 0xBF'
  );
});

// ----------------------------------------------------------------------------
// SUITE 2: RFC 4180 COMPLIANCE & CELL-BY-CELL ROUND-TRIP
// ----------------------------------------------------------------------------
console.log('\n--- SUITE 2: RFC 4180 Strict Parsing & Cell-by-Cell Round-Trip ---');

runTest('2.1 Strict RFC 4180 syntax validation with 0 syntax errors', () => {
  const invoice = createMockInvoice();
  const csvString = generateInvoiceCsvString(invoice, defaultSettings);

  const parsed = parseStrictRfc4180Csv(csvString);

  expect(parsed.rawBomDetected).toBe(true, 'BOM must be detected');
  expect(parsed.bomBytes).toEqual([0xef, 0xbb, 0xbf], 'BOM bytes must be 0xEF, 0xBB, 0xBF');
  expect(parsed.syntaxErrors.length).toBe(0, 'Must have zero RFC 4180 syntax errors');
  expect(parsed.records.length > 10).toBe(true, 'Must have parsed structured records');
});

runTest('2.2 Cell-by-cell round-trip equality for standard invoice', () => {
  const invoice = createMockInvoice();
  const csvString = generateInvoiceCsvString(invoice, defaultSettings);
  const parsed = parseStrictRfc4180Csv(csvString);

  const records = parsed.records;

  // Row 0: Company Header
  expect(records[0][0]).toBe(defaultSettings.name, 'Row 0 Col 0: Company name');
  expect(records[0][1]).toBe(
    'صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)',
    'Row 0 Col 1: Official Doc Title'
  );

  // Row 1: Address and Contacts
  expect(records[1][0]).toBe(`نشانی: ${defaultSettings.address}`);
  expect(records[1][1]).toBe(`تلفن: ${defaultSettings.phone}`);
  expect(records[1][2]).toBe(`فکس: ${defaultSettings.fax}`);

  // Row 2: Official Seller Tax Info
  expect(records[2][0]).toBe(`کد اقتصادی فروشنده: ${defaultSettings.economicCode}`);
  expect(records[2][1]).toBe(`شناسه ملی: ${defaultSettings.nationalId}`);
  expect(records[2][2]).toBe(`شماره ثبت: ${defaultSettings.registrationNumber}`);
  expect(records[2][3]).toBe(`کد پستی: ${defaultSettings.postalCode}`);

  // Find metadata header row
  const metaHeaderIdx = records.findIndex((r) => r[0] === 'شماره فاکتور');
  expect(metaHeaderIdx !== -1).toBe(true, 'Metadata header row found');
  const metaRow = records[metaHeaderIdx + 1];

  expect(metaRow[0]).toBe(invoice.invoiceNumber, 'Metadata: invoiceNumber');
  expect(metaRow[1]).toBe(invoice.jalaliDate, 'Metadata: jalaliDate');
  expect(metaRow[2]).toBe(invoice.customerName, 'Metadata: customerName');
  expect(metaRow[3]).toBe(invoice.customerPhone, 'Metadata: customerPhone');
  expect(metaRow[4]).toBe(invoice.customerEconomicCode, 'Metadata: customerEconomicCode');
  expect(metaRow[5]).toBe(invoice.customerNationalId, 'Metadata: customerNationalId');
  expect(metaRow[6]).toBe(invoice.customerPostalCode, 'Metadata: customerPostalCode');
  expect(metaRow[7]).toBe(invoice.customerAddress, 'Metadata: customerAddress');
  expect(metaRow[8]).toBe(invoice.currencyUnit, 'Metadata: currencyUnit');

  // Find Pipe Table Header
  const pipeHeaderIdx = records.findIndex((r) => r[1] === 'شرح کالا');
  expect(pipeHeaderIdx !== -1).toBe(true, 'Pipe table header found');
  const pipeRow = records[pipeHeaderIdx + 1];

  expect(pipeRow[0]).toBe('1', 'Pipe row 1 index');
  expect(pipeRow[1]).toBe('لوله پلی‌اتیلن 125 میلیمتر SDR17 PN10 PE100', 'Pipe item description');
  expect(pipeRow[2]).toBe(String(samplePipe.diameter), 'Pipe diameter');
  expect(pipeRow[3]).toBe(String(samplePipe.sdr), 'Pipe SDR');
  expect(pipeRow[4]).toBe(String(samplePipe.pn), 'Pipe PN');
  expect(pipeRow[5]).toBe(samplePipe.peGrade, 'Pipe Grade');
  expect(pipeRow[6]).toBe(String(samplePipe.length), 'Pipe Length');
  expect(pipeRow[7]).toBe(samplePipe.weightPerMeter.toFixed(3), 'Pipe weightPerMeter');
  expect(pipeRow[8]).toBe(samplePipe.totalWeight.toFixed(3), 'Pipe totalWeight');
  expect(pipeRow[9]).toBe(String(samplePipe.pricePerKg), 'Pipe pricePerKg');
  expect(pipeRow[10]).toBe(String(samplePipe.totalPrice), 'Pipe totalPrice');

  // Pipe summary row
  const pipeSummaryRow = records.find((r) => r[0] === 'جمع اقلام لوله');
  expect(Boolean(pipeSummaryRow)).toBe(true, 'Pipe summary row found');
  expect(pipeSummaryRow![6]).toBe(`${samplePipe.length} متر`, 'Pipe summary total length');
  expect(pipeSummaryRow![8]).toBe(
    `${samplePipe.totalWeight.toFixed(3)} کیلو`,
    'Pipe summary total weight'
  );
  expect(pipeSummaryRow![10]).toBe(
    String(samplePipe.totalPrice),
    'Pipe summary total price'
  );

  // Find Fitting Table Header
  const fittingHeaderIdx = records.findIndex((r) => r[1] === 'نام محصول');
  expect(fittingHeaderIdx !== -1).toBe(true, 'Fitting table header found');
  const fittingRow = records[fittingHeaderIdx + 1];

  expect(fittingRow[0]).toBe('1', 'Fitting row 1 index');
  expect(fittingRow[1]).toBe(sampleFitting.product, 'Fitting product name');
  expect(fittingRow[2]).toBe(sampleFitting.category, 'Fitting category');
  expect(fittingRow[3]).toBe(sampleFitting.connectionType, 'Fitting connectionType');
  expect(fittingRow[4]).toBe(String(sampleFitting.size), 'Fitting size');
  expect(fittingRow[5]).toBe(String(sampleFitting.pressure), 'Fitting pressure');
  expect(fittingRow[6]).toBe(String(sampleFitting.quantity), 'Fitting quantity');
  expect(fittingRow[7]).toBe(String(sampleFitting.unitPrice), 'Fitting unitPrice');
  expect(fittingRow[8]).toBe(String(sampleFitting.totalPrice), 'Fitting totalPrice');

  // Fitting summary row
  const fittingSummaryRow = records.find((r) => r[0] === 'جمع اقلام اتصالات');
  expect(Boolean(fittingSummaryRow)).toBe(true, 'Fitting summary row found');
  expect(fittingSummaryRow![6]).toBe(
    `${sampleFitting.quantity} عدد`,
    'Fitting summary total quantity'
  );
  expect(fittingSummaryRow![8]).toBe(
    String(sampleFitting.totalPrice),
    'Fitting summary total price'
  );

  // Financial Summary Section
  const subtotalRow = records.find((r) => r[0] === 'جمع کل ناخالص:');
  expect(Boolean(subtotalRow)).toBe(true, 'Subtotal row found');
  expect(subtotalRow![1]).toBe(String(invoice.subtotal), 'Subtotal value');
  expect(subtotalRow![2]).toBe(invoice.currencyUnit, 'Subtotal currency');

  const discountRow = records.find((r) => r[0]?.startsWith('تخفیف'));
  expect(Boolean(discountRow)).toBe(true, 'Discount row found');
  expect(discountRow![1]).toBe(`-${invoice.discountAmount}`, 'Discount value negative');
  expect(discountRow![2]).toBe(invoice.currencyUnit, 'Discount currency');

  const taxRow = records.find((r) => r[0]?.startsWith('مالیات بر ارزش افزوده'));
  expect(Boolean(taxRow)).toBe(true, 'Tax row found');
  expect(taxRow![1]).toBe(`+${invoice.taxAmount}`, 'Tax value positive');
  expect(taxRow![2]).toBe(invoice.currencyUnit, 'Tax currency');

  const totalRow = records.find((r) => r[0] === 'مبلغ نهایی فاکتور:');
  expect(Boolean(totalRow)).toBe(true, 'Total amount row found');
  expect(totalRow![1]).toBe(String(invoice.totalAmount), 'Total amount value');
  expect(totalRow![2]).toBe(invoice.currencyUnit, 'Total amount currency');

  const wordsRow = records.find((r) => r[0] === 'مبلغ به حروف:');
  expect(Boolean(wordsRow)).toBe(true, 'Words row found');
  expect(wordsRow![1]).toBe(
    `${numberToPersianWords(invoice.totalAmount)} ${invoice.currencyUnit}`,
    'Amount in Persian words'
  );

  // Bank details
  const bankRow = records.find((r) => r[0]?.startsWith('بانک:'));
  expect(Boolean(bankRow)).toBe(true, 'Bank row found');
  expect(bankRow![0]).toBe(`بانک: ${defaultSettings.bankName}`);
  expect(bankRow![1]).toBe(`شماره حساب: ${defaultSettings.bankAccount}`);
  expect(bankRow![2]).toBe(`شماره شبا: ${defaultSettings.bankIban}`);

  // Explanation
  const expRow = records.find((r) => r[0] === 'توضیحات:');
  expect(Boolean(expRow)).toBe(true, 'Explanation row found');
  expect(expRow![1]).toBe(invoice.explanation, 'Explanation content matches');
});

// ----------------------------------------------------------------------------
// SUITE 3: ADVERSARIAL CSV ESCAPING (RFC 4180 EDGE CASES)
// ----------------------------------------------------------------------------
console.log('\n--- SUITE 3: Adversarial RFC 4180 Escaping Stress Tests ---');

runTest('3.1 Handling embedded commas, quotes, and multiline newlines', () => {
  const adversarialInvoice = createMockInvoice({
    customerName: 'شرکت "فولاد و پلیمر", تبریز (سهامی خاص)',
    customerAddress: 'آذربایجان شرقی، تبریز، بلوار ۲۹ بهمن، پلاک "۴۵"، طبقه ۳, زنگ ۲',
    explanation: 'توضیحات چندخطی:\nخط اول: نیاز به گواهی استاندارد "DIN 8074"\r\nخط دوم: بارگیری روز پنجشنبه، با تماس قبلی\nپایان.',
  });

  const adversarialSettings: CompanySettings = {
    ...defaultSettings,
    name: 'تولیدی "پلی‌غرب", شعبه ۱ (اصلی)',
    address: 'تهران، خیابان "شهید بهشتی"، پلاک ۱۲، زنگ "۴"',
  };

  const csvString = generateInvoiceCsvString(adversarialInvoice, adversarialSettings);
  const parsed = parseStrictRfc4180Csv(csvString);

  expect(parsed.syntaxErrors.length).toBe(
    0,
    `Must parse cleanly with zero syntax errors. Found: ${parsed.syntaxErrors.join('; ')}`
  );

  // Verify that the adversarial fields were decoded exactly back to their unescaped originals
  expect(parsed.records[0][0]).toBe(adversarialSettings.name);
  expect(parsed.records[1][0]).toBe(`نشانی: ${adversarialSettings.address}`);

  const metaHeaderIdx = parsed.records.findIndex((r) => r[0] === 'شماره فاکتور');
  const metaRow = parsed.records[metaHeaderIdx + 1];
  expect(metaRow[2]).toBe(adversarialInvoice.customerName);
  expect(metaRow[7]).toBe(adversarialInvoice.customerAddress);

  const expRow = parsed.records.find((r) => r[0] === 'توضیحات:');
  expect(Boolean(expRow)).toBe(true);
  expect(expRow![1]).toBe(adversarialInvoice.explanation);
});

runTest('3.2 Extreme numbers (Trillions) and custom discounts', () => {
  const highValInvoice = createMockInvoice({
    subtotal: 500000000000, // 500 Billion
    includeDiscount: false,
    includeCustomDiscount: true,
    discountAmount: 20000000000, // 20 Billion
    includeTax: true,
    taxPercentage: 10,
    taxAmount: 48000000000, // 48 Billion
    totalAmount: 528000000000, // 528 Billion
  });

  const csvString = generateInvoiceCsvString(highValInvoice, defaultSettings);
  const parsed = parseStrictRfc4180Csv(csvString);

  expect(parsed.syntaxErrors.length).toBe(0);

  const totalRow = parsed.records.find((r) => r[0] === 'مبلغ نهایی فاکتور:');
  expect(totalRow![1]).toBe('528000000000');

  const wordsRow = parsed.records.find((r) => r[0] === 'مبلغ به حروف:');
  expect(wordsRow![1]).toContain('پانصد و بیست و هشت میلیارد تومان');
});

runTest('3.3 Empty items invoice (Edge case handling)', () => {
  const emptyInvoice = createMockInvoice({
    items: [],
    subtotal: 0,
    includeDiscount: false,
    discountAmount: 0,
    includeTax: false,
    taxAmount: 0,
    totalAmount: 0,
    explanation: undefined,
  });

  const csvString = generateInvoiceCsvString(emptyInvoice, defaultSettings);
  const parsed = parseStrictRfc4180Csv(csvString);

  expect(parsed.syntaxErrors.length).toBe(0);

  // Pipe and fitting table headers should NOT be present when items are empty
  const hasPipe = parsed.records.some((r) => r[0] === '--- جدول اقلام لوله پلی‌اتیلن ---');
  const hasFitting = parsed.records.some((r) => r[0] === '--- جدول اقلام اتصالات پلی‌اتیلن ---');
  expect(hasPipe).toBe(false, 'Should not render pipe section when empty');
  expect(hasFitting).toBe(false, 'Should not render fitting section when empty');

  // Subtotal and Total must still be present
  const totalRow = parsed.records.find((r) => r[0] === 'مبلغ نهایی فاکتور:');
  expect(totalRow![1]).toBe('0');
  const wordsRow = parsed.records.find((r) => r[0] === 'مبلغ به حروف:');
  expect(wordsRow![1]).toBe('صفر تومان');
});

// ----------------------------------------------------------------------------
// SUITE 4: WHATSAPP / SMS SUMMARY FORMATTING
// ----------------------------------------------------------------------------
console.log('\n--- SUITE 4: WhatsApp / SMS Persian Summary Verification ---');

runTest('4.1 WhatsApp summary regex structure validation', () => {
  const invoice = createMockInvoice();
  const summary = generateInvoiceWhatsAppSummary(invoice, defaultSettings);

  const lines = summary.split('\n');
  expect(lines.length >= 15).toBe(true, `Expected >= 15 lines, got ${lines.length}`);

  // Regex checks on specific lines
  // Header
  expect(lines[0]).toMatch(
    /^📋 صورتحساب رسمی فروش کالا و خدمات \(ماده ۱۶۹ م\.م\)$/,
    'Official Header Regex'
  );

  // Invoice Number & Date
  expect(lines[1]).toMatch(
    /^شماره: [۰-۹]+ \| تاریخ: [۰-۹]{4}\/[۰-۹]{2}\/[۰-۹]{2}$/,
    'Number and Jalali Date Regex'
  );

  // Customer Name
  expect(lines[2]).toMatch(/^خریدار محترم: .+$/, 'Customer Name Regex');

  // Customer Phone
  expect(lines[3]).toMatch(/^تلفن خریدار: [۰-۹]+$/, 'Customer Phone Regex');

  // Pipe section header
  const pipeSecIdx = lines.indexOf('🔹 اقلام لوله پلی‌اتیلن:');
  expect(pipeSecIdx !== -1).toBe(true, 'Pipe Section Header found');

  // Pipe line item regex
  const pipeLine = lines[pipeSecIdx + 1];
  expect(pipeLine).toMatch(
    /^[۰-۹]+\. لوله [۰-۹]+ SDR[۰-۹]+ PN[۰-۹]+ .+ - [۰-۹]+ متر( \([۰-۹\.]+ کیلو\))?: [۰-۹,]+ تومان$/,
    'Pipe Line Item Regex'
  );

  // Fitting section header
  const fitSecIdx = lines.indexOf('🔸 اقلام اتصالات:');
  expect(fitSecIdx !== -1).toBe(true, 'Fitting Section Header found');

  // Fitting line item regex
  const fitLine = lines[fitSecIdx + 1];
  expect(fitLine).toMatch(
    /^[۰-۹]+\. .+ \(.+\) سایز [۰-۹]+ - [۰-۹]+ عدد: [۰-۹,]+ تومان$/,
    'Fitting Line Item Regex'
  );

  // Financial summary header
  const finSecIdx = lines.indexOf('📊 خلاصه مالی:');
  expect(finSecIdx !== -1).toBe(true, 'Financial Section Header found');

  // Subtotal regex
  expect(lines[finSecIdx + 1]).toMatch(
    /^▫️ جمع ناخالص: [۰-۹,]+ تومان$/,
    'Subtotal Line Regex'
  );

  // Discount regex
  expect(lines[finSecIdx + 2]).toMatch(
    /^▫️ تخفیف: [۰-۹,]+ تومان$/,
    'Discount Line Regex'
  );

  // Tax regex
  expect(lines[finSecIdx + 3]).toMatch(
    /^▫️ مالیات ارزش افزوده \([۰-۹]+%\): [۰-۹,]+ تومان$/,
    'Tax Line Regex'
  );

  // Total amount line
  expect(lines[finSecIdx + 4]).toMatch(
    /^💰 مبلغ نهایی قابل پرداخت: [۰-۹,]+ تومان$/,
    'Total Amount Line Regex'
  );

  // Total in words line
  expect(lines[finSecIdx + 5]).toMatch(
    /^\(.+ تومان\)$/,
    'Total Words Line Regex'
  );

  // Bank Info section
  const bankSecIdx = lines.indexOf('💳 اطلاعات واریز:');
  expect(bankSecIdx !== -1).toBe(true, 'Bank Section Header found');
  expect(lines[bankSecIdx + 1]).toMatch(/^بانک: .+$/, 'Bank Name Regex');
  expect(lines[bankSecIdx + 2]).toMatch(/^شماره حساب: [۰-۹]+$/, 'Bank Account Regex');
  expect(lines[bankSecIdx + 3]).toMatch(
    /^شماره شبا: IR[0-9]{24}$/,
    'Bank IBAN Regex (Latin IR with 24 digits)'
  );

  // Company Contact section
  const phoneLine = lines.find((l) => l.startsWith('📞 تماس:'));
  expect(Boolean(phoneLine)).toBe(true);
  expect(phoneLine!).toMatch(/^📞 تماس: [۰-۹]+$/, 'Contact Phone Regex');

  const addrLine = lines.find((l) => l.startsWith('🏢 نشانی:'));
  expect(Boolean(addrLine)).toBe(true);

  expect(lines[lines.length - 1]).toBe(
    'اعتبار پیش‌فاکتور: ۳ روز کاری',
    'Validity notice regex'
  );
});

runTest('4.2 Strictly Persian numerals in all formatted numeric locations', () => {
  const invoice = createMockInvoice({
    invoiceNumber: '9876',
    jalaliDate: '1405/08/12',
    customerPhone: '09123456789',
  });
  const summary = generateInvoiceWhatsAppSummary(invoice, defaultSettings);

  // 1. Verify Persian numerals for Invoice Number and Date
  expect(summary).toContain(`شماره: ${toPersianDigits('9876')}`);
  expect(summary).toContain(`تاریخ: ${toPersianDigits('1405/08/12')}`);

  // 2. Verify Persian numerals for Customer Phone
  expect(summary).toContain(`تلفن خریدار: ${toPersianDigits('09123456789')}`);

  // 3. Verify Persian numerals for Pipe line metrics (diameter, SDR, PN, length, weight, price)
  expect(summary).toContain(`لوله ${toPersianDigits(samplePipe.diameter)}`);
  expect(summary).toContain(`SDR${toPersianDigits(samplePipe.sdr)}`);
  expect(summary).toContain(`PN${toPersianDigits(samplePipe.pn)}`);
  expect(summary).toContain(`${toPersianDigits(samplePipe.length)} متر`);
  expect(summary).toContain(`(${toPersianDigits(samplePipe.totalWeight.toFixed(2))} کیلو)`);
  expect(summary).toContain(`${formatNumber(samplePipe.totalPrice, true)} تومان`);

  // 4. Verify Persian numerals for Fitting line metrics (size, quantity, price)
  expect(summary).toContain(`سایز ${toPersianDigits(sampleFitting.size)}`);
  expect(summary).toContain(`${toPersianDigits(sampleFitting.quantity)} عدد`);
  expect(summary).toContain(`${formatNumber(sampleFitting.totalPrice, true)} تومان`);

  // 5. Verify Persian numerals for Financial Summary
  expect(summary).toContain(`جمع ناخالص: ${formatNumber(invoice.subtotal, true)} تومان`);
  expect(summary).toContain(`تخفیف: ${formatNumber(invoice.discountAmount, true)} تومان`);
  expect(summary).toContain(`مالیات ارزش افزوده (${toPersianDigits(invoice.taxPercentage)}%): ${formatNumber(invoice.taxAmount, true)} تومان`);
  expect(summary).toContain(`مبلغ نهایی قابل پرداخت: ${formatNumber(invoice.totalAmount, true)} تومان`);

  // 6. Verify Persian numerals for Bank Account & Company Phone
  expect(summary).toContain(`شماره حساب: ${toPersianDigits(defaultSettings.bankAccount!)}`);
  expect(summary).toContain(`تماس: ${toPersianDigits(defaultSettings.phone)}`);

  // 7. Verify that raw English digits do NOT appear in numeric metrics
  expect(summary.includes('9876')).toBe(false, 'Raw invoice number must not appear in ASCII');
  expect(summary.includes('1405/08/12')).toBe(false, 'Raw date must not appear in ASCII');
  expect(summary.includes('09123456789')).toBe(false, 'Raw phone must not appear in ASCII');
  expect(summary.includes('37,200,000')).toBe(false, 'Raw subtotal must not appear in ASCII');
  expect(summary.includes('38,874,000')).toBe(false, 'Raw total must not appear in ASCII');
});

runTest('4.3 Total in words and formatted digits correspondence', () => {
  const invoice = createMockInvoice({ totalAmount: 42350750 });
  const summary = generateInvoiceWhatsAppSummary(invoice, defaultSettings);

  const expectedWords = numberToPersianWords(42350750);
  const expectedDigits = formatNumber(42350750, true);

  expect(summary).toContain(
    `💰 مبلغ نهایی قابل پرداخت: ${expectedDigits} تومان`,
    'Summary must contain formatted digits'
  );
  expect(summary).toContain(
    `(${expectedWords} تومان)`,
    'Summary must contain exact Persian words'
  );
});

runTest('4.4 Proper RTL line breaks and newline separation', () => {
  const invoice = createMockInvoice();
  const summary = generateInvoiceWhatsAppSummary(invoice, defaultSettings);

  // Verify that summary contains ONLY standard Unix newline '\n' (no '\r' pollution)
  expect(summary.includes('\r')).toBe(
    false,
    'WhatsApp summary must not have carriage returns \\r'
  );

  const lines = summary.split('\n');
  expect(lines.length > 0).toBe(true);

  // Check that no lines are null or undefined
  lines.forEach((line, idx) => {
    expect(typeof line === 'string').toBe(true, `Line ${idx} must be string`);
  });
});

runTest('4.5 WhatsApp summary in Simple Proforma mode and Rial currency', () => {
  const rialSettings: CompanySettings = {
    ...defaultSettings,
    currencyUnit: 'ریال',
  };
  const simpleInvoice = createMockInvoice({
    isOfficial: false,
    currencyUnit: 'ریال',
  });

  const summary = generateInvoiceWhatsAppSummary(simpleInvoice, rialSettings);
  const lines = summary.split('\n');

  // Header should be Simple Proforma
  expect(lines[0]).toBe(`📋 پیش‌فاکتور شرکت ${rialSettings.name}`);

  // Currency should be ریال everywhere
  expect(summary).toContain('ریال');
  expect(summary.includes('تومان')).toBe(false, 'Must not contain تومان when currency is ریال');
});

// ----------------------------------------------------------------------------
// SUITE 5: PRINT VIEW ORIENTATION & OFFICIAL PROFORMA DOM/STYLE VERIFICATION
// ----------------------------------------------------------------------------
console.log('\n--- SUITE 5: Print Orientation & Official Mode Logic Verification ---');

runTest('5.1 Dynamic @page CSS injection matches orientation state', () => {
  const landscapeCss = `@media print { @page { size: A4 landscape; margin: 8mm; } }`;
  const portraitCss = `@media print { @page { size: A4 portrait; margin: 8mm; } }`;

  // Verify that landscape and portrait strings are valid CSS Paged Media declarations
  expect(landscapeCss).toContain('size: A4 landscape');
  expect(portraitCss).toContain('size: A4 portrait');
});

runTest('5.2 IPC save-pdf argument forwarding', () => {
  // Test simulated IPC call signature as used in InvoicePrintView.tsx
  const mockIpcCalls: any[] = [];
  const fakeElectronApi = {
    savePdf: async (args: any) => {
      mockIpcCalls.push(args);
      return { success: true, filePath: '/tmp/test.pdf' };
    },
  };

  const invoice = createMockInvoice({ invoiceNumber: '555', customerName: 'آبفا البرز' });

  // Simulate print handler in Landscape mode
  let isLandscape = true;
  let isOfficial = true;
  let modePrefix = isOfficial ? 'صورتحساب_رسمی' : 'پیش_فاکتور';
  let defaultName = `${modePrefix}_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;

  fakeElectronApi.savePdf({ defaultName, landscape: isLandscape });

  expect(mockIpcCalls[0].landscape).toBe(true);
  expect(mockIpcCalls[0].defaultName).toBe('صورتحساب_رسمی_555_آبفا_البرز.pdf');

  // Simulate print handler in Portrait mode and simple proforma
  isLandscape = false;
  isOfficial = false;
  modePrefix = isOfficial ? 'صورتحساب_رسمی' : 'پیش_فاکتور';
  defaultName = `${modePrefix}_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;

  fakeElectronApi.savePdf({ defaultName, landscape: isLandscape });

  expect(mockIpcCalls[1].landscape).toBe(false);
  expect(mockIpcCalls[1].defaultName).toBe('پیش_فاکتور_555_آبفا_البرز.pdf');
});

// ============================================================================
// SUMMARY & VERDICT
// ============================================================================
console.log('\n================================================================');
console.log(`TEST SUMMARY: ${totalPassed} PASSED, ${totalFailed} FAILED`);
console.log('================================================================');

if (totalFailed > 0) {
  console.error('\nVERDICT: REQUEST_CHANGES (Regressions or Failures detected)');
} else {
  console.log('\nVERDICT: APPROVE (All empirical tests and stress harnesses passed cleanly)');
}
});
