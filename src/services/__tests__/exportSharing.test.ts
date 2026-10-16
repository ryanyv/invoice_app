import { CompanySettings, Invoice, PipeItem, FittingItem } from '../../types';
import {
  generateInvoiceCsvString,
  generateInvoiceWhatsAppSummary,
} from '../../utils/exportUtils';
import { numberToPersianWords, toPersianDigits } from '../persianUtils';

function assert(condition: boolean, message: string) {
  if (!condition) {
    throw new Error(`Assertion failed: ${message}`);
  }
}

function assertEquals<T>(actual: T, expected: T, message: string) {
  if (actual !== expected) {
    throw new Error(`Assertion failed: ${message} (Expected: ${String(expected)}, Got: ${String(actual)})`);
  }
}

console.log('=== RUNNING ADVANCED PRINT, EXPORT & SHARING TEST SUITE (R3) ===');

// Mock data fixtures
const mockSettings: CompanySettings = {
  name: 'شرکت پلی غرب اتصال ایرانیان',
  address: 'تهران، بلوار میرداماد، میدان مادر، پلاک ۳۴',
  phone: '02122922999',
  fax: '02122278331',
  taxRate: 10,
  invoiceCounter: 105,
  currencyUnit: 'تومان',
  basePipePrices: {
    PE100: 95000,
    PE80: 88000,
    PE63: 82000,
  },
  economicCode: '411456789123',
  nationalId: '10103567890',
  registrationNumber: '384521',
  postalCode: '1548763211',
  bankName: 'بانک ملت شعبه میرداماد',
  bankAccount: '1234567890',
  bankIban: 'IR120120000000012345678901',
};

const samplePipeItem: PipeItem = {
  id: 'pipe_1',
  type: 'pipe',
  peGrade: 'PE100',
  diameter: 110,
  sdr: 11,
  pn: 16,
  length: 100,
  weightPerMeter: 2.38,
  totalWeight: 238,
  pricePerKg: 95000,
  totalPrice: 22610000,
};

const sampleFittingItem: FittingItem = {
  id: 'fit_1',
  type: 'fitting',
  category: 'اتصالات جوشی فشار قوی',
  connectionType: 'جوشی',
  product: 'فلنج پایه بلند',
  material: 'HDPE100',
  pressure: '16',
  size: '110',
  quantity: 4,
  unitPrice: 600000,
  totalPrice: 2400000,
};

const mockInvoice: Invoice = {
  id: 'inv_101',
  invoiceNumber: '101',
  customerName: 'شرکت مهندسی آب و خاک کاسپین',
  customerPhone: '09121112233',
  customerAddress: 'قزوین، شهرک صنعتی کاسپین، بلوار صنعت',
  customerEconomicCode: '411122334455',
  customerNationalId: '10100987654',
  customerPostalCode: '3414156789',
  currencyUnit: 'تومان',
  isOfficial: true,
  jalaliDate: '1405/07/25',
  items: [samplePipeItem, sampleFittingItem],
  subtotal: 25010000,
  includeDiscount: true,
  discountPercentage: 5,
  discountAmount: 1250500,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: true,
  taxPercentage: 10,
  taxAmount: 2375950,
  totalAmount: 26135450,
  explanation: 'تحویل درب کارخانه، هزینه حمل با خریدار است.',
  status: 'finalized',
  createdAt: '2026-10-16T00:00:00Z',
};

// -------------------------------------------------------------
// Test Suite 1: CSV Generation & UTF-8 BOM
// -------------------------------------------------------------
console.log('--- Test Suite 1: CSV Export Formatting & UTF-8 BOM ---');

const csvResult = generateInvoiceCsvString(mockInvoice, mockSettings);

// 1.1 UTF-8 BOM verification
assert(csvResult.startsWith('\uFEFF'), 'CSV must start with UTF-8 Byte Order Mark (\\uFEFF)');
console.log('✓ Test 1.1 Passed: UTF-8 BOM verified.');

// 1.2 Company info in CSV
assert(csvResult.includes(mockSettings.name), 'CSV must contain company name');
assert(csvResult.includes(mockSettings.phone), 'CSV must contain company phone');
assert(csvResult.includes(mockSettings.address), 'CSV must contain company address');
assert(csvResult.includes('کد اقتصادی فروشنده: 411456789123'), 'CSV must contain seller economic code in official mode');
console.log('✓ Test 1.2 Passed: Company header and official tax registration verified in CSV.');

// 1.3 Invoice & Customer metadata in CSV
assert(csvResult.includes(mockInvoice.invoiceNumber), 'CSV must contain invoice number');
assert(csvResult.includes(mockInvoice.jalaliDate), 'CSV must contain Jalali date');
assert(csvResult.includes(mockInvoice.customerName), 'CSV must contain customer name');
assert(csvResult.includes(mockInvoice.customerPhone!), 'CSV must contain customer phone');
assert(csvResult.includes('411122334455'), 'CSV must contain customer economic code');
assert(csvResult.includes('10100987654'), 'CSV must contain customer national ID');
console.log('✓ Test 1.3 Passed: Invoice metadata and buyer details verified in CSV.');

// 1.4 Tabular pipe items in CSV
assert(csvResult.includes('لوله پلی‌اتیلن 110 میلیمتر SDR11 PN16 PE100'), 'CSV must contain pipe item description');
assert(csvResult.includes('2.380'), 'CSV must contain pipe weight per meter');
assert(csvResult.includes('238.000'), 'CSV must contain pipe total weight');
assert(csvResult.includes('22610000'), 'CSV must contain pipe total price');
console.log('✓ Test 1.4 Passed: Pipe items table and row data verified in CSV.');

// 1.5 Tabular fitting items in CSV
assert(csvResult.includes('فلنج پایه بلند'), 'CSV must contain fitting product name');
assert(csvResult.includes('اتصالات جوشی فشار قوی'), 'CSV must contain fitting category');
assert(csvResult.includes('2400000'), 'CSV must contain fitting total price');
console.log('✓ Test 1.5 Passed: Fitting items table and row data verified in CSV.');

// 1.6 Financial summary & totals in CSV
assert(csvResult.includes('25010000'), 'CSV must contain subtotal');
assert(csvResult.includes('-1250500'), 'CSV must contain discount amount');
assert(csvResult.includes('+2375950'), 'CSV must contain tax amount');
assert(csvResult.includes('26135450'), 'CSV must contain total amount');
assert(csvResult.includes(numberToPersianWords(mockInvoice.totalAmount)), 'CSV must contain total amount in words');
console.log('✓ Test 1.6 Passed: Financial summary and Persian words verified in CSV.');

// 1.7 Bank payment details in CSV
assert(csvResult.includes(mockSettings.bankName!), 'CSV must contain bank name');
assert(csvResult.includes(mockSettings.bankAccount!), 'CSV must contain bank account number');
assert(csvResult.includes(mockSettings.bankIban!), 'CSV must contain bank IBAN');
console.log('✓ Test 1.7 Passed: Official bank payment settlement details verified in CSV.');

// -------------------------------------------------------------
// Test Suite 2: WhatsApp / SMS Persian Summary Generator
// -------------------------------------------------------------
console.log('--- Test Suite 2: WhatsApp / SMS Persian Summary Generator ---');

const waResult = generateInvoiceWhatsAppSummary(mockInvoice, mockSettings);

// 2.1 Header and official title
assert(waResult.includes('صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)'), 'WhatsApp summary must display official title when isOfficial is true');
assert(waResult.includes(toPersianDigits(mockInvoice.invoiceNumber)), 'WhatsApp summary must contain invoice number in Persian digits');
assert(waResult.includes(toPersianDigits(mockInvoice.jalaliDate)), 'WhatsApp summary must contain date in Persian digits');
assert(waResult.includes(mockInvoice.customerName), 'WhatsApp summary must contain customer name');
console.log('✓ Test 2.1 Passed: WhatsApp summary header, numbers, and customer verified.');

// 2.2 Items listing in WhatsApp summary
assert(waResult.includes('لوله'), 'WhatsApp summary must list pipe items');
assert(waResult.includes('فلنج پایه بلند'), 'WhatsApp summary must list fitting items');
assert(waResult.includes('🔹 اقلام لوله پلی‌اتیلن:'), 'WhatsApp summary must have pipe section header');
assert(waResult.includes('🔸 اقلام اتصالات:'), 'WhatsApp summary must have fitting section header');
console.log('✓ Test 2.2 Passed: Pipe and fitting line items verified in WhatsApp summary.');

// 2.3 Financial totals and Persian words in WhatsApp summary
assert(
  waResult.includes(toPersianDigits('26,135,450')) || waResult.includes('۲۶,۱۳۵,۴۵۰'),
  'WhatsApp summary must contain total formatted price'
);
assert(waResult.includes(numberToPersianWords(mockInvoice.totalAmount)), 'WhatsApp summary must contain total in Persian words');
assert(waResult.includes('تومان'), 'WhatsApp summary must contain currency unit');
console.log('✓ Test 2.3 Passed: Financial totals and Persian words verified in WhatsApp summary.');

// 2.4 Banking and company contacts
assert(waResult.includes(mockSettings.bankName!), 'WhatsApp summary must contain bank name');
assert(waResult.includes(mockSettings.bankIban!), 'WhatsApp summary must contain IBAN');
assert(waResult.includes(toPersianDigits(mockSettings.phone)), 'WhatsApp summary must contain company phone');
assert(waResult.includes('اعتبار پیش‌فاکتور: ۳ روز کاری'), 'WhatsApp summary must contain validity period');
console.log('✓ Test 2.4 Passed: Banking info and company contact verified in WhatsApp summary.');

// -------------------------------------------------------------
// Test Suite 3: Official vs Simple Proforma Mode
// -------------------------------------------------------------
console.log('--- Test Suite 3: Official vs Simple Proforma Modes ---');

const simpleInvoice: Invoice = {
  ...mockInvoice,
  isOfficial: false,
};

const simpleCsv = generateInvoiceCsvString(simpleInvoice, mockSettings);
const simpleWa = generateInvoiceWhatsAppSummary(simpleInvoice, mockSettings);

// Simple mode checks
assert(simpleCsv.includes('پیش‌فاکتور فروش کالا و خدمات'), 'Simple CSV must have standard proforma header');
assert(!simpleCsv.includes('کد اقتصادی فروشنده'), 'Simple CSV should not include official tax declaration header');
assert(simpleWa.includes(`پیش‌فاکتور شرکت ${mockSettings.name}`), 'Simple WhatsApp summary must have proforma company header');
assert(!simpleWa.includes('ماده ۱۶۹ مکرر'), 'Simple WhatsApp summary should not reference ماده ۱۶۹');
console.log('✓ Test 3 Passed: Mode toggling between Simple and Official verified in CSV and WhatsApp outputs.');

// -------------------------------------------------------------
// Test Suite 4: Edge Cases & RFC 4180 Escaping
// -------------------------------------------------------------
console.log('--- Test Suite 4: Edge Cases & Special Characters Escaping ---');

const trickyInvoice: Invoice = {
  ...mockInvoice,
  customerName: 'شرکت "گسترش آب"، شعبه ۱',
  customerAddress: 'تهران، خیابان ولیعصر\nکوچه دوم، پلاک ۱۰',
  explanation: 'توضیحات با "نقل قول" و کاما، خط اول\nخط دوم',
  items: [],
  subtotal: 0,
  includeDiscount: false,
  discountAmount: 0,
  includeTax: false,
  taxAmount: 0,
  totalAmount: 0,
};

const trickyCsv = generateInvoiceCsvString(trickyInvoice, mockSettings);

// RFC 4180 requires double quotes to be escaped as "" and fields with quotes/commas/newlines to be wrapped in quotes
assert(trickyCsv.includes('""گسترش آب""'), 'Quotes inside customer name must be doubled for CSV');
assert(trickyCsv.includes('"تهران، خیابان ولیعصر\nکوچه دوم، پلاک ۱۰"'), 'Newlines inside address must be preserved inside quoted CSV field');
assert(trickyCsv.includes('""نقل قول""'), 'Quotes inside explanation must be doubled');

const trickyWa = generateInvoiceWhatsAppSummary(trickyInvoice, mockSettings);
assert(trickyWa.includes('شرکت "گسترش آب"، شعبه ۱'), 'WhatsApp summary handles quotes and special characters without corruption');
assert(trickyWa.includes('صفر تومان'), 'Zero total converts correctly to Persian words');
console.log('✓ Test 4 Passed: RFC 4180 escaping and edge cases verified.');

// -------------------------------------------------------------
// Test Suite 5: Currency Consistency (Rial vs Toman)
// -------------------------------------------------------------
console.log('--- Test Suite 5: Currency Consistency (Rial vs Toman) ---');

const rialInvoice: Invoice = {
  ...mockInvoice,
  currencyUnit: 'ریال',
  totalAmount: 261354500,
};

const rialCsv = generateInvoiceCsvString(rialInvoice, mockSettings);
assert(rialCsv.includes('ریال'), 'CSV must respect invoice currencyUnit (ریال)');

const rialWa = generateInvoiceWhatsAppSummary(rialInvoice, mockSettings);
assert(rialWa.includes('ریال'), 'WhatsApp summary must respect invoice currencyUnit (ریال)');
assert(rialWa.includes('دویست و شصت و یک میلیون و سیصد و پنجاه و چهار هزار و پانصد ریال'), 'WhatsApp summary words must reflect Rial amount');
console.log('✓ Test 5 Passed: Currency consistency strictly maintained across export formats.');

// -------------------------------------------------------------
// Test Suite 6: Adversarial High-Load, Extreme Figures & Null Robustness
// -------------------------------------------------------------
console.log('--- Test Suite 6: Adversarial High-Load, Extreme Figures & Null Robustness ---');

// 6.1 Empty metadata invoice
const bareMinimumInvoice: Invoice = {
  id: 'inv_bare',
  invoiceNumber: '0',
  customerName: 'مشتری بدون مشخصات',
  jalaliDate: '1405/01/01',
  items: [],
  subtotal: 0,
  includeDiscount: false,
  discountPercentage: 0,
  discountAmount: 0,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 0,
  taxAmount: 0,
  totalAmount: 0,
  status: 'draft',
  createdAt: '2026-10-16T00:00:00Z',
};

const bareSettings: CompanySettings = {
  name: 'شرکت پایه',
  address: 'تهران',
  phone: '02100000000',
  fax: '',
  taxRate: 0,
  invoiceCounter: 1,
  currencyUnit: 'تومان',
  basePipePrices: { PE100: 95000, PE80: 88000, PE63: 82000 },
};

const bareCsv = generateInvoiceCsvString(bareMinimumInvoice, bareSettings);
assert(bareCsv.startsWith('\uFEFF'), 'Bare CSV must have BOM');
assert(bareCsv.includes('مشتری بدون مشخصات'), 'Bare CSV must have customer name');

const bareWa = generateInvoiceWhatsAppSummary(bareMinimumInvoice, bareSettings);
assert(bareWa.includes('مشتری بدون مشخصات'), 'Bare WhatsApp must have customer name');
console.log('✓ Test 6.1 Passed: Bare-minimum invoice with zero items and null fields handled gracefully.');

// 6.2 High-load invoice with 500 items
const bigItems: (PipeItem | FittingItem)[] = [];
for (let i = 0; i < 250; i++) {
  bigItems.push({
    id: `bulk_pipe_${i}`,
    type: 'pipe',
    peGrade: i % 2 === 0 ? 'PE100' : 'PE80',
    diameter: 63 + (i % 5) * 20,
    sdr: 11,
    pn: 16,
    length: 12,
    weightPerMeter: 1.25,
    totalWeight: 15.0,
    pricePerKg: 95000,
    totalPrice: 1425000,
  });
  bigItems.push({
    id: `bulk_fit_${i}`,
    type: 'fitting',
    category: 'اتصالات جوشی',
    connectionType: 'جوشی',
    product: `زانو ۹۰ درجه ${63 + (i % 5) * 20} میلیمتر`,
    material: 'PE100',
    pressure: '16',
    size: String(63 + (i % 5) * 20),
    quantity: 10,
    unitPrice: 150000,
    totalPrice: 1500000,
  });
}

const largeInvoice: Invoice = {
  ...mockInvoice,
  items: bigItems,
  subtotal: 500 * 1425000,
  totalAmount: 500 * 1425000,
};

const startTime = Date.now();
const bigCsv = generateInvoiceCsvString(largeInvoice, mockSettings);
const bigWa = generateInvoiceWhatsAppSummary(largeInvoice, mockSettings);
const elapsedMs = Date.now() - startTime;

assert(bigCsv.length > 50000, 'Large CSV must contain extensive data');
assert(bigWa.length > 20000, 'Large WhatsApp summary must contain extensive lines');
assert(elapsedMs < 1000, `High load processing should complete in under 1000ms (took ${elapsedMs}ms)`);
console.log(`✓ Test 6.2 Passed: 500-item invoice generated in ${elapsedMs}ms without performance degradation.`);

// 6.3 Trillion figure conversion
const trillionInvoice: Invoice = {
  ...mockInvoice,
  currencyUnit: 'ریال',
  totalAmount: 1500000000000, // 1.5 Trillion Rials
};
const trillionWa = generateInvoiceWhatsAppSummary(trillionInvoice, mockSettings);
assert(trillionWa.includes('یک تریلیون و پانصد میلیارد ریال'), '1.5 Trillion Rial figure converts correctly to words');
console.log('✓ Test 6.3 Passed: Trillion figures scale cleanly in Persian words.');

console.log('======================================================');
console.log('ALL ADVANCED PRINT, EXPORT & SHARING TESTS PASSED! (6/6)');
console.log('======================================================');
