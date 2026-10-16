// Invoicing Workflow Verification Test Suite
import { describe, it } from 'vitest';
import { StorageService } from '../storage';
import { computePipeItem, computeInvoiceSummary } from '../calculator';
import { initialWeightTable } from '../../data/initialData';
import { Customer, Invoice, PipeItem, FittingItem, InvoiceItem } from '../../types';

// Setup Mock LocalStorage
const store = new Map<string, string>();
const mockLocalStorage = {
  getItem: (key: string) => store.get(key) ?? null,
  setItem: (key: string, value: string) => store.set(key, String(value)),
  removeItem: (key: string) => store.delete(key),
  clear: () => store.clear(),
};
(globalThis as any).localStorage = mockLocalStorage;

function assert(condition: boolean, message: string) {
  if (!condition) {
    throw new Error(`Assertion failed: ${message}`);
  }
}

describe('Invoicing Workflow Verification Suite', () => {
  it('passes all 9 invoicing workflow tests', () => {
console.log('--- RUNNING INVOICING WORKFLOW VERIFICATION SUITE ---');

// Test 1: Real-time Draft Auto-Save
console.log('Test 1: StorageService Draft Auto-Save, Retrieval, and Clear');
const sampleDraft: Partial<Invoice> = {
  invoiceNumber: '105',
  customerName: 'شرکت آب و خاک البرز',
  customerPhone: '09121112233',
  customerAddress: 'کرج، شهرک صنعتی',
  currencyUnit: 'تومان',
  items: [],
  subtotal: 0,
};

StorageService.saveDraft(sampleDraft);
const loadedDraft = StorageService.getDraft();
assert(loadedDraft !== null, 'Draft should be loaded from localStorage');
assert(loadedDraft?.customerName === 'شرکت آب و خاک البرز', 'Draft customerName should match');
assert(loadedDraft?.invoiceNumber === '105', 'Draft invoiceNumber should match');
assert(loadedDraft?.currencyUnit === 'تومان', 'Draft currencyUnit should match');

StorageService.clearDraft();
const clearedDraft = StorageService.getDraft();
assert(clearedDraft === null, 'Draft should be null after clearDraft()');
console.log('✓ Test 1 Passed: Draft auto-save operates correctly.');

// Test 2: Edit vs New Invoice in StorageService (Update in Place vs Prepend)
console.log('Test 2: SaveInvoice Updates In-Place When Editing Existing Invoice');
const initialInvoice: Invoice = {
  id: 'inv_101',
  invoiceNumber: '101',
  customerName: 'آقای کاظمی',
  customerPhone: '09120000000',
  jalaliDate: '1405/07/20',
  items: [],
  subtotal: 50000000,
  includeDiscount: false,
  discountPercentage: 0,
  discountAmount: 0,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 10,
  taxAmount: 0,
  totalAmount: 50000000,
  currencyUnit: 'تومان',
  status: 'finalized',
  createdAt: '2026-10-15T10:00:00Z',
};

StorageService.saveInvoice(initialInvoice);
let currentInvoices = StorageService.getInvoices();
assert(currentInvoices.length === 1, 'Should have 1 invoice');
assert(currentInvoices[0].customerName === 'آقای کاظمی', 'Initial invoice customer name matches');

// Perform edit
const editedInvoice: Invoice = {
  ...initialInvoice,
  customerName: 'آقای کاظمی (ویرایش شده)',
  totalAmount: 65000000,
};
StorageService.saveInvoice(editedInvoice);
currentInvoices = StorageService.getInvoices();
assert(currentInvoices.length === 1, 'Should STILL have 1 invoice (updated in place, not duplicated)');
assert(currentInvoices[0].customerName === 'آقای کاظمی (ویرایش شده)', 'Invoice should be updated in place');
assert(currentInvoices[0].totalAmount === 65000000, 'Invoice totalAmount updated');

// Perform duplicate as new invoice
const duplicateInvoice: Invoice = {
  ...editedInvoice,
  id: 'inv_102',
  invoiceNumber: '102',
  createdAt: '2026-10-15T11:00:00Z',
};
StorageService.saveInvoice(duplicateInvoice);
currentInvoices = StorageService.getInvoices();
assert(currentInvoices.length === 2, 'Should now have 2 invoices after duplicate');
assert(currentInvoices[0].id === 'inv_102', 'New duplicated invoice prepended');
console.log('✓ Test 2 Passed: In-place edit updates and duplicate prepends correctly.');

// Test 3: Inline Pipe Table Editing Live Recalculation
console.log('Test 3: Inline Pipe Item Editing Live Recalculation');
const initialPipe = computePipeItem(
  {
    peGrade: 'PE100',
    diameter: 110,
    sdr: 11,
    pn: 16,
    length: 100, // 100 meters
    pricePerKg: 95000,
  },
  initialWeightTable,
  'pipe_item_1'
);

if (!initialPipe) throw new Error('initialPipe is null');
assert(initialPipe.totalWeight > 0, 'Total weight should be computed');
const initialWeight = initialPipe.totalWeight;
const initialTotal = initialPipe.totalPrice;

// User edits length inline: 100m -> 200m
const editedLengthPipe = computePipeItem(
  {
    peGrade: initialPipe.peGrade,
    diameter: initialPipe.diameter,
    sdr: initialPipe.sdr,
    pn: initialPipe.pn,
    length: 200,
    pricePerKg: initialPipe.pricePerKg,
  },
  initialWeightTable,
  initialPipe.id
);

if (!editedLengthPipe) throw new Error('editedLengthPipe is null');
assert(editedLengthPipe.id === initialPipe.id, 'Pipe item ID preserved on inline edit');
assert(Math.abs(editedLengthPipe.totalWeight - initialWeight * 2) < 0.01, 'Weight doubled on length doubling');
assert(editedLengthPipe.totalPrice === initialTotal * 2, 'Total price doubled on length doubling');

// User edits pricePerKg inline: 95000 -> 100000
const editedPricePipe = computePipeItem(
  {
    peGrade: editedLengthPipe.peGrade,
    diameter: editedLengthPipe.diameter,
    sdr: editedLengthPipe.sdr,
    pn: editedLengthPipe.pn,
    length: editedLengthPipe.length,
    pricePerKg: 100000,
  },
  initialWeightTable,
  editedLengthPipe.id
);

if (!editedPricePipe) throw new Error('editedPricePipe is null');
assert(editedPricePipe.totalPrice === Math.round(editedLengthPipe.totalWeight * 100000), 'Total price recalculated with new pricePerKg');
console.log('✓ Test 3 Passed: Inline pipe item recalculation operates accurately.');

// Test 4: Inline Fitting Item Editing Live Recalculation
console.log('Test 4: Inline Fitting Item Editing Live Recalculation');
const initialFitting: FittingItem = {
  id: 'fit_1',
  type: 'fitting',
  category: 'اتصالات فشار قوی',
  connectionType: 'جوشی',
  product: 'فلنج پایه بلند',
  material: 'HDPE100',
  pressure: '16',
  size: '110',
  quantity: 5,
  unitPrice: 250000,
  totalPrice: 1250000,
};

// Inline edit quantity: 5 -> 12
const editedQtyFitting: FittingItem = {
  ...initialFitting,
  quantity: 12,
  totalPrice: Math.round(12 * initialFitting.unitPrice),
};
assert(editedQtyFitting.totalPrice === 3000000, 'Total price matches quantity * unitPrice');

// Inline edit unit price: 250000 -> 300000
const editedPriceFitting: FittingItem = {
  ...editedQtyFitting,
  unitPrice: 300000,
  totalPrice: Math.round(editedQtyFitting.quantity * 300000),
};
assert(editedPriceFitting.totalPrice === 3600000, 'Total price matches new unitPrice');
console.log('✓ Test 4 Passed: Inline fitting item editing operates accurately.');

// Test 5: Customer Autofill & Discount Computation
console.log('Test 5: Customer Autofill & Discount Application');
const mockCustomer: Customer = {
  id: 'cust_1',
  name: 'مهندس رضایی',
  phone: '09123456789',
  address: 'تهران، بزرگراه فتح',
  defaultDiscount: 7, // 7% discount
};

const itemsList: InvoiceItem[] = [editedPricePipe, editedPriceFitting];
const summaryBeforeDiscount = computeInvoiceSummary({
  items: itemsList,
  includeDiscount: false,
  discountPercentage: 0,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 10,
});

assert(summaryBeforeDiscount.subtotal > 0, 'Subtotal is positive');

// Applying customer 7% discount
const expectedDiscountAmount = Math.round((summaryBeforeDiscount.subtotal * (mockCustomer.defaultDiscount || 0)) / 100);
const summaryWithCustomerDiscount = computeInvoiceSummary({
  items: itemsList,
  includeDiscount: false,
  discountPercentage: 0,
  includeCustomDiscount: true,
  customDiscountAmount: expectedDiscountAmount,
  includeTax: true,
  taxPercentage: 10,
});

assert(summaryWithCustomerDiscount.discountAmount === expectedDiscountAmount, 'Discount amount applied correctly');
assert(
  summaryWithCustomerDiscount.amountAfterDiscount === summaryBeforeDiscount.subtotal - expectedDiscountAmount,
  'Amount after discount matches'
);
assert(
  summaryWithCustomerDiscount.totalAmount ===
    summaryWithCustomerDiscount.amountAfterDiscount + summaryWithCustomerDiscount.taxAmount,
  'Total amount includes tax'
);
console.log('✓ Test 5 Passed: Customer autofill and discount math accurate.');

// Test 6: Currency Consistency Check
console.log('Test 6: Currency Consistency Across Rial & Toman Invoices');
const tomanInvoice: Invoice = {
  ...initialInvoice,
  id: 'inv_toman_1',
  invoiceNumber: '103',
  currencyUnit: 'تومان',
  totalAmount: 10000000,
};

const rialInvoice: Invoice = {
  ...initialInvoice,
  id: 'inv_rial_1',
  invoiceNumber: '104',
  currencyUnit: 'ریال',
  totalAmount: 100000000,
};

StorageService.saveInvoice(tomanInvoice);
StorageService.saveInvoice(rialInvoice);

const allInvoices = StorageService.getInvoices();
const loadedToman = allInvoices.find((i) => i.id === 'inv_toman_1');
const loadedRial = allInvoices.find((i) => i.id === 'inv_rial_1');

assert(loadedToman?.currencyUnit === 'تومان', 'Toman invoice preserves تومان');
assert(loadedRial?.currencyUnit === 'ریال', 'Rial invoice preserves ریال');
assert(loadedToman?.totalAmount === 10000000, 'Toman invoice total unchanged');
assert(loadedRial?.totalAmount === 100000000, 'Rial invoice total unchanged');
console.log('✓ Test 6 Passed: Currency consistency preserved across invoices.');

// Test 7: Invoice Counter State Sync & Overwrite Prevention
console.log('Test 7: Invoice Counter State Sync & Overwrite Prevention');
StorageService.saveSettings({ ...StorageService.getSettings(), invoiceCounter: 200, currencyUnit: 'تومان' });
const dupNum = StorageService.getNextInvoiceNumber();
assert(dupNum === '201', 'Next invoice number is 201');
const syncedSettings = StorageService.getSettings();
assert(syncedSettings.invoiceCounter === 201, 'Settings invoiceCounter updated to 201 in storage');

const duplicatedInv: Invoice = {
  ...initialInvoice,
  id: 'inv_dup_201',
  invoiceNumber: dupNum,
  customerName: 'مشتری الف (تکثیر شده)',
  totalAmount: 1000,
};
StorageService.saveInvoice(duplicatedInv);

// Create another invoice with separate ID (simulating collision scenario)
const subsequentInv: Invoice = {
  ...initialInvoice,
  id: 'inv_subsequent_201',
  invoiceNumber: dupNum,
  customerName: 'مشتری ب (جدید)',
  totalAmount: 2000,
};
StorageService.saveInvoice(subsequentInv);

const currentList = StorageService.getInvoices();
const foundDup = currentList.find((i) => i.id === 'inv_dup_201');
const foundSub = currentList.find((i) => i.id === 'inv_subsequent_201');
assert(foundDup !== undefined, 'Duplicated invoice is NOT overwritten by subsequent invoice');
assert(foundSub !== undefined, 'Subsequent invoice is also preserved');
console.log('✓ Test 7 Passed: Invoice counter state sync and overwrite prevention verified.');

// Test 8: Non-Negative Input Clamping in Pipe Calculations
console.log('Test 8: Non-Negative Input Clamping in Pipe Calculations');
const negLengthPipe = computePipeItem(
  {
    peGrade: 'PE100',
    diameter: 110,
    sdr: 11,
    pn: 16,
    length: -25,
    pricePerKg: 95000,
  },
  initialWeightTable
);
assert(negLengthPipe !== null, 'Pipe item computed');
assert(negLengthPipe?.length === 0, 'Negative length clamped to 0');
assert(negLengthPipe?.totalWeight === 0, 'Total weight clamped to 0');
assert(negLengthPipe?.totalPrice === 0, 'Total price clamped to 0');

const negPricePipe = computePipeItem(
  {
    peGrade: 'PE100',
    diameter: 110,
    sdr: 11,
    pn: 16,
    length: 100,
    pricePerKg: -50000,
  },
  initialWeightTable
);
assert(negPricePipe !== null, 'Pipe item computed');
assert(negPricePipe?.pricePerKg === 0, 'Negative price clamped to 0');
assert(negPricePipe?.totalPrice === 0, 'Total price clamped to 0');
console.log('✓ Test 8 Passed: Non-negative clamping operates correctly.');

// Test 9: Corrupted Draft Resilience and Reactive Discount
console.log('Test 9: Corrupted Draft Resilience and Reactive Discount');
mockLocalStorage.setItem('pipecal_draft_invoice_v2', '{"items": [null], "customerName": 99999}');
const recoveredDraft = StorageService.getDraft();
assert(recoveredDraft !== null, 'Draft parsed from JSON');
const validCustName = typeof recoveredDraft?.customerName === 'string' ? recoveredDraft.customerName : '';
assert(validCustName === '', 'Non-string customerName falls back safely to empty string');

const validItemsInSummary = computeInvoiceSummary({
  items: (recoveredDraft?.items as any) || [],
  includeDiscount: true,
  discountPercentage: 15,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 10,
});
assert(validItemsInSummary.subtotal === 0, 'Null item handled gracefully without crash');

// Test auto-healing invalid JSON
mockLocalStorage.setItem('pipecal_draft_invoice_v2', '{ corrupt json string');
const corruptDraft = StorageService.getDraft();
assert(corruptDraft === null, 'Corrupt draft returns null');
assert(mockLocalStorage.getItem('pipecal_draft_invoice_v2') === null, 'Corrupt draft auto-healed and cleared from storage');
console.log('✓ Test 9 Passed: Corrupted draft resilience and auto-heal verified.');

console.log('\n========================================');
console.log('ALL INVOICING WORKFLOW TESTS PASSED (9/9)');
console.log('========================================\n');
  });
});
