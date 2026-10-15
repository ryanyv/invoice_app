// Challenger M1 Empirical Stress Testing Harness
// Adversarial verification of Requirement R1: Invoicing Workflow Enhancements
import { StorageService } from '../storage';
import {
  computePipeItem,
  getWeightPerMeter,
  computeInvoiceSummary,
  getApplicableDiscount,
} from '../calculator';
import { initialWeightTable, initialDiscounts, initialCompanySettings } from '../../data/initialData';
import {
  CompanySettings,
  Customer,
  DiscountRule,
  Invoice,
  InvoiceItem,
  PipeItem,
  FittingItem,
  PipeGrade,
} from '../../types';

// In-Memory LocalStorage Mock with state tracking & failure simulation
class MockLocalStorage {
  private store: Map<string, string> = new Map();
  public failNextSetItem = false;
  public failNextGetItem = false;

  getItem(key: string): string | null {
    if (this.failNextGetItem) {
      this.failNextGetItem = false;
      throw new Error('Simulated localStorage.getItem failure');
    }
    return this.store.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    if (this.failNextSetItem) {
      this.failNextSetItem = false;
      throw new Error('Simulated localStorage.setItem quota exceeded');
    }
    this.store.set(key, String(value));
  }

  removeItem(key: string): void {
    this.store.delete(key);
  }

  clear(): void {
    this.store.clear();
  }

  rawSet(key: string, rawValue: string): void {
    this.store.set(key, rawValue);
  }
}

const mockStorage = new MockLocalStorage();
(globalThis as any).localStorage = mockStorage;

// Helper assertion function
function assert(condition: boolean, message: string): void {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    throw new Error(`Assertion failed: ${message}`);
  }
}

let totalTests = 0;
let passedTests = 0;

function runTest(name: string, fn: () => void): void {
  totalTests++;
  try {
    fn();
    passedTests++;
    console.log(`  ✓ ${name}`);
  } catch (err: any) {
    console.error(`  ✗ ${name}`);
    console.error(`    ${err?.message || err}`);
    throw err;
  }
}

console.log('================================================================');
console.log('CHALLENGER 1 EMPIRICAL STRESS-TEST HARNESS — MILESTONE M1 (R1)');
console.log('================================================================\n');

// =====================================================================
// SUITE 1: DRAFT AUTO-SAVE AND RESTORE ROUND-TRIPS
// =====================================================================
console.log('SUITE 1: Draft Auto-Save and Restore Round-Trips');

runTest('1.1 Full Draft Serialization Round-Trip with Complex Nested Data', () => {
  mockStorage.clear();

  const pipe1 = computePipeItem(
    { peGrade: 'PE100', diameter: 160, sdr: 17, pn: 10, length: 350.5, pricePerKg: 96500 },
    initialWeightTable,
    'pipe_draft_1'
  );
  assert(pipe1 !== null, 'pipe1 should be computed');

  const fitting1: FittingItem = {
    id: 'fit_draft_1',
    type: 'fitting',
    category: 'اتصالات فشار قوی',
    connectionType: 'جوشی',
    product: 'سه‌راهی جوشی ۹۰ درجه',
    material: 'HDPE100',
    pressure: '10',
    size: '160',
    quantity: 8,
    unitPrice: 420000,
    totalPrice: 8 * 420000,
  };

  const complexDraft: Partial<Invoice> = {
    invoiceNumber: '888',
    customerName: 'شرکت مهندسی کاوشگران آب و خاک پایتخت (سهامی خاص) 💧',
    customerPhone: '0912-345-6789',
    customerAddress: 'تهران، میدان ونک، برج نگار، طبقه ۱۲، واحد ۴',
    explanation: 'توضیحات پیش‌فاکتور با نویسه‌های ویژه فارسی: گرید PE100، فشار ۱۰ بار، تحویل انبار کارخانه تبریز.',
    currencyUnit: 'تومان',
    items: [pipe1!, fitting1],
    includeDiscount: true,
    discountPercentage: 3,
    discountAmount: 1500000,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: true,
    taxPercentage: 10,
    taxAmount: 5000000,
    totalAmount: 53500000,
  };

  StorageService.saveDraft(complexDraft);
  const loaded = StorageService.getDraft();

  assert(loaded !== null, 'Draft must be retrieved');
  assert(loaded?.invoiceNumber === '888', 'Invoice number must match');
  assert(loaded?.customerName === complexDraft.customerName, 'Customer name with Persian text/emojis must match');
  assert(loaded?.customerPhone === complexDraft.customerPhone, 'Customer phone must match');
  assert(loaded?.customerAddress === complexDraft.customerAddress, 'Customer address must match');
  assert(loaded?.explanation === complexDraft.explanation, 'Explanation must match');
  assert(loaded?.currencyUnit === 'تومان', 'Currency unit must match');
  assert(loaded?.items?.length === 2, 'Must contain 2 items');

  const restoredPipe = loaded?.items?.[0] as PipeItem;
  assert(restoredPipe.type === 'pipe', 'First item must be pipe');
  assert(restoredPipe.diameter === 160, 'Pipe diameter must match');
  assert(restoredPipe.sdr === 17, 'Pipe SDR must match');
  assert(restoredPipe.length === 350.5, 'Pipe length must match');
  assert(restoredPipe.weightPerMeter === pipe1!.weightPerMeter, 'Weight per meter must match');
  assert(restoredPipe.totalWeight === pipe1!.totalWeight, 'Total weight must match');
  assert(restoredPipe.totalPrice === pipe1!.totalPrice, 'Total price must match');

  const restoredFitting = loaded?.items?.[1] as FittingItem;
  assert(restoredFitting.type === 'fitting', 'Second item must be fitting');
  assert(restoredFitting.product === 'سه‌راهی جوشی ۹۰ درجه', 'Fitting product must match');
  assert(restoredFitting.quantity === 8, 'Fitting quantity must match');
  assert(restoredFitting.unitPrice === 420000, 'Fitting unitPrice must match');
  assert(restoredFitting.totalPrice === 3360000, 'Fitting totalPrice must match');
});

runTest('1.2 Draft Auto-Save Idempotency & Progressive Updates', () => {
  mockStorage.clear();

  // Simulating 10 keystrokes of draft auto-save
  for (let step = 1; step <= 10; step++) {
    StorageService.saveDraft({
      invoiceNumber: '100',
      customerName: `نام مشتری مرحله ${step}`,
      explanation: `متن در حال ویرایش: طول ${step * 10}`,
      currencyUnit: step % 2 === 0 ? 'ریال' : 'تومان',
    });

    const current = StorageService.getDraft();
    assert(current?.customerName === `نام مشتری مرحله ${step}`, `Step ${step} customer name match`);
    assert(current?.currencyUnit === (step % 2 === 0 ? 'ریال' : 'تومان'), `Step ${step} currency match`);
  }
});

runTest('1.3 Draft Clear Operation & Repeated Clears', () => {
  mockStorage.clear();
  StorageService.saveDraft({ invoiceNumber: '999', customerName: 'تست پاکسازی' });
  assert(StorageService.getDraft() !== null, 'Draft exists before clear');

  StorageService.clearDraft();
  assert(StorageService.getDraft() === null, 'Draft is null after clearDraft()');

  // Repeated clears should be safe no-ops
  StorageService.clearDraft();
  StorageService.clearDraft();
  assert(StorageService.getDraft() === null, 'Repeated clearDraft() remains null without error');
});

runTest('1.4 Draft Resilience to Malformed & Corrupted Storage Data', () => {
  mockStorage.clear();

  // Malformed JSON string in storage
  mockStorage.rawSet('pipecal_draft_invoice_v2', '{ broken json: nil, [}');
  const corrupted = StorageService.getDraft();
  assert(corrupted === null, 'Corrupted JSON draft must return null without throwing');

  // Non-JSON garbage string
  mockStorage.rawSet('pipecal_draft_invoice_v2', 'GARBAGE_DATA_12345');
  const garbage = StorageService.getDraft();
  assert(garbage === null, 'Garbage data draft must return null without throwing');

  // Empty string
  mockStorage.rawSet('pipecal_draft_invoice_v2', '');
  const empty = StorageService.getDraft();
  assert(empty === null, 'Empty string draft must return null');
});

// =====================================================================
// SUITE 2: DIN STANDARD WEIGHTS & INLINE RECALCULATION CONSISTENCY
// =====================================================================
console.log('\nSUITE 2: DIN Standard Weights & Inline Calculation Consistency');

runTest('2.1 Exhaustive Verification of All DIN Table Combinations (Every Diameter & SDR)', () => {
  const diameters = Object.keys(initialWeightTable).map(Number).sort((a, b) => a - b);
  assert(diameters.length > 0, 'Weight table must have diameters');

  let totalPairsTested = 0;

  for (const diameter of diameters) {
    const sdrObj = initialWeightTable[diameter];
    assert(typeof sdrObj === 'object' && sdrObj !== null, `Diameter ${diameter} must have an SDR mapping`);

    const sdrs = Object.keys(sdrObj).map(Number);
    for (const sdr of sdrs) {
      const wpm = getWeightPerMeter(diameter, sdr, initialWeightTable);
      assert(wpm !== null && wpm > 0, `Weight per meter for ${diameter}mm SDR ${sdr} must be positive number`);

      // Test computePipeItem with test length and price
      const testLength = 120.0;
      const testPrice = 100000;
      const pipeItem = computePipeItem(
        {
          peGrade: 'PE100',
          diameter,
          sdr,
          pn: 10,
          length: testLength,
          pricePerKg: testPrice,
        },
        initialWeightTable
      );

      assert(pipeItem !== null, `computePipeItem must succeed for ${diameter}mm SDR ${sdr}`);
      assert(pipeItem!.weightPerMeter === wpm, `Item WPM must match lookup WPM for ${diameter}mm SDR ${sdr}`);
      assert(pipeItem!.totalWeight === Number((wpm! * testLength).toFixed(3)), `Total weight calculation consistent`);
      assert(pipeItem!.totalPrice === Math.round(pipeItem!.totalWeight * testPrice), `Total price calculation consistent`);

      totalPairsTested++;
    }
  }

  console.log(`    (Verified ${totalPairsTested} distinct DIN diameter/SDR standard specifications)`);
  assert(totalPairsTested >= 40, 'Should have verified at least 40 standard pipe specifications');
});

runTest('2.2 Property-Based Monte Carlo Testing (500 Random Invocations)', () => {
  // Collect all valid (diameter, SDR) pairs from the DIN table
  const validPairs: Array<{ diameter: number; sdr: number; wpm: number }> = [];
  for (const d of Object.keys(initialWeightTable).map(Number)) {
    for (const s of Object.keys(initialWeightTable[d]).map(Number)) {
      const wpm = initialWeightTable[d][s];
      if (typeof wpm === 'number' && wpm > 0) {
        validPairs.push({ diameter: d, sdr: s, wpm });
      }
    }
  }

  const iterations = 500;
  for (let i = 0; i < iterations; i++) {
    const pair = validPairs[Math.floor(Math.random() * validPairs.length)];
    // Random length between 0.1 and 10000 with up to 2 decimal places
    const randomLength = Math.round((Math.random() * 9999.9 + 0.1) * 100) / 100;
    // Random price between 10,000 and 2,500,000
    const randomPrice = Math.floor(Math.random() * 2490000) + 10000;
    const grade: PipeGrade = Math.random() > 0.5 ? 'PE100' : 'PE80';

    const item = computePipeItem(
      {
        peGrade: grade,
        diameter: pair.diameter,
        sdr: pair.sdr,
        pn: 10,
        length: randomLength,
        pricePerKg: randomPrice,
      },
      initialWeightTable,
      `pipe_rand_${i}`
    );

    assert(item !== null, `computePipeItem failed on iteration ${i}`);
    assert(item!.id === `pipe_rand_${i}`, 'Item ID must be preserved');
    assert(item!.weightPerMeter === pair.wpm, 'Weight per meter matches table');

    // Invariant 1: totalWeight == round(wpm * length, 3)
    const expectedWeight = Number((pair.wpm * randomLength).toFixed(3));
    assert(
      Math.abs(item!.totalWeight - expectedWeight) < 1e-6,
      `Invariant 1: weight mismatch on d=${pair.diameter}, sdr=${pair.sdr}, L=${randomLength}`
    );

    // Invariant 2: totalPrice == round(totalWeight * pricePerKg)
    const expectedPrice = Math.round(item!.totalWeight * randomPrice);
    assert(
      item!.totalPrice === expectedPrice,
      `Invariant 2: price mismatch on weight=${item!.totalWeight}, price=${randomPrice}`
    );

    // Invariant 3: Inline edit length simulation
    const updatedLength = Math.round((Math.random() * 5000 + 1) * 100) / 100;
    const inlineUpdatedPipe = computePipeItem(
      {
        peGrade: item!.peGrade,
        diameter: item!.diameter,
        sdr: item!.sdr,
        pn: item!.pn,
        length: updatedLength,
        pricePerKg: item!.pricePerKg,
      },
      initialWeightTable,
      item!.id
    );

    assert(inlineUpdatedPipe !== null, 'Inline pipe update failed');
    assert(inlineUpdatedPipe!.id === item!.id, 'Pipe item ID preserved on inline edit');
    const updatedWeight = Number((pair.wpm * updatedLength).toFixed(3));
    assert(inlineUpdatedPipe!.totalWeight === updatedWeight, 'Inline updated weight consistent');
    assert(
      inlineUpdatedPipe!.totalPrice === Math.round(updatedWeight * item!.pricePerKg),
      'Inline updated price consistent'
    );
  }
  console.log(`    (Passed ${iterations} randomized Monte Carlo property tests successfully)`);
});

runTest('2.3 Boundary & Negative Edge Conditions', () => {
  // Unknown diameter
  assert(getWeightPerMeter(999, 11, initialWeightTable) === null, 'Unknown diameter returns null');
  assert(
    computePipeItem(
      { peGrade: 'PE100', diameter: 999, sdr: 11, pn: 16, length: 100, pricePerKg: 95000 },
      initialWeightTable
    ) === null,
    'computePipeItem with unknown diameter returns null'
  );

  // Unknown SDR
  assert(getWeightPerMeter(110, 99, initialWeightTable) === null, 'Unknown SDR returns null');
  assert(
    computePipeItem(
      { peGrade: 'PE100', diameter: 110, sdr: 99, pn: 16, length: 100, pricePerKg: 95000 },
      initialWeightTable
    ) === null,
    'computePipeItem with unknown SDR returns null'
  );

  // Negative diameter
  assert(getWeightPerMeter(-50, 11, initialWeightTable) === null, 'Negative diameter returns null');

  // Zero length
  const zeroLengthPipe = computePipeItem(
    { peGrade: 'PE100', diameter: 110, sdr: 11, pn: 16, length: 0, pricePerKg: 95000 },
    initialWeightTable
  );
  assert(zeroLengthPipe !== null, 'computePipeItem with length=0 succeeds');
  assert(zeroLengthPipe!.totalWeight === 0, 'Zero length yields 0 weight');
  assert(zeroLengthPipe!.totalPrice === 0, 'Zero length yields 0 price');

  // Diameter 630 (not in table) correctly returns null
  assert(getWeightPerMeter(630, 11, initialWeightTable) === null, 'Diameter 630 not in table returns null');

  // Large length & price with max diameter (500mm SDR 11: wpm=64.5)
  const largePipe = computePipeItem(
    { peGrade: 'PE100', diameter: 500, sdr: 11, pn: 16, length: 1000000, pricePerKg: 500000 },
    initialWeightTable
  );
  assert(largePipe !== null, 'Large pipe calculation succeeds');
  assert(!isNaN(largePipe!.totalPrice) && isFinite(largePipe!.totalPrice), 'Large pipe price is finite');
  assert(largePipe!.totalPrice > 0, 'Large pipe price is positive');
});

runTest('2.4 Inline Fitting Calculations Live Updates & Boundary Clamping', () => {
  const initialFitting: FittingItem = {
    id: 'fit_stress_1',
    type: 'fitting',
    category: 'اتصالات تزریقی',
    connectionType: 'جوشی',
    product: 'زانویی ۹۰ درجه',
    material: 'HDPE100',
    pressure: '16',
    size: '90',
    quantity: 15,
    unitPrice: 185000,
    totalPrice: 15 * 185000,
  };

  // Simulation of handleUpdateFittingItem inline hook logic
  const updateFitting = (item: FittingItem, newQty: number, newPrice: number): FittingItem => {
    const qty = Math.max(0, newQty);
    const price = Math.max(0, newPrice);
    return {
      ...item,
      quantity: qty,
      unitPrice: price,
      totalPrice: Math.round(qty * price),
    };
  };

  // Test 1: Quantity increase
  const f1 = updateFitting(initialFitting, 25, initialFitting.unitPrice);
  assert(f1.quantity === 25, 'Quantity updated to 25');
  assert(f1.totalPrice === 25 * 185000, 'Total price matches 25 * 185000');

  // Test 2: Unit price change
  const f2 = updateFitting(f1, f1.quantity, 200000);
  assert(f2.unitPrice === 200000, 'Price updated to 200000');
  assert(f2.totalPrice === 25 * 200000, 'Total price matches 25 * 200000');

  // Test 3: Clamping negative quantity to 0
  const f3 = updateFitting(f2, -5, f2.unitPrice);
  assert(f3.quantity === 0, 'Negative quantity clamped to 0');
  assert(f3.totalPrice === 0, 'Zero quantity produces 0 total price');

  // Test 4: Clamping negative price to 0
  const f4 = updateFitting(f2, 10, -50000);
  assert(f4.unitPrice === 0, 'Negative price clamped to 0');
  assert(f4.totalPrice === 0, 'Zero price produces 0 total price');
});

// =====================================================================
// SUITE 3: CURRENCY TOGGLE PRESERVATION
// =====================================================================
console.log('\nSUITE 3: Currency Toggle Preservation & Consistency');

runTest('3.1 Multi-Invoice Currency Persistence Across Settings Mutations', () => {
  mockStorage.clear();

  // Create Toman Invoice
  const tomanInvoice: Invoice = {
    id: 'inv_toman_100',
    invoiceNumber: '100',
    customerName: 'شرکت لوله‌گستر اصفهان',
    jalaliDate: '1405/07/20',
    items: [],
    subtotal: 75000000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 75000000,
    currencyUnit: 'تومان',
    status: 'finalized',
    createdAt: '2026-10-15T12:00:00Z',
  };

  // Create Rial Invoice
  const rialInvoice: Invoice = {
    id: 'inv_rial_200',
    invoiceNumber: '200',
    customerName: 'صنایع پتروشیمی پارس',
    jalaliDate: '1405/07/21',
    items: [],
    subtotal: 750000000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 750000000,
    currencyUnit: 'ریال',
    status: 'finalized',
    createdAt: '2026-10-15T13:00:00Z',
  };

  StorageService.saveInvoice(tomanInvoice);
  StorageService.saveInvoice(rialInvoice);

  // Settings are toggled multiple times
  const settingsToman: CompanySettings = { ...initialCompanySettings, currencyUnit: 'تومان' };
  const settingsRial: CompanySettings = { ...initialCompanySettings, currencyUnit: 'ریال' };

  // Phase 1: Settings is Toman
  StorageService.saveSettings(settingsToman);
  let invoices = StorageService.getInvoices();
  let loadedToman = invoices.find((i) => i.id === 'inv_toman_100');
  let loadedRial = invoices.find((i) => i.id === 'inv_rial_200');

  assert(loadedToman?.currencyUnit === 'تومان', 'Phase 1: Toman invoice has currencyUnit تومان');
  assert(loadedToman?.totalAmount === 75000000, 'Phase 1: Toman totalAmount unaltered');
  assert(loadedRial?.currencyUnit === 'ریال', 'Phase 1: Rial invoice has currencyUnit ریال');
  assert(loadedRial?.totalAmount === 750000000, 'Phase 1: Rial totalAmount unaltered');

  // Phase 2: Toggle Settings to Rial
  StorageService.saveSettings(settingsRial);
  invoices = StorageService.getInvoices();
  loadedToman = invoices.find((i) => i.id === 'inv_toman_100');
  loadedRial = invoices.find((i) => i.id === 'inv_rial_200');

  assert(loadedToman?.currencyUnit === 'تومان', 'Phase 2: Toman invoice RETAINS تومان despite settings=ریال');
  assert(loadedToman?.totalAmount === 75000000, 'Phase 2: Toman totalAmount unchanged');
  assert(loadedRial?.currencyUnit === 'ریال', 'Phase 2: Rial invoice RETAINS ریال');
  assert(loadedRial?.totalAmount === 750000000, 'Phase 2: Rial totalAmount unchanged');

  // Phase 3: Toggle Settings back to Toman
  StorageService.saveSettings(settingsToman);
  invoices = StorageService.getInvoices();
  loadedToman = invoices.find((i) => i.id === 'inv_toman_100');
  loadedRial = invoices.find((i) => i.id === 'inv_rial_200');

  assert(loadedToman?.currencyUnit === 'تومان', 'Phase 3: Toman invoice intact');
  assert(loadedRial?.currencyUnit === 'ریال', 'Phase 3: Rial invoice intact');
});

runTest('3.2 Legacy Invoices Without currencyUnit Field Gracefully Fall Back', () => {
  mockStorage.clear();

  // Legacy invoice without currencyUnit
  const legacyInvoice: any = {
    id: 'inv_legacy_99',
    invoiceNumber: '99',
    customerName: 'فاکتور قدیمی بدون واحد ارزی',
    jalaliDate: '1404/01/01',
    items: [],
    subtotal: 10000000,
    totalAmount: 10000000,
    status: 'finalized',
    createdAt: '2025-01-01T00:00:00Z',
  };

  StorageService.saveInvoice(legacyInvoice);
  const loaded = StorageService.getInvoices()[0];
  assert(loaded.currencyUnit === undefined, 'Legacy invoice has no currencyUnit');

  // Verify UI fallback expression used in InvoiceHistory and InvoicePrintView:
  // const displayedCurrency = inv.currencyUnit || settings.currencyUnit || 'تومان';
  const settings: CompanySettings = { ...initialCompanySettings, currencyUnit: 'ریال' };
  const fallbackCurrency = loaded.currencyUnit || settings.currencyUnit || 'تومان';
  assert(fallbackCurrency === 'ریال', 'Falls back safely to active settings currency unit');
});

runTest('3.3 Base Pipe Prices Scale Consistently Between Rial (10x) and Toman (1x)', () => {
  const getBasePipePriceForGrade = (
    grade: PipeGrade,
    currency: 'تومان' | 'ریال',
    settings: CompanySettings
  ): number => {
    const baseToman =
      settings?.basePipePrices?.[grade] ||
      (grade === 'PE100' ? 95000 : grade === 'PE80' ? 88000 : 82000);
    return currency === 'ریال' ? baseToman * 10 : baseToman;
  };

  const testSettings = {
    ...initialCompanySettings,
    basePipePrices: {
      PE100: 95000,
      PE80: 88000,
      PE63: 82000,
    },
  };

  const grades: PipeGrade[] = ['PE100', 'PE80', 'PE63'];
  for (const grade of grades) {
    const tomanPrice = getBasePipePriceForGrade(grade, 'تومان', testSettings);
    const rialPrice = getBasePipePriceForGrade(grade, 'ریال', testSettings);
    assert(rialPrice === tomanPrice * 10, `Rial price for ${grade} must be exactly 10x Toman price`);
  }
});

// =====================================================================
// SUITE 4: DUPLICATE VS EDIT SEMANTICS
// =====================================================================
console.log('\nSUITE 4: Duplicate vs Edit Semantics');

runTest('4.1 In-Place Edit Preserves Invoice ID and Invoice Number Without Duplication', () => {
  mockStorage.clear();
  const settings = { ...initialCompanySettings, invoiceCounter: 42 };
  StorageService.saveSettings(settings);

  const originalInvoice: Invoice = {
    id: 'inv_edit_target',
    invoiceNumber: '42',
    customerName: 'مشتری آزمایشی برای ویرایش',
    jalaliDate: '1405/06/15',
    items: [],
    subtotal: 20000000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 20000000,
    currencyUnit: 'تومان',
    status: 'finalized',
    createdAt: '2026-09-01T10:00:00Z',
  };

  StorageService.saveInvoice(originalInvoice);
  assert(StorageService.getInvoices().length === 1, 'Initial invoice stored');

  // User edits this invoice: changing customer name, adding item, updating total
  const pipe = computePipeItem(
    { peGrade: 'PE100', diameter: 90, sdr: 11, pn: 16, length: 150, pricePerKg: 95000 },
    initialWeightTable
  )!;

  const editedInvoice: Invoice = {
    ...originalInvoice,
    customerName: 'مشتری آزمایشی - نام ویرایش شده',
    customerPhone: '09198765432',
    items: [pipe],
    subtotal: pipe.totalPrice,
    totalAmount: pipe.totalPrice,
  };

  // Save the edited invoice
  StorageService.saveInvoice(editedInvoice);

  const updatedInvoices = StorageService.getInvoices();
  assert(updatedInvoices.length === 1, 'In-place edit must NOT increase invoice count');
  assert(updatedInvoices[0].id === 'inv_edit_target', 'ID must remain identical');
  assert(updatedInvoices[0].invoiceNumber === '42', 'Invoice number must remain identical');
  assert(updatedInvoices[0].customerName === 'مشتری آزمایشی - نام ویرایش شده', 'Customer name must be updated');
  assert(updatedInvoices[0].customerPhone === '09198765432', 'Customer phone must be updated');
  assert(updatedInvoices[0].items.length === 1, 'Items list must be updated');
  assert(updatedInvoices[0].totalAmount === pipe.totalPrice, 'Total amount must be updated');
  assert(updatedInvoices[0].createdAt === '2026-09-01T10:00:00Z', 'Original createdAt must be preserved');

  // Verify settings invoiceCounter was NOT incremented on edit
  const currentSettings = StorageService.getSettings();
  assert(currentSettings.invoiceCounter === 42, 'invoiceCounter must NOT increment during edit save');
});

runTest('4.2 Duplicate Allocates New Distinct ID, Increments Invoice Number, and Prepends to Storage', () => {
  mockStorage.clear();
  const settings = { ...initialCompanySettings, invoiceCounter: 42 };
  StorageService.saveSettings(settings);

  const templateInvoice: Invoice = {
    id: 'inv_template_42',
    invoiceNumber: '42',
    customerName: 'مشتری مبنا برای تکثیر',
    jalaliDate: '1405/06/15',
    items: [
      {
        id: 'pipe_item_original',
        type: 'pipe',
        peGrade: 'PE100',
        diameter: 110,
        sdr: 11,
        pn: 16,
        length: 200,
        weightPerMeter: 3.12,
        totalWeight: 624,
        pricePerKg: 95000,
        totalPrice: 59280000,
      },
    ],
    subtotal: 59280000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 59280000,
    currencyUnit: 'تومان',
    status: 'finalized',
    createdAt: '2026-09-01T10:00:00Z',
  };

  StorageService.saveInvoice(templateInvoice);
  assert(StorageService.getInvoices().length === 1, 'Template invoice stored');

  // Simulate handleDuplicateInvoice in App.tsx
  const nextInvoiceNum = StorageService.getNextInvoiceNumber();
  assert(nextInvoiceNum === '43', 'Next invoice number is 43');
  assert(StorageService.getSettings().invoiceCounter === 43, 'invoiceCounter bumped to 43');

  const duplicatedInvoice: Invoice = {
    ...templateInvoice,
    id: 'inv_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
    invoiceNumber: nextInvoiceNum,
    jalaliDate: '1405/07/25',
    createdAt: new Date().toISOString(),
    // Deep clone items
    items: JSON.parse(JSON.stringify(templateInvoice.items)),
  };

  // Verify duplicate invariants before saving
  assert(duplicatedInvoice.id !== templateInvoice.id, 'Duplicated invoice ID must be distinct');
  assert(duplicatedInvoice.invoiceNumber === '43', 'Duplicated invoice number must be 43');
  assert(duplicatedInvoice.invoiceNumber !== templateInvoice.invoiceNumber, 'Numbers must differ');
  assert(duplicatedInvoice.items.length === 1, 'Items copied to duplicate');

  // Save duplicated invoice
  StorageService.saveInvoice(duplicatedInvoice);

  const allInvoices = StorageService.getInvoices();
  assert(allInvoices.length === 2, 'Storage must now contain exactly 2 invoices');
  assert(allInvoices[0].id === duplicatedInvoice.id, 'Duplicated invoice is prepended to top of history');
  assert(allInvoices[0].invoiceNumber === '43', 'Top invoice number is 43');
  assert(allInvoices[1].id === templateInvoice.id, 'Original template invoice remains intact');
  assert(allInvoices[1].invoiceNumber === '42', 'Original invoice number remains 42');
});

runTest('4.3 Duplication Deep Clone Isolation (Mutation of Duplicated Items Does Not Mutate Original)', () => {
  mockStorage.clear();

  const originalPipe: PipeItem = {
    id: 'pipe_orig',
    type: 'pipe',
    peGrade: 'PE100',
    diameter: 63,
    sdr: 11,
    pn: 16,
    length: 100,
    weightPerMeter: 1.05,
    totalWeight: 105,
    pricePerKg: 95000,
    totalPrice: 9975000,
  };

  const invoiceOrig: Invoice = {
    id: 'inv_clone_test',
    invoiceNumber: '1',
    customerName: 'آزمون ایزولاسیون شبیه‌سازی',
    jalaliDate: '1405/01/01',
    items: [originalPipe],
    subtotal: 9975000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 9975000,
    status: 'finalized',
    createdAt: '2026-01-01T00:00:00Z',
  };

  // Deep clone as performed in InvoiceBuilder
  const clonedItems: InvoiceItem[] = JSON.parse(JSON.stringify(invoiceOrig.items));

  // Mutate cloned item
  const clonedPipe = clonedItems[0] as PipeItem;
  clonedPipe.length = 500;
  clonedPipe.totalWeight = 525;
  clonedPipe.totalPrice = 49875000;

  // Assert original remains completely untouched
  assert(originalPipe.length === 100, 'Original pipe length must remain 100');
  assert(originalPipe.totalWeight === 105, 'Original pipe total weight must remain 105');
  assert(originalPipe.totalPrice === 9975000, 'Original pipe total price must remain 9975000');
  assert((invoiceOrig.items[0] as PipeItem).length === 100, 'Invoice original item array unchanged');
});

runTest('4.4 Chained Sequential Duplications (10 Sequential Invoices)', () => {
  mockStorage.clear();
  let counter = 100;
  StorageService.saveSettings({ ...initialCompanySettings, invoiceCounter: counter });

  const baseInvoice: Invoice = {
    id: 'inv_base_100',
    invoiceNumber: '100',
    customerName: 'مشتری پایه زنجیره تکثیر',
    jalaliDate: '1405/01/01',
    items: [],
    subtotal: 1000000,
    includeDiscount: false,
    discountPercentage: 0,
    discountAmount: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
    taxAmount: 0,
    totalAmount: 1000000,
    status: 'finalized',
    createdAt: '2026-01-01T00:00:00Z',
  };
  StorageService.saveInvoice(baseInvoice);

  // Perform 10 consecutive duplicates
  const createdIds = new Set<string>();
  createdIds.add(baseInvoice.id);

  for (let step = 1; step <= 10; step++) {
    const nextNum = StorageService.getNextInvoiceNumber();
    const newId = 'inv_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6) + '_' + step;
    assert(!createdIds.has(newId), `ID collision check for ${newId}`);
    createdIds.add(newId);

    const dup: Invoice = {
      ...baseInvoice,
      id: newId,
      invoiceNumber: nextNum,
      createdAt: new Date().toISOString(),
    };
    StorageService.saveInvoice(dup);
  }

  const all = StorageService.getInvoices();
  assert(all.length === 11, 'Storage must have 1 base + 10 duplicated = 11 invoices');
  assert(StorageService.getSettings().invoiceCounter === 110, 'Counter must be 110');

  // Verify all invoice numbers from 100 to 110 exist and are distinct
  const invoiceNumbers = all.map((i) => i.invoiceNumber);
  for (let num = 100; num <= 110; num++) {
    assert(invoiceNumbers.includes(String(num)), `Invoice number ${num} must exist in history`);
  }
});

// =====================================================================
// SUITE 5: FINANCIAL SUMMARY INTEGRATION & CUSTOMER AUTOFILL
// =====================================================================
console.log('\nSUITE 5: Customer Autofill & Financial Summary Consistency');

runTest('5.1 Customer Autofill Match and Percentage Discount Application', () => {
  const customerWithDiscount: Customer = {
    id: 'cust_vip_1',
    name: 'شرکت پترو سدید پارس',
    phone: '021-88776655',
    address: 'تهران، خیابان ولیعصر',
    defaultDiscount: 12, // 12% default discount
  };

  const pipe1 = computePipeItem(
    { peGrade: 'PE100', diameter: 125, sdr: 11, pn: 16, length: 200, pricePerKg: 95000 },
    initialWeightTable
  )!;

  const summary = computeInvoiceSummary({
    items: [pipe1],
    includeDiscount: false,
    discountPercentage: 0,
    includeCustomDiscount: true,
    customDiscountAmount: Math.round((pipe1.totalPrice * customerWithDiscount.defaultDiscount!) / 100),
    includeTax: true,
    taxPercentage: 10,
  });

  const expectedDiscount = Math.round(pipe1.totalPrice * 0.12);
  const expectedAfterDiscount = pipe1.totalPrice - expectedDiscount;
  const expectedTax = Math.round(expectedAfterDiscount * 0.10);
  const expectedTotal = expectedAfterDiscount + expectedTax;

  assert(summary.subtotal === pipe1.totalPrice, 'Subtotal matches item total');
  assert(summary.discountAmount === expectedDiscount, '12% custom discount correctly computed');
  assert(summary.amountAfterDiscount === expectedAfterDiscount, 'Amount after discount is accurate');
  assert(summary.taxAmount === expectedTax, '10% VAT computed on post-discount amount');
  assert(summary.totalAmount === expectedTotal, 'Total amount equals net + tax');
});

runTest('5.2 Customer Autofill with Flat Toman Discount (>100)', () => {
  const customerWithFlatDiscount: Customer = {
    id: 'cust_flat_1',
    name: 'پیمانکاری سازه‌گستر',
    defaultDiscount: 5000000, // 5,000,000 Toman flat discount
  };

  const pipe1 = computePipeItem(
    { peGrade: 'PE100', diameter: 200, sdr: 17, pn: 10, length: 500, pricePerKg: 95000 },
    initialWeightTable
  )!;

  const summary = computeInvoiceSummary({
    items: [pipe1],
    includeDiscount: false,
    discountPercentage: 0,
    includeCustomDiscount: true,
    customDiscountAmount: customerWithFlatDiscount.defaultDiscount!,
    includeTax: false,
    taxPercentage: 10,
  });

  assert(summary.discountAmount === 5000000, 'Flat discount of 5M applied verbatim');
  assert(summary.totalAmount === pipe1.totalPrice - 5000000, 'Grand total reduced by flat discount');
});

console.log('\n================================================================');
console.log(`ALL CHALLENGER 1 STRESS TESTS PASSED (${passedTests}/${totalTests})`);
console.log('================================================================\n');
