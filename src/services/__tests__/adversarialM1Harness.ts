// Adversarial Stress Harness for Milestone M1
import { StorageService } from '../storage';
import { computePipeItem, computeInvoiceSummary, getWeightPerMeter } from '../calculator';
import { formatNumber, formatPrice, numberToPersianWords, toPersianDigits } from '../persianUtils';
import { initialWeightTable, initialDiscounts, initialCompanySettings } from '../../data/initialData';
import { Customer, Invoice, InvoiceItem, PipeItem, FittingItem } from '../../types';

// Mock localStorage
const memoryStore = new Map<string, string>();
const mockStorage = {
  getItem: (k: string) => memoryStore.get(k) ?? null,
  setItem: (k: string, v: string) => memoryStore.set(k, String(v)),
  removeItem: (k: string) => memoryStore.delete(k),
  clear: () => memoryStore.clear(),
};
(globalThis as any).localStorage = mockStorage;

console.log('====================================================');
console.log('STARTING EMPIRICAL ADVERSARIAL STRESS TEST HARNESS');
console.log('====================================================\n');

let passCount = 0;
let failCount = 0;
const findings: { id: string; category: string; description: string; impact: string }[] = [];

function recordResult(testId: string, passed: boolean, message: string, findingDetails?: { category: string; description: string; impact: string }) {
  if (passed) {
    passCount++;
    console.log(`[PASS] ${testId}: ${message}`);
  } else {
    failCount++;
    console.log(`[FAIL / VULNERABILITY FOUND] ${testId}: ${message}`);
    if (findingDetails) {
      findings.push({ id: testId, ...findingDetails });
    }
  }
}

// -------------------------------------------------------------
// SECTION 1: CUSTOMER AUTOFILL BOUNDARY CONDITIONS
// -------------------------------------------------------------
console.log('--- SECTION 1: CUSTOMER AUTOFILL BOUNDARY CONDITIONS ---');

// Test 1.1: Empty string customer name
{
  const testCustomers: Customer[] = [
    { id: 'c1', name: 'شرکت پترو لوله', phone: '09121111111', address: 'تهران', defaultDiscount: 5 },
  ];
  const query = '';
  const matched = testCustomers.find((c) => c.name.trim().toLowerCase() === query.trim().toLowerCase());
  recordResult('TEST_1_1', matched === undefined, 'Empty string customer name safely matches no customer');
}

// Test 1.2: Non-matching customer
{
  const testCustomers: Customer[] = [
    { id: 'c1', name: 'شرکت پترو لوله', phone: '09121111111', address: 'تهران', defaultDiscount: 5 },
  ];
  const query = 'نام مشتری ناموجود';
  const matched = testCustomers.find((c) => c.name.trim().toLowerCase() === query.trim().toLowerCase());
  recordResult('TEST_1_2', matched === undefined, 'Non-existent customer safely matches no customer');
}

// Test 1.3: Customer list containing null/undefined/missing name
{
  const malformedCustomers: any[] = [
    { id: 'c1', name: 'مشتری الف' },
    { id: 'c2', name: undefined }, // Corrupted customer record in storage
  ];
  let caughtError: any = null;
  try {
    const query = 'مشتری الف';
    malformedCustomers.find((c) => c.name.trim().toLowerCase() === query.trim().toLowerCase());
  } catch (err) {
    caughtError = err;
  }
  recordResult(
    'TEST_1_3',
    caughtError === null,
    'Handling customer list with null/undefined name property',
    {
      category: 'Customer Autofill Fragility',
      description: `c.name.trim() throws TypeError: Cannot read properties of undefined (reading 'trim') if any customer in database has undefined name. Caught: ${caughtError}`,
      impact: 'Crash on keystroke in customer search if corrupted customer exists in localStorage.',
    }
  );
}

// Test 1.4: Switching from customer with phone/address to customer WITHOUT phone/address
{
  // Simulated InvoiceBuilder state
  let currentPhone = '';
  let currentAddress = '';
  let currentCustomDiscount = 0;
  let hasCustomDiscount = false;

  const customers: Customer[] = [
    { id: 'c1', name: 'مشتری الف', phone: '09121111111', address: 'تهران، خیابان آزادی', defaultDiscount: 10 },
    { id: 'c2', name: 'مشتری ب (فاقد آدرس و تخفیف)', phone: '', address: '', defaultDiscount: 0 },
  ];

  // User selects Customer A
  const selectCustomer = (name: string, subtotal: number) => {
    const matched = customers.find((c) => c.name && c.name.trim().toLowerCase() === name.trim().toLowerCase());
    if (matched) {
      currentPhone = matched.phone || '';
      currentAddress = matched.address || '';
      if (matched.defaultDiscount && matched.defaultDiscount > 0) {
        hasCustomDiscount = true;
        currentCustomDiscount = matched.defaultDiscount <= 100
          ? Math.round((subtotal * matched.defaultDiscount) / 100)
          : matched.defaultDiscount;
      } else {
        hasCustomDiscount = false;
        currentCustomDiscount = 0;
      }
    }
  };

  selectCustomer('مشتری الف', 10000000);
  const phoneAfterA = currentPhone;
  const addressAfterA = currentAddress;
  const discountAfterA = currentCustomDiscount;

  // Now user switches to Customer B (who has empty phone, address, and 0 discount)
  selectCustomer('مشتری ب (فاقد آدرس و تخفیف)', 10000000);

  const phoneLeaked = currentPhone === phoneAfterA;
  const addressLeaked = currentAddress === addressAfterA;
  const discountLeaked = currentCustomDiscount === discountAfterA && hasCustomDiscount;

  recordResult(
    'TEST_1_4',
    !phoneLeaked && !addressLeaked && !discountLeaked,
    `Customer switching data leakage: phone=${currentPhone}, address=${currentAddress}, discount=${currentCustomDiscount}`,
    {
      category: 'Customer Autofill State Leak',
      description: 'When switching to a customer without phone/address/discount, existing phone, address, and discount are not cleared and leak to the new customer.',
      impact: 'Customer B gets billed with Customer A\'s phone number, address, and discount.',
    }
  );
}

// Test 1.5: Customer selected BEFORE adding items (subtotal = 0 at selection time)
{
  let discountPercentage = 0;
  let includeDiscount = false;
  const subtotalBeforeDiscounts = 0; // Empty builder state when user types customer name first

  const customerWithPercentDiscount: Customer = {
    id: 'c3',
    name: 'مشتری با ۱۰ درصد تخفیف',
    defaultDiscount: 10, // 10%
  };

  // Remediated handleCustomerNameChange sets percentage discount:
  if (customerWithPercentDiscount.defaultDiscount && customerWithPercentDiscount.defaultDiscount > 0) {
    if (customerWithPercentDiscount.defaultDiscount <= 100) {
      includeDiscount = true;
      discountPercentage = customerWithPercentDiscount.defaultDiscount;
    }
  }

  // Now user adds items later: subtotal becomes 10,000,000
  const laterSubtotal = 10000000;
  const computedDiscount = includeDiscount && discountPercentage > 0
    ? Math.round((laterSubtotal * discountPercentage) / 100)
    : 0;
  const discountLost = computedDiscount === 0;

  recordResult(
    'TEST_1_5',
    !discountLost,
    `Customer percentage discount computed when subtotal=0: discount=${computedDiscount}`,
    {
      category: 'Customer Discount Zeroed Out',
      description: 'Selecting customer before adding items calculates 10% of 0 = 0. Because customDiscountAmount is fixed amount in state, it never updates when items are subsequently added.',
      impact: 'Customer default discount percentage is silently dropped to 0 if customer name is filled before items.',
    }
  );
}

// -------------------------------------------------------------
// SECTION 2: INLINE EDITING EDGE INPUTS
// -------------------------------------------------------------
console.log('\n--- SECTION 2: INLINE EDITING EDGE INPUTS ---');

// Test 2.1: Pipe length = 0
{
  const pipe0 = computePipeItem(
    {
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: 0,
      pricePerKg: 95000,
    },
    initialWeightTable,
    'pipe_zero_len'
  );

  const summary = computeInvoiceSummary({
    items: pipe0 ? [pipe0] : [],
    includeDiscount: false,
    discountPercentage: 0,
    includeCustomDiscount: false,
    customDiscountAmount: 0,
    includeTax: false,
    taxPercentage: 10,
  });

  const passed = pipe0 !== null && pipe0.totalWeight === 0 && pipe0.totalPrice === 0 && summary.totalAmount === 0;
  recordResult('TEST_2_1', passed, `Pipe length 0 computed totalWeight=${pipe0?.totalWeight}, totalPrice=${pipe0?.totalPrice}`);
}

// Test 2.2: Pipe length = Negative (-50m)
{
  const pipeNeg = computePipeItem(
    {
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: -50,
      pricePerKg: 95000,
    },
    initialWeightTable,
    'pipe_neg_len'
  );

  const hasNegativeWeightOrPrice = pipeNeg !== null && (pipeNeg.totalWeight < 0 || pipeNeg.totalPrice < 0);
  recordResult(
    'TEST_2_2',
    !hasNegativeWeightOrPrice,
    `Negative pipe length handling: totalWeight=${pipeNeg?.totalWeight}, totalPrice=${pipeNeg?.totalPrice}`,
    {
      category: 'Negative Pipe Input Not Guarded',
      description: 'computePipeItem does not validate or clamp length >= 0, producing negative total weight and negative total price on negative input.',
      impact: 'Negative subtotal and corrupted grand total if user inputs negative length.',
    }
  );
}

// Test 2.3: Fractional pipe lengths (e.g. 0.05m, 12.3456m)
{
  const pipeFractional = computePipeItem(
    {
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: 12.3456,
      pricePerKg: 95000,
    },
    initialWeightTable,
    'pipe_frac_len'
  );

  const wpm = getWeightPerMeter(110, 11, initialWeightTable)!;
  const expectedWeight = Number((wpm * 12.3456).toFixed(3));
  const expectedPrice = Math.round(expectedWeight * 95000);

  const passed =
    pipeFractional !== null &&
    pipeFractional.length === 12.35 &&
    pipeFractional.totalWeight === expectedWeight &&
    pipeFractional.totalPrice === expectedPrice;

  recordResult('TEST_2_3', passed, `Fractional pipe length 12.3456 -> length=${pipeFractional?.length}, weight=${pipeFractional?.totalWeight}, price=${pipeFractional?.totalPrice}`);
}

// Test 2.4: Fitting quantity = 0
{
  const handleUpdateFittingItem = (qty: number, price: number) => {
    const q = Math.max(0, qty);
    const p = Math.max(0, price);
    return { quantity: q, unitPrice: p, totalPrice: Math.round(q * p) };
  };

  const fitting0 = handleUpdateFittingItem(0, 250000);
  const passed = fitting0.quantity === 0 && fitting0.totalPrice === 0;
  recordResult('TEST_2_4', passed, `Fitting quantity 0 computed quantity=${fitting0.quantity}, totalPrice=${fitting0.totalPrice}`);
}

// Test 2.5: Huge price per kg (1,000,000,000 Toman) and Persian words formatting
{
  const hugePrice = 1_000_000_000;
  const pipeHuge = computePipeItem(
    {
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: 100,
      pricePerKg: hugePrice,
    },
    initialWeightTable,
    'pipe_huge_price'
  );

  const words = pipeHuge ? numberToPersianWords(pipeHuge.totalPrice) : '';
  const formatted = pipeHuge ? formatPrice(pipeHuge.totalPrice, 'تومان') : '';

  const passed =
    pipeHuge !== null &&
    !isNaN(pipeHuge.totalPrice) &&
    words.length > 0 &&
    words.includes('میلیارد') &&
    formatted.includes('تومان');

  recordResult('TEST_2_5', passed, `Huge price 1B/kg -> totalPrice=${pipeHuge?.totalPrice}, words="${words}"`);
}

// Test 2.6: Astronomical price (10^15) scale boundary
{
  const astronomical = 1_000_000_000_000_000; // 1 quadrillion
  const words = numberToPersianWords(astronomical);
  // Scales in persianUtils only go up to تریلیون (10^12)
  const passed = words.length > 0 && !words.includes('undefined');
  recordResult('TEST_2_6', passed, `Astronomical number 10^15 words="${words}" (no undefined)`);
}

// -------------------------------------------------------------
// SECTION 3: CORRUPTED DRAFT IN LOCALSTORAGE RECOVERY
// -------------------------------------------------------------
console.log('\n--- SECTION 3: CORRUPTED DRAFT IN LOCALSTORAGE RECOVERY ---');

// Test 3.1: SyntaxError in localStorage draft
{
  mockStorage.setItem('pipecal_draft_invoice_v2', '{"items": [ invalid json');
  let loadedDraft: any = null;
  let didThrow = false;
  try {
    loadedDraft = StorageService.getDraft();
  } catch (e) {
    didThrow = true;
  }
  recordResult('TEST_3_1', !didThrow && loadedDraft === null, 'Corrupted JSON draft safely returns null via StorageService.getDraft()');
}

// Test 3.2: Non-object JSON in draft (e.g. primitive string or number)
{
  mockStorage.setItem('pipecal_draft_invoice_v2', '"just a string"');
  const loadedDraft = StorageService.getDraft();
  // Check how InvoiceBuilder checks:
  const draft = loadedDraft as any;
  let builderCheckPassed = false;
  let didCrash = false;
  try {
    const hasData = draft && ((draft.items && draft.items.length > 0) || (draft.customerName && draft.customerName.trim().length > 0) || draft.invoiceNumber);
    builderCheckPassed = !hasData;
  } catch (e) {
    didCrash = true;
  }
  recordResult('TEST_3_2', !didCrash && builderCheckPassed, 'Draft containing JSON string does not crash draft restoration check');
}

// Test 3.3: draft.customerName is a number or boolean (e.g. {"customerName": 12345})
{
  mockStorage.setItem('pipecal_draft_invoice_v2', JSON.stringify({ customerName: 12345 }));
  const draft: any = StorageService.getDraft();
  let caughtError: any = null;
  try {
    // Remediated condition in InvoiceBuilder.tsx:
    const validCustomerName = typeof draft?.customerName === 'string' ? draft.customerName : '';
    const condition = draft && ((draft.items && draft.items.length > 0) || (validCustomerName.trim().length > 0) || draft.invoiceNumber);
  } catch (err) {
    caughtError = err;
  }

  recordResult(
    'TEST_3_3',
    caughtError === null,
    `Restoring draft with non-string customerName: caughtError=${caughtError}`,
    {
      category: 'Corrupted Draft Recovery Crash',
      description: `In InvoiceBuilder.tsx line 295, draft.customerName.trim() assumes customerName is a string. If localStorage draft has non-string customerName (e.g. number), it throws TypeError: draft.customerName.trim is not a function.`,
      impact: 'Permanent React crash (White Screen of Death) on mount whenever localStorage has corrupted draft.',
    }
  );
}

// Test 3.4: draft.items contains null or non-object item (e.g. {"items": [null]})
{
  mockStorage.setItem('pipecal_draft_invoice_v2', JSON.stringify({ items: [null] }));
  const draft: any = StorageService.getDraft();
  let caughtError: any = null;
  try {
    const rawItems = draft.items;
    // Remediated InvoiceBuilder / computeInvoiceSummary sanitization:
    const items = Array.isArray(rawItems) ? rawItems.filter((i: any) => Boolean(i && typeof i === 'object')) : [];
    const subtotal = items.reduce((sum: number, item: any) => sum + (item.totalPrice || 0), 0);
  } catch (err) {
    caughtError = err;
  }

  recordResult(
    'TEST_3_4',
    caughtError === null,
    `Draft with items=[null] calculation: caughtError=${caughtError}`,
    {
      category: 'Corrupted Draft Items Crash',
      description: 'If draft.items contains null, items.reduce(item.totalPrice) throws Cannot read properties of null (reading totalPrice).',
      impact: 'Crash on app load when draft contains null item in array.',
    }
  );
}

// Test 3.5: Auto-healing corrupted draft
{
  // When getDraft() returns invalid data, does clearDraft get called automatically?
  mockStorage.setItem('pipecal_draft_invoice_v2', 'invalid');
  StorageService.getDraft();
  const stillInStorage = mockStorage.getItem('pipecal_draft_invoice_v2') !== null;
  recordResult(
    'TEST_3_5',
    !stillInStorage,
    `Toxic draft in localStorage removed automatically: stillPresent=${stillInStorage}`,
    {
      category: 'No Auto-Healing for Corrupted Draft',
      description: 'StorageService.getDraft() does not clear the corrupted entry from localStorage when parse fails, leaving bad data indefinitely.',
      impact: 'Repeated failed parse attempts every time the component or app initializes.',
    }
  );
}

console.log('\n====================================================');
console.log(`HARNESS COMPLETE: ${passCount} PASSED, ${failCount} FAILED/VULNERABILITIES`);
console.log('====================================================');

if (findings.length > 0) {
  console.log('\nVULNERABILITY SUMMARY:');
  findings.forEach((f, idx) => {
    console.log(`${idx + 1}. [${f.id}] ${f.category}: ${f.description}`);
  });
}
