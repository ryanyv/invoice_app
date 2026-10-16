import { describe, it, expect, beforeEach } from 'vitest';
import { AccountingService } from '../accountingService';
import { StorageService } from '../storage';
import { Customer, Invoice, PaymentRecord, CompanySettings } from '../../types';

// Setup Mock LocalStorage
const store = new Map<string, string>();
const mockLocalStorage = {
  getItem: (key: string) => store.get(key) ?? null,
  setItem: (key: string, value: string) => store.set(key, String(value)),
  removeItem: (key: string) => store.delete(key),
  clear: () => store.clear(),
};
(globalThis as any).localStorage = mockLocalStorage;

describe('AccountingService & CRM Lite Accounting Tests', () => {
  const mockSettings: CompanySettings = {
    name: 'شرکت پلی غرب',
    address: 'کرمانشاه، شهرک صنعتی فرامان',
    phone: '08334274000',
    fax: '08334274001',
    taxRate: 10,
    invoiceCounter: 100,
    currencyUnit: 'تومان',
    basePipePrices: { PE100: 95000, PE80: 88000, PE63: 82000 },
    economicCode: '411456789123',
    nationalId: '10103567890',
  };

  const sampleCustomer: Customer = {
    id: 'cust_101',
    name: 'شرکت مهندسی آبراه غرب',
    phone: '09121112233',
    nationalId: '10101112223',
    economicCode: '411122233344',
  };

  const sampleCustomer2: Customer = {
    id: 'cust_102',
    name: 'آقای صادقی',
    phone: '09185556677',
  };

  const sampleInvoices: Invoice[] = [
    {
      id: 'inv_1',
      invoiceNumber: '1001',
      customerName: 'شرکت مهندسی آبراه غرب',
      jalaliDate: '1403/07/01',
      items: [],
      subtotal: 100000000,
      includeDiscount: true,
      discountPercentage: 5,
      discountAmount: 5000000,
      includeCustomDiscount: false,
      customDiscountAmount: 0,
      includeTax: true,
      taxPercentage: 10,
      taxAmount: 9500000,
      totalAmount: 104500000,
      currencyUnit: 'تومان',
      isOfficial: true,
      status: 'finalized',
      createdAt: '2024-09-22T10:00:00Z',
    },
    {
      id: 'inv_2',
      invoiceNumber: '1002',
      customerName: 'شرکت مهندسی آبراه غرب',
      jalaliDate: '1403/07/15',
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
      isOfficial: false,
      status: 'finalized',
      createdAt: '2024-10-06T12:00:00Z',
    },
    {
      id: 'inv_3',
      invoiceNumber: '1003',
      customerName: 'آقای صادقی',
      jalaliDate: '1403/07/10',
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
      createdAt: '2024-10-01T08:00:00Z',
    },
  ];

  const samplePayments: PaymentRecord[] = [
    {
      id: 'pay_1',
      customerId: 'cust_101',
      customerName: 'شرکت مهندسی آبراه غرب',
      invoiceId: 'inv_1',
      invoiceNumber: '1001',
      amount: 104500000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/05',
      method: 'transfer',
      referenceNumber: 'TRX-987654',
      bankName: 'ملت',
      status: 'cleared',
      createdAt: '2024-09-26T14:00:00Z',
    },
    {
      id: 'pay_2',
      customerId: 'cust_101',
      customerName: 'شرکت مهندسی آبراه غرب',
      invoiceId: 'inv_2',
      invoiceNumber: '1002',
      amount: 20000000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/18',
      method: 'cheque',
      referenceNumber: 'CHK-1234567890123456',
      bankName: 'صادرات',
      chequeDueDate: '1403/09/01',
      status: 'pending', // Pending cheque
      createdAt: '2024-10-09T09:00:00Z',
    },
    {
      id: 'pay_3',
      customerId: 'cust_102',
      customerName: 'آقای صادقی',
      amount: 20000000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/12',
      method: 'pos',
      status: 'cleared',
      createdAt: '2024-10-03T11:00:00Z',
    },
  ];

  it('1. Computes customer account summaries correctly', () => {
    const summaries = AccountingService.computeCustomerSummaries(
      [sampleCustomer, sampleCustomer2],
      sampleInvoices,
      samplePayments
    );

    expect(summaries).toHaveLength(2);

    const cust1 = summaries.find((s) => s.customerId === 'cust_101')!;
    expect(cust1).toBeDefined();
    expect(cust1.totalInvoiced).toBe(154500000); // 104500000 + 50000000
    expect(cust1.totalPaid).toBe(104500000); // Only pay_1 is cleared
    expect(cust1.totalPendingCheques).toBe(20000000); // pay_2 is pending cheque
    expect(cust1.balance).toBe(50000000); // 154500000 - 104500000
    expect(cust1.balanceStatus).toBe('debtor');
    expect(cust1.invoiceCount).toBe(2);
    expect(cust1.paymentCount).toBe(2);

    const cust2 = summaries.find((s) => s.customerId === 'cust_102')!;
    expect(cust2).toBeDefined();
    expect(cust2.totalInvoiced).toBe(20000000);
    expect(cust2.totalPaid).toBe(20000000);
    expect(cust2.balance).toBe(0);
    expect(cust2.balanceStatus).toBe('settled');
  });

  it('2. Computes invoice settlement status correctly', () => {
    const settlements = AccountingService.computeInvoiceSettlementList(
      sampleInvoices,
      samplePayments
    );

    expect(settlements).toHaveLength(3);

    const inv1 = settlements.find((s) => s.invoiceId === 'inv_1')!;
    expect(inv1.status).toBe('fully_paid');
    expect(inv1.paidAmount).toBe(104500000);
    expect(inv1.remainingAmount).toBe(0);

    const inv2 = settlements.find((s) => s.invoiceId === 'inv_2')!;
    // pay_2 was pending cheque, so cleared paidAmount is 0
    expect(inv2.status).toBe('unpaid');
    expect(inv2.remainingAmount).toBe(50000000);
  });

  it('3. Generates complete customer statement ledger with running balance', () => {
    const report = AccountingService.generateCustomerStatement(
      sampleCustomer,
      sampleInvoices,
      samplePayments,
      '1403/07/20'
    );

    expect(report.customer.name).toBe('شرکت مهندسی آبراه غرب');
    expect(report.totalDebit).toBe(154500000);
    expect(report.totalCredit).toBe(104500000);
    expect(report.finalBalance).toBe(50000000);
    expect(report.diagnosis).toBe('بدهکار');
    expect(report.pendingCheques).toBe(20000000);

    // Verify chronological order of entries
    expect(report.entries.length).toBeGreaterThanOrEqual(3);
    expect(report.entries[0].date).toBe('1403/07/01');
    expect(report.entries[0].type).toBe('invoice');
    expect(report.entries[0].debit).toBe(104500000);
    expect(report.entries[0].balance).toBe(104500000);
    expect(report.entries[0].diagnosis).toBe('بدهکار');

    expect(report.entries[1].date).toBe('1403/07/05');
    expect(report.entries[1].type).toBe('payment');
    expect(report.entries[1].credit).toBe(104500000);
    expect(report.entries[1].balance).toBe(0);
    expect(report.entries[1].diagnosis).toBe('تسویه');
  });

  it('4. Computes official tax rows compliant with Article 169', () => {
    const taxRows = AccountingService.computeOfficialTaxRows(sampleInvoices);
    // Only inv_1 has isOfficial: true
    expect(taxRows).toHaveLength(1);
    expect(taxRows[0].invoiceNumber).toBe('1001');
    expect(taxRows[0].grossAmount).toBe(100000000);
    expect(taxRows[0].discountAmount).toBe(5000000);
    expect(taxRows[0].taxableAmount).toBe(95000000);
    expect(taxRows[0].taxAmount).toBe(9500000);
    expect(taxRows[0].totalAmount).toBe(104500000);
  });

  it('5. Exports customer statement and official tax report to RFC 4180 CSV with UTF-8 BOM', () => {
    const report = AccountingService.generateCustomerStatement(
      sampleCustomer,
      sampleInvoices,
      samplePayments,
      '1403/07/20'
    );
    const csvStatement = AccountingService.exportCustomerStatementToCsv(report, mockSettings);
    expect(csvStatement.startsWith('\uFEFF')).toBe(true);
    expect(csvStatement).toContain('شرکت مهندسی آبراه غرب');
    expect(csvStatement).toContain('104500000');

    const csvTax = AccountingService.exportOfficialTaxToCsv(sampleInvoices, mockSettings);
    expect(csvTax.startsWith('\uFEFF')).toBe(true);
    expect(csvTax).toContain('ماده ۱۶۹');
    expect(csvTax).toContain('1001');
    expect(csvTax).toContain('95000000');
  });

  it('6. StorageService payment persistence and CSV export operations', () => {
    StorageService.savePayments(samplePayments);
    const loaded = StorageService.getPayments();
    expect(loaded).toHaveLength(3);
    expect(loaded[0].id).toBe('pay_1');

    // Add new payment
    StorageService.addPayment({
      id: 'pay_new_99',
      customerId: 'cust_101',
      customerName: 'شرکت مهندسی آبراه غرب',
      amount: 15000000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/21',
      method: 'transfer',
      status: 'cleared',
      createdAt: new Date().toISOString(),
    });

    const updated = StorageService.getPayments();
    expect(updated).toHaveLength(4);
    expect(updated[0].id).toBe('pay_new_99');

    // Export payments CSV
    const csv = StorageService.exportPaymentsToCsv();
    expect(csv.startsWith('\uFEFF')).toBe(true);
    expect(csv).toContain('pay_new_99');
    expect(csv).toContain('حواله بانکی');

    // Delete payment
    StorageService.deletePayment('pay_new_99');
    expect(StorageService.getPayments()).toHaveLength(3);
  });

  it('7. Handles empty inputs gracefully without throwing', () => {
    const emptySummaries = AccountingService.computeCustomerSummaries([], [], []);
    expect(emptySummaries).toEqual([]);

    const emptySettlements = AccountingService.computeInvoiceSettlementList([], []);
    expect(emptySettlements).toEqual([]);

    const emptyTax = AccountingService.computeOfficialTaxRows([]);
    expect(emptyTax).toEqual([]);

    const emptyReport = AccountingService.generateCustomerStatement(
      sampleCustomer,
      [],
      [],
      '1403/07/20'
    );
    expect(emptyReport.entries).toEqual([]);
    expect(emptyReport.totalDebit).toBe(0);
    expect(emptyReport.totalCredit).toBe(0);
    expect(emptyReport.finalBalance).toBe(0);
    expect(emptyReport.diagnosis).toBe('تسویه');
  });

  it('8. Handles customer overpayment correctly (creditor diagnosis)', () => {
    const overpaymentCustomer: Customer = {
      id: 'cust_over',
      name: 'آقای شمس',
    };
    const smallInvoice: Invoice = {
      id: 'inv_small',
      invoiceNumber: '2001',
      customerName: 'آقای شمس',
      jalaliDate: '1403/07/01',
      items: [],
      subtotal: 10000000,
      includeDiscount: false,
      discountPercentage: 0,
      discountAmount: 0,
      includeCustomDiscount: false,
      customDiscountAmount: 0,
      includeTax: false,
      taxPercentage: 10,
      taxAmount: 0,
      totalAmount: 10000000,
      currencyUnit: 'تومان',
      status: 'finalized',
      createdAt: '2024-09-22T10:00:00Z',
    };
    const largePayment: PaymentRecord = {
      id: 'pay_large',
      customerId: 'cust_over',
      customerName: 'آقای شمس',
      amount: 15000000, // 5M overpayment
      currencyUnit: 'تومان',
      paymentDate: '1403/07/02',
      method: 'transfer',
      status: 'cleared',
      createdAt: '2024-09-23T10:00:00Z',
    };

    const summaries = AccountingService.computeCustomerSummaries(
      [overpaymentCustomer],
      [smallInvoice],
      [largePayment]
    );
    expect(summaries[0].balance).toBe(-5000000);
    expect(summaries[0].balanceStatus).toBe('creditor');

    const statement = AccountingService.generateCustomerStatement(
      overpaymentCustomer,
      [smallInvoice],
      [largePayment],
      '1403/07/05'
    );
    expect(statement.finalBalance).toBe(-5000000);
    expect(statement.diagnosis).toBe('بستانکار');
  });

  it('9. Robust matching with trimmed whitespace and case insensitivity', () => {
    const cust: Customer = { id: 'cust_trim', name: '   پترو پیشرو غرب   ' };
    const inv: Invoice = {
      ...sampleInvoices[0],
      id: 'inv_trim',
      customerName: 'پترو پیشرو غرب',
      totalAmount: 30000000,
    };
    const pay: PaymentRecord = {
      id: 'pay_trim',
      customerId: '', // No ID, relies on name matching
      customerName: '  پترو پیشرو غرب  ',
      amount: 30000000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/10',
      method: 'pos',
      status: 'cleared',
      createdAt: '2024-10-01T10:00:00Z',
    };

    const summaries = AccountingService.computeCustomerSummaries([cust], [inv], [pay]);
    expect(summaries[0].totalInvoiced).toBe(30000000);
    expect(summaries[0].totalPaid).toBe(30000000);
    expect(summaries[0].balance).toBe(0);
    expect(summaries[0].balanceStatus).toBe('settled');
  });

  it('10. Full backup and restore round-trip preserves payments array', () => {
    StorageService.savePayments(samplePayments);
    const backupJson = StorageService.createFullBackup();
    const parsed = JSON.parse(backupJson);
    expect(parsed.payments).toBeDefined();
    expect(parsed.payments).toHaveLength(3);

    // Clear and restore
    StorageService.savePayments([]);
    expect(StorageService.getPayments()).toHaveLength(0);

    const restoreSuccess = StorageService.restoreFromBackup(backupJson);
    expect(restoreSuccess).toBe(true);
    expect(StorageService.getPayments()).toHaveLength(3);
  });

  it('11. Cheque status lifecycle (pending -> cleared) updates balance reactively', () => {
    const cust: Customer = { id: 'cust_chk', name: 'خریدار چک' };
    const inv: Invoice = {
      ...sampleInvoices[0],
      id: 'inv_chk',
      customerName: 'خریدار چک',
      totalAmount: 50000000,
    };
    const chequePayment: PaymentRecord = {
      id: 'pay_chk',
      customerId: 'cust_chk',
      customerName: 'خریدار چک',
      amount: 50000000,
      currencyUnit: 'تومان',
      paymentDate: '1403/07/10',
      method: 'cheque',
      chequeDueDate: '1403/09/10',
      status: 'pending', // Pending
      createdAt: '2024-10-01T10:00:00Z',
    };

    // While pending: balance is still 50M, pendingCheques is 50M
    let summaries = AccountingService.computeCustomerSummaries([cust], [inv], [chequePayment]);
    expect(summaries[0].balance).toBe(50000000);
    expect(summaries[0].balanceStatus).toBe('debtor');
    expect(summaries[0].totalPendingCheques).toBe(50000000);

    // Once cleared: balance becomes 0, pendingCheques becomes 0
    const clearedCheque: PaymentRecord = { ...chequePayment, status: 'cleared' };
    summaries = AccountingService.computeCustomerSummaries([cust], [inv], [clearedCheque]);
    expect(summaries[0].balance).toBe(0);
    expect(summaries[0].balanceStatus).toBe('settled');
    expect(summaries[0].totalPendingCheques).toBe(0);
  });
});
