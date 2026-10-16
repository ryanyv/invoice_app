import {
  CompanySettings,
  Customer,
  CustomerAccountSummary,
  Invoice,
  InvoiceSettlementInfo,
  PaymentRecord,
} from '../types';
import { formatNumber, toPersianDigits } from './persianUtils';

export interface CustomerLedgerEntry {
  id: string;
  date: string; // Jalali date
  type: 'invoice' | 'payment';
  documentNumber: string; // شماره فاکتور یا شماره سند/چک
  description: string; // شرح رویداد مالی
  debit: number; // بدهکار (افزایش بدهی مشتری بابت خرید)
  credit: number; // بستانکار (کاهش بدهی مشتری بابت پرداخت)
  balance: number; // مانده حساب جاری بعد از این تراکنش
  diagnosis: 'بدهکار' | 'بستانکار' | 'تسویه';
  referenceId?: string;
  method?: string;
  status?: string;
}

export interface CustomerStatementReport {
  customer: Customer;
  entries: CustomerLedgerEntry[];
  totalDebit: number;
  totalCredit: number;
  pendingCheques: number;
  finalBalance: number;
  diagnosis: 'بدهکار' | 'بستانکار' | 'تسویه';
  generatedDate: string;
}

export interface OfficialTaxRow {
  invoiceNumber: string;
  jalaliDate: string;
  customerName: string;
  customerNationalId: string;
  customerEconomicCode: string;
  customerPostalCode: string;
  grossAmount: number;
  discountAmount: number;
  taxableAmount: number;
  taxAmount: number;
  totalAmount: number;
  currencyUnit: string;
}

export class AccountingService {
  /**
   * Normalizes strings for robust matching between customers, invoices, and payments
   */
  static normalize(str?: string): string {
    return (str || '').trim().toLowerCase();
  }

  /**
   * Computes financial summaries for all customers
   */
  static computeCustomerSummaries(
    customers: Customer[],
    invoices: Invoice[],
    payments: PaymentRecord[]
  ): CustomerAccountSummary[] {
    return customers.map((cust) => {
      const custNormName = this.normalize(cust.name);

      // Invoices for this customer
      const custInvoices = invoices.filter(
        (inv) => this.normalize(inv.customerName) === custNormName
      );
      const totalInvoiced = custInvoices.reduce(
        (sum, inv) => sum + (Number(inv.totalAmount) || 0),
        0
      );

      // Payments for this customer (matched by customerId or normalized customerName)
      const custPayments = payments.filter((p) => {
        if (p.customerId && p.customerId === cust.id) return true;
        return this.normalize(p.customerName) === custNormName;
      });

      // Cleared payments contribute to totalPaid
      const clearedPayments = custPayments.filter((p) => p.status === 'cleared');
      const totalPaid = clearedPayments.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

      // Pending cheques (in-process collections)
      const pendingCheques = custPayments.filter(
        (p) => p.status === 'pending' && p.method === 'cheque'
      );
      const totalPendingCheques = pendingCheques.reduce(
        (sum, p) => sum + (Number(p.amount) || 0),
        0
      );

      // Net balance: positive = owes company (debtor), negative = company owes customer (creditor)
      const balance = totalInvoiced - totalPaid;
      let balanceStatus: 'settled' | 'debtor' | 'creditor' = 'settled';
      if (balance > 1) {
        balanceStatus = 'debtor';
      } else if (balance < -1) {
        balanceStatus = 'creditor';
      }

      // Latest dates
      const lastInvoiceDate = custInvoices.length > 0 ? custInvoices[0].jalaliDate : null;
      const lastPaymentDate = custPayments.length > 0 ? custPayments[0].paymentDate : null;

      return {
        customerId: cust.id,
        customerName: cust.name,
        totalInvoiced,
        totalPaid,
        totalPendingCheques,
        balance,
        balanceStatus,
        invoiceCount: custInvoices.length,
        paymentCount: custPayments.length,
        lastInvoiceDate,
        lastPaymentDate,
      };
    });
  }

  /**
   * Computes settlement tracking info for all invoices
   */
  static computeInvoiceSettlementList(
    invoices: Invoice[],
    payments: PaymentRecord[]
  ): InvoiceSettlementInfo[] {
    return invoices.map((inv) => {
      // Find payments explicitly tied to this invoice by invoiceId or invoiceNumber
      const matchedPayments = payments.filter((p) => {
        if (p.invoiceId && p.invoiceId === inv.id) return true;
        if (p.invoiceNumber && p.invoiceNumber === inv.invoiceNumber) return true;
        return false;
      });

      const totalPaid = matchedPayments
        .filter((p) => p.status === 'cleared')
        .reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

      const invTotal = Number(inv.totalAmount) || 0;
      const remainingAmount = Math.max(0, invTotal - totalPaid);

      let status: 'fully_paid' | 'partially_paid' | 'unpaid' = 'unpaid';
      if (totalPaid >= invTotal && invTotal > 0) {
        status = 'fully_paid';
      } else if (totalPaid > 0) {
        status = 'partially_paid';
      }

      return {
        invoiceId: inv.id,
        invoiceNumber: inv.invoiceNumber,
        customerName: inv.customerName,
        date: inv.jalaliDate,
        totalAmount: invTotal,
        currencyUnit: inv.currencyUnit || 'تومان',
        paidAmount: totalPaid,
        remainingAmount,
        status,
        paymentCount: matchedPayments.length,
      };
    });
  }

  /**
   * Generates a standard double-entry customer ledger statement (کارت حساب معین خریدار)
   */
  static generateCustomerStatement(
    customer: Customer,
    invoices: Invoice[],
    payments: PaymentRecord[],
    currentDate: string
  ): CustomerStatementReport {
    const custNormName = this.normalize(customer.name);

    // Invoices for this customer
    const custInvoices = invoices.filter(
      (inv) => this.normalize(inv.customerName) === custNormName
    );

    // Payments for this customer
    const custPayments = payments.filter((p) => {
      if (p.customerId && p.customerId === customer.id) return true;
      return this.normalize(p.customerName) === custNormName;
    });

    interface RawEvent {
      id: string;
      date: string;
      type: 'invoice' | 'payment';
      documentNumber: string;
      description: string;
      debit: number;
      credit: number;
      method?: string;
      status?: string;
      timestamp: number;
    }

    const events: RawEvent[] = [];

    // Map Invoices -> Debits
    for (const inv of custInvoices) {
      const invAmt = Number(inv.totalAmount) || 0;
      const isOfficial = Boolean(inv.isOfficial);
      events.push({
        id: inv.id,
        date: inv.jalaliDate || currentDate,
        type: 'invoice',
        documentNumber: inv.invoiceNumber,
        description: `فاکتور فروش شماره ${inv.invoiceNumber} (${isOfficial ? 'رسمی ماده ۱۶۹' : 'پیش‌فاکتور عادی'}) - شامل ${inv.items.length} ردیف کالا`,
        debit: invAmt,
        credit: 0,
        timestamp: new Date(inv.createdAt || Date.now()).getTime(),
      });
    }

    // Map Payments -> Credits
    const methodLabels: Record<string, string> = {
      transfer: 'حواله بانکی',
      cheque: 'چک صیادی',
      pos: 'کارت‌خوان / پوز',
      cash: 'دریافت نقدی',
      deposit: 'بیعانه / پیش‌پرداخت',
    };

    for (const p of custPayments) {
      const pAmt = Number(p.amount) || 0;
      const methodLabel = methodLabels[p.method] || p.method;
      let desc = `دریافت وجه (${methodLabel})`;
      if (p.bankName) desc += ` - بانک ${p.bankName}`;
      if (p.referenceNumber) desc += ` - پیگیری/شناسه: ${p.referenceNumber}`;
      if (p.chequeDueDate) desc += ` - سررسید چک: ${p.chequeDueDate}`;
      if (p.invoiceNumber) desc += ` - بابت فاکتور شماره ${p.invoiceNumber}`;
      if (p.notes) desc += ` (${p.notes})`;

      // Only cleared payments reduce ledger balance; pending cheques are tracked as pending documents
      const isCleared = p.status === 'cleared';
      events.push({
        id: p.id,
        date: p.paymentDate || currentDate,
        type: 'payment',
        documentNumber: p.referenceNumber || p.id.slice(-6),
        description: desc + (!isCleared ? ` [وضعیت: ${p.status === 'pending' ? 'در جریان وصول' : 'برگشتی'}]` : ''),
        debit: 0,
        credit: isCleared ? pAmt : 0,
        method: p.method,
        status: p.status,
        timestamp: new Date(p.createdAt || Date.now()).getTime(),
      });
    }

    // Sort chronologically by date string (YYYY/MM/DD) and fallback timestamp
    events.sort((a, b) => {
      if (a.date !== b.date) {
        return a.date.localeCompare(b.date);
      }
      return a.timestamp - b.timestamp;
    });

    // Compute running balance
    let runningBalance = 0;
    let totalDebit = 0;
    let totalCredit = 0;

    const entries: CustomerLedgerEntry[] = events.map((ev) => {
      totalDebit += ev.debit;
      totalCredit += ev.credit;
      runningBalance = runningBalance + ev.debit - ev.credit;

      let diagnosis: 'بدهکار' | 'بستانکار' | 'تسویه' = 'تسویه';
      if (runningBalance > 1) {
        diagnosis = 'بدهکار';
      } else if (runningBalance < -1) {
        diagnosis = 'بستانکار';
      }

      return {
        id: ev.id,
        date: ev.date,
        type: ev.type,
        documentNumber: ev.documentNumber,
        description: ev.description,
        debit: ev.debit,
        credit: ev.credit,
        balance: Math.abs(runningBalance),
        diagnosis,
        method: ev.method,
        status: ev.status,
      };
    });

    const pendingCheques = custPayments
      .filter((p) => p.status === 'pending' && p.method === 'cheque')
      .reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

    const finalBalance = runningBalance;
    let diagnosis: 'بدهکار' | 'بستانکار' | 'تسویه' = 'تسویه';
    if (finalBalance > 1) {
      diagnosis = 'بدهکار';
    } else if (finalBalance < -1) {
      diagnosis = 'بستانکار';
    }

    return {
      customer,
      entries,
      totalDebit,
      totalCredit,
      pendingCheques,
      finalBalance,
      diagnosis,
      generatedDate: currentDate,
    };
  }

  /**
   * Computes official VAT / Tax report data compliant with Iranian Tax Article 169
   */
  static computeOfficialTaxRows(invoices: Invoice[]): OfficialTaxRow[] {
    const officialInvoices = invoices.filter((inv) => Boolean(inv.isOfficial));

    return officialInvoices.map((inv) => {
      const gross = Number(inv.subtotal) || 0;
      const discount = (Number(inv.discountAmount) || 0) + (Number(inv.customDiscountAmount) || 0);
      const taxable = Math.max(0, gross - discount);
      const tax = Number(inv.taxAmount) || 0;
      const total = Number(inv.totalAmount) || 0;

      return {
        invoiceNumber: inv.invoiceNumber,
        jalaliDate: inv.jalaliDate,
        customerName: inv.customerName,
        customerNationalId: inv.customerNationalId || '',
        customerEconomicCode: inv.customerEconomicCode || '',
        customerPostalCode: inv.customerPostalCode || '',
        grossAmount: gross,
        discountAmount: discount,
        taxableAmount: taxable,
        taxAmount: tax,
        totalAmount: total,
        currencyUnit: inv.currencyUnit || 'تومان',
      };
    });
  }

  /**
   * Exports official tax report to RFC 4180 CSV with UTF-8 BOM
   */
  static exportOfficialTaxToCsv(invoices: Invoice[], settings: CompanySettings): string {
    const rows = this.computeOfficialTaxRows(invoices);
    const headers = [
      'شماره فاکتور',
      'تاریخ صدور',
      'نام خریدار',
      'شناسه ملی خریدار',
      'کد اقتصادی خریدار',
      'کد پستی خریدار',
      'مبلغ ناخالص',
      'مبلغ تخفیف',
      'مبلغ مشمول مالیات',
      'مالیات بر ارزش افزوده (۱۰٪)',
      'مبلغ کل نهایی',
      'واحد پول',
    ];

    const escapeCell = (val: string | number | undefined) => {
      if (val === undefined || val === null) return '""';
      return `"${String(val).replace(/"/g, '""')}"`;
    };

    const dataLines = rows.map((r) => [
      escapeCell(r.invoiceNumber),
      escapeCell(r.jalaliDate),
      escapeCell(r.customerName),
      escapeCell(r.customerNationalId),
      escapeCell(r.customerEconomicCode),
      escapeCell(r.customerPostalCode),
      r.grossAmount,
      r.discountAmount,
      r.taxableAmount,
      r.taxAmount,
      r.totalAmount,
      escapeCell(r.currencyUnit),
    ]);

    const titleLine = `گزارش صورت معاملات فصلی و مالیات بر ارزش افزوده ماده ۱۶۹ - ${settings.name}`;
    return '\uFEFF' + [titleLine, headers.join(','), ...dataLines.map((l) => l.join(','))].join('\n');
  }

  /**
   * Exports customer statement of account to RFC 4180 CSV with UTF-8 BOM
   */
  static exportCustomerStatementToCsv(
    report: CustomerStatementReport,
    settings: CompanySettings
  ): string {
    const headers = [
      'ردیف',
      'تاریخ',
      'نوع رویداد',
      'شماره سند / فاکتور',
      'شرح رویداد مالی',
      'بدهکار',
      'بستانکار',
      'مانده حساب',
      'وضعیت تشخیص',
    ];

    const escapeCell = (val: string | number | undefined) => {
      if (val === undefined || val === null) return '""';
      return `"${String(val).replace(/"/g, '""')}"`;
    };

    const typeLabels = { invoice: 'فاکتور فروش', payment: 'دریافت وجه' };

    const dataLines = report.entries.map((e, idx) => [
      idx + 1,
      escapeCell(e.date),
      escapeCell(typeLabels[e.type] || e.type),
      escapeCell(e.documentNumber),
      escapeCell(e.description),
      e.debit,
      e.credit,
      e.balance,
      escapeCell(e.diagnosis),
    ]);

    const headerLines = [
      `صورت‌حساب گردش مالی خریدار: ${report.customer.name}`,
      `شرکت طرف حساب: ${settings.name} | تلفن: ${settings.phone}`,
      `تاریخ تهیه گزارش: ${report.generatedDate}`,
      `شناسه ملی خریدار: ${report.customer.nationalId || '-'} | کد اقتصادی: ${report.customer.economicCode || '-'}`,
      `مجموع بدهکار (فاکتورها): ${report.totalDebit} | مجموع بستانکار (دریافتی‌ها): ${report.totalCredit}`,
      `مانده نهایی حساب: ${Math.abs(report.finalBalance)} (${report.diagnosis})`,
      '',
    ];

    return '\uFEFF' + [...headerLines, headers.join(','), ...dataLines.map((l) => l.join(','))].join('\n');
  }
}
