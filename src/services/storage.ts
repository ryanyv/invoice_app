import {
  CompanySettings,
  Customer,
  DiscountRule,
  FittingCatalogItem,
  GradeSdrPn,
  Invoice,
  WeightTable,
} from '../types';
import {
  initialCompanySettings,
  initialDiscounts,
  initialFittings,
  initialGradeSdrPn,
  initialWeightTable,
} from '../data/initialData';

const KEYS = {
  FITTINGS: 'pipecal_fittings_v2',
  SETTINGS: 'pipecal_settings_v2',
  DISCOUNTS: 'pipecal_discounts_v2',
  CUSTOMERS: 'pipecal_customers_v2',
  INVOICES: 'pipecal_invoices_v2',
  WEIGHT_TABLE: 'pipecal_weight_table_v2',
  GRADE_SDR_PN: 'pipecal_grade_sdr_pn_v2',
  DRAFT: 'pipecal_draft_invoice_v2',
};

export class StorageService {
  // --- Fittings Catalog ---
  static getFittings(): FittingCatalogItem[] {
    try {
      const data = localStorage.getItem(KEYS.FITTINGS);
      return data ? JSON.parse(data) : initialFittings;
    } catch {
      return initialFittings;
    }
  }

  static saveFittings(fittings: FittingCatalogItem[]): void {
    localStorage.setItem(KEYS.FITTINGS, JSON.stringify(fittings));
    this.triggerDiskBackup();
  }

  static updateFittingPrice(id: string, newUnitPriceRial: number): void {
    const fittings = this.getFittings();
    const index = fittings.findIndex((f) => f.id === id);
    if (index !== -1) {
      fittings[index].unitPriceRial = newUnitPriceRial;
      this.saveFittings(fittings);
    }
  }

  static bulkUpdateFittingPrices(
    filter: { category?: string; connectionType?: string; product?: string },
    percentageDelta: number // e.g. +10 or -5
  ): void {
    const fittings = this.getFittings();
    const factor = 1 + percentageDelta / 100;
    const updated = fittings.map((f) => {
      let match = true;
      if (filter.category && f.category !== filter.category) match = false;
      if (filter.connectionType && f.connectionType !== filter.connectionType) match = false;
      if (filter.product && f.product !== filter.product) match = false;

      if (match) {
        return {
          ...f,
          unitPriceRial: Math.round(f.unitPriceRial * factor),
        };
      }
      return f;
    });
    this.saveFittings(updated);
  }

  // --- Company Settings & Base Pipe Prices ---
  static getSettings(): CompanySettings {
    const fallbackPrices = {
      PE100: 95000,
      PE80: 88000,
      PE63: 82000,
    };
    try {
      const data = localStorage.getItem(KEYS.SETTINGS);
      if (!data) return { ...initialCompanySettings };
      const parsed = JSON.parse(data);
      const userPrices = parsed?.basePipePrices || {};
      return {
        ...initialCompanySettings,
        ...parsed,
        basePipePrices: {
          PE100: Number(userPrices.PE100) || initialCompanySettings.basePipePrices?.PE100 || fallbackPrices.PE100,
          PE80: Number(userPrices.PE80) || initialCompanySettings.basePipePrices?.PE80 || fallbackPrices.PE80,
          PE63: Number(userPrices.PE63) || initialCompanySettings.basePipePrices?.PE63 || fallbackPrices.PE63,
        },
      };
    } catch {
      return { ...initialCompanySettings };
    }
  }

  static saveSettings(settings: CompanySettings): void {
    localStorage.setItem(KEYS.SETTINGS, JSON.stringify(settings));
    this.triggerDiskBackup();
  }

  static getNextInvoiceNumber(): string {
    const settings = this.getSettings();
    const nextNum = (settings.invoiceCounter || 1) + 1;
    settings.invoiceCounter = nextNum;
    this.saveSettings(settings);
    return String(nextNum);
  }

  // --- Standards ---
  static getWeightTable(): WeightTable {
    try {
      const data = localStorage.getItem(KEYS.WEIGHT_TABLE);
      return data ? JSON.parse(data) : initialWeightTable;
    } catch {
      return initialWeightTable;
    }
  }

  static getGradeSdrPn(): GradeSdrPn {
    try {
      const data = localStorage.getItem(KEYS.GRADE_SDR_PN);
      return data ? JSON.parse(data) : initialGradeSdrPn;
    } catch {
      return initialGradeSdrPn;
    }
  }

  // --- Discount Rules ---
  static getDiscounts(): DiscountRule[] {
    try {
      const data = localStorage.getItem(KEYS.DISCOUNTS);
      return data ? JSON.parse(data) : initialDiscounts;
    } catch {
      return initialDiscounts;
    }
  }

  static saveDiscounts(discounts: DiscountRule[]): void {
    localStorage.setItem(KEYS.DISCOUNTS, JSON.stringify(discounts));
    this.triggerDiskBackup();
  }

  // --- Customers ---
  static getCustomers(): Customer[] {
    try {
      const data = localStorage.getItem(KEYS.CUSTOMERS);
      return data ? JSON.parse(data) : [];
    } catch {
      return [];
    }
  }

  static saveCustomers(customers: Customer[]): void {
    localStorage.setItem(KEYS.CUSTOMERS, JSON.stringify(customers));
    this.triggerDiskBackup();
  }

  static addOrUpdateCustomer(customer: Customer): void {
    const list = this.getCustomers();
    const idx = list.findIndex(
      (c) =>
        (customer.id && c.id === customer.id) ||
        c.name.trim().toLowerCase() === customer.name.trim().toLowerCase()
    );
    if (idx !== -1) {
      const existing = list[idx];
      list[idx] = {
        ...existing,
        id: existing.id,
        name: customer.name.trim(),
        phone: customer.phone !== undefined ? customer.phone : existing.phone,
        address: customer.address !== undefined ? customer.address : existing.address,
        nationalId: customer.nationalId !== undefined ? customer.nationalId : existing.nationalId,
        economicCode: customer.economicCode !== undefined ? customer.economicCode : existing.economicCode,
        postalCode: customer.postalCode !== undefined ? customer.postalCode : existing.postalCode,
        defaultDiscount:
          customer.defaultDiscount !== undefined ? customer.defaultDiscount : existing.defaultDiscount,
      };
    } else {
      list.push({
        ...customer,
        id: customer.id || 'cust_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
        name: customer.name.trim(),
      });
    }
    this.saveCustomers(list);
  }

  static deleteCustomer(id: string): void {
    const list = this.getCustomers().filter((c) => c.id !== id);
    this.saveCustomers(list);
  }

  static exportCustomersToCsv(): string {
    const customers = this.getCustomers();
    const headers = [
      'شناسه',
      'نام مشتری / شرکت',
      'تلفن تماس',
      'کد اقتصادی (ماده ۱۶۹)',
      'شناسه ملی / کد ملی',
      'کد پستی',
      'تخفیف پیش‌فرض',
      'نشانی کامل',
    ];
    const escapeCell = (val: string | number | undefined) => {
      if (val === undefined || val === null) return '""';
      return `"${String(val).replace(/"/g, '""')}"`;
    };
    const rows = customers.map((c) => [
      escapeCell(c.id),
      escapeCell(c.name),
      escapeCell(c.phone || ''),
      escapeCell(c.economicCode || ''),
      escapeCell(c.nationalId || ''),
      escapeCell(c.postalCode || ''),
      c.defaultDiscount || 0,
      escapeCell(c.address || ''),
    ]);
    return '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
  }

  // --- Invoices History ---
  static getInvoices(): Invoice[] {
    try {
      const data = localStorage.getItem(KEYS.INVOICES);
      return data ? JSON.parse(data) : [];
    } catch {
      return [];
    }
  }

  static saveInvoices(invoices: Invoice[]): void {
    localStorage.setItem(KEYS.INVOICES, JSON.stringify(invoices));
    this.triggerDiskBackup();
  }

  static saveInvoice(invoice: Invoice): void {
    const list = this.getInvoices();
    const idx = list.findIndex((i) => i.id === invoice.id);
    if (idx !== -1) {
      list[idx] = invoice;
    } else {
      list.unshift(invoice);
    }
    this.saveInvoices(list);
  }

  static deleteInvoice(id: string): void {
    const list = this.getInvoices().filter((i) => i.id !== id);
    this.saveInvoices(list);
  }

  // --- Real-Time Draft Auto-Save ---
  static getDraft(): Partial<Invoice> | null {
    try {
      const data = localStorage.getItem(KEYS.DRAFT);
      if (!data) return null;
      const parsed = JSON.parse(data);
      if (!parsed || typeof parsed !== 'object') {
        this.clearDraft();
        return null;
      }
      return parsed;
    } catch {
      this.clearDraft();
      return null;
    }
  }

  static saveDraft(draft: Partial<Invoice>): void {
    try {
      localStorage.setItem(KEYS.DRAFT, JSON.stringify(draft));
    } catch (e) {
      console.error('Failed to save draft invoice', e);
    }
  }

  static clearDraft(): void {
    localStorage.removeItem(KEYS.DRAFT);
  }

  // --- Export & Backup ---
  static createFullBackup(): string {
    const backup = {
      version: '2.0.0',
      exportedAt: new Date().toISOString(),
      fittings: this.getFittings(),
      settings: this.getSettings(),
      discounts: this.getDiscounts(),
      customers: this.getCustomers(),
      invoices: this.getInvoices(),
      weightTable: this.getWeightTable(),
      gradeSdrPn: this.getGradeSdrPn(),
    };
    return JSON.stringify(backup, null, 2);
  }

  static restoreFromBackup(jsonString: string): boolean {
    try {
      const data = JSON.parse(jsonString);
      if (data.fittings) localStorage.setItem(KEYS.FITTINGS, JSON.stringify(data.fittings));
      if (data.settings) localStorage.setItem(KEYS.SETTINGS, JSON.stringify(data.settings));
      if (data.discounts) localStorage.setItem(KEYS.DISCOUNTS, JSON.stringify(data.discounts));
      if (data.customers) localStorage.setItem(KEYS.CUSTOMERS, JSON.stringify(data.customers));
      if (data.invoices) localStorage.setItem(KEYS.INVOICES, JSON.stringify(data.invoices));
      if (data.weightTable) localStorage.setItem(KEYS.WEIGHT_TABLE, JSON.stringify(data.weightTable));
      if (data.gradeSdrPn) localStorage.setItem(KEYS.GRADE_SDR_PN, JSON.stringify(data.gradeSdrPn));
      this.triggerDiskBackup();
      return true;
    } catch {
      return false;
    }
  }

  // --- Automated Rolling Disk Backups (Layer 2) ---
  static async triggerDiskBackup(): Promise<boolean> {
    try {
      if (typeof window !== 'undefined' && window.electronAPI?.saveBackup) {
        const backupJson = this.createFullBackup();
        const res = await window.electronAPI.saveBackup({ backupJson });
        return Boolean(res?.success);
      }
    } catch (err) {
      console.warn('[StorageService] Automated disk backup failed:', err);
    }
    return false;
  }

  // --- Export Fittings to CSV ---
  static exportFittingsToCsv(): string {
    const fittings = this.getFittings();
    const headers = [
      'شناسه',
      'نوع',
      'نوع محصول',
      'محصول',
      'جنس',
      'فشار قابل تحمل',
      'اندازه (mm)',
      'قیمت واحد (ریال)',
      'تعداد در کارتن',
      'وضعیت فعال',
    ];
    const rows = fittings.map((f) => [
      f.id,
      `"${f.category}"`,
      `"${f.connectionType}"`,
      `"${f.product}"`,
      `"${f.material}"`,
      `"${f.pressure}"`,
      `"${f.size}"`,
      f.unitPriceRial,
      f.packageCount,
      f.isActive ? '1' : '0',
    ]);
    return '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
  }

  static resetToDefaults(): void {
    localStorage.removeItem(KEYS.FITTINGS);
    localStorage.removeItem(KEYS.SETTINGS);
    localStorage.removeItem(KEYS.DISCOUNTS);
  }
}

