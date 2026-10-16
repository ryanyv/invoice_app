import React, { useState, useMemo } from 'react';
import {
  Users,
  Search,
  Plus,
  Edit2,
  Trash2,
  FilePlus,
  History,
  Download,
  Upload,
  Check,
  X,
  ShieldCheck,
  Percent,
  Phone,
  MapPin,
  TrendingUp,
  Printer,
  Copy,
  Edit3,
  Building2,
} from 'lucide-react';
import { CompanySettings, Customer, Invoice } from '../../types';
import { StorageService } from '../../services/storage';
import { formatNumber, toPersianDigits } from '../../services/persianUtils';

interface CustomerManagerProps {
  customers: Customer[];
  onUpdateCustomers: (customers: Customer[]) => void;
  invoices: Invoice[];
  settings: CompanySettings;
  onStartInvoiceForCustomer: (customer: Customer) => void;
  onPreviewPrintInvoice: (invoice: Invoice) => void;
  onEditInvoice: (invoice: Invoice) => void;
  onDuplicateInvoice: (invoice: Invoice) => void;
}

type CustomerFilter = 'all' | 'official' | 'vip' | 'with_orders';

interface CustomerStats {
  invoiceCount: number;
  totalSpent: number;
  lastInvoiceDate: string | null;
  customerInvoices: Invoice[];
}

export const CustomerManager: React.FC<CustomerManagerProps> = ({
  customers,
  onUpdateCustomers,
  invoices,
  settings,
  onStartInvoiceForCustomer,
  onPreviewPrintInvoice,
  onEditInvoice,
  onDuplicateInvoice,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [activeFilter, setActiveFilter] = useState<CustomerFilter>('all');

  // Add / Edit Modal State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingCustomer, setEditingCustomer] = useState<Customer | null>(null);
  const [formData, setFormData] = useState<Partial<Customer>>({
    name: '',
    phone: '',
    address: '',
    economicCode: '',
    nationalId: '',
    postalCode: '',
    defaultDiscount: 0,
  });

  // Customer Purchase History Modal State
  const [historyCustomer, setHistoryCustomer] = useState<Customer | null>(null);

  // Inline Delete Confirmation State (avoids blocking window.confirm)
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  // Toast Notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  const currencyUnit = settings.currencyUnit || 'تومان';

  // Compute purchase analytics per customer (normalized by trimmed lowercase name)
  const statsByCustomerId = useMemo(() => {
    const map = new Map<string, CustomerStats>();
    for (const c of customers) {
      const normalizedName = (c.name || '').trim().toLowerCase();
      const matchedInvoices = invoices.filter(
        (inv) => (inv.customerName || '').trim().toLowerCase() === normalizedName
      );
      const totalSpent = matchedInvoices.reduce((acc, inv) => acc + (inv.totalAmount || 0), 0);
      const lastInvoiceDate = matchedInvoices.length > 0 ? matchedInvoices[0].jalaliDate : null;
      map.set(c.id, {
        invoiceCount: matchedInvoices.length,
        totalSpent,
        lastInvoiceDate,
        customerInvoices: matchedInvoices,
      });
    }
    return map;
  }, [customers, invoices]);

  // Overall CRM KPIs
  const kpis = useMemo(() => {
    let officialCount = 0;
    let vipCount = 0;
    let totalCustomerRevenue = 0;
    for (const c of customers) {
      if (c.economicCode || c.nationalId) officialCount++;
      if (c.defaultDiscount && c.defaultDiscount > 0) vipCount++;
      const st = statsByCustomerId.get(c.id);
      if (st) totalCustomerRevenue += st.totalSpent;
    }
    return {
      totalCustomers: customers.length,
      officialCount,
      vipCount,
      totalCustomerRevenue,
    };
  }, [customers, statsByCustomerId]);

  // Filtered customers list
  const filteredCustomers = useMemo(() => {
    return customers.filter((c) => {
      const st = statsByCustomerId.get(c.id);
      if (activeFilter === 'official' && !c.economicCode && !c.nationalId) return false;
      if (activeFilter === 'vip' && (!c.defaultDiscount || c.defaultDiscount <= 0)) return false;
      if (activeFilter === 'with_orders' && (!st || st.invoiceCount === 0)) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const haystack = [
          c.name,
          c.phone,
          c.address,
          c.economicCode,
          c.nationalId,
          c.postalCode,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(q)) return false;
      }

      return true;
    });
  }, [customers, statsByCustomerId, activeFilter, searchQuery]);

  // Open modal for adding new customer
  const handleOpenAddModal = () => {
    setEditingCustomer(null);
    setFormData({
      name: '',
      phone: '',
      address: '',
      economicCode: '',
      nationalId: '',
      postalCode: '',
      defaultDiscount: 0,
    });
    setIsModalOpen(true);
  };

  // Open modal for editing existing customer
  const handleOpenEditModal = (customer: Customer) => {
    setEditingCustomer(customer);
    setFormData({
      id: customer.id,
      name: customer.name,
      phone: customer.phone || '',
      address: customer.address || '',
      economicCode: customer.economicCode || '',
      nationalId: customer.nationalId || '',
      postalCode: customer.postalCode || '',
      defaultDiscount: customer.defaultDiscount || 0,
    });
    setIsModalOpen(true);
  };

  // Save customer (Add or Edit)
  const handleSaveCustomer = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanName = (formData.name || '').trim();
    if (!cleanName) return;

    const payload: Customer = {
      id: editingCustomer?.id || 'cust_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      name: cleanName,
      phone: (formData.phone || '').trim() || undefined,
      address: (formData.address || '').trim() || undefined,
      economicCode: (formData.economicCode || '').trim() || undefined,
      nationalId: (formData.nationalId || '').trim() || undefined,
      postalCode: (formData.postalCode || '').trim() || undefined,
      defaultDiscount: Math.max(0, Number(formData.defaultDiscount) || 0),
    };

    StorageService.addOrUpdateCustomer(payload);
    onUpdateCustomers(StorageService.getCustomers());
    setIsModalOpen(false);
    showToast(
      editingCustomer
        ? `اطلاعات مشتری «${cleanName}» با موفقیت بروزرسانی شد.`
        : `مشتری جدید «${cleanName}» در سیستم ثبت شد.`
    );
  };

  // Execute customer deletion
  const handleExecuteDelete = (customer: Customer) => {
    StorageService.deleteCustomer(customer.id);
    onUpdateCustomers(StorageService.getCustomers());
    setConfirmDeleteId(null);
    showToast(`مشتری «${customer.name}» از فهرست حذف شد.`);
  };

  // Export Customers to CSV
  const handleExportCsv = () => {
    const csvContent = StorageService.exportCustomersToCsv();
    const defaultName = `customers_crm_${Date.now()}.csv`;
    if (typeof window !== 'undefined' && window.electronAPI?.exportFile) {
      window.electronAPI
        .exportFile({ defaultName, content: csvContent, extensions: ['csv'] })
        .then((res) => {
          if (res?.success) showToast('فهرست مشتریان با موفقیت در فایل اکسل/CSV ذخیره شد.');
        });
      return;
    }
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', defaultName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('فایل اکسل/CSV مشتریان با موفقیت دانلود شد.');
  };

  // RFC 4180 CSV Row Parser
  const parseCsvRow = (line: string): string[] => {
    const cells: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          cells.push(current.trim());
          current = '';
        } else {
          current += ch;
        }
      }
    }
    cells.push(current.trim());
    return cells;
  };

  // Import Customers from CSV
  const handleImportCsv = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const text = (evt.target?.result as string) || '';
        const lines = text.replace(/^\uFEFF/, '').trim().split(/\r?\n/);
        if (lines.length < 2) {
          showToast('فرمت فایل CSV مشتریان نامعتبر است.');
          return;
        }

        let importedCount = 0;
        for (let i = 1; i < lines.length; i++) {
          if (!lines[i].trim()) continue;
          const row = parseCsvRow(lines[i]);
          const name = (row[1] || row[0] || '').trim();
          if (!name) continue;
          StorageService.addOrUpdateCustomer({
            id: row[0]?.startsWith('cust_') ? row[0] : 'cust_' + Date.now() + '_' + i,
            name,
            phone: row[2] || undefined,
            economicCode: row[3] || undefined,
            nationalId: row[4] || undefined,
            postalCode: row[5] || undefined,
            defaultDiscount: Math.max(0, Number(row[6]) || 0),
            address: row[7] || undefined,
          });
          importedCount++;
        }

        if (importedCount > 0) {
          onUpdateCustomers(StorageService.getCustomers());
          showToast(`${toPersianDigits(importedCount)} مشتری با موفقیت از فایل CSV همگام‌سازی شد.`);
        }
      } catch {
        showToast('خطا در خواندن فایل CSV مشتریان.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 left-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-xl shadow-2xl flex items-center gap-3 animate-fade-in">
          <Check className="w-5 h-5" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Header Banner */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-800/80 p-6 rounded-2xl border border-slate-700 shadow-lg">
        <div>
          <h1 className="text-2xl font-bold text-white flex items-center gap-3">
            <Users className="w-7 h-7 text-blue-400" />
            <span>باشگاه و مدیریت مشتریان (CRM Lite)</span>
          </h1>
          <p className="text-slate-400 text-sm mt-1">
            مدیریت پرونده خریداران، مشخصات مالیاتی ماده ۱۶۹، تخفیف‌های اختصاصی و سوابق خرید هر مشتری
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={handleExportCsv}
            className="flex items-center gap-2 px-3.5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium transition-all cursor-pointer"
            title="خروجی اکسل / CSV از فهرست مشتریان"
          >
            <Download className="w-4 h-4" />
            <span>خروجی اکسل</span>
          </button>

          <label className="flex items-center gap-2 px-3.5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium transition-all cursor-pointer">
            <Upload className="w-4 h-4" />
            <span>ورود از CSV</span>
            <input type="file" accept=".csv" onChange={handleImportCsv} className="hidden" />
          </label>

          <button
            onClick={handleOpenAddModal}
            className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
          >
            <Plus className="w-4 h-4" />
            <span>ثبت مشتری جدید</span>
          </button>
        </div>
      </div>

      {/* KPI Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex items-center gap-4">
          <div className="p-3 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/30">
            <Users className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs text-slate-400 block">کل مشتریان ثبت‌شده</span>
            <strong className="text-xl font-bold text-white font-mono">
              {toPersianDigits(kpis.totalCustomers)}
            </strong>
          </div>
        </div>

        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex items-center gap-4">
          <div className="p-3 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-500/30">
            <ShieldCheck className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs text-slate-400 block">مشتریان رسمی (ماده ۱۶۹)</span>
            <strong className="text-xl font-bold text-amber-300 font-mono">
              {toPersianDigits(kpis.officialCount)}
            </strong>
          </div>
        </div>

        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex items-center gap-4">
          <div className="p-3 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
            <Percent className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs text-slate-400 block">دارای تخفیف اختصاصی</span>
            <strong className="text-xl font-bold text-emerald-400 font-mono">
              {toPersianDigits(kpis.vipCount)}
            </strong>
          </div>
        </div>

        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex items-center gap-4">
          <div className="p-3 rounded-xl bg-indigo-500/20 text-indigo-400 border border-indigo-500/30">
            <TrendingUp className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs text-slate-400 block">مجموع خرید مشتریان ({currencyUnit})</span>
            <strong className="text-base font-bold text-indigo-300 font-mono">
              {formatNumber(kpis.totalCustomerRevenue)}
            </strong>
          </div>
        </div>
      </div>

      {/* Search & Filter Controls */}
      <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700 flex flex-col lg:flex-row items-center justify-between gap-4">
        <div className="relative w-full lg:w-96">
          <Search className="w-5 h-5 absolute right-3 top-2.5 text-slate-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="جستجو با نام، تلفن، کد اقتصادی، شناسه ملی یا آدرس..."
            className="w-full bg-slate-900 border border-slate-700 rounded-xl pr-10 pl-4 py-2 text-sm text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none"
          />
        </div>

        <div className="flex flex-wrap items-center gap-2 w-full lg:w-auto justify-end">
          {(
            [
              { id: 'all', label: `همه (${toPersianDigits(customers.length)})` },
              { id: 'official', label: `رسمی ماده ۱۶۹ (${toPersianDigits(kpis.officialCount)})` },
              { id: 'vip', label: `دارای تخفیف (${toPersianDigits(kpis.vipCount)})` },
              { id: 'with_orders', label: 'دارای سابقه خرید' },
            ] as Array<{ id: CustomerFilter; label: string }>
          ).map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveFilter(tab.id)}
              className={`px-3.5 py-1.5 rounded-xl text-xs font-medium transition-all cursor-pointer ${
                activeFilter === tab.id
                  ? 'bg-blue-600 text-white shadow'
                  : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700/70'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      {/* Customers Table */}
      <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-right text-sm">
            <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
              <tr>
                <th className="py-3.5 px-4">نام مشتری / سازمان</th>
                <th className="py-3.5 px-4">اطلاعات تماس و نشانی</th>
                <th className="py-3.5 px-4">مشخصات رسمی (ماده ۱۶۹)</th>
                <th className="py-3.5 px-4 text-center">تخفیف پیش‌فرض</th>
                <th className="py-3.5 px-4 text-center">سابقه و حجم خرید</th>
                <th className="py-3.5 px-4 text-center">عملیات سریع</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-700/60">
              {filteredCustomers.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-slate-500">
                    {searchQuery || activeFilter !== 'all'
                      ? 'مشتری مطابق با فیلتر یا عبارت جستجو شده یافت نشد.'
                      : 'هنوز مشتری در سیستم ثبت نشده است. روی «ثبت مشتری جدید» کلیک کنید.'}
                  </td>
                </tr>
              ) : (
                filteredCustomers.map((c) => {
                  const stats = statsByCustomerId.get(c.id) || {
                    invoiceCount: 0,
                    totalSpent: 0,
                    lastInvoiceDate: null,
                    customerInvoices: [],
                  };
                  const isOfficialCustomer = Boolean(c.economicCode || c.nationalId);

                  return (
                    <tr key={c.id} className="hover:bg-slate-700/30 transition-colors">
                      {/* Name & Official Badge */}
                      <td className="py-3.5 px-4">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-white text-sm">{c.name}</span>
                          {isOfficialCustomer && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-500/15 text-amber-300 border border-amber-500/30">
                              <ShieldCheck className="w-3 h-3" />
                              <span>رسمی</span>
                            </span>
                          )}
                        </div>
                      </td>

                      {/* Phone & Address */}
                      <td className="py-3.5 px-4 text-xs space-y-1">
                        {c.phone ? (
                          <div className="flex items-center gap-1.5 text-slate-200 font-mono">
                            <Phone className="w-3.5 h-3.5 text-blue-400 shrink-0" />
                            <span>{toPersianDigits(c.phone)}</span>
                          </div>
                        ) : (
                          <span className="text-slate-500">بدون تلفن</span>
                        )}
                        {c.address && (
                          <div className="flex items-center gap-1.5 text-slate-400 max-w-xs truncate" title={c.address}>
                            <MapPin className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                            <span className="truncate">{c.address}</span>
                          </div>
                        )}
                      </td>

                      {/* Official Tax IDs */}
                      <td className="py-3.5 px-4 text-xs font-mono">
                        {isOfficialCustomer || c.postalCode ? (
                          <div className="space-y-0.5 text-slate-300">
                            {c.economicCode && (
                              <div>
                                <span className="text-slate-500 font-sans ml-1">اقتصادی:</span>
                                <span>{toPersianDigits(c.economicCode)}</span>
                              </div>
                            )}
                            {c.nationalId && (
                              <div>
                                <span className="text-slate-500 font-sans ml-1">شناسه/ملی:</span>
                                <span>{toPersianDigits(c.nationalId)}</span>
                              </div>
                            )}
                            {c.postalCode && (
                              <div>
                                <span className="text-slate-500 font-sans ml-1">کد پستی:</span>
                                <span>{toPersianDigits(c.postalCode)}</span>
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-500 font-sans">-</span>
                        )}
                      </td>

                      {/* Default Discount */}
                      <td className="py-3.5 px-4 text-center">
                        {c.defaultDiscount && c.defaultDiscount > 0 ? (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 font-mono text-xs font-bold">
                            {c.defaultDiscount <= 100
                              ? `${toPersianDigits(c.defaultDiscount)}٪ تخفیف`
                              : `${formatNumber(c.defaultDiscount)} ${currencyUnit}`}
                          </span>
                        ) : (
                          <span className="text-slate-500 text-xs">پله‌ای عادی</span>
                        )}
                      </td>

                      {/* Purchase Analytics */}
                      <td className="py-3.5 px-4 text-center">
                        {stats.invoiceCount > 0 ? (
                          <div className="space-y-0.5">
                            <div className="font-mono font-bold text-emerald-400 text-xs">
                              {formatNumber(stats.totalSpent)} {currencyUnit}
                            </div>
                            <div className="text-[11px] text-slate-400">
                              {toPersianDigits(stats.invoiceCount)} فاکتور | آخرین: {toPersianDigits(stats.lastInvoiceDate || '')}
                            </div>
                          </div>
                        ) : (
                          <span className="text-xs text-slate-500">بدون فاکتور ثبت‌شده</span>
                        )}
                      </td>

                      {/* Actions */}
                      <td className="py-3.5 px-4 text-center">
                        <div className="flex items-center justify-center gap-1.5 flex-wrap">
                          {/* Quick Issue Invoice */}
                          <button
                            onClick={() => onStartInvoiceForCustomer(c)}
                            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold shadow-sm transition-colors cursor-pointer"
                            title="صدور فوری پیش‌فاکتور جدید برای این مشتری"
                          >
                            <FilePlus className="w-3.5 h-3.5" />
                            <span>صدور فاکتور</span>
                          </button>

                          {/* Customer Invoice History */}
                          {stats.invoiceCount > 0 && (
                            <button
                              onClick={() => setHistoryCustomer(c)}
                              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/30 text-xs font-medium transition-colors cursor-pointer"
                              title="مشاهده فاکتورهای قبلی این مشتری"
                            >
                              <History className="w-3.5 h-3.5" />
                              <span>سوابق ({toPersianDigits(stats.invoiceCount)})</span>
                            </button>
                          )}

                          {/* Edit Customer */}
                          <button
                            onClick={() => handleOpenEditModal(c)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-amber-300 hover:bg-amber-500/10 transition-colors cursor-pointer"
                            title="ویرایش مشخصات مشتری"
                          >
                            <Edit2 className="w-4 h-4" />
                          </button>

                          {/* Inline Delete Confirmation */}
                          {confirmDeleteId === c.id ? (
                            <div className="flex items-center gap-1 bg-rose-950/60 border border-rose-500/40 rounded-lg px-2 py-1">
                              <button
                                onClick={() => handleExecuteDelete(c)}
                                className="text-[11px] font-bold text-rose-300 hover:text-white cursor-pointer"
                              >
                                حذف شود
                              </button>
                              <span className="text-slate-600">|</span>
                              <button
                                onClick={() => setConfirmDeleteId(null)}
                                className="text-[11px] text-slate-400 hover:text-white cursor-pointer"
                              >
                                لغو
                              </button>
                            </div>
                          ) : (
                            <button
                              onClick={() => setConfirmDeleteId(c.id)}
                              className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 transition-colors cursor-pointer"
                              title="حذف مشتری"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* MODAL 1: ADD / EDIT CUSTOMER */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm p-4">
          <form
            onSubmit={handleSaveCustomer}
            className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-xl p-6 space-y-5 shadow-2xl"
          >
            <div className="flex items-center justify-between border-b border-slate-700 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="p-2 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/30">
                  <Building2 className="w-5 h-5" />
                </div>
                <h3 className="text-lg font-bold text-white">
                  {editingCustomer ? `ویرایش پرونده مشتری: ${editingCustomer.name}` : 'ثبت پرونده مشتری جدید'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-700 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              <div className="sm:col-span-2">
                <label className="block text-slate-300 mb-1 font-medium">
                  نام شخص حقیقی / شرکت یا سازمان خریدار (الزامی):
                </label>
                <input
                  type="text"
                  required
                  autoFocus
                  value={formData.name || ''}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  placeholder="مثال: شرکت مهندسی آب و خاک غرب"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white focus:border-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">شماره تلفن / همراه:</label>
                <input
                  type="text"
                  value={formData.phone || ''}
                  onChange={(e) => setFormData({ ...formData, phone: e.target.value })}
                  placeholder="مثال: ۰۹۱۲۳۴۵۶۷۸۹"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">
                  تخفیف اختصاصی پیش‌فرض (% یا مبلغ ثابت):
                </label>
                <input
                  type="number"
                  min="0"
                  step="any"
                  value={formData.defaultDiscount || ''}
                  onChange={(e) => setFormData({ ...formData, defaultDiscount: Number(e.target.value) || 0 })}
                  placeholder="مثال: 5 برای ۵٪ تخفیف"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-emerald-400 font-mono font-bold text-left focus:border-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-amber-300/90 mb-1 font-medium">شماره اقتصادی (ماده ۱۶۹):</label>
                <input
                  type="text"
                  value={formData.economicCode || ''}
                  onChange={(e) => setFormData({ ...formData, economicCode: e.target.value })}
                  placeholder="مثال: ۴۱۱۴۵۶۷۸۹۱۲۳"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-amber-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-amber-300/90 mb-1 font-medium">شناسه ملی / کد ملی:</label>
                <input
                  type="text"
                  value={formData.nationalId || ''}
                  onChange={(e) => setFormData({ ...formData, nationalId: e.target.value })}
                  placeholder="مثال: ۱۰۱۰۳۵۶۷۸۹۰"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-amber-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-amber-300/90 mb-1 font-medium">کد پستی ۱۰ رقمی:</label>
                <input
                  type="text"
                  value={formData.postalCode || ''}
                  onChange={(e) => setFormData({ ...formData, postalCode: e.target.value })}
                  placeholder="مثال: ۱۵۴۸۷۶۳۲۱۱"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-amber-500 focus:outline-none"
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-slate-300 mb-1 font-medium">نشانی کامل دفتر / پروژه تحویل:</label>
                <input
                  type="text"
                  value={formData.address || ''}
                  onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                  placeholder="استان، شهر، شهرک صنعتی / نشانی پروژه..."
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-blue-500 focus:outline-none"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-700">
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="px-4 py-2 text-sm text-slate-300 hover:text-white cursor-pointer"
              >
                انصراف
              </button>
              <button
                type="submit"
                className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-bold shadow-lg shadow-blue-600/30 cursor-pointer"
              >
                <Check className="w-4 h-4" />
                <span>{editingCustomer ? 'ذخیره تغییرات مشتری' : 'ثبت مشتری در سیستم'}</span>
              </button>
            </div>
          </form>
        </div>
      )}

      {/* MODAL 2: CUSTOMER PURCHASE HISTORY */}
      {historyCustomer && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm p-4">
          <div className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-4xl max-h-[85vh] flex flex-col p-6 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-700 pb-3">
              <div>
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <History className="w-5 h-5 text-indigo-400" />
                  <span>سوابق پیش‌فاکتورهای «{historyCustomer.name}»</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  مشاهده، چاپ مجدد، ویرایش یا تکثیر فاکتورهای صادر شده برای این مشتری
                </p>
              </div>
              <button
                onClick={() => setHistoryCustomer(null)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-700 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="overflow-y-auto flex-1">
              <table className="w-full text-right text-sm">
                <thead className="bg-slate-900 text-xs text-slate-400 border-b border-slate-700 sticky top-0">
                  <tr>
                    <th className="py-3 px-4 text-center">شماره فاکتور</th>
                    <th className="py-3 px-4">تاریخ صدور</th>
                    <th className="py-3 px-4 text-center">تعداد اقلام</th>
                    <th className="py-3 px-4 text-center">مبلغ نهایی</th>
                    <th className="py-3 px-4 text-center">عملیات</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {(statsByCustomerId.get(historyCustomer.id)?.customerInvoices || []).map((inv) => (
                    <tr key={inv.id} className="hover:bg-slate-700/30">
                      <td className="py-3 px-4 text-center font-mono font-bold text-blue-400">
                        {toPersianDigits(inv.invoiceNumber)}
                      </td>
                      <td className="py-3 px-4 font-mono text-slate-300">
                        {toPersianDigits(inv.jalaliDate)}
                      </td>
                      <td className="py-3 px-4 text-center font-mono text-slate-300">
                        {toPersianDigits(inv.items.length)} ردیف
                      </td>
                      <td className="py-3 px-4 text-center font-mono font-bold text-emerald-400">
                        {formatNumber(inv.totalAmount)} {inv.currencyUnit || currencyUnit}
                      </td>
                      <td className="py-3 px-4 text-center">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            onClick={() => {
                              setHistoryCustomer(null);
                              onPreviewPrintInvoice(inv);
                            }}
                            className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-blue-600/20 hover:bg-blue-600/40 text-blue-300 border border-blue-500/30 text-xs cursor-pointer"
                          >
                            <Printer className="w-3.5 h-3.5" />
                            <span>چاپ / PDF</span>
                          </button>
                          <button
                            onClick={() => {
                              setHistoryCustomer(null);
                              onEditInvoice(inv);
                            }}
                            className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-600/20 hover:bg-amber-600/40 text-amber-300 border border-amber-500/30 text-xs cursor-pointer"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                            <span>ویرایش</span>
                          </button>
                          <button
                            onClick={() => {
                              setHistoryCustomer(null);
                              onDuplicateInvoice(inv);
                            }}
                            className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/30 text-xs cursor-pointer"
                          >
                            <Copy className="w-3.5 h-3.5" />
                            <span>تکثیر</span>
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="flex items-center justify-between pt-3 border-t border-slate-700">
              <button
                onClick={() => {
                  const cust = historyCustomer;
                  setHistoryCustomer(null);
                  onStartInvoiceForCustomer(cust);
                }}
                className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold cursor-pointer"
              >
                <FilePlus className="w-4 h-4" />
                <span>صدور پیش‌فاکتور جدید برای {historyCustomer.name}</span>
              </button>
              <button
                onClick={() => setHistoryCustomer(null)}
                className="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium cursor-pointer"
              >
                بستن
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
