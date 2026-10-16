import React, { useState, useMemo } from 'react';
import {
  Receipt,
  Wallet,
  CreditCard,
  TrendingUp,
  CheckCircle2,
  AlertCircle,
  Calendar,
  Building2,
  Search,
  Plus,
  Filter,
  Download,
  Printer,
  Copy,
  Edit2,
  Trash2,
  FileText,
  Check,
  X,
  Clock,
  FileSpreadsheet,
  Layers,
  Scale,
  Phone,
  ArrowUpRight,
  ArrowDownLeft,
  Share2,
} from 'lucide-react';
import {
  CompanySettings,
  Customer,
  Invoice,
  PaymentMethod,
  PaymentRecord,
  PaymentStatus,
} from '../../types';
import { StorageService } from '../../services/storage';
import {
  AccountingService,
  CustomerStatementReport,
} from '../../services/accountingService';
import {
  formatNumber,
  getCurrentJalaliDate,
  numberToPersianWords,
  parseFormattedNumber,
  toPersianDigits,
} from '../../services/persianUtils';

interface AccountingManagerProps {
  customers: Customer[];
  invoices: Invoice[];
  payments: PaymentRecord[];
  onUpdatePayments: (payments: PaymentRecord[]) => void;
  settings: CompanySettings;
  onNavigateToCustomer?: (customerId: string) => void;
  onStartInvoiceForCustomer: (customer: Customer) => void;
  onPreviewPrintInvoice: (invoice: Invoice) => void;
}

type AccountingTab = 'customers' | 'payments' | 'invoices' | 'tax';
type CustomerFilter = 'all' | 'debtor' | 'settled' | 'pending_cheques' | 'official';
type PaymentFilter = 'all' | 'transfer' | 'cheque' | 'pos' | 'cash';

export const AccountingManager: React.FC<AccountingManagerProps> = ({
  customers,
  invoices,
  payments,
  onUpdatePayments,
  settings,
  onStartInvoiceForCustomer,
  onPreviewPrintInvoice,
}) => {
  const [activeTab, setActiveTab] = useState<AccountingTab>('customers');
  const [searchQuery, setSearchQuery] = useState('');

  // Sub-filters
  const [customerFilter, setCustomerFilter] = useState<CustomerFilter>('all');
  const [paymentMethodFilter, setPaymentMethodFilter] = useState<PaymentFilter>('all');
  const [paymentStatusFilter, setPaymentStatusFilter] = useState<'all' | PaymentStatus>('all');
  const [invoiceSettlementFilter, setInvoiceSettlementFilter] = useState<'all' | 'unpaid' | 'partially_paid' | 'fully_paid'>('all');

  // Modals state
  const [isPaymentModalOpen, setIsPaymentModalOpen] = useState(false);
  const [editingPayment, setEditingPayment] = useState<PaymentRecord | null>(null);
  const [activeStatementCustomer, setActiveStatementCustomer] = useState<Customer | null>(null);
  const [activeVoucherPayment, setActiveVoucherPayment] = useState<PaymentRecord | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

  // Toast notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  const currencyUnit = settings.currencyUnit || 'تومان';
  const currentDate = getCurrentJalaliDate();

  // Form State for Payment Modal
  const [paymentForm, setPaymentForm] = useState<{
    customerId: string;
    customerName: string;
    invoiceId: string;
    invoiceNumber: string;
    amount: string;
    paymentDate: string;
    method: PaymentMethod;
    referenceNumber: string;
    bankName: string;
    chequeDueDate: string;
    status: PaymentStatus;
    notes: string;
  }>({
    customerId: '',
    customerName: '',
    invoiceId: '',
    invoiceNumber: '',
    amount: '',
    paymentDate: currentDate,
    method: 'transfer',
    referenceNumber: '',
    bankName: '',
    chequeDueDate: '',
    status: 'cleared',
    notes: '',
  });

  // Calculate Customer Summaries
  const customerSummaries = useMemo(() => {
    return AccountingService.computeCustomerSummaries(customers, invoices, payments);
  }, [customers, invoices, payments]);

  const customerSummariesMap = useMemo(() => {
    const map = new Map<string, (typeof customerSummaries)[0]>();
    customerSummaries.forEach((s) => map.set(s.customerId, s));
    return map;
  }, [customerSummaries]);

  // Calculate Invoice Settlement List
  const invoiceSettlements = useMemo(() => {
    return AccountingService.computeInvoiceSettlementList(invoices, payments);
  }, [invoices, payments]);

  // Overall Financial KPIs
  const kpis = useMemo(() => {
    const totalSales = invoices.reduce((sum, inv) => sum + (Number(inv.totalAmount) || 0), 0);
    const clearedPayments = payments.filter((p) => p.status === 'cleared');
    const totalReceived = clearedPayments.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

    const pendingChequesList = payments.filter((p) => p.status === 'pending' && p.method === 'cheque');
    const totalPendingCheques = pendingChequesList.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);

    // Total outstanding receivables from debtor customers
    let totalReceivables = 0;
    let debtorCount = 0;
    customerSummaries.forEach((s) => {
      if (s.balance > 1) {
        totalReceivables += s.balance;
        debtorCount++;
      }
    });

    // Collection rate percentage
    const collectionRate = totalSales > 0 ? Math.min(100, Math.round((totalReceived / totalSales) * 100)) : 100;

    // Official VAT (10%)
    const officialInvoices = invoices.filter((i) => Boolean(i.isOfficial));
    const officialVatTotal = officialInvoices.reduce((sum, i) => sum + (Number(i.taxAmount) || 0), 0);

    return {
      totalSales,
      totalReceived,
      totalReceivables,
      debtorCount,
      totalPendingCheques,
      pendingChequeCount: pendingChequesList.length,
      collectionRate,
      officialVatTotal,
      officialCount: officialInvoices.length,
      totalInvoicesCount: invoices.length,
      totalPaymentsCount: payments.length,
    };
  }, [invoices, payments, customerSummaries]);

  // Filtered Customer Summaries
  const filteredCustomerSummaries = useMemo(() => {
    return customerSummaries.filter((s) => {
      const cust = customers.find((c) => c.id === s.customerId);
      if (customerFilter === 'debtor' && s.balanceStatus !== 'debtor') return false;
      if (customerFilter === 'settled' && s.balanceStatus !== 'settled') return false;
      if (customerFilter === 'pending_cheques' && s.totalPendingCheques <= 0) return false;
      if (customerFilter === 'official' && !cust?.economicCode && !cust?.nationalId) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const haystack = [
          s.customerName,
          cust?.phone,
          cust?.economicCode,
          cust?.nationalId,
          cust?.address,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(q)) return false;
      }

      return true;
    });
  }, [customerSummaries, customers, customerFilter, searchQuery]);

  // Filtered Payments Journal
  const filteredPayments = useMemo(() => {
    return payments.filter((p) => {
      if (paymentMethodFilter !== 'all' && p.method !== paymentMethodFilter) return false;
      if (paymentStatusFilter !== 'all' && p.status !== paymentStatusFilter) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const haystack = [
          p.customerName,
          p.referenceNumber,
          p.bankName,
          p.invoiceNumber,
          p.notes,
          p.paymentDate,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(q)) return false;
      }

      return true;
    });
  }, [payments, paymentMethodFilter, paymentStatusFilter, searchQuery]);

  // Filtered Invoice Settlements
  const filteredInvoiceSettlements = useMemo(() => {
    return invoiceSettlements.filter((item) => {
      if (invoiceSettlementFilter !== 'all' && item.status !== invoiceSettlementFilter) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const haystack = [
          item.invoiceNumber,
          item.customerName,
          item.date,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!haystack.includes(q)) return false;
      }

      return true;
    });
  }, [invoiceSettlements, invoiceSettlementFilter, searchQuery]);

  // Open Payment Modal for adding or editing
  const handleOpenPaymentModal = (preset?: {
    customer?: Customer;
    invoice?: Invoice;
    payment?: PaymentRecord;
  }) => {
    if (preset?.payment) {
      setEditingPayment(preset.payment);
      setPaymentForm({
        customerId: preset.payment.customerId || '',
        customerName: preset.payment.customerName || '',
        invoiceId: preset.payment.invoiceId || '',
        invoiceNumber: preset.payment.invoiceNumber || '',
        amount: formatNumber(preset.payment.amount),
        paymentDate: preset.payment.paymentDate || currentDate,
        method: preset.payment.method || 'transfer',
        referenceNumber: preset.payment.referenceNumber || '',
        bankName: preset.payment.bankName || '',
        chequeDueDate: preset.payment.chequeDueDate || '',
        status: preset.payment.status || 'cleared',
        notes: preset.payment.notes || '',
      });
    } else {
      setEditingPayment(null);
      const cust = preset?.customer || (preset?.invoice ? customers.find((c) => c.name === preset.invoice?.customerName) : null);
      const inv = preset?.invoice;
      const invSettlement = inv ? invoiceSettlements.find((s) => s.invoiceId === inv.id) : null;
      const suggestedAmount = invSettlement ? invSettlement.remainingAmount : 0;

      setPaymentForm({
        customerId: cust?.id || '',
        customerName: cust?.name || inv?.customerName || (customers[0]?.name ?? ''),
        invoiceId: inv?.id || '',
        invoiceNumber: inv?.invoiceNumber || '',
        amount: suggestedAmount > 0 ? formatNumber(suggestedAmount) : '',
        paymentDate: currentDate,
        method: 'transfer',
        referenceNumber: '',
        bankName: settings.bankName?.split(' ')[1] || 'ملت',
        chequeDueDate: '',
        status: 'cleared',
        notes: inv ? `بابت تسویه پیش‌فاکتور شماره ${inv.invoiceNumber}` : '',
      });
    }
    setIsPaymentModalOpen(true);
  };

  // Handle Customer Selection in Payment Modal
  const handleCustomerSelectInModal = (selectedName: string) => {
    const cust = customers.find((c) => c.name === selectedName);
    setPaymentForm((prev) => ({
      ...prev,
      customerName: selectedName,
      customerId: cust?.id || '',
      invoiceId: '',
      invoiceNumber: '',
    }));
  };

  // Invoices for selected customer in modal (for linking)
  const candidateInvoicesForSelectedCustomer = useMemo(() => {
    if (!paymentForm.customerName) return [];
    const norm = AccountingService.normalize(paymentForm.customerName);
    return invoices.filter((i) => AccountingService.normalize(i.customerName) === norm);
  }, [paymentForm.customerName, invoices]);

  // Handle Invoice Selection in Modal
  const handleInvoiceSelectInModal = (invoiceId: string) => {
    if (!invoiceId) {
      setPaymentForm((prev) => ({
        ...prev,
        invoiceId: '',
        invoiceNumber: '',
      }));
      return;
    }
    const inv = invoices.find((i) => i.id === invoiceId);
    if (!inv) return;

    const settlement = invoiceSettlements.find((s) => s.invoiceId === inv.id);
    const remaining = settlement ? settlement.remainingAmount : inv.totalAmount;

    setPaymentForm((prev) => ({
      ...prev,
      invoiceId: inv.id,
      invoiceNumber: inv.invoiceNumber,
      amount: remaining > 0 ? formatNumber(remaining) : prev.amount,
      notes: prev.notes || `بابت تسویه فاکتور شماره ${inv.invoiceNumber}`,
    }));
  };

  // Save Payment Record (Add / Edit)
  const handleSavePayment = (e: React.FormEvent) => {
    e.preventDefault();
    const rawAmount = parseFormattedNumber(paymentForm.amount);
    if (rawAmount <= 0) {
      alert('لطفاً مبلغ معتبر برای سند دریافت وارد نمایید.');
      return;
    }
    if (!paymentForm.customerName.trim()) {
      alert('لطفاً نام مشتری را مشخص کنید.');
      return;
    }

    const payload: PaymentRecord = {
      id: editingPayment?.id || 'pay_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      customerId: paymentForm.customerId || (customers.find((c) => c.name === paymentForm.customerName)?.id || ''),
      customerName: paymentForm.customerName.trim(),
      invoiceId: paymentForm.invoiceId || undefined,
      invoiceNumber: paymentForm.invoiceNumber || undefined,
      amount: rawAmount,
      currencyUnit: (settings.currencyUnit as any) || 'تومان',
      paymentDate: paymentForm.paymentDate || currentDate,
      method: paymentForm.method,
      referenceNumber: paymentForm.referenceNumber.trim() || undefined,
      bankName: paymentForm.bankName.trim() || undefined,
      chequeDueDate: paymentForm.method === 'cheque' ? paymentForm.chequeDueDate.trim() || undefined : undefined,
      status: paymentForm.status,
      notes: paymentForm.notes.trim() || undefined,
      createdAt: editingPayment?.createdAt || new Date().toISOString(),
    };

    if (editingPayment) {
      StorageService.updatePayment(payload);
      showToast(`سند دریافت شماره «${payload.referenceNumber || payload.id.slice(-6)}» با موفقیت ویرایش شد.`);
    } else {
      StorageService.addPayment(payload);
      showToast(`سند دریافت به مبلغ ${formatNumber(payload.amount)} ${currencyUnit} با موفقیت ثبت شد.`);
    }

    onUpdatePayments(StorageService.getPayments());
    setIsPaymentModalOpen(false);
  };

  // Delete Payment Record
  const handleDeletePayment = (id: string) => {
    StorageService.deletePayment(id);
    onUpdatePayments(StorageService.getPayments());
    setDeleteConfirmId(null);
    showToast('سند دریافت از سیستم حذف شد.');
  };

  // Quick toggle payment status
  const handleToggleStatus = (p: PaymentRecord, newStatus: PaymentStatus) => {
    const updated = { ...p, status: newStatus };
    StorageService.updatePayment(updated);
    onUpdatePayments(StorageService.getPayments());
    showToast(`وضعیت سند به «${newStatus === 'cleared' ? 'وصول شده' : newStatus === 'pending' ? 'در جریان وصول' : 'برگشتی'}» تغییر یافت.`);
  };

  // Export Journal to CSV
  const handleExportPaymentsCsv = () => {
    const csvContent = StorageService.exportPaymentsToCsv();
    const defaultName = `payments_journal_${Date.now()}.csv`;
    if (typeof window !== 'undefined' && window.electronAPI?.exportFile) {
      window.electronAPI
        .exportFile({ defaultName, content: csvContent, extensions: ['csv'] })
        .then((res) => {
          if (res?.success) showToast('دفتر روزنامه دریافت‌ها با موفقیت ذخیره شد.');
        });
      return;
    }
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = defaultName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('فایل اکسل دفتر روزنامه دریافت‌ها دانلود شد.');
  };

  // Export Official Tax Report to CSV
  const handleExportTaxCsv = () => {
    const csvContent = AccountingService.exportOfficialTaxToCsv(invoices, settings);
    const defaultName = `official_tax_vat_169_${Date.now()}.csv`;
    if (typeof window !== 'undefined' && window.electronAPI?.exportFile) {
      window.electronAPI
        .exportFile({ defaultName, content: csvContent, extensions: ['csv'] })
        .then((res) => {
          if (res?.success) showToast('گزارش رسمی مالیات و ارزش افزوده ماده ۱۶۹ ذخیره شد.');
        });
      return;
    }
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = defaultName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('فایل اکسل گزارش فصلی ارزش افزوده ماده ۱۶۹ دانلود شد.');
  };

  // Method Labels and Badges
  const getMethodBadge = (method: PaymentMethod) => {
    switch (method) {
      case 'transfer':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-blue-500/15 text-blue-300 border border-blue-500/30">
            <ArrowDownLeft className="w-3 h-3" />
            <span>حواله بانکی</span>
          </span>
        );
      case 'cheque':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-indigo-500/15 text-indigo-300 border border-indigo-500/30">
            <CreditCard className="w-3 h-3" />
            <span>چک صیادی</span>
          </span>
        );
      case 'pos':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">
            <Wallet className="w-3 h-3" />
            <span>دستگاه پوز / کارت</span>
          </span>
        );
      case 'cash':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-500/15 text-amber-300 border border-amber-500/30">
            <Receipt className="w-3 h-3" />
            <span>نقدی</span>
          </span>
        );
      case 'deposit':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-purple-500/15 text-purple-300 border border-purple-500/30">
            <Layers className="w-3 h-3" />
            <span>پیش‌پرداخت / بیعانه</span>
          </span>
        );
      default:
        return null;
    }
  };

  const getStatusBadge = (status: PaymentStatus) => {
    switch (status) {
      case 'cleared':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">
            <CheckCircle2 className="w-3 h-3" />
            <span>وصول شده</span>
          </span>
        );
      case 'pending':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/40">
            <Clock className="w-3 h-3" />
            <span>در جریان وصول</span>
          </span>
        );
      case 'bounced':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-500/20 text-rose-300 border border-rose-500/40">
            <AlertCircle className="w-3 h-3" />
            <span>برگشت خورده</span>
          </span>
        );
    }
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 left-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-xl shadow-2xl flex items-center gap-3 animate-fade-in">
          <Check className="w-5 h-5" />
          <span className="text-sm font-medium">{toastMessage}</span>
        </div>
      )}

      {/* Header Banner */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-800/80 p-6 rounded-2xl border border-slate-700 shadow-lg">
        <div>
          <h1 className="text-2xl font-bold text-white flex items-center gap-3">
            <Receipt className="w-7 h-7 text-emerald-400" />
            <span>مدیریت مالی و حسابداری (CRM Lite)</span>
          </h1>
          <p className="text-slate-400 text-sm mt-1">
            دفتر حساب معین خریداران، ثبت واریزی‌ها و چک‌های صیادی، تراز مطالبات و گزارش رسمی ارزش افزوده ماده ۱۶۹
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={handleExportPaymentsCsv}
            className="flex items-center gap-2 px-3.5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium transition-all cursor-pointer"
            title="خروجی اکسل از کلیه واریزی‌ها و دریافت‌ها"
          >
            <Download className="w-4 h-4" />
            <span>اکسل دریافت‌ها</span>
          </button>

          <button
            onClick={handleExportTaxCsv}
            className="flex items-center gap-2 px-3.5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium transition-all cursor-pointer"
            title="خروجی اکسل گزارش فصلی مالیاتی ماده ۱۶۹"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>گزارش ماده ۱۶۹</span>
          </button>

          <button
            onClick={() => handleOpenPaymentModal()}
            className="flex items-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-sm font-bold shadow-lg shadow-emerald-600/30 transition-all cursor-pointer"
          >
            <Plus className="w-4 h-4" />
            <span>ثبت سند دریافت جدید</span>
          </button>
        </div>
      </div>

      {/* KPI Financial Overview Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
        {/* 1. Total Sales */}
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-400 text-xs mb-2">
            <span>مجموع فروش و فاکتورها</span>
            <div className="p-2 rounded-xl bg-blue-500/20 text-blue-400">
              <TrendingUp className="w-4 h-4" />
            </div>
          </div>
          <div className="text-lg font-bold text-white font-mono">
            {formatNumber(kpis.totalSales)}{' '}
            <span className="text-xs font-sans text-slate-400">{currencyUnit}</span>
          </div>
          <span className="text-[11px] text-slate-400 mt-1">
            {toPersianDigits(kpis.totalInvoicesCount)} پیش‌فاکتور صادر شده
          </span>
        </div>

        {/* 2. Total Realized Collections */}
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-400 text-xs mb-2">
            <span>مجموع وجوه وصول‌شده</span>
            <div className="p-2 rounded-xl bg-emerald-500/20 text-emerald-400">
              <CheckCircle2 className="w-4 h-4" />
            </div>
          </div>
          <div className="text-lg font-bold text-emerald-400 font-mono">
            {formatNumber(kpis.totalReceived)}{' '}
            <span className="text-xs font-sans text-slate-400">{currencyUnit}</span>
          </div>
          <div className="flex items-center gap-1.5 mt-1 text-[11px] text-emerald-400 font-medium">
            <span>نرخ تسویه: {toPersianDigits(kpis.collectionRate)}٪</span>
          </div>
        </div>

        {/* 3. Total Outstanding Receivables */}
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-400 text-xs mb-2">
            <span>کل مطالبات معوق (بدهی خریداران)</span>
            <div className="p-2 rounded-xl bg-rose-500/20 text-rose-400">
              <Scale className="w-4 h-4" />
            </div>
          </div>
          <div className="text-lg font-bold text-rose-400 font-mono">
            {formatNumber(kpis.totalReceivables)}{' '}
            <span className="text-xs font-sans text-slate-400">{currencyUnit}</span>
          </div>
          <span className="text-[11px] text-rose-300/80 mt-1">
            {toPersianDigits(kpis.debtorCount)} مشتری دارای مانده بدهی
          </span>
        </div>

        {/* 4. Pending Cheques */}
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-400 text-xs mb-2">
            <span>چک‌های در جریان وصول</span>
            <div className="p-2 rounded-xl bg-indigo-500/20 text-indigo-400">
              <CreditCard className="w-4 h-4" />
            </div>
          </div>
          <div className="text-lg font-bold text-indigo-300 font-mono">
            {formatNumber(kpis.totalPendingCheques)}{' '}
            <span className="text-xs font-sans text-slate-400">{currencyUnit}</span>
          </div>
          <span className="text-[11px] text-indigo-400/80 mt-1">
            {toPersianDigits(kpis.pendingChequeCount)} فقره چک سررسیدنشده
          </span>
        </div>

        {/* 5. Official VAT Article 169 */}
        <div className="bg-slate-800/80 p-4 rounded-2xl border border-slate-700/80 flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-400 text-xs mb-2">
            <span>مالیات بر ارزش افزوده (۱۰٪)</span>
            <div className="p-2 rounded-xl bg-cyan-500/20 text-cyan-400">
              <Building2 className="w-4 h-4" />
            </div>
          </div>
          <div className="text-lg font-bold text-cyan-300 font-mono">
            {formatNumber(kpis.officialVatTotal)}{' '}
            <span className="text-xs font-sans text-slate-400">{currencyUnit}</span>
          </div>
          <span className="text-[11px] text-cyan-400/80 mt-1">
            {toPersianDigits(kpis.officialCount)} فاکتور رسمی ماده ۱۶۹
          </span>
        </div>
      </div>

      {/* Sub-Navigation Tabs */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-700/80 pb-3">
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setActiveTab('customers')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'customers'
                ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                : 'bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700'
            }`}
          >
            <Scale className="w-4 h-4" />
            <span>تراز و مانده حساب خریداران</span>
            <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-white/20 font-mono">
              {toPersianDigits(customerSummaries.length)}
            </span>
          </button>

          <button
            onClick={() => setActiveTab('payments')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'payments'
                ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-600/25'
                : 'bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700'
            }`}
          >
            <Wallet className="w-4 h-4" />
            <span>دفتر ثبت دریافت‌ها و اسناد</span>
            <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-white/20 font-mono">
              {toPersianDigits(payments.length)}
            </span>
          </button>

          <button
            onClick={() => setActiveTab('invoices')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'invoices'
                ? 'bg-indigo-600 text-white shadow-lg shadow-indigo-600/25'
                : 'bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700'
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>وضعیت تسویه فاکتورها</span>
            <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-white/20 font-mono">
              {toPersianDigits(invoices.length)}
            </span>
          </button>

          <button
            onClick={() => setActiveTab('tax')}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'tax'
                ? 'bg-cyan-600 text-white shadow-lg shadow-cyan-600/25'
                : 'bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700'
            }`}
          >
            <Building2 className="w-4 h-4" />
            <span>گزارش فصلی ارزش افزوده (ماده ۱۶۹)</span>
            <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-white/20 font-mono">
              {toPersianDigits(kpis.officialCount)}
            </span>
          </button>
        </div>

        {/* Global Search Input */}
        <div className="relative w-full sm:w-72">
          <Search className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="جستجو در نام، شماره سند، کد پیگیری..."
            className="w-full bg-slate-900 border border-slate-700 rounded-xl pr-9 pl-3 py-1.5 text-xs text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none"
          />
        </div>
      </div>

      {/* ============================================================== */}
      {/* TAB 1: CUSTOMERS ACCOUNTS LEDGER & RECEIVABLES                  */}
      {/* ============================================================== */}
      {activeTab === 'customers' && (
        <div className="space-y-4">
          {/* Sub Filters */}
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="flex items-center gap-2">
              <Filter className="w-4 h-4 text-slate-400" />
              <span className="text-xs text-slate-400">فیلتر خریداران:</span>
              {(
                [
                  { id: 'all', label: `همه (${toPersianDigits(customerSummaries.length)})` },
                  { id: 'debtor', label: `بدهکاران معوق (${toPersianDigits(kpis.debtorCount)})` },
                  { id: 'settled', label: 'تسویه‌شده‌ها' },
                  { id: 'pending_cheques', label: `دارای چک در جریان (${toPersianDigits(kpis.pendingChequeCount)})` },
                  { id: 'official', label: `رسمی ماده ۱۶۹ (${toPersianDigits(kpis.officialCount)})` },
                ] as Array<{ id: CustomerFilter; label: string }>
              ).map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setCustomerFilter(tab.id)}
                  className={`px-3 py-1 rounded-xl text-xs font-medium transition-all cursor-pointer ${
                    customerFilter === tab.id
                      ? 'bg-blue-600 text-white shadow'
                      : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700/70'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            <button
              onClick={() => {
                const csv = StorageService.exportCustomersToCsv();
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `customers_ledger_${Date.now()}.csv`;
                a.click();
                URL.revokeObjectURL(url);
              }}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>خروجی اکسل مانده حساب‌ها</span>
            </button>
          </div>

          {/* Table */}
          <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-right text-sm">
                <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
                  <tr>
                    <th className="py-3.5 px-4">نام مشتری / سازمان خریدار</th>
                    <th className="py-3.5 px-4">مشخصات تماس و اقتصادی</th>
                    <th className="py-3.5 px-4 text-center">مجموع فاکتورها (بدهکار)</th>
                    <th className="py-3.5 px-4 text-center">مجموع دریافتی‌ها (بستانکار)</th>
                    <th className="py-3.5 px-4 text-center">اسناد در جریان (چک)</th>
                    <th className="py-3.5 px-4 text-center">مانده نهایی حساب</th>
                    <th className="py-3.5 px-4 text-center">عملیات مالی</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {filteredCustomerSummaries.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-12 text-center text-slate-500">
                        مشتری مطابق با فیلتر جستجو شده یافت نشد.
                      </td>
                    </tr>
                  ) : (
                    filteredCustomerSummaries.map((summary) => {
                      const cust = customers.find((c) => c.id === summary.customerId);
                      const isDebtor = summary.balanceStatus === 'debtor';
                      const isSettled = summary.balanceStatus === 'settled';
                      const isCreditor = summary.balanceStatus === 'creditor';

                      return (
                        <tr key={summary.customerId} className="hover:bg-slate-700/30 transition-colors">
                          <td className="py-3.5 px-4">
                            <div className="font-bold text-white text-sm">{summary.customerName}</div>
                            {cust?.address && (
                              <div className="text-xs text-slate-400 truncate max-w-xs">{cust.address}</div>
                            )}
                          </td>

                          <td className="py-3.5 px-4 text-xs font-mono space-y-0.5">
                            {cust?.phone && (
                              <div className="flex items-center gap-1 text-slate-300">
                                <Phone className="w-3 h-3 text-blue-400 shrink-0" />
                                <span>{toPersianDigits(cust.phone)}</span>
                              </div>
                            )}
                            {cust?.nationalId && (
                              <div className="text-slate-400">
                                <span className="font-sans ml-1 text-slate-500">ملی:</span>
                                <span>{toPersianDigits(cust.nationalId)}</span>
                              </div>
                            )}
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono font-bold text-slate-200">
                            <div>{formatNumber(summary.totalInvoiced)} {currencyUnit}</div>
                            <span className="text-[11px] font-sans font-normal text-slate-400">
                              {toPersianDigits(summary.invoiceCount)} فاکتور
                            </span>
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono font-bold text-emerald-400">
                            <div>{formatNumber(summary.totalPaid)} {currencyUnit}</div>
                            <span className="text-[11px] font-sans font-normal text-slate-400">
                              {toPersianDigits(summary.paymentCount)} سند دریافت
                            </span>
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono text-xs">
                            {summary.totalPendingCheques > 0 ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-indigo-500/15 text-indigo-300 border border-indigo-500/30 font-bold">
                                {formatNumber(summary.totalPendingCheques)} {currencyUnit}
                              </span>
                            ) : (
                              <span className="text-slate-500">-</span>
                            )}
                          </td>

                          <td className="py-3.5 px-4 text-center">
                            {isSettled ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                <span>تسویه کامل</span>
                              </span>
                            ) : isDebtor ? (
                              <div className="inline-flex flex-col items-center">
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-rose-500/20 text-rose-300 border border-rose-500/30 font-mono">
                                  <span>بدهکار: {formatNumber(summary.balance)}</span>
                                </span>
                              </div>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30 font-mono">
                                <span>بستانکار: {formatNumber(Math.abs(summary.balance))}</span>
                              </span>
                            )}
                          </td>

                          <td className="py-3.5 px-4 text-center">
                            <div className="flex items-center justify-center gap-1.5 flex-wrap">
                              {/* Quick Record Payment */}
                              <button
                                onClick={() => handleOpenPaymentModal({ customer: cust })}
                                className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold shadow-sm transition-colors cursor-pointer"
                                title="ثبت سند دریافت جدید برای این مشتری"
                              >
                                <Plus className="w-3.5 h-3.5" />
                                <span>ثبت دریافت</span>
                              </button>

                              {/* Customer Account Statement */}
                              <button
                                onClick={() => setActiveStatementCustomer(cust || { id: summary.customerId, name: summary.customerName })}
                                className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/30 text-xs font-medium transition-colors cursor-pointer"
                                title="مشاهده و چاپ صورت‌حساب رسمی معین خریدار"
                              >
                                <FileText className="w-3.5 h-3.5" />
                                <span>صورت‌حساب</span>
                              </button>

                              {/* Quick New Invoice */}
                              {cust && (
                                <button
                                  onClick={() => onStartInvoiceForCustomer(cust)}
                                  className="p-1.5 rounded-lg text-slate-400 hover:text-blue-300 hover:bg-blue-500/10 transition-colors cursor-pointer"
                                  title="صدور پیش‌فاکتور جدید برای این مشتری"
                                >
                                  <ArrowUpRight className="w-4 h-4" />
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
        </div>
      )}

      {/* ============================================================== */}
      {/* TAB 2: RECEIPTS & CASH JOURNAL                                 */}
      {/* ============================================================== */}
      {activeTab === 'payments' && (
        <div className="space-y-4">
          {/* Sub Filters & Actions */}
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="flex items-center gap-2 flex-wrap">
              <Filter className="w-4 h-4 text-slate-400" />
              <span className="text-xs text-slate-400">شیوه دریافت:</span>
              {(
                [
                  { id: 'all', label: 'همه روش‌ها' },
                  { id: 'transfer', label: 'حواله بانکی' },
                  { id: 'cheque', label: 'چک صیادی' },
                  { id: 'pos', label: 'کارت‌خوان' },
                  { id: 'cash', label: 'نقدی' },
                ] as Array<{ id: PaymentFilter; label: string }>
              ).map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setPaymentMethodFilter(tab.id)}
                  className={`px-3 py-1 rounded-xl text-xs font-medium transition-all cursor-pointer ${
                    paymentMethodFilter === tab.id
                      ? 'bg-emerald-600 text-white shadow'
                      : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700/70'
                  }`}
                >
                  {tab.label}
                </button>
              ))}

              <span className="text-slate-600 mx-1">|</span>

              <span className="text-xs text-slate-400">وضعیت:</span>
              {(
                [
                  { id: 'all', label: 'همه' },
                  { id: 'cleared', label: 'وصول شده' },
                  { id: 'pending', label: 'در جریان' },
                  { id: 'bounced', label: 'برگشتی' },
                ] as Array<{ id: 'all' | PaymentStatus; label: string }>
              ).map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setPaymentStatusFilter(tab.id)}
                  className={`px-2.5 py-1 rounded-xl text-xs font-medium transition-all cursor-pointer ${
                    paymentStatusFilter === tab.id
                      ? 'bg-slate-700 text-white border border-slate-500'
                      : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700/70'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            <button
              onClick={() => handleOpenPaymentModal()}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition-colors cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              <span>ثبت دریافت جدید</span>
            </button>
          </div>

          {/* Payments Table */}
          <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-right text-sm">
                <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
                  <tr>
                    <th className="py-3 px-3 text-center">ردیف</th>
                    <th className="py-3 px-4">تاریخ دریافت</th>
                    <th className="py-3 px-4">نام خریدار / مشتری</th>
                    <th className="py-3 px-3 text-center">شیوه دریافت</th>
                    <th className="py-3 px-4">شماره پیگیری / شناسه صیادی</th>
                    <th className="py-3 px-3">بانک عامل</th>
                    <th className="py-3 px-3 text-center">سررسید چک</th>
                    <th className="py-3 px-4 text-center">مبلغ دریافتی</th>
                    <th className="py-3 px-3 text-center">وضعیت وصول</th>
                    <th className="py-3 px-4">شرح / بابت</th>
                    <th className="py-3 px-4 text-center">عملیات</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {filteredPayments.length === 0 ? (
                    <tr>
                      <td colSpan={11} className="py-12 text-center text-slate-500">
                        سند دریافتی مطابق با فیلتر یافت نشد. روی «ثبت سند دریافت جدید» کلیک کنید.
                      </td>
                    </tr>
                  ) : (
                    filteredPayments.map((p, idx) => (
                      <tr key={p.id} className="hover:bg-slate-700/30 transition-colors">
                        <td className="py-3 px-3 text-center text-xs text-slate-500 font-mono">
                          {toPersianDigits(idx + 1)}
                        </td>

                        <td className="py-3 px-4 font-mono text-xs text-slate-300">
                          {toPersianDigits(p.paymentDate)}
                        </td>

                        <td className="py-3 px-4 font-bold text-white text-xs">
                          {p.customerName}
                          {p.invoiceNumber && (
                            <span className="block text-[11px] text-blue-400 font-normal mt-0.5">
                              بابت فاکتور #{toPersianDigits(p.invoiceNumber)}
                            </span>
                          )}
                        </td>

                        <td className="py-3 px-3 text-center">
                          {getMethodBadge(p.method)}
                        </td>

                        <td className="py-3 px-4 font-mono text-xs text-slate-300">
                          {p.referenceNumber ? toPersianDigits(p.referenceNumber) : '-'}
                        </td>

                        <td className="py-3 px-3 text-xs text-slate-300">
                          {p.bankName ? `بانک ${p.bankName}` : '-'}
                        </td>

                        <td className="py-3 px-3 text-center font-mono text-xs text-indigo-300">
                          {p.chequeDueDate ? toPersianDigits(p.chequeDueDate) : '-'}
                        </td>

                        <td className="py-3 px-4 text-center font-mono font-bold text-emerald-400 text-sm">
                          {formatNumber(p.amount)} {p.currencyUnit}
                        </td>

                        <td className="py-3 px-3 text-center">
                          <button
                            onClick={() => {
                              const nextStatus: PaymentStatus =
                                p.status === 'cleared' ? 'pending' : p.status === 'pending' ? 'bounced' : 'cleared';
                              handleToggleStatus(p, nextStatus);
                            }}
                            className="cursor-pointer"
                            title="برای تغییر وضعیت کلیک کنید"
                          >
                            {getStatusBadge(p.status)}
                          </button>
                        </td>

                        <td className="py-3 px-4 text-xs text-slate-400 max-w-xs truncate" title={p.notes}>
                          {p.notes || '-'}
                        </td>

                        <td className="py-3 px-4 text-center">
                          <div className="flex items-center justify-center gap-1">
                            {/* Print Voucher */}
                            <button
                              onClick={() => setActiveVoucherPayment(p)}
                              className="p-1.5 rounded-lg text-slate-400 hover:text-blue-300 hover:bg-blue-500/10 transition-colors cursor-pointer"
                              title="چاپ رسید صندوق / سند مالی"
                            >
                              <Printer className="w-3.5 h-3.5" />
                            </button>

                            {/* Edit */}
                            <button
                              onClick={() => handleOpenPaymentModal({ payment: p })}
                              className="p-1.5 rounded-lg text-slate-400 hover:text-amber-300 hover:bg-amber-500/10 transition-colors cursor-pointer"
                              title="ویرایش سند"
                            >
                              <Edit2 className="w-3.5 h-3.5" />
                            </button>

                            {/* Delete */}
                            {deleteConfirmId === p.id ? (
                              <div className="flex items-center gap-1 bg-rose-950/60 border border-rose-500/40 rounded-lg px-1.5 py-0.5">
                                <button
                                  onClick={() => handleDeletePayment(p.id)}
                                  className="text-[10px] font-bold text-rose-300 hover:text-white cursor-pointer"
                                >
                                  حذف
                                </button>
                                <span className="text-slate-600">|</span>
                                <button
                                  onClick={() => setDeleteConfirmId(null)}
                                  className="text-[10px] text-slate-400 hover:text-white cursor-pointer"
                                >
                                  لغو
                                </button>
                              </div>
                            ) : (
                              <button
                                onClick={() => setDeleteConfirmId(p.id)}
                                className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 transition-colors cursor-pointer"
                                title="حذف سند"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ============================================================== */}
      {/* TAB 3: INVOICES SETTLEMENT TRACKING & AGING                   */}
      {/* ============================================================== */}
      {activeTab === 'invoices' && (
        <div className="space-y-4">
          {/* Sub Filters */}
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-slate-400" />
            <span className="text-xs text-slate-400">وضعیت فاکتورها:</span>
            {(
              [
                { id: 'all', label: `همه (${toPersianDigits(invoices.length)})` },
                { id: 'unpaid', label: 'تسویه نشده / معوق' },
                { id: 'partially_paid', label: 'پرداخت ناقص / بیعانه' },
                { id: 'fully_paid', label: 'تسویه کامل' },
              ] as Array<{ id: 'all' | 'unpaid' | 'partially_paid' | 'fully_paid'; label: string }>
            ).map((tab) => (
              <button
                key={tab.id}
                onClick={() => setInvoiceSettlementFilter(tab.id)}
                className={`px-3 py-1 rounded-xl text-xs font-medium transition-all cursor-pointer ${
                  invoiceSettlementFilter === tab.id
                    ? 'bg-indigo-600 text-white shadow'
                    : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700/70'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* Table */}
          <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-right text-sm">
                <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
                  <tr>
                    <th className="py-3 px-4 text-center">شماره فاکتور</th>
                    <th className="py-3 px-4">تاریخ صدور</th>
                    <th className="py-3 px-4">نام خریدار</th>
                    <th className="py-3 px-4 text-center">مبلغ کل فاکتور</th>
                    <th className="py-3 px-4 text-center">مبلغ پرداخت‌شده</th>
                    <th className="py-3 px-4 text-center">مانده تسویه نشده</th>
                    <th className="py-3 px-4 text-center">پیشرفت تسویه</th>
                    <th className="py-3 px-4 text-center">وضعیت</th>
                    <th className="py-3 px-4 text-center">اقدام سریع</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {filteredInvoiceSettlements.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-12 text-center text-slate-500">
                        فاکتوری مطابق با فیلتر یافت نشد.
                      </td>
                    </tr>
                  ) : (
                    filteredInvoiceSettlements.map((item) => {
                      const inv = invoices.find((i) => i.id === item.invoiceId);
                      const percent =
                        item.totalAmount > 0
                          ? Math.min(100, Math.round((item.paidAmount / item.totalAmount) * 100))
                          : 100;

                      return (
                        <tr key={item.invoiceId} className="hover:bg-slate-700/30 transition-colors">
                          <td className="py-3.5 px-4 text-center font-mono font-bold text-blue-400">
                            {toPersianDigits(item.invoiceNumber)}
                          </td>

                          <td className="py-3.5 px-4 font-mono text-xs text-slate-300">
                            {toPersianDigits(item.date)}
                          </td>

                          <td className="py-3.5 px-4 font-bold text-white text-xs">
                            {item.customerName}
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono font-bold text-slate-200">
                            {formatNumber(item.totalAmount)} {item.currencyUnit}
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono font-bold text-emerald-400">
                            {formatNumber(item.paidAmount)} {item.currencyUnit}
                          </td>

                          <td className="py-3.5 px-4 text-center font-mono font-bold text-rose-400">
                            {item.remainingAmount > 0
                              ? `${formatNumber(item.remainingAmount)} ${item.currencyUnit}`
                              : '۰'}
                          </td>

                          <td className="py-3.5 px-4 text-center">
                            <div className="w-28 mx-auto space-y-1">
                              <div className="w-full bg-slate-900 rounded-full h-2 overflow-hidden border border-slate-700">
                                <div
                                  className={`h-full transition-all duration-300 ${
                                    percent >= 100
                                      ? 'bg-emerald-500'
                                      : percent > 0
                                      ? 'bg-amber-500'
                                      : 'bg-slate-600'
                                  }`}
                                  style={{ width: `${percent}%` }}
                                />
                              </div>
                              <span className="text-[10px] font-mono text-slate-400">
                                {toPersianDigits(percent)}٪ تسویه
                              </span>
                            </div>
                          </td>

                          <td className="py-3.5 px-4 text-center">
                            {item.status === 'fully_paid' ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                <span>تسویه کامل</span>
                              </span>
                            ) : item.status === 'partially_paid' ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                                <Clock className="w-3.5 h-3.5" />
                                <span>پرداخت ناقص</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-xl text-xs font-bold bg-rose-500/20 text-rose-300 border border-rose-500/30">
                                <AlertCircle className="w-3.5 h-3.5" />
                                <span>پرداخت‌نشده</span>
                              </span>
                            )}
                          </td>

                          <td className="py-3.5 px-4 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              {item.remainingAmount > 0 && inv && (
                                <button
                                  onClick={() => handleOpenPaymentModal({ invoice: inv })}
                                  className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold shadow-sm transition-colors cursor-pointer"
                                  title="ثبت واریزی و تسویه برای این فاکتور"
                                >
                                  <Plus className="w-3.5 h-3.5" />
                                  <span>ثبت تسویه</span>
                                </button>
                              )}

                              {inv && (
                                <button
                                  onClick={() => onPreviewPrintInvoice(inv)}
                                  className="p-1.5 rounded-lg text-slate-400 hover:text-blue-300 hover:bg-blue-500/10 transition-colors cursor-pointer"
                                  title="مشاهده فاکتور"
                                >
                                  <Printer className="w-4 h-4" />
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
        </div>
      )}

      {/* ============================================================== */}
      {/* TAB 4: OFFICIAL VAT REPORT (ARTICLE 169)                       */}
      {/* ============================================================== */}
      {activeTab === 'tax' && (
        <div className="space-y-4">
          <div className="bg-slate-800/80 p-5 rounded-2xl border border-slate-700 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <Building2 className="w-5 h-5 text-cyan-400" />
                <span>گزارش فصلی صورت معاملات و مالیات بر ارزش افزوده (ماده ۱۶۹ م.م)</span>
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                ویژه حسابداری و ارسال به سامانه مودیان و سامانه صورت معاملات فصلی با احتساب ۱۰٪ مالیات بر ارزش افزوده
              </p>
            </div>

            <button
              onClick={handleExportTaxCsv}
              className="flex items-center gap-2 px-4 py-2 bg-cyan-600 hover:bg-cyan-500 text-white rounded-xl text-xs font-bold transition-all shadow-lg shadow-cyan-600/30 cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>خروجی اکسل استاندارد ماده ۱۶۹</span>
            </button>
          </div>

          {/* Tax Table */}
          <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-right text-sm">
                <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
                  <tr>
                    <th className="py-3 px-4 text-center">شماره فاکتور</th>
                    <th className="py-3 px-4">تاریخ صدور</th>
                    <th className="py-3 px-4">نام خریدار رسمی</th>
                    <th className="py-3 px-4">شناسه ملی / کد اقتصادی</th>
                    <th className="py-3 px-4 text-center">مبلغ ناخالص</th>
                    <th className="py-3 px-4 text-center">تخفیف اعطایی</th>
                    <th className="py-3 px-4 text-center">خالص مشمول مالیات</th>
                    <th className="py-3 px-4 text-center">ارزش افزوده (۱۰٪)</th>
                    <th className="py-3 px-4 text-center">جمع کل نهایی</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {AccountingService.computeOfficialTaxRows(invoices).length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-12 text-center text-slate-500">
                        هنوز فاکتور رسمی در سیستم ثبت نشده است. هنگام صدور فاکتور گزینه «فاکتور رسمی ماده ۱۶۹» را فعال کنید.
                      </td>
                    </tr>
                  ) : (
                    AccountingService.computeOfficialTaxRows(invoices).map((row) => (
                      <tr key={row.invoiceNumber} className="hover:bg-slate-700/30 transition-colors">
                        <td className="py-3.5 px-4 text-center font-mono font-bold text-blue-400">
                          {toPersianDigits(row.invoiceNumber)}
                        </td>

                        <td className="py-3.5 px-4 font-mono text-xs text-slate-300">
                          {toPersianDigits(row.jalaliDate)}
                        </td>

                        <td className="py-3.5 px-4 font-bold text-white text-xs">
                          {row.customerName}
                        </td>

                        <td className="py-3.5 px-4 font-mono text-xs text-slate-300 space-y-0.5">
                          {row.customerNationalId && <div>ملی: {toPersianDigits(row.customerNationalId)}</div>}
                          {row.customerEconomicCode && <div>اقتصادی: {toPersianDigits(row.customerEconomicCode)}</div>}
                        </td>

                        <td className="py-3.5 px-4 text-center font-mono text-slate-300">
                          {formatNumber(row.grossAmount)}
                        </td>

                        <td className="py-3.5 px-4 text-center font-mono text-amber-300">
                          {row.discountAmount > 0 ? formatNumber(row.discountAmount) : '-'}
                        </td>

                        <td className="py-3.5 px-4 text-center font-mono font-bold text-slate-200">
                          {formatNumber(row.taxableAmount)}
                        </td>

                        <td className="py-3.5 px-4 text-center font-mono font-bold text-cyan-400">
                          {formatNumber(row.taxAmount)}
                        </td>

                        <td className="py-3.5 px-4 text-center font-mono font-bold text-emerald-400">
                          {formatNumber(row.totalAmount)} {row.currencyUnit}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ============================================================== */}
      {/* MODAL 1: REGISTER / EDIT PAYMENT VOUCHER                       */}
      {/* ============================================================== */}
      {isPaymentModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm p-4">
          <form
            onSubmit={handleSavePayment}
            className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-xl p-6 space-y-4 shadow-2xl"
          >
            <div className="flex items-center justify-between border-b border-slate-700 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="p-2 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                  <Wallet className="w-5 h-5" />
                </div>
                <h3 className="text-lg font-bold text-white">
                  {editingPayment ? 'ویرایش سند دریافت وجه / چک' : 'ثبت سند دریافت وجه و اسناد مالی'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsPaymentModalOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-700 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              {/* Customer Selector */}
              <div className="sm:col-span-2">
                <label className="block text-slate-300 mb-1 font-medium">خریدار / طرف حساب (الزامی):</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    required
                    list="accounting-customers-datalist"
                    value={paymentForm.customerName}
                    onChange={(e) => handleCustomerSelectInModal(e.target.value)}
                    placeholder="نام خریدار را تایپ یا از فهرست انتخاب کنید..."
                    className="flex-1 bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-emerald-500 focus:outline-none"
                  />
                  <datalist id="accounting-customers-datalist">
                    {customers.map((c) => (
                      <option key={c.id} value={c.name} />
                    ))}
                  </datalist>
                </div>
              </div>

              {/* Link to Invoice (Optional) */}
              <div className="sm:col-span-2">
                <label className="block text-slate-400 mb-1 font-medium">
                  بابت پیش‌فاکتور خاص (اختیاری):
                </label>
                <select
                  value={paymentForm.invoiceId}
                  onChange={(e) => handleInvoiceSelectInModal(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                >
                  <option value="">واریز به حساب کلی مشتری (بدون تخصیص فاکتور)</option>
                  {candidateInvoicesForSelectedCustomer.map((inv) => {
                    const st = invoiceSettlements.find((s) => s.invoiceId === inv.id);
                    const rem = st ? st.remainingAmount : inv.totalAmount;
                    return (
                      <option key={inv.id} value={inv.id}>
                        فاکتور #{toPersianDigits(inv.invoiceNumber)} - تاریخ {toPersianDigits(inv.jalaliDate)} (مانده: {formatNumber(rem)} {currencyUnit})
                      </option>
                    );
                  })}
                </select>
              </div>

              {/* Amount */}
              <div>
                <label className="block text-emerald-300 mb-1 font-medium">مبلغ دریافتی ({currencyUnit}):</label>
                <input
                  type="text"
                  required
                  value={paymentForm.amount}
                  onChange={(e) => {
                    const clean = e.target.value.replace(/[^0-9]/g, '');
                    setPaymentForm({ ...paymentForm, amount: clean ? formatNumber(clean) : '' });
                  }}
                  placeholder="مثال: ۲۵,۰۰۰,۰۰۰"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm font-bold text-emerald-400 font-mono text-left focus:border-emerald-500 focus:outline-none"
                />
              </div>

              {/* Payment Date */}
              <div>
                <label className="block text-slate-300 mb-1 font-medium">تاریخ دریافت (شمسی):</label>
                <input
                  type="text"
                  required
                  value={paymentForm.paymentDate}
                  onChange={(e) => setPaymentForm({ ...paymentForm, paymentDate: e.target.value })}
                  placeholder="۱۴۰۳/۰۷/۲۵"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-emerald-500 focus:outline-none"
                />
              </div>

              {/* Live Amount in Persian Words */}
              {parseFormattedNumber(paymentForm.amount) > 0 && (
                <div className="sm:col-span-2 bg-emerald-950/40 border border-emerald-500/30 rounded-xl p-2.5 text-xs text-emerald-300">
                  <span className="text-slate-400 ml-1">مبلغ به حروف:</span>
                  <span className="font-medium">
                    {numberToPersianWords(parseFormattedNumber(paymentForm.amount))} {currencyUnit}
                  </span>
                </div>
              )}

              {/* Payment Method */}
              <div>
                <label className="block text-slate-300 mb-1 font-medium">شیوه پرداخت:</label>
                <select
                  value={paymentForm.method}
                  onChange={(e) =>
                    setPaymentForm({
                      ...paymentForm,
                      method: e.target.value as PaymentMethod,
                      status: e.target.value === 'cheque' ? 'pending' : 'cleared',
                    })
                  }
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                >
                  <option value="transfer">حواله بانکی / پایا / ساتنا</option>
                  <option value="cheque">چک صیادی</option>
                  <option value="pos">دستگاه پوز / کارت‌خوان</option>
                  <option value="cash">پرداخت نقدی</option>
                  <option value="deposit">پیش‌پرداخت / بیعانه</option>
                </select>
              </div>

              {/* Status */}
              <div>
                <label className="block text-slate-300 mb-1 font-medium">وضعیت سند:</label>
                <select
                  value={paymentForm.status}
                  onChange={(e) => setPaymentForm({ ...paymentForm, status: e.target.value as PaymentStatus })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
                >
                  <option value="cleared">وصول شده (نهایی)</option>
                  <option value="pending">در جریان وصول (چک مدت‌دار)</option>
                  <option value="bounced">برگشت‌خورده</option>
                </select>
              </div>

              {/* Bank Name */}
              <div>
                <label className="block text-slate-300 mb-1 font-medium">بانک عامل:</label>
                <input
                  type="text"
                  value={paymentForm.bankName}
                  onChange={(e) => setPaymentForm({ ...paymentForm, bankName: e.target.value })}
                  placeholder="مثال: ملت، صادرات، ملی، تجارت..."
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-emerald-500 focus:outline-none"
                />
              </div>

              {/* Reference / Sayad ID */}
              <div>
                <label className="block text-slate-300 mb-1 font-medium">
                  {paymentForm.method === 'cheque' ? 'شناسه صیادی ۱۶ رقمی / شماره چک:' : 'شماره پیگیری / فیش واریز:'}
                </label>
                <input
                  type="text"
                  value={paymentForm.referenceNumber}
                  onChange={(e) => setPaymentForm({ ...paymentForm, referenceNumber: e.target.value })}
                  placeholder="مثال: ۱۰۲۵۴۸۷۹۶۳..."
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white font-mono text-left focus:border-emerald-500 focus:outline-none"
                />
              </div>

              {/* Cheque Due Date (if Cheque) */}
              {paymentForm.method === 'cheque' && (
                <div className="sm:col-span-2">
                  <label className="block text-indigo-300 mb-1 font-medium">تاریخ سررسید چک (شمسی):</label>
                  <input
                    type="text"
                    value={paymentForm.chequeDueDate}
                    onChange={(e) => setPaymentForm({ ...paymentForm, chequeDueDate: e.target.value })}
                    placeholder="مثال: ۱۴۰۳/۰۹/۱۵"
                    className="w-full bg-slate-900 border border-indigo-500/50 rounded-xl px-3 py-2 text-sm text-indigo-300 font-mono text-left focus:border-indigo-400 focus:outline-none"
                  />
                </div>
              )}

              {/* Notes */}
              <div className="sm:col-span-2">
                <label className="block text-slate-300 mb-1 font-medium">توضیحات و شرح سند:</label>
                <input
                  type="text"
                  value={paymentForm.notes}
                  onChange={(e) => setPaymentForm({ ...paymentForm, notes: e.target.value })}
                  placeholder="مثال: واریز نقدی بابت فاکتور لوله‌های سایز ۱۱۰..."
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-sm text-white focus:border-emerald-500 focus:outline-none"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-3 pt-3 border-t border-slate-700">
              <button
                type="button"
                onClick={() => setIsPaymentModalOpen(false)}
                className="px-4 py-2 text-sm text-slate-300 hover:text-white cursor-pointer"
              >
                انصراف
              </button>
              <button
                type="submit"
                className="flex items-center gap-2 px-6 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-sm font-bold shadow-lg shadow-emerald-600/30 cursor-pointer"
              >
                <Check className="w-4 h-4" />
                <span>{editingPayment ? 'ذخیره تغییرات سند' : 'ثبت و صدور سند دریافت'}</span>
              </button>
            </div>
          </form>
        </div>
      )}

      {/* ============================================================== */}
      {/* MODAL 2: CUSTOMER STATEMENT OF ACCOUNT (کارت معین خریدار)        */}
      {/* ============================================================== */}
      {activeStatementCustomer && (
        <CustomerStatementModal
          customer={activeStatementCustomer}
          invoices={invoices}
          payments={payments}
          settings={settings}
          currentDate={currentDate}
          onClose={() => setActiveStatementCustomer(null)}
          onShowToast={showToast}
        />
      )}

      {/* ============================================================== */}
      {/* MODAL 3: PAYMENT VOUCHER RECEIPT (رسید صندوق و امور مالی)        */}
      {/* ============================================================== */}
      {activeVoucherPayment && (
        <PaymentVoucherModal
          payment={activeVoucherPayment}
          settings={settings}
          onClose={() => setActiveVoucherPayment(null)}
          onShowToast={showToast}
        />
      )}
    </div>
  );
};

/* ==================================================================== */
/* SUB-COMPONENT: CUSTOMER STATEMENT MODAL (صورت‌حساب رسمی معین خریدار) */
/* ==================================================================== */
interface CustomerStatementModalProps {
  customer: Customer;
  invoices: Invoice[];
  payments: PaymentRecord[];
  settings: CompanySettings;
  currentDate: string;
  onClose: () => void;
  onShowToast: (msg: string) => void;
}

const CustomerStatementModal: React.FC<CustomerStatementModalProps> = ({
  customer,
  invoices,
  payments,
  settings,
  currentDate,
  onClose,
  onShowToast,
}) => {
  const report: CustomerStatementReport = useMemo(() => {
    return AccountingService.generateCustomerStatement(customer, invoices, payments, currentDate);
  }, [customer, invoices, payments, currentDate]);

  const currencyUnit = settings.currencyUnit || 'تومان';

  // Handle Print
  const handlePrint = () => {
    window.print();
  };

  // Handle CSV Download
  const handleDownloadCsv = () => {
    const csv = AccountingService.exportCustomerStatementToCsv(report, settings);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `statement_${customer.name}_${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    onShowToast('فایل اکسل صورت‌حساب مشتری دانلود شد.');
  };

  // Copy Summary to WhatsApp
  const handleCopySummary = async () => {
    const diagText =
      report.diagnosis === 'بدهکار'
        ? `بدهکار به مبلغ ${formatNumber(report.finalBalance)} ${currencyUnit}`
        : report.diagnosis === 'بستانکار'
        ? `بستانکار به مبلغ ${formatNumber(Math.abs(report.finalBalance))} ${currencyUnit}`
        : 'حساب کاملاً تسویه شده (بی‌حساب)';

    const text = [
      `📋 خلاصه صورت‌حساب مالی خریدار`,
      `خریدار: ${customer.name}`,
      `فروشنده: ${settings.name}`,
      `تاریخ: ${toPersianDigits(currentDate)}`,
      `-----------------------`,
      `مجموع فاکتورها: ${formatNumber(report.totalDebit)} ${currencyUnit}`,
      `مجموع واریزی‌ها: ${formatNumber(report.totalCredit)} ${currencyUnit}`,
      report.pendingCheques > 0 ? `چک‌های در جریان: ${formatNumber(report.pendingCheques)} ${currencyUnit}` : '',
      `وضعیت حساب: ${diagText}`,
      `-----------------------`,
      `جهت هماهنگی با واحد مالی: ${toPersianDigits(settings.phone)}`,
    ]
      .filter(Boolean)
      .join('\n');

    try {
      await navigator.clipboard.writeText(text);
      onShowToast('متن خلاصه صورت‌حساب با موفقیت در کلیپ‌بورد کپی شد.');
    } catch {
      onShowToast('خطا در کپی کردن متن.');
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4 overflow-y-auto">
      <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-4xl max-h-[92vh] flex flex-col shadow-2xl text-white">
        {/* Modal Top Bar (Hidden in Print) */}
        <div className="no-print flex items-center justify-between p-4 border-b border-slate-700 bg-slate-800 rounded-t-2xl">
          <div>
            <h3 className="text-base font-bold flex items-center gap-2">
              <FileText className="w-5 h-5 text-indigo-400" />
              <span>صورت‌حساب رسمی گردش مالی (کارت معین): {customer.name}</span>
            </h3>
            <span className="text-xs text-slate-400">
              شامل کلیه فاکتورهای فروش و دریافتی‌های قطعی خریدار به همراه مانده جاری
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleCopySummary}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium cursor-pointer"
              title="کپی متن خلاصه صورت‌حساب جهت ارسال در پیامک یا واتساپ"
            >
              <Copy className="w-3.5 h-3.5" />
              <span>کپی پیامک</span>
            </button>

            <button
              onClick={handleDownloadCsv}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-medium cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>خروجی اکسل</span>
            </button>

            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold shadow-lg shadow-blue-600/30 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>چاپ / PDF</span>
            </button>

            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-700 cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Printable Statement Canvas */}
        <div className="p-6 overflow-y-auto flex-1 space-y-5 bg-slate-900 print:bg-white print:text-black">
          {/* Header */}
          <div className="flex items-center justify-between border-b pb-4 border-slate-700 print:border-black">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-xl bg-blue-500/20 border border-blue-500/30 flex items-center justify-center p-1">
                <img src="./logo.png" alt="پلی غرب" className="w-full h-full object-contain" />
              </div>
              <div>
                <h2 className="text-xl font-bold text-white print:text-black">{settings.name}</h2>
                <p className="text-xs text-slate-400 print:text-gray-600">
                  صورت‌حساب گردش مالی خریدار (کارت حساب معین)
                </p>
              </div>
            </div>

            <div className="text-left text-xs font-mono text-slate-300 print:text-gray-800 space-y-1">
              <div>تاریخ گزارش: {toPersianDigits(report.generatedDate)}</div>
              <div>تلفن شرکت: {toPersianDigits(settings.phone)}</div>
              {settings.economicCode && <div>کد اقتصادی فروشنده: {toPersianDigits(settings.economicCode)}</div>}
            </div>
          </div>

          {/* Customer Metadata Card */}
          <div className="bg-slate-800/60 print:bg-gray-100 p-4 rounded-xl border border-slate-700 print:border-gray-300 grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div>
              <span className="text-slate-400 print:text-gray-600 block">نام خریدار:</span>
              <strong className="text-white print:text-black">{customer.name}</strong>
            </div>
            <div>
              <span className="text-slate-400 print:text-gray-600 block">شماره تماس:</span>
              <span className="font-mono">{customer.phone ? toPersianDigits(customer.phone) : '-'}</span>
            </div>
            <div>
              <span className="text-slate-400 print:text-gray-600 block">شناسه / کد ملی:</span>
              <span className="font-mono">{customer.nationalId ? toPersianDigits(customer.nationalId) : '-'}</span>
            </div>
            <div>
              <span className="text-slate-400 print:text-gray-600 block">کد اقتصادی:</span>
              <span className="font-mono">{customer.economicCode ? toPersianDigits(customer.economicCode) : '-'}</span>
            </div>
          </div>

          {/* KPI Mini-Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div className="bg-slate-800/80 print:bg-gray-50 p-3 rounded-xl border border-slate-700 print:border-gray-300">
              <span className="text-slate-400 print:text-gray-600 block">مجموع بدهکار (فاکتورها)</span>
              <strong className="text-sm font-mono text-white print:text-black">
                {formatNumber(report.totalDebit)} {currencyUnit}
              </strong>
            </div>

            <div className="bg-slate-800/80 print:bg-gray-50 p-3 rounded-xl border border-slate-700 print:border-gray-300">
              <span className="text-slate-400 print:text-gray-600 block">مجموع بستانکار (دریافتی‌ها)</span>
              <strong className="text-sm font-mono text-emerald-400 print:text-emerald-700">
                {formatNumber(report.totalCredit)} {currencyUnit}
              </strong>
            </div>

            <div className="bg-slate-800/80 print:bg-gray-50 p-3 rounded-xl border border-slate-700 print:border-gray-300">
              <span className="text-slate-400 print:text-gray-600 block">چک‌های در جریان وصول</span>
              <strong className="text-sm font-mono text-indigo-300 print:text-indigo-800">
                {formatNumber(report.pendingCheques)} {currencyUnit}
              </strong>
            </div>

            <div className="bg-slate-800/80 print:bg-gray-50 p-3 rounded-xl border border-slate-700 print:border-gray-300">
              <span className="text-slate-400 print:text-gray-600 block">مانده نهایی حساب</span>
              <strong
                className={`text-sm font-mono ${
                  report.diagnosis === 'بدهکار'
                    ? 'text-rose-400 print:text-rose-700'
                    : report.diagnosis === 'بستانکار'
                    ? 'text-blue-400 print:text-blue-700'
                    : 'text-emerald-400 print:text-emerald-700'
                }`}
              >
                {formatNumber(Math.abs(report.finalBalance))} {currencyUnit} ({report.diagnosis})
              </strong>
            </div>
          </div>

          {/* Ledger Statement Table */}
          <div className="border border-slate-700 print:border-gray-300 rounded-xl overflow-hidden">
            <table className="w-full text-right text-xs">
              <thead className="bg-slate-800 print:bg-gray-200 text-slate-300 print:text-black border-b border-slate-700 print:border-gray-300">
                <tr>
                  <th className="py-2.5 px-3 text-center">ردیف</th>
                  <th className="py-2.5 px-3">تاریخ</th>
                  <th className="py-2.5 px-3">شماره سند / فاکتور</th>
                  <th className="py-2.5 px-4">شرح رویداد مالی</th>
                  <th className="py-2.5 px-3 text-center">بدهکار ({currencyUnit})</th>
                  <th className="py-2.5 px-3 text-center">بستانکار ({currencyUnit})</th>
                  <th className="py-2.5 px-3 text-center">مانده ({currencyUnit})</th>
                  <th className="py-2.5 px-3 text-center">تشخیص</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800 print:divide-gray-200">
                {report.entries.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-8 text-center text-slate-500">
                      هیچ تراکنش یا فاکتوری برای این مشتری ثبت نشده است.
                    </td>
                  </tr>
                ) : (
                  report.entries.map((entry, idx) => (
                    <tr key={entry.id} className="hover:bg-slate-800/40 print:hover:bg-transparent">
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400 print:text-gray-600">
                        {toPersianDigits(idx + 1)}
                      </td>
                      <td className="py-2.5 px-3 font-mono text-slate-300 print:text-black">
                        {toPersianDigits(entry.date)}
                      </td>
                      <td className="py-2.5 px-3 font-mono font-bold text-blue-400 print:text-blue-800">
                        {toPersianDigits(entry.documentNumber)}
                      </td>
                      <td className="py-2.5 px-4 text-slate-300 print:text-black">
                        {entry.description}
                      </td>
                      <td className="py-2.5 px-3 text-center font-mono text-slate-200 print:text-black">
                        {entry.debit > 0 ? formatNumber(entry.debit) : '-'}
                      </td>
                      <td className="py-2.5 px-3 text-center font-mono font-bold text-emerald-400 print:text-emerald-700">
                        {entry.credit > 0 ? formatNumber(entry.credit) : '-'}
                      </td>
                      <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-100 print:text-black">
                        {formatNumber(entry.balance)}
                      </td>
                      <td className="py-2.5 px-3 text-center font-bold">
                        <span
                          className={`text-[10px] px-2 py-0.5 rounded-full ${
                            entry.diagnosis === 'بدهکار'
                              ? 'bg-rose-500/20 text-rose-300 print:text-rose-700'
                              : entry.diagnosis === 'بستانکار'
                              ? 'bg-blue-500/20 text-blue-300 print:text-blue-700'
                              : 'bg-emerald-500/20 text-emerald-300 print:text-emerald-700'
                          }`}
                        >
                          {entry.diagnosis}
                        </span>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
              <tfoot className="bg-slate-800/90 print:bg-gray-100 font-bold border-t border-slate-700 print:border-gray-300">
                <tr>
                  <td colSpan={4} className="py-3 px-4 text-left">
                    جمع کل گردش و مانده نهایی:
                  </td>
                  <td className="py-3 px-3 text-center font-mono text-white print:text-black">
                    {formatNumber(report.totalDebit)}
                  </td>
                  <td className="py-3 px-3 text-center font-mono text-emerald-400 print:text-emerald-700">
                    {formatNumber(report.totalCredit)}
                  </td>
                  <td className="py-3 px-3 text-center font-mono text-white print:text-black">
                    {formatNumber(Math.abs(report.finalBalance))}
                  </td>
                  <td className="py-3 px-3 text-center text-xs">
                    {report.diagnosis}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>

          {/* Amount in words & Signature section */}
          <div className="bg-slate-800/40 print:bg-gray-50 p-4 rounded-xl border border-slate-700 print:border-gray-300 space-y-3">
            <div className="text-xs">
              <span className="text-slate-400 print:text-gray-600">مانده نهایی به حروف:</span>{' '}
              <strong className="text-white print:text-black">
                {numberToPersianWords(Math.abs(report.finalBalance))} {currencyUnit} ({report.diagnosis})
              </strong>
            </div>

            <div className="grid grid-cols-2 pt-6 text-center text-xs text-slate-400 print:text-gray-700">
              <div>امضاء و تایید امور مالی {settings.name}</div>
              <div>امضاء و تایید خریدار / نماینده قانونی</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

/* ==================================================================== */
/* SUB-COMPONENT: PAYMENT VOUCHER MODAL (رسید صندوق و سند مالی دریافت)   */
/* ==================================================================== */
interface PaymentVoucherModalProps {
  payment: PaymentRecord;
  settings: CompanySettings;
  onClose: () => void;
  onShowToast: (msg: string) => void;
}

const PaymentVoucherModal: React.FC<PaymentVoucherModalProps> = ({
  payment,
  settings,
  onClose,
}) => {
  const currencyUnit = payment.currencyUnit || settings.currencyUnit || 'تومان';

  const methodNames: Record<string, string> = {
    transfer: 'حواله بانکی / ساتنا / پایا',
    cheque: 'چک صیادی',
    pos: 'کارت‌خوان / دستگاه پوز',
    cash: 'پرداخت نقدی',
    deposit: 'بیعانه / پیش‌پرداخت',
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-2xl p-6 shadow-2xl text-white space-y-5">
        {/* Top Controls */}
        <div className="no-print flex items-center justify-between border-b border-slate-700 pb-3">
          <h3 className="text-base font-bold flex items-center gap-2">
            <Printer className="w-5 h-5 text-emerald-400" />
            <span>رسید رسمی دریافت وجه و اسناد تجاری</span>
          </h3>
          <div className="flex items-center gap-2">
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>چاپ رسید</span>
            </button>
            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-700 cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Printable Voucher Paper */}
        <div className="bg-slate-800/80 print:bg-white print:text-black border border-slate-700 print:border-black rounded-xl p-6 space-y-4 text-xs">
          {/* Voucher Header */}
          <div className="flex items-center justify-between border-b pb-3 border-slate-700 print:border-black">
            <div>
              <h2 className="text-base font-bold text-white print:text-black">{settings.name}</h2>
              <span className="text-[11px] text-slate-400 print:text-gray-600">رسید دریافت صندوق و اسناد مالی</span>
            </div>
            <div className="text-left font-mono text-[11px] text-slate-300 print:text-black space-y-0.5">
              <div>شماره سند: {payment.id.slice(-8)}</div>
              <div>تاریخ: {toPersianDigits(payment.paymentDate)}</div>
            </div>
          </div>

          {/* Voucher Body Grid */}
          <div className="grid grid-cols-2 gap-3 leading-relaxed">
            <div>
              <span className="text-slate-400 print:text-gray-600 block">دریافت شد از:</span>
              <strong className="text-sm text-white print:text-black">{payment.customerName}</strong>
            </div>

            <div>
              <span className="text-slate-400 print:text-gray-600 block">مبلغ دریافتی:</span>
              <strong className="text-sm font-mono text-emerald-400 print:text-emerald-700">
                {formatNumber(payment.amount)} {currencyUnit}
              </strong>
            </div>

            <div className="col-span-2 bg-slate-900/50 print:bg-gray-100 p-2.5 rounded-lg border border-slate-700/60 print:border-gray-300">
              <span className="text-slate-400 print:text-gray-600 ml-1">مبلغ به حروف:</span>
              <strong className="text-white print:text-black">
                {numberToPersianWords(payment.amount)} {currencyUnit}
              </strong>
            </div>

            <div>
              <span className="text-slate-400 print:text-gray-600 block">روش پرداخت:</span>
              <span>{methodNames[payment.method] || payment.method}</span>
            </div>

            <div>
              <span className="text-slate-400 print:text-gray-600 block">بانک عامل:</span>
              <span>{payment.bankName ? `بانک ${payment.bankName}` : '-'}</span>
            </div>

            <div>
              <span className="text-slate-400 print:text-gray-600 block">شماره پیگیری / شناسه صیادی:</span>
              <span className="font-mono">{payment.referenceNumber ? toPersianDigits(payment.referenceNumber) : '-'}</span>
            </div>

            {payment.chequeDueDate && (
              <div>
                <span className="text-slate-400 print:text-gray-600 block">سررسید چک:</span>
                <span className="font-mono font-bold text-indigo-300 print:text-indigo-800">
                  {toPersianDigits(payment.chequeDueDate)}
                </span>
              </div>
            )}

            {payment.invoiceNumber && (
              <div>
                <span className="text-slate-400 print:text-gray-600 block">بابت فاکتور:</span>
                <span className="font-mono text-blue-400 print:text-blue-800">
                  پیش‌فاکتور شماره {toPersianDigits(payment.invoiceNumber)}
                </span>
              </div>
            )}

            {payment.notes && (
              <div className="col-span-2">
                <span className="text-slate-400 print:text-gray-600 block">توضیحات:</span>
                <span>{payment.notes}</span>
              </div>
            )}
          </div>

          {/* Signatures */}
          <div className="grid grid-cols-2 pt-8 text-center text-slate-400 print:text-gray-700 border-t border-slate-700/60 print:border-black">
            <div>امضاء تحویل‌دهنده</div>
            <div>مهر و امضاء امور مالی / صندوق</div>
          </div>
        </div>
      </div>
    </div>
  );
};
