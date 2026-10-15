import React, { useState, useMemo } from 'react';
import { Search, Trash2, Printer, FileText, Edit3, Copy } from 'lucide-react';
import { CompanySettings, Invoice } from '../../types';
import { formatNumber, toPersianDigits } from '../../services/persianUtils';
import { StorageService } from '../../services/storage';

interface InvoiceHistoryProps {
  invoices: Invoice[];
  onUpdateInvoices: (invoices: Invoice[]) => void;
  settings: CompanySettings;
  onPreviewPrint: (invoice: Invoice) => void;
  onEditInvoice?: (invoice: Invoice) => void;
  onDuplicateInvoice?: (invoice: Invoice) => void;
}

export const InvoiceHistory: React.FC<InvoiceHistoryProps> = ({
  invoices,
  onUpdateInvoices,
  settings,
  onPreviewPrint,
  onEditInvoice,
  onDuplicateInvoice,
}) => {
  const [search, setSearch] = useState('');

  const filteredInvoices = useMemo(() => {
    if (!search.trim()) return invoices;
    const q = search.toLowerCase().trim();
    return invoices.filter(
      (inv) =>
        inv.invoiceNumber.toLowerCase().includes(q) ||
        inv.customerName.toLowerCase().includes(q) ||
        (inv.customerPhone && inv.customerPhone.includes(q)) ||
        inv.jalaliDate.includes(q)
    );
  }, [invoices, search]);

  const handleDelete = (id: string, num: string) => {
    if (confirm(`آیا از حذف پیش‌فاکتور شماره ${num} اطمینان دارید؟`)) {
      StorageService.deleteInvoice(id);
      onUpdateInvoices(StorageService.getInvoices());
    }
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-800/80 p-6 rounded-2xl border border-slate-700 shadow-lg">
        <div>
          <h1 className="text-2xl font-bold text-white flex items-center gap-3">
            <FileText className="w-7 h-7 text-blue-400" />
            <span>سوابق و آرشیو پیش‌فاکتورها</span>
          </h1>
          <p className="text-slate-400 text-sm mt-1">
            مشاهده، جستجو، چاپ مجدد و مدیریت پیش‌فاکتورهای صادر شده
          </p>
        </div>

        {/* Search Input */}
        <div className="relative w-full sm:w-80">
          <Search className="w-5 h-5 absolute right-3 top-3 text-slate-400" />
          <input
            type="text"
            placeholder="جستجو با شماره فاکتور، نام مشتری یا تلفن..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="w-full bg-slate-900 border border-slate-700 rounded-xl pr-10 pl-4 py-2.5 text-sm text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none"
          />
        </div>
      </div>

      {/* Invoices List Table */}
      <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-right text-sm">
            <thead className="bg-slate-900/90 text-xs font-semibold text-slate-400 border-b border-slate-700">
              <tr>
                <th className="py-3.5 px-4 text-center">شماره فاکتور</th>
                <th className="py-3.5 px-4">تاریخ صدور</th>
                <th className="py-3.5 px-4">نام خریدار / مشتری</th>
                <th className="py-3.5 px-4 text-center">تعداد اقلام</th>
                <th className="py-3.5 px-4 text-center">مبلغ کل فاکتور</th>
                <th className="py-3.5 px-4 text-center">عملیات</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-700/60">
              {filteredInvoices.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-slate-500">
                    {search ? 'فاکتوری مطابق با عبارت جستجو شده پیدا نشد.' : 'هنوز فاکتوری در سیستم ثبت نشده است.'}
                  </td>
                </tr>
              ) : (
                filteredInvoices.map((inv) => (
                  <tr key={inv.id} className="hover:bg-slate-700/30 transition-colors">
                    <td className="py-3.5 px-4 text-center font-mono font-bold text-blue-400">
                      {toPersianDigits(inv.invoiceNumber)}
                    </td>
                    <td className="py-3.5 px-4 text-slate-300 font-mono">
                      {toPersianDigits(inv.jalaliDate)}
                    </td>
                    <td className="py-3.5 px-4 font-medium text-white">
                      <div>{inv.customerName}</div>
                      {inv.customerPhone && (
                        <div className="text-xs text-slate-400 font-mono">{inv.customerPhone}</div>
                      )}
                    </td>
                    <td className="py-3.5 px-4 text-center font-mono text-slate-300">
                      {toPersianDigits(inv.items.length)} ردیف
                    </td>
                    <td className="py-3.5 px-4 text-center font-mono font-bold text-emerald-400">
                      {formatNumber(inv.totalAmount)}{' '}
                      <span className="text-xs font-normal text-slate-400 font-sans">
                        {inv.currencyUnit || settings.currencyUnit}
                      </span>
                    </td>
                    <td className="py-3.5 px-4 text-center">
                      <div className="flex items-center justify-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => onEditInvoice?.(inv)}
                          className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-amber-600/20 hover:bg-amber-600/40 text-amber-300 border border-amber-500/30 text-xs font-medium transition-colors cursor-pointer"
                          title="ویرایش پیش‌فاکتور در فرم سازنده"
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                          <span>ویرایش</span>
                        </button>

                        <button
                          onClick={() => onDuplicateInvoice?.(inv)}
                          className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/30 text-xs font-medium transition-colors cursor-pointer"
                          title="تکثیر این فاکتور با شماره جدید"
                        >
                          <Copy className="w-3.5 h-3.5" />
                          <span>تکثیر</span>
                        </button>

                        <button
                          onClick={() => onPreviewPrint(inv)}
                          className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-blue-600/20 hover:bg-blue-600/40 text-blue-300 border border-blue-500/30 text-xs font-medium transition-colors cursor-pointer"
                          title="مشاهده و چاپ"
                        >
                          <Printer className="w-3.5 h-3.5" />
                          <span>چاپ / PDF</span>
                        </button>

                        <button
                          onClick={() => handleDelete(inv.id, inv.invoiceNumber)}
                          className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 transition-colors cursor-pointer"
                          title="حذف از سوابق"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
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
  );
};

