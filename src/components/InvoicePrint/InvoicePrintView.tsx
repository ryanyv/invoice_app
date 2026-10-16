import React, { useState } from 'react';
import {
  Printer,
  ArrowRight,
  FileSpreadsheet,
  Share2,
  Check,
  ShieldCheck,
  FileText,
} from 'lucide-react';
import { CompanySettings, Invoice, PipeItem, FittingItem } from '../../types';
import { formatNumber, numberToPersianWords, toPersianDigits } from '../../services/persianUtils';
import { exportInvoiceToCsv, copyInvoiceWhatsAppSummary } from '../../utils/exportUtils';

interface InvoicePrintViewProps {
  invoice: Invoice;
  settings: CompanySettings;
  onBack: () => void;
}

export const InvoicePrintView: React.FC<InvoicePrintViewProps> = ({
  invoice,
  settings,
  onBack,
}) => {
  // 1. Orientation state: Landscape (default) vs Portrait
  const [isLandscape, setIsLandscape] = useState(true);

  // 2. Official Mode state: Simple Proforma vs Official Tax Invoice (ماده ۱۶۹ م.م)
  const [isOfficial, setIsOfficial] = useState(Boolean(invoice.isOfficial));

  // 3. WhatsApp copy feedback state
  const [copiedToast, setCopiedToast] = useState(false);

  const currencyUnit = invoice.currencyUnit || settings.currencyUnit || 'تومان';

  const handlePrint = React.useCallback(async () => {
    // Check if running inside Electron with savePdf IPC
    if ((window as any).electronAPI?.savePdf) {
      const modePrefix = isOfficial ? 'صورتحساب_رسمی' : 'پیش_فاکتور';
      const defaultName = `${modePrefix}_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;
      const res = await (window as any).electronAPI.savePdf({
        defaultName,
        landscape: isLandscape,
      });
      if (res?.success) {
        alert(`فایل PDF با موفقیت ذخیره شد:\n${res.filePath}`);
      } else if (!res?.canceled) {
        window.print();
      }
    } else {
      window.print();
    }
  }, [invoice, isOfficial, isLandscape]);

  React.useEffect(() => {
    const onTriggerPrint = () => {
      handlePrint();
    };
    const onEscape = () => {
      onBack();
    };
    window.addEventListener('pipecal:trigger-print', onTriggerPrint);
    window.addEventListener('pipecal:escape', onEscape);
    return () => {
      window.removeEventListener('pipecal:trigger-print', onTriggerPrint);
      window.removeEventListener('pipecal:escape', onEscape);
    };
  }, [handlePrint, onBack]);

  const handleExportCsv = () => {
    exportInvoiceToCsv({ ...invoice, isOfficial, currencyUnit }, settings);
  };

  const handleCopyWhatsApp = async () => {
    const success = await copyInvoiceWhatsAppSummary(
      { ...invoice, isOfficial, currencyUnit },
      settings
    );
    if (success) {
      setCopiedToast(true);
      setTimeout(() => setCopiedToast(false), 2500);
    }
  };

  const pipeItems = invoice.items.filter((i): i is PipeItem => i.type === 'pipe');
  const fittingItems = invoice.items.filter((i): i is FittingItem => i.type === 'fitting');

  const totalPipeWeight = pipeItems.reduce((acc, p) => acc + (p.totalWeight || 0), 0);
  const totalAmountWords = numberToPersianWords(invoice.totalAmount);

  return (
    <div className="min-h-screen bg-slate-900 py-8 px-4 print:p-0 print:bg-white text-slate-900">
      {/* Dynamic Print Styles for A4 Page Orientation */}
      <style>{`
        @media print {
          @page {
            size: A4 ${isLandscape ? 'landscape' : 'portrait'};
            margin: 8mm;
          }
        }
      `}</style>

      {/* Action Bar (Hidden when printing) */}
      <div className="no-print max-w-6xl mx-auto mb-6 flex flex-wrap items-center justify-between gap-4 bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-xl">
        <button
          onClick={onBack}
          className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-xl text-sm font-medium transition-colors cursor-pointer"
        >
          <ArrowRight className="w-4 h-4" />
          <span>بازگشت به ویرایش پیش‌فاکتور</span>
        </button>

        <div className="flex flex-wrap items-center gap-3">
          {/* Orientation Toggle */}
          <div className="bg-slate-900 p-1 rounded-xl border border-slate-700 flex items-center gap-1 text-xs font-medium">
            <button
              onClick={() => setIsLandscape(true)}
              className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                isLandscape
                  ? 'bg-blue-600 text-white shadow-sm font-bold'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              افقی (A4 Landscape)
            </button>
            <button
              onClick={() => setIsLandscape(false)}
              className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                !isLandscape
                  ? 'bg-blue-600 text-white shadow-sm font-bold'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              عمودی (A4 Portrait)
            </button>
          </div>

          {/* Official vs Simple Proforma Toggle */}
          <div className="bg-slate-900 p-1 rounded-xl border border-slate-700 flex items-center gap-1 text-xs font-medium">
            <button
              onClick={() => setIsOfficial(false)}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                !isOfficial
                  ? 'bg-emerald-600 text-white shadow-sm font-bold'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <FileText className="w-3.5 h-3.5" />
              <span>پیش‌فاکتور معمولی</span>
            </button>
            <button
              onClick={() => setIsOfficial(true)}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                isOfficial
                  ? 'bg-amber-600 text-white shadow-sm font-bold'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <ShieldCheck className="w-3.5 h-3.5" />
              <span>پیش‌فاکتور رسمی (ماده ۱۶۹)</span>
            </button>
          </div>

          {/* Export to CSV/Excel */}
          <button
            onClick={handleExportCsv}
            className="flex items-center gap-2 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-medium shadow-md transition-all cursor-pointer"
            title="خروجی مستقیم به فایل اکسل / CSV با پشتیبانی زبان فارسی"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>خروجی اکسل / CSV</span>
          </button>

          {/* WhatsApp / SMS Summary Copy */}
          <button
            onClick={handleCopyWhatsApp}
            className={`flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-medium transition-all cursor-pointer ${
              copiedToast
                ? 'bg-emerald-500 text-white shadow-lg'
                : 'bg-teal-600 hover:bg-teal-500 text-white shadow-md'
            }`}
            title="کپی متن خلاصه برای پیامک، واتساپ یا شبکه‌های اجتماعی"
          >
            {copiedToast ? <Check className="w-4 h-4" /> : <Share2 className="w-4 h-4" />}
            <span>{copiedToast ? '✓ خلاصه فاکتور کپی شد' : 'کپی خلاصه پیامک / واتساپ'}</span>
          </button>

          {/* Print / Save PDF */}
          <button
            onClick={handlePrint}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
          >
            <Printer className="w-4 h-4" />
            <span>چاپ / PDF</span>
          </button>
        </div>
      </div>

      {/* Printable Sheet */}
      <div
        className={`print-container mx-auto bg-white rounded-xl shadow-2xl print:shadow-none print:rounded-none border border-slate-200 text-black ${
          isLandscape
            ? 'max-w-5xl p-8 sm:p-12 print:p-4 text-sm'
            : 'max-w-3xl p-6 sm:p-8 print:p-2 text-xs'
        }`}
      >
        {/* ===================== MODE 1: OFFICIAL TAX INVOICE (ماده ۱۶۹ م.م) ===================== */}
        {isOfficial ? (
          <div>
            {/* Official Header */}
            <div className="border-2 border-slate-900 rounded-lg p-3 mb-4 bg-slate-50">
              <div className="flex items-center justify-between">
                <div className="w-16 h-16 border border-slate-300 rounded-md p-1 flex items-center justify-center bg-white">
                  <img src="./logo.png" alt="لوگو شرکت" className="max-w-full max-h-full object-contain" />
                </div>
                <div className="text-center flex-1">
                  <h1 className="text-base sm:text-lg font-black text-slate-950">
                    صورتحساب رسمی فروش کالا و خدمات
                  </h1>
                  <span className="text-[11px] font-bold text-slate-600 block mt-0.5">
                    (ماده ۱۶۹ مکرر قانون مالیات‌های مستقیم)
                  </span>
                </div>
                <div className="text-right text-[11px] space-y-1 bg-white border border-slate-300 rounded p-2 min-w-[170px]">
                  <div className="flex justify-between">
                    <span className="text-slate-500">شماره سریال:</span>
                    <strong className="font-mono text-slate-900 font-bold">{toPersianDigits(invoice.invoiceNumber)}</strong>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-slate-500">تاریخ صدور:</span>
                    <span className="font-mono text-slate-800 font-medium">{toPersianDigits(invoice.jalaliDate)}</span>
                  </div>
                </div>
              </div>
            </div>

            {/* Official 2-Box: Seller Table & Buyer Table */}
            <div className="space-y-3 mb-4">
              {/* Seller Table */}
              <div className="border border-slate-400 rounded-lg overflow-hidden text-[11px]">
                <div className="bg-slate-200 py-1 px-3 font-bold text-slate-900 border-b border-slate-400">
                  مشخصات فروشنده
                </div>
                <div className="p-2.5 bg-slate-50/60 grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <div>
                    <span className="text-slate-500 ml-1">نام شخص حقیقی/حقوقی:</span>
                    <strong className="text-slate-900">{settings.name}</strong>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">شماره اقتصادی:</span>
                    <span className="font-mono font-bold text-slate-900">{toPersianDigits(settings.economicCode || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">شناسه ملی / کد ملی:</span>
                    <span className="font-mono font-bold text-slate-900">{toPersianDigits(settings.nationalId || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">شماره ثبت:</span>
                    <span className="font-mono text-slate-900">{toPersianDigits(settings.registrationNumber || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">کد پستی ۱۰ رقمی:</span>
                    <span className="font-mono text-slate-900">{toPersianDigits(settings.postalCode || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">تلفن و نمابر:</span>
                    <span className="font-mono text-slate-900">{toPersianDigits(settings.phone)} {settings.fax ? `/ ${toPersianDigits(settings.fax)}` : ''}</span>
                  </div>
                  <div className="sm:col-span-3">
                    <span className="text-slate-500 ml-1">نشانی کامل:</span>
                    <span className="text-slate-800">{settings.address}</span>
                  </div>
                </div>
              </div>

              {/* Buyer Table */}
              <div className="border border-slate-400 rounded-lg overflow-hidden text-[11px]">
                <div className="bg-slate-200 py-1 px-3 font-bold text-slate-900 border-b border-slate-400">
                  مشخصات خریدار
                </div>
                <div className="p-2.5 bg-slate-50/60 grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <div>
                    <span className="text-slate-500 ml-1">نام شخص حقیقی/حقوقی:</span>
                    <strong className="text-slate-900">{invoice.customerName}</strong>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">شماره اقتصادی:</span>
                    <span className="font-mono font-bold text-slate-900">{toPersianDigits(invoice.customerEconomicCode || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">شناسه ملی / کد ملی:</span>
                    <span className="font-mono font-bold text-slate-900">{toPersianDigits(invoice.customerNationalId || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">کد پستی ۱۰ رقمی:</span>
                    <span className="font-mono text-slate-900">{toPersianDigits(invoice.customerPostalCode || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">تلفن تماس:</span>
                    <span className="font-mono text-slate-900">{toPersianDigits(invoice.customerPhone || '-')}</span>
                  </div>
                  <div>
                    <span className="text-slate-500 ml-1">واحد پول صورتحساب:</span>
                    <span className="font-bold text-blue-900">{currencyUnit}</span>
                  </div>
                  <div className="sm:col-span-3">
                    <span className="text-slate-500 ml-1">نشانی کامل خریدار:</span>
                    <span className="text-slate-800">{invoice.customerAddress || '-'}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* ===================== MODE 2: STANDARD PROFORMA ===================== */
          <div>
            {/* Header Section */}
            <div className="flex items-start justify-between border-b-2 border-slate-900 pb-4 mb-5">
              {/* Logo & Company info */}
              <div className="flex items-center gap-4">
                <div className="w-16 h-16 sm:w-20 sm:h-20 border border-slate-300 rounded-lg p-1 flex items-center justify-center">
                  <img src="./logo.png" alt="پلی غرب" className="max-w-full max-h-full object-contain" />
                </div>
                <div>
                  <h1 className="text-lg sm:text-xl font-extrabold text-slate-900">{settings.name}</h1>
                  <p className="text-xs text-slate-600 mt-0.5 max-w-md">{settings.address}</p>
                  <div className="flex items-center gap-4 text-xs text-slate-600 mt-1">
                    <span>تلفن: {toPersianDigits(settings.phone)}</span>
                    {settings.fax && <span>فکس: {toPersianDigits(settings.fax)}</span>}
                  </div>
                </div>
              </div>

              {/* Invoice Meta */}
              <div className="border border-slate-300 rounded-lg p-3 bg-slate-50 min-w-[200px] text-xs space-y-1.5 text-right">
                <div className="text-center font-bold text-sm text-blue-900 pb-1 border-b border-slate-200">
                  پیش‌فاکتور فروش
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">شماره پیش‌فاکتور:</span>
                  <strong className="text-slate-900 font-mono text-sm">{toPersianDigits(invoice.invoiceNumber)}</strong>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">تاریخ:</span>
                  <span className="text-slate-800 font-medium">{toPersianDigits(invoice.jalaliDate)}</span>
                </div>
              </div>
            </div>

            {/* Customer Information Banner */}
            <div className="bg-slate-100/80 rounded-lg p-3 mb-5 border border-slate-200 flex flex-wrap items-center justify-between text-xs gap-3">
              <div>
                <span className="text-slate-500 ml-2">نام مشتری / خریدار:</span>
                <strong className="text-slate-900 text-sm">{invoice.customerName}</strong>
              </div>
              {invoice.customerPhone && (
                <div>
                  <span className="text-slate-500 ml-2">تلفن تماس:</span>
                  <span className="text-slate-800 font-mono">{toPersianDigits(invoice.customerPhone)}</span>
                </div>
              )}
              {invoice.customerAddress && (
                <div className="w-full text-slate-700">
                  <span className="text-slate-500 ml-2">آدرس خریدار:</span>
                  <span>{invoice.customerAddress}</span>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Section 1: Pipe Items (if any) */}
        {pipeItems.length > 0 && (
          <div className="mb-5">
            <div className="font-bold text-xs sm:text-sm text-slate-800 mb-2 border-r-4 border-blue-600 pr-2">
              اقلام لوله پلی‌اتیلن
            </div>
            <table className={`w-full text-center border-collapse border border-slate-300 ${
              isLandscape ? 'text-xs' : 'text-[11px]'
            }`}>
              <thead className="bg-slate-200 text-slate-800 font-bold">
                <tr>
                  <th className="border border-slate-300 py-1.5 px-1 w-8">ردیف</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">قطر (mm)</th>
                  <th className="border border-slate-300 py-1.5 px-1">SDR</th>
                  <th className="border border-slate-300 py-1.5 px-1">فشار (PN)</th>
                  <th className="border border-slate-300 py-1.5 px-1">گرید</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">طول (متر)</th>
                  <th className="border border-slate-300 py-1.5 px-1">وزن/متر (kg)</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">وزن کل (kg)</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">قیمت/kg ({currencyUnit})</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">قیمت کل ({currencyUnit})</th>
                </tr>
              </thead>
              <tbody>
                {pipeItems.map((item, idx) => (
                  <tr key={item.id} className={idx % 2 === 1 ? 'bg-slate-50' : 'bg-white'}>
                    <td className="border border-slate-300 py-1 px-1 font-mono">{toPersianDigits(idx + 1)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-bold">{toPersianDigits(item.diameter)}</td>
                    <td className="border border-slate-300 py-1 px-1 font-mono">{toPersianDigits(item.sdr)}</td>
                    <td className="border border-slate-300 py-1 px-1 font-mono">{toPersianDigits(item.pn)}</td>
                    <td className="border border-slate-300 py-1 px-1">{item.peGrade}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono">{toPersianDigits(item.length)}</td>
                    <td className="border border-slate-300 py-1 px-1 font-mono">{toPersianDigits(item.weightPerMeter.toFixed(3))}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono font-medium">{toPersianDigits(item.totalWeight.toFixed(3))}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono">{formatNumber(item.pricePerKg, true)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono font-bold text-slate-900">{formatNumber(item.totalPrice, true)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Section 2: Fitting Items (if any) */}
        {fittingItems.length > 0 && (
          <div className="mb-5">
            <div className="font-bold text-xs sm:text-sm text-slate-800 mb-2 border-r-4 border-emerald-600 pr-2">
              اقلام اتصالات پلی‌اتیلن
            </div>
            <table className={`w-full text-center border-collapse border border-slate-300 ${
              isLandscape ? 'text-xs' : 'text-[11px]'
            }`}>
              <thead className="bg-slate-200 text-slate-800 font-bold">
                <tr>
                  <th className="border border-slate-300 py-1.5 px-1 w-8">ردیف</th>
                  <th className="border border-slate-300 py-1.5 px-2 text-right">نوع اتصال</th>
                  <th className="border border-slate-300 py-1.5 px-2 text-right">نام محصول</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">فشار (bar)</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">سایز (mm)</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">تعداد</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">قیمت واحد ({currencyUnit})</th>
                  <th className="border border-slate-300 py-1.5 px-1 sm:px-2">قیمت کل ({currencyUnit})</th>
                </tr>
              </thead>
              <tbody>
                {fittingItems.map((item, idx) => (
                  <tr key={item.id} className={idx % 2 === 1 ? 'bg-slate-50' : 'bg-white'}>
                    <td className="border border-slate-300 py-1 px-1 font-mono">{toPersianDigits(idx + 1)}</td>
                    <td className="border border-slate-300 py-1 px-2 text-right">{item.category} ({item.connectionType})</td>
                    <td className="border border-slate-300 py-1 px-2 text-right font-medium">{item.product}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono">{toPersianDigits(item.pressure)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-bold">{toPersianDigits(item.size)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono">{toPersianDigits(item.quantity)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono">{formatNumber(item.unitPrice, true)}</td>
                    <td className="border border-slate-300 py-1 px-1 sm:px-2 font-mono font-bold text-slate-900">{formatNumber(item.totalPrice, true)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Financial Summary & Total Box */}
        <div className="flex flex-col md:flex-row justify-between items-start gap-4 border-t-2 border-slate-300 pt-3 mb-5">
          {/* Notes & Amount in Words */}
          <div className="flex-1 space-y-2.5">
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-2.5 text-xs">
              <span className="font-bold text-blue-950 block mb-0.5">مبلغ کل قابل پرداخت به حروف:</span>
              <span className="text-blue-900 font-bold text-xs sm:text-sm">
                {totalAmountWords} {currencyUnit}
              </span>
            </div>

            {invoice.explanation && (
              <div className="text-xs text-slate-700 bg-slate-50 border border-slate-200 rounded-lg p-2.5">
                <span className="font-bold block mb-0.5 text-slate-900">توضیحات پیش‌فاکتور:</span>
                <p className="whitespace-pre-line leading-relaxed">{invoice.explanation}</p>
              </div>
            )}

            {totalPipeWeight > 0 && (
              <div className="text-xs text-slate-600">
                مجموع وزن لوله‌ها: <strong>{toPersianDigits(totalPipeWeight.toFixed(3))} کیلوگرم</strong>
              </div>
            )}
          </div>

          {/* Totals Table */}
          <div className="w-full md:w-80 border border-slate-300 rounded-lg overflow-hidden text-xs">
            <table className="w-full">
              <tbody>
                <tr className="border-b border-slate-200 bg-slate-50">
                  <td className="py-1.5 px-3 text-slate-600">جمع کل ناخالص:</td>
                  <td className="py-1.5 px-3 text-left font-mono font-bold text-slate-800">
                    {formatNumber(invoice.subtotal, true)} {currencyUnit}
                  </td>
                </tr>

                {(invoice.includeDiscount || invoice.includeCustomDiscount) && invoice.discountAmount > 0 && (
                  <tr className="border-b border-slate-200 bg-emerald-50 text-emerald-900">
                    <td className="py-1.5 px-3">
                      تخفیف ({invoice.includeCustomDiscount ? 'مبلغ دلخواه' : `${toPersianDigits(invoice.discountPercentage)}%`}):
                    </td>
                    <td className="py-1.5 px-3 text-left font-mono font-bold">
                      - {formatNumber(invoice.discountAmount, true)} {currencyUnit}
                    </td>
                  </tr>
                )}

                {invoice.includeTax && invoice.taxAmount > 0 && (
                  <tr className="border-b border-slate-200 bg-amber-50 text-amber-900">
                    <td className="py-1.5 px-3">
                      مالیات بر ارزش افزوده ({toPersianDigits(invoice.taxPercentage)}%):
                    </td>
                    <td className="py-1.5 px-3 text-left font-mono font-bold">
                      + {formatNumber(invoice.taxAmount, true)} {currencyUnit}
                    </td>
                  </tr>
                )}

                <tr className="bg-slate-900 text-white text-xs sm:text-sm font-bold">
                  <td className="py-2.5 px-3">مبلغ نهایی فاکتور:</td>
                  <td className="py-2.5 px-3 text-left font-mono text-sm sm:text-base text-yellow-300">
                    {formatNumber(invoice.totalAmount, true)} {currencyUnit}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Bank Settlement Block (Displayed in Official Mode or when Bank details exist) */}
        {(isOfficial || settings.bankName || settings.bankAccount || settings.bankIban) && (
          <div className="border border-slate-300 rounded-lg p-2.5 mb-5 bg-slate-50/80 text-[11px] text-slate-800">
            <div className="font-bold text-slate-900 mb-1 border-b border-slate-200 pb-1 flex items-center justify-between">
              <span>اطلاعات حساب و واریز بانکی جهت تسویه</span>
              <span className="text-[10px] text-slate-500 font-normal">لطفاً پس از واریز، فیش مربوطه را ارسال فرمایید</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <div>
                <span className="text-slate-500 ml-1">نام بانک:</span>
                <strong className="text-slate-900">{settings.bankName || 'بانک ملت'}</strong>
              </div>
              <div>
                <span className="text-slate-500 ml-1">شماره حساب:</span>
                <span className="font-mono font-bold text-slate-900">{toPersianDigits(settings.bankAccount || '-')}</span>
              </div>
              <div>
                <span className="text-slate-500 ml-1">شماره شبا (IBAN):</span>
                <span className="font-mono font-bold text-blue-900">{settings.bankIban || '-'}</span>
              </div>
            </div>
          </div>
        )}

        {/* Footer & Signature Stamps */}
        <div className="border-t border-slate-200 pt-4 mt-6 grid grid-cols-2 text-center text-xs text-slate-600">
          <div className="h-24 flex flex-col justify-between items-center border-l border-slate-200 pl-4">
            <span className="font-bold text-slate-800">مهر و امضای خریدار</span>
            <span className="text-[10px] text-slate-400">تأیید مشخصات و قیمت‌های مندرج</span>
          </div>
          <div className="h-24 flex flex-col justify-between items-center pr-4">
            <span className="font-bold text-slate-800">مهر و امضای فروشنده ({settings.name})</span>
            <span className="text-[10px] text-slate-400">اعتبار این پیش‌فاکتور به مدت ۳ روز کاری می‌باشد</span>
          </div>
        </div>
      </div>
    </div>
  );
};
