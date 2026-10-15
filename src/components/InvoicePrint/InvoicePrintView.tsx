import React from 'react';
import { Printer, ArrowRight } from 'lucide-react';
import { CompanySettings, Invoice, PipeItem, FittingItem } from '../../types';
import { formatNumber, numberToPersianWords, toPersianDigits } from '../../services/persianUtils';

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
  const handlePrint = async () => {
    // Check if running inside Electron with savePdf IPC
    if ((window as any).electronAPI?.savePdf) {
      const defaultName = `پیش_فاکتور_${invoice.invoiceNumber}_${invoice.customerName.replace(/\s+/g, '_')}.pdf`;
      const res = await (window as any).electronAPI.savePdf({ defaultName });
      if (res.success) {
        alert(`فایل PDF با موفقیت ذخیره شد:\n${res.filePath}`);
      } else if (!res.canceled) {
        window.print();
      }
    } else {
      window.print();
    }
  };

  const pipeItems = invoice.items.filter((i): i is PipeItem => i.type === 'pipe');
  const fittingItems = invoice.items.filter((i): i is FittingItem => i.type === 'fitting');

  const currencyUnit = invoice.currencyUnit || settings.currencyUnit || 'تومان';
  const totalPipeWeight = pipeItems.reduce((acc, p) => acc + (p.totalWeight || 0), 0);
  const totalAmountWords = numberToPersianWords(invoice.totalAmount);

  return (
    <div className="min-h-screen bg-slate-900 py-8 px-4 print:p-0 print:bg-white text-slate-900">
      {/* Action Bar (Hidden when printing) */}
      <div className="no-print max-w-5xl mx-auto mb-6 flex items-center justify-between bg-slate-800 p-4 rounded-2xl border border-slate-700 shadow-xl">
        <button
          onClick={onBack}
          className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-xl text-sm font-medium transition-colors cursor-pointer"
        >
          <ArrowRight className="w-4 h-4" />
          <span>بازگشت به ویرایش پیش‌فاکتور</span>
        </button>

        <div className="flex items-center gap-3">
          <button
            onClick={handlePrint}
            className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl font-medium shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
          >
            <Printer className="w-5 h-5" />
            <span>چاپ / ذخیره به صورت PDF</span>
          </button>
        </div>
      </div>

      {/* Printable Sheet (Standard A4 Landscape) */}
      <div className="print-container max-w-5xl mx-auto bg-white p-8 sm:p-12 rounded-xl shadow-2xl print:shadow-none print:rounded-none print:p-4 border border-slate-200 text-black text-sm">
        
        {/* Header Section */}
        <div className="flex items-start justify-between border-b-2 border-slate-900 pb-5 mb-6">
          {/* Logo & Company info */}
          <div className="flex items-center gap-4">
            <div className="w-20 h-20 border border-slate-300 rounded-lg p-1 flex items-center justify-center">
              <img src="/logo.png" alt="پلی غرب" className="max-w-full max-h-full object-contain" />
            </div>
            <div>
              <h1 className="text-xl font-extrabold text-slate-900">{settings.name}</h1>
              <p className="text-xs text-slate-600 mt-1 max-w-md">{settings.address}</p>
              <div className="flex items-center gap-4 text-xs text-slate-600 mt-1">
                <span>تلفن: {settings.phone}</span>
                {settings.fax && <span>فکس: {settings.fax}</span>}
              </div>
            </div>
          </div>

          {/* Invoice Meta */}
          <div className="border border-slate-300 rounded-lg p-3 bg-slate-50 min-w-[220px] text-xs space-y-1.5 text-right">
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
        <div className="bg-slate-100/80 rounded-lg p-3.5 mb-6 border border-slate-200 flex flex-wrap items-center justify-between text-xs gap-3">
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

        {/* Section 1: Pipe Items (if any) */}
        {pipeItems.length > 0 && (
          <div className="mb-6">
            <div className="font-bold text-sm text-slate-800 mb-2 border-r-4 border-blue-600 pr-2">
              اقلام لوله پلی‌اتیلن
            </div>
            <table className="w-full text-center border-collapse border border-slate-300 text-xs">
              <thead className="bg-slate-200 text-slate-800 font-bold">
                <tr>
                  <th className="border border-slate-300 py-2 px-1 w-10">ردیف</th>
                  <th className="border border-slate-300 py-2 px-2">قطر (mm)</th>
                  <th className="border border-slate-300 py-2 px-2">SDR</th>
                  <th className="border border-slate-300 py-2 px-2">فشار (PN)</th>
                  <th className="border border-slate-300 py-2 px-2">گرید</th>
                  <th className="border border-slate-300 py-2 px-2">طول (متر)</th>
                  <th className="border border-slate-300 py-2 px-2">وزن/متر (kg)</th>
                  <th className="border border-slate-300 py-2 px-2">وزن کل (kg)</th>
                  <th className="border border-slate-300 py-2 px-2">قیمت/kg ({currencyUnit})</th>
                  <th className="border border-slate-300 py-2 px-2">قیمت کل ({currencyUnit})</th>
                </tr>
              </thead>
              <tbody>
                {pipeItems.map((item, idx) => (
                  <tr key={item.id} className={idx % 2 === 1 ? 'bg-slate-50' : 'bg-white'}>
                    <td className="border border-slate-300 py-1.5 px-1 font-mono">{toPersianDigits(idx + 1)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-bold">{toPersianDigits(item.diameter)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.sdr)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.pn)}</td>
                    <td className="border border-slate-300 py-1.5 px-2">{item.peGrade}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.length)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.weightPerMeter.toFixed(3))}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono font-medium">{toPersianDigits(item.totalWeight.toFixed(3))}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{formatNumber(item.pricePerKg, true)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono font-bold text-slate-900">{formatNumber(item.totalPrice, true)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Section 2: Fitting Items (if any) */}
        {fittingItems.length > 0 && (
          <div className="mb-6">
            <div className="font-bold text-sm text-slate-800 mb-2 border-r-4 border-emerald-600 pr-2">
              اقلام اتصالات پلی‌اتیلن
            </div>
            <table className="w-full text-center border-collapse border border-slate-300 text-xs">
              <thead className="bg-slate-200 text-slate-800 font-bold">
                <tr>
                  <th className="border border-slate-300 py-2 px-1 w-10">ردیف</th>
                  <th className="border border-slate-300 py-2 px-3 text-right">نوع اتصال</th>
                  <th className="border border-slate-300 py-2 px-3 text-right">نام محصول</th>
                  <th className="border border-slate-300 py-2 px-2">فشار (bar)</th>
                  <th className="border border-slate-300 py-2 px-2">سایز (mm)</th>
                  <th className="border border-slate-300 py-2 px-2">تعداد</th>
                  <th className="border border-slate-300 py-2 px-2">قیمت واحد ({currencyUnit})</th>
                  <th className="border border-slate-300 py-2 px-2">قیمت کل ({currencyUnit})</th>
                </tr>
              </thead>
              <tbody>
                {fittingItems.map((item, idx) => (
                  <tr key={item.id} className={idx % 2 === 1 ? 'bg-slate-50' : 'bg-white'}>
                    <td className="border border-slate-300 py-1.5 px-1 font-mono">{toPersianDigits(idx + 1)}</td>
                    <td className="border border-slate-300 py-1.5 px-3 text-right">{item.category} ({item.connectionType})</td>
                    <td className="border border-slate-300 py-1.5 px-3 text-right font-medium">{item.product}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.pressure)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-bold">{toPersianDigits(item.size)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{toPersianDigits(item.quantity)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono">{formatNumber(item.unitPrice, true)}</td>
                    <td className="border border-slate-300 py-1.5 px-2 font-mono font-bold text-slate-900">{formatNumber(item.totalPrice, true)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Financial Summary & Total Box */}
        <div className="flex flex-col md:flex-row justify-between items-start gap-6 border-t-2 border-slate-300 pt-4 mb-6">
          {/* Notes & Amount in Words */}
          <div className="flex-1 space-y-3">
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-xs">
              <span className="font-bold text-blue-950 block mb-1">مبلغ کل قابل پرداخت به حروف:</span>
              <span className="text-blue-900 font-bold text-sm">
                {totalAmountWords} {currencyUnit}
              </span>
            </div>

            {invoice.explanation && (
              <div className="text-xs text-slate-700 bg-slate-50 border border-slate-200 rounded-lg p-3">
                <span className="font-bold block mb-1 text-slate-900">توضیحات پیش‌فاکتور:</span>
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
                  <td className="py-2 px-3 text-slate-600">جمع کل ناخالص:</td>
                  <td className="py-2 px-3 text-left font-mono font-bold text-slate-800">
                    {formatNumber(invoice.subtotal, true)} {currencyUnit}
                  </td>
                </tr>

                {(invoice.includeDiscount || invoice.includeCustomDiscount) && invoice.discountAmount > 0 && (
                  <tr className="border-b border-slate-200 bg-emerald-50 text-emerald-900">
                    <td className="py-2 px-3">
                      تخفیف ({invoice.includeCustomDiscount ? 'مبلغ دلخواه' : `${toPersianDigits(invoice.discountPercentage)}%`}):
                    </td>
                    <td className="py-2 px-3 text-left font-mono font-bold">
                      - {formatNumber(invoice.discountAmount, true)} {currencyUnit}
                    </td>
                  </tr>
                )}

                {invoice.includeTax && invoice.taxAmount > 0 && (
                  <tr className="border-b border-slate-200 bg-amber-50 text-amber-900">
                    <td className="py-2 px-3">
                      مالیات بر ارزش افزوده ({toPersianDigits(invoice.taxPercentage)}%):
                    </td>
                    <td className="py-2 px-3 text-left font-mono font-bold">
                      + {formatNumber(invoice.taxAmount, true)} {currencyUnit}
                    </td>
                  </tr>
                )}

                <tr className="bg-slate-900 text-white text-sm font-bold">
                  <td className="py-3 px-3">مبلغ نهایی فاکتور:</td>
                  <td className="py-3 px-3 text-left font-mono text-base text-yellow-300">
                    {formatNumber(invoice.totalAmount, true)} {currencyUnit}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Footer & Signature Stamps */}
        <div className="border-t border-slate-200 pt-6 mt-8 grid grid-cols-2 text-center text-xs text-slate-600">
          <div className="h-28 flex flex-col justify-between items-center border-l border-slate-200 pl-4">
            <span className="font-bold text-slate-800">مهر و امضای خریدار</span>
            <span className="text-[10px] text-slate-400">تأیید مشخصات و قیمت‌های مندرج</span>
          </div>
          <div className="h-28 flex flex-col justify-between items-center pr-4">
            <span className="font-bold text-slate-800">مهر و امضای شرکت پلی غرب</span>
            <span className="text-[10px] text-slate-400">اعتبار این پیش‌فاکتور به مدت ۳ روز کاری می‌باشد</span>
          </div>
        </div>

      </div>
    </div>
  );
};

