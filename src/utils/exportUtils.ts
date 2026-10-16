import { CompanySettings, FittingItem, Invoice, PipeItem } from '../types';
import {
  formatNumber,
  numberToPersianWords,
  toPersianDigits,
} from '../services/persianUtils';

/**
 * Escapes a cell value for RFC 4180 CSV compliance
 */
function escapeCsv(value: string | number | boolean | null | undefined): string {
  if (value === null || value === undefined) return '""';
  const str = String(value);
  // If string contains quotes, commas, newlines, or carriage returns, wrap in quotes and escape internal quotes
  const escaped = str.replace(/"/g, '""');
  return `"${escaped}"`;
}

/**
 * Generates an Excel-compatible RFC 4180 CSV string with UTF-8 BOM
 * ensuring Persian / RTL characters open cleanly without mojibake.
 */
export function generateInvoiceCsvString(
  invoice: Invoice,
  settings: CompanySettings
): string {
  const currency = invoice.currencyUnit || settings.currencyUnit || 'تومان';
  const isOfficial = Boolean(invoice.isOfficial);
  const rows: string[] = [];

  // UTF-8 Byte Order Mark
  const BOM = '\uFEFF';

  // 1. Company Header
  const docTitle = isOfficial
    ? 'صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)'
    : 'پیش‌فاکتور فروش کالا و خدمات';

  rows.push([escapeCsv(settings.name), escapeCsv(docTitle)].join(','));
  rows.push(
    [
      escapeCsv(`نشانی: ${settings.address}`),
      escapeCsv(`تلفن: ${settings.phone}`),
      settings.fax ? escapeCsv(`فکس: ${settings.fax}`) : escapeCsv(''),
    ].join(',')
  );

  if (isOfficial) {
    rows.push(
      [
        escapeCsv(`کد اقتصادی فروشنده: ${settings.economicCode || '-'}`),
        escapeCsv(`شناسه ملی: ${settings.nationalId || '-'}`),
        escapeCsv(`شماره ثبت: ${settings.registrationNumber || '-'}`),
        escapeCsv(`کد پستی: ${settings.postalCode || '-'}`),
      ].join(',')
    );
  }

  rows.push(''); // Blank row

  // 2. Invoice & Customer Metadata
  rows.push(
    [
      escapeCsv('شماره فاکتور'),
      escapeCsv('تاریخ صدور'),
      escapeCsv('خریدار'),
      escapeCsv('تلفن خریدار'),
      escapeCsv('کد اقتصادی خریدار'),
      escapeCsv('شناسه/کد ملی خریدار'),
      escapeCsv('کد پستی خریدار'),
      escapeCsv('نشانی خریدار'),
      escapeCsv('واحد پول'),
    ].join(',')
  );

  rows.push(
    [
      escapeCsv(invoice.invoiceNumber),
      escapeCsv(invoice.jalaliDate),
      escapeCsv(invoice.customerName),
      escapeCsv(invoice.customerPhone || '-'),
      escapeCsv(invoice.customerEconomicCode || '-'),
      escapeCsv(invoice.customerNationalId || '-'),
      escapeCsv(invoice.customerPostalCode || '-'),
      escapeCsv(invoice.customerAddress || '-'),
      escapeCsv(currency),
    ].join(',')
  );

  rows.push(''); // Blank row

  // 3. Pipe Items Section
  const pipeItems = invoice.items.filter((i): i is PipeItem => i.type === 'pipe');
  if (pipeItems.length > 0) {
    rows.push([escapeCsv('--- جدول اقلام لوله پلی‌اتیلن ---')].join(','));
    rows.push(
      [
        escapeCsv('ردیف'),
        escapeCsv('شرح کالا'),
        escapeCsv('قطر (mm)'),
        escapeCsv('SDR'),
        escapeCsv('فشار (PN)'),
        escapeCsv('گرید مواد'),
        escapeCsv('طول (متر)'),
        escapeCsv('وزن واحد (kg/m)'),
        escapeCsv('وزن کل (kg)'),
        escapeCsv(`قیمت هر کیلو (${currency})`),
        escapeCsv(`قیمت کل (${currency})`),
      ].join(',')
    );

    let totalPipeLength = 0;
    let totalPipeWeight = 0;
    let totalPipePrice = 0;

    pipeItems.forEach((p, idx) => {
      totalPipeLength += p.length || 0;
      totalPipeWeight += p.totalWeight || 0;
      totalPipePrice += p.totalPrice || 0;

      rows.push(
        [
          escapeCsv(idx + 1),
          escapeCsv(`لوله پلی‌اتیلن ${p.diameter} میلیمتر SDR${p.sdr} PN${p.pn} ${p.peGrade}`),
          escapeCsv(p.diameter),
          escapeCsv(p.sdr),
          escapeCsv(p.pn),
          escapeCsv(p.peGrade),
          escapeCsv(p.length),
          escapeCsv(p.weightPerMeter.toFixed(3)),
          escapeCsv((p.totalWeight || 0).toFixed(3)),
          escapeCsv(p.pricePerKg),
          escapeCsv(p.totalPrice),
        ].join(',')
      );
    });

    // Pipe summary row
    rows.push(
      [
        escapeCsv('جمع اقلام لوله'),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(`${totalPipeLength} متر`),
        escapeCsv(''),
        escapeCsv(`${totalPipeWeight.toFixed(3)} کیلو`),
        escapeCsv(''),
        escapeCsv(totalPipePrice),
      ].join(',')
    );

    rows.push(''); // Blank row
  }

  // 4. Fitting Items Section
  const fittingItems = invoice.items.filter((i): i is FittingItem => i.type === 'fitting');
  if (fittingItems.length > 0) {
    rows.push([escapeCsv('--- جدول اقلام اتصالات پلی‌اتیلن ---')].join(','));
    rows.push(
      [
        escapeCsv('ردیف'),
        escapeCsv('نام محصول'),
        escapeCsv('دسته‌بندی'),
        escapeCsv('نوع اتصال'),
        escapeCsv('سایز (mm)'),
        escapeCsv('فشار (bar)'),
        escapeCsv('تعداد'),
        escapeCsv(`قیمت واحد (${currency})`),
        escapeCsv(`قیمت کل (${currency})`),
      ].join(',')
    );

    let totalFittingQty = 0;
    let totalFittingPrice = 0;

    fittingItems.forEach((f, idx) => {
      totalFittingQty += f.quantity || 0;
      totalFittingPrice += f.totalPrice || 0;

      rows.push(
        [
          escapeCsv(idx + 1),
          escapeCsv(f.product),
          escapeCsv(f.category),
          escapeCsv(f.connectionType),
          escapeCsv(f.size),
          escapeCsv(f.pressure),
          escapeCsv(f.quantity),
          escapeCsv(f.unitPrice),
          escapeCsv(f.totalPrice),
        ].join(',')
      );
    });

    // Fitting summary row
    rows.push(
      [
        escapeCsv('جمع اقلام اتصالات'),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(''),
        escapeCsv(`${totalFittingQty} عدد`),
        escapeCsv(''),
        escapeCsv(totalFittingPrice),
      ].join(',')
    );

    rows.push(''); // Blank row
  }

  // 5. Financial Summary Section
  rows.push([escapeCsv('--- خلاصه مالی پیش‌فاکتور ---')].join(','));
  rows.push([escapeCsv('جمع کل ناخالص:'), escapeCsv(invoice.subtotal), escapeCsv(currency)].join(','));

  if ((invoice.includeDiscount || invoice.includeCustomDiscount) && invoice.discountAmount > 0) {
    const discLabel = invoice.includeCustomDiscount
      ? 'تخفیف (مبلغ دلخواه):'
      : `تخفیف (${invoice.discountPercentage}%):`;
    rows.push([escapeCsv(discLabel), escapeCsv(`-${invoice.discountAmount}`), escapeCsv(currency)].join(','));
  }

  if (invoice.includeTax && invoice.taxAmount > 0) {
    rows.push(
      [
        escapeCsv(`مالیات بر ارزش افزوده (${invoice.taxPercentage}%):`),
        escapeCsv(`+${invoice.taxAmount}`),
        escapeCsv(currency),
      ].join(',')
    );
  }

  rows.push([escapeCsv('مبلغ نهایی فاکتور:'), escapeCsv(invoice.totalAmount), escapeCsv(currency)].join(','));
  rows.push(
    [
      escapeCsv('مبلغ به حروف:'),
      escapeCsv(`${numberToPersianWords(invoice.totalAmount)} ${currency}`),
    ].join(',')
  );

  // 6. Bank Details (if present)
  if (settings.bankName || settings.bankAccount || settings.bankIban) {
    rows.push('');
    rows.push([escapeCsv('--- اطلاعات واریز بانکی ---')].join(','));
    rows.push(
      [
        escapeCsv(`بانک: ${settings.bankName || '-'}`),
        escapeCsv(`شماره حساب: ${settings.bankAccount || '-'}`),
        escapeCsv(`شماره شبا: ${settings.bankIban || '-'}`),
      ].join(',')
    );
  }

  // 7. Explanation
  if (invoice.explanation) {
    rows.push('');
    rows.push([escapeCsv('توضیحات:'), escapeCsv(invoice.explanation)].join(','));
  }

  return BOM + rows.join('\r\n');
}

/**
 * Triggers direct download or Electron IPC file save for CSV
 */
export function exportInvoiceToCsv(invoice: Invoice, settings: CompanySettings): void {
  const csvContent = generateInvoiceCsvString(invoice, settings);
  const cleanCustomer = (invoice.customerName || 'مشتری').replace(/\s+/g, '_');
  const filename = `پیش_فاکتور_${invoice.invoiceNumber}_${cleanCustomer}.csv`;

  // Electron IPC save if available
  if (typeof window !== 'undefined' && (window as any).electronAPI?.exportFile) {
    (window as any).electronAPI
      .exportFile({
        defaultName: filename,
        content: csvContent,
        extensions: ['csv'],
      })
      .then((res: any) => {
        if (res?.success) {
          alert(`فایل اکسل/CSV با موفقیت ذخیره شد:\n${res.filePath}`);
        }
      })
      .catch((err: any) => {
        console.error('Electron CSV export error:', err);
        fallbackDownloadCsv(filename, csvContent);
      });
  } else {
    fallbackDownloadCsv(filename, csvContent);
  }
}

function fallbackDownloadCsv(filename: string, content: string): void {
  if (typeof document === 'undefined') return;
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Formats a clean Persian WhatsApp / SMS text snippet with emojis and RTL layout
 */
export function generateInvoiceWhatsAppSummary(
  invoice: Invoice,
  settings: CompanySettings
): string {
  const currency = invoice.currencyUnit || settings.currencyUnit || 'تومان';
  const isOfficial = Boolean(invoice.isOfficial);

  const lines: string[] = [];

  // Header
  lines.push(
    isOfficial
      ? `📋 صورتحساب رسمی فروش کالا و خدمات (ماده ۱۶۹ م.م)`
      : `📋 پیش‌فاکتور شرکت ${settings.name}`
  );
  lines.push(
    `شماره: ${toPersianDigits(invoice.invoiceNumber)} | تاریخ: ${toPersianDigits(invoice.jalaliDate)}`
  );
  lines.push(`خریدار محترم: ${invoice.customerName}`);
  if (invoice.customerPhone) {
    lines.push(`تلفن خریدار: ${toPersianDigits(invoice.customerPhone)}`);
  }

  // Items
  const pipeItems = invoice.items.filter((i): i is PipeItem => i.type === 'pipe');
  const fittingItems = invoice.items.filter((i): i is FittingItem => i.type === 'fitting');

  if (pipeItems.length > 0) {
    lines.push('');
    lines.push('🔹 اقلام لوله پلی‌اتیلن:');
    pipeItems.forEach((p, idx) => {
      const weightStr = p.totalWeight ? ` (${toPersianDigits(p.totalWeight.toFixed(2))} کیلو)` : '';
      lines.push(
        `${toPersianDigits(idx + 1)}. لوله ${toPersianDigits(p.diameter)} SDR${toPersianDigits(p.sdr)} PN${toPersianDigits(p.pn)} ${p.peGrade} - ${toPersianDigits(p.length)} متر${weightStr}: ${formatNumber(p.totalPrice, true)} ${currency}`
      );
    });
  }

  if (fittingItems.length > 0) {
    lines.push('');
    lines.push('🔸 اقلام اتصالات:');
    fittingItems.forEach((f, idx) => {
      lines.push(
        `${toPersianDigits(idx + 1)}. ${f.product} (${f.category}) سایز ${toPersianDigits(f.size)} - ${toPersianDigits(f.quantity)} عدد: ${formatNumber(f.totalPrice, true)} ${currency}`
      );
    });
  }

  // Totals
  lines.push('');
  lines.push('📊 خلاصه مالی:');
  lines.push(`▫️ جمع ناخالص: ${formatNumber(invoice.subtotal, true)} ${currency}`);

  if ((invoice.includeDiscount || invoice.includeCustomDiscount) && invoice.discountAmount > 0) {
    lines.push(`▫️ تخفیف: ${formatNumber(invoice.discountAmount, true)} ${currency}`);
  }

  if (invoice.includeTax && invoice.taxAmount > 0) {
    lines.push(
      `▫️ مالیات ارزش افزوده (${toPersianDigits(invoice.taxPercentage)}%): ${formatNumber(invoice.taxAmount, true)} ${currency}`
    );
  }

  lines.push(`💰 مبلغ نهایی قابل پرداخت: ${formatNumber(invoice.totalAmount, true)} ${currency}`);
  lines.push(`(${numberToPersianWords(invoice.totalAmount)} ${currency})`);

  // Banking
  if (settings.bankName || settings.bankAccount || settings.bankIban) {
    lines.push('');
    lines.push('💳 اطلاعات واریز:');
    if (settings.bankName) lines.push(`بانک: ${settings.bankName}`);
    if (settings.bankAccount) lines.push(`شماره حساب: ${toPersianDigits(settings.bankAccount)}`);
    if (settings.bankIban) lines.push(`شماره شبا: ${settings.bankIban}`);
  }

  // Company Contact
  lines.push('');
  lines.push(`📞 تماس: ${toPersianDigits(settings.phone)}`);
  if (settings.address) {
    lines.push(`🏢 نشانی: ${settings.address}`);
  }
  lines.push('اعتبار پیش‌فاکتور: ۳ روز کاری');

  return lines.join('\n');
}

/**
 * Copies the WhatsApp summary directly to clipboard
 */
export async function copyInvoiceWhatsAppSummary(
  invoice: Invoice,
  settings: CompanySettings
): Promise<boolean> {
  const summary = generateInvoiceWhatsAppSummary(invoice, settings);
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(summary);
      return true;
    } catch {
      return fallbackCopyText(summary);
    }
  } else {
    return fallbackCopyText(summary);
  }
}

function fallbackCopyText(text: string): boolean {
  if (typeof document === 'undefined') return false;
  try {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    const successful = document.execCommand('copy');
    document.body.removeChild(textArea);
    return successful;
  } catch (err) {
    console.error('Fallback copy failed:', err);
    return false;
  }
}
