import { toJalaali } from 'jalaali-js';

/**
 * Returns the current date in Jalali format: YYYY/MM/DD
 */
export function getCurrentJalaliDate(): string {
  const now = new Date();
  const j = toJalaali(now.getFullYear(), now.getMonth() + 1, now.getDate());
  const month = String(j.jm).padStart(2, '0');
  const day = String(j.jd).padStart(2, '0');
  return `${j.jy}/${month}/${day}`;
}

/**
 * Converts English digits in a string or number to Persian digits
 */
export function toPersianDigits(input: string | number): string {
  if (input === null || input === undefined) return '';
  const str = String(input);
  const persianDigits = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
  return str.replace(/[0-9]/g, (w) => persianDigits[parseInt(w, 10)]);
}

/**
 * Formats a number with comma separators (e.g., 1,500,000)
 */
export function formatNumber(num: number | string, usePersianDigits = false): string {
  if (num === null || num === undefined || isNaN(Number(num))) return '۰';
  const parts = Math.round(Number(num)).toString().split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const formatted = parts.join('.');
  return usePersianDigits ? toPersianDigits(formatted) : formatted;
}

/**
 * Formats price with currency unit
 */
export function formatPrice(amount: number, unit = 'تومان', usePersianDigits = false): string {
  return `${formatNumber(amount, usePersianDigits)} ${unit}`;
}

// Persian number words mapping
const ones = ['', 'یک', 'دو', 'سه', 'چهار', 'پنج', 'شش', 'هفت', 'هشت', 'نه'];
const teens = [
  'ده',
  'یازده',
  'دوازده',
  'سیزده',
  'چهارده',
  'پانزده',
  'شانزده',
  'هفده',
  'هجده',
  'نوزده',
];
const tens = ['', '', 'بیست', 'سی', 'چهل', 'پنجاه', 'شصت', 'هفتاد', 'هشتاد', 'نود'];
const hundreds = [
  '',
  'صد',
  'دویست',
  'سیصد',
  'چهارصد',
  'پانصد',
  'ششصد',
  'هفتصد',
  'هشتصد',
  'نهصد',
];
const scales = ['', 'هزار', 'میلیون', 'میلیارد', 'تریلیون'];

function convertThreeDigits(num: number): string {
  if (num === 0) return '';
  const result: string[] = [];

  const h = Math.floor(num / 100);
  const remainder = num % 100;

  if (h > 0) {
    result.push(hundreds[h]);
  }

  if (remainder > 0) {
    if (remainder < 10) {
      result.push(ones[remainder]);
    } else if (remainder < 20) {
      result.push(teens[remainder - 10]);
    } else {
      const t = Math.floor(remainder / 10);
      const o = remainder % 10;
      result.push(tens[t]);
      if (o > 0) {
        result.push(ones[o]);
      }
    }
  }

  return result.join(' و ');
}

/**
 * Converts any positive integer to Persian words
 */
export function numberToPersianWords(value: number): string {
  const num = Math.floor(Math.abs(Number(value) || 0));
  if (num === 0) return 'صفر';

  const chunks: number[] = [];
  let temp = num;

  while (temp > 0) {
    chunks.push(temp % 1000);
    temp = Math.floor(temp / 1000);
  }

  const wordsParts: string[] = [];

  for (let i = chunks.length - 1; i >= 0; i--) {
    const chunk = chunks[i];
    if (chunk === 0) continue;

    const chunkWords = convertThreeDigits(chunk);
    const scale = scales[i];

    if (scale) {
      wordsParts.push(`${chunkWords} ${scale}`);
    } else {
      wordsParts.push(chunkWords);
    }
  }

  return wordsParts.join(' و ');
}

