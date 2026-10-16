import { describe, it, expect } from 'vitest';
import {
  toPersianDigits,
  toEnglishDigits,
  formatNumber,
  parseFormattedNumber,
  getCurrentJalaliDate,
  numberToPersianWords,
  formatPrice,
} from '../persianUtils';

describe('Persian Utils Service - Digits, Dates, Formatting & Words', () => {
  describe('toPersianDigits', () => {
    it('Case 15: converts English digits "1234567890" to Persian "۱۲۳۴۵۶۷۸۹۰"', () => {
      expect(toPersianDigits('1234567890')).toBe('۱۲۳۴۵۶۷۸۹۰');
      expect(toPersianDigits(1234567890)).toBe('۱۲۳۴۵۶۷۸۹۰');
    });

    it('handles mixed strings and preserves non-digits', () => {
      expect(toPersianDigits('SDR 11 - 110mm')).toBe('SDR ۱۱ - ۱۱۰mm');
      expect(toPersianDigits('1,500,000 تومان')).toBe('۱,۵۰۰,۰۰۰ تومان');
    });

    it('handles empty or nullish inputs', () => {
      expect(toPersianDigits('')).toBe('');
      expect(toPersianDigits(null as unknown as string)).toBe('');
      expect(toPersianDigits(undefined as unknown as string)).toBe('');
    });
  });

  describe('toEnglishDigits', () => {
    it('Case 16: converts Persian digits "۱۲۳۴۵۶۷۸۹۰" to English "1234567890"', () => {
      expect(toEnglishDigits('۱۲۳۴۵۶۷۸۹۰')).toBe('1234567890');
    });

    it('Case 16 (cont): converts Arabic digits "١٢٣٤٥٦٧٨٩٠" to English "1234567890"', () => {
      expect(toEnglishDigits('١٢٣٤٥٦٧٨٩٠')).toBe('1234567890');
    });

    it('handles mixed Persian and Arabic digits', () => {
      expect(toEnglishDigits('۱۲۳۴-١٢٣٤')).toBe('1234-1234');
    });

    it('handles empty or nullish inputs', () => {
      expect(toEnglishDigits('')).toBe('');
      expect(toEnglishDigits(null as unknown as string)).toBe('');
      expect(toEnglishDigits(undefined as unknown as string)).toBe('');
    });
  });

  describe('formatNumber & parseFormattedNumber', () => {
    it('Case 17: adds comma separation every 3 digits for English numerals', () => {
      expect(formatNumber(1500000)).toBe('1,500,000');
      expect(formatNumber(100)).toBe('100');
      expect(formatNumber(1000)).toBe('1,000');
      expect(formatNumber(29830000)).toBe('29,830,000');
    });

    it('Case 18: formats with comma separation and Persian digits when usePersianDigits is true', () => {
      expect(formatNumber(1500000, true)).toBe('۱,۵۰۰,۰۰۰');
      expect(formatNumber(29830000, true)).toBe('۲۹,۸۳۰,۰۰۰');
    });

    it('supports decimal values without loss of precision', () => {
      expect(formatNumber(12.5)).toBe('12.5');
      expect(formatNumber(12.5, true)).toBe('۱۲.۵');
      expect(formatNumber(1500.5)).toBe('1,500.5');
      expect(formatNumber(1500.5, true)).toBe('۱,۵۰۰.۵');
      expect(formatNumber(1500.555, false, 2)).toBe('1,500.56');
    });

    it('handles zero, null, and NaN safely', () => {
      expect(formatNumber(0)).toBe('0');
      expect(formatNumber(0, true)).toBe('۰');
      expect(formatNumber(null as unknown as number)).toBe('0');
      expect(formatNumber(undefined as unknown as number)).toBe('0');
      expect(formatNumber('invalid')).toBe('0');
    });

    it('parseFormattedNumber cleanly parses comma-separated Persian and Arabic numbers', () => {
      expect(parseFormattedNumber('۱,۵۰۰,۰۰۰')).toBe(1500000);
      expect(parseFormattedNumber('١,٥٠٠,٠٠٠')).toBe(1500000);
      expect(parseFormattedNumber('1,500,000')).toBe(1500000);
      expect(parseFormattedNumber('۱۲.۵')).toBe(12.5);
      expect(parseFormattedNumber('')).toBe(0);
      expect(parseFormattedNumber(null as unknown as string)).toBe(0);
    });

    it('formatPrice formats with specified currency unit', () => {
      expect(formatPrice(1500000, 'تومان')).toBe('1,500,000 تومان');
      expect(formatPrice(15000000, 'ریال', true)).toBe('۱۵,۰۰۰,۰۰۰ ریال');
    });
  });

  describe('getCurrentJalaliDate', () => {
    it('Case 19: returns current Jalali date matching YYYY/MM/DD pattern', () => {
      const jalaliDate = getCurrentJalaliDate();
      const jalaliPattern = /^\d{4}\/\d{2}\/\d{2}$/;
      expect(jalaliDate).toMatch(jalaliPattern);

      const parts = jalaliDate.split('/');
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      const day = parseInt(parts[2], 10);

      // Jalali year should be in current era (e.g. 1400-1450)
      expect(year).toBeGreaterThanOrEqual(1400);
      expect(year).toBeLessThan(1500);
      expect(month).toBeGreaterThanOrEqual(1);
      expect(month).toBeLessThanOrEqual(12);
      expect(day).toBeGreaterThanOrEqual(1);
      expect(day).toBeLessThanOrEqual(31);
    });
  });

  describe('numberToPersianWords', () => {
    it('Case 20: converts 0 to "صفر"', () => {
      expect(numberToPersianWords(0)).toBe('صفر');
    });

    it('Case 21: converts 1000 to "یک هزار"', () => {
      expect(numberToPersianWords(1000)).toBe('یک هزار');
    });

    it('Case 22: converts 1,250,000 to "یک میلیون و دویست و پنجاه هزار"', () => {
      expect(numberToPersianWords(1250000)).toBe('یک میلیون و دویست و پنجاه هزار');
    });

    it('Case 23: converts 1,000,000,000 to "یک میلیارد"', () => {
      expect(numberToPersianWords(1000000000)).toBe('یک میلیارد');
    });

    it('converts complex numbers and trillions', () => {
      expect(numberToPersianWords(29830000)).toBe('بیست و نه میلیون و هشتصد و سی هزار');
      expect(numberToPersianWords(1000000000000)).toBe('یک تریلیون');
      expect(numberToPersianWords(3542150)).toBe(
        'سه میلیون و پانصد و چهل و دو هزار و صد و پنجاه'
      );
    });
  });
});
