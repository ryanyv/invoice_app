import { describe, it, expect } from 'vitest';
import seedData from '../../data/seedData.json';
import {
  getWeightPerMeter,
  computePipeItem,
  calculateLengthFromMass,
  calculatePricePerKgFromTotal,
  getApplicableDiscount,
  computeInvoiceSummary,
} from '../calculator';
import { WeightTable, DiscountRule, PipeItem, FittingItem } from '../../types';

const weightTable = seedData.weightTable as unknown as WeightTable;

describe('Calculator Service - DIN 8074 Lookups & Pipe Calculations', () => {
  describe('DIN 8074 weight lookup accuracy (getWeightPerMeter)', () => {
    it('Case 1: correctly looks up 110mm SDR11 -> 3.14 kg/m', () => {
      const wpm = getWeightPerMeter(110, 11, weightTable);
      expect(wpm).toBe(3.14);
    });

    it('Case 2: correctly looks up 63mm SDR11 -> 1.05 kg/m', () => {
      const wpm = getWeightPerMeter(63, 11, weightTable);
      expect(wpm).toBe(1.05);
    });

    it('Case 3: correctly looks up 200mm SDR17 -> 7.05 kg/m', () => {
      const wpm = getWeightPerMeter(200, 17, weightTable);
      expect(wpm).toBe(7.05);
    });

    it('Case 4: returns null for invalid diameter (999mm)', () => {
      const wpm = getWeightPerMeter(999, 11, weightTable);
      expect(wpm).toBeNull();
    });

    it('Case 5: returns null for invalid SDR (SDR 99) on valid diameter (110mm)', () => {
      const wpm = getWeightPerMeter(110, 99, weightTable);
      expect(wpm).toBeNull();
    });
  });

  describe('Forward Pipe Calculation (computePipeItem)', () => {
    it('Case 6: calculates 100m 110mm SDR11 @ 95,000 Toman/kg correctly', () => {
      const item = computePipeItem(
        {
          peGrade: 'PE100',
          diameter: 110,
          sdr: 11,
          pn: 16,
          length: 100,
          pricePerKg: 95000,
        },
        weightTable
      );

      expect(item).not.toBeNull();
      if (!item) return;

      expect(item.weightPerMeter).toBe(3.14);
      expect(item.totalWeight).toBe(314);
      expect(item.totalPrice).toBe(29830000); // 314 kg * 95,000 = 29,830,000
      expect(item.length).toBe(100);
      expect(item.type).toBe('pipe');
    });

    it('Case 7: calculates decimal length 12.5m 63mm SDR11 @ 88,000 Toman/kg correctly', () => {
      const item = computePipeItem(
        {
          peGrade: 'PE100',
          diameter: 63,
          sdr: 11,
          pn: 16,
          length: 12.5,
          pricePerKg: 88000,
        },
        weightTable
      );

      expect(item).not.toBeNull();
      if (!item) return;

      expect(item.weightPerMeter).toBe(1.05);
      expect(item.totalWeight).toBe(13.125); // 1.05 * 12.5 = 13.125
      expect(item.totalPrice).toBe(1155000); // 13.125 * 88,000 = 1,155,000
      expect(item.length).toBe(12.5);
    });

    it('returns null for nonexistent diameter/SDR combinations', () => {
      const item = computePipeItem(
        {
          peGrade: 'PE100',
          diameter: 999,
          sdr: 11,
          pn: 16,
          length: 100,
          pricePerKg: 90000,
        },
        weightTable
      );
      expect(item).toBeNull();
    });
  });

  describe('Reverse Pipe Calculations', () => {
    it('Case 8: calculates pipe length from mass: 314 kg of 110mm SDR11 -> 100.00 meters', () => {
      const length = calculateLengthFromMass(314, 110, 11, weightTable);
      expect(length).toBe(100.0);
    });

    it('returns null for zero, negative mass, or unknown diameter', () => {
      expect(calculateLengthFromMass(0, 110, 11, weightTable)).toBeNull();
      expect(calculateLengthFromMass(-50, 110, 11, weightTable)).toBeNull();
      expect(calculateLengthFromMass(314, 999, 11, weightTable)).toBeNull();
    });

    it('Case 9: calculates unit price per kg from total price: 29,830,000 / 314 kg -> 95,000', () => {
      const pricePerKg = calculatePricePerKgFromTotal(29830000, 314);
      expect(pricePerKg).toBe(95000);
    });

    it('returns null for invalid inputs in calculatePricePerKgFromTotal', () => {
      expect(calculatePricePerKgFromTotal(0, 314)).toBeNull();
      expect(calculatePricePerKgFromTotal(29830000, 0)).toBeNull();
      expect(calculatePricePerKgFromTotal(29830000, -10)).toBeNull();
    });
  });

  describe('Stepped Discount Tiers (getApplicableDiscount)', () => {
    const discountRules: DiscountRule[] = [
      { threshold: 10000000, percentage: 3 },
      { threshold: 25000000, percentage: 5 },
      { threshold: 50000000, percentage: 8 },
    ];

    it('Case 10: evaluates stepped discount tiers accurately', () => {
      // Below lowest tier
      expect(getApplicableDiscount(5000000, discountRules)).toBe(0);
      // Exactly at tier 1
      expect(getApplicableDiscount(10000000, discountRules)).toBe(3);
      // In between tier 1 and tier 2
      expect(getApplicableDiscount(15000000, discountRules)).toBe(3);
      // Exactly at tier 2
      expect(getApplicableDiscount(25000000, discountRules)).toBe(5);
      // Between tier 2 and tier 3
      expect(getApplicableDiscount(40000000, discountRules)).toBe(5);
      // At or above tier 3
      expect(getApplicableDiscount(50000000, discountRules)).toBe(8);
      expect(getApplicableDiscount(100000000, discountRules)).toBe(8);
    });

    it('returns 0 when discount rules list is empty', () => {
      expect(getApplicableDiscount(50000000, [])).toBe(0);
    });
  });

  describe('Invoice Financial Summary (computeInvoiceSummary)', () => {
    const samplePipeItem: PipeItem = {
      id: 'pipe-1',
      type: 'pipe',
      peGrade: 'PE100',
      diameter: 110,
      sdr: 11,
      pn: 16,
      length: 100,
      weightPerMeter: 3.14,
      totalWeight: 314,
      pricePerKg: 95000,
      totalPrice: 29830000,
    };

    const sampleFittingItem: FittingItem = {
      id: 'fitting-1',
      type: 'fitting',
      category: 'electrofusion',
      connectionType: 'electrofusion',
      product: 'سه راهی ۹۰ مساوی ۱۱۰',
      material: 'HDPE100',
      pressure: '16',
      size: '110',
      quantity: 2,
      unitPrice: 500000,
      totalPrice: 1000000,
    };

    it('Case 11: computes invoice summary without discount and tax', () => {
      const summary = computeInvoiceSummary({
        items: [samplePipeItem, sampleFittingItem],
        includeDiscount: false,
        discountPercentage: 0,
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: false,
        taxPercentage: 0,
      });

      expect(summary.subtotal).toBe(30830000);
      expect(summary.discountAmount).toBe(0);
      expect(summary.amountAfterDiscount).toBe(30830000);
      expect(summary.taxAmount).toBe(0);
      expect(summary.totalAmount).toBe(30830000);
      expect(summary.totalPipeWeight).toBe(314);
      expect(summary.totalFittingCount).toBe(2);
    });

    it('Case 12: applies 10% VAT tax computation', () => {
      const mockItem: PipeItem = {
        ...samplePipeItem,
        totalPrice: 10000000,
      };

      const summary = computeInvoiceSummary({
        items: [mockItem],
        includeDiscount: false,
        discountPercentage: 0,
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(10000000);
      expect(summary.discountAmount).toBe(0);
      expect(summary.amountAfterDiscount).toBe(10000000);
      expect(summary.taxAmount).toBe(1000000); // 10% of 10,000,000
      expect(summary.totalAmount).toBe(11000000); // 10,000,000 + 1,000,000
    });

    it('Case 13: deducts percentage discount first, then computes VAT tax on post-discount amount', () => {
      const mockItem: PipeItem = {
        ...samplePipeItem,
        totalPrice: 10000000,
      };

      const summary = computeInvoiceSummary({
        items: [mockItem],
        includeDiscount: true,
        discountPercentage: 10,
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(10000000);
      expect(summary.discountAmount).toBe(1000000); // 10% of 10,000,000
      expect(summary.amountAfterDiscount).toBe(9000000); // 10,000,000 - 1,000,000
      expect(summary.taxAmount).toBe(900000); // 10% of 9,000,000
      expect(summary.totalAmount).toBe(9900000); // 9,000,000 + 900,000
    });

    it('Case 14: clamps custom discount when it exceeds subtotal (totalAmount >= 0)', () => {
      const mockItem: PipeItem = {
        ...samplePipeItem,
        totalPrice: 5000000,
      };

      const summary = computeInvoiceSummary({
        items: [mockItem],
        includeDiscount: false,
        discountPercentage: 0,
        includeCustomDiscount: true,
        customDiscountAmount: 10000000, // Exceeds 5,000,000 subtotal
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(5000000);
      expect(summary.discountAmount).toBe(5000000); // Clamped to subtotal
      expect(summary.amountAfterDiscount).toBe(0);
      expect(summary.taxAmount).toBe(0);
      expect(summary.totalAmount).toBe(0);
      expect(summary.totalAmount).toBeGreaterThanOrEqual(0);
    });
  });
});
