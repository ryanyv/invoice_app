import { describe, it, expect } from 'vitest';
import {
  computeInvoiceSummary,
  estimateFreightVehicle,
  computePackaging,
  isCoilAllowed,
  FREIGHT_VEHICLES,
} from '../calculator';
import { InvoiceItem, PipeItem, FittingItem } from '../../types';

describe('Challenger M5 Adversarial Probes', () => {
  describe('Probe Suite 1: computeInvoiceSummary Edge Cases & Boundary Conditions', () => {
    it('1.1: 0 items array returns all-zero financial summary and 0 vehicles', () => {
      const summary = computeInvoiceSummary({
        items: [],
        includeDiscount: true,
        discountPercentage: 20,
        includeCustomDiscount: true,
        customDiscountAmount: 50000,
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(0);
      expect(summary.discountAmount).toBe(0);
      expect(summary.amountAfterDiscount).toBe(0);
      expect(summary.taxAmount).toBe(0);
      expect(summary.totalAmount).toBe(0);
      expect(summary.totalPipeWeight).toBe(0);
      expect(summary.totalFittingCount).toBe(0);
      expect(summary.maxBranchLengthMeters).toBe(0);
      expect(summary.freight.vehicleCount).toBe(0);
      expect(summary.freight.utilizationPercentage).toBe(0);
      expect(summary.freight.recommendationNote).toContain('باری برای تخصیص ناوگان ثبت نشده است');
    });

    it('1.2: Extreme percentage discount (>100%) is clamped to subtotal (never negative)', () => {
      const item: PipeItem = {
        id: 'p1',
        type: 'pipe',
        peGrade: 'PE100',
        diameter: 63,
        sdr: 11,
        pn: 16,
        length: 10,
        weightPerMeter: 1.05,
        totalWeight: 10.5,
        pricePerKg: 100000,
        totalPrice: 1050000,
      };

      const summary = computeInvoiceSummary({
        items: [item],
        includeDiscount: true,
        discountPercentage: 250, // 250% discount
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(1050000);
      expect(summary.discountAmount).toBe(1050000); // Clamped to subtotal
      expect(summary.amountAfterDiscount).toBe(0);
      expect(summary.taxAmount).toBe(0);
      expect(summary.totalAmount).toBe(0);
      expect(summary.totalAmount).toBeGreaterThanOrEqual(0);
    });

    it('1.3: Extreme custom discount (orders of magnitude larger than subtotal) is clamped to subtotal', () => {
      const item: FittingItem = {
        id: 'f1',
        type: 'fitting',
        category: 'electrofusion',
        connectionType: 'electrofusion',
        product: 'کوپلر ۹۰',
        material: 'HDPE100',
        pressure: '16',
        size: '90',
        quantity: 1,
        unitPrice: 200000,
        totalPrice: 200000,
      };

      const summary = computeInvoiceSummary({
        items: [item],
        includeDiscount: false,
        discountPercentage: 0,
        includeCustomDiscount: true,
        customDiscountAmount: 1000000000000, // 1 Trillion Toman discount on 200k item
        includeTax: true,
        taxPercentage: 10,
      });

      expect(summary.subtotal).toBe(200000);
      expect(summary.discountAmount).toBe(200000); // Clamped
      expect(summary.amountAfterDiscount).toBe(0);
      expect(summary.taxAmount).toBe(0);
      expect(summary.totalAmount).toBe(0);
    });

    it('1.4: Negative discount and tax inputs are ignored safely', () => {
      const item: PipeItem = {
        id: 'p1',
        type: 'pipe',
        peGrade: 'PE100',
        diameter: 63,
        sdr: 11,
        pn: 16,
        length: 10,
        weightPerMeter: 1.05,
        totalWeight: 10.5,
        pricePerKg: 100000,
        totalPrice: 1000000,
      };

      const summary = computeInvoiceSummary({
        items: [item],
        includeDiscount: true,
        discountPercentage: -25, // Negative discount
        includeCustomDiscount: true,
        customDiscountAmount: -50000, // Negative custom discount
        includeTax: true,
        taxPercentage: -10, // Negative tax
      });

      expect(summary.subtotal).toBe(1000000);
      expect(summary.discountAmount).toBe(0); // Not applied
      expect(summary.amountAfterDiscount).toBe(1000000);
      expect(summary.taxAmount).toBe(0); // Not applied
      expect(summary.totalAmount).toBe(1000000);
    });

    it('1.5: Fractional quantities, floating point lengths, and non-integer prices', () => {
      const pipeItem: PipeItem = {
        id: 'p1',
        type: 'pipe',
        peGrade: 'PE100',
        diameter: 63,
        sdr: 11,
        pn: 16,
        length: 15.75, // Fractional length
        weightPerMeter: 1.05,
        totalWeight: 16.5375, // Fractional weight
        pricePerKg: 95432.12,
        totalPrice: 1578210,
      };

      const fittingItem: FittingItem = {
        id: 'f1',
        type: 'fitting',
        category: 'butt_fusion',
        connectionType: 'butt_fusion',
        product: 'تبدیل جوشی ۹۰ به ۶۳',
        material: 'HDPE100',
        pressure: '16',
        size: '90x63',
        quantity: 2.5, // Fractional quantity (e.g. meter-based or split pack)
        unitPrice: 100000,
        totalPrice: 250000,
      };

      const summary = computeInvoiceSummary({
        items: [pipeItem, fittingItem],
        includeDiscount: true,
        discountPercentage: 7.5,
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: true,
        taxPercentage: 9.5,
      });

      expect(summary.subtotal).toBe(1828210);
      // discount = round(1828210 * 0.075) = 137116
      expect(summary.discountAmount).toBe(137116);
      expect(summary.amountAfterDiscount).toBe(1691094);
      // tax = round(1691094 * 0.095) = 160654
      expect(summary.taxAmount).toBe(160654);
      expect(summary.totalAmount).toBe(1851748);
      expect(summary.totalPipeWeight).toBe(16.538); // toFixed(3)
      expect(summary.totalFittingCount).toBe(2.5);
    });

    it('1.6: Items with missing or null optional fields', () => {
      const sparsePipe: InvoiceItem = {
        id: 'sparse-1',
        type: 'pipe',
        peGrade: 'PE100',
        diameter: 110,
        sdr: 11,
        pn: 16,
        length: 50,
        weightPerMeter: 3.14,
        totalWeight: 157,
        pricePerKg: 90000,
        totalPrice: 14130000,
        // packaging omitted
      } as unknown as InvoiceItem;

      const summary = computeInvoiceSummary({
        items: [sparsePipe],
        includeDiscount: false,
        discountPercentage: 0,
        includeCustomDiscount: false,
        customDiscountAmount: 0,
        includeTax: false,
        taxPercentage: 0,
      });

      expect(summary.subtotal).toBe(14130000);
      expect(summary.totalPipeWeight).toBe(157);
      expect(summary.maxBranchLengthMeters).toBe(0);
      expect(summary.freight.vehicle.id).toBe('nissan');
    });
  });

  describe('Probe Suite 2: estimateFreightVehicle Adversarial Cases', () => {
    it('2.1: Zero and negative payload weight returns 0 vehicles safely', () => {
      const recZero = estimateFreightVehicle(0, 0);
      expect(recZero.vehicleCount).toBe(0);
      expect(recZero.totalWeightKg).toBe(0);
      expect(recZero.totalTonnage).toBe(0);
      expect(recZero.utilizationPercentage).toBe(0);
      expect(recZero.recommendationNote).toContain('باری برای تخصیص ناوگان ثبت نشده است');

      const recNeg = estimateFreightVehicle(-1500, 12);
      expect(recNeg.vehicleCount).toBe(0);
      expect(recNeg.totalWeightKg).toBe(0);
      expect(recNeg.totalTonnage).toBe(0);
      expect(recNeg.utilizationPercentage).toBe(0);
    });

    it('2.2: Tiny positive payload (1 gram = 0.001 kg) with 12m branch mandates Flatbed Trailer', () => {
      const recTiny = estimateFreightVehicle(0.001, 12);
      expect(recTiny.vehicle.id).toBe('trailer_flatbed');
      expect(recTiny.vehicleCount).toBe(1);
      expect(recTiny.isLengthConstrained).toBe(true);
      expect(recTiny.utilizationPercentage).toBe(0); // rounds to 0%
    });

    it('2.3: Tiny positive payload (1 gram = 0.001 kg) with 6m branch mandates 6-Wheel Truck', () => {
      const recTiny6m = estimateFreightVehicle(0.001, 6);
      expect(recTiny6m.vehicle.id).toBe('truck_6w');
      expect(recTiny6m.vehicleCount).toBe(1);
      expect(recTiny6m.isLengthConstrained).toBe(true);
    });

    it('2.4: Exact boundary weights across Iranian logistics tiers', () => {
      // Exactly 2000 kg -> Nissan Pickup
      const rec2000 = estimateFreightVehicle(2000, 0);
      expect(rec2000.vehicle.id).toBe('nissan');
      expect(rec2000.vehicleCount).toBe(1);
      expect(rec2000.utilizationPercentage).toBe(100);

      // Exactly 2001 kg -> Khavar
      const rec2001 = estimateFreightVehicle(2001, 0);
      expect(rec2001.vehicle.id).toBe('khavar');
      expect(rec2001.vehicleCount).toBe(1);

      // Exactly 4000 kg -> Khavar (100% util)
      const rec4000 = estimateFreightVehicle(4000, 0);
      expect(rec4000.vehicle.id).toBe('khavar');
      expect(rec4000.vehicleCount).toBe(1);
      expect(rec4000.utilizationPercentage).toBe(100);

      // Exactly 4001 kg -> 6-Wheel Single Truck
      const rec4001 = estimateFreightVehicle(4001, 0);
      expect(rec4001.vehicle.id).toBe('truck_6w');
      expect(rec4001.vehicleCount).toBe(1);

      // Exactly 10000 kg -> 6-Wheel Truck (100% util)
      const rec10000 = estimateFreightVehicle(10000, 0);
      expect(rec10000.vehicle.id).toBe('truck_6w');
      expect(rec10000.vehicleCount).toBe(1);
      expect(rec10000.utilizationPercentage).toBe(100);

      // Exactly 10001 kg -> 10-Wheel Double Truck
      const rec10001 = estimateFreightVehicle(10001, 0);
      expect(rec10001.vehicle.id).toBe('truck_10w');
      expect(rec10001.vehicleCount).toBe(1);

      // Exactly 15000 kg -> 10-Wheel Double Truck (100% util)
      const rec15000 = estimateFreightVehicle(15000, 0);
      expect(rec15000.vehicle.id).toBe('truck_10w');
      expect(rec15000.vehicleCount).toBe(1);
      expect(rec15000.utilizationPercentage).toBe(100);

      // Exactly 15001 kg -> Flatbed Trailer
      const rec15001 = estimateFreightVehicle(15001, 0);
      expect(rec15001.vehicle.id).toBe('trailer_flatbed');
      expect(rec15001.vehicleCount).toBe(1);

      // Exactly 24000 kg -> 1 Flatbed Trailer (100% util)
      const rec24000 = estimateFreightVehicle(24000, 0);
      expect(rec24000.vehicle.id).toBe('trailer_flatbed');
      expect(rec24000.vehicleCount).toBe(1);
      expect(rec24000.utilizationPercentage).toBe(100);

      // Exactly 24001 kg -> 2 Flatbed Trailers
      const rec24001 = estimateFreightVehicle(24001, 0);
      expect(rec24001.vehicle.id).toBe('trailer_flatbed');
      expect(rec24001.vehicleCount).toBe(2);
      expect(rec24001.utilizationPercentage).toBe(50); // 24001 / 48000 = 50%
    });

    it('2.5: Branch length boundary transitions (2.2m vs 2.21m and 6.0m vs 6.01m)', () => {
      // 2.2m or lower with 1000 kg -> Nissan
      const rec22 = estimateFreightVehicle(1000, 2.2);
      expect(rec22.vehicle.id).toBe('nissan');

      // 2.21m with 1000 kg -> Truck 6W (Nissan & Khavar cannot take >2.2m)
      const rec221 = estimateFreightVehicle(1000, 2.21);
      expect(rec221.vehicle.id).toBe('truck_6w');
      expect(rec221.isLengthConstrained).toBe(true);

      // 6.0m with 1000 kg -> Truck 6W
      const rec60 = estimateFreightVehicle(1000, 6.0);
      expect(rec60.vehicle.id).toBe('truck_6w');

      // 6.01m with 1000 kg -> Trailer Flatbed
      const rec601 = estimateFreightVehicle(1000, 6.01);
      expect(rec601.vehicle.id).toBe('trailer_flatbed');
      expect(rec601.isLengthConstrained).toBe(true);
    });
  });

  describe('Probe Suite 3: Packaging Rules Adversarial Cases', () => {
    it('3.1: Boundary diameters for isCoilAllowed', () => {
      expect(isCoilAllowed(0)).toBe(true);
      expect(isCoilAllowed(125)).toBe(true);
      expect(isCoilAllowed(125.0)).toBe(true);
      expect(isCoilAllowed(125.0001)).toBe(false);
      expect(isCoilAllowed(126)).toBe(false);
      expect(isCoilAllowed(140)).toBe(false);
      expect(isCoilAllowed(500)).toBe(false);
    });

    it('3.2: computePackaging fallback when invalid coil or branch length is provided', () => {
      // Invalid coil length (e.g. 0 or negative) falls back to default 100m
      const p1 = computePackaging(63, 200, { preferredType: 'coil', coilLength: 0 });
      expect(p1.unitLength).toBe(100);
      expect(p1.unitCount).toBe(2);

      const p2 = computePackaging(63, 200, { preferredType: 'coil', coilLength: -50 });
      expect(p2.unitLength).toBe(100);
      expect(p2.unitCount).toBe(2);

      // Valid 5m branch length is supported alongside 6m and 12m
      const p5m = computePackaging(110, 30, { preferredType: 'branch', branchLength: 5 });
      expect(p5m.unitLength).toBe(5);
      expect(p5m.unitCount).toBe(6);

      // Invalid branch length (e.g. 8 instead of 5, 6, or 12) falls back to default (6 for <=125, 12 for >125)
      const p3 = computePackaging(110, 30, { preferredType: 'branch', branchLength: 8 as unknown as 6 });
      expect(p3.unitLength).toBe(6);
      expect(p3.unitCount).toBe(5);

      const p4 = computePackaging(160, 30, { preferredType: 'branch', branchLength: 8 as unknown as 12 });
      expect(p4.unitLength).toBe(12);
      expect(p4.unitCount).toBe(3);
    });

    it('3.3: Strict enforcement: >125mm pipe CANNOT be packaged as coil under any condition', () => {
      const p = computePackaging(140, 1000, { preferredType: 'coil', coilLength: 200 });
      expect(p.packagingType).toBe('branch');
      expect(p.isCoilAllowed).toBe(false);
      expect(p.unitLength).toBe(12);
      expect(p.unitCount).toBe(84); // ceil(1000 / 12) = 84
    });
  });
});
