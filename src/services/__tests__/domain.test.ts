import { describe, it, expect } from 'vitest';
import {
  isCoilAllowed,
  computePackaging,
  estimateFreightVehicle,
  calculateRawMaterialCost,
  FREIGHT_VEHICLES,
} from '../calculator';

describe('PE Domain Capabilities Service - Packaging, Freight & Raw Materials', () => {
  describe('Packaging Constraint Rules (isCoilAllowed)', () => {
    it('Case 24: allows coil packaging for 63mm (diameter <= 125mm)', () => {
      expect(isCoilAllowed(63)).toBe(true);
    });

    it('Case 25: allows coil packaging for boundary 125mm (diameter <= 125mm)', () => {
      expect(isCoilAllowed(125)).toBe(true);
    });

    it('Case 26: strictly prohibits coil packaging for 160mm (>125mm must be branch only)', () => {
      expect(isCoilAllowed(160)).toBe(false);
      expect(isCoilAllowed(140)).toBe(false);
      expect(isCoilAllowed(200)).toBe(false);
      expect(isCoilAllowed(500)).toBe(false);
    });
  });

  describe('Packaging Unit & Length Calculations (computePackaging)', () => {
    it('Case 27: computes 500m of 63mm coil -> 5 coils of 100m', () => {
      const packaging = computePackaging(63, 500, {
        preferredType: 'coil',
        coilLength: 100,
      });

      expect(packaging.packagingType).toBe('coil');
      expect(packaging.isCoilAllowed).toBe(true);
      expect(packaging.unitCount).toBe(5);
      expect(packaging.unitLength).toBe(100);
      expect(packaging.nominalTotalLength).toBe(500);
      expect(packaging.actualLength).toBe(500);
      expect(packaging.wasteOrSurplusMeters).toBe(0);
      expect(packaging.labelPersian).toContain('۵ کلاف ۱۰۰ متری');
    });

    it('Case 28: computes 100m of 160mm branch -> 9 branches of 12m (108m nominal, 8m surplus)', () => {
      const packaging = computePackaging(160, 100, {
        branchLength: 12,
      });

      expect(packaging.packagingType).toBe('branch');
      expect(packaging.isCoilAllowed).toBe(false); // 160mm cannot be coiled
      expect(packaging.unitCount).toBe(9); // Math.ceil(100 / 12) = 9
      expect(packaging.unitLength).toBe(12);
      expect(packaging.nominalTotalLength).toBe(108); // 9 * 12 = 108
      expect(packaging.actualLength).toBe(100);
      expect(packaging.wasteOrSurplusMeters).toBe(8); // 108 - 100 = 8
      expect(packaging.labelPersian).toContain('۹ شاخه ۱۲ متری');
    });

    it('enforces branch packaging for diameter > 125 even if preferredType is set to coil', () => {
      const packaging = computePackaging(200, 50, {
        preferredType: 'coil',
        branchLength: 6,
      });

      expect(packaging.packagingType).toBe('branch');
      expect(packaging.isCoilAllowed).toBe(false);
      expect(packaging.unitLength).toBe(6);
      expect(packaging.unitCount).toBe(9); // ceil(50 / 6) = 9
    });
  });

  describe('Freight Vehicle Estimator & Logistics Constraints (estimateFreightVehicle)', () => {
    it('Case 29: recommends وانت نیسان (2t) for 1.5t payload with coil packaging', () => {
      const recommendation = estimateFreightVehicle(1500, 0);

      expect(recommendation.vehicle.id).toBe('nissan');
      expect(recommendation.vehicleCount).toBe(1);
      expect(recommendation.totalWeightKg).toBe(1500);
      expect(recommendation.totalTonnage).toBe(1.5);
      expect(recommendation.isLengthConstrained).toBe(false);
      expect(recommendation.vehicle.nameFa).toContain('وانت نیسان');
    });

    it('Case 30: recommends کامیونت خاور / ایسوزو (4t) for 3.2t payload', () => {
      const recommendation = estimateFreightVehicle(3200, 0);

      expect(recommendation.vehicle.id).toBe('khavar');
      expect(recommendation.vehicleCount).toBe(1);
      expect(recommendation.totalWeightKg).toBe(3200);
      expect(recommendation.totalTonnage).toBe(3.2);
      expect(recommendation.isLengthConstrained).toBe(false);
      expect(recommendation.vehicle.nameFa).toContain('خاور');
    });

    it('Case 31: recommends کامیون تک (10t) for 8t payload with 6m branch length', () => {
      const recommendation = estimateFreightVehicle(8000, 6);

      expect(recommendation.vehicle.id).toBe('truck_6w');
      expect(recommendation.vehicleCount).toBe(1);
      expect(recommendation.totalWeightKg).toBe(8000);
      expect(recommendation.totalTonnage).toBe(8.0);
      expect(recommendation.vehicle.nameFa).toContain('کامیون تک');
    });

    it('Case 32: recommends کامیون جفت (15t) for 13t payload with 6m branch length', () => {
      const recommendation = estimateFreightVehicle(13000, 6);

      expect(recommendation.vehicle.id).toBe('truck_10w');
      expect(recommendation.vehicleCount).toBe(1);
      expect(recommendation.totalWeightKg).toBe(13000);
      expect(recommendation.totalTonnage).toBe(13.0);
      expect(recommendation.vehicle.nameFa).toContain('کامیون جفت');
    });

    it('Case 33: mandates تریلی کفی (24t) for 4t payload with 12m branch length due to cargo bed length', () => {
      // 4t would fit on a Khavar by weight, but 12m branch length strictly mandates Flatbed Trailer
      const recommendation = estimateFreightVehicle(4000, 12);

      expect(recommendation.vehicle.id).toBe('trailer_flatbed');
      expect(recommendation.vehicleCount).toBe(1);
      expect(recommendation.totalWeightKg).toBe(4000);
      expect(recommendation.totalTonnage).toBe(4.0);
      expect(recommendation.isLengthConstrained).toBe(true);
      expect(recommendation.vehicle.nameFa).toContain('تریلی کفی');
      expect(recommendation.recommendationNote).toContain('شاخه ۱۲ متری');
    });

    it('calculates multiple flatbeds for tonnage exceeding single vehicle payload', () => {
      const recommendation = estimateFreightVehicle(50000, 12);

      expect(recommendation.vehicle.id).toBe('trailer_flatbed');
      expect(recommendation.vehicleCount).toBe(3); // ceil(50,000 / 24,000) = 3
      expect(recommendation.totalTonnage).toBe(50.0);
    });

    it('handles zero weight gracefully', () => {
      const recommendation = estimateFreightVehicle(0, 0);
      expect(recommendation.vehicleCount).toBe(0);
      expect(recommendation.totalWeightKg).toBe(0);
    });
  });

  describe('Raw Material Cost Formula (calculateRawMaterialCost)', () => {
    it('Case 34: accurately computes raw material cost breakdown and suggested selling price', () => {
      const result = calculateRawMaterialCost({
        polymerPricePerKg: 80000,
        masterbatchPercentage: 2,
        masterbatchPricePerKg: 120000,
        wastePercentage: 1.5,
        conversionCostPerKg: 10000,
        profitMarginPercentage: 10,
      });

      // 1. Base material cost: 80000 * 0.98 + 120000 * 0.02 = 78400 + 2400 = 80800
      expect(result.baseMaterialCostPerKg).toBe(80800);

      // 2. Waste cost: 80800 * 0.015 = 1212
      expect(result.wasteCostPerKg).toBe(1212);

      // 3. Adjusted material cost: 80800 + 1212 = 82012
      expect(result.adjustedMaterialCostPerKg).toBe(82012);

      // 4. Conversion cost: 10000
      expect(result.conversionCostPerKg).toBe(10000);

      // 5. Production cost: 82012 + 10000 = 92012
      expect(result.productionCostPerKg).toBe(92012);

      // 6. Profit amount: 92012 * 0.10 = 9201.2 -> 9201
      expect(result.profitAmountPerKg).toBe(9201);

      // 7. Suggested price per kg: round(92012 + 9201.2) = 101213
      expect(result.suggestedPricePerKg).toBe(101213);
    });

    it('computes pre-colored black polymer with 0% masterbatch and defaults', () => {
      const result = calculateRawMaterialCost({
        polymerPricePerKg: 85000,
        masterbatchPercentage: 0,
        wastePercentage: 1.0,
        conversionCostPerKg: 12000,
        profitMarginPercentage: 12,
      });

      expect(result.baseMaterialCostPerKg).toBe(85000);
      expect(result.wasteCostPerKg).toBe(850); // 85000 * 0.01 = 850
      expect(result.adjustedMaterialCostPerKg).toBe(85850);
      expect(result.conversionCostPerKg).toBe(12000);
      expect(result.productionCostPerKg).toBe(97850);
      expect(result.suggestedPricePerKg).toBe(Math.round(97850 * 1.12));
    });
  });
});
