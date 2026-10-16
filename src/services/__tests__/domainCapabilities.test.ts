import seedData from '../../data/seedData.json';
import {
  isCoilAllowed,
  computePackaging,
  FREIGHT_VEHICLES,
  estimateFreightVehicle,
  calculateLengthFromMass,
  calculatePricePerKgFromTotal,
  calculateReversePipe,
  calculateRawMaterialCost,
  computePipeItem,
  computeInvoiceSummary,
} from '../calculator';
import { WeightTable, PipeItem } from '../../types';

function assert(condition: boolean, message: string): asserts condition {
  if (!condition) {
    throw new Error(`Assertion Failed: ${message}`);
  }
}

function assertEqual<T>(actual: T, expected: T, message?: string) {
  if (actual !== expected) {
    throw new Error(`Assertion Failed: expected ${JSON.stringify(expected)} but got ${JSON.stringify(actual)}. ${message || ''}`);
  }
}

const weightTable = seedData.weightTable as unknown as WeightTable;

console.log('=== RUNNING PE DOMAIN CAPABILITIES TEST SUITE ===');

// ---------------------------------------------------------------------------
// 1. Packaging Rules: isCoilAllowed and computePackaging
// ---------------------------------------------------------------------------
console.log('\n--- Test 1: Packaging Rules (isCoilAllowed & computePackaging) ---');

// Diameters <= 125mm must allow coil
const coilAllowedSizes = [10, 12, 16, 20, 25, 32, 40, 50, 63, 75, 90, 110, 125];
for (const d of coilAllowedSizes) {
  assertEqual(isCoilAllowed(d), true, `Diameter ${d}mm should allow coil`);
}

// Diameters > 125mm must strictly prohibit coil
const branchOnlySizes = [140, 160, 180, 200, 225, 250, 280, 315, 355, 400, 450, 500];
for (const d of branchOnlySizes) {
  assertEqual(isCoilAllowed(d), false, `Diameter ${d}mm should prohibit coil`);
}

// Compute packaging: 63mm with coil
const packCoil = computePackaging(63, 500, { preferredType: 'coil', coilLength: 100 });
assertEqual(packCoil.packagingType, 'coil');
assertEqual(packCoil.unitCount, 5);
assertEqual(packCoil.unitLength, 100);
assertEqual(packCoil.nominalTotalLength, 500);
assertEqual(packCoil.wasteOrSurplusMeters, 0);
assertEqual(packCoil.labelPersian, '۵ کلاف ۱۰۰ متری');

// Compute packaging: 63mm coil with surplus
const packCoilSurplus = computePackaging(63, 230, { preferredType: 'coil', coilLength: 100 });
assertEqual(packCoilSurplus.unitCount, 3);
assertEqual(packCoilSurplus.nominalTotalLength, 300);
assertEqual(packCoilSurplus.wasteOrSurplusMeters, 70);
assertEqual(packCoilSurplus.labelPersian, '۳ کلاف ۱۰۰ متری');

// Compute packaging: 160mm forced to branch even when user requested coil
const packForcedBranch = computePackaging(160, 100, { preferredType: 'coil', branchLength: 12 });
assertEqual(packForcedBranch.packagingType, 'branch');
assertEqual(packForcedBranch.isCoilAllowed, false);
assertEqual(packForcedBranch.unitLength, 12);
assertEqual(packForcedBranch.unitCount, 9); // ceil(100/12) = 9
assertEqual(packForcedBranch.nominalTotalLength, 108);
assertEqual(packForcedBranch.wasteOrSurplusMeters, 8);
assertEqual(packForcedBranch.labelPersian, '۹ شاخه ۱۲ متری');

// Compute packaging: 110mm with branch option (6m)
const packBranch6m = computePackaging(110, 100, { preferredType: 'branch', branchLength: 6 });
assertEqual(packBranch6m.packagingType, 'branch');
assertEqual(packBranch6m.isCoilAllowed, true);
assertEqual(packBranch6m.unitLength, 6);
assertEqual(packBranch6m.unitCount, 17); // ceil(100/6) = 17
assertEqual(packBranch6m.nominalTotalLength, 102);
assertEqual(packBranch6m.wasteOrSurplusMeters, 2);
assertEqual(packBranch6m.labelPersian, '۱۷ شاخه ۶ متری');

console.log('✓ Test 1 Passed: Packaging rules and calculations verified.');

// ---------------------------------------------------------------------------
// 2. Freight Vehicle Estimator: 5 Iranian Logistics Tiers & Dual Constraints
// ---------------------------------------------------------------------------
console.log('\n--- Test 2: Freight Vehicle Estimator ---');

// Verify 5 vehicle tiers definition
assertEqual(FREIGHT_VEHICLES.length, 5);
const expectedVehicleIds = ['nissan', 'khavar', 'truck_6w', 'truck_10w', 'trailer_flatbed'];
for (let i = 0; i < 5; i++) {
  assertEqual(FREIGHT_VEHICLES[i].id, expectedVehicleIds[i]);
}

// Tier 1: Nissan (weight <= 2000kg, coils)
const recNissan = estimateFreightVehicle(1500, 0);
assertEqual(recNissan.vehicle.id, 'nissan');
assertEqual(recNissan.vehicleCount, 1);
assertEqual(recNissan.utilizationPercentage, 75);
assertEqual(recNissan.isLengthConstrained, false);

// Tier 2: Khavar (2000kg < weight <= 4000kg, coils)
const recKhavar = estimateFreightVehicle(3600, 0);
assertEqual(recKhavar.vehicle.id, 'khavar');
assertEqual(recKhavar.vehicleCount, 1);
assertEqual(recKhavar.utilizationPercentage, 90);
assertEqual(recKhavar.isLengthConstrained, false);

// Tier 3: 6-Wheel Truck (4000kg < weight <= 10000kg)
const recTruck6w = estimateFreightVehicle(8000, 6);
assertEqual(recTruck6w.vehicle.id, 'truck_6w');
assertEqual(recTruck6w.vehicleCount, 1);
assertEqual(recTruck6w.utilizationPercentage, 80);

// Tier 4: 10-Wheel Truck (10000kg < weight <= 15000kg)
const recTruck10w = estimateFreightVehicle(12000, 6);
assertEqual(recTruck10w.vehicle.id, 'truck_10w');
assertEqual(recTruck10w.vehicleCount, 1);
assertEqual(recTruck10w.utilizationPercentage, 80);

// Tier 5: Flatbed Trailer (15000kg < weight <= 24000kg)
const recTrailer = estimateFreightVehicle(20000, 6);
assertEqual(recTrailer.vehicle.id, 'trailer_flatbed');
assertEqual(recTrailer.vehicleCount, 1);
assertEqual(recTrailer.utilizationPercentage, 83);

// Multi-Trailer: weight > 24000kg
const recMultiTrailer = estimateFreightVehicle(36000, 0);
assertEqual(recMultiTrailer.vehicle.id, 'trailer_flatbed');
assertEqual(recMultiTrailer.vehicleCount, 2);
assertEqual(recMultiTrailer.utilizationPercentage, 75);

// Dual Constraint 1: 12m branch strictly mandates Flatbed Trailer even if weight is low (e.g. 3000kg)
const recLength12m = estimateFreightVehicle(3000, 12);
assertEqual(recLength12m.vehicle.id, 'trailer_flatbed');
assertEqual(recLength12m.isLengthConstrained, true);
assert(recLength12m.recommendationNote.includes('شاخه ۱۲ متری'), 'Should mention 12m branch');

// Dual Constraint 2: 6m branch mandates at least 6-wheel truck even if weight is low (e.g. 1500kg)
const recLength6m = estimateFreightVehicle(1500, 6);
assertEqual(recLength6m.vehicle.id, 'truck_6w');
assertEqual(recLength6m.isLengthConstrained, true);
assert(recLength6m.recommendationNote.includes('شاخه ۶ متری'), 'Should mention 6m branch');

// Edge Case: 0kg or empty order
const recZero = estimateFreightVehicle(0, 0);
assertEqual(recZero.vehicleCount, 0);
assertEqual(recZero.utilizationPercentage, 0);

console.log('✓ Test 2 Passed: Freight vehicle estimator and dual constraints verified.');

// ---------------------------------------------------------------------------
// 3. Raw Material Cost Calculator (بورس کالا و هزینه تبدیل)
// ---------------------------------------------------------------------------
console.log('\n--- Test 3: Raw Material Cost Calculator ---');

// Natural polymer scenario: 85,000 polymer, 2% MB @ 130,000, 1.5% waste, 10,000 conv, 10% margin
const costResult = calculateRawMaterialCost({
  polymerPricePerKg: 85000,
  masterbatchPercentage: 2.0,
  masterbatchPricePerKg: 130000,
  wastePercentage: 1.5,
  conversionCostPerKg: 10000,
  profitMarginPercentage: 10,
});

assertEqual(costResult.baseMaterialCostPerKg, 85900); // 85000*0.98 + 130000*0.02
assertEqual(costResult.wasteCostPerKg, 1289); // Math.round(85900 * 0.015)
assertEqual(costResult.adjustedMaterialCostPerKg, 87189);
assertEqual(costResult.conversionCostPerKg, 10000);
assertEqual(costResult.productionCostPerKg, 97189);
assertEqual(costResult.profitAmountPerKg, 9719); // Math.round(97188.5 * 0.10)
assertEqual(costResult.suggestedPricePerKg, 106907); // Math.round(97188.5 + 9718.85)

// Pre-colored polymer scenario (0% MB, 0% waste, 0% profit margin)
const simpleCost = calculateRawMaterialCost({
  polymerPricePerKg: 90000,
  conversionCostPerKg: 8000,
  profitMarginPercentage: 0,
});
assertEqual(simpleCost.baseMaterialCostPerKg, 90000);
assertEqual(simpleCost.wasteCostPerKg, 0);
assertEqual(simpleCost.productionCostPerKg, 98000);
assertEqual(simpleCost.suggestedPricePerKg, 98000);

console.log('✓ Test 3 Passed: Raw material cost formulas verified.');

// ---------------------------------------------------------------------------
// 4. Reverse Pipe Calculations (Mass -> Length, Total -> Price/Kg, Budget -> Length)
// ---------------------------------------------------------------------------
console.log('\n--- Test 4: Reverse Pipe Calculations ---');

// 110mm SDR11 DIN weight = 3.14 kg/m
assertEqual(weightTable[110][11], 3.14);

// calculateLengthFromMass: 314 kg -> 100.00 m
const lenFromMass = calculateLengthFromMass(314, 110, 11, weightTable);
assertEqual(lenFromMass, 100.00);

// calculatePricePerKgFromTotal: 29,830,000 / 314 kg -> 95,000 Toman/kg
const pricePerKg = calculatePricePerKgFromTotal(29830000, 314);
assertEqual(pricePerKg, 95000);

// High-level calculateReversePipe in tonnage mode: 3.14 tons -> 1000m
const revTonnage = calculateReversePipe(
  {
    mode: 'tonnage',
    diameter: 110,
    sdr: 11,
    peGrade: 'PE100',
    pricePerKg: 95000,
    targetTonnage: 3.14,
    packagingPreference: 'coil',
    coilLengthPreference: 100,
  },
  weightTable
);
assert(revTonnage !== null, 'revTonnage should not be null');
assertEqual(revTonnage.weightPerMeter, 3.14);
assertEqual(revTonnage.calculatedLength, 1000.00);
assertEqual(revTonnage.totalWeightKg, 3140);
assertEqual(revTonnage.totalTonnage, 3.14);
assertEqual(revTonnage.totalPrice, 298300000);
assertEqual(revTonnage.packaging.packagingType, 'coil');
assertEqual(revTonnage.packaging.unitCount, 10);
assertEqual(revTonnage.packaging.labelPersian, '۱۰ کلاف ۱۰۰ متری');

// High-level calculateReversePipe in budget mode: 29,830,000 Toman budget -> 100m
const revBudget = calculateReversePipe(
  {
    mode: 'budget',
    diameter: 110,
    sdr: 11,
    peGrade: 'PE100',
    pricePerKg: 95000,
    targetBudget: 29830000,
    packagingPreference: 'branch',
    branchLengthPreference: 6,
  },
  weightTable
);
assert(revBudget !== null, 'revBudget should not be null');
assertEqual(revBudget.calculatedLength, 100.00);
assertEqual(revBudget.totalPrice, 29830000);
assertEqual(revBudget.packaging.packagingType, 'branch');
assertEqual(revBudget.packaging.unitLength, 6);
assertEqual(revBudget.packaging.unitCount, 17);

// Reverse calculation error safety on invalid diameter
const revInvalid = calculateReversePipe(
  {
    mode: 'tonnage',
    diameter: 999,
    sdr: 11,
    peGrade: 'PE100',
    pricePerKg: 95000,
    targetTonnage: 5,
  },
  weightTable
);
assertEqual(revInvalid, null);

console.log('✓ Test 4 Passed: Reverse pipe calculation functions verified.');

// ---------------------------------------------------------------------------
// 5. computePipeItem with Packaging & computeInvoiceSummary with Freight
// ---------------------------------------------------------------------------
console.log('\n--- Test 5: Pipe Item Packaging & Invoice Summary Freight ---');

// Compute pipe item attaching packaging
const pipeItem1 = computePipeItem(
  {
    peGrade: 'PE100',
    diameter: 160,
    sdr: 11,
    pn: 16,
    length: 120,
    pricePerKg: 95000,
    packagingOptions: { branchLength: 12 },
  },
  weightTable
);
assert(pipeItem1 !== null, 'pipeItem1 should not be null');
const validPipeItem: PipeItem = pipeItem1;
assertEqual(validPipeItem.packaging?.packagingType, 'branch');
assertEqual(validPipeItem.packaging?.unitLength, 12);
assertEqual(validPipeItem.packaging?.unitCount, 10);
assertEqual(validPipeItem.packaging?.labelPersian, '۱۰ شاخه ۱۲ متری');

// Invoice Summary with pipe item 1 (120m of 160mm @ 6.67 kg/m = 800.4 kg, but 12m branch!)
const summary = computeInvoiceSummary({
  items: [validPipeItem],
  includeDiscount: false,
  discountPercentage: 0,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 10,
});

assertEqual(summary.maxBranchLengthMeters, 12);
assertEqual(summary.freight.vehicle.id, 'trailer_flatbed');
assertEqual(summary.freight.isLengthConstrained, true);
assertEqual(summary.freight.vehicleCount, 1);

console.log('✓ Test 5 Passed: Invoice summary freight and pipe item packaging verified.');

// ---------------------------------------------------------------------------
// 6. Adversarial Boundary, Edge Cases & Multi-Pipe Mixed Shipments
// ---------------------------------------------------------------------------
console.log('\n--- Test 6: Boundary & Multi-Pipe Mixed Logistics Stress Tests ---');

// Boundary test: Diameter 125mm is the exact maximum allowed for coil
assertEqual(isCoilAllowed(125), true, '125mm must allow coil');
assertEqual(isCoilAllowed(126), false, '126mm must prohibit coil');
assertEqual(isCoilAllowed(140), false, '140mm must prohibit coil');

// High tonnage stress test: 100 tons (100,000 kg) requires ceil(100,000 / 24,000) = 5 trailers
const rec100t = estimateFreightVehicle(100000, 0);
assertEqual(rec100t.vehicle.id, 'trailer_flatbed');
assertEqual(rec100t.vehicleCount, 5);
assertEqual(rec100t.utilizationPercentage, 83); // 100000 / (5 * 24000) = 83.33%

// Mixed shipment test:
// Pipe 1: 500m of 63mm coil (totalWeight = 525 kg, coil)
// Pipe 2: 120m of 110mm with 6m branch (totalWeight = 376.8 kg, branch 6m)
// Pipe 3: 60m of 200mm (diameter > 125mm, default 12m branch, totalWeight = 423 kg)
const mixedPipes: PipeItem[] = [
  computePipeItem(
    { peGrade: 'PE100', diameter: 63, sdr: 11, pn: 16, length: 500, pricePerKg: 95000, packagingOptions: { preferredType: 'coil' } },
    weightTable
  )!,
  computePipeItem(
    { peGrade: 'PE100', diameter: 110, sdr: 11, pn: 16, length: 120, pricePerKg: 95000, packagingOptions: { preferredType: 'branch', branchLength: 6 } },
    weightTable
  )!,
  computePipeItem(
    { peGrade: 'PE100', diameter: 200, sdr: 17, pn: 10, length: 60, pricePerKg: 95000, packagingOptions: { branchLength: 12 } },
    weightTable
  )!,
];

const mixedSummary = computeInvoiceSummary({
  items: mixedPipes,
  includeDiscount: false,
  discountPercentage: 0,
  includeCustomDiscount: false,
  customDiscountAmount: 0,
  includeTax: false,
  taxPercentage: 10,
});

assertEqual(mixedSummary.maxBranchLengthMeters, 12, 'Mixed order should detect max branch length of 12m');
assertEqual(mixedSummary.freight.vehicle.id, 'trailer_flatbed', 'Presence of 12m branch mandates flatbed trailer');
assertEqual(mixedSummary.freight.isLengthConstrained, true);

// Negative and zero values edge tests
assertEqual(calculateLengthFromMass(0, 110, 11, weightTable), null);
assertEqual(calculateLengthFromMass(-50, 110, 11, weightTable), null);
assertEqual(calculatePricePerKgFromTotal(0, 500), null);
assertEqual(calculatePricePerKgFromTotal(500000, 0), null);
assertEqual(calculatePricePerKgFromTotal(500000, -10), null);

// Packaging edge cases: length 0
const packZero = computePackaging(63, 0);
assertEqual(packZero.unitCount, 0);
assertEqual(packZero.wasteOrSurplusMeters, 0);

console.log('✓ Test 6 Passed: Boundary conditions and mixed freight shipments verified.');

console.log('\n======================================================');
console.log('ALL PE DOMAIN CAPABILITIES TESTS PASSED SUCCESSFULLY! (6/6)');
console.log('======================================================\n');
