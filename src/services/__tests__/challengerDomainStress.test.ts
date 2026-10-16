/**
 * Empirical Challenger Stress Test Harness for Milestone M2 (Requirement R2: PE Domain Capabilities)
 * 
 * Verifies:
 * 1. Reverse calculations against DIN 8074 standard weights (exhaustive & randomized across diameters, SDRs, and tonnages).
 * 2. Packaging unit counts, nominal length, and surplus meters across coil and branch configurations.
 * 3. Freight vehicle estimator recommendations across the full weight spectrum (100kg to 50,000kg) and branch lengths (0 to 12m).
 * 4. Raw material cost and IME pricing breakdown equations.
 */

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
  getWeightPerMeter,
} from '../calculator';
import { WeightTable, PackagingOptions } from '../../types';

function assert(condition: boolean, msg: string): asserts condition {
  if (!condition) {
    throw new Error(`[CHALLENGER FAILURE] ${msg}`);
  }
}

function assertCloseTo(actual: number, expected: number, tolerance = 0.05, msg = ''): void {
  const diff = Math.abs(actual - expected);
  if (diff > tolerance) {
    throw new Error(`[CHALLENGER FAILURE] Expected ${expected} ± ${tolerance}, but got ${actual}. ${msg}`);
  }
}

const weightTable = seedData.weightTable as unknown as WeightTable;

console.log('================================================================');
console.log('       EMPIRICAL CHALLENGER STRESS HARNESS — MILESTONE M2       ');
console.log('================================================================');

let totalAssertions = 0;

// ============================================================================
// PART 1: REVERSE CALCULATIONS STRESS HARNESS (DIN 8074 STANDARD WEIGHTS)
// ============================================================================
console.log('\n>>> SUITE 1: Reverse Calculations against DIN 8074 Standard Weights');

// 1.1 Extract all valid pairs of (diameter, SDR) in the DIN weight table
interface ValidPipeSpec {
  diameter: number;
  sdr: number;
  weightPerMeter: number;
}

const validSpecs: ValidPipeSpec[] = [];
const diameters = Object.keys(weightTable).map(Number).sort((a, b) => a - b);

for (const d of diameters) {
  const sdrObj = weightTable[d];
  if (!sdrObj) continue;
  for (const sdrStr of Object.keys(sdrObj)) {
    const sdr = Number(sdrStr);
    const wpm = sdrObj[sdr];
    if (typeof wpm === 'number' && wpm > 0) {
      validSpecs.push({ diameter: d, sdr, weightPerMeter: wpm });
    }
  }
}

console.log(`- Loaded ${validSpecs.length} distinct DIN 8074 (diameter, SDR) standard specifications.`);
assert(validSpecs.length > 50, `Expected > 50 valid DIN specifications, found ${validSpecs.length}`);
totalAssertions++;

// 1.2 Deterministic & Random Target Tonnages sweep
const deterministicTonnages = [0.05, 0.1, 0.25, 0.5, 1.0, 2.5, 3.14, 5.0, 10.0, 24.0, 50.0, 100.0];
let reverseChecksCount = 0;

for (const spec of validSpecs) {
  for (const ton of deterministicTonnages) {
    const massKg = ton * 1000;
    const len = calculateLengthFromMass(massKg, spec.diameter, spec.sdr, weightTable);
    assert(len !== null && len > 0, `Length for ${ton}t of D=${spec.diameter} SDR=${spec.sdr} must be positive`);
    
    // Theoretical oracle
    const expectedLen = Number((massKg / spec.weightPerMeter).toFixed(2));
    assert(len === expectedLen, `Expected ${expectedLen}m, got ${len}m for D=${spec.diameter} SDR=${spec.sdr}`);

    // Round-trip verification: mass reconstructed from length
    const reconstructedMass = len * spec.weightPerMeter;
    const massTolerance = 0.005 * spec.weightPerMeter + 0.05; // 0.005m rounding discretization
    assertCloseTo(reconstructedMass, massKg, massTolerance, `Mass roundtrip failure for ${ton}t`);

    totalAssertions += 3;
    reverseChecksCount++;
  }
}
console.log(`- Completed ${reverseChecksCount} deterministic reverse mass->length calculations across all DIN specs.`);

// 1.3 High-level calculateReversePipe in Tonnage, WeightKg, and Budget modes
let highLevelChecksCount = 0;
// Test 200 randomized scenarios
for (let i = 0; i < 200; i++) {
  const spec = validSpecs[Math.floor(Math.random() * validSpecs.length)];
  const randomTonnage = Number((0.1 + Math.random() * 40).toFixed(3));
  const pricePerKg = Math.round(50000 + Math.random() * 100000);

  // Mode: tonnage
  const resTonnage = calculateReversePipe(
    {
      mode: 'tonnage',
      diameter: spec.diameter,
      sdr: spec.sdr,
      peGrade: 'PE100',
      pricePerKg,
      targetTonnage: randomTonnage,
      packagingPreference: isCoilAllowed(spec.diameter) ? 'coil' : 'branch',
    },
    weightTable
  );

  assert(resTonnage !== null, `Reverse pipe tonnage calculation failed for D=${spec.diameter} SDR=${spec.sdr}`);
  assert(resTonnage.weightPerMeter === spec.weightPerMeter, `WPM mismatch in result`);
  assert(resTonnage.totalTonnage === Number((resTonnage.totalWeightKg / 1000).toFixed(3)), `Tonnage inconsistency`);
  assert(resTonnage.totalPrice === Math.round(resTonnage.totalWeightKg * pricePerKg), `Total price mismatch`);
  totalAssertions += 4;

  // Mode: budget
  const budget = Math.round(randomTonnage * 1000 * pricePerKg);
  const resBudget = calculateReversePipe(
    {
      mode: 'budget',
      diameter: spec.diameter,
      sdr: spec.sdr,
      peGrade: 'PE100',
      pricePerKg,
      targetBudget: budget,
      packagingPreference: 'branch',
      branchLengthPreference: 12,
    },
    weightTable
  );

  assert(resBudget !== null, `Reverse pipe budget calculation failed`);
  assert(resBudget.calculatedLength > 0, `Budget calculated length must be positive`);
  const maxPriceDiscretization = (0.006 * spec.weightPerMeter + 0.002) * pricePerKg + 10;
  assertCloseTo(resBudget.totalPrice, budget, maxPriceDiscretization, `Budget total price mismatch`);
  totalAssertions += 3;

  highLevelChecksCount++;
}
console.log(`- Completed ${highLevelChecksCount} randomized calculateReversePipe checks (tonnage & budget modes).`);

// 1.4 Invalid & Boundary Inputs for Reverse Calculation
assert(calculateLengthFromMass(0, 110, 11, weightTable) === null, '0 mass must return null');
assert(calculateLengthFromMass(-100, 110, 11, weightTable) === null, 'Negative mass must return null');
assert(calculateLengthFromMass(1000, 9999, 11, weightTable) === null, 'Non-existent diameter must return null');
assert(calculateLengthFromMass(1000, 110, 99, weightTable) === null, 'Non-existent SDR must return null');
assert(calculatePricePerKgFromTotal(0, 500) === null, '0 total price must return null');
assert(calculatePricePerKgFromTotal(500, 0) === null, '0 total weight must return null');
assert(calculatePricePerKgFromTotal(500, -10) === null, 'Negative total weight must return null');
assert(calculatePricePerKgFromTotal(-500, 10) === null, 'Negative total price must return null');

const nullRev1 = calculateReversePipe({ mode: 'tonnage', diameter: 110, sdr: 11, peGrade: 'PE100', pricePerKg: 95000, targetTonnage: 0 }, weightTable);
assert(nullRev1 === null, 'targetTonnage = 0 must return null');
const nullRev2 = calculateReversePipe({ mode: 'budget', diameter: 110, sdr: 11, peGrade: 'PE100', pricePerKg: 0, targetBudget: 1000000 }, weightTable);
assert(nullRev2 === null, 'pricePerKg = 0 in budget mode must return null');
totalAssertions += 10;
console.log('✓ Suite 1 PASSED: Reverse calculations rigorously verified against DIN standards.');

// ============================================================================
// PART 2: PACKAGING UNIT COUNTS AND SURPLUS METERS STRESS HARNESS
// ============================================================================
console.log('\n>>> SUITE 2: Packaging Rules, Unit Counts & Surplus Meters');

// 2.1 Bending Radius & Coiling Rule Verification
for (const d of diameters) {
  const allowed = isCoilAllowed(d);
  if (d <= 125) {
    assert(allowed === true, `Diameter ${d} <= 125 must allow coil`);
  } else {
    assert(allowed === false, `Diameter ${d} > 125 must strictly prohibit coil`);
  }
  totalAssertions++;
}

// 2.2 Forced Branch Conversion for Diameter > 125mm
const largeDiameters = diameters.filter((d) => d > 125);
for (const d of largeDiameters) {
  const pack = computePackaging(d, 50, { preferredType: 'coil', coilLength: 100 });
  assert(pack.packagingType === 'branch', `Diameter ${d} > 125 MUST override preferredType coil to branch`);
  assert(pack.isCoilAllowed === false, `isCoilAllowed must be false for ${d}`);
  assert(pack.unitLength === 12, `Default branch length for ${d} > 125 must be 12m`);
  totalAssertions += 3;
}

// 2.3 Comprehensive Sweep of Lengths and Unit Sizes
const testLengths = [0, 0.1, 1, 5.9, 6, 6.1, 11.9, 12, 12.1, 49.9, 50, 50.1, 99.9, 100, 100.1, 230, 500, 1000, 1000.5, 9999];
const coilSizes = [50, 100, 150, 200, 400];
const branchSizes: (6 | 12)[] = [6, 12];

let packagingChecksCount = 0;

// Test coils for diameter <= 125
for (const coilLen of coilSizes) {
  for (const len of testLengths) {
    const res = computePackaging(63, len, { preferredType: 'coil', coilLength: coilLen });
    assert(res.packagingType === 'coil', 'Must be coil');
    assert(res.unitLength === coilLen, 'Unit length must match coil size');
    
    const expectedCount = len > 0 ? Math.ceil(len / coilLen) : 0;
    assert(res.unitCount === expectedCount, `Expected unitCount ${expectedCount}, got ${res.unitCount} for len=${len}, coil=${coilLen}`);
    
    const expectedNominal = expectedCount * coilLen;
    assert(res.nominalTotalLength === expectedNominal, `Nominal total mismatch: expected ${expectedNominal}, got ${res.nominalTotalLength}`);
    
    const expectedSurplus = Number(Math.max(0, expectedNominal - len).toFixed(2));
    assert(res.wasteOrSurplusMeters === expectedSurplus, `Surplus mismatch: expected ${expectedSurplus}, got ${res.wasteOrSurplusMeters}`);
    
    if (len > 0) {
      assert(res.wasteOrSurplusMeters >= 0 && res.wasteOrSurplusMeters < coilLen, `Surplus ${res.wasteOrSurplusMeters} out of bounds [0, ${coilLen})`);
    } else {
      assert(res.wasteOrSurplusMeters === 0 && res.unitCount === 0, 'Zero length must have 0 units and 0 surplus');
    }
    
    assert(res.labelPersian.includes('کلاف'), `Label must include کلاف: ${res.labelPersian}`);
    totalAssertions += 7;
    packagingChecksCount++;
  }
}

// Test branches (6m and 12m)
for (const branchLen of branchSizes) {
  for (const len of testLengths) {
    const res = computePackaging(160, len, { branchLength: branchLen });
    assert(res.packagingType === 'branch', 'Must be branch');
    assert(res.unitLength === branchLen, 'Unit length must match branch size');
    
    const expectedCount = len > 0 ? Math.ceil(len / branchLen) : 0;
    assert(res.unitCount === expectedCount, `Expected unitCount ${expectedCount}, got ${res.unitCount} for len=${len}, branch=${branchLen}`);
    
    const expectedNominal = expectedCount * branchLen;
    assert(res.nominalTotalLength === expectedNominal, `Nominal total mismatch: expected ${expectedNominal}, got ${res.nominalTotalLength}`);
    
    const expectedSurplus = Number(Math.max(0, expectedNominal - len).toFixed(2));
    assert(res.wasteOrSurplusMeters === expectedSurplus, `Surplus mismatch for len=${len}`);
    
    if (len > 0) {
      assert(res.wasteOrSurplusMeters >= 0 && res.wasteOrSurplusMeters < branchLen, `Surplus ${res.wasteOrSurplusMeters} out of bounds [0, ${branchLen})`);
    } else {
      assert(res.wasteOrSurplusMeters === 0 && res.unitCount === 0, 'Zero length branch check');
    }
    
    assert(res.labelPersian.includes('شاخه'), `Label must include شاخه: ${res.labelPersian}`);
    totalAssertions += 7;
    packagingChecksCount++;
  }
}
console.log(`- Completed ${packagingChecksCount} packaging unit count & surplus meter tests across configurations.`);
console.log('✓ Suite 2 PASSED: Packaging calculations, unit counts, and surplus meters verified.');

// ============================================================================
// PART 3: FREIGHT VEHICLE ESTIMATOR STRESS HARNESS (100kg to 50,000kg, 0 to 12m)
// ============================================================================
console.log('\n>>> SUITE 3: Freight Vehicle Estimator across Weight Spectrum & Branch Lengths');

// 3.1 Verify 5 Iranian Freight Vehicle specifications
assert(FREIGHT_VEHICLES.length === 5, 'Must have exactly 5 freight vehicle tiers');
const [vNissan, vKhavar, vTruck6w, vTruck10w, vTrailer] = FREIGHT_VEHICLES;
assert(vNissan.id === 'nissan' && vNissan.maxPayloadKg === 2000 && vNissan.maxBranchLengthMeters === 2.2, 'Nissan spec');
assert(vKhavar.id === 'khavar' && vKhavar.maxPayloadKg === 4000 && vKhavar.maxBranchLengthMeters === 4.5, 'Khavar spec');
assert(vTruck6w.id === 'truck_6w' && vTruck6w.maxPayloadKg === 10000 && vTruck6w.maxBranchLengthMeters === 6.0, '6W truck spec');
assert(vTruck10w.id === 'truck_10w' && vTruck10w.maxPayloadKg === 15000 && vTruck10w.maxBranchLengthMeters === 6.0, '10W truck spec');
assert(vTrailer.id === 'trailer_flatbed' && vTrailer.maxPayloadKg === 24000 && vTrailer.maxBranchLengthMeters === 12.0, 'Trailer spec');
totalAssertions += 6;

// 3.2 Weight Spectrum & Branch Length Combinations
const testWeights: number[] = [];
// Step by 250kg from 100kg to 50,000kg
for (let w = 100; w <= 50000; w += 250) {
  testWeights.push(w);
}
// Include exact boundary transition points
const boundaryWeights = [
  1999, 2000, 2001,
  3999, 4000, 4001,
  9999, 10000, 10001,
  14999, 15000, 15001,
  23999, 24000, 24001,
  47999, 48000, 48001,
];
for (const bw of boundaryWeights) {
  if (!testWeights.includes(bw)) testWeights.push(bw);
}
testWeights.sort((a, b) => a - b);

const testBranchLengths = [0, 1.5, 2.2, 2.21, 3.0, 4.5, 6.0, 6.01, 8.0, 12.0];

let freightChecksCount = 0;

for (const w of testWeights) {
  for (const maxLen of testBranchLengths) {
    const rec = estimateFreightVehicle(w, maxLen);

    // Invariant 1: Total weight and tonnage consistency
    assert(rec.totalWeightKg === w, `Total weight must be ${w}`);
    assert(rec.totalTonnage === Number((w / 1000).toFixed(3)), `Tonnage inconsistency`);
    assert(rec.vehicleCount >= 1, `Vehicle count must be >= 1 for positive weight ${w}`);
    assert(rec.utilizationPercentage >= 0 && rec.utilizationPercentage <= 100, `Utilization ${rec.utilizationPercentage}% out of range`);

    // Invariant 2: Capacity must satisfy total weight
    const totalCapacity = rec.vehicleCount * rec.vehicle.maxPayloadKg;
    assert(totalCapacity >= w, `Insufficient capacity: ${totalCapacity}kg < ${w}kg`);

    // Constraint Rule 1: 12m branch (length > 6m) MUST mandate trailer_flatbed
    if (maxLen > 6.0) {
      assert(rec.vehicle.id === 'trailer_flatbed', `Length ${maxLen}m > 6m MUST mandate trailer_flatbed, got ${rec.vehicle.id}`);
      assert(rec.isLengthConstrained === true, '12m branch must flag isLengthConstrained = true');
      const expectedCount = Math.max(1, Math.ceil(w / 24000));
      assert(rec.vehicleCount === expectedCount, `Expected ${expectedCount} flatbeds for ${w}kg, got ${rec.vehicleCount}`);
    }

    // Constraint Rule 2: 6m branch (2.2m < length <= 6m) CANNOT be Nissan or Khavar
    else if (maxLen > 2.2) {
      assert(rec.vehicle.id !== 'nissan' && rec.vehicle.id !== 'khavar', `Length ${maxLen}m cannot use nissan or khavar! Got ${rec.vehicle.id}`);
      
      if (w <= 10000) {
        assert(rec.vehicle.id === 'truck_6w', `Expected truck_6w for ${w}kg with 6m branch, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, 'Expected 1 truck_6w');
        // If weight <= 4000kg, length forced this vehicle
        if (w <= 4000) {
          assert(rec.isLengthConstrained === true, `Length constraint should be true for ${w}kg with 6m branch`);
        }
      } else if (w <= 15000) {
        assert(rec.vehicle.id === 'truck_10w', `Expected truck_10w for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, 'Expected 1 truck_10w');
      } else if (w <= 24000) {
        assert(rec.vehicle.id === 'trailer_flatbed', `Expected trailer_flatbed for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, 'Expected 1 trailer');
      } else {
        assert(rec.vehicle.id === 'trailer_flatbed', `Expected trailer_flatbed for ${w}kg > 24t`);
        const expectedCount = Math.ceil(w / 24000);
        assert(rec.vehicleCount === expectedCount, `Expected ${expectedCount} trailers for ${w}kg`);
      }
    }

    // Constraint Rule 3: Pure weight tiers (coils / short fittings <= 2.2m)
    else {
      assert(rec.isLengthConstrained === false, 'isLengthConstrained should be false for <= 2.2m');
      if (w <= 2000) {
        assert(rec.vehicle.id === 'nissan', `Expected nissan for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, `Expected 1 nissan for ${w}kg`);
      } else if (w <= 4000) {
        assert(rec.vehicle.id === 'khavar', `Expected khavar for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, `Expected 1 khavar for ${w}kg`);
      } else if (w <= 10000) {
        assert(rec.vehicle.id === 'truck_6w', `Expected truck_6w for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, `Expected 1 truck_6w for ${w}kg`);
      } else if (w <= 15000) {
        assert(rec.vehicle.id === 'truck_10w', `Expected truck_10w for ${w}kg, got ${rec.vehicle.id}`);
        assert(rec.vehicleCount === 1, `Expected 1 truck_10w for ${w}kg`);
      } else {
        assert(rec.vehicle.id === 'trailer_flatbed', `Expected trailer_flatbed for ${w}kg, got ${rec.vehicle.id}`);
        const expectedCount = Math.ceil(w / 24000);
        assert(rec.vehicleCount === expectedCount, `Expected ${expectedCount} trailers for ${w}kg`);
      }
    }

    totalAssertions += 6;
    freightChecksCount++;
  }
}

// 3.3 Zero and Negative Weight Edge Cases
const recZero = estimateFreightVehicle(0, 0);
assert(recZero.vehicleCount === 0, 'Zero weight must have 0 vehicle count');
assert(recZero.utilizationPercentage === 0, 'Zero weight must have 0% utilization');
const recNeg = estimateFreightVehicle(-500, 6);
assert(recNeg.vehicleCount === 0, 'Negative weight must return 0 vehicle count');
totalAssertions += 3;

console.log(`- Completed ${freightChecksCount} freight estimator checks across the weight spectrum (100kg-50,000kg) and lengths (0-12m).`);
console.log('✓ Suite 3 PASSED: Freight vehicle estimator and dual constraints verified.');

// ============================================================================
// PART 4: RAW MATERIAL COST CALCULATOR STRESS HARNESS (IME PRICING FORMULA)
// ============================================================================
console.log('\n>>> SUITE 4: Raw Material Cost & IME Formula Stress Tests');

let costChecksCount = 0;
// Test parametric sweep of pricing components
const testPolymers = [50000, 85000, 120000];
const testMbPcts = [0, 1.5, 2.0, 2.5, 5.0];
const testWastePcts = [0, 1.0, 1.5, 2.0, 3.0];
const testConvCosts = [0, 8000, 12000, 20000];
const testMargins = [0, 5, 10, 15, 25];

for (const pPolymer of testPolymers) {
  for (const mbPct of testMbPcts) {
    for (const wastePct of testWastePcts) {
      for (const convCost of testConvCosts) {
        for (const margin of testMargins) {
          const mbPrice = pPolymer * 1.5;
          const res = calculateRawMaterialCost({
            polymerPricePerKg: pPolymer,
            masterbatchPercentage: mbPct,
            masterbatchPricePerKg: mbPrice,
            wastePercentage: wastePct,
            conversionCostPerKg: convCost,
            profitMarginPercentage: margin,
          });

          // Theoretical calculation
          const expectedBase = pPolymer * (1 - mbPct / 100) + mbPrice * (mbPct / 100);
          const expectedWaste = expectedBase * (wastePct / 100);
          const expectedAdj = expectedBase + expectedWaste;
          const expectedProd = expectedAdj + convCost;
          const expectedProfit = expectedProd * (margin / 100);
          const expectedSuggested = Math.round(expectedProd + expectedProfit);

          assert(res.baseMaterialCostPerKg === Math.round(expectedBase), `Base cost mismatch`);
          assert(res.wasteCostPerKg === Math.round(expectedWaste), `Waste cost mismatch`);
          assert(res.adjustedMaterialCostPerKg === Math.round(expectedAdj), `Adjusted cost mismatch`);
          assert(res.conversionCostPerKg === Math.round(convCost), `Conversion cost mismatch`);
          assert(res.productionCostPerKg === Math.round(expectedProd), `Production cost mismatch`);
          assert(res.profitAmountPerKg === Math.round(expectedProfit), `Profit mismatch`);
          assert(res.suggestedPricePerKg === expectedSuggested, `Suggested price mismatch`);

          totalAssertions += 7;
          costChecksCount++;
        }
      }
    }
  }
}
console.log(`- Completed ${costChecksCount} raw material IME formula parametric calculations.`);
console.log('✓ Suite 4 PASSED: Raw material cost calculator equations verified.');

// ============================================================================
// FINAL SUMMARY
// ============================================================================
console.log('\n================================================================');
console.log(`  ALL CHALLENGER EMPIRICAL TESTS PASSED! (${totalAssertions} assertions verified)  `);
console.log('  VERDICT: APPROVE                                              ');
console.log('================================================================\n');
