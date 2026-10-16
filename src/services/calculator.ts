import {
  DiscountRule,
  FreightRecommendation,
  FreightVehicle,
  InvoiceItem,
  PackagingOptions,
  PackagingResult,
  PackagingType,
  PipeGrade,
  PipeItem,
  RawMaterialCostParams,
  RawMaterialCostResult,
  ReversePipeCalculationParams,
  ReversePipeCalculationResult,
  WeightTable,
} from '../types';
import { toPersianDigits } from './persianUtils';

/**
 * Look up mass per meter (kg/m) from the DIN standards weight table
 */
export function getWeightPerMeter(
  diameter: number,
  sdr: number,
  weightTable: WeightTable
): number | null {
  const dObj = weightTable[diameter];
  if (!dObj) return null;
  const weight = dObj[sdr];
  return typeof weight === 'number' && !isNaN(weight) ? weight : null;
}

/**
 * Packaging constraint rule:
 * Outer diameter <= 125mm allows Coil (کلاف) or Branch (شاخه).
 * Outer diameter > 125mm is strictly prohibited from coiling and MUST be Branch (شاخه 6m or 12m).
 */
export function isCoilAllowed(diameter: number): boolean {
  return diameter <= 125;
}

/**
 * Computes packaging specifications (coils or 6m/12m branches)
 */
export function computePackaging(
  diameter: number,
  length: number,
  options?: PackagingOptions
): PackagingResult {
  const coilAllowed = isCoilAllowed(diameter);
  const type: PackagingType = coilAllowed && options?.preferredType === 'coil' ? 'coil' : 'branch';

  if (type === 'coil') {
    const coilLen = options?.coilLength && options.coilLength > 0 ? options.coilLength : 100;
    const unitCount = length > 0 ? Math.ceil(length / coilLen) : 0;
    const nominalTotalLength = unitCount * coilLen;
    const wasteOrSurplusMeters = Number(Math.max(0, nominalTotalLength - length).toFixed(2));

    return {
      packagingType: 'coil',
      isCoilAllowed: true,
      unitCount,
      unitLength: coilLen,
      nominalTotalLength,
      actualLength: Number(length.toFixed(2)),
      wasteOrSurplusMeters,
      labelPersian: `${toPersianDigits(unitCount)} کلاف ${toPersianDigits(coilLen)} متری`,
    };
  } else {
    const defaultBranchLen: 6 | 12 = diameter > 125 ? 12 : 6;
    const branchLen = options?.branchLength === 6 || options?.branchLength === 12
      ? options.branchLength
      : defaultBranchLen;
    const unitCount = length > 0 ? Math.ceil(length / branchLen) : 0;
    const nominalTotalLength = unitCount * branchLen;
    const wasteOrSurplusMeters = Number(Math.max(0, nominalTotalLength - length).toFixed(2));

    return {
      packagingType: 'branch',
      isCoilAllowed: coilAllowed,
      unitCount,
      unitLength: branchLen,
      nominalTotalLength,
      actualLength: Number(length.toFixed(2)),
      wasteOrSurplusMeters,
      labelPersian: `${toPersianDigits(unitCount)} شاخه ${toPersianDigits(branchLen)} متری`,
    };
  }
}

/**
 * 5 Iranian Freight Vehicle Tiers
 */
export const FREIGHT_VEHICLES: FreightVehicle[] = [
  {
    id: 'nissan',
    nameFa: 'وانت نیسان (۲ تن)',
    nameEn: 'Nissan Pickup (2t)',
    maxPayloadKg: 2000,
    maxBranchLengthMeters: 2.2,
    descriptionFa: 'مناسب بارهای سبک، اتصالات و کلاف‌های تا ۲ تن',
  },
  {
    id: 'khavar',
    nameFa: 'کامیونت خاور / ایسوزو (۴ تن)',
    nameEn: 'Khavar / Isuzu Light Truck (4t)',
    maxPayloadKg: 4000,
    maxBranchLengthMeters: 4.5,
    descriptionFa: 'مناسب کلاف‌های متوسط و اتصالات تا ۴ تن (بدون شاخه ۶ یا ۱۲ متری)',
  },
  {
    id: 'truck_6w',
    nameFa: 'کامیون تک ۶ چرخ (۱۰ تن)',
    nameEn: '6-Wheel Single Truck (10t)',
    maxPayloadKg: 10000,
    maxBranchLengthMeters: 6.0,
    descriptionFa: 'مناسب حمل شاخه‌های ۶ متری و کلاف‌های سنگین تا ۱۰ تن',
  },
  {
    id: 'truck_10w',
    nameFa: 'کامیون جفت ۱۰ چرخ (۱۵ تن)',
    nameEn: '10-Wheel Double Truck (15t)',
    maxPayloadKg: 15000,
    maxBranchLengthMeters: 6.0,
    descriptionFa: 'مناسب پروژه‌های بزرگ با شاخه‌های ۶ متری تا ۱۵ تن',
  },
  {
    id: 'trailer_flatbed',
    nameFa: 'تریلی کفی ۱۸ چرخ (۲۴ تن)',
    nameEn: '18-Wheel Flatbed Trailer (24t)',
    maxPayloadKg: 24000,
    maxBranchLengthMeters: 12.0,
    descriptionFa: 'تنها وسیله مجاز برای حمل شاخه‌های ۱۲ متری و تناژ بالا تا ۲۴ تن',
  },
];

/**
 * Estimate logistics vehicle requirement with dual constraints:
 * 1. Payload weight capacity
 * 2. Cargo bed pipe length limits (12m branch mandates Flatbed Trailer; 6m branch mandates at least 6W truck)
 */
export function estimateFreightVehicle(
  totalWeightKg: number,
  maxBranchLengthMeters: number = 0
): FreightRecommendation {
  const totalTonnage = Number((totalWeightKg / 1000).toFixed(3));

  if (totalWeightKg <= 0) {
    const defaultVehicle = FREIGHT_VEHICLES[0];
    return {
      vehicle: defaultVehicle,
      vehicleCount: 0,
      totalWeightKg: 0,
      totalTonnage: 0,
      utilizationPercentage: 0,
      isLengthConstrained: false,
      recommendationNote: 'باری برای تخصیص ناوگان ثبت نشده است.',
    };
  }

  // Constraint 1: 12m branch strictly mandates Flatbed Trailer
  if (maxBranchLengthMeters > 6) {
    const flatbed = FREIGHT_VEHICLES.find((v) => v.id === 'trailer_flatbed')!;
    const count = Math.max(1, Math.ceil(totalWeightKg / flatbed.maxPayloadKg));
    const util = Math.min(100, Math.round((totalWeightKg / (count * flatbed.maxPayloadKg)) * 100));
    return {
      vehicle: flatbed,
      vehicleCount: count,
      totalWeightKg,
      totalTonnage,
      utilizationPercentage: util,
      isLengthConstrained: true,
      recommendationNote: `الزام به ${count} دستگاه تریلی کفی ۱۸ چرخ به دلیل وجود شاخه ۱۲ متری (ضریب بار ${util}٪).`,
    };
  }

  // Constraint 2: 6m branch mandates at least 6-wheel truck
  if (maxBranchLengthMeters > 2.2) {
    if (totalWeightKg <= 10000) {
      const truck6w = FREIGHT_VEHICLES.find((v) => v.id === 'truck_6w')!;
      const util = Math.min(100, Math.round((totalWeightKg / truck6w.maxPayloadKg) * 100));
      return {
        vehicle: truck6w,
        vehicleCount: 1,
        totalWeightKg,
        totalTonnage,
        utilizationPercentage: util,
        isLengthConstrained: totalWeightKg <= 4000,
        recommendationNote: `پیشنهاد ۱ دستگاه کامیون تک ۶ چرخ (۱۰ تن) مناسب شاخه ۶ متری (ضریب بار ${util}٪).`,
      };
    } else if (totalWeightKg <= 15000) {
      const truck10w = FREIGHT_VEHICLES.find((v) => v.id === 'truck_10w')!;
      const util = Math.min(100, Math.round((totalWeightKg / truck10w.maxPayloadKg) * 100));
      return {
        vehicle: truck10w,
        vehicleCount: 1,
        totalWeightKg,
        totalTonnage,
        utilizationPercentage: util,
        isLengthConstrained: false,
        recommendationNote: `پیشنهاد ۱ دستگاه کامیون جفت ۱۰ چرخ (۱۵ تن) مناسب شاخه ۶ متری (ضریب بار ${util}٪).`,
      };
    } else if (totalWeightKg <= 24000) {
      const flatbed = FREIGHT_VEHICLES.find((v) => v.id === 'trailer_flatbed')!;
      const util = Math.min(100, Math.round((totalWeightKg / flatbed.maxPayloadKg) * 100));
      return {
        vehicle: flatbed,
        vehicleCount: 1,
        totalWeightKg,
        totalTonnage,
        utilizationPercentage: util,
        isLengthConstrained: false,
        recommendationNote: `پیشنهاد ۱ دستگاه تریلی کفی ۱۸ چرخ (۲۴ تن) با ضریب بار ${util}٪.`,
      };
    } else {
      const flatbed = FREIGHT_VEHICLES.find((v) => v.id === 'trailer_flatbed')!;
      const count = Math.ceil(totalWeightKg / flatbed.maxPayloadKg);
      const util = Math.min(100, Math.round((totalWeightKg / (count * flatbed.maxPayloadKg)) * 100));
      return {
        vehicle: flatbed,
        vehicleCount: count,
        totalWeightKg,
        totalTonnage,
        utilizationPercentage: util,
        isLengthConstrained: false,
        recommendationNote: `پیشنهاد ${count} دستگاه تریلی کفی ۱۸ چرخ (۲۴ تن) با ضریب بار ${util}٪.`,
      };
    }
  }

  // Constraint 3: Coils or short fittings (weight tiers only)
  let chosen = FREIGHT_VEHICLES[0];
  if (totalWeightKg <= 2000) {
    chosen = FREIGHT_VEHICLES.find((v) => v.id === 'nissan')!;
  } else if (totalWeightKg <= 4000) {
    chosen = FREIGHT_VEHICLES.find((v) => v.id === 'khavar')!;
  } else if (totalWeightKg <= 10000) {
    chosen = FREIGHT_VEHICLES.find((v) => v.id === 'truck_6w')!;
  } else if (totalWeightKg <= 15000) {
    chosen = FREIGHT_VEHICLES.find((v) => v.id === 'truck_10w')!;
  } else {
    chosen = FREIGHT_VEHICLES.find((v) => v.id === 'trailer_flatbed')!;
  }

  const count = Math.max(1, Math.ceil(totalWeightKg / chosen.maxPayloadKg));
  const util = Math.min(100, Math.round((totalWeightKg / (count * chosen.maxPayloadKg)) * 100));

  return {
    vehicle: chosen,
    vehicleCount: count,
    totalWeightKg,
    totalTonnage,
    utilizationPercentage: util,
    isLengthConstrained: false,
    recommendationNote: `پیشنهاد ${count} دستگاه ${chosen.nameFa} با ضریب بار ${util}٪.`,
  };
}

/**
 * Calculates full pipe item parameters given length, diameter, SDR, and price per kg
 */
export function computePipeItem(
  params: {
    peGrade: PipeGrade;
    diameter: number;
    sdr: number;
    pn: number;
    length: number;
    pricePerKg: number;
    packagingOptions?: PackagingOptions;
  },
  weightTable: WeightTable,
  existingId?: string
): PipeItem | null {
  const wpm = getWeightPerMeter(params.diameter, params.sdr, weightTable);
  if (wpm === null || wpm <= 0) return null;

  const safeLength = Math.max(0, params.length);
  const safePricePerKg = Math.max(0, params.pricePerKg);

  const totalWeight = Number((wpm * safeLength).toFixed(3));
  const totalPrice = Math.round(totalWeight * safePricePerKg);
  const packaging = computePackaging(params.diameter, safeLength, params.packagingOptions);

  return {
    id: existingId || 'pipe_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
    type: 'pipe',
    peGrade: params.peGrade,
    diameter: params.diameter,
    sdr: params.sdr,
    pn: params.pn,
    length: Number(safeLength.toFixed(2)),
    weightPerMeter: wpm,
    totalWeight,
    pricePerKg: safePricePerKg,
    totalPrice,
    packaging,
  };
}

/**
 * Calculates pipe length from total mass (in kilograms)
 */
export function calculateLengthFromMass(
  totalMass: number,
  diameter: number,
  sdr: number,
  weightTable: WeightTable
): number | null {
  const wpm = getWeightPerMeter(diameter, sdr, weightTable);
  if (!wpm || wpm <= 0 || totalMass <= 0) return null;
  return Number((totalMass / wpm).toFixed(2));
}

/**
 * Calculates unit price per kg from total price and total weight
 */
export function calculatePricePerKgFromTotal(
  totalPrice: number,
  totalWeight: number
): number | null {
  if (!totalWeight || totalWeight <= 0 || totalPrice <= 0) return null;
  return Math.round(totalPrice / totalWeight);
}

/**
 * High-level reverse pipe calculation supporting both mass/tonnage and total budget modes
 */
export function calculateReversePipe(
  params: ReversePipeCalculationParams,
  weightTable: WeightTable
): ReversePipeCalculationResult | null {
  const wpm = getWeightPerMeter(params.diameter, params.sdr, weightTable);
  if (!wpm || wpm <= 0) return null;

  let massKg = 0;
  if (params.mode === 'budget') {
    if (!params.targetBudget || params.targetBudget <= 0 || !params.pricePerKg || params.pricePerKg <= 0) {
      return null;
    }
    massKg = params.targetBudget / params.pricePerKg;
  } else {
    // Mode tonnage or weight
    if (params.targetTonnage && params.targetTonnage > 0) {
      massKg = params.targetTonnage * 1000;
    } else if (params.targetWeightKg && params.targetWeightKg > 0) {
      massKg = params.targetWeightKg;
    } else {
      return null;
    }
  }

  const calculatedLength = Number((massKg / wpm).toFixed(2));
  if (calculatedLength <= 0) return null;

  const totalWeightKg = Number((calculatedLength * wpm).toFixed(3));
  const totalTonnage = Number((totalWeightKg / 1000).toFixed(3));
  const totalPrice = Math.round(totalWeightKg * params.pricePerKg);

  const packaging = computePackaging(params.diameter, calculatedLength, {
    preferredType: params.packagingPreference,
    branchLength: params.branchLengthPreference,
    coilLength: params.coilLengthPreference,
  });

  return {
    weightPerMeter: wpm,
    calculatedLength,
    totalWeightKg,
    totalTonnage,
    totalPrice,
    unitPricePerKg: params.pricePerKg,
    packaging,
  };
}

/**
 * Calculate raw material and manufacturing production cost indexed against Iran Commodity Exchange (بورس کالا)
 */
export function calculateRawMaterialCost(params: RawMaterialCostParams): RawMaterialCostResult {
  const mbPct = Math.max(0, (params.masterbatchPercentage ?? 0) / 100);
  const mbPrice = params.masterbatchPricePerKg ?? params.polymerPricePerKg;
  const baseMaterialCost = params.polymerPricePerKg * (1 - mbPct) + mbPrice * mbPct;

  const wastePct = Math.max(0, (params.wastePercentage ?? 0) / 100);
  const wasteCost = baseMaterialCost * wastePct;
  const adjustedMaterialCost = baseMaterialCost + wasteCost;

  const conversionCost = params.conversionCostPerKg ?? 0;
  const productionCost = adjustedMaterialCost + conversionCost;

  const marginPct = Math.max(0, (params.profitMarginPercentage ?? 0) / 100);
  const profitAmount = productionCost * marginPct;
  const suggestedPrice = Math.round(productionCost + profitAmount);

  return {
    baseMaterialCostPerKg: Math.round(baseMaterialCost),
    wasteCostPerKg: Math.round(wasteCost),
    adjustedMaterialCostPerKg: Math.round(adjustedMaterialCost),
    conversionCostPerKg: Math.round(conversionCost),
    productionCostPerKg: Math.round(productionCost),
    profitAmountPerKg: Math.round(profitAmount),
    suggestedPricePerKg: suggestedPrice,
  };
}

/**
 * Returns standard discount percentage based on order total
 */
export function getApplicableDiscount(
  orderTotal: number,
  discountRules: DiscountRule[]
): number {
  if (!discountRules || discountRules.length === 0) return 0;
  const sorted = [...discountRules].sort((a, b) => a.threshold - b.threshold);
  let discount = 0;
  for (const rule of sorted) {
    if (orderTotal >= rule.threshold) {
      discount = rule.percentage;
    } else {
      break;
    }
  }
  return discount;
}

/**
 * Computes complete invoice financial summary including freight recommendation
 */
export function computeInvoiceSummary(params: {
  items: InvoiceItem[];
  includeDiscount: boolean;
  discountPercentage: number;
  includeCustomDiscount: boolean;
  customDiscountAmount: number;
  includeTax: boolean;
  taxPercentage: number;
}) {
  const validItems = Array.isArray(params.items)
    ? params.items.filter((i): i is InvoiceItem => Boolean(i && typeof i === 'object'))
    : [];
  const subtotal = validItems.reduce((sum, item) => sum + (item.totalPrice || 0), 0);

  let discountAmount = 0;
  if (params.includeCustomDiscount && params.customDiscountAmount > 0) {
    discountAmount = params.customDiscountAmount;
  } else if (params.includeDiscount && params.discountPercentage > 0) {
    discountAmount = Math.round((subtotal * params.discountPercentage) / 100);
  }

  // Ensure discount does not exceed subtotal
  discountAmount = Math.min(discountAmount, subtotal);

  const amountAfterDiscount = Math.max(0, subtotal - discountAmount);

  let taxAmount = 0;
  if (params.includeTax && params.taxPercentage > 0) {
    taxAmount = Math.round((amountAfterDiscount * params.taxPercentage) / 100);
  }

  const totalAmount = amountAfterDiscount + taxAmount;

  const totalPipeWeight = validItems
    .filter((i): i is PipeItem => i.type === 'pipe')
    .reduce((sum, p) => sum + (p.totalWeight || 0), 0);

  const totalFittingCount = validItems
    .filter((i) => i.type === 'fitting')
    .reduce((sum, f) => sum + (f.quantity || 0), 0);

  // Compute maximum branch length to enforce freight bed length constraints
  let maxBranchLengthMeters = 0;
  for (const item of validItems) {
    if (item.type === 'pipe') {
      if (item.packaging?.packagingType === 'branch') {
        maxBranchLengthMeters = Math.max(maxBranchLengthMeters, item.packaging.unitLength);
      } else if (item.diameter > 125) {
        maxBranchLengthMeters = Math.max(maxBranchLengthMeters, 12);
      }
    }
  }

  const freight = estimateFreightVehicle(totalPipeWeight, maxBranchLengthMeters);

  return {
    subtotal,
    discountAmount,
    amountAfterDiscount,
    taxAmount,
    totalAmount,
    totalPipeWeight: Number(totalPipeWeight.toFixed(3)),
    totalFittingCount,
    maxBranchLengthMeters,
    freight,
  };
}

