import { DiscountRule, InvoiceItem, PipeGrade, PipeItem, WeightTable } from '../types';

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
  },
  weightTable: WeightTable,
  existingId?: string
): PipeItem | null {
  const wpm = getWeightPerMeter(params.diameter, params.sdr, weightTable);
  if (wpm === null || wpm <= 0) return null;

  const totalWeight = Number((wpm * params.length).toFixed(3));
  const totalPrice = Math.round(totalWeight * params.pricePerKg);

  return {
    id: existingId || 'pipe_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
    type: 'pipe',
    peGrade: params.peGrade,
    diameter: params.diameter,
    sdr: params.sdr,
    pn: params.pn,
    length: Number(params.length.toFixed(2)),
    weightPerMeter: wpm,
    totalWeight,
    pricePerKg: params.pricePerKg,
    totalPrice,
  };
}

/**
 * Calculates pipe length from total mass
 */
export function calculateLengthFromMass(
  totalMass: number,
  diameter: number,
  sdr: number,
  weightTable: WeightTable
): number | null {
  const wpm = getWeightPerMeter(diameter, sdr, weightTable);
  if (!wpm || wpm <= 0) return null;
  return Number((totalMass / wpm).toFixed(2));
}

/**
 * Calculates unit price per kg from total price and total weight
 */
export function calculatePricePerKgFromTotal(
  totalPrice: number,
  totalWeight: number
): number | null {
  if (!totalWeight || totalWeight <= 0) return null;
  return Math.round(totalPrice / totalWeight);
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
 * Computes complete invoice financial summary
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
  const subtotal = params.items.reduce((sum, item) => sum + (item.totalPrice || 0), 0);

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

  const totalPipeWeight = params.items
    .filter((i): i is PipeItem => i.type === 'pipe')
    .reduce((sum, p) => sum + (p.totalWeight || 0), 0);

  const totalFittingCount = params.items
    .filter((i) => i.type === 'fitting')
    .reduce((sum, f) => sum + (f.quantity || 0), 0);

  return {
    subtotal,
    discountAmount,
    amountAfterDiscount,
    taxAmount,
    totalAmount,
    totalPipeWeight: Number(totalPipeWeight.toFixed(3)),
    totalFittingCount,
  };
}

