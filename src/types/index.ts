export type PipeGrade = 'PE100' | 'PE80' | 'PE63';

// Packaging Types
export type PackagingType = 'coil' | 'branch';

export interface PackagingOptions {
  preferredType?: PackagingType;
  branchLength?: 6 | 12;
  coilLength?: number;
}

export interface PackagingResult {
  packagingType: PackagingType;
  isCoilAllowed: boolean;
  unitCount: number;
  unitLength: number;
  nominalTotalLength: number;
  actualLength: number;
  wasteOrSurplusMeters: number;
  labelPersian: string;
}

// Freight Types
export interface FreightVehicle {
  id: string;
  nameFa: string;
  nameEn: string;
  maxPayloadKg: number;
  maxBranchLengthMeters: number;
  descriptionFa: string;
}

export interface FreightRecommendation {
  vehicle: FreightVehicle;
  vehicleCount: number;
  totalWeightKg: number;
  totalTonnage: number;
  utilizationPercentage: number;
  isLengthConstrained: boolean;
  recommendationNote: string;
}

// Raw Material Cost Calculation Types
export interface RawMaterialCostParams {
  polymerPricePerKg: number; // Base polymer price from IME (بورس کالا)
  masterbatchPercentage?: number; // Carbon black masterbatch % (e.g. 2.0%)
  masterbatchPricePerKg?: number; // Masterbatch price per kg
  wastePercentage?: number; // Extrusion waste % (e.g. 1.5%)
  conversionCostPerKg?: number; // Factory overhead & conversion per kg (e.g. 10,000)
  profitMarginPercentage?: number; // Profit margin % (e.g. 10%)
}

export interface RawMaterialCostResult {
  baseMaterialCostPerKg: number;
  wasteCostPerKg: number;
  adjustedMaterialCostPerKg: number;
  conversionCostPerKg: number;
  productionCostPerKg: number;
  profitAmountPerKg: number;
  suggestedPricePerKg: number;
}

// Reverse Pipe Calculation Types
export interface ReversePipeCalculationParams {
  mode: 'tonnage' | 'budget';
  diameter: number;
  sdr: number;
  peGrade: PipeGrade;
  pricePerKg: number;
  targetTonnage?: number; // In tons (e.g. 2.5)
  targetWeightKg?: number; // In kg (e.g. 2500)
  targetBudget?: number; // In active currency
  packagingPreference?: PackagingType;
  branchLengthPreference?: 6 | 12;
  coilLengthPreference?: number;
}

export interface ReversePipeCalculationResult {
  weightPerMeter: number;
  calculatedLength: number;
  totalWeightKg: number;
  totalTonnage: number;
  totalPrice: number;
  unitPricePerKg: number;
  packaging: PackagingResult;
}

export interface PipeItem {
  id: string;
  type: 'pipe';
  peGrade: PipeGrade;
  diameter: number; // mm
  sdr: number;
  pn: number;
  length: number; // meters
  weightPerMeter: number; // kg/m
  totalWeight: number; // kg
  pricePerKg: number; // Toman
  totalPrice: number; // Toman
  packaging?: PackagingResult;
}

export interface FittingCatalogItem {
  id: string;
  category: string; // نوع: مثلاً اتصالات فشار قوی
  connectionType: string; // نوع محصول: مثلاً جوشی
  product: string; // محصول: مثلاً فلنچ پایه بلند
  material: string; // جنس: HDPE100
  pressure: string; // فشار: 6, 10, 16
  size: string; // سایز: 63, 75, 90
  unitPriceRial: number; // قیمت ریال
  packageCount: number; // تعداد در کارتن/کیسه
  isActive: boolean;
}

export interface FittingItem {
  id: string;
  type: 'fitting';
  catalogId?: string;
  category: string;
  connectionType: string;
  product: string;
  material: string;
  pressure: string;
  size: string;
  quantity: number;
  unitPrice: number; // in active currency (Toman or Rial)
  totalPrice: number;
}

export type InvoiceItem = PipeItem | FittingItem;

export interface Customer {
  id: string;
  name: string;
  phone?: string;
  address?: string;
  nationalId?: string;
  economicCode?: string;
  postalCode?: string;
  defaultDiscount?: number;
}

export interface DiscountRule {
  threshold: number; // Toman
  percentage: number;
}

export interface CompanySettings {
  name: string;
  address: string;
  phone: string;
  fax: string;
  taxRate: number; // percentage (e.g. 10%)
  invoiceCounter: number;
  currencyUnit: 'تومان' | 'ریال';
  basePipePrices: Record<PipeGrade, number>;
  economicCode?: string;
  nationalId?: string;
  registrationNumber?: string;
  postalCode?: string;
  bankName?: string;
  bankAccount?: string;
  bankIban?: string;
}

export interface Invoice {
  id: string;
  invoiceNumber: string;
  customerName: string;
  customerPhone?: string;
  customerAddress?: string;
  customerNationalId?: string;
  customerEconomicCode?: string;
  customerPostalCode?: string;
  currencyUnit?: 'تومان' | 'ریال';
  isOfficial?: boolean;
  jalaliDate: string;
  items: InvoiceItem[];
  subtotal: number;
  // Discounts
  includeDiscount: boolean;
  discountPercentage: number;
  discountAmount: number;
  includeCustomDiscount: boolean;
  customDiscountAmount: number;
  // Tax / Added Value
  includeTax: boolean;
  taxPercentage: number;
  taxAmount: number;
  // Final total
  totalAmount: number;
  explanation?: string;
  status: 'draft' | 'finalized';
  createdAt: string;
}

export interface WeightTable {
  [diameter: number]: {
    [sdr: number]: number;
  };
}

export interface GradeSdrPn {
  [grade: string]: {
    [sdr: number]: number;
  };
}

