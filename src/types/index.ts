export type PipeGrade = 'PE100' | 'PE80' | 'PE63';

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
}

export interface Invoice {
  id: string;
  invoiceNumber: string;
  customerName: string;
  customerPhone?: string;
  customerAddress?: string;
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

