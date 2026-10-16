import seedData from './seedData.json';
import { CompanySettings, DiscountRule, FittingCatalogItem, GradeSdrPn, WeightTable } from '../types';

export const initialWeightTable: WeightTable = seedData.weightTable as unknown as WeightTable;
export const initialGradeSdrPn: GradeSdrPn = seedData.gradeSdrPn as unknown as GradeSdrPn;
export const initialFittings: FittingCatalogItem[] = seedData.fittings as unknown as FittingCatalogItem[];
export const initialDiscounts: DiscountRule[] = seedData.discounts as unknown as DiscountRule[];
export const defaultBasePipePrices = {
  PE100: 95000,
  PE80: 88000,
  PE63: 82000,
};

export const initialCompanySettings: CompanySettings = {
  ...(seedData.companyInfo as any),
  basePipePrices: (seedData as any).basePipePrices || defaultBasePipePrices,
  economicCode: '411456789123',
  nationalId: '10103567890',
  registrationNumber: '384521',
  postalCode: '1548763211',
  bankName: 'بانک ملت (شعبه میرداماد)',
  bankAccount: '1234567890',
  bankIban: 'IR120120000000012345678901',
};

