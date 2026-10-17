import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Plus,
  Trash2,
  Printer,
  Save,
  Calculator,
  Layers,
  CheckCircle,
  Check,
  FileCheck,
  Edit3,
  Copy,
  RotateCcw,
  Truck,
  Factory,
  Package,
  Info,
  X,
  Sparkles,
  AlertCircle,
  ArrowRight,
  Wifi,
  WifiOff,
  RefreshCw,
  Globe,
  TrendingUp,
} from 'lucide-react';
import {
  CompanySettings,
  Customer,
  DiscountRule,
  FittingCatalogItem,
  GradeSdrPn,
  Invoice,
  InvoiceItem,
  PipeGrade,
  PipeItem,
  FittingItem,
  WeightTable,
  PackagingType,
  PackagingOptions,
  PackagingResult,
  FreightRecommendation,
  RawMaterialCostResult,
  MarketPolymerItem,
  MarketPriceSnapshot,
} from '../../types';
import { MarketPriceService } from '../../services/marketPriceService';
import {
  computeInvoiceSummary,
  computePipeItem,
  getApplicableDiscount,
  getWeightPerMeter,
  isCoilAllowed,
  computePackaging,
  estimateFreightVehicle,
  calculateLengthFromMass,
  calculatePricePerKgFromTotal,
  calculateReversePipe,
  calculateRawMaterialCost,
  FREIGHT_VEHICLES,
} from '../../services/calculator';
import {
  formatNumber,
  formatPrice,
  getCurrentJalaliDate,
  numberToPersianWords,
  toPersianDigits,
} from '../../services/persianUtils';
import { StorageService } from '../../services/storage';

interface InvoiceBuilderProps {
  settings: CompanySettings;
  fittings: FittingCatalogItem[];
  discounts: DiscountRule[];
  customers: Customer[];
  weightTable: WeightTable;
  gradeSdrPn: GradeSdrPn;
  editingSession?: { invoice: Invoice; isDuplicate: boolean } | null;
  onClearEditingSession?: () => void;
  onPreviewPrint: (invoice: Invoice) => void;
  onInvoiceSaved?: () => void;
  onUpdateSettings?: (settings: CompanySettings) => void;
}

export const InvoiceBuilder: React.FC<InvoiceBuilderProps> = ({
  settings,
  fittings,
  discounts,
  customers,
  weightTable,
  gradeSdrPn,
  editingSession,
  onClearEditingSession,
  onPreviewPrint,
  onInvoiceSaved,
  onUpdateSettings,
}) => {
  // Active currency unit: from editing session if editing, or current settings
  const activeCurrency = editingSession?.invoice?.currencyUnit || settings.currencyUnit || 'تومان';

  // Base pipe price helper: settings store basePipePrices in Toman. In Rial, multiply by 10.
  const getBasePipePriceForGrade = (grade: PipeGrade, currency: 'تومان' | 'ریال') => {
    const baseToman = settings?.basePipePrices?.[grade] || (grade === 'PE100' ? 95000 : grade === 'PE80' ? 88000 : 82000);
    return currency === 'ریال' ? baseToman * 10 : baseToman;
  };

  // Invoice Metadata
  const [invoiceNumber, setInvoiceNumber] = useState<string>('');
  const [jalaliDate, setJalaliDate] = useState<string>(getCurrentJalaliDate());
  const [customerName, setCustomerName] = useState<string>('');
  const [customerPhone, setCustomerPhone] = useState<string>('');
  const [customerAddress, setCustomerAddress] = useState<string>('');
  const [customerNationalId, setCustomerNationalId] = useState<string>('');
  const [customerEconomicCode, setCustomerEconomicCode] = useState<string>('');
  const [customerPostalCode, setCustomerPostalCode] = useState<string>('');
  const [isOfficial, setIsOfficial] = useState<boolean>(false);
  const [explanation, setExplanation] = useState<string>('');

  // Items List
  const [items, setItems] = useState<InvoiceItem[]>([]);

  // Active form section
  const [activeItemTab, setActiveItemTab] = useState<'pipe' | 'fitting'>('pipe');

  // --- PIPE FORM STATE ---
  const [pipeGrade, setPipeGrade] = useState<PipeGrade>('PE100');
  const availableDiameters = useMemo(() => {
    return Object.keys(weightTable)
      .map(Number)
      .sort((a, b) => a - b);
  }, [weightTable]);

  const [pipeDiameter, setPipeDiameter] = useState<number>(110);
  const [pipeSdr, setPipeSdr] = useState<number>(26);
  const [pipePn, setPipePn] = useState<number>(6.3);
  const [pipeLength, setPipeLength] = useState<number>(5);
  const [pipePricePerKg, setPipePricePerKg] = useState<number>(() =>
    getBasePipePriceForGrade('PE100', activeCurrency)
  );

  // Available SDRs for selected diameter
  const availableSdrs = useMemo(() => {
    const dObj = weightTable?.[pipeDiameter];
    if (!dObj) return [];
    return Object.keys(dObj)
      .map(Number)
      .sort((a, b) => a - b);
  }, [weightTable, pipeDiameter]);

  // Sync PN when grade or SDR changes
  useEffect(() => {
    const gradeMap = gradeSdrPn?.[pipeGrade];
    if (gradeMap && gradeMap[pipeSdr] !== undefined) {
      setPipePn(gradeMap[pipeSdr]);
    }
  }, [pipeGrade, pipeSdr, gradeSdrPn]);

  // Update default price when grade or active currency changes
  useEffect(() => {
    setPipePricePerKg(getBasePipePriceForGrade(pipeGrade, activeCurrency));
  }, [pipeGrade, activeCurrency, settings?.basePipePrices]);

  // Compute live pipe weight per meter
  const currentWpm = useMemo(() => {
    return getWeightPerMeter(pipeDiameter, pipeSdr, weightTable);
  }, [pipeDiameter, pipeSdr, weightTable]);

  // Live pipe price per meter, total weight, and total price
  const currentPipePricePerMeter = currentWpm ? Math.round(currentWpm * pipePricePerKg) : 0;
  const currentPipeWeight = currentWpm ? Number((currentWpm * pipeLength).toFixed(3)) : 0;
  const currentPipeTotalPrice = Math.round(currentPipeWeight * pipePricePerKg);

  // --- PACKAGING STATE FOR PIPE FORM ---
  const [pipePackagingType, setPipePackagingType] = useState<PackagingType>('branch');
  const [pipeBranchLength, setPipeBranchLength] = useState<5 | 6 | 12>(5);
  const [pipeCoilLength, setPipeCoilLength] = useState<number>(100);

  // Sync packaging mode if diameter > 125mm (strictly branch only)
  useEffect(() => {
    if (!isCoilAllowed(pipeDiameter)) {
      setPipePackagingType('branch');
    }
  }, [pipeDiameter]);

  const currentPackaging = useMemo(() => {
    return computePackaging(pipeDiameter, pipeLength, {
      preferredType: isCoilAllowed(pipeDiameter) ? pipePackagingType : 'branch',
      branchLength: pipeBranchLength,
      coilLength: pipeCoilLength,
    });
  }, [pipeDiameter, pipeLength, pipePackagingType, pipeBranchLength, pipeCoilLength]);

  // --- REVERSE PIPE CALCULATION STATE ---
  const [showReversePanel, setShowReversePanel] = useState<boolean>(false);
  const [reverseMode, setReverseMode] = useState<'tonnage' | 'budget'>('tonnage');
  const [reverseWeightUnit, setReverseWeightUnit] = useState<'ton' | 'kg'>('ton');
  const [reverseTargetValue, setReverseTargetValue] = useState<number>(5); // default 5 tons
  const [reverseDiameter, setReverseDiameter] = useState<number>(110);
  const [reverseSdr, setReverseSdr] = useState<number>(26);
  const [reverseGrade, setReverseGrade] = useState<PipeGrade>('PE100');
  const [reversePn, setReversePn] = useState<number>(6.3);
  const [reversePricePerKg, setReversePricePerKg] = useState<number>(() =>
    getBasePipePriceForGrade('PE100', activeCurrency)
  );
  const [reversePackagingType, setReversePackagingType] = useState<PackagingType>('branch');
  const [reverseBranchLength, setReverseBranchLength] = useState<5 | 6 | 12>(5);
  const [reverseCoilLength, setReverseCoilLength] = useState<number>(100);

  // Available SDRs for reverse diameter
  const reverseAvailableSdrs = useMemo(() => {
    const dObj = weightTable?.[reverseDiameter];
    if (!dObj) return [];
    return Object.keys(dObj).map(Number).sort((a, b) => a - b);
  }, [weightTable, reverseDiameter]);

  // Sync PN for reverse
  useEffect(() => {
    const gradeMap = gradeSdrPn?.[reverseGrade];
    if (gradeMap && gradeMap[reverseSdr] !== undefined) {
      setReversePn(gradeMap[reverseSdr]);
    }
  }, [reverseGrade, reverseSdr, gradeSdrPn]);

  // Sync reverse price with grade and active currency
  useEffect(() => {
    setReversePricePerKg(getBasePipePriceForGrade(reverseGrade, activeCurrency));
  }, [reverseGrade, activeCurrency, settings?.basePipePrices]);

  // Sync packaging allowance for reverse diameter
  useEffect(() => {
    if (!isCoilAllowed(reverseDiameter)) {
      setReversePackagingType('branch');
    }
  }, [reverseDiameter]);

  // Live reverse calculation
  const reverseResult = useMemo(() => {
    const targetTonnage = reverseMode === 'tonnage' && reverseWeightUnit === 'ton' ? reverseTargetValue : undefined;
    const targetWeightKg = reverseMode === 'tonnage' && reverseWeightUnit === 'kg' ? reverseTargetValue : undefined;
    const targetBudget = reverseMode === 'budget' ? reverseTargetValue : undefined;

    return calculateReversePipe(
      {
        mode: reverseMode,
        diameter: reverseDiameter,
        sdr: reverseSdr,
        peGrade: reverseGrade,
        pricePerKg: reversePricePerKg,
        targetTonnage,
        targetWeightKg,
        targetBudget,
        packagingPreference: isCoilAllowed(reverseDiameter) ? reversePackagingType : 'branch',
        branchLengthPreference: reverseBranchLength,
        coilLengthPreference: reverseCoilLength,
      },
      weightTable
    );
  }, [
    reverseMode,
    reverseWeightUnit,
    reverseTargetValue,
    reverseDiameter,
    reverseSdr,
    reverseGrade,
    reversePricePerKg,
    reversePackagingType,
    reverseBranchLength,
    reverseCoilLength,
    weightTable,
  ]);

  // --- RAW MATERIAL COST CALCULATOR STATE ---
  const [showCostModal, setShowCostModal] = useState<boolean>(false);
  const [costTargetForm, setCostTargetForm] = useState<'pipe' | 'reverse'>('pipe');
  const [rawPolymerPrice, setRawPolymerPrice] = useState<number>(() => 85000 * (activeCurrency === 'ریال' ? 10 : 1));
  const [rawMasterbatchPct, setRawMasterbatchPct] = useState<number>(2.0);
  const [rawMasterbatchPrice, setRawMasterbatchPrice] = useState<number>(() => 130000 * (activeCurrency === 'ریال' ? 10 : 1));
  const [rawWastePct, setRawWastePct] = useState<number>(1.5);
  const [rawConversionCost, setRawConversionCost] = useState<number>(() => 10000 * (activeCurrency === 'ریال' ? 10 : 1));
  const [rawProfitMarginPct, setRawProfitMarginPct] = useState<number>(10);

  const rawCostResult = useMemo(() => {
    return calculateRawMaterialCost({
      polymerPricePerKg: rawPolymerPrice,
      masterbatchPercentage: rawMasterbatchPct,
      masterbatchPricePerKg: rawMasterbatchPrice,
      wastePercentage: rawWastePct,
      conversionCostPerKg: rawConversionCost,
      profitMarginPercentage: rawProfitMarginPct,
    });
  }, [
    rawPolymerPrice,
    rawMasterbatchPct,
    rawMasterbatchPrice,
    rawWastePct,
    rawConversionCost,
    rawProfitMarginPct,
  ]);

  // --- ONLINE & OFFLINE MARKET POLYMER PRICES STATE (IME / BOURSE KALA) ---
  const [marketSnapshot, setMarketSnapshot] = useState<MarketPriceSnapshot>(() =>
    MarketPriceService.getStoredMarketPrices()
  );
  const [isFetchingMarket, setIsFetchingMarket] = useState<boolean>(false);

  const handleRefreshMarketPrices = async () => {
    setIsFetchingMarket(true);
    try {
      const live = await MarketPriceService.fetchLiveMarketPrices();
      setMarketSnapshot(live);
      if (live.isOnline) {
        notify('✓ استعلام زنده: آخرین نرخ‌های بورس کالا با موفقیت دریافت شدند.');
      } else {
        notify('حالت آفلاین: آخرین نرخ مصوب رسمی پتروشیمی از حافظه برنامه بارگذاری شد.');
      }
    } catch {
      notify('عدم برقراری ارتباط اینترنت؛ نرخ‌های ذخیره‌شده آفلاین معتبر هستند.');
    } finally {
      setIsFetchingMarket(false);
    }
  };

  const handleSelectMarketItem = (item: MarketPolymerItem) => {
    const unitPrice = activeCurrency === 'ریال' ? item.pricePerKgToman * 10 : item.pricePerKgToman;
    if (item.grade === 'Masterbatch') {
      setRawMasterbatchPrice(unitPrice);
      notify(`نرخ مستربچ دوده به ${formatNumber(unitPrice)} ${activeCurrency} تنظیم شد.`);
    } else {
      setRawPolymerPrice(unitPrice);
      notify(`نرخ پایه پلیمر (${item.grade}) به ${formatNumber(unitPrice)} ${activeCurrency} تنظیم شد.`);
    }
  };

  const handleApplyAllToSystem = () => {
    const updated = MarketPriceService.applyAllMarketPrices(marketSnapshot);
    onUpdateSettings?.(updated);
    notify('✓ نرخ‌های رسمی بورس کالا بر روی قیمت‌های پایه کل سیستم اعمال گردید.');
  };

  const handleOpenCostModal = (target: 'pipe' | 'reverse') => {
    setCostTargetForm(target);
    const targetGrade = target === 'pipe' ? pipeGrade : reverseGrade;
    const mult = activeCurrency === 'ریال' ? 10 : 1;
    const polymerItem = marketSnapshot.items.find((item) => item.grade === targetGrade) || marketSnapshot.items[0];
    const mbItem = marketSnapshot.items.find((item) => item.grade === 'Masterbatch');
    if (polymerItem) {
      setRawPolymerPrice(polymerItem.pricePerKgToman * mult);
    }
    if (mbItem) {
      setRawMasterbatchPrice(mbItem.pricePerKgToman * mult);
    }
    setRawConversionCost(10000 * mult);
    setShowCostModal(true);
  };

  // --- FITTING FORM STATE ---
  const activeFittings = useMemo(() => fittings.filter((f) => f.isActive), [fittings]);

  const fittingCategories = useMemo(() => {
    return Array.from(new Set(activeFittings.map((f) => f.category)));
  }, [activeFittings]);

  const [fittingCategory, setFittingCategory] = useState<string>(
    fittingCategories[0] || 'اتصالات فشار قوی'
  );

  const fittingTypes = useMemo(() => {
    return Array.from(
      new Set(
        activeFittings
          .filter((f) => f.category === fittingCategory)
          .map((f) => f.connectionType)
      )
    );
  }, [activeFittings, fittingCategory]);

  const [fittingType, setFittingType] = useState<string>(fittingTypes[0] || 'جوشی');

  const fittingProducts = useMemo(() => {
    return Array.from(
      new Set(
        activeFittings
          .filter(
            (f) => f.category === fittingCategory && f.connectionType === fittingType
          )
          .map((f) => f.product)
      )
    );
  }, [activeFittings, fittingCategory, fittingType]);

  const [fittingProduct, setFittingProduct] = useState<string>(fittingProducts[0] || '');

  const availablePressures = useMemo(() => {
    return Array.from(
      new Set(
        activeFittings
          .filter(
            (f) =>
              f.category === fittingCategory &&
              f.connectionType === fittingType &&
              f.product === fittingProduct
          )
          .map((f) => f.pressure)
      )
    );
  }, [activeFittings, fittingCategory, fittingType, fittingProduct]);

  const [fittingPressure, setFittingPressure] = useState<string>(availablePressures[0] || '10');

  const availableSizes = useMemo(() => {
    return Array.from(
      new Set(
        activeFittings
          .filter(
            (f) =>
              f.category === fittingCategory &&
              f.connectionType === fittingType &&
              f.product === fittingProduct &&
              f.pressure === fittingPressure
          )
          .map((f) => f.size)
      )
    );
  }, [activeFittings, fittingCategory, fittingType, fittingProduct, fittingPressure]);

  const [fittingSize, setFittingSize] = useState<string>(availableSizes[0] || '');
  const [fittingQuantity, setFittingQuantity] = useState<number>(10);
  const [fittingUnitPrice, setFittingUnitPrice] = useState<number>(0);

  // Selected fitting catalog match
  const selectedFittingCatalog = useMemo(() => {
    return activeFittings.find(
      (f) =>
        f.category === fittingCategory &&
        f.connectionType === fittingType &&
        f.product === fittingProduct &&
        f.pressure === fittingPressure &&
        f.size === fittingSize
    );
  }, [activeFittings, fittingCategory, fittingType, fittingProduct, fittingPressure, fittingSize]);

  // Update unit price when matched fitting changes or active currency changes
  useEffect(() => {
    if (selectedFittingCatalog) {
      // In Toman or Rial
      const price =
        activeCurrency === 'تومان'
          ? Math.round(selectedFittingCatalog.unitPriceRial / 10)
          : selectedFittingCatalog.unitPriceRial;
      setFittingUnitPrice(price);
    }
  }, [selectedFittingCatalog, activeCurrency]);

  // Reset cascading selections when parent changes
  useEffect(() => {
    if (fittingTypes.length && !fittingTypes.includes(fittingType)) {
      setFittingType(fittingTypes[0]);
    }
  }, [fittingTypes]);

  useEffect(() => {
    if (fittingProducts.length && !fittingProducts.includes(fittingProduct)) {
      setFittingProduct(fittingProducts[0]);
    }
  }, [fittingProducts]);

  useEffect(() => {
    if (availablePressures.length && !availablePressures.includes(fittingPressure)) {
      setFittingPressure(availablePressures[0]);
    }
  }, [availablePressures]);

  useEffect(() => {
    if (availableSizes.length && !availableSizes.includes(fittingSize)) {
      setFittingSize(availableSizes[0]);
    }
  }, [availableSizes]);

  // --- DISCOUNT & TAX STATE ---
  const [includeDiscount, setIncludeDiscount] = useState<boolean>(false);
  const [discountPercentage, setDiscountPercentage] = useState<number>(0);
  const [includeCustomDiscount, setIncludeCustomDiscount] = useState<boolean>(false);
  const [customDiscountAmount, setCustomDiscountAmount] = useState<number>(0);
  const [includeTax, setIncludeTax] = useState<boolean>(false);
  const [taxPercentage, setTaxPercentage] = useState<number>(settings.taxRate || 10);

  // Sync taxPercentage when settings.taxRate changes (unless editing an existing finalized invoice)
  useEffect(() => {
    if (!editingSession || editingSession.isDuplicate) {
      setTaxPercentage(settings.taxRate || 10);
    }
  }, [settings.taxRate, editingSession]);

  // Hydration tracking
  const [isHydrated, setIsHydrated] = useState<boolean>(false);

  // Restore draft or load editingSession
  useEffect(() => {
    if (editingSession) {
      const inv = editingSession.invoice;
      setInvoiceNumber(inv.invoiceNumber);
      setJalaliDate(editingSession.isDuplicate ? getCurrentJalaliDate() : (inv.jalaliDate || getCurrentJalaliDate()));
      setCustomerName(inv.customerName || '');
      setCustomerPhone(inv.customerPhone || '');
      setCustomerAddress(inv.customerAddress || '');
      setCustomerNationalId(inv.customerNationalId || '');
      setCustomerEconomicCode(inv.customerEconomicCode || '');
      setCustomerPostalCode(inv.customerPostalCode || '');
      setIsOfficial(Boolean(inv.isOfficial));
      setExplanation(inv.explanation || '');
      setItems(inv.items ? JSON.parse(JSON.stringify(inv.items)) : []);
      setIncludeDiscount(Boolean(inv.includeDiscount));
      setDiscountPercentage(inv.discountPercentage || 0);
      setIncludeCustomDiscount(Boolean(inv.includeCustomDiscount));
      setCustomDiscountAmount(inv.customDiscountAmount || 0);
      setIncludeTax(Boolean(inv.includeTax));
      setTaxPercentage(typeof inv.taxPercentage === 'number' ? inv.taxPercentage : (settings.taxRate || 10));
      setIsHydrated(true);
      if (editingSession.isDuplicate) {
        notify(`فاکتور شماره ${inv.invoiceNumber} به عنوان پیش‌فاکتور جدید تکثیر گردید.`);
      } else {
        notify(`فاکتور شماره ${inv.invoiceNumber} جهت ویرایش بارگذاری شد.`);
      }
    } else {
      // Check for saved draft in localStorage
      const draft = StorageService.getDraft();
      if (draft && typeof draft === 'object') {
        const validCustomerName = typeof draft.customerName === 'string' ? draft.customerName : '';
        const validInvoiceNumber =
          typeof draft.invoiceNumber === 'string' || typeof draft.invoiceNumber === 'number'
            ? String(draft.invoiceNumber)
            : '';
        const hasItems = Array.isArray(draft.items) && draft.items.length > 0;
        const hasCustomerName = validCustomerName.trim().length > 0;
        const hasInvoiceNumber = validInvoiceNumber.trim().length > 0;

        if (hasItems || hasCustomerName || hasInvoiceNumber) {
          if (validInvoiceNumber) setInvoiceNumber(validInvoiceNumber);
          if (typeof draft.jalaliDate === 'string' && draft.jalaliDate) setJalaliDate(draft.jalaliDate);
          if (validCustomerName) setCustomerName(validCustomerName);
          if (typeof draft.customerPhone === 'string') setCustomerPhone(draft.customerPhone);
          if (typeof draft.customerAddress === 'string') setCustomerAddress(draft.customerAddress);
          if (typeof draft.customerNationalId === 'string') setCustomerNationalId(draft.customerNationalId);
          if (typeof draft.customerEconomicCode === 'string') setCustomerEconomicCode(draft.customerEconomicCode);
          if (typeof draft.customerPostalCode === 'string') setCustomerPostalCode(draft.customerPostalCode);
          if (typeof draft.isOfficial === 'boolean') setIsOfficial(draft.isOfficial);
          if (typeof draft.explanation === 'string') setExplanation(draft.explanation);
          if (Array.isArray(draft.items)) {
            const sanitizedItems = draft.items.filter((item): item is InvoiceItem => Boolean(item && typeof item === 'object' && item.type));
            setItems(sanitizedItems);
          }
          if (typeof draft.includeDiscount === 'boolean') setIncludeDiscount(draft.includeDiscount);
          if (typeof draft.discountPercentage === 'number') setDiscountPercentage(draft.discountPercentage);
          if (typeof draft.includeCustomDiscount === 'boolean') setIncludeCustomDiscount(draft.includeCustomDiscount);
          if (typeof draft.customDiscountAmount === 'number') setCustomDiscountAmount(draft.customDiscountAmount);
          if (typeof draft.includeTax === 'boolean') setIncludeTax(draft.includeTax);
          if (typeof draft.taxPercentage === 'number') setTaxPercentage(draft.taxPercentage);
          notify('پیش‌نویس ذخیره‌شده فاکتور بازیابی گردید.');
        } else {
          const freshSettings = StorageService.getSettings();
          const nextNum = (freshSettings.invoiceCounter || settings.invoiceCounter || 1) + 1;
          setInvoiceNumber(String(nextNum));
        }
      } else {
        const freshSettings = StorageService.getSettings();
        const nextNum = (freshSettings.invoiceCounter || settings.invoiceCounter || 1) + 1;
        setInvoiceNumber(String(nextNum));
      }
      setIsHydrated(true);
    }
  }, [editingSession]);

  // Real-time draft auto-save
  useEffect(() => {
    if (!isHydrated) return;
    // Do not overwrite draft if editing an existing invoice
    if (editingSession && !editingSession.isDuplicate) return;

    const hasData =
      items.length > 0 ||
      (typeof customerName === 'string' && customerName.trim().length > 0) ||
      (typeof customerPhone === 'string' && customerPhone.trim().length > 0) ||
      (typeof customerAddress === 'string' && customerAddress.trim().length > 0) ||
      (typeof customerNationalId === 'string' && customerNationalId.trim().length > 0) ||
      (typeof customerEconomicCode === 'string' && customerEconomicCode.trim().length > 0) ||
      (typeof customerPostalCode === 'string' && customerPostalCode.trim().length > 0) ||
      (typeof explanation === 'string' && explanation.trim().length > 0);

    if (hasData) {
      StorageService.saveDraft({
        invoiceNumber,
        jalaliDate,
        customerName,
        customerPhone,
        customerAddress,
        customerNationalId,
        customerEconomicCode,
        customerPostalCode,
        isOfficial,
        explanation,
        items,
        includeDiscount,
        discountPercentage,
        includeCustomDiscount,
        customDiscountAmount,
        includeTax,
        taxPercentage,
        currencyUnit: activeCurrency,
      });
    } else {
      StorageService.clearDraft();
    }
  }, [
    isHydrated,
    editingSession,
    items,
    customerName,
    customerPhone,
    customerAddress,
    customerNationalId,
    customerEconomicCode,
    customerPostalCode,
    isOfficial,
    explanation,
    includeDiscount,
    discountPercentage,
    includeCustomDiscount,
    customDiscountAmount,
    includeTax,
    taxPercentage,
    invoiceNumber,
    jalaliDate,
    activeCurrency,
  ]);

  // Inline edit for pipe item: length and pricePerKg (preserving existing packaging options)
  const handleUpdatePipeItem = (id: string, newLength: number, newPricePerKg: number) => {
    const len = Math.max(0, Number(newLength) || 0);
    const price = Math.max(0, Number(newPricePerKg) || 0);
    setItems((prev) =>
      prev.map((item) => {
        if (item.id === id && item.type === 'pipe') {
          const packagingOptions: PackagingOptions | undefined = item.packaging
            ? {
                preferredType: item.packaging.packagingType,
                branchLength:
                  item.packaging.unitLength === 12
                    ? 12
                    : item.packaging.unitLength === 5
                    ? 5
                    : 6,
                coilLength:
                  item.packaging.packagingType === 'coil' ? item.packaging.unitLength : undefined,
              }
            : undefined;
          const updated = computePipeItem(
            {
              peGrade: item.peGrade,
              diameter: item.diameter,
              sdr: item.sdr,
              pn: item.pn,
              length: len,
              pricePerKg: price,
              packagingOptions,
            },
            weightTable,
            item.id
          );
          return updated || item;
        }
        return item;
      })
    );
  };

  // Inline edit for fitting item: quantity and unitPrice
  const handleUpdateFittingItem = (id: string, newQuantity: number, newUnitPrice: number) => {
    setItems((prev) =>
      prev.map((item) => {
        if (item.id === id && item.type === 'fitting') {
          const qty = Math.max(0, newQuantity);
          const price = Math.max(0, newUnitPrice);
          return {
            ...item,
            quantity: qty,
            unitPrice: price,
            totalPrice: Math.round(qty * price),
          };
        }
        return item;
      })
    );
  };

  // Calculate live summary
  const subtotalBeforeDiscounts = items.reduce(
    (sum, item) => sum + (item && item.totalPrice ? item.totalPrice : 0),
    0
  );
  const autoDiscountPct = getApplicableDiscount(subtotalBeforeDiscounts, discounts);
  const effectiveDiscountPct = discountPercentage > 0 ? discountPercentage : autoDiscountPct;

  const summary = useMemo(() => {
    return computeInvoiceSummary({
      items,
      includeDiscount,
      discountPercentage: effectiveDiscountPct,
      includeCustomDiscount,
      customDiscountAmount,
      includeTax,
      taxPercentage,
    });
  }, [
    items,
    includeDiscount,
    effectiveDiscountPct,
    includeCustomDiscount,
    customDiscountAmount,
    includeTax,
    taxPercentage,
  ]);

  // Notification message
  const [notification, setNotification] = useState<string | null>(null);
  const notify = (msg: string) => {
    setNotification(msg);
    setTimeout(() => setNotification(null), 3000);
  };

  // Customer selection & autofill
  const handleCustomerNameChange = (name: string) => {
    setCustomerName(name);
    const matched = customers.find(
      (c) => c && c.name && typeof c.name === 'string' && c.name.trim().toLowerCase() === name.trim().toLowerCase()
    );
    if (matched) {
      setCustomerPhone(matched.phone || '');
      setCustomerAddress(matched.address || '');
      setCustomerNationalId(matched.nationalId || '');
      setCustomerEconomicCode(matched.economicCode || '');
      setCustomerPostalCode(matched.postalCode || '');
      if (matched.nationalId || matched.economicCode) {
        setIsOfficial(true);
      }
      if (matched.defaultDiscount && matched.defaultDiscount > 0) {
        if (matched.defaultDiscount <= 100) {
          setIncludeDiscount(true);
          setDiscountPercentage(matched.defaultDiscount);
          setIncludeCustomDiscount(false);
          setCustomDiscountAmount(0);
          notify(`اطلاعات و درصد تخفیف ${toPersianDigits(matched.defaultDiscount)}٪ مشتری اعمال شد.`);
        } else {
          setIncludeCustomDiscount(true);
          setCustomDiscountAmount(matched.defaultDiscount);
          setIncludeDiscount(false);
          setDiscountPercentage(0);
          notify(`اطلاعات و تخفیف ${formatNumber(matched.defaultDiscount)} ${activeCurrency} مشتری اعمال شد.`);
        }
      } else {
        setDiscountPercentage(0);
        setIncludeCustomDiscount(false);
        setCustomDiscountAmount(0);
        notify('اطلاعات مشتری (تلفن و نشانی) بارگذاری شد.');
      }
    }
  };

  // Add Pipe Item handler
  const handleAddPipeItem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentWpm || currentWpm <= 0) {
      alert('ترکیب قطر و SDR انتخاب شده در جدول استانداردهای DIN معتبر نیست.');
      return;
    }

    if (pipeLength <= 0) {
      alert('لطفاً طول معتبر به متر وارد کنید.');
      return;
    }

    const item = computePipeItem(
      {
        peGrade: pipeGrade,
        diameter: pipeDiameter,
        sdr: pipeSdr,
        pn: pipePn,
        length: pipeLength,
        pricePerKg: pipePricePerKg,
        packagingOptions: {
          preferredType: isCoilAllowed(pipeDiameter) ? pipePackagingType : 'branch',
          branchLength: pipeBranchLength,
          coilLength: pipeCoilLength,
        },
      },
      weightTable
    );

    if (item) {
      setItems([...items, item]);
      notify(`لوله سایز ${pipeDiameter} با گرید ${pipeGrade} (${item.packaging?.labelPersian}) افزوده شد.`);
    }
  };

  // Insert Reverse Calculated Pipe Item handler
  const handleInsertReverseItem = () => {
    if (!reverseResult) {
      alert('محاسبه معکوس با مقادیر فعلی معتبر نیست.');
      return;
    }

    const item = computePipeItem(
      {
        peGrade: reverseGrade,
        diameter: reverseDiameter,
        sdr: reverseSdr,
        pn: reversePn,
        length: reverseResult.calculatedLength,
        pricePerKg: reverseResult.unitPricePerKg,
        packagingOptions: {
          preferredType: reverseResult.packaging.packagingType,
          branchLength:
            reverseResult.packaging.unitLength === 12
              ? 12
              : reverseResult.packaging.unitLength === 5
              ? 5
              : 6,
          coilLength: reverseResult.packaging.unitLength,
        },
      },
      weightTable
    );

    if (item) {
      setItems((prev) => [...prev, item]);
      notify(`لوله معکوس محاسبه‌شده (${reverseResult.calculatedLength} متر - ${reverseResult.packaging.labelPersian}) به پیش‌فاکتور افزوده شد.`);
      setShowReversePanel(false);
    }
  };

  // Apply suggested raw material cost price
  const handleApplyCostPrice = () => {
    if (costTargetForm === 'pipe') {
      setPipePricePerKg(rawCostResult.suggestedPricePerKg);
      notify(`نرخ پیشنهادی بورس کالا (${formatNumber(rawCostResult.suggestedPricePerKg)} ${activeCurrency}) در فرم لوله درج شد.`);
    } else {
      setReversePricePerKg(rawCostResult.suggestedPricePerKg);
      notify(`نرخ پیشنهادی بورس کالا (${formatNumber(rawCostResult.suggestedPricePerKg)} ${activeCurrency}) در محاسبه معکوس درج شد.`);
    }
    setShowCostModal(false);
  };

  // Add Fitting Item handler
  const handleAddFittingItem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!fittingProduct || !fittingSize) {
      alert('لطفاً مشخصات اتصال را تکمیل کنید.');
      return;
    }

    if (fittingQuantity <= 0) {
      alert('تعداد باید حداقل ۱ باشد.');
      return;
    }

    const item: FittingItem = {
      id: 'fit_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      type: 'fitting',
      catalogId: selectedFittingCatalog?.id,
      category: fittingCategory,
      connectionType: fittingType,
      product: fittingProduct,
      material: selectedFittingCatalog?.material || 'HDPE100',
      pressure: fittingPressure,
      size: fittingSize,
      quantity: fittingQuantity,
      unitPrice: fittingUnitPrice,
      totalPrice: fittingQuantity * fittingUnitPrice,
    };

    setItems([...items, item]);
    notify(`${fittingProduct} سایز ${fittingSize} به تعداد ${fittingQuantity} افزوده شد.`);
  };

  // Remove Item
  const handleRemoveItem = (id: string) => {
    setItems(items.filter((i) => i.id !== id));
  };

  // Clear / Reset all
  const resetToNewDraft = () => {
    setItems([]);
    setExplanation('');
    setCustomerName('');
    setCustomerPhone('');
    setCustomerAddress('');
    setCustomerNationalId('');
    setCustomerEconomicCode('');
    setCustomerPostalCode('');
    setIsOfficial(false);
    setIncludeDiscount(false);
    setDiscountPercentage(0);
    setIncludeCustomDiscount(false);
    setCustomDiscountAmount(0);
    setIncludeTax(false);
    setTaxPercentage(settings.taxRate || 10);
    StorageService.clearDraft();
    onClearEditingSession?.();
    const freshSettings = StorageService.getSettings();
    const nextNum = (freshSettings.invoiceCounter || settings.invoiceCounter || 1) + 1;
    setInvoiceNumber(String(nextNum));
    setJalaliDate(getCurrentJalaliDate());
    notify('پیش‌فاکتور جدید آماده گردید.');
  };

  const handleClearAll = () => {
    if (items.length > 0 && confirm('آیا از پاک کردن تمام اقلام و شروع فاکتور جدید اطمینان دارید؟')) {
      resetToNewDraft();
    }
  };

  // Construct current invoice object
  const buildCurrentInvoice = (status: 'draft' | 'finalized' = 'finalized'): Invoice => {
    const isEditMode = Boolean(editingSession && !editingSession.isDuplicate);
    const invoiceId = isEditMode && editingSession?.invoice.id ? editingSession.invoice.id : 'inv_' + Date.now();
    return {
      id: invoiceId,
      invoiceNumber: invoiceNumber.trim() || '1',
      customerName: customerName.trim() || 'مشتری آزاد',
      customerPhone: customerPhone.trim(),
      customerAddress: customerAddress.trim(),
      customerNationalId: customerNationalId.trim() || undefined,
      customerEconomicCode: customerEconomicCode.trim() || undefined,
      customerPostalCode: customerPostalCode.trim() || undefined,
      isOfficial,
      jalaliDate,
      items,
      subtotal: summary.subtotal,
      includeDiscount,
      discountPercentage: effectiveDiscountPct,
      discountAmount: summary.discountAmount,
      includeCustomDiscount,
      customDiscountAmount,
      includeTax,
      taxPercentage,
      taxAmount: summary.taxAmount,
      totalAmount: summary.totalAmount,
      explanation,
      currencyUnit: activeCurrency,
      status,
      createdAt: isEditMode && editingSession?.invoice.createdAt ? editingSession.invoice.createdAt : new Date().toISOString(),
    };
  };

  // Save invoice
  const handleSaveInvoice = () => {
    if (items.length === 0) {
      alert('پیش‌فاکتور خالی است. ابتدا حداقل یک کالا اضافه کنید.');
      return;
    }

    const isEditMode = Boolean(editingSession && !editingSession.isDuplicate);
    const inv = buildCurrentInvoice('finalized');
    StorageService.saveInvoice(inv);
    StorageService.clearDraft();

    // Save customer if name provided
    if (customerName.trim()) {
      StorageService.addOrUpdateCustomer({
        id: 'cust_' + Date.now(),
        name: customerName.trim(),
        phone: customerPhone.trim(),
        address: customerAddress.trim(),
        nationalId: customerNationalId.trim() || undefined,
        economicCode: customerEconomicCode.trim() || undefined,
        postalCode: customerPostalCode.trim() || undefined,
      });
    }

    // Increment invoice counter only if NEW invoice (not editing)
    if (!isEditMode) {
      const currentNum = Number(invoiceNumber) || 1;
      if (currentNum >= (settings.invoiceCounter || 0)) {
        const updatedSettings = {
          ...settings,
          invoiceCounter: currentNum,
        };
        StorageService.saveSettings(updatedSettings);
      }
    }

    notify(
      isEditMode
        ? `پیش‌فاکتور شماره ${invoiceNumber} با موفقیت ویرایش و بروزرسانی شد.`
        : `پیش‌فاکتور شماره ${invoiceNumber} در سوابق ذخیره شد.`
    );
    onInvoiceSaved?.();
  };

  // Preview & Print
  const handlePreview = () => {
    if (items.length === 0) {
      alert('لطفاً قبل از پیش‌نمایش، حداقل یک کالا اضافه کنید.');
      return;
    }
    const inv = buildCurrentInvoice('finalized');
    onPreviewPrint(inv);
  };

  // Stable ref for global shortcut handlers so listeners bind only once on mount
  const shortcutHandlersRef = useRef({
    onSave: handleSaveInvoice,
    onPreview: handlePreview,
    onReset: resetToNewDraft,
  });
  shortcutHandlersRef.current = {
    onSave: handleSaveInvoice,
    onPreview: handlePreview,
    onReset: resetToNewDraft,
  };

  // Listen for global shortcut events
  useEffect(() => {
    const onSave = () => shortcutHandlersRef.current.onSave();
    const onPreview = () => shortcutHandlersRef.current.onPreview();
    const onReset = () => shortcutHandlersRef.current.onReset();

    window.addEventListener('pipecal:save-invoice', onSave);
    window.addEventListener('pipecal:preview-invoice', onPreview);
    window.addEventListener('pipecal:reset-draft', onReset);

    return () => {
      window.removeEventListener('pipecal:save-invoice', onSave);
      window.removeEventListener('pipecal:preview-invoice', onPreview);
      window.removeEventListener('pipecal:reset-draft', onReset);
    };
  }, []);

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Toast Notification */}
      {notification && (
        <div className="fixed bottom-6 left-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-xl shadow-2xl flex items-center gap-3 animate-fade-in">
          <CheckCircle className="w-5 h-5" />
          <span>{notification}</span>
        </div>
      )}

      {/* Edit Mode Banner */}
      {editingSession && !editingSession.isDuplicate && (
        <div className="bg-amber-500/10 border-2 border-amber-500/40 rounded-2xl p-4 flex flex-col sm:flex-row items-center justify-between gap-4 shadow-lg animate-fade-in">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-amber-500/20 rounded-xl text-amber-400">
              <Edit3 className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-amber-300 text-base">
                  حالت ویرایش فاکتور شماره {toPersianDigits(editingSession.invoice.invoiceNumber)}
                </span>
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 font-mono">
                  {editingSession.invoice.id}
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-1">
                تغییرات شما به صورت مستقیم روی این فاکتور در سوابق ذخیره و بروزرسانی می‌شود.
              </p>
            </div>
          </div>
          <button
            onClick={resetToNewDraft}
            className="flex items-center gap-1.5 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-semibold transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>انصراف از ویرایش / ایجاد فاکتور جدید</span>
          </button>
        </div>
      )}

      {/* Duplicate Mode Banner */}
      {editingSession && editingSession.isDuplicate && (
        <div className="bg-indigo-500/10 border-2 border-indigo-500/40 rounded-2xl p-4 flex flex-col sm:flex-row items-center justify-between gap-4 shadow-lg animate-fade-in">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-indigo-500/20 rounded-xl text-indigo-400">
              <Copy className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-indigo-300 text-base">
                  تکثیر از فاکتور شماره {toPersianDigits(editingSession.invoice.invoiceNumber)}
                </span>
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-300">
                  پیش‌فاکتور جدید
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-1">
                اقلام با موفقیت بارگذاری شدند. فاکتور تحت شماره جدید {toPersianDigits(invoiceNumber)} صادر خواهد شد.
              </p>
            </div>
          </div>
          <button
            onClick={resetToNewDraft}
            className="flex items-center gap-1.5 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-xs font-semibold transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>پاکسازی و شروع فاکتور خالی</span>
          </button>
        </div>
      )}

      {/* TOP: Customer & Invoice Details Card */}
      <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 shadow-lg space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-700 pb-4">
          <h2 className="text-xl font-bold text-white flex items-center gap-2">
            <FileCheck className="w-6 h-6 text-blue-400" />
            <span>مشخصات پیش‌فاکتور و مشتری</span>
          </h2>

          <div className="flex flex-wrap items-center gap-3">
            <button
              type="button"
              onClick={() => setIsOfficial((prev) => !prev)}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border transition-all cursor-pointer ${
                isOfficial
                  ? 'bg-amber-600/20 text-amber-300 border-amber-500/50 shadow-sm'
                  : 'bg-slate-900/80 text-slate-400 hover:text-slate-200 border-slate-700'
              }`}
              title="فعال‌سازی مشخصات مالیاتی خریدار برای صورتحساب رسمی ماده ۱۶۹"
            >
              <FileCheck className="w-3.5 h-3.5" />
              <span>{isOfficial ? '✓ پیش‌فاکتور رسمی (ماده ۱۶۹)' : 'پیش‌فاکتور رسمی (ماده ۱۶۹)'}</span>
            </button>

            <div className="flex items-center gap-2 text-sm bg-slate-900/80 px-3 py-1.5 rounded-xl border border-slate-700">
              <span className="text-slate-400">شماره پیش‌فاکتور:</span>
              <input
                type="text"
                value={invoiceNumber}
                onChange={(e) => setInvoiceNumber(e.target.value)}
                className="w-20 bg-slate-800 border border-slate-600 rounded px-2 py-0.5 text-center font-mono font-bold text-white focus:outline-none"
              />
            </div>

            <div className="flex items-center gap-2 text-sm bg-slate-900/80 px-3 py-1.5 rounded-xl border border-slate-700">
              <span className="text-slate-400">تاریخ:</span>
              <input
                type="text"
                value={jalaliDate}
                onChange={(e) => setJalaliDate(e.target.value)}
                className="w-28 bg-slate-800 border border-slate-600 rounded px-2 py-0.5 text-center font-mono text-white focus:outline-none"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div>
            <label className="block text-xs text-slate-400 mb-1">نام مشتری / سازمان خریدار:</label>
            <div className="relative">
              <input
                type="text"
                value={customerName}
                onChange={(e) => handleCustomerNameChange(e.target.value)}
                placeholder="مثال: آقای رضایی / شرکت آب و خاک"
                list="customers-list"
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
              />
              <datalist id="customers-list">
                {customers.map((c) => (
                  <option key={c.id} value={c.name} />
                ))}
              </datalist>
            </div>
          </div>

          <div>
            <label className="block text-xs text-slate-400 mb-1">شماره تلفن یا همراه:</label>
            <input
              type="text"
              value={customerPhone}
              onChange={(e) => setCustomerPhone(e.target.value)}
              placeholder="مثال: ۰۹۱۲۳۴۵۶۷۸۹"
              className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm font-mono text-left"
            />
          </div>

          <div>
            <label className="block text-xs text-slate-400 mb-1">نشانی و آدرس تحویل:</label>
            <input
              type="text"
              value={customerAddress}
              onChange={(e) => setCustomerAddress(e.target.value)}
              placeholder="شهر، پروژه، آدرس انبار..."
              className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
            />
          </div>
        </div>

        {/* Official Tax Invoice Buyer Details (ماده ۱۶۹ م.م) */}
        {(isOfficial || customerEconomicCode || customerNationalId || customerPostalCode) && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-3 border-t border-slate-700/70 animate-fade-in">
            <div>
              <label className="block text-xs text-amber-300/90 mb-1">شماره اقتصادی خریدار (ماده ۱۶۹):</label>
              <input
                type="text"
                value={customerEconomicCode}
                onChange={(e) => setCustomerEconomicCode(e.target.value)}
                placeholder="مثال: ۴۱۱۴۵۶۷۸۹۱۲۳"
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-amber-500 focus:outline-none text-sm font-mono text-left"
              />
            </div>

            <div>
              <label className="block text-xs text-amber-300/90 mb-1">شناسه ملی / کد ملی خریدار:</label>
              <input
                type="text"
                value={customerNationalId}
                onChange={(e) => setCustomerNationalId(e.target.value)}
                placeholder="مثال: ۱۰۱۰۳۵۶۷۸۹۰"
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-amber-500 focus:outline-none text-sm font-mono text-left"
              />
            </div>

            <div>
              <label className="block text-xs text-amber-300/90 mb-1">کد پستی ۱۰ رقمی خریدار:</label>
              <input
                type="text"
                value={customerPostalCode}
                onChange={(e) => setCustomerPostalCode(e.target.value)}
                placeholder="مثال: ۱۵۴۸۷۶۳۲۱۱"
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white placeholder-slate-500 focus:ring-2 focus:ring-amber-500 focus:outline-none text-sm font-mono text-left"
              />
            </div>
          </div>
        )}
      </div>

      {/* MIDDLE: Add Item Tabs (Pipe vs Fitting) */}
      <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-lg overflow-hidden">
        {/* Sub-tabs header */}
        <div className="flex border-b border-slate-700 bg-slate-900/50">
          <button
            onClick={() => setActiveItemTab('pipe')}
            className={`flex-1 py-3.5 px-6 font-bold text-sm flex items-center justify-center gap-2 border-b-2 transition-all cursor-pointer ${
              activeItemTab === 'pipe'
                ? 'border-blue-500 text-blue-400 bg-slate-800/90'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Calculator className="w-4 h-4" />
            <span>افزودن لوله پلی‌اتیلن (PE Pipes)</span>
          </button>

          <button
            onClick={() => setActiveItemTab('fitting')}
            className={`flex-1 py-3.5 px-6 font-bold text-sm flex items-center justify-center gap-2 border-b-2 transition-all cursor-pointer ${
              activeItemTab === 'fitting'
                ? 'border-emerald-500 text-emerald-400 bg-slate-800/90'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Layers className="w-4 h-4" />
            <span>افزودن اتصالات پلی‌اتیلن (Fittings)</span>
          </button>
        </div>

        {/* --- FORM 1: ADD PIPE --- */}
        {activeItemTab === 'pipe' && (
          <div className="p-6 space-y-6">
            {/* Top Toolbar for PE Domain Capabilities */}
            <div className="flex flex-wrap items-center justify-between gap-3 p-3.5 bg-slate-900/70 rounded-xl border border-slate-700/70">
              <div className="flex items-center gap-2">
                <span className="text-xs text-slate-400 font-medium">ابزارهای تخصصی صنعت پلی‌اتیلن:</span>
                <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30">DIN 8074/8075</span>
              </div>

              <div className="flex items-center gap-2">
                {/* Reverse Calculation Toggle */}
                <button
                  type="button"
                  onClick={() => setShowReversePanel((prev) => !prev)}
                  className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border ${
                    showReversePanel
                      ? 'bg-blue-600 text-white border-blue-500 shadow-md shadow-blue-600/30'
                      : 'bg-slate-800 text-blue-400 hover:text-blue-300 border-slate-700 hover:border-slate-600'
                  }`}
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>{showReversePanel ? 'بستن محاسبه معکوس' : 'محاسبه معکوس (تناژ / بودجه)'}</span>
                </button>

                {/* Raw Material Cost Calculator Modal */}
                <button
                  type="button"
                  onClick={() => handleOpenCostModal('pipe')}
                  className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold bg-slate-800 text-emerald-400 hover:text-emerald-300 border border-slate-700 hover:border-slate-600 transition-all cursor-pointer"
                >
                  <Factory className="w-3.5 h-3.5" />
                  <span>آنالیز بورس کالا و هزینه تبدیل</span>
                </button>
              </div>
            </div>

            {/* --- REVERSE CALCULATION PANEL (TOGGLEABLE) --- */}
            {showReversePanel && (
              <div className="bg-gradient-to-br from-slate-900 via-slate-850 to-slate-900 p-5 rounded-2xl border-2 border-blue-500/50 shadow-xl space-y-5 animate-fade-in">
                <div className="flex items-center justify-between border-b border-slate-700/80 pb-3">
                  <div className="flex items-center gap-2.5">
                    <div className="p-2 bg-blue-500/20 text-blue-400 rounded-lg">
                      <RotateCcw className="w-5 h-5" />
                    </div>
                    <div>
                      <h4 className="font-bold text-white text-sm">محاسبه‌گر معکوس متراژ لوله (Reverse Pipe Calculator)</h4>
                      <p className="text-xs text-slate-400">تخمین متراژ مورد نیاز بر اساس تناژ درخواستی، وزن کل یا بودجه پروژه</p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => setShowReversePanel(false)}
                    className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>

                {/* Mode Selector */}
                <div className="flex items-center gap-3">
                  <span className="text-xs text-slate-400 font-medium">مبنای محاسبه معکوس:</span>
                  <div className="flex bg-slate-950 p-1 rounded-xl border border-slate-800">
                    <button
                      type="button"
                      onClick={() => setReverseMode('tonnage')}
                      className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                        reverseMode === 'tonnage'
                          ? 'bg-blue-600 text-white shadow-sm'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      بر اساس وزن / تناژ کل
                    </button>
                    <button
                      type="button"
                      onClick={() => setReverseMode('budget')}
                      className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                        reverseMode === 'budget'
                          ? 'bg-blue-600 text-white shadow-sm'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      بر اساس بودجه کل پروژه
                    </button>
                  </div>
                </div>

                {/* Input Fields */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 text-xs">
                  {/* Target Input */}
                  {reverseMode === 'tonnage' ? (
                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-slate-400">وزن یا تناژ هدف:</label>
                        <div className="flex text-[10px] gap-1 bg-slate-950 px-1 py-0.5 rounded border border-slate-800">
                          <button
                            type="button"
                            onClick={() => setReverseWeightUnit('ton')}
                            className={`px-1 rounded ${reverseWeightUnit === 'ton' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}
                          >
                            تن
                          </button>
                          <button
                            type="button"
                            onClick={() => setReverseWeightUnit('kg')}
                            className={`px-1 rounded ${reverseWeightUnit === 'kg' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}
                          >
                            کیلوگرم
                          </button>
                        </div>
                      </div>
                      <input
                        type="number"
                        step="any"
                        value={reverseTargetValue || ''}
                        onChange={(e) => setReverseTargetValue(Number(e.target.value) || 0)}
                        placeholder={reverseWeightUnit === 'ton' ? 'مثال: ۵ (تن)' : 'مثال: ۵۰۰۰ (کیلوگرم)'}
                        className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-blue-500 font-bold"
                      />
                    </div>
                  ) : (
                    <div>
                      <label className="block text-slate-400 mb-1">بودجه کل پروژه ({activeCurrency}):</label>
                      <input
                        type="number"
                        value={reverseTargetValue || ''}
                        onChange={(e) => setReverseTargetValue(Number(e.target.value) || 0)}
                        placeholder="مثال: ۵۰,۰۰۰,۰۰۰"
                        className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-blue-500 font-bold"
                      />
                    </div>
                  )}

                  {/* Reverse Grade */}
                  <div>
                    <label className="block text-slate-400 mb-1">گرید لوله:</label>
                    <select
                      value={reverseGrade}
                      onChange={(e) => setReverseGrade(e.target.value as PipeGrade)}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-bold focus:outline-none"
                    >
                      <option value="PE100">PE100</option>
                      <option value="PE80">PE80</option>
                      <option value="PE63">PE63</option>
                    </select>
                  </div>

                  {/* Reverse Diameter */}
                  <div>
                    <label className="block text-slate-400 mb-1">قطر خارجی (mm):</label>
                    <select
                      value={reverseDiameter}
                      onChange={(e) => {
                        const d = Number(e.target.value);
                        setReverseDiameter(d);
                        const sdrs = Object.keys(weightTable[d] || {}).map(Number);
                        if (sdrs.length && !sdrs.includes(reverseSdr)) {
                          setReverseSdr(sdrs.includes(26) ? 26 : sdrs[0]);
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-bold focus:outline-none"
                    >
                      {availableDiameters.map((d) => (
                        <option key={d} value={d}>
                          {d} میلی‌متر
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Reverse SDR */}
                  <div>
                    <label className="block text-slate-400 mb-1">نسبت SDR:</label>
                    <select
                      value={reverseSdr}
                      onChange={(e) => setReverseSdr(Number(e.target.value))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono focus:outline-none"
                    >
                      {reverseAvailableSdrs.map((s) => (
                        <option key={s} value={s}>
                          SDR {s} ({gradeSdrPn?.[reverseGrade]?.[s] ? `${gradeSdrPn[reverseGrade][s]} bar` : ''})
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Reverse Price/Kg */}
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="text-slate-400">قیمت هر کیلو ({activeCurrency}):</label>
                      <button
                        type="button"
                        onClick={() => handleOpenCostModal('reverse')}
                        className="text-[10px] text-emerald-400 hover:underline cursor-pointer"
                      >
                        آنالیز بورس
                      </button>
                    </div>
                    <input
                      type="number"
                      value={reversePricePerKg || ''}
                      onChange={(e) => setReversePricePerKg(Number(e.target.value) || 0)}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-emerald-400 font-mono font-bold text-left focus:outline-none"
                    />
                  </div>
                </div>

                {/* Packaging Mode in Reverse */}
                <div className="bg-slate-950/70 p-3.5 rounded-xl border border-slate-800 space-y-2">
                  <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                    <span className="text-slate-300 font-medium">نحوه بسته‌بندی در محاسبه معکوس:</span>
                    {isCoilAllowed(reverseDiameter) ? (
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setReversePackagingType('coil')}
                          className={`px-2.5 py-1 rounded-lg text-xs font-medium cursor-pointer transition-colors ${
                            reversePackagingType === 'coil'
                              ? 'bg-blue-600 text-white'
                              : 'bg-slate-800 text-slate-400 hover:text-white'
                          }`}
                        >
                          کلاف (Coil)
                        </button>
                        <button
                          type="button"
                          onClick={() => setReversePackagingType('branch')}
                          className={`px-2.5 py-1 rounded-lg text-xs font-medium cursor-pointer transition-colors ${
                            reversePackagingType === 'branch'
                              ? 'bg-blue-600 text-white'
                              : 'bg-slate-800 text-slate-400 hover:text-white'
                          }`}
                        >
                          شاخه مستقیم
                        </button>

                        {reversePackagingType === 'branch' ? (
                          <div className="flex items-center gap-1 text-[11px] bg-slate-900 px-2 py-0.5 rounded border border-slate-700 font-mono">
                            <span className="text-slate-400">طول شاخه:</span>
                            <button
                              type="button"
                              onClick={() => setReverseBranchLength(5)}
                              className={`px-1.5 rounded ${reverseBranchLength === 5 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                            >
                              ۵ متری
                            </button>
                            <button
                              type="button"
                              onClick={() => setReverseBranchLength(6)}
                              className={`px-1.5 rounded ${reverseBranchLength === 6 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                            >
                              ۶ متری
                            </button>
                            <button
                              type="button"
                              onClick={() => setReverseBranchLength(12)}
                              className={`px-1.5 rounded ${reverseBranchLength === 12 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                            >
                              ۱۲ متری
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center gap-1 text-[11px] bg-slate-900 px-2 py-0.5 rounded border border-slate-700 font-mono">
                            <span className="text-slate-400">متراژ هر کلاف:</span>
                            <select
                              value={reverseCoilLength}
                              onChange={(e) => setReverseCoilLength(Number(e.target.value))}
                              className="bg-transparent text-blue-300 font-bold focus:outline-none"
                            >
                              <option value="50">۵۰ متری</option>
                              <option value="100">۱۰۰ متری</option>
                              <option value="150">۱۵۰ متری</option>
                              <option value="200">۲۰۰ متری</option>
                              <option value="400">۴۰۰ متری</option>
                            </select>
                          </div>
                        )}
                      </div>
                    ) : (
                      <div className="flex items-center gap-2">
                        <span className="text-amber-400 text-[11px] flex items-center gap-1 bg-amber-500/10 px-2.5 py-1 rounded border border-amber-500/20">
                          <AlertCircle className="w-3.5 h-3.5" />
                          سایز {reverseDiameter}mm الزماً شاخه مستقیم است (کلاف غیرمجاز است)
                        </span>
                        <div className="flex items-center gap-1 text-[11px] bg-slate-900 px-2 py-0.5 rounded border border-slate-700 font-mono">
                          <button
                            type="button"
                            onClick={() => setReverseBranchLength(5)}
                            className={`px-2 py-0.5 rounded ${reverseBranchLength === 5 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۵ متری
                          </button>
                          <button
                            type="button"
                            onClick={() => setReverseBranchLength(6)}
                            className={`px-2 py-0.5 rounded ${reverseBranchLength === 6 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۶ متری
                          </button>
                          <button
                            type="button"
                            onClick={() => setReverseBranchLength(12)}
                            className={`px-2 py-0.5 rounded ${reverseBranchLength === 12 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۱۲ متری
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {/* Live Calculation Output Card */}
                {reverseResult ? (
                  <div className="bg-slate-950 p-4 rounded-xl border border-blue-500/30 flex flex-wrap items-center justify-between gap-4">
                    <div className="grid grid-cols-2 sm:grid-cols-6 gap-4 text-xs flex-1">
                      <div>
                        <span className="text-slate-400 block">وزن هر متر (DIN):</span>
                        <strong className="text-slate-200 text-sm font-mono">{reverseResult.weightPerMeter} kg/m</strong>
                      </div>

                      <div>
                        <span className="text-slate-400 block">قیمت هر متر:</span>
                        <strong className="text-cyan-400 text-sm font-mono font-bold">
                          {formatPrice(reverseResult.pricePerMeter ?? Math.round(reverseResult.weightPerMeter * reverseResult.unitPricePerKg), activeCurrency)}
                        </strong>
                      </div>

                      <div>
                        <span className="text-slate-400 block">متراژ دقیق حاصله:</span>
                        <strong className="text-blue-400 text-base font-mono font-bold">
                          {formatNumber(reverseResult.calculatedLength)} متر
                        </strong>
                      </div>

                      <div>
                        <span className="text-slate-400 block">وزن کل نهایی:</span>
                        <strong className="text-slate-200 text-sm font-mono">
                          {formatNumber(reverseResult.totalWeightKg)} kg ({toPersianDigits(reverseResult.totalTonnage)} تن)
                        </strong>
                      </div>

                      <div>
                        <span className="text-slate-400 block">بسته‌بندی پیشنهادی:</span>
                        <strong className="text-amber-300 text-sm font-mono block">
                          {reverseResult.packaging.labelPersian}
                        </strong>
                        {reverseResult.packaging.wasteOrSurplusMeters > 0 && (
                          <span className="text-[10px] text-slate-500 font-mono">
                            ({toPersianDigits(reverseResult.packaging.wasteOrSurplusMeters)}م مازاد)
                          </span>
                        )}
                      </div>

                      <div>
                        <span className="text-slate-400 block">مبلغ کل این ردیف:</span>
                        <strong className="text-emerald-400 text-base font-bold font-mono">
                          {formatPrice(reverseResult.totalPrice, activeCurrency)}
                        </strong>
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={handleInsertReverseItem}
                      className="flex items-center gap-2 px-5 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer shrink-0"
                    >
                      <Plus className="w-4 h-4" />
                      <span>افزودن این لوله به پیش‌فاکتور</span>
                    </button>
                  </div>
                ) : (
                  <div className="p-3 bg-amber-500/10 border border-amber-500/20 text-amber-300 rounded-xl text-xs text-center">
                    لطفاً مقادیر معتبر برای محاسبه معکوس متراژ وارد کنید.
                  </div>
                )}
              </div>
            )}

            {/* Standard Pipe Calculation Form */}
            <form onSubmit={handleAddPipeItem} className="space-y-5">
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-4 text-sm">
                {/* Grade */}
                <div>
                  <label className="block text-xs text-slate-400 mb-1">گرید لوله:</label>
                  <select
                    value={pipeGrade}
                    onChange={(e) => setPipeGrade(e.target.value as PipeGrade)}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-bold focus:outline-none"
                  >
                    <option value="PE100">PE100</option>
                    <option value="PE80">PE80</option>
                    <option value="PE63">PE63</option>
                  </select>
                </div>

                {/* Diameter */}
                <div>
                  <label className="block text-xs text-slate-400 mb-1">قطر خارجی (mm):</label>
                  <select
                    value={pipeDiameter}
                    onChange={(e) => {
                      const d = Number(e.target.value);
                      setPipeDiameter(d);
                      const sdrs = Object.keys(weightTable[d] || {}).map(Number);
                      if (sdrs.length && !sdrs.includes(pipeSdr)) {
                        setPipeSdr(sdrs.includes(26) ? 26 : sdrs[0]);
                      }
                    }}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-bold focus:outline-none"
                  >
                    {availableDiameters.map((d) => (
                      <option key={d} value={d}>
                        {d} میلی‌متر
                      </option>
                    ))}
                  </select>
                </div>

                {/* SDR */}
                <div>
                  <label className="block text-xs text-slate-400 mb-1">نسبت SDR:</label>
                  <select
                    value={pipeSdr}
                    onChange={(e) => setPipeSdr(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono focus:outline-none"
                  >
                    {availableSdrs.map((s) => (
                      <option key={s} value={s}>
                        SDR {s}
                      </option>
                    ))}
                  </select>
                </div>

                {/* PN (Auto-calculated) */}
                <div>
                  <label className="block text-xs text-slate-400 mb-1">فشار کاری (PN):</label>
                  <div className="w-full bg-slate-900/60 border border-slate-700/60 rounded-xl px-3 py-2 text-blue-400 font-mono font-bold">
                    {pipePn} bar
                  </div>
                </div>

                {/* Length */}
                <div>
                  <label className="block text-xs text-slate-400 mb-1">طول لوله (متر):</label>
                  <input
                    type="number"
                    step="any"
                    value={pipeLength || ''}
                    onChange={(e) => setPipeLength(Number(e.target.value) || 0)}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                  />
                </div>

                {/* Price per Kg */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs text-slate-400">
                      قیمت هر کیلو ({activeCurrency}):
                    </label>
                    <button
                      type="button"
                      onClick={() => handleOpenCostModal('pipe')}
                      className="text-[10px] text-emerald-400 hover:underline cursor-pointer"
                    >
                      آنالیز بورس
                    </button>
                  </div>
                  <input
                    type="number"
                    value={pipePricePerKg || ''}
                    onChange={(e) => setPipePricePerKg(Number(e.target.value) || 0)}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-emerald-400 font-mono font-bold text-left focus:outline-none"
                  />
                </div>
              </div>

              {/* Packaging Rule Section (Coil vs Branch) */}
              <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-700/60 space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                  <div className="flex items-center gap-2">
                    <Package className="w-4 h-4 text-blue-400" />
                    <span className="font-bold text-white">شیوه بسته‌بندی و تحویل کالا:</span>
                  </div>

                  {isCoilAllowed(pipeDiameter) ? (
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setPipePackagingType('coil')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-medium cursor-pointer transition-colors ${
                          pipePackagingType === 'coil'
                            ? 'bg-blue-600 text-white shadow-sm'
                            : 'bg-slate-800 text-slate-400 hover:text-white'
                        }`}
                      >
                        کلاف (Coil)
                      </button>

                      <button
                        type="button"
                        onClick={() => setPipePackagingType('branch')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-medium cursor-pointer transition-colors ${
                          pipePackagingType === 'branch'
                            ? 'bg-blue-600 text-white shadow-sm'
                            : 'bg-slate-800 text-slate-400 hover:text-white'
                        }`}
                      >
                        شاخه مستقیم
                      </button>

                      {pipePackagingType === 'branch' ? (
                        <div className="flex items-center gap-1.5 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-700 text-xs font-mono">
                          <span className="text-slate-400">طول شاخه:</span>
                          <button
                            type="button"
                            onClick={() => setPipeBranchLength(5)}
                            className={`px-2 py-0.5 rounded ${pipeBranchLength === 5 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۵ متری
                          </button>
                          <button
                            type="button"
                            onClick={() => setPipeBranchLength(6)}
                            className={`px-2 py-0.5 rounded ${pipeBranchLength === 6 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۶ متری
                          </button>
                          <button
                            type="button"
                            onClick={() => setPipeBranchLength(12)}
                            className={`px-2 py-0.5 rounded ${pipeBranchLength === 12 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                          >
                            ۱۲ متری
                          </button>
                        </div>
                      ) : (
                        <div className="flex items-center gap-1.5 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-700 text-xs font-mono">
                          <span className="text-slate-400">متراژ هر کلاف:</span>
                          <select
                            value={pipeCoilLength}
                            onChange={(e) => setPipeCoilLength(Number(e.target.value))}
                            className="bg-transparent text-blue-300 font-bold focus:outline-none"
                          >
                            <option value="50">۵۰ متری</option>
                            <option value="100">۱۰۰ متری</option>
                            <option value="150">۱۵۰ متری</option>
                            <option value="200">۲۰۰ متری</option>
                            <option value="400">۴۰۰ متری</option>
                          </select>
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="flex flex-wrap items-center gap-2">
                      <div className="flex items-center gap-1 text-[11px] text-amber-400 bg-amber-500/10 px-2.5 py-1 rounded-lg border border-amber-500/20">
                        <AlertCircle className="w-3.5 h-3.5" />
                        <span>قطر {pipeDiameter}mm: کلاف‌پیچی غیرمجاز است (الزام شاخه مستقیم ۵، ۶ یا ۱۲ متری)</span>
                      </div>
                      <div className="flex items-center gap-1.5 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-700 text-xs font-mono">
                        <button
                          type="button"
                          onClick={() => setPipeBranchLength(5)}
                          className={`px-2.5 py-0.5 rounded ${pipeBranchLength === 5 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                        >
                          ۵ متری
                        </button>
                        <button
                          type="button"
                          onClick={() => setPipeBranchLength(6)}
                          className={`px-2.5 py-0.5 rounded ${pipeBranchLength === 6 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                        >
                          ۶ متری
                        </button>
                        <button
                          type="button"
                          onClick={() => setPipeBranchLength(12)}
                          className={`px-2.5 py-0.5 rounded ${pipeBranchLength === 12 ? 'bg-blue-500 text-white' : 'text-slate-400'}`}
                        >
                          ۱۲ متری
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* Packaging preview text */}
                <div className="text-xs text-slate-300 flex items-center justify-between pt-1 border-t border-slate-800">
                  <span>
                    واحد بسته‌بندی محاسباتی: <strong className="text-blue-300 font-mono">{currentPackaging.labelPersian}</strong>
                  </span>
                  {currentPackaging.wasteOrSurplusMeters > 0 && (
                    <span className="text-amber-400 font-mono text-[11px]">
                      (شامل {toPersianDigits(currentPackaging.wasteOrSurplusMeters)} متر سرجمع اضافه متراژ ناشی از برش استاندارد)
                    </span>
                  )}
                </div>
              </div>

              {/* Calculations Preview Bar */}
              <div className="bg-slate-900/70 p-4 rounded-xl border border-slate-700/70 flex flex-wrap items-center justify-between gap-4 text-xs">
                <div className="flex flex-wrap items-center gap-6">
                  <div>
                    <span className="text-slate-400 block">وزن هر متر (DIN):</span>
                    <strong className="text-slate-200 text-sm font-mono">
                      {currentWpm ? `${currentWpm} kg/m` : 'نامعتبر'}
                    </strong>
                  </div>

                  <div>
                    <span className="text-slate-400 block">قیمت هر متر:</span>
                    <strong className="text-cyan-400 text-sm font-mono font-bold">
                      {currentWpm ? formatPrice(currentPipePricePerMeter, activeCurrency) : 'نامعتبر'}
                    </strong>
                  </div>

                  <div>
                    <span className="text-slate-400 block">وزن کل لوله:</span>
                    <strong className="text-blue-400 text-sm font-mono">
                      {formatNumber(currentPipeWeight)} کیلوگرم
                    </strong>
                  </div>

                  <div>
                    <span className="text-slate-400 block">بسته‌بندی:</span>
                    <strong className="text-amber-300 text-sm font-mono">
                      {currentPackaging.labelPersian}
                    </strong>
                  </div>

                  <div>
                    <span className="text-slate-400 block">قیمت کل این ردیف:</span>
                    <strong className="text-emerald-400 text-base font-bold font-mono">
                      {formatPrice(currentPipeTotalPrice, activeCurrency)}
                    </strong>
                  </div>
                </div>

                <button
                  type="submit"
                  disabled={!currentWpm || pipeLength <= 0}
                  className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white rounded-xl font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
                >
                  <Plus className="w-5 h-5" />
                  <span>افزودن این لوله به پیش‌فاکتور</span>
                </button>
              </div>
            </form>
          </div>
        )}

        {/* --- FORM 2: ADD FITTING --- */}
        {activeItemTab === 'fitting' && (
          <form onSubmit={handleAddFittingItem} className="p-6 space-y-5">
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-4 text-sm">
              {/* Category */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">دسته‌بندی (نوع):</label>
                <select
                  value={fittingCategory}
                  onChange={(e) => setFittingCategory(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                >
                  {fittingCategories.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>

              {/* Type */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">نوع محصول:</label>
                <select
                  value={fittingType}
                  onChange={(e) => setFittingType(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                >
                  {fittingTypes.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </div>

              {/* Product */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">نام محصول:</label>
                <select
                  value={fittingProduct}
                  onChange={(e) => setFittingProduct(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-medium focus:outline-none"
                >
                  {fittingProducts.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>

              {/* Pressure */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">فشار کاری (bar):</label>
                <select
                  value={fittingPressure}
                  onChange={(e) => setFittingPressure(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono focus:outline-none"
                >
                  {availablePressures.map((pr) => (
                    <option key={pr} value={pr}>
                      {pr} بار
                    </option>
                  ))}
                </select>
              </div>

              {/* Size */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">سایز (mm):</label>
                <select
                  value={fittingSize}
                  onChange={(e) => setFittingSize(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-blue-400 font-bold focus:outline-none"
                >
                  {availableSizes.map((s) => (
                    <option key={s} value={s}>
                      {s} mm
                    </option>
                  ))}
                </select>
              </div>

              {/* Quantity */}
              <div>
                <label className="block text-xs text-slate-400 mb-1">تعداد درخواستی:</label>
                <input
                  type="number"
                  min="1"
                  value={fittingQuantity}
                  onChange={(e) => setFittingQuantity(Number(e.target.value) || 1)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
            </div>

            {/* Fitting Calculation & Price Preview */}
            <div className="bg-slate-900/70 p-4 rounded-xl border border-slate-700/70 flex flex-wrap items-center justify-between gap-4 text-xs">
              <div className="flex items-center gap-6">
                <div>
                  <span className="text-slate-400 block">بسته‌بندی در کارتن:</span>
                  <strong className="text-slate-200 text-sm font-mono">
                    {selectedFittingCatalog?.packageCount || 1} عدد
                  </strong>
                </div>

                <div>
                  <span className="text-slate-400 block">قیمت واحد ({activeCurrency}):</span>
                  <input
                    type="number"
                    value={fittingUnitPrice}
                    onChange={(e) => setFittingUnitPrice(Number(e.target.value) || 0)}
                    className="w-36 bg-slate-800 border border-slate-600 rounded-lg px-2 py-1 text-emerald-400 font-mono font-bold text-left focus:outline-none text-sm"
                  />
                </div>

                <div>
                  <span className="text-slate-400 block">قیمت کل این اتصال:</span>
                  <strong className="text-emerald-400 text-base font-bold font-mono">
                    {formatPrice(fittingQuantity * fittingUnitPrice, activeCurrency)}
                  </strong>
                </div>
              </div>

              <button
                type="submit"
                disabled={!fittingProduct || !fittingSize || fittingQuantity <= 0}
                className="flex items-center gap-2 px-6 py-2.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white rounded-xl font-bold shadow-lg shadow-emerald-600/30 transition-all cursor-pointer"
              >
                <Plus className="w-5 h-5" />
                <span>افزودن این اتصال به پیش‌فاکتور</span>
              </button>
            </div>
          </form>
        )}
      </div>

      {/* ITEMS TABLE */}
      <div className="bg-slate-800/80 rounded-2xl border border-slate-700 shadow-xl overflow-hidden">
        <div className="px-6 py-4 bg-slate-800 border-b border-slate-700 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <h3 className="font-bold text-white">اقلام پیش‌فاکتور</h3>
            <span className="text-xs px-2.5 py-0.5 rounded-full bg-blue-500/20 text-blue-300 font-medium">
              {toPersianDigits(items.length)} ردیف
            </span>
          </div>

          {items.length > 0 && (
            <button
              onClick={handleClearAll}
              className="text-xs text-rose-400 hover:text-rose-300 flex items-center gap-1.5 p-1.5 rounded-lg hover:bg-rose-500/10 transition-colors cursor-pointer"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>پاک کردن تمام اقلام</span>
            </button>
          )}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-right text-sm">
            <thead className="bg-slate-900/80 text-xs font-semibold text-slate-400 border-b border-slate-700">
              <tr>
                <th className="py-3 px-4">ردیف</th>
                <th className="py-3 px-4">نوع قلم</th>
                <th className="py-3 px-4">شرح کالا و مشخصات فنی</th>
                <th className="py-3 px-4 text-center">مقدار / طول / تعداد</th>
                <th className="py-3 px-4 text-center">وزن واحد / کل (kg)</th>
                <th className="py-3 px-4 text-center">قیمت واحد / هر متر ({activeCurrency})</th>
                <th className="py-3 px-4 text-center">قیمت کل ({activeCurrency})</th>
                <th className="py-3 px-4 text-center">حذف</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-700/60">
              {items.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-12 text-center text-slate-500">
                    هنوز کالایی به این پیش‌فاکتور افزوده نشده است. از فرم‌های بالا اقلام لوله یا اتصالات را اضافه کنید.
                  </td>
                </tr>
              ) : (
                items.map((item, idx) => {
                  const isPipe = item.type === 'pipe';
                  return (
                    <tr key={item.id} className="hover:bg-slate-700/30 transition-colors">
                      <td className="py-3 px-4 font-mono text-slate-400 text-xs">{toPersianDigits(idx + 1)}</td>
                      <td className="py-3 px-4">
                        <span
                          className={`text-xs px-2 py-0.5 rounded-full font-medium ${
                            isPipe
                              ? 'bg-blue-500/20 text-blue-300 border border-blue-500/30'
                              : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                          }`}
                        >
                          {isPipe ? 'لوله پلی‌اتیلن' : 'اتصال جوشی'}
                        </span>
                      </td>

                      <td className="py-3 px-4">
                        {isPipe ? (
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-white">لوله قطر {item.diameter} mm</span>
                              <span className="text-slate-400 text-xs">
                                (گرید {item.peGrade} | SDR {item.sdr} | PN {item.pn})
                              </span>
                            </div>
                            {item.packaging && (
                              <div className="mt-1 flex items-center gap-1.5">
                                <span className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-md bg-slate-900 border border-slate-700/80 text-blue-300 font-mono">
                                  <Package className="w-3 h-3 text-blue-400" />
                                  <span>{item.packaging.labelPersian}</span>
                                </span>
                                {item.packaging.wasteOrSurplusMeters > 0 && (
                                  <span className="text-[10px] text-amber-400 font-mono">
                                    ({toPersianDigits(item.packaging.wasteOrSurplusMeters)}م مازاد)
                                  </span>
                                )}
                              </div>
                            )}
                          </div>
                        ) : (
                          <div>
                            <span className="font-bold text-white">{item.product}</span>
                            <span className="text-slate-400 text-xs mr-2">
                              (سایز {item.size} mm | فشار {item.pressure} bar | {item.connectionType})
                            </span>
                          </div>
                        )}
                      </td>

                      <td className="py-3 px-3 text-center">
                        {isPipe ? (
                          <div className="flex items-center justify-center gap-1.5">
                            <input
                              type="number"
                              min="0.1"
                              step="any"
                              value={item.length}
                              onChange={(e) => {
                                const val = parseFloat(e.target.value) || 0;
                                handleUpdatePipeItem(item.id, val, item.pricePerKg);
                              }}
                              className="w-24 bg-slate-900 border border-slate-700 focus:border-blue-500 rounded-lg px-2 py-1 text-center font-mono text-white text-xs focus:ring-1 focus:ring-blue-500 focus:outline-none"
                              title="ویرایش آنلاین متراژ لوله (متر)"
                            />
                            <span className="text-xs text-slate-400">متر</span>
                          </div>
                        ) : (
                          <div className="flex items-center justify-center gap-1.5">
                            <input
                              type="number"
                              min="1"
                              step="1"
                              value={item.quantity}
                              onChange={(e) => {
                                const val = parseInt(e.target.value, 10) || 0;
                                handleUpdateFittingItem(item.id, val, item.unitPrice);
                              }}
                              className="w-20 bg-slate-900 border border-slate-700 focus:border-emerald-500 rounded-lg px-2 py-1 text-center font-mono text-white text-xs focus:ring-1 focus:ring-emerald-500 focus:outline-none"
                              title="ویرایش آنلاین تعداد اتصالات (عدد)"
                            />
                            <span className="text-xs text-slate-400">عدد</span>
                          </div>
                        )}
                      </td>

                      <td className="py-3 px-4 text-center font-mono text-xs text-slate-300">
                        {isPipe ? (
                          <div>
                            <div className="font-bold text-white">{item.totalWeight} kg</div>
                            <div className="text-[10px] text-slate-500">({item.weightPerMeter} kg/m)</div>
                          </div>
                        ) : (
                          '-'
                        )}
                      </td>

                      <td className="py-3 px-3 text-center">
                        {isPipe ? (
                          <div className="flex flex-col items-center justify-center gap-1">
                            <div className="flex items-center justify-center gap-1">
                              <input
                                type="number"
                                min="0"
                                step="100"
                                value={item.pricePerKg}
                                onChange={(e) => {
                                  const val = parseFloat(e.target.value) || 0;
                                  handleUpdatePipeItem(item.id, item.length, val);
                                }}
                                className="w-28 bg-slate-900 border border-slate-700 focus:border-blue-500 rounded-lg px-2 py-1 text-center font-mono text-white text-xs focus:ring-1 focus:ring-blue-500 focus:outline-none"
                                title={`ویرایش آنلاین قیمت هر کیلوگرم (${activeCurrency})`}
                              />
                              <span className="text-[10px] text-slate-400">/kg</span>
                            </div>
                            <div className="text-[11px] text-cyan-400 font-mono">
                              هر متر: {formatNumber(item.pricePerMeter ?? Math.round(item.weightPerMeter * item.pricePerKg))}
                            </div>
                          </div>
                        ) : (
                          <div className="flex items-center justify-center gap-1">
                            <input
                              type="number"
                              min="0"
                              step="1000"
                              value={item.unitPrice}
                              onChange={(e) => {
                                const val = parseFloat(e.target.value) || 0;
                                handleUpdateFittingItem(item.id, item.quantity, val);
                              }}
                              className="w-28 bg-slate-900 border border-slate-700 focus:border-emerald-500 rounded-lg px-2 py-1 text-center font-mono text-white text-xs focus:ring-1 focus:ring-emerald-500 focus:outline-none"
                              title={`ویرایش آنلاین قیمت واحد اتصال (${activeCurrency})`}
                            />
                          </div>
                        )}
                      </td>

                      <td className="py-3 px-4 text-center font-mono font-bold text-emerald-400">
                        {formatNumber(item.totalPrice)}
                      </td>

                      <td className="py-3 px-4 text-center">
                        <button
                          onClick={() => handleRemoveItem(item.id)}
                          className="p-1 rounded text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 transition-colors cursor-pointer"
                          title="حذف ردیف"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* BOTTOM: Summary & Financial Adjustments */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Notes & Explanations & Logistics */}
        <div className="lg:col-span-2 space-y-4">
          <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-4">
            <label className="block text-sm font-bold text-white">توضیحات و شرایط پیش‌فاکتور (اختیاری):</label>
            <textarea
              rows={3}
              value={explanation}
              onChange={(e) => setExplanation(e.target.value)}
              placeholder="مثال: هزینه حمل و نقل به عهده خریدار می‌باشد. تحویل کالا ظرف مدت ۵ روز کاری از زمان واریز پیش‌پرداخت..."
              className="w-full bg-slate-900 border border-slate-700 rounded-xl p-3 text-sm text-white placeholder-slate-500 focus:ring-2 focus:ring-blue-500 focus:outline-none"
            />

            {summary.totalAmount > 0 && (
              <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-700/60 text-xs space-y-1">
                <span className="text-slate-400 block font-medium">مبلغ نهایی فاکتور به حروف:</span>
                <strong className="text-blue-300 text-sm block">
                  {numberToPersianWords(summary.totalAmount)} {activeCurrency}
                </strong>
              </div>
            )}
          </div>

          {/* Freight Logistics Estimator Card */}
          {summary.totalPipeWeight > 0 && summary.freight && (
            <div className="bg-gradient-to-br from-slate-900 via-slate-850 to-slate-900 p-5 rounded-2xl border border-blue-500/30 space-y-3.5 shadow-lg">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="p-2.5 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/30">
                    <Truck className="w-6 h-6" />
                  </div>
                  <div>
                    <span className="text-xs text-slate-400 block font-medium">برآورد ناوگان لجستیک و حمل (استاندارد ایران):</span>
                    <strong className="text-white text-base font-bold flex items-center gap-2">
                      <span>
                        {summary.freight.vehicleCount > 1 ? `${toPersianDigits(summary.freight.vehicleCount)} دستگاه ` : ''}
                        {summary.freight.vehicle.nameFa}
                      </span>
                    </strong>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <span className="px-3 py-1 rounded-full text-xs font-bold font-mono bg-blue-500/20 text-blue-300 border border-blue-500/30">
                    ضریب بار: {toPersianDigits(summary.freight.utilizationPercentage)}٪
                  </span>
                </div>
              </div>

              {/* Capacity meter */}
              <div className="w-full bg-slate-950 rounded-full h-2.5 overflow-hidden border border-slate-700/60">
                <div
                  className={`h-full transition-all duration-500 rounded-full ${
                    summary.freight.utilizationPercentage > 95
                      ? 'bg-rose-500'
                      : summary.freight.utilizationPercentage > 80
                      ? 'bg-amber-400'
                      : 'bg-blue-500'
                  }`}
                  style={{ width: `${Math.min(100, summary.freight.utilizationPercentage)}%` }}
                />
              </div>

              <div className="flex flex-wrap items-center justify-between text-xs text-slate-300 pt-2 border-t border-slate-700/60 gap-2">
                <span>
                  وزن کل لوله‌ها: <strong className="text-blue-400 font-mono">{formatNumber(summary.totalPipeWeight)} کیلوگرم</strong> ({toPersianDigits((summary.totalPipeWeight / 1000).toFixed(2))} تن)
                </span>
                {summary.freight.isLengthConstrained && (
                  <span className="text-amber-400 text-xs font-medium flex items-center gap-1 bg-amber-500/10 px-2.5 py-0.5 rounded-lg border border-amber-500/20">
                    <AlertCircle className="w-3.5 h-3.5" />
                    الزام طول بارگیر به دلیل شاخه مستقیم {toPersianDigits(summary.maxBranchLengthMeters)} متری
                  </span>
                )}
              </div>

              <div className="text-xs text-slate-300 bg-slate-950/60 p-3 rounded-xl border border-slate-800 flex items-start gap-2">
                <Info className="w-4 h-4 text-blue-400 shrink-0 mt-0.5" />
                <span>{summary.freight.recommendationNote}</span>
              </div>
            </div>
          )}
        </div>

        {/* Financial Calculation Box */}
        <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-4 shadow-xl">
          <h4 className="font-bold text-white border-b border-slate-700 pb-3 flex items-center justify-between">
            <span>محاسبات نهایی مالی</span>
            <span className="text-xs text-slate-400">واحد: {activeCurrency}</span>
          </h4>

          <div className="space-y-3 text-sm">
            {/* Subtotal */}
            <div className="flex justify-between items-center text-slate-300">
              <span>جمع اقلام:</span>
              <span className="font-mono font-bold">{formatNumber(summary.subtotal)}</span>
            </div>

            {/* Recommended Freight Tag in Financial Box */}
            {summary.totalPipeWeight > 0 && summary.freight && (
              <div className="flex justify-between items-center text-xs text-slate-400 pt-1 pb-1">
                <span className="flex items-center gap-1">
                  <Truck className="w-3.5 h-3.5 text-blue-400" />
                  <span>ناوگان حمل پیشنهادی:</span>
                </span>
                <span className="font-medium text-blue-300">
                  {summary.freight.vehicleCount > 1 ? `${toPersianDigits(summary.freight.vehicleCount)} × ` : ''}
                  {summary.freight.vehicle.nameFa}
                </span>
              </div>
            )}

            {/* Discount Option */}
            <div className="pt-2 border-t border-slate-700/60 space-y-2">
              <label className="flex items-center justify-between cursor-pointer">
                <span className="flex items-center gap-2 text-xs text-slate-300">
                  <input
                    type="checkbox"
                    checked={includeDiscount}
                    onChange={(e) => {
                      setIncludeDiscount(e.target.checked);
                      if (e.target.checked) setIncludeCustomDiscount(false);
                    }}
                    className="rounded text-blue-600 focus:ring-0"
                  />
                  <span>
                    {discountPercentage > 0
                      ? `تخفیف مشتری (${toPersianDigits(discountPercentage)}٪):`
                      : `تخفیف پله‌ای خودکار (${toPersianDigits(autoDiscountPct)}٪):`}
                  </span>
                </span>
                {includeDiscount && (
                  <span className="font-mono text-emerald-400 text-xs font-bold">
                    - {formatNumber(summary.discountAmount)}
                  </span>
                )}
              </label>

              {/* Custom Discount Option */}
              <label className="flex items-center justify-between cursor-pointer">
                <span className="flex items-center gap-2 text-xs text-slate-300">
                  <input
                    type="checkbox"
                    checked={includeCustomDiscount}
                    onChange={(e) => {
                      setIncludeCustomDiscount(e.target.checked);
                      if (e.target.checked) setIncludeDiscount(false);
                    }}
                    className="rounded text-blue-600 focus:ring-0"
                  />
                  <span>تخفیف مبلغ دلخواه:</span>
                </span>
              </label>

              {includeCustomDiscount && (
                <div className="relative pl-2">
                  <input
                    type="number"
                    placeholder={`مبلغ تخفیف به ${activeCurrency}`}
                    value={customDiscountAmount || ''}
                    onChange={(e) => setCustomDiscountAmount(Number(e.target.value) || 0)}
                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-white font-mono text-left focus:outline-none"
                  />
                </div>
              )}
            </div>

            {/* Tax Option */}
            <div className="pt-2 border-t border-slate-700/60">
              <label className="flex items-center justify-between cursor-pointer">
                <span className="flex items-center gap-2 text-xs text-slate-300">
                  <input
                    type="checkbox"
                    checked={includeTax}
                    onChange={(e) => setIncludeTax(e.target.checked)}
                    className="rounded text-blue-600 focus:ring-0"
                  />
                  <span>ارزش افزوده ({taxPercentage}%):</span>
                </span>
                {includeTax && (
                  <span className="font-mono text-amber-400 text-xs font-bold">
                    + {formatNumber(summary.taxAmount)}
                  </span>
                )}
              </label>
            </div>

            {/* Final Total */}
            <div className="pt-4 border-t-2 border-slate-700 flex justify-between items-center">
              <span className="font-extrabold text-white">مبلغ کل فاکتور:</span>
              <span className="font-mono font-black text-xl text-emerald-400">
                {formatNumber(summary.totalAmount)}
              </span>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="pt-4 space-y-2">
            <button
              onClick={handlePreview}
              disabled={items.length === 0}
              className="w-full flex items-center justify-center gap-2 py-3 bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white rounded-xl font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
            >
              <Printer className="w-5 h-5" />
              <span>پیش‌نمایش و چاپ / PDF</span>
            </button>

            <button
              onClick={handleSaveInvoice}
              disabled={items.length === 0}
              className="w-full flex items-center justify-center gap-2 py-2.5 bg-slate-700 hover:bg-slate-600 disabled:opacity-50 text-slate-200 rounded-xl font-medium transition-all cursor-pointer"
            >
              <Save className="w-4 h-4" />
              <span>
                {editingSession && !editingSession.isDuplicate
                  ? 'بروزرسانی و ثبت تغییرات فاکتور'
                  : 'ذخیره در سوابق فاکتورها'}
              </span>
            </button>
          </div>
        </div>
      </div>

      {/* --- MODAL: RAW MATERIAL COST CALCULATOR (بورس کالا و هزینه تبدیل) --- */}
      {showCostModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-3xl max-h-[92vh] overflow-y-auto w-full p-6 shadow-2xl space-y-6">
            <div className="flex items-center justify-between border-b border-slate-700 pb-4">
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                  <Factory className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-white">محاسبه‌گر بهای تمام‌شده و نرخ‌گذاری بورس کالا</h3>
                  <p className="text-xs text-slate-400">فرمول صنعتی محاسبه نرخ هر کیلوگرم لوله پلی‌اتیلن بر اساس بورس کالا (IME)</p>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setShowCostModal(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Official Iranian Market Prices Banner & Ticker */}
            <div className="bg-slate-950/90 rounded-xl p-4 border border-slate-800 space-y-3">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-800">
                <div className="flex items-center gap-2.5">
                  <div className={`p-2 rounded-lg border flex items-center justify-center ${
                    marketSnapshot.isOnline
                      ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30'
                      : 'bg-amber-500/20 text-amber-400 border-amber-500/30'
                  }`}>
                    {marketSnapshot.isOnline ? <Wifi className="w-4 h-4" /> : <WifiOff className="w-4 h-4" />}
                  </div>
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-white">نرخ‌های رسمی بورس کالای ایران (IME)</span>
                      <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium border ${
                        marketSnapshot.isOnline
                          ? 'bg-emerald-500/10 text-emerald-300 border-emerald-500/20'
                          : 'bg-amber-500/10 text-amber-300 border-amber-500/20'
                      }`}>
                        {marketSnapshot.isOnline ? '🟢 استعلام زنده آنلاین' : '🟡 آخرین نرخ مصوب رسمی (آفلاین)'}
                      </span>
                    </div>
                    <span className="text-[11px] text-slate-400 block mt-0.5">
                      مرجع: {marketSnapshot.source} (تاریخ: {marketSnapshot.lastUpdated})
                    </span>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={handleRefreshMarketPrices}
                    disabled={isFetchingMarket}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-500 disabled:opacity-60 text-white rounded-lg text-xs font-medium transition-all shadow-sm cursor-pointer"
                    title="استعلام آخرین نرخ‌های رسمی بورس کالا"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isFetchingMarket ? 'animate-spin' : ''}`} />
                    <span>{isFetchingMarket ? 'در حال استعلام...' : 'استعلام آنلاین نرخ بورس'}</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleApplyAllToSystem}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-lg text-xs font-medium transition-all cursor-pointer"
                    title="بروزرسانی قیمت پایه لوله در کل سیستم با نرخ‌های رسمی بورس"
                  >
                    <TrendingUp className="w-3.5 h-3.5 text-emerald-400" />
                    <span>اعمال روی نرخ پایه کل سیستم</span>
                  </button>
                </div>
              </div>

              {/* Categorized Market Prices: 1. Polyethylene Resins, 2. Masterbatch */}
              <div className="space-y-4 pt-1">
                {/* Section A: Polymer Resins (Base Material) */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-xs text-slate-300 font-bold border-b border-slate-800/80 pb-1">
                    <span className="flex items-center gap-1.5 text-blue-400">
                      <span className="w-2 h-2 rounded-full bg-blue-500 inline-block"></span>
                      <span>رزین‌های پایه پلی‌اتیلن لوله (بورس کالا)</span>
                    </span>
                    <span className="text-[11px] text-slate-400 font-normal">
                      انتخاب برای «قیمت پایه پلیمر»
                    </span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                    {marketSnapshot.items
                      .filter((item) => item.grade !== 'Masterbatch')
                      .map((item) => {
                        const displayPrice = activeCurrency === 'ریال' ? item.pricePerKgToman * 10 : item.pricePerKgToman;
                        const isSelected = rawPolymerPrice === displayPrice;

                        return (
                          <div
                            key={item.id}
                            className={`p-3 rounded-xl border transition-all text-xs flex flex-col justify-between ${
                              isSelected
                                ? 'bg-blue-950/40 border-blue-500/60 ring-1 ring-blue-500/40'
                                : 'bg-slate-900/70 border-slate-800 hover:border-slate-700'
                            }`}
                          >
                            <div>
                              <div className="flex items-start justify-between gap-1 mb-1.5">
                                <span className="font-bold text-white text-xs leading-snug">{item.name}</span>
                                {item.changePercent !== undefined && (
                                  <span
                                    dir="ltr"
                                    className={`text-[10px] font-mono font-bold px-1.5 py-0.5 rounded inline-flex shrink-0 ${
                                      item.changePercent >= 0
                                        ? 'bg-emerald-500/10 text-emerald-400'
                                        : 'bg-rose-500/10 text-rose-400'
                                    }`}
                                  >
                                    {item.changePercent >= 0 ? `+${item.changePercent}` : item.changePercent}%
                                  </span>
                                )}
                              </div>
                              <div className="text-[11px] text-slate-400 leading-tight">{item.producer}</div>
                              <div className="mt-1.5">
                                <span
                                  dir="ltr"
                                  className="inline-block font-mono text-[10px] text-slate-400 bg-slate-950 px-1.5 py-0.5 rounded border border-slate-800"
                                >
                                  {item.code}
                                </span>
                              </div>
                            </div>

                            <div className="flex items-center justify-between pt-2.5 mt-2.5 border-t border-slate-800/80">
                              <div className="flex items-baseline gap-1 text-emerald-400 font-bold">
                                <span className="font-mono text-sm">{formatNumber(displayPrice)}</span>
                                <span className="text-[10px] text-slate-300 font-normal">{activeCurrency}/kg</span>
                              </div>

                              {isSelected ? (
                                <span className="inline-flex items-center gap-1 text-[11px] font-bold text-blue-400 bg-blue-500/15 px-2 py-0.5 rounded-lg border border-blue-500/30">
                                  <Check className="w-3.5 h-3.5" />
                                  <span>پلیمر فعال</span>
                                </span>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => handleSelectMarketItem(item)}
                                  className="px-2 py-0.5 rounded-lg text-[11px] font-medium bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 transition-colors cursor-pointer"
                                >
                                  انتخاب نرخ
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                  </div>
                </div>

                {/* Section B: Masterbatch Additive */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-xs text-slate-300 font-bold border-b border-slate-800/80 pb-1">
                    <span className="flex items-center gap-1.5 text-amber-400">
                      <span className="w-2 h-2 rounded-full bg-amber-500 inline-block"></span>
                      <span>افزودنی مستربچ دوده (Carbon Black)</span>
                    </span>
                    <span className="text-[11px] text-slate-400 font-normal">
                      انتخاب برای «قیمت مستربچ دوده»
                    </span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    {marketSnapshot.items
                      .filter((item) => item.grade === 'Masterbatch')
                      .map((item) => {
                        const displayPrice = activeCurrency === 'ریال' ? item.pricePerKgToman * 10 : item.pricePerKgToman;
                        const isSelected = rawMasterbatchPrice === displayPrice;

                        return (
                          <div
                            key={item.id}
                            className={`p-3 rounded-xl border transition-all text-xs flex flex-col justify-between ${
                              isSelected
                                ? 'bg-amber-950/30 border-amber-500/50 ring-1 ring-amber-500/30'
                                : 'bg-slate-900/70 border-slate-800 hover:border-slate-700'
                            }`}
                          >
                            <div>
                              <div className="flex items-start justify-between gap-1 mb-1.5">
                                <span className="font-bold text-white text-xs leading-snug">{item.name}</span>
                                {item.changePercent !== undefined && (
                                  <span
                                    dir="ltr"
                                    className={`text-[10px] font-mono font-bold px-1.5 py-0.5 rounded inline-flex shrink-0 ${
                                      item.changePercent >= 0
                                        ? 'bg-emerald-500/10 text-emerald-400'
                                        : 'bg-rose-500/10 text-rose-400'
                                    }`}
                                  >
                                    {item.changePercent >= 0 ? `+${item.changePercent}` : item.changePercent}%
                                  </span>
                                )}
                              </div>
                              <div className="text-[11px] text-slate-400 leading-tight">{item.producer}</div>
                              <div className="mt-1.5">
                                <span
                                  dir="ltr"
                                  className="inline-block font-mono text-[10px] text-slate-400 bg-slate-950 px-1.5 py-0.5 rounded border border-slate-800"
                                >
                                  {item.code}
                                </span>
                              </div>
                            </div>

                            <div className="flex items-center justify-between pt-2.5 mt-2.5 border-t border-slate-800/80">
                              <div className="flex items-baseline gap-1 text-amber-400 font-bold">
                                <span className="font-mono text-sm">{formatNumber(displayPrice)}</span>
                                <span className="text-[10px] text-slate-300 font-normal">{activeCurrency}/kg</span>
                              </div>

                              {isSelected ? (
                                <span className="inline-flex items-center gap-1 text-[11px] font-bold text-amber-400 bg-amber-500/15 px-2 py-0.5 rounded-lg border border-amber-500/30">
                                  <Check className="w-3.5 h-3.5" />
                                  <span>مستربچ فعال</span>
                                </span>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => handleSelectMarketItem(item)}
                                  className="px-2 py-0.5 rounded-lg text-[11px] font-medium bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 transition-colors cursor-pointer"
                                >
                                  انتخاب نرخ
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                  </div>
                </div>
              </div>
            </div>

            {/* Inputs Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-300 mb-1 font-medium">قیمت پایه پلیمر بورس کالا ({activeCurrency}/kg):</label>
                <input
                  type="number"
                  value={rawPolymerPrice || ''}
                  onChange={(e) => setRawPolymerPrice(Number(e.target.value) || 0)}
                  placeholder="مثال: ۸۵۰۰۰"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">درصد مستربچ دوده (Carbon Black %):</label>
                <input
                  type="number"
                  step="0.1"
                  value={rawMasterbatchPct}
                  onChange={(e) => setRawMasterbatchPct(Number(e.target.value) || 0)}
                  placeholder="پیش‌فرض: ۲.۰٪ (پلیمر خودرنگ ۰)"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">قیمت هر کیلو مستربچ دوده ({activeCurrency}/kg):</label>
                <input
                  type="number"
                  value={rawMasterbatchPrice || ''}
                  onChange={(e) => setRawMasterbatchPrice(Number(e.target.value) || 0)}
                  placeholder="مثال: ۱۳۰۰۰۰"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">درصد افت و ضایعات تولید (Waste %):</label>
                <input
                  type="number"
                  step="0.1"
                  value={rawWastePct}
                  onChange={(e) => setRawWastePct(Number(e.target.value) || 0)}
                  placeholder="پیش‌فرض: ۱.۵٪"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">هزینه تبدیل و سربار کارخانه ({activeCurrency}/kg):</label>
                <input
                  type="number"
                  value={rawConversionCost || ''}
                  onChange={(e) => setRawConversionCost(Number(e.target.value) || 0)}
                  placeholder="برق، استهلاک خط، دستمزد، بسته‌بندی (مثال: ۱۰۰۰۰)"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>

              <div>
                <label className="block text-slate-300 mb-1 font-medium">درصد حاشیه سود تولیدکننده (Margin %):</label>
                <input
                  type="number"
                  step="0.5"
                  value={rawProfitMarginPct}
                  onChange={(e) => setRawProfitMarginPct(Number(e.target.value) || 0)}
                  placeholder="پیش‌فرض: ۱۰٪"
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none focus:border-emerald-500 font-bold"
                />
              </div>
            </div>

            {/* Live Cost Breakdown */}
            <div className="bg-slate-950/80 p-4 rounded-xl border border-slate-800 space-y-2 text-xs">
              <div className="flex justify-between items-center text-slate-400">
                <span>بهای مواد اولیه با احتساب دوده:</span>
                <span className="font-mono text-slate-200">{formatNumber(rawCostResult.baseMaterialCostPerKg)} {activeCurrency}</span>
              </div>
              <div className="flex justify-between items-center text-slate-400">
                <span>هزینه افت ضایعات خط اکسترودر:</span>
                <span className="font-mono text-slate-200">+ {formatNumber(rawCostResult.wasteCostPerKg)} {activeCurrency}</span>
              </div>
              <div className="flex justify-between items-center text-slate-400">
                <span>هزینه تبدیل، برق و سربار تولید:</span>
                <span className="font-mono text-slate-200">+ {formatNumber(rawCostResult.conversionCostPerKg)} {activeCurrency}</span>
              </div>
              <div className="flex justify-between items-center text-slate-300 font-semibold pt-1 border-t border-slate-800">
                <span>بهای تمام‌شده تولید (Production Cost):</span>
                <span className="font-mono text-blue-300">{formatNumber(rawCostResult.productionCostPerKg)} {activeCurrency}</span>
              </div>
              <div className="flex justify-between items-center text-slate-400">
                <span>سود ناخالص تولیدکننده:</span>
                <span className="font-mono text-slate-200">+ {formatNumber(rawCostResult.profitAmountPerKg)} {activeCurrency}</span>
              </div>
              <div className="flex justify-between items-center text-white font-bold text-sm pt-2 border-t-2 border-slate-800">
                <span>نرخ فروش پیشنهادی هر کیلوگرم:</span>
                <span className="font-mono text-emerald-400 text-base">{formatNumber(rawCostResult.suggestedPricePerKg)} {activeCurrency}</span>
              </div>
            </div>

            {/* Modal Actions */}
            <div className="flex items-center justify-end gap-3 pt-2">
              <button
                type="button"
                onClick={() => setShowCostModal(false)}
                className="px-4 py-2.5 rounded-xl text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
              >
                انصراف
              </button>

              <button
                type="button"
                onClick={handleApplyCostPrice}
                className="flex items-center gap-2 px-6 py-2.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg shadow-emerald-600/30 transition-all cursor-pointer"
              >
                <CheckCircle className="w-4 h-4" />
                <span>اعمال این نرخ ({formatNumber(rawCostResult.suggestedPricePerKg)} {activeCurrency})</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

