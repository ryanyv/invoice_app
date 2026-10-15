import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Plus,
  Trash2,
  Printer,
  Save,
  Calculator,
  Layers,
  CheckCircle,
  FileCheck,
  Edit3,
  Copy,
  RotateCcw,
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
} from '../../types';
import {
  computeInvoiceSummary,
  computePipeItem,
  getApplicableDiscount,
  getWeightPerMeter,
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
  const [pipeSdr, setPipeSdr] = useState<number>(11);
  const [pipePn, setPipePn] = useState<number>(16);
  const [pipeLength, setPipeLength] = useState<number>(100);
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

  // Live pipe total weight and total price
  const currentPipeWeight = currentWpm ? Number((currentWpm * pipeLength).toFixed(3)) : 0;
  const currentPipeTotalPrice = Math.round(currentPipeWeight * pipePricePerKg);

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
  const [includeCustomDiscount, setIncludeCustomDiscount] = useState<boolean>(false);
  const [customDiscountAmount, setCustomDiscountAmount] = useState<number>(0);
  const [includeTax, setIncludeTax] = useState<boolean>(false);
  const [taxPercentage] = useState<number>(settings.taxRate || 10);

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
      setExplanation(inv.explanation || '');
      setItems(inv.items ? JSON.parse(JSON.stringify(inv.items)) : []);
      setIncludeDiscount(Boolean(inv.includeDiscount));
      setIncludeCustomDiscount(Boolean(inv.includeCustomDiscount));
      setCustomDiscountAmount(inv.customDiscountAmount || 0);
      setIncludeTax(Boolean(inv.includeTax));
      setIsHydrated(true);
      if (editingSession.isDuplicate) {
        notify(`فاکتور شماره ${inv.invoiceNumber} به عنوان پیش‌فاکتور جدید تکثیر گردید.`);
      } else {
        notify(`فاکتور شماره ${inv.invoiceNumber} جهت ویرایش بارگذاری شد.`);
      }
    } else {
      // Check for saved draft in localStorage
      const draft = StorageService.getDraft();
      if (draft && ((draft.items && draft.items.length > 0) || (draft.customerName && draft.customerName.trim().length > 0) || draft.invoiceNumber)) {
        if (draft.invoiceNumber) setInvoiceNumber(draft.invoiceNumber);
        if (draft.jalaliDate) setJalaliDate(draft.jalaliDate);
        if (draft.customerName) setCustomerName(draft.customerName);
        if (draft.customerPhone) setCustomerPhone(draft.customerPhone);
        if (draft.customerAddress) setCustomerAddress(draft.customerAddress);
        if (draft.explanation !== undefined) setExplanation(draft.explanation);
        if (draft.items && Array.isArray(draft.items)) setItems(draft.items);
        if (draft.includeDiscount !== undefined) setIncludeDiscount(draft.includeDiscount);
        if (draft.includeCustomDiscount !== undefined) setIncludeCustomDiscount(draft.includeCustomDiscount);
        if (draft.customDiscountAmount !== undefined) setCustomDiscountAmount(draft.customDiscountAmount);
        if (draft.includeTax !== undefined) setIncludeTax(draft.includeTax);
        notify('پیش‌نویس ذخیره‌شده فاکتور بازیابی گردید.');
      } else {
        const nextNum = (settings.invoiceCounter || 1) + 1;
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
      customerName.trim().length > 0 ||
      customerPhone.trim().length > 0 ||
      customerAddress.trim().length > 0 ||
      explanation.trim().length > 0;

    if (hasData) {
      StorageService.saveDraft({
        invoiceNumber,
        jalaliDate,
        customerName,
        customerPhone,
        customerAddress,
        explanation,
        items,
        includeDiscount,
        includeCustomDiscount,
        customDiscountAmount,
        includeTax,
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
    explanation,
    includeDiscount,
    includeCustomDiscount,
    customDiscountAmount,
    includeTax,
    invoiceNumber,
    jalaliDate,
    activeCurrency,
  ]);

  // Inline edit for pipe item: length and pricePerKg
  const handleUpdatePipeItem = (id: string, newLength: number, newPricePerKg: number) => {
    setItems((prev) =>
      prev.map((item) => {
        if (item.id === id && item.type === 'pipe') {
          const updated = computePipeItem(
            {
              peGrade: item.peGrade,
              diameter: item.diameter,
              sdr: item.sdr,
              pn: item.pn,
              length: newLength,
              pricePerKg: newPricePerKg,
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
  const subtotalBeforeDiscounts = items.reduce((sum, item) => sum + (item.totalPrice || 0), 0);
  const autoDiscountPct = getApplicableDiscount(subtotalBeforeDiscounts, discounts);

  const summary = useMemo(() => {
    return computeInvoiceSummary({
      items,
      includeDiscount,
      discountPercentage: autoDiscountPct,
      includeCustomDiscount,
      customDiscountAmount,
      includeTax,
      taxPercentage,
    });
  }, [
    items,
    includeDiscount,
    autoDiscountPct,
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
      (c) => c.name.trim().toLowerCase() === name.trim().toLowerCase()
    );
    if (matched) {
      if (matched.phone) setCustomerPhone(matched.phone);
      if (matched.address) setCustomerAddress(matched.address);
      if (matched.defaultDiscount && matched.defaultDiscount > 0) {
        setIncludeCustomDiscount(true);
        setIncludeDiscount(false);
        if (matched.defaultDiscount <= 100) {
          const discountVal = Math.round((subtotalBeforeDiscounts * matched.defaultDiscount) / 100);
          setCustomDiscountAmount(discountVal);
          notify(`اطلاعات و درصد تخفیف ${toPersianDigits(matched.defaultDiscount)}٪ مشتری اعمال شد.`);
        } else {
          setCustomDiscountAmount(matched.defaultDiscount);
          notify(`اطلاعات و تخفیف ${formatNumber(matched.defaultDiscount)} ${activeCurrency} مشتری اعمال شد.`);
        }
      } else {
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
      },
      weightTable
    );

    if (item) {
      setItems([...items, item]);
      notify(`لوله سایز ${pipeDiameter} با گرید ${pipeGrade} افزوده شد.`);
    }
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
    setIncludeDiscount(false);
    setIncludeCustomDiscount(false);
    setCustomDiscountAmount(0);
    setIncludeTax(false);
    StorageService.clearDraft();
    onClearEditingSession?.();
    const nextNum = (settings.invoiceCounter || 1) + 1;
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
      jalaliDate,
      items,
      subtotal: summary.subtotal,
      includeDiscount,
      discountPercentage: autoDiscountPct,
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

          <div className="flex items-center gap-3">
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
          <form onSubmit={handleAddPipeItem} className="p-6 space-y-5">
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
                      setPipeSdr(sdrs[0]);
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
                <label className="block text-xs text-slate-400 mb-1">
                  قیمت هر کیلوگرم ({activeCurrency}):
                </label>
                <input
                  type="number"
                  value={pipePricePerKg || ''}
                  onChange={(e) => setPipePricePerKg(Number(e.target.value) || 0)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-emerald-400 font-mono font-bold text-left focus:outline-none"
                />
              </div>
            </div>

            {/* Calculations Preview Bar */}
            <div className="bg-slate-900/70 p-4 rounded-xl border border-slate-700/70 flex flex-wrap items-center justify-between gap-4 text-xs">
              <div className="flex items-center gap-6">
                <div>
                  <span className="text-slate-400 block">وزن هر متر (DIN):</span>
                  <strong className="text-slate-200 text-sm font-mono">
                    {currentWpm ? `${currentWpm} kg/m` : 'نامعتبر'}
                  </strong>
                </div>

                <div>
                  <span className="text-slate-400 block">وزن کل لوله:</span>
                  <strong className="text-blue-400 text-sm font-mono">
                    {formatNumber(currentPipeWeight)} کیلوگرم
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
                <th className="py-3 px-4 text-center">قیمت واحد ({activeCurrency})</th>
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
                            <span className="font-bold text-white">لوله قطر {item.diameter} mm</span>
                            <span className="text-slate-400 text-xs mr-2">
                              (گرید {item.peGrade} | SDR {item.sdr} | PN {item.pn})
                            </span>
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
        {/* Notes & Explanations */}
        <div className="lg:col-span-2 bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-4">
          <label className="block text-sm font-bold text-white">توضیحات و شرایط پیش‌فاکتور (اختیاری):</label>
          <textarea
            rows={4}
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
                  <span>تخفیف پله‌ای خودکار ({autoDiscountPct}%):</span>
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
    </div>
  );
};

