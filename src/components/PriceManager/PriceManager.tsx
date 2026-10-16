import React, { useState, useMemo } from 'react';
import {
  Search,
  Plus,
  TrendingUp,
  Download,
  Upload,
  Save,
  Check,
  Edit2,
  Trash2,
  RotateCcw,
  Percent,
  Layers,
  Wifi,
  WifiOff,
  RefreshCw,
  Users,
} from 'lucide-react';
import { CompanySettings, Customer, DiscountRule, FittingCatalogItem, PipeGrade } from '../../types';
import { StorageService } from '../../services/storage';
import { MarketPriceService } from '../../services/marketPriceService';
import { formatNumber, formatPrice, toPersianDigits } from '../../services/persianUtils';

interface PriceManagerProps {
  settings: CompanySettings;
  onUpdateSettings: (settings: CompanySettings) => void;
  fittings: FittingCatalogItem[];
  onUpdateFittings: (fittings: FittingCatalogItem[]) => void;
  discounts: DiscountRule[];
  onUpdateDiscounts: (discounts: DiscountRule[]) => void;
  customers?: Customer[];
  onUpdateCustomers?: (customers: Customer[]) => void;
}

export const PriceManager: React.FC<PriceManagerProps> = ({
  settings,
  onUpdateSettings,
  fittings,
  onUpdateFittings,
  discounts,
  onUpdateDiscounts,
  customers = [],
  onUpdateCustomers,
}) => {
  const [activeSubTab, setActiveSubTab] = useState<'pipes' | 'fittings' | 'discounts' | 'customers'>('fittings');

  // Pipe prices state
  const [pipePrices, setPipePrices] = useState<Record<PipeGrade, number>>(() => ({
    PE100: settings?.basePipePrices?.PE100 || 95000,
    PE80: settings?.basePipePrices?.PE80 || 88000,
    PE63: settings?.basePipePrices?.PE63 || 82000,
  }));
  const [pipeSavedNotice, setPipeSavedNotice] = useState(false);

  // Fittings filtering & search
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedType, setSelectedType] = useState<string>('all');
  const [selectedProduct, setSelectedProduct] = useState<string>('all');

  // Inline editing state
  const [editingId, setEditingId] = useState<string | null>(null);
  const [tempPriceRial, setTempPriceRial] = useState<number>(0);

  // Bulk update modal state
  const [showBulkModal, setShowBulkModal] = useState(false);
  const [bulkPercent, setBulkPercent] = useState<number>(10);
  const [bulkTarget, setBulkTarget] = useState<'filtered' | 'all'>('filtered');

  // Add new fitting modal state
  const [showAddModal, setShowAddModal] = useState(false);
  const [newFitting, setNewFitting] = useState<Partial<FittingCatalogItem>>({
    category: 'اتصالات فشار قوی',
    connectionType: 'جوشی',
    product: '',
    material: 'HDPE100',
    pressure: '10',
    size: '63',
    unitPriceRial: 100000,
    packageCount: 10,
    isActive: true,
  });

  // Notification banner
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  // Market live price sync state
  const [isFetchingMarket, setIsFetchingMarket] = useState(false);
  const [marketSnapshot, setMarketSnapshot] = useState(() => MarketPriceService.getStoredMarketPrices());

  const handleSyncIMEPrices = async () => {
    setIsFetchingMarket(true);
    try {
      const snapshot = await MarketPriceService.fetchLiveMarketPrices();
      setMarketSnapshot(snapshot);
      const newPrices = { ...pipePrices };
      for (const item of snapshot.items) {
        if (item.grade === 'PE100' && item.id.includes('black')) {
          newPrices.PE100 = item.pricePerKgToman;
        } else if (item.grade === 'PE80') {
          newPrices.PE80 = item.pricePerKgToman;
        } else if (item.grade === 'PE63') {
          newPrices.PE63 = item.pricePerKgToman;
        }
      }
      setPipePrices(newPrices);
      const updatedSettings = {
        ...settings,
        basePipePrices: newPrices,
      };
      onUpdateSettings(updatedSettings);
      StorageService.saveSettings(updatedSettings);
      showToast(
        snapshot.isOnline
          ? '✓ نرخ‌های رسمی بورس کالا به صورت آنلاین دریافت و اعمال گردید.'
          : 'حالت آفلاین: آخرین نرخ رسمی مصوب پتروشیمی اعمال گردید.'
      );
    } catch {
      showToast('خطا در استعلام آنلاین؛ از نرخ‌های آفلاین استفاده شد.');
    } finally {
      setIsFetchingMarket(false);
    }
  };

  // Save pipe base prices
  const handleSavePipePrices = () => {
    const updated = {
      ...settings,
      basePipePrices: pipePrices,
    };
    onUpdateSettings(updated);
    StorageService.saveSettings(updated);
    setPipeSavedNotice(true);
    setTimeout(() => setPipeSavedNotice(false), 2500);
    showToast('قیمت‌های پایه لوله با موفقیت ذخیره شدند.');
  };

  // Unique filter lists
  const categories = useMemo(() => {
    const set = new Set(fittings.map((f) => f.category).filter(Boolean));
    return Array.from(set);
  }, [fittings]);

  const connectionTypes = useMemo(() => {
    let list = fittings;
    if (selectedCategory !== 'all') {
      list = list.filter((f) => f.category === selectedCategory);
    }
    const set = new Set(list.map((f) => f.connectionType).filter(Boolean));
    return Array.from(set);
  }, [fittings, selectedCategory]);

  const products = useMemo(() => {
    let list = fittings;
    if (selectedCategory !== 'all') {
      list = list.filter((f) => f.category === selectedCategory);
    }
    if (selectedType !== 'all') {
      list = list.filter((f) => f.connectionType === selectedType);
    }
    const set = new Set(list.map((f) => f.product).filter(Boolean));
    return Array.from(set);
  }, [fittings, selectedCategory, selectedType]);

  // Filtered fittings list
  const filteredFittings = useMemo(() => {
    return fittings.filter((item) => {
      if (selectedCategory !== 'all' && item.category !== selectedCategory) return false;
      if (selectedType !== 'all' && item.connectionType !== selectedType) return false;
      if (selectedProduct !== 'all' && item.product !== selectedProduct) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const text = `${item.category} ${item.connectionType} ${item.product} ${item.material} ${item.pressure} ${item.size}`.toLowerCase();
        if (!text.includes(q)) return false;
      }

      return true;
    });
  }, [fittings, selectedCategory, selectedType, selectedProduct, searchQuery]);

  // Handle inline price update
  const startEditing = (item: FittingCatalogItem) => {
    setEditingId(item.id);
    setTempPriceRial(item.unitPriceRial);
  };

  const saveInlinePrice = (id: string) => {
    const updated = fittings.map((f) =>
      f.id === id ? { ...f, unitPriceRial: Math.max(0, tempPriceRial) } : f
    );
    onUpdateFittings(updated);
    StorageService.saveFittings(updated);
    setEditingId(null);
    showToast('قیمت بروزرسانی شد.');
  };

  // Toggle active status
  const toggleItemActive = (id: string) => {
    const updated = fittings.map((f) =>
      f.id === id ? { ...f, isActive: !f.isActive } : f
    );
    onUpdateFittings(updated);
    StorageService.saveFittings(updated);
  };

  // Delete fitting
  const handleDeleteFitting = (id: string) => {
    if (confirm('آیا از حذف این کالا از کاتالوگ اطمینان دارید؟')) {
      const updated = fittings.filter((f) => f.id !== id);
      onUpdateFittings(updated);
      StorageService.saveFittings(updated);
      showToast('کالا حذف شد.');
    }
  };

  // Bulk price update
  const executeBulkUpdate = () => {
    const factor = 1 + bulkPercent / 100;
    const targetIds = new Set(
      bulkTarget === 'filtered' ? filteredFittings.map((f) => f.id) : fittings.map((f) => f.id)
    );

    const updated = fittings.map((f) => {
      if (targetIds.has(f.id)) {
        return {
          ...f,
          unitPriceRial: Math.round(f.unitPriceRial * factor),
        };
      }
      return f;
    });

    onUpdateFittings(updated);
    StorageService.saveFittings(updated);
    setShowBulkModal(false);
    showToast(`قیمت ${targetIds.size} کالا به میزان ${bulkPercent}% بروزرسانی شد.`);
  };

  // Add new fitting submit
  const handleAddNewFitting = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFitting.product || !newFitting.size) {
      alert('لطفاً عنوان محصول و سایز را وارد کنید.');
      return;
    }

    const item: FittingCatalogItem = {
      id: 'conn_custom_' + Date.now(),
      category: newFitting.category || 'اتصالات فشار قوی',
      connectionType: newFitting.connectionType || 'جوشی',
      product: newFitting.product,
      material: newFitting.material || 'HDPE100',
      pressure: newFitting.pressure || '10',
      size: newFitting.size,
      unitPriceRial: Number(newFitting.unitPriceRial) || 0,
      packageCount: Number(newFitting.packageCount) || 1,
      isActive: true,
    };

    const updated = [item, ...fittings];
    onUpdateFittings(updated);
    StorageService.saveFittings(updated);
    setShowAddModal(false);
    showToast('کالای جدید به کاتالوگ افزوده شد.');
  };

  // Customer form state
  const [newCustomer, setNewCustomer] = useState<Partial<Customer>>({
    name: '',
    phone: '',
    address: '',
    economicCode: '',
    nationalId: '',
    postalCode: '',
    defaultDiscount: 0,
  });

  const handleSaveCustomer = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCustomer.name?.trim()) return;
    const cust: Customer = {
      id: 'cust_' + Date.now(),
      name: newCustomer.name.trim(),
      phone: newCustomer.phone?.trim() || undefined,
      address: newCustomer.address?.trim() || undefined,
      economicCode: newCustomer.economicCode?.trim() || undefined,
      nationalId: newCustomer.nationalId?.trim() || undefined,
      postalCode: newCustomer.postalCode?.trim() || undefined,
      defaultDiscount: Number(newCustomer.defaultDiscount) || 0,
    };
    StorageService.addOrUpdateCustomer(cust);
    onUpdateCustomers?.(StorageService.getCustomers());
    setNewCustomer({
      name: '',
      phone: '',
      address: '',
      economicCode: '',
      nationalId: '',
      postalCode: '',
      defaultDiscount: 0,
    });
    showToast('اطلاعات مشتری با موفقیت ذخیره شد.');
  };

  const handleDeleteCustomer = (id: string) => {
    const updated = customers.filter((c) => c.id !== id);
    StorageService.saveCustomers(updated);
    onUpdateCustomers?.(updated);
    showToast('مشتری از فهرست حذف شد.');
  };

  // Export CSV
  const handleExportCsv = () => {
    const csvContent = StorageService.exportFittingsToCsv();
    const defaultName = `fittings_catalog_${Date.now()}.csv`;
    if (typeof window !== 'undefined' && window.electronAPI?.exportFile) {
      window.electronAPI
        .exportFile({ defaultName, content: csvContent, extensions: ['csv'] })
        .then((res) => {
          if (res?.success) showToast('فایل CSV با موفقیت ذخیره شد.');
        });
      return;
    }
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', defaultName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('فایل CSV با موفقیت دانلود شد.');
  };

  // RFC 4180 CSV line parser supporting quoted fields with commas and escaped quotes
  const parseCsvRow = (line: string): string[] => {
    const cells: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          cells.push(current.trim());
          current = '';
        } else {
          current += ch;
        }
      }
    }
    cells.push(current.trim());
    return cells;
  };

  // Import CSV
  const handleImportCsv = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const text = evt.target?.result as string;
        const lines = text.replace(/^\uFEFF/, '').trim().split(/\r?\n/);
        if (lines.length < 2) {
          alert('فرمت فایل نامعتبر است.');
          return;
        }

        const imported: FittingCatalogItem[] = [];
        for (let i = 1; i < lines.length; i++) {
          if (!lines[i].trim()) continue;
          const row = parseCsvRow(lines[i]);
          if (row.length >= 8) {
            imported.push({
              id: row[0] || 'conn_' + i + '_' + Date.now(),
              category: row[1] || 'اتصالات',
              connectionType: row[2] || '',
              product: row[3] || '',
              material: row[4] || '',
              pressure: row[5] || '',
              size: row[6] || '',
              unitPriceRial: Math.max(0, Number(row[7]) || 0),
              packageCount: Math.max(1, Number(row[8]) || 1),
              isActive: row[9] !== '0',
            });
          }
        }

        if (imported.length > 0) {
          onUpdateFittings(imported);
          StorageService.saveFittings(imported);
          showToast(`${imported.length} رکورد از فایل CSV با موفقیت بارگذاری شد.`);
        }
      } catch (err) {
        alert('خطا در خواندن فایل CSV.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Toast banner */}
      {toastMessage && (
        <div className="fixed bottom-6 left-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-xl shadow-2xl flex items-center gap-3 animate-bounce">
          <Check className="w-5 h-5" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Header and Sub Tabs */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-800/80 p-6 rounded-2xl border border-slate-700">
        <div>
          <h1 className="text-2xl font-bold text-white flex items-center gap-3">
            <Layers className="w-7 h-7 text-blue-400" />
            <span>مدیریت قیمت‌ها، کاتالوگ و مشتریان</span>
          </h1>
          <p className="text-slate-400 text-sm mt-1">
            ویرایش مستقیم قیمت پایه لوله‌ها، کاتالوگ اتصالات، پله‌های تخفیف و مشتریان دائمی
          </p>
        </div>

        {/* Tab switchers */}
        <div className="flex flex-wrap items-center gap-2 bg-slate-900/80 p-1.5 rounded-xl border border-slate-700/60">
          <button
            onClick={() => setActiveSubTab('fittings')}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-all cursor-pointer ${
              activeSubTab === 'fittings'
                ? 'bg-blue-600 text-white shadow'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            کاتالوگ اتصالات ({toPersianDigits(fittings.length)})
          </button>

          <button
            onClick={() => setActiveSubTab('pipes')}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-all cursor-pointer ${
              activeSubTab === 'pipes'
                ? 'bg-blue-600 text-white shadow'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            قیمت پایه مواد اولیه لوله
          </button>

          <button
            onClick={() => setActiveSubTab('discounts')}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-all cursor-pointer ${
              activeSubTab === 'discounts'
                ? 'bg-blue-600 text-white shadow'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            پله‌های تخفیف خودکار
          </button>

          <button
            onClick={() => setActiveSubTab('customers')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium transition-all cursor-pointer ${
              activeSubTab === 'customers'
                ? 'bg-blue-600 text-white shadow'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            <Users className="w-4 h-4" />
            <span>مشتریان ({toPersianDigits(customers.length)})</span>
          </button>
        </div>
      </div>

      {/* SUB-TAB 1: PIPE BASE PRICES */}
      {activeSubTab === 'pipes' && (
        <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-6">
          <div className="border-b border-slate-700 pb-4">
            <h2 className="text-lg font-bold text-white">نرخ پایه مواد اولیه لوله پلی‌اتیلن (هر کیلوگرم)</h2>
            <p className="text-slate-400 text-sm mt-1">
              هنگام افزودن لوله به پیش‌فاکتور، نرخ پیش‌فرض هر کیلوگرم از این جدول خوانده می‌شود.
            </p>
          </div>

          {/* IME Market Price Live Sync Card */}
          <div className="bg-slate-900/90 p-4 rounded-xl border border-slate-700/80 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className={`p-2 rounded-xl border flex items-center justify-center ${
                marketSnapshot.isOnline
                  ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30'
                  : 'bg-amber-500/20 text-amber-400 border-amber-500/30'
              }`}>
                {marketSnapshot.isOnline ? <Wifi className="w-5 h-5" /> : <WifiOff className="w-5 h-5" />}
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-bold text-white">همگام‌سازی با نرخ رسمی بورس کالای ایران (IME)</span>
                  <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium border ${
                    marketSnapshot.isOnline
                      ? 'bg-emerald-500/10 text-emerald-300 border-emerald-500/20'
                      : 'bg-amber-500/10 text-amber-300 border-amber-500/20'
                  }`}>
                    {marketSnapshot.isOnline ? '🟢 آنلاین (متصل به بازار)' : '🟡 آفلاین (آخرین نرخ رسمی مصوب)'}
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-0.5">
                  مرجع: {marketSnapshot.source} (تاریخ: {marketSnapshot.lastUpdated})
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={handleSyncIMEPrices}
              disabled={isFetchingMarket}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-60 text-white rounded-xl text-xs font-bold transition-all shadow-md cursor-pointer whitespace-nowrap"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isFetchingMarket ? 'animate-spin' : ''}`} />
              <span>{isFetchingMarket ? 'در حال استعلام...' : 'دریافت آنلاین و همگام‌سازی نرخ بورس'}</span>
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {(['PE100', 'PE80', 'PE63'] as PipeGrade[]).map((grade) => (
              <div
                key={grade}
                className="bg-slate-900/60 p-5 rounded-xl border border-slate-700/70 hover:border-blue-500/50 transition-all space-y-3"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-lg text-blue-400">{grade}</span>
                  <span className="text-xs px-2 py-1 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20">
                    پلی‌اتیلن گرید {grade}
                  </span>
                </div>

                <div className="space-y-1">
                  <label className="text-xs text-slate-400">قیمت هر کیلوگرم (تومان):</label>
                  <div className="relative">
                    <input
                      type="number"
                      value={pipePrices[grade] || ''}
                      onChange={(e) =>
                        setPipePrices({
                          ...pipePrices,
                          [grade]: Number(e.target.value) || 0,
                        })
                      }
                      className="w-full bg-slate-800 border border-slate-600 rounded-lg px-3 py-2 text-white font-mono text-left focus:ring-2 focus:ring-blue-500 focus:outline-none"
                    />
                    <span className="absolute right-3 top-2.5 text-xs text-slate-400 pointer-events-none">
                      تومان
                    </span>
                  </div>
                </div>

                <div className="text-xs text-slate-400 pt-1">
                  معادل ریالی: {formatNumber((pipePrices[grade] || 0) * 10)} ریال
                </div>
              </div>
            ))}
          </div>

          <div className="flex items-center justify-between pt-4 border-t border-slate-700">
            <span className="text-sm text-slate-400">
              {pipeSavedNotice ? '✓ قیمت‌ها ذخیره شدند' : 'پس از تغییر، روی ذخیره کلیک کنید.'}
            </span>
            <button
              onClick={handleSavePipePrices}
              className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl font-medium shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
            >
              <Save className="w-5 h-5" />
              <span>ذخیره تغییرات قیمت لوله</span>
            </button>
          </div>
        </div>
      )}

      {/* SUB-TAB 2: FITTINGS CATALOG & INLINE EDITING */}
      {activeSubTab === 'fittings' && (
        <div className="space-y-4">
          {/* Controls: Search, Filters, Bulk, Add, Import/Export */}
          <div className="bg-slate-800/80 p-5 rounded-2xl border border-slate-700 space-y-4">
            <div className="flex flex-col lg:flex-row items-center justify-between gap-4">
              {/* Search bar */}
              <div className="relative w-full lg:w-96">
                <Search className="w-5 h-5 absolute right-3 top-3 text-slate-400" />
                <input
                  type="text"
                  placeholder="جستجو در محصول، نوع، اندازه، فشار یا گرید..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl pr-10 pl-4 py-2.5 text-sm text-white placeholder-slate-400 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              {/* Action Buttons: Add, Bulk Update, CSV Export/Import */}
              <div className="flex flex-wrap items-center gap-2 w-full lg:w-auto justify-end">
                <button
                  onClick={() => setShowAddModal(true)}
                  className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-sm font-medium transition-all shadow cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>افزودن اتصال جدید</span>
                </button>

                <button
                  onClick={() => setShowBulkModal(true)}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-medium transition-all shadow cursor-pointer"
                >
                  <TrendingUp className="w-4 h-4" />
                  <span>تغییر درصدی قیمت‌ها</span>
                </button>

                <button
                  onClick={handleExportCsv}
                  className="flex items-center gap-2 px-3 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-sm transition-all cursor-pointer"
                  title="خروجی فایل اکسل / CSV"
                >
                  <Download className="w-4 h-4" />
                  <span>خروجی اکسل</span>
                </button>

                <label className="flex items-center gap-2 px-3 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-sm transition-all cursor-pointer">
                  <Upload className="w-4 h-4" />
                  <span>بارگذاری CSV</span>
                  <input
                    type="file"
                    accept=".csv"
                    onChange={handleImportCsv}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {/* Filter Dropdowns */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-3 border-t border-slate-700/60">
              <div>
                <label className="block text-xs text-slate-400 mb-1">دسته‌بندی (نوع):</label>
                <select
                  value={selectedCategory}
                  onChange={(e) => {
                    setSelectedCategory(e.target.value);
                    setSelectedType('all');
                    setSelectedProduct('all');
                  }}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none"
                >
                  <option value="all">تمام دسته‌بندی‌ها</option>
                  {categories.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">نوع محصول:</label>
                <select
                  value={selectedType}
                  onChange={(e) => {
                    setSelectedType(e.target.value);
                    setSelectedProduct('all');
                  }}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none"
                >
                  <option value="all">تمام انواع</option>
                  {connectionTypes.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">نام محصول:</label>
                <select
                  value={selectedProduct}
                  onChange={(e) => setSelectedProduct(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none"
                >
                  <option value="all">تمام محصولات</option>
                  {products.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Fittings Data Table */}
          <div className="bg-slate-800/80 rounded-2xl border border-slate-700 overflow-hidden shadow-xl">
            <div className="px-6 py-4 bg-slate-800 border-b border-slate-700 flex items-center justify-between">
              <span className="text-sm font-medium text-slate-300">
                تعداد کالاهای مطابق با فیلتر: <strong className="text-blue-400">{toPersianDigits(filteredFittings.length)}</strong>
              </span>
              <span className="text-xs text-slate-400">
                💡 برای ویرایش سریع، روی قیمت هر ردیف کلیک کنید
              </span>
            </div>

            <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
              <table className="w-full text-right border-collapse text-sm">
                <thead className="bg-slate-900/90 sticky top-0 z-10 text-xs font-semibold text-slate-400 border-b border-slate-700">
                  <tr>
                    <th className="py-3 px-4">ردیف</th>
                    <th className="py-3 px-4">دسته‌بندی / نوع</th>
                    <th className="py-3 px-4">نام محصول</th>
                    <th className="py-3 px-4">سایز (mm)</th>
                    <th className="py-3 px-4">فشار (bar)</th>
                    <th className="py-3 px-4">جنس</th>
                    <th className="py-3 px-4">تعداد/بسته</th>
                    <th className="py-3 px-4 text-center">قیمت واحد (ریال / تومان)</th>
                    <th className="py-3 px-4 text-center">وضعیت</th>
                    <th className="py-3 px-4 text-center">عملیات</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/60">
                  {filteredFittings.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="py-12 text-center text-slate-500">
                        موردی با مشخصات جستجو شده یافت نشد.
                      </td>
                    </tr>
                  ) : (
                    filteredFittings.map((item, index) => {
                      const isEditing = editingId === item.id;
                      return (
                        <tr
                          key={item.id}
                          className={`hover:bg-slate-700/30 transition-colors ${
                            !item.isActive ? 'opacity-50 bg-slate-900/20' : ''
                          }`}
                        >
                          <td className="py-3 px-4 text-slate-400 font-mono text-xs">
                            {toPersianDigits(index + 1)}
                          </td>
                          <td className="py-3 px-4 text-slate-300">
                            <div>{item.category}</div>
                            <div className="text-xs text-slate-400">{item.connectionType}</div>
                          </td>
                          <td className="py-3 px-4 font-medium text-white">{item.product}</td>
                          <td className="py-3 px-4 text-blue-400 font-bold">{toPersianDigits(item.size)}</td>
                          <td className="py-3 px-4 text-slate-300 font-mono">{toPersianDigits(item.pressure)}</td>
                          <td className="py-3 px-4 text-slate-400 text-xs">{item.material}</td>
                          <td className="py-3 px-4 text-slate-400 font-mono text-xs">
                            {toPersianDigits(item.packageCount)}
                          </td>

                          {/* Editable Price Column */}
                          <td className="py-3 px-4 text-center">
                            {isEditing ? (
                              <div className="flex items-center justify-center gap-1">
                                <input
                                  type="number"
                                  autoFocus
                                  value={tempPriceRial}
                                  onChange={(e) => setTempPriceRial(Number(e.target.value) || 0)}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') saveInlinePrice(item.id);
                                    if (e.key === 'Escape') setEditingId(null);
                                  }}
                                  className="w-32 bg-slate-900 border border-blue-500 rounded px-2 py-1 text-white font-mono text-sm text-left focus:outline-none"
                                />
                                <button
                                  onClick={() => saveInlinePrice(item.id)}
                                  className="p-1 rounded bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer"
                                  title="تایید"
                                >
                                  <Check className="w-4 h-4" />
                                </button>
                                <button
                                  onClick={() => setEditingId(null)}
                                  className="p-1 rounded bg-slate-600 hover:bg-slate-500 text-white cursor-pointer"
                                  title="لغو"
                                >
                                  <RotateCcw className="w-4 h-4" />
                                </button>
                              </div>
                            ) : (
                              <button
                                onClick={() => startEditing(item)}
                                className="group inline-flex items-center gap-2 px-3 py-1 rounded-lg hover:bg-slate-700 border border-transparent hover:border-slate-600 transition-all cursor-pointer"
                                title="برای ویرایش کلیک کنید"
                              >
                                <div className="text-left font-mono">
                                  <div className="text-emerald-400 font-bold">
                                    {formatNumber(item.unitPriceRial)} ریال
                                  </div>
                                  <div className="text-xs text-slate-400">
                                    ({formatNumber(Math.round(item.unitPriceRial / 10))} تومان)
                                  </div>
                                </div>
                                <Edit2 className="w-3.5 h-3.5 text-slate-500 group-hover:text-blue-400 opacity-0 group-hover:opacity-100 transition-opacity" />
                              </button>
                            )}
                          </td>

                          {/* Active / Inactive Toggle */}
                          <td className="py-3 px-4 text-center">
                            <button
                              onClick={() => toggleItemActive(item.id)}
                              className={`text-xs px-2.5 py-1 rounded-full font-medium transition-colors cursor-pointer ${
                                item.isActive
                                  ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                                  : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'
                              }`}
                            >
                              {item.isActive ? 'فعال' : 'غیرفعال'}
                            </button>
                          </td>

                          {/* Delete Action */}
                          <td className="py-3 px-4 text-center">
                            <button
                              onClick={() => handleDeleteFitting(item.id)}
                              className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 transition-colors cursor-pointer"
                              title="حذف کالا"
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
        </div>
      )}

      {/* SUB-TAB 3: DISCOUNT RULES */}
      {activeSubTab === 'discounts' && (
        <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-6">
          <div className="flex items-center justify-between border-b border-slate-700 pb-4">
            <div>
              <h2 className="text-lg font-bold text-white">پله‌های محاسبه تخفیف خودکار پیش‌فاکتور</h2>
              <p className="text-slate-400 text-sm mt-1">
                بر اساس مبلغ کل خرید مشتری، درصد تخفیف به صورت خودکار پیشنهاد داده می‌شود.
              </p>
            </div>
            <button
              onClick={() => {
                const newRule: DiscountRule = { threshold: 10000000, percentage: 5 };
                const updated = [...discounts, newRule].sort((a, b) => a.threshold - b.threshold);
                onUpdateDiscounts(updated);
                StorageService.saveDiscounts(updated);
              }}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-medium cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              <span>افزودن پله تخفیف</span>
            </button>
          </div>

          <div className="space-y-3">
            {discounts.map((rule, idx) => (
              <div
                key={idx}
                className="flex items-center gap-4 bg-slate-900/60 p-4 rounded-xl border border-slate-700/60"
              >
                <div className="flex-1">
                  <label className="text-xs text-slate-400 block mb-1">از مبلغ سفارش حداقل (تومان):</label>
                  <input
                    type="number"
                    value={rule.threshold}
                    onChange={(e) => {
                      const updated = [...discounts];
                      updated[idx].threshold = Number(e.target.value) || 0;
                      onUpdateDiscounts(updated);
                      StorageService.saveDiscounts(updated);
                    }}
                    className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                  />
                  <div className="text-xs text-slate-400 mt-1 font-mono">
                    {formatNumber(rule.threshold)} تومان
                  </div>
                </div>

                <div className="w-40">
                  <label className="text-xs text-slate-400 block mb-1">درصد تخفیف (%):</label>
                  <div className="relative">
                    <input
                      type="number"
                      value={rule.percentage}
                      onChange={(e) => {
                        const updated = [...discounts];
                        updated[idx].percentage = Number(e.target.value) || 0;
                        onUpdateDiscounts(updated);
                        StorageService.saveDiscounts(updated);
                      }}
                      className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                    />
                    <Percent className="w-4 h-4 absolute right-3 top-3 text-slate-400" />
                  </div>
                </div>

                <button
                  onClick={() => {
                    const updated = discounts.filter((_, i) => i !== idx);
                    onUpdateDiscounts(updated);
                    StorageService.saveDiscounts(updated);
                  }}
                  className="mt-5 p-2 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 cursor-pointer"
                  title="حذف پله"
                >
                  <Trash2 className="w-5 h-5" />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* SUB-TAB 4: CUSTOMERS MANAGEMENT */}
      {activeSubTab === 'customers' && (
        <div className="bg-slate-800/80 p-6 rounded-2xl border border-slate-700 space-y-6">
          <div className="border-b border-slate-700 pb-4">
            <h2 className="text-lg font-bold text-white">مدیریت مشتریان، مشخصات رسمی و تخفیف اختصاصی</h2>
            <p className="text-slate-400 text-sm mt-1">
              مشتریان ثبت‌شده در این بخش هنگام تایپ نام در فرم صدور پیش‌فاکتور به صورت خودکار تکمیل می‌شوند.
            </p>
          </div>

          <form onSubmit={handleSaveCustomer} className="bg-slate-900/60 p-4 rounded-xl border border-slate-700/70 space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1">نام مشتری / شرکت (الزامی):</label>
                <input
                  type="text"
                  required
                  value={newCustomer.name || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, name: e.target.value })}
                  placeholder="مثال: شرکت آب و فاضلاب"
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-slate-400 mb-1">شماره تماس:</label>
                <input
                  type="text"
                  value={newCustomer.phone || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, phone: e.target.value })}
                  placeholder="۰۹۱۲..."
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-slate-400 mb-1">تخفیف پیش‌فرض (% یا مبلغ):</label>
                <input
                  type="number"
                  value={newCustomer.defaultDiscount || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, defaultDiscount: Number(e.target.value) || 0 })}
                  placeholder="مثال: 5 (درصد)"
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-slate-400 mb-1">شماره اقتصادی (ماده ۱۶۹):</label>
                <input
                  type="text"
                  value={newCustomer.economicCode || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, economicCode: e.target.value })}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-slate-400 mb-1">شناسه ملی / کد ملی:</label>
                <input
                  type="text"
                  value={newCustomer.nationalId || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, nationalId: e.target.value })}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-slate-400 mb-1">کد پستی ۱۰ رقمی:</label>
                <input
                  type="text"
                  value={newCustomer.postalCode || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, postalCode: e.target.value })}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
              <div className="sm:col-span-2">
                <label className="block text-slate-400 mb-1">نشانی کامل:</label>
                <input
                  type="text"
                  value={newCustomer.address || ''}
                  onChange={(e) => setNewCustomer({ ...newCustomer, address: e.target.value })}
                  placeholder="آدرس تحویل..."
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>
            </div>
            <div className="flex justify-end">
              <button
                type="submit"
                className="flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>افزودن / بروزرسانی مشتری</span>
              </button>
            </div>
          </form>

          <div className="overflow-x-auto">
            <table className="w-full text-right text-sm">
              <thead className="bg-slate-900/90 text-xs text-slate-400 border-b border-slate-700">
                <tr>
                  <th className="py-3 px-4">نام مشتری</th>
                  <th className="py-3 px-4">تلفن</th>
                  <th className="py-3 px-4">کد اقتصادی / شناسه ملی</th>
                  <th className="py-3 px-4 text-center">تخفیف پیش‌فرض</th>
                  <th className="py-3 px-4">نشانی</th>
                  <th className="py-3 px-4 text-center">عملیات</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-700/60">
                {customers.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="py-8 text-center text-slate-500">
                      هنوز مشتری ذخیره شده‌ای وجود ندارد.
                    </td>
                  </tr>
                ) : (
                  customers.map((c) => (
                    <tr key={c.id} className="hover:bg-slate-700/30">
                      <td className="py-3 px-4 font-bold text-white">{c.name}</td>
                      <td className="py-3 px-4 font-mono text-slate-300 text-xs">{c.phone || '-'}</td>
                      <td className="py-3 px-4 font-mono text-slate-400 text-xs">
                        {c.economicCode || c.nationalId ? `${c.economicCode || '-'} / ${c.nationalId || '-'}` : '-'}
                      </td>
                      <td className="py-3 px-4 text-center font-mono text-emerald-400 text-xs">
                        {c.defaultDiscount
                          ? c.defaultDiscount <= 100
                            ? `${toPersianDigits(c.defaultDiscount)}٪`
                            : `${formatNumber(c.defaultDiscount)} ${settings.currencyUnit}`
                          : '-'}
                      </td>
                      <td className="py-3 px-4 text-slate-400 text-xs">{c.address || '-'}</td>
                      <td className="py-3 px-4 text-center">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            onClick={() => setNewCustomer(c)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-blue-400 hover:bg-blue-500/10 cursor-pointer"
                            title="ویرایش مشتری"
                          >
                            <Edit2 className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => handleDeleteCustomer(c.id)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 cursor-pointer"
                            title="حذف مشتری"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* MODAL: BULK PRICE UPDATE */}
      {showBulkModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
          <div className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-md p-6 space-y-5 shadow-2xl">
            <h3 className="text-lg font-bold text-white flex items-center gap-2">
              <TrendingUp className="w-5 h-5 text-blue-400" />
              <span>تغییر درصدی قیمت‌ها (عملیات گروهی)</span>
            </h3>

            <div className="space-y-4 text-sm">
              <div>
                <label className="block text-slate-300 mb-2">محدوده اعمال تغییرات:</label>
                <div className="space-y-2">
                  <label className="flex items-center gap-2 p-3 rounded-xl bg-slate-900/60 border border-slate-700 cursor-pointer">
                    <input
                      type="radio"
                      name="bulkTarget"
                      checked={bulkTarget === 'filtered'}
                      onChange={() => setBulkTarget('filtered')}
                      className="text-blue-600"
                    />
                    <span>فقط کالاهای فیلتر شده فعلی ({toPersianDigits(filteredFittings.length)} کالا)</span>
                  </label>
                  <label className="flex items-center gap-2 p-3 rounded-xl bg-slate-900/60 border border-slate-700 cursor-pointer">
                    <input
                      type="radio"
                      name="bulkTarget"
                      checked={bulkTarget === 'all'}
                      onChange={() => setBulkTarget('all')}
                      className="text-blue-600"
                    />
                    <span>کل کاتالوگ اتصالات ({toPersianDigits(fittings.length)} کالا)</span>
                  </label>
                </div>
              </div>

              <div>
                <label className="block text-slate-300 mb-1">
                  درصد تغییر قیمت (مثبت برای افزایش، منفی برای کاهش):
                </label>
                <div className="relative">
                  <input
                    type="number"
                    value={bulkPercent}
                    onChange={(e) => setBulkPercent(Number(e.target.value) || 0)}
                    className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2.5 text-white font-mono text-left focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  />
                  <span className="absolute right-4 top-2.5 text-slate-400 pointer-events-none">%</span>
                </div>
                <p className="text-xs text-slate-400 mt-1">
                  مثال: عدد 10 تمام قیمت‌های مورد نظر را ۱۰٪ افزایش می‌دهد.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-3 pt-3 border-t border-slate-700">
              <button
                type="button"
                onClick={() => setShowBulkModal(false)}
                className="px-4 py-2 text-slate-300 hover:text-white cursor-pointer"
              >
                انصراف
              </button>
              <button
                type="button"
                onClick={executeBulkUpdate}
                className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl font-medium shadow cursor-pointer"
              >
                اعمال تغییرات قیمت
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: ADD NEW FITTING */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
          <form
            onSubmit={handleAddNewFitting}
            className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-lg p-6 space-y-4 shadow-2xl"
          >
            <h3 className="text-lg font-bold text-white flex items-center gap-2">
              <Plus className="w-5 h-5 text-emerald-400" />
              <span>افزودن کالای جدید به کاتالوگ اتصالات</span>
            </h3>

            <div className="grid grid-cols-2 gap-3 text-sm">
              <div>
                <label className="block text-slate-300 text-xs mb-1">دسته‌بندی (نوع):</label>
                <input
                  type="text"
                  required
                  value={newFitting.category || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, category: e.target.value })}
                  placeholder="مثال: اتصالات فشار قوی"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 text-xs mb-1">نوع محصول:</label>
                <input
                  type="text"
                  required
                  value={newFitting.connectionType || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, connectionType: e.target.value })}
                  placeholder="مثال: جوشی"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div className="col-span-2">
                <label className="block text-slate-300 text-xs mb-1">نام محصول:</label>
                <input
                  type="text"
                  required
                  value={newFitting.product || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, product: e.target.value })}
                  placeholder="مثال: فلنچ پایه بلند"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 text-xs mb-1">سایز (mm):</label>
                <input
                  type="text"
                  required
                  value={newFitting.size || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, size: e.target.value })}
                  placeholder="مثال: 110"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 text-xs mb-1">فشار کاری (bar):</label>
                <input
                  type="text"
                  required
                  value={newFitting.pressure || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, pressure: e.target.value })}
                  placeholder="مثال: 10 یا 16"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 text-xs mb-1">جنس کالا:</label>
                <input
                  type="text"
                  value={newFitting.material || 'HDPE100'}
                  onChange={(e) => setNewFitting({ ...newFitting, material: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 text-xs mb-1">تعداد در بسته:</label>
                <input
                  type="number"
                  value={newFitting.packageCount || 1}
                  onChange={(e) => setNewFitting({ ...newFitting, packageCount: Number(e.target.value) || 1 })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div className="col-span-2">
                <label className="block text-slate-300 text-xs mb-1">قیمت واحد (ریال):</label>
                <input
                  type="number"
                  required
                  value={newFitting.unitPriceRial || ''}
                  onChange={(e) => setNewFitting({ ...newFitting, unitPriceRial: Number(e.target.value) || 0 })}
                  placeholder="مثال: 1250000"
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
                <div className="text-xs text-slate-400 mt-1">
                  معادل تومان: {formatNumber(Math.round((newFitting.unitPriceRial || 0) / 10))} تومان
                </div>
              </div>
            </div>

            <div className="flex items-center justify-end gap-3 pt-3 border-t border-slate-700">
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="px-4 py-2 text-slate-300 hover:text-white cursor-pointer"
              >
                انصراف
              </button>
              <button
                type="submit"
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-medium shadow cursor-pointer"
              >
                افزودن به کاتالوگ
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};

