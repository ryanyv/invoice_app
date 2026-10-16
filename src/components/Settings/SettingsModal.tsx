import React, { useState } from 'react';
import { X, Save, RotateCcw, Download, Upload, Building2 } from 'lucide-react';
import { CompanySettings } from '../../types';
import { StorageService } from '../../services/storage';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: CompanySettings;
  onUpdateSettings: (settings: CompanySettings) => void;
  onRefreshAllData: () => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  settings,
  onUpdateSettings,
  onRefreshAllData,
}) => {
  const [formData, setFormData] = useState<CompanySettings>(settings);
  const [saveNotice, setSaveNotice] = useState(false);

  if (!isOpen) return null;

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdateSettings(formData);
    StorageService.saveSettings(formData);
    setSaveNotice(true);
    setTimeout(() => {
      setSaveNotice(false);
      onClose();
    }, 1200);
  };

  const handleDownloadBackup = () => {
    const backupJson = StorageService.createFullBackup();
    const blob = new Blob([backupJson], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `pipecal_backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleRestoreBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const text = evt.target?.result as string;
        const ok = StorageService.restoreFromBackup(text);
        if (ok) {
          alert('اطلاعات با موفقیت بازیابی شد.');
          onRefreshAllData();
          onClose();
        } else {
          alert('خطا در ساختار فایل پشتیبان.');
        }
      } catch {
        alert('فایل پشتیبان نامعتبر است.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const handleResetDefaults = () => {
    if (confirm('هشدار: آیا می‌خواهید تمام داده‌های کاتالوگ و قیمت‌ها به حالت اولیه بازگردند؟')) {
      StorageService.resetToDefaults();
      onRefreshAllData();
      alert('تنظیمات به حالت اولیه بازنشانی شدند.');
      onClose();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm p-4">
      <div className="bg-slate-800 border border-slate-700 rounded-3xl w-full max-w-2xl max-h-[90vh] overflow-y-auto p-6 sm:p-8 space-y-6 shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-700 pb-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 rounded-xl bg-blue-600/20 border border-blue-500/30 text-blue-400">
              <Building2 className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-xl font-bold text-white">تنظیمات شرکت و سیستم</h2>
              <p className="text-xs text-slate-400">مشخصات سربرگ پیش‌فاکتور، مالیات، واحد پول و پشتیبان‌گیری</p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-2 rounded-xl text-slate-400 hover:text-white hover:bg-slate-700 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Settings Form */}
        <form onSubmit={handleSave} className="space-y-6">
          <div className="space-y-4">
            <h3 className="text-sm font-bold text-blue-400">مشخصات سربرگ شرکت در پیش‌فاکتور</h3>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-sm">
              <div className="sm:col-span-2">
                <label className="block text-xs text-slate-400 mb-1">نام رسمی شرکت:</label>
                <input
                  type="text"
                  required
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-xs text-slate-400 mb-1">آدرس دفتر / کارخانه:</label>
                <input
                  type="text"
                  required
                  value={formData.address}
                  onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">تلفن تماس:</label>
                <input
                  type="text"
                  value={formData.phone}
                  onChange={(e) => setFormData({ ...formData, phone: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">شماره فکس (اختیاری):</label>
                <input
                  type="text"
                  value={formData.fax || ''}
                  onChange={(e) => setFormData({ ...formData, fax: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Official Registration & Banking Details (R3 ماده ۱۶۹ م.م) */}
          <div className="space-y-4 pt-4 border-t border-slate-700">
            <h3 className="text-sm font-bold text-blue-400">اطلاعات رسمی، کد اقتصادی و حساب بانکی (فاکتور رسمی ماده ۱۶۹)</h3>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-sm">
              <div>
                <label className="block text-xs text-slate-400 mb-1">کد اقتصادی شرکت:</label>
                <input
                  type="text"
                  value={formData.economicCode || ''}
                  onChange={(e) => setFormData({ ...formData, economicCode: e.target.value })}
                  placeholder="مثال: ۴۱۱۴۵۶۷۸۹۱۲۳"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">شناسه ملی شرکت:</label>
                <input
                  type="text"
                  value={formData.nationalId || ''}
                  onChange={(e) => setFormData({ ...formData, nationalId: e.target.value })}
                  placeholder="مثال: ۱۰۱۰۳۵۶۷۸۹۰"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">شماره ثبت:</label>
                <input
                  type="text"
                  value={formData.registrationNumber || ''}
                  onChange={(e) => setFormData({ ...formData, registrationNumber: e.target.value })}
                  placeholder="مثال: ۳۸۴۵۲۱"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">کد پستی ۱۰ رقمی:</label>
                <input
                  type="text"
                  value={formData.postalCode || ''}
                  onChange={(e) => setFormData({ ...formData, postalCode: e.target.value })}
                  placeholder="مثال: ۱۵۴۸۷۶۳۲۱۱"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">نام بانک:</label>
                <input
                  type="text"
                  value={formData.bankName || ''}
                  onChange={(e) => setFormData({ ...formData, bankName: e.target.value })}
                  placeholder="مثال: بانک ملت شعبه میرداماد"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">شماره حساب بانک:</label>
                <input
                  type="text"
                  value={formData.bankAccount || ''}
                  onChange={(e) => setFormData({ ...formData, bankAccount: e.target.value })}
                  placeholder="مثال: ۱۲۳۴۵۶۷۸۹۰"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-xs text-slate-400 mb-1">شماره شبا (IBAN):</label>
                <input
                  type="text"
                  value={formData.bankIban || ''}
                  onChange={(e) => setFormData({ ...formData, bankIban: e.target.value })}
                  placeholder="مثال: IR120120000000012345678901"
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left uppercase focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Financial & Invoicing Defaults */}
          <div className="space-y-4 pt-4 border-t border-slate-700">
            <h3 className="text-sm font-bold text-blue-400">تنظیمات مالی و شماره‌گذاری</h3>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-sm">
              <div>
                <label className="block text-xs text-slate-400 mb-1">درصد مالیات ارزش افزوده (%):</label>
                <input
                  type="number"
                  value={formData.taxRate}
                  onChange={(e) => setFormData({ ...formData, taxRate: Number(e.target.value) || 0 })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">شماره آخرین پیش‌فاکتور:</label>
                <input
                  type="number"
                  value={formData.invoiceCounter}
                  onChange={(e) => setFormData({ ...formData, invoiceCounter: Number(e.target.value) || 1 })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-mono text-left focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs text-slate-400 mb-1">واحد پول پیش‌فاکتور:</label>
                <select
                  value={formData.currencyUnit}
                  onChange={(e) => setFormData({ ...formData, currencyUnit: e.target.value as 'تومان' | 'ریال' })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white focus:outline-none"
                >
                  <option value="تومان">تومان</option>
                  <option value="ریال">ریال</option>
                </select>
              </div>
            </div>
          </div>

          {/* Backup & Restore */}
          <div className="space-y-4 pt-4 border-t border-slate-700">
            <h3 className="text-sm font-bold text-blue-400">پشتیبان‌گیری و بازیابی اطلاعات</h3>
            
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={handleDownloadBackup}
                className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-xl text-xs font-medium transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>دانلود فایل کامل پشتیبان (JSON)</span>
              </button>

              <label className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-xl text-xs font-medium transition-colors cursor-pointer">
                <Upload className="w-4 h-4" />
                <span>بازیابی اطلاعات از فایل پشتیبان</span>
                <input
                  type="file"
                  accept=".json"
                  onChange={handleRestoreBackup}
                  className="hidden"
                />
              </label>

              <button
                type="button"
                onClick={handleResetDefaults}
                className="flex items-center gap-2 px-4 py-2 bg-rose-500/20 hover:bg-rose-500/30 text-rose-300 border border-rose-500/30 rounded-xl text-xs font-medium transition-colors mr-auto cursor-pointer"
              >
                <RotateCcw className="w-4 h-4" />
                <span>بازنشانی به حالت کارخانه</span>
              </button>
            </div>
          </div>

          {/* Modal Footer */}
          <div className="flex items-center justify-between pt-6 border-t border-slate-700">
            <span className="text-xs text-emerald-400 font-medium">
              {saveNotice && '✓ تنظیمات با موفقیت ذخیره شدند'}
            </span>

            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={onClose}
                className="px-5 py-2.5 text-sm text-slate-300 hover:text-white cursor-pointer"
              >
                انصراف
              </button>

              <button
                type="submit"
                className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-bold shadow-lg shadow-blue-600/30 transition-all cursor-pointer"
              >
                <Save className="w-4 h-4" />
                <span>ذخیره تنظیمات</span>
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};

