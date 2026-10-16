import React, { useState, useEffect } from 'react';
import { X, Save, RotateCcw, Download, Upload, Building2, HardDrive, RefreshCw } from 'lucide-react';
import { CompanySettings } from '../../types';
import { StorageService } from '../../services/storage';
import { formatBackupSize } from '../../services/backupUtils';

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
  const [desktopBackups, setDesktopBackups] = useState<Array<{ filename: string; filePath: string; size: number; createdAt: string }>>([]);
  const [isLoadingBackups, setIsLoadingBackups] = useState(false);
  const [backupMessage, setBackupMessage] = useState<{ text: string; isError?: boolean } | null>(null);

  const loadDesktopBackups = async () => {
    if (typeof window !== 'undefined' && window.electronAPI?.listBackups) {
      setIsLoadingBackups(true);
      try {
        const res = await window.electronAPI.listBackups();
        if (res.success && res.backups) {
          setDesktopBackups(res.backups);
        } else if (res.success && res.files) {
          const mapped = res.files.map((f: any) =>
            typeof f === 'string' ? { filename: f, filePath: f, size: 0, createdAt: '' } : f
          );
          setDesktopBackups(mapped);
        }
      } catch (err) {
        console.error('Failed to list backups:', err);
      } finally {
        setIsLoadingBackups(false);
      }
    }
  };

  useEffect(() => {
    if (isOpen) {
      setFormData(settings);
      loadDesktopBackups();
    }
  }, [isOpen, settings]);

  const handleRestoreDesktopBackup = async (filename: string) => {
    if (!confirm(`آیا از بازیابی نسخه پشتیبان «${filename}» اطمینان دارید؟ داده‌های فعلی نرم‌افزار جایگزین خواهند شد.`)) {
      return;
    }
    if (window.electronAPI?.loadBackup) {
      try {
        const res = await window.electronAPI.loadBackup({ filename });
        if (res.success && res.content) {
          const ok = StorageService.restoreFromBackup(res.content);
          if (ok) {
            setBackupMessage({ text: `نسخه پشتیبان «${filename}» با موفقیت بازیابی شد.` });
            onRefreshAllData();
            setTimeout(() => setBackupMessage(null), 3500);
          } else {
            setBackupMessage({ text: 'خطا در محتوای نسخه پشتیبان.', isError: true });
          }
        } else {
          setBackupMessage({ text: res.error || 'خطا در بارگذاری فایل پشتیبان', isError: true });
        }
      } catch (err) {
        setBackupMessage({ text: 'خطا در بازیابی نسخه پشتیبان.', isError: true });
      }
    }
  };

  const handleCreateSnapshot = async () => {
    if (window.electronAPI?.saveBackup) {
      const ok = await StorageService.triggerDiskBackup();
      if (ok) {
        setBackupMessage({ text: 'نسخه پشتیبان جدید با موفقیت روی دیسک ایجاد شد.' });
        await loadDesktopBackups();
        setTimeout(() => setBackupMessage(null), 3000);
      } else {
        setBackupMessage({ text: 'خطا در ایجاد نسخه پشتیبان روی دیسک.', isError: true });
      }
    }
  };

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

            {/* Automated Desktop Backups (R4) */}
            <div className="mt-4 pt-4 border-t border-slate-700/60 space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="text-xs font-bold text-blue-400 flex items-center gap-2">
                    <HardDrive className="w-4 h-4 text-blue-400" />
                    <span>نسخه‌های پشتیبان خودکار دسکتاپ</span>
                  </h4>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    ذخیره‌سازی ایمن روی دیسک با نگهداری خودکار ۲۰ نسخه اخیر
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={loadDesktopBackups}
                    disabled={isLoadingBackups}
                    className="p-1.5 bg-slate-700 hover:bg-slate-600 text-slate-300 rounded-lg text-xs transition-colors cursor-pointer"
                    title="بروزرسانی فهرست پشتیبان‌ها"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isLoadingBackups ? 'animate-spin' : ''}`} />
                  </button>
                  <button
                    type="button"
                    onClick={handleCreateSnapshot}
                    className="px-2.5 py-1 bg-blue-600/20 hover:bg-blue-600/40 text-blue-300 border border-blue-500/30 rounded-lg text-xs font-medium transition-colors cursor-pointer"
                  >
                    + ذخیره فوری
                  </button>
                </div>
              </div>

              {backupMessage && (
                <div
                  className={`p-2 rounded-xl text-xs font-medium ${
                    backupMessage.isError
                      ? 'bg-rose-900/40 text-rose-300 border border-rose-700/50'
                      : 'bg-emerald-900/40 text-emerald-300 border border-emerald-700/50'
                  }`}
                >
                  {backupMessage.text}
                </div>
              )}

              {desktopBackups.length > 0 ? (
                <div className="max-h-48 overflow-y-auto space-y-2 pr-1 border border-slate-700/60 rounded-xl p-2 bg-slate-900/40">
                  {desktopBackups.map((b) => (
                    <div
                      key={b.filename}
                      className="flex items-center justify-between p-2 rounded-lg bg-slate-800/80 hover:bg-slate-800 border border-slate-700/50 text-xs transition-colors"
                    >
                      <div className="flex flex-col gap-0.5">
                        <span className="font-mono text-slate-200 text-[11px] dir-ltr text-right">
                          {b.filename}
                        </span>
                        <div className="flex items-center gap-3 text-[10px] text-slate-400">
                          <span>حجم: {formatBackupSize(b.size)}</span>
                          {b.createdAt && (
                            <span>
                              تاریخ:{' '}
                              {new Date(b.createdAt).toLocaleDateString('fa-IR', {
                                year: 'numeric',
                                month: 'short',
                                day: 'numeric',
                                hour: '2-digit',
                                minute: '2-digit',
                              })}
                            </span>
                          )}
                        </div>
                      </div>

                      <button
                        type="button"
                        onClick={() => handleRestoreDesktopBackup(b.filename)}
                        className="flex items-center gap-1 px-2.5 py-1 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 rounded-lg font-medium transition-colors cursor-pointer text-xs"
                      >
                        <RotateCcw className="w-3.5 h-3.5" />
                        <span>بازیابی</span>
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-3 px-2 bg-slate-900/30 border border-slate-700/40 rounded-xl text-xs text-slate-400">
                  {typeof window !== 'undefined' && window.electronAPI
                    ? 'هنوز نسخه پشتیبان خودکاری روی دیسک ایجاد نشده است.'
                    : 'نسخه‌های پشتیبان خودکار دسکتاپ در محیط دسکتاپ (Electron) فعال خواهند بود.'}
                </div>
              )}
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

