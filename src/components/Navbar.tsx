import React from 'react';
import { FileText, Database, History, Settings, Users, Receipt } from 'lucide-react';
import { getCurrentJalaliDate, toPersianDigits } from '../services/persianUtils';

export type NavTab = 'builder' | 'prices' | 'customers' | 'accounting' | 'history';

interface NavbarProps {
  activeTab: NavTab;
  setActiveTab: (tab: NavTab) => void;
  onOpenSettings: () => void;
  currencyUnit: string;
  customerCount?: number;
  debtorCount?: number;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  setActiveTab,
  onOpenSettings,
  currencyUnit,
  customerCount,
  debtorCount,
}) => {
  const currentDate = getCurrentJalaliDate();

  return (
    <header className="no-print bg-slate-800/90 backdrop-blur border-b border-slate-700/80 sticky top-0 z-40">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          {/* Logo & App Title */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600/20 border border-blue-500/30 flex items-center justify-center overflow-hidden p-1">
              <img src="./logo.png" alt="پلی غرب" className="w-full h-full object-contain" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-lg text-white">شرکت پلی غرب</span>
                <span className="text-xs px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-300 border border-blue-500/30 font-medium">
                  نسخه ۲.۲
                </span>
              </div>
              <p className="text-xs text-slate-400">سامانه صدور پیش‌فاکتور و مدیریت قیمت اتصالات و لوله‌ها</p>
            </div>
          </div>

          {/* Navigation Tabs */}
          <nav className="flex items-center gap-1 bg-slate-900/60 p-1.5 rounded-xl border border-slate-700/60">
            <button
              onClick={() => setActiveTab('builder')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer ${
                activeTab === 'builder'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                  : 'text-slate-300 hover:text-white hover:bg-slate-800'
              }`}
            >
              <FileText className="w-4 h-4" />
              <span>صدور پیش‌فاکتور</span>
            </button>

            <button
              onClick={() => setActiveTab('prices')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer ${
                activeTab === 'prices'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                  : 'text-slate-300 hover:text-white hover:bg-slate-800'
              }`}
            >
              <Database className="w-4 h-4" />
              <span>مدیریت قیمت‌ها و کاتالوگ</span>
            </button>

            <button
              onClick={() => setActiveTab('customers')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer ${
                activeTab === 'customers'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                  : 'text-slate-300 hover:text-white hover:bg-slate-800'
              }`}
            >
              <Users className="w-4 h-4" />
              <span>باشگاه مشتریان</span>
              {typeof customerCount === 'number' && customerCount > 0 && (
                <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-white/15 font-mono">
                  {toPersianDigits(customerCount)}
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('accounting')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer ${
                activeTab === 'accounting'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                  : 'text-slate-300 hover:text-white hover:bg-slate-800'
              }`}
            >
              <Receipt className="w-4 h-4" />
              <span>حسابداری و مالی</span>
              {typeof debtorCount === 'number' && debtorCount > 0 && (
                <span className="text-[11px] px-1.5 py-0.2 rounded-full bg-rose-500/20 text-rose-300 border border-rose-500/30 font-mono">
                  {toPersianDigits(debtorCount)}
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('history')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer ${
                activeTab === 'history'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25'
                  : 'text-slate-300 hover:text-white hover:bg-slate-800'
              }`}
            >
              <History className="w-4 h-4" />
              <span>سوابق فاکتورها</span>
            </button>
          </nav>

          {/* Right Tools: Date, Currency badge, Settings */}
          <div className="flex items-center gap-3">
            <div className="hidden md:flex flex-col items-end text-xs text-slate-400 bg-slate-900/40 px-3 py-1.5 rounded-lg border border-slate-800">
              <span className="text-slate-200 font-medium">امروز: {toPersianDigits(currentDate)}</span>
              <span className="text-slate-400">واحد پول: {currencyUnit}</span>
            </div>

            <button
              onClick={onOpenSettings}
              className="p-2.5 rounded-xl bg-slate-700/60 hover:bg-slate-700 text-slate-200 hover:text-white border border-slate-600/50 transition-colors cursor-pointer"
              title="تنظیمات سیستم"
            >
              <Settings className="w-5 h-5" />
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};

