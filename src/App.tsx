import React, { useState, useEffect } from 'react';
import { Navbar, NavTab } from './components/Navbar';
import { InvoiceBuilder } from './components/InvoiceBuilder/InvoiceBuilder';
import { PriceManager } from './components/PriceManager/PriceManager';
import { InvoiceHistory } from './components/InvoiceHistory/InvoiceHistory';
import { InvoicePrintView } from './components/InvoicePrint/InvoicePrintView';
import { SettingsModal } from './components/Settings/SettingsModal';
import { StorageService } from './services/storage';
import {
  CompanySettings,
  Customer,
  DiscountRule,
  FittingCatalogItem,
  GradeSdrPn,
  Invoice,
  WeightTable,
} from './types';

export const App: React.FC = () => {
  // Navigation State
  const [activeTab, setActiveTab] = useState<NavTab>('builder');

  useEffect(() => {
    console.log('[Pipecal Pro App] UI mounted and running successfully!');
  }, []);

  // Application Data State
  const [settings, setSettings] = useState<CompanySettings>(() => StorageService.getSettings());
  const [fittings, setFittings] = useState<FittingCatalogItem[]>(() => StorageService.getFittings());
  const [discounts, setDiscounts] = useState<DiscountRule[]>(() => StorageService.getDiscounts());
  const [customers, setCustomers] = useState<Customer[]>(() => StorageService.getCustomers());
  const [invoices, setInvoices] = useState<Invoice[]>(() => StorageService.getInvoices());

  // Standards (DIN & SDR tables)
  const [weightTable] = useState<WeightTable>(() => StorageService.getWeightTable());
  const [gradeSdrPn] = useState<GradeSdrPn>(() => StorageService.getGradeSdrPn());

  // Modals & Full views
  const [previewInvoice, setPreviewInvoice] = useState<Invoice | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  // Reload data after backup restore
  const handleRefreshAllData = () => {
    setSettings(StorageService.getSettings());
    setFittings(StorageService.getFittings());
    setDiscounts(StorageService.getDiscounts());
    setCustomers(StorageService.getCustomers());
    setInvoices(StorageService.getInvoices());
  };

  // If viewing printable preview
  if (previewInvoice) {
    return (
      <InvoicePrintView
        invoice={previewInvoice}
        settings={settings}
        onBack={() => setPreviewInvoice(null)}
      />
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col font-sans">
      {/* Top Navbar */}
      <Navbar
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        onOpenSettings={() => setIsSettingsOpen(true)}
        currencyUnit={settings.currencyUnit}
      />

      {/* Main Content Area */}
      <main className="flex-1 pb-16">
        {activeTab === 'builder' && (
          <InvoiceBuilder
            settings={settings}
            fittings={fittings}
            discounts={discounts}
            customers={customers}
            weightTable={weightTable}
            gradeSdrPn={gradeSdrPn}
            onPreviewPrint={(inv) => setPreviewInvoice(inv)}
            onInvoiceSaved={() => setInvoices(StorageService.getInvoices())}
          />
        )}

        {activeTab === 'prices' && (
          <PriceManager
            settings={settings}
            onUpdateSettings={setSettings}
            fittings={fittings}
            onUpdateFittings={setFittings}
            discounts={discounts}
            onUpdateDiscounts={setDiscounts}
          />
        )}

        {activeTab === 'history' && (
          <InvoiceHistory
            invoices={invoices}
            onUpdateInvoices={setInvoices}
            settings={settings}
            onPreviewPrint={(inv) => setPreviewInvoice(inv)}
          />
        )}
      </main>

      {/* Settings Modal */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        settings={settings}
        onUpdateSettings={setSettings}
        onRefreshAllData={handleRefreshAllData}
      />
    </div>
  );
};

export default App;

