import React, { useState, useEffect } from 'react';
import { Navbar, NavTab } from './components/Navbar';
import { InvoiceBuilder } from './components/InvoiceBuilder/InvoiceBuilder';
import { PriceManager } from './components/PriceManager/PriceManager';
import { CustomerManager } from './components/CustomerManager/CustomerManager';
import { InvoiceHistory } from './components/InvoiceHistory/InvoiceHistory';
import { InvoicePrintView } from './components/InvoicePrint/InvoicePrintView';
import { SettingsModal } from './components/Settings/SettingsModal';
import { StorageService } from './services/storage';
import { getCurrentJalaliDate } from './services/persianUtils';
import { useKeyboardShortcuts } from './hooks/useKeyboardShortcuts';
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

  // Active editing session: { invoice, isDuplicate }
  const [editingSession, setEditingSession] = useState<{
    invoice: Invoice;
    isDuplicate: boolean;
  } | null>(null);

  const handleEditInvoice = (inv: Invoice) => {
    setEditingSession({ invoice: inv, isDuplicate: false });
    setActiveTab('builder');
  };

  const handleDuplicateInvoice = (inv: Invoice) => {
    const freshSettings = StorageService.getSettings();
    const nextInvoiceNum = String((freshSettings.invoiceCounter || 1) + 1);
    const duplicated: Invoice = {
      ...inv,
      id: 'inv_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      invoiceNumber: nextInvoiceNum,
      jalaliDate: getCurrentJalaliDate(),
      createdAt: new Date().toISOString(),
    };
    setEditingSession({ invoice: duplicated, isDuplicate: true });
    setActiveTab('builder');
  };

  const handleStartInvoiceForCustomer = (cust: Customer) => {
    const freshSettings = StorageService.getSettings();
    const nextInvoiceNum = String((freshSettings.invoiceCounter || 1) + 1);
    const hasDefaultPct =
      typeof cust.defaultDiscount === 'number' && cust.defaultDiscount > 0 && cust.defaultDiscount <= 100;
    const hasDefaultFixed =
      typeof cust.defaultDiscount === 'number' && cust.defaultDiscount > 100;
    const isOfficialCustomer = Boolean(cust.nationalId || cust.economicCode);

    StorageService.saveDraft({
      invoiceNumber: nextInvoiceNum,
      jalaliDate: getCurrentJalaliDate(),
      customerName: cust.name,
      customerPhone: cust.phone || '',
      customerAddress: cust.address || '',
      customerNationalId: cust.nationalId || '',
      customerEconomicCode: cust.economicCode || '',
      customerPostalCode: cust.postalCode || '',
      isOfficial: isOfficialCustomer,
      explanation: '',
      items: [],
      includeDiscount: hasDefaultPct,
      discountPercentage: hasDefaultPct ? cust.defaultDiscount : 0,
      includeCustomDiscount: hasDefaultFixed,
      customDiscountAmount: hasDefaultFixed ? cust.defaultDiscount : 0,
      includeTax: isOfficialCustomer,
      taxPercentage: freshSettings.taxRate || 10,
    });
    setEditingSession(null);
    setActiveTab('builder');
  };

  // Reload data after backup restore
  const handleRefreshAllData = () => {
    setSettings(StorageService.getSettings());
    setFittings(StorageService.getFittings());
    setDiscounts(StorageService.getDiscounts());
    setCustomers(StorageService.getCustomers());
    setInvoices(StorageService.getInvoices());
  };

  // Global Desktop Keyboard Shortcuts (Ctrl/Cmd+S, Ctrl/Cmd+P, Ctrl/Cmd+N, Escape)
  useKeyboardShortcuts({
    onSave: () => {
      if (previewInvoice) {
        window.dispatchEvent(new CustomEvent('pipecal:trigger-print'));
      } else {
        if (activeTab !== 'builder') setActiveTab('builder');
        window.dispatchEvent(new CustomEvent('pipecal:save-invoice'));
      }
    },
    onPrint: () => {
      if (previewInvoice) {
        window.dispatchEvent(new CustomEvent('pipecal:trigger-print'));
      } else {
        if (activeTab !== 'builder') setActiveTab('builder');
        window.dispatchEvent(new CustomEvent('pipecal:preview-invoice'));
      }
    },
    onNewDraft: () => {
      if (previewInvoice) {
        setPreviewInvoice(null);
      }
      if (activeTab !== 'builder') {
        setActiveTab('builder');
      }
      window.dispatchEvent(new CustomEvent('pipecal:reset-draft'));
    },
    onEscape: () => {
      if (isSettingsOpen) {
        setIsSettingsOpen(false);
      } else if (previewInvoice) {
        setPreviewInvoice(null);
      } else {
        window.dispatchEvent(new CustomEvent('pipecal:escape'));
      }
    },
  });

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
        customerCount={customers.length}
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
            editingSession={editingSession}
            onClearEditingSession={() => setEditingSession(null)}
            onPreviewPrint={(inv) => setPreviewInvoice(inv)}
            onInvoiceSaved={() => {
              setInvoices(StorageService.getInvoices());
              setSettings(StorageService.getSettings());
              setCustomers(StorageService.getCustomers());
              setEditingSession(null);
            }}
            onUpdateSettings={setSettings}
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

        {activeTab === 'customers' && (
          <CustomerManager
            customers={customers}
            onUpdateCustomers={setCustomers}
            invoices={invoices}
            settings={settings}
            onStartInvoiceForCustomer={handleStartInvoiceForCustomer}
            onPreviewPrintInvoice={(inv) => setPreviewInvoice(inv)}
            onEditInvoice={handleEditInvoice}
            onDuplicateInvoice={handleDuplicateInvoice}
          />
        )}

        {activeTab === 'history' && (
          <InvoiceHistory
            invoices={invoices}
            onUpdateInvoices={setInvoices}
            settings={settings}
            onPreviewPrint={(inv) => setPreviewInvoice(inv)}
            onEditInvoice={handleEditInvoice}
            onDuplicateInvoice={handleDuplicateInvoice}
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

