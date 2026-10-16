/// <reference types="vite/client" />

export interface BackupFileEntry {
  filename: string;
  filePath: string;
  size: number;
  createdAt: string;
  mtimeMs?: number;
}

export interface ElectronAPI {
  isElectron?: boolean;
  savePdf?: (options: { defaultName?: string; landscape?: boolean }) => Promise<{ success: boolean; filePath?: string; canceled?: boolean; error?: string }>;
  exportFile?: (options: { defaultName?: string; content: string; extensions?: string[] }) => Promise<{ success: boolean; filePath?: string; canceled?: boolean; error?: string }>;
  importFile?: (options: { extensions?: string[] }) => Promise<{ success: boolean; content?: string; filePath?: string; canceled?: boolean; error?: string }>;
  saveBackup?: (options: { backupJson: string }) => Promise<{ success: boolean; filePath?: string; filename?: string; error?: string }>;
  listBackups?: () => Promise<{ success: boolean; backups?: BackupFileEntry[]; files?: string[]; error?: string }>;
  loadBackup?: (options: { filename: string }) => Promise<{ success: boolean; content?: string; filePath?: string; error?: string }>;
  fetchMarketPrices?: () => Promise<{ success: boolean; isOnline: boolean; snapshot?: any; error?: string }>;
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}

