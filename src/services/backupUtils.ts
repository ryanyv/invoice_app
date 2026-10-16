/**
 * Safe Storage Persistence & Automated Rolling Backups Utilities
 * Milestone M4 (Requirement R4: Offline & Desktop Reliability)
 */

export interface BackupItem {
  filename: string;
  filePath?: string;
  mtimeMs: number;
  size?: number;
  createdAt?: string;
}

export const BACKUP_RETENTION_LIMIT = 20;

/**
 * Generates a standard rolling backup filename in the format:
 * pipecal-backup-YYYY-MM-DD-HHmmss.json
 */
export function generateBackupFilename(date: Date = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  const y = date.getFullYear();
  const m = pad(date.getMonth() + 1);
  const d = pad(date.getDate());
  const h = pad(date.getHours());
  const min = pad(date.getMinutes());
  const s = pad(date.getSeconds());
  return `pipecal-backup-${y}-${m}-${d}-${h}${min}${s}.json`;
}

/**
 * Parses timestamp from a backup filename.
 */
export function parseBackupTimestamp(filename: string): Date | null {
  const match = filename.match(/pipecal-backup-(\d{4})-(\d{2})-(\d{2})-(\d{2})(\d{2})(\d{2})\.json/);
  if (!match) return null;
  const [, y, m, d, h, min, s] = match;
  return new Date(Number(y), Number(m) - 1, Number(d), Number(h), Number(min), Number(s));
}

/**
 * Prunes backups exceeding retention limit.
 * Keeps the newest `maxRetention` items, returns the list of older items to delete.
 */
export function pruneBackups<T extends { filename: string; mtimeMs: number }>(
  backups: T[],
  maxRetention: number = BACKUP_RETENTION_LIMIT
): T[] {
  // Sort newest first
  const sorted = [...backups].sort((a, b) => b.mtimeMs - a.mtimeMs);
  if (sorted.length <= maxRetention) {
    return [];
  }
  return sorted.slice(maxRetention);
}

/**
 * Formats byte size into human readable string.
 */
export function formatBackupSize(bytes: number = 0): string {
  if (bytes <= 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}
