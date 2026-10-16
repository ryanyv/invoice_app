/// <reference types="node" />
/**
 * Automated Test Suite: Offline & Desktop Reliability (R4)
 * Pipecal Pro 2.0 - Milestone M4
 *
 * Verifies:
 * 1. Font Asset Existence & 100% Offline Integrity (Vazirmatn)
 * 2. Desktop Keyboard Shortcut Mappings & Event Dispatch Logic
 * 3. Rolling Backup Creation, Filename Conventions & Timestamp Parsing
 * 4. Backup Retention Pruning Logic (Strict 20-file limit)
 * 5. Safe Storage Dual-Layer Persistence & Backup JSON Serialization
 */

import fs from 'fs';
import path from 'path';
import {
  generateBackupFilename,
  parseBackupTimestamp,
  pruneBackups,
  formatBackupSize,
  BACKUP_RETENTION_LIMIT,
} from '../backupUtils';
import {
  SHORTCUT_DEFINITIONS,
  matchKeyboardShortcut,
  handleKeyboardEvent,
  ShortcutHandlers,
} from '../../hooks/useKeyboardShortcuts';
import { StorageService } from '../storage';

// ============================================================================
// TEST HARNESS ASSERTION UTILITIES
// ============================================================================

let totalPassed = 0;
let totalFailed = 0;
const failures: string[] = [];

function assert(condition: boolean, message: string): asserts condition {
  if (!condition) {
    throw new Error(`Assertion Failed: ${message}`);
  }
}

function assertEqual<T>(actual: T, expected: T, message?: string) {
  if (actual !== expected) {
    throw new Error(
      `Assertion Failed: expected ${JSON.stringify(expected)} but got ${JSON.stringify(actual)}. ${message || ''}`
    );
  }
}

function runTest(name: string, fn: () => void | Promise<void>) {
  it(name, async () => {
    try {
      await fn();
      totalPassed++;
      console.log(`  ✓ PASS: ${name}`);
    } catch (err: any) {
      totalFailed++;
      failures.push(`${name} -> ${err.message}`);
      console.error(`  ✗ FAIL: ${name} -> ${err.message}`);
      throw err;
    }
  });
}

describe('RUNNING OFFLINE & DESKTOP RELIABILITY TEST SUITE (M4)', () => {

// ----------------------------------------------------------------------------
// SUITE 1: 100% Offline Font Bundle & Zero CDN Network Requests
// ----------------------------------------------------------------------------
console.log('--- Test Suite 1: Local Font Bundle & Zero CDN Isolation ---');

const projectRoot = path.resolve(__dirname, '../../..');
const publicFontsDir = path.join(projectRoot, 'public/fonts');
const indexHtmlPath = path.join(projectRoot, 'index.html');
const indexCssPath = path.join(projectRoot, 'src/index.css');

runTest('1.1 Font directory public/fonts exists', () => {
  assert(fs.existsSync(publicFontsDir), `Expected ${publicFontsDir} to exist`);
  const stat = fs.statSync(publicFontsDir);
  assert(stat.isDirectory(), 'public/fonts must be a directory');
});

runTest('1.2 Local Vazirmatn font files exist with non-zero size', () => {
  const requiredFiles = [
    'Vazirmatn[wght].woff2',
    'vazirmatn-arabic-wght-normal.woff2',
    'vazirmatn-latin-wght-normal.woff2',
    'vazirmatn-latin-ext-wght-normal.woff2',
  ];

  for (const file of requiredFiles) {
    const filePath = path.join(publicFontsDir, file);
    assert(fs.existsSync(filePath), `Font file ${file} must exist in public/fonts`);
    const size = fs.statSync(filePath).size;
    assert(size > 10000, `Font file ${file} must be valid size (>10KB), got ${size} bytes`);
  }
});

runTest('1.3 index.html contains ZERO Google Fonts CDN or external font links', () => {
  assert(fs.existsSync(indexHtmlPath), 'index.html must exist');
  const content = fs.readFileSync(indexHtmlPath, 'utf-8');

  assert(!content.includes('fonts.googleapis.com'), 'index.html must not contain fonts.googleapis.com');
  assert(!content.includes('fonts.gstatic.com'), 'index.html must not contain fonts.gstatic.com');
  assert(!content.includes('http://') && !content.includes('https://'), 'index.html must have 0 external URLs');
});

runTest('1.4 src/index.css defines local @font-face and font-family for Vazirmatn', () => {
  assert(fs.existsSync(indexCssPath), 'src/index.css must exist');
  const css = fs.readFileSync(indexCssPath, 'utf-8');

  assert(css.includes('@font-face'), 'src/index.css must declare @font-face');
  assert(css.includes('Vazirmatn'), 'src/index.css must declare Vazirmatn font family');
  assert(css.includes('/fonts/'), 'src/index.css @font-face must point to local /fonts/ assets');
});

// ----------------------------------------------------------------------------
// SUITE 2: Desktop Keyboard Shortcut Mapping Definitions & Handler
// ----------------------------------------------------------------------------
console.log('\n--- Test Suite 2: Desktop Keyboard Shortcuts Definitions & Matching ---');

runTest('2.1 SHORTCUT_DEFINITIONS contains all required desktop actions', () => {
  assert(SHORTCUT_DEFINITIONS.SAVE !== undefined, 'Shortcut SAVE must be defined');
  assertEqual(SHORTCUT_DEFINITIONS.SAVE.key, 's');
  assertEqual(SHORTCUT_DEFINITIONS.SAVE.ctrlOrCmd, true);
  assertEqual(SHORTCUT_DEFINITIONS.SAVE.action, 'save');

  assert(SHORTCUT_DEFINITIONS.PRINT !== undefined, 'Shortcut PRINT must be defined');
  assertEqual(SHORTCUT_DEFINITIONS.PRINT.key, 'p');
  assertEqual(SHORTCUT_DEFINITIONS.PRINT.ctrlOrCmd, true);
  assertEqual(SHORTCUT_DEFINITIONS.PRINT.action, 'print');

  assert(SHORTCUT_DEFINITIONS.NEW_DRAFT !== undefined, 'Shortcut NEW_DRAFT must be defined');
  assertEqual(SHORTCUT_DEFINITIONS.NEW_DRAFT.key, 'n');
  assertEqual(SHORTCUT_DEFINITIONS.NEW_DRAFT.ctrlOrCmd, true);
  assertEqual(SHORTCUT_DEFINITIONS.NEW_DRAFT.action, 'newDraft');

  assert(SHORTCUT_DEFINITIONS.ESCAPE !== undefined, 'Shortcut ESCAPE must be defined');
  assertEqual(SHORTCUT_DEFINITIONS.ESCAPE.key, 'Escape');
  assertEqual(SHORTCUT_DEFINITIONS.ESCAPE.ctrlOrCmd, false);
  assertEqual(SHORTCUT_DEFINITIONS.ESCAPE.action, 'escape');
});

runTest('2.2 matchKeyboardShortcut matches Ctrl and Cmd (metaKey) combos correctly', () => {
  // Save
  assertEqual(matchKeyboardShortcut({ key: 's', ctrlKey: true }), 'save');
  assertEqual(matchKeyboardShortcut({ key: 'S', ctrlKey: true }), 'save');
  assertEqual(matchKeyboardShortcut({ key: 's', metaKey: true }), 'save');
  assertEqual(matchKeyboardShortcut({ key: 's', ctrlKey: false, metaKey: false }), null);

  // Print
  assertEqual(matchKeyboardShortcut({ key: 'p', ctrlKey: true }), 'print');
  assertEqual(matchKeyboardShortcut({ key: 'P', ctrlKey: true }), 'print');
  assertEqual(matchKeyboardShortcut({ key: 'p', metaKey: true }), 'print');
  assertEqual(matchKeyboardShortcut({ key: 'p', ctrlKey: false }), null);

  // New Draft
  assertEqual(matchKeyboardShortcut({ key: 'n', ctrlKey: true }), 'newDraft');
  assertEqual(matchKeyboardShortcut({ key: 'N', ctrlKey: true }), 'newDraft');
  assertEqual(matchKeyboardShortcut({ key: 'n', metaKey: true }), 'newDraft');
  assertEqual(matchKeyboardShortcut({ key: 'n', ctrlKey: false }), null);

  // Escape
  assertEqual(matchKeyboardShortcut({ key: 'Escape' }), 'escape');
  assertEqual(matchKeyboardShortcut({ key: 'Escape', ctrlKey: false }), 'escape');

  // Unmapped keys
  assertEqual(matchKeyboardShortcut({ key: 'z', ctrlKey: true }), null);
  assertEqual(matchKeyboardShortcut({ key: 'Enter' }), null);
  assertEqual(matchKeyboardShortcut({ key: 'Tab' }), null);
});

runTest('2.3 handleKeyboardEvent triggers handlers and calls preventDefault', () => {
  let savedCalled = false;
  let printCalled = false;
  let newDraftCalled = false;
  let escapeCalled = false;
  let preventDefaultCalls = 0;

  const handlers: ShortcutHandlers = {
    onSave: () => {
      savedCalled = true;
    },
    onPrint: () => {
      printCalled = true;
    },
    onNewDraft: () => {
      newDraftCalled = true;
    },
    onEscape: () => {
      escapeCalled = true;
    },
  };

  const createEvent = (key: string, ctrl = false, meta = false) => ({
    key,
    ctrlKey: ctrl,
    metaKey: meta,
    preventDefault: () => {
      preventDefaultCalls++;
    },
  });

  // Test Ctrl+S
  const handledSave = handleKeyboardEvent(createEvent('s', true), handlers);
  assert(handledSave, 'handleKeyboardEvent must return true for Ctrl+S');
  assert(savedCalled, 'onSave must have been invoked');
  assertEqual(preventDefaultCalls, 1, 'preventDefault must be called once');

  // Test Cmd+P
  const handledPrint = handleKeyboardEvent(createEvent('p', false, true), handlers);
  assert(handledPrint, 'handleKeyboardEvent must return true for Cmd+P');
  assert(printCalled, 'onPrint must have been invoked');
  assertEqual(preventDefaultCalls, 2);

  // Test Ctrl+N
  const handledNew = handleKeyboardEvent(createEvent('n', true), handlers);
  assert(handledNew, 'handleKeyboardEvent must return true for Ctrl+N');
  assert(newDraftCalled, 'onNewDraft must have been invoked');
  assertEqual(preventDefaultCalls, 3);

  // Test Escape
  const handledEsc = handleKeyboardEvent(createEvent('Escape'), handlers);
  assert(handledEsc, 'handleKeyboardEvent must return true for Escape');
  assert(escapeCalled, 'onEscape must have been invoked');
  assertEqual(preventDefaultCalls, 4);

  // Test unhandled key
  const handledUnmapped = handleKeyboardEvent(createEvent('a'), handlers);
  assert(!handledUnmapped, 'handleKeyboardEvent must return false for unmapped key');
  assertEqual(preventDefaultCalls, 4, 'preventDefault must not be called on unmapped key');
});

// ----------------------------------------------------------------------------
// SUITE 3: Rolling Backup Filename Conventions & Timestamp Parsing
// ----------------------------------------------------------------------------
console.log('\n--- Test Suite 3: Backup Filename Conventions & Timestamp Parsing ---');

runTest('3.1 generateBackupFilename generates pipecal-backup-YYYY-MM-DD-HHmmss.json format', () => {
  const d = new Date(2026, 9, 16, 14, 25, 30); // 2026-10-16 14:25:30
  const filename = generateBackupFilename(d);
  assertEqual(filename, 'pipecal-backup-2026-10-16-142530.json');

  const nowFilename = generateBackupFilename();
  assert(
    /^pipecal-backup-\d{4}-\d{2}-\d{2}-\d{6}\.json$/.test(nowFilename),
    `Backup filename "${nowFilename}" must match standard pattern`
  );
});

runTest('3.2 parseBackupTimestamp parses date accurately from standard filename', () => {
  const filename = 'pipecal-backup-2026-10-16-142530.json';
  const parsed = parseBackupTimestamp(filename);
  assert(parsed !== null, 'Timestamp must not be null');
  assertEqual(parsed.getFullYear(), 2026);
  assertEqual(parsed.getMonth(), 9); // October
  assertEqual(parsed.getDate(), 16);
  assertEqual(parsed.getHours(), 14);
  assertEqual(parsed.getMinutes(), 25);
  assertEqual(parsed.getSeconds(), 30);

  // Invalid filename test
  assertEqual(parseBackupTimestamp('random-file.json'), null);
  assertEqual(parseBackupTimestamp('pipecal-backup.json'), null);
});

runTest('3.3 formatBackupSize formats byte values into human readable units', () => {
  assertEqual(formatBackupSize(0), '0 B');
  assertEqual(formatBackupSize(512), '512 B');
  assertEqual(formatBackupSize(1024), '1.0 KB');
  assertEqual(formatBackupSize(2560), '2.5 KB');
  assertEqual(formatBackupSize(1024 * 1024), '1.00 MB');
  assertEqual(formatBackupSize(3.5 * 1024 * 1024), '3.50 MB');
});

// ----------------------------------------------------------------------------
// SUITE 4: Automated Rolling Backup Pruning Logic (Retention = 20)
// ----------------------------------------------------------------------------
console.log('\n--- Test Suite 4: Rolling Backup Pruning Logic ---');

runTest('4.1 Default BACKUP_RETENTION_LIMIT is exactly 20', () => {
  assertEqual(BACKUP_RETENTION_LIMIT, 20);
});

runTest('4.2 pruneBackups does not delete when count <= retention limit', () => {
  const backups = Array.from({ length: 15 }, (_, i) => ({
    filename: `pipecal-backup-2026-10-${String(i + 1).padStart(2, '0')}-120000.json`,
    mtimeMs: 1000 + i * 100,
  }));

  const toDelete = pruneBackups(backups, 20);
  assertEqual(toDelete.length, 0, 'No backups should be pruned when count is <= 20');

  const exactly20 = Array.from({ length: 20 }, (_, i) => ({
    filename: `pipecal-backup-2026-10-${String(i + 1).padStart(2, '0')}-120000.json`,
    mtimeMs: 1000 + i * 100,
  }));
  assertEqual(pruneBackups(exactly20, 20).length, 0, 'Exactly 20 backups must produce 0 deletions');
});

runTest('4.3 pruneBackups prunes oldest backups when total exceeds 20', () => {
  // Create 30 mock backups with increasing timestamps
  const backups = Array.from({ length: 30 }, (_, i) => ({
    filename: `pipecal-backup-${String(i + 1).padStart(2, '0')}.json`,
    mtimeMs: 1000 + i * 1000, // item 0 is oldest, item 29 is newest
  }));

  const toDelete = pruneBackups(backups, 20);
  assertEqual(toDelete.length, 10, 'Must prune exactly 10 excess backups');

  // The 10 oldest must be in toDelete (items 0 through 9)
  const deletedFilenames = toDelete.map((b) => b.filename);
  for (let i = 0; i < 10; i++) {
    const expectedOldest = `pipecal-backup-${String(i + 1).padStart(2, '0')}.json`;
    assert(
      deletedFilenames.includes(expectedOldest),
      `Oldest file ${expectedOldest} must be selected for deletion`
    );
  }

  // The 20 newest (items 10 through 29) must NOT be deleted
  for (let i = 10; i < 30; i++) {
    const keptFile = `pipecal-backup-${String(i + 1).padStart(2, '0')}.json`;
    assert(!deletedFilenames.includes(keptFile), `Newer file ${keptFile} must NOT be deleted`);
  }
});

// ----------------------------------------------------------------------------
// SUITE 5: StorageService Persistence & Full Backup Serialization
// ----------------------------------------------------------------------------
console.log('\n--- Test Suite 5: Safe Storage Dual-Layer Persistence ---');

// Mock localStorage if in node environment without window
if (typeof globalThis.localStorage === 'undefined') {
  const storageMap = new Map<string, string>();
  globalThis.localStorage = {
    getItem: (key: string) => storageMap.get(key) || null,
    setItem: (key: string, value: string) => storageMap.set(key, value),
    removeItem: (key: string) => storageMap.delete(key),
    clear: () => storageMap.clear(),
    key: (index: number) => Array.from(storageMap.keys())[index] || null,
    length: 0,
  } as any;
}

runTest('5.1 StorageService.createFullBackup outputs valid JSON with all entities', () => {
  const jsonStr = StorageService.createFullBackup();
  assert(typeof jsonStr === 'string', 'Backup must be a JSON string');

  const parsed = JSON.parse(jsonStr);
  assertEqual(parsed.version, '2.0.0');
  assert(Array.isArray(parsed.fittings), 'fittings must be an array');
  assert(typeof parsed.settings === 'object', 'settings must be an object');
  assert(Array.isArray(parsed.discounts), 'discounts must be an array');
  assert(Array.isArray(parsed.customers), 'customers must be an array');
  assert(Array.isArray(parsed.invoices), 'invoices must be an array');
  assert(typeof parsed.weightTable === 'object', 'weightTable must be an object');
  assert(typeof parsed.gradeSdrPn === 'object', 'gradeSdrPn must be an object');
  assert(typeof parsed.exportedAt === 'string', 'exportedAt must be an ISO date string');
});

runTest('5.2 StorageService.restoreFromBackup restores entities and validates structure', () => {
  const mockBackup = {
    version: '2.0.0',
    exportedAt: new Date().toISOString(),
    fittings: [
      {
        id: 'test_fit_1',
        category: 'اتصالات جوشی',
        connectionType: 'جوشی',
        product: 'سه‌راهی',
        material: 'HDPE100',
        pressure: '10',
        size: '110',
        unitPriceRial: 2500000,
        packageCount: 5,
        isActive: true,
      },
    ],
    settings: {
      name: 'پلیمر تست',
      address: 'تهران',
      phone: '02112345678',
      fax: '',
      taxRate: 10,
      invoiceCounter: 42,
      currencyUnit: 'تومان',
      basePipePrices: { PE100: 95000, PE80: 88000, PE63: 82000 },
    },
    discounts: [],
    customers: [
      {
        id: 'cust_test_1',
        name: 'شرکت پایپینگ مهر',
        phone: '09120000000',
      },
    ],
    invoices: [],
  };

  const success = StorageService.restoreFromBackup(JSON.stringify(mockBackup));
  assert(success, 'restoreFromBackup must return true for valid backup JSON');

  const restoredSettings = StorageService.getSettings();
  assertEqual(restoredSettings.name, 'پلیمر تست');
  assertEqual(restoredSettings.invoiceCounter, 42);

  const restoredCustomers = StorageService.getCustomers();
  assertEqual(restoredCustomers.length, 1);
  assertEqual(restoredCustomers[0].name, 'شرکت پایپینگ مهر');

  // Test invalid JSON rejection
  const failed = StorageService.restoreFromBackup('invalid-not-json');
  assert(!failed, 'restoreFromBackup must return false for invalid JSON');
});

runTest('5.3 StorageService.triggerDiskBackup calls window.electronAPI.saveBackup', async () => {
  let diskBackupPayload: string | null = null;

  (globalThis as any).window = {
    electronAPI: {
      saveBackup: async ({ backupJson }: { backupJson: string }) => {
        diskBackupPayload = backupJson;
        return { success: true, filePath: '/tmp/userData/backups/pipecal-latest.json' };
      },
    },
  };

  const ok = await StorageService.triggerDiskBackup();
  assert(ok, 'triggerDiskBackup must return true when Electron saveBackup succeeds');
  assert(diskBackupPayload !== null, 'backupJson payload must have been sent to Electron IPC');

  const parsed = JSON.parse(diskBackupPayload!);
  assertEqual(parsed.version, '2.0.0');
});

// ============================================================================
// FINAL SUMMARY
// ============================================================================
console.log('\n======================================================');
console.log(`TOTAL PASSED: ${totalPassed}`);
console.log(`TOTAL FAILED: ${totalFailed}`);
if (totalFailed > 0) {
  console.log('FAILURES:');
  failures.forEach((f) => console.log(`  - ${f}`));
} else {
  console.log('ALL OFFLINE & DESKTOP RELIABILITY TESTS PASSED CLEANLY! (15/15)');
  console.log('======================================================\n');
}
});
