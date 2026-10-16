/// <reference types="node" />
/**
 * CHALLENGER EMPIRICAL STRESS TEST HARNESS: Requirement R4 (Offline & Desktop Reliability)
 * Milestone M4 Verification
 *
 * Requirements Tested:
 * 1. Static Grep Inspection: 0 occurrences of fonts.googleapis.com / fonts.gstatic.com
 * 2. Font File Integrity: 4 .woff2 files in public/fonts/ are valid, readable, and non-empty (WOFF2 magic check)
 * 3. Keyboard Shortcut Event Simulation: Synthetic KeyboardEvents for Ctrl/Cmd+S, P, N, Escape (enabled vs disabled)
 * 4. Rolling Backup Retention Stress: 50 sequential backup entries pruned strictly to 20 newest files
 */

import fs from 'fs';
import path from 'path';
import os from 'os';
import {
  pruneBackups,
  BACKUP_RETENTION_LIMIT,
  generateBackupFilename,
  parseBackupTimestamp,
  formatBackupSize,
} from '../backupUtils';
import {
  SHORTCUT_DEFINITIONS,
  matchKeyboardShortcut,
  handleKeyboardEvent,
  ShortcutHandlers,
} from '../../hooks/useKeyboardShortcuts';
import { StorageService } from '../storage';

// ============================================================================
// TEST RUNNER INFRASTRUCTURE
// ============================================================================

let passedCount = 0;
let failedCount = 0;
const failureDetails: string[] = [];

function assert(condition: boolean, message?: string): asserts condition {
  if (!condition) {
    throw new Error(`Assertion Failed: ${message || 'condition not met'}`);
  }
}

function assertEqual<T>(actual: T, expected: T, message?: string) {
  if (actual !== expected) {
    throw new Error(
      `Assertion Failed: expected ${JSON.stringify(expected)} but got ${JSON.stringify(actual)}. ${message || ''}`
    );
  }
}

function runSyncTest(name: string, fn: () => void) {
  try {
    fn();
    passedCount++;
    console.log(`  ✓ PASS: ${name}`);
  } catch (err: any) {
    failedCount++;
    failureDetails.push(`${name} -> ${err.message}`);
    console.error(`  ✗ FAIL: ${name} -> ${err.message}`);
    throw err;
  }
}

async function runAsyncTest(name: string, fn: () => Promise<void>) {
  try {
    await fn();
    passedCount++;
    console.log(`  ✓ PASS: ${name}`);
  } catch (err: any) {
    failedCount++;
    failureDetails.push(`${name} -> ${err.message}`);
    console.error(`  ✗ FAIL: ${name} -> ${err.message}`);
    throw err;
  }
}

const projectRoot = path.resolve(__dirname, '../../..');

describe('CHALLENGER EMPIRICAL VERIFICATION HARNESS (Milestone M4: R4)', () => {
  it('runs all offline stress tests', async () => {
  console.log('\n================================================================');
  console.log('CHALLENGER EMPIRICAL VERIFICATION HARNESS (Milestone M4: R4)');
  console.log('================================================================\n');

  // ============================================================================
  // PART 1: STATIC GREP INSPECTION ACROSS ALL PROJECT FILES
  // ============================================================================
  console.log('--- 1. Static Grep Inspection (Zero External Google Fonts CDN) ---');

  runSyncTest('1.1 Zero occurrences of fonts.googleapis.com across all source and html files', () => {
    const ignoreDirs = new Set(['.git', 'node_modules', 'dist', '.agents']);
    const cdnTargets = ['fonts.googleapis.com', 'fonts.gstatic.com'];

    function scanDir(dir: string): string[] {
      const findings: string[] = [];
      const entries = fs.readdirSync(dir, { withFileTypes: true });

      for (const entry of entries) {
        if (entry.isDirectory()) {
          if (!ignoreDirs.has(entry.name)) {
            findings.push(...scanDir(path.join(dir, entry.name)));
          }
        } else if (entry.isFile()) {
          const fullPath = path.join(dir, entry.name);
          const relPath = path.relative(projectRoot, fullPath);

          // Skip test assertion files that contain the query string as a test assert target
          if (relPath.includes('__tests__')) {
            continue;
          }

          const ext = path.extname(entry.name).toLowerCase();
          if (['.html', '.css', '.ts', '.tsx', '.js', '.jsx', '.cjs', '.mjs', '.json'].includes(ext)) {
            const content = fs.readFileSync(fullPath, 'utf-8');
            for (const target of cdnTargets) {
              if (content.includes(target)) {
                findings.push(`${relPath} contains prohibited CDN reference: "${target}"`);
              }
            }
          }
        }
      }
      return findings;
    }

    const forbiddenMatches = scanDir(projectRoot);
    assertEqual(
      forbiddenMatches.length,
      0,
      `Found forbidden external font CDN references:\n${forbiddenMatches.join('\n')}`
    );
  });

  runSyncTest('1.2 index.html contains no external stylesheet or preconnect links', () => {
    const indexHtmlPath = path.join(projectRoot, 'index.html');
    assert(fs.existsSync(indexHtmlPath), 'index.html must exist at project root');
    const html = fs.readFileSync(indexHtmlPath, 'utf-8');

    assert(!html.includes('fonts.googleapis.com'), 'index.html must not contain fonts.googleapis.com');
    assert(!html.includes('fonts.gstatic.com'), 'index.html must not contain fonts.gstatic.com');
    assert(!html.includes('href="http'), 'index.html must not link to any external HTTP/HTTPS stylesheet');
  });

  runSyncTest('1.3 src/index.css references strictly local /fonts/ assets', () => {
    const cssPath = path.join(projectRoot, 'src/index.css');
    assert(fs.existsSync(cssPath), 'src/index.css must exist');
    const css = fs.readFileSync(cssPath, 'utf-8');

    assert(css.includes('@font-face'), 'src/index.css must declare @font-face');
    assert(css.includes("font-family: 'Vazirmatn'"), 'src/index.css must specify Vazirmatn font-family');
    assert(css.includes("url('/fonts/"), 'src/index.css @font-face must link to local /fonts/ directory');
  });

  // ============================================================================
  // PART 2: FONT FILE INTEGRITY & WOFF2 BINARY HEADER VERIFICATION
  // ============================================================================
  console.log('\n--- 2. Font File Integrity & WOFF2 Binary Magic Inspection ---');

  const requiredWoff2Files = [
    'Vazirmatn[wght].woff2',
    'vazirmatn-arabic-wght-normal.woff2',
    'vazirmatn-latin-wght-normal.woff2',
    'vazirmatn-latin-ext-wght-normal.woff2',
  ];

  runSyncTest('2.1 All 4 expected .woff2 font files exist in public/fonts/', () => {
    const fontsDir = path.join(projectRoot, 'public/fonts');
    assert(fs.existsSync(fontsDir), 'public/fonts directory must exist');

    for (const fontName of requiredWoff2Files) {
      const fullPath = path.join(fontsDir, fontName);
      assert(fs.existsSync(fullPath), `Font file ${fontName} must exist in public/fonts/`);
      const stat = fs.statSync(fullPath);
      assert(stat.isFile(), `${fontName} must be a regular file`);
    }
  });

  runSyncTest('2.2 All 4 font files are non-empty and satisfy size thresholds (>20KB)', () => {
    const fontsDir = path.join(projectRoot, 'public/fonts');

    for (const fontName of requiredWoff2Files) {
      const fullPath = path.join(fontsDir, fontName);
      const size = fs.statSync(fullPath).size;
      assert(size > 20000, `Font file ${fontName} must be > 20,000 bytes (got ${size} bytes)`);
    }
  });

  runSyncTest('2.3 All 4 font files contain valid WOFF2 binary headers (magic "wOF2" + valid length)', () => {
    const fontsDir = path.join(projectRoot, 'public/fonts');

    for (const fontName of requiredWoff2Files) {
      const fullPath = path.join(fontsDir, fontName);
      const buffer = fs.readFileSync(fullPath);
      const statSize = fs.statSync(fullPath).size;

      assert(buffer.length >= 48, `${fontName} header must be at least 48 bytes`);

      // Byte 0-3: Magic number 'wOF2' (0x77, 0x4F, 0x46, 0x32)
      const magic = buffer.toString('ascii', 0, 4);
      assertEqual(magic, 'wOF2', `${fontName} must begin with WOFF2 magic bytes 'wOF2'`);

      // Byte 8-11: Length of WOFF2 file (uint32 big-endian)
      const headerLength = buffer.readUInt32BE(8);
      assertEqual(
        headerLength,
        statSize,
        `${fontName} header length (${headerLength}) must match exact disk size (${statSize})`
      );

      // Byte 12-13: numTables (uint16 big-endian)
      const numTables = buffer.readUInt16BE(12);
      assert(numTables > 0, `${fontName} numTables must be > 0 (got ${numTables})`);

      // Byte 16-19: totalSfntSize (uint32 big-endian)
      const totalSfntSize = buffer.readUInt32BE(16);
      assert(totalSfntSize > 0, `${fontName} totalSfntSize must be > 0`);

      // Byte 20-23: totalCompressedSize (uint32 big-endian)
      const totalCompressedSize = buffer.readUInt32BE(20);
      assert(totalCompressedSize > 0, `${fontName} totalCompressedSize must be > 0`);
    }
  });

  // ============================================================================
  // PART 3: KEYBOARD SHORTCUT EVENT SIMULATION (ENABLED VS DISABLED)
  // ============================================================================
  console.log('\n--- 3. Keyboard Shortcut Event Simulation (Enabled vs Disabled) ---');

  runSyncTest('3.1 Synthetic KeyboardEvents for Ctrl+S and Cmd+S trigger onSave and preventDefault', () => {
    let saveCount = 0;
    let preventCount = 0;

    const handlers: ShortcutHandlers = {
      onSave: () => {
        saveCount++;
      },
    };

    // Ctrl+s
    const ctrlSEvent = {
      key: 's',
      ctrlKey: true,
      metaKey: false,
      preventDefault: () => {
        preventCount++;
      },
    };
    const handled1 = handleKeyboardEvent(ctrlSEvent, handlers);
    assert(handled1, 'Ctrl+s must be handled');
    assertEqual(saveCount, 1);
    assertEqual(preventCount, 1);

    // Cmd+s (macOS)
    const cmdSEvent = {
      key: 's',
      ctrlKey: false,
      metaKey: true,
      preventDefault: () => {
        preventCount++;
      },
    };
    const handled2 = handleKeyboardEvent(cmdSEvent, handlers);
    assert(handled2, 'Cmd+s must be handled');
    assertEqual(saveCount, 2);
    assertEqual(preventCount, 2);

    // Uppercase Ctrl+S
    const ctrlUpperS = {
      key: 'S',
      ctrlKey: true,
      metaKey: false,
      preventDefault: () => {
        preventCount++;
      },
    };
    const handled3 = handleKeyboardEvent(ctrlUpperS, handlers);
    assert(handled3, 'Ctrl+S must be handled');
    assertEqual(saveCount, 3);
    assertEqual(preventCount, 3);
  });

  runSyncTest('3.2 Synthetic KeyboardEvents for Ctrl+P and Cmd+P trigger onPrint and preventDefault', () => {
    let printCount = 0;
    let preventCount = 0;

    const handlers: ShortcutHandlers = {
      onPrint: () => {
        printCount++;
      },
    };

    // Ctrl+p
    const handled1 = handleKeyboardEvent(
      {
        key: 'p',
        ctrlKey: true,
        metaKey: false,
        preventDefault: () => {
          preventCount++;
        },
      },
      handlers
    );
    assert(handled1, 'Ctrl+p must be handled');
    assertEqual(printCount, 1);
    assertEqual(preventCount, 1);

    // Cmd+P (macOS)
    const handled2 = handleKeyboardEvent(
      {
        key: 'P',
        ctrlKey: false,
        metaKey: true,
        preventDefault: () => {
          preventCount++;
        },
      },
      handlers
    );
    assert(handled2, 'Cmd+P must be handled');
    assertEqual(printCount, 2);
    assertEqual(preventCount, 2);
  });

  runSyncTest('3.3 Synthetic KeyboardEvents for Ctrl+N and Cmd+N trigger onNewDraft and preventDefault', () => {
    let newDraftCount = 0;
    let preventCount = 0;

    const handlers: ShortcutHandlers = {
      onNewDraft: () => {
        newDraftCount++;
      },
    };

    // Ctrl+n
    const handled1 = handleKeyboardEvent(
      {
        key: 'n',
        ctrlKey: true,
        metaKey: false,
        preventDefault: () => {
          preventCount++;
        },
      },
      handlers
    );
    assert(handled1, 'Ctrl+n must be handled');
    assertEqual(newDraftCount, 1);
    assertEqual(preventCount, 1);

    // Cmd+N (macOS)
    const handled2 = handleKeyboardEvent(
      {
        key: 'N',
        ctrlKey: false,
        metaKey: true,
        preventDefault: () => {
          preventCount++;
        },
      },
      handlers
    );
    assert(handled2, 'Cmd+N must be handled');
    assertEqual(newDraftCount, 2);
    assertEqual(preventCount, 2);
  });

  runSyncTest('3.4 Synthetic KeyboardEvents for Escape trigger onEscape and preventDefault', () => {
    let escapeCount = 0;
    let preventCount = 0;

    const handlers: ShortcutHandlers = {
      onEscape: () => {
        escapeCount++;
      },
    };

    const handled = handleKeyboardEvent(
      {
        key: 'Escape',
        ctrlKey: false,
        metaKey: false,
        preventDefault: () => {
          preventCount++;
        },
      },
      handlers
    );
    assert(handled, 'Escape must be handled');
    assertEqual(escapeCount, 1);
    assertEqual(preventCount, 1);
  });

  runSyncTest('3.5 DISABLED STATE: Unregistered handlers return false and do NOT call preventDefault', () => {
    let preventCount = 0;
    const emptyHandlers: ShortcutHandlers = {};

    const createEvent = (key: string, ctrl = false, meta = false) => ({
      key,
      ctrlKey: ctrl,
      metaKey: meta,
      preventDefault: () => {
        preventCount++;
      },
    });

    // Ctrl+S without onSave
    assert(!handleKeyboardEvent(createEvent('s', true), emptyHandlers), 'Must return false when onSave is disabled');
    assertEqual(preventCount, 0, 'Must NOT preventDefault when onSave is disabled');

    // Cmd+P without onPrint
    assert(!handleKeyboardEvent(createEvent('p', false, true), emptyHandlers), 'Must return false when onPrint is disabled');
    assertEqual(preventCount, 0, 'Must NOT preventDefault when onPrint is disabled');

    // Ctrl+N without onNewDraft
    assert(!handleKeyboardEvent(createEvent('n', true), emptyHandlers), 'Must return false when onNewDraft is disabled');
    assertEqual(preventCount, 0, 'Must NOT preventDefault when onNewDraft is disabled');

    // Escape without onEscape
    assert(!handleKeyboardEvent(createEvent('Escape'), emptyHandlers), 'Must return false when onEscape is disabled');
    assertEqual(preventCount, 0, 'Must NOT preventDefault when onEscape is disabled');
  });

  runSyncTest('3.6 Regular typing without Ctrl/Cmd does NOT trigger shortcuts or preventDefault', () => {
    let calledCount = 0;
    let preventCount = 0;

    const fullHandlers: ShortcutHandlers = {
      onSave: () => calledCount++,
      onPrint: () => calledCount++,
      onNewDraft: () => calledCount++,
      onEscape: () => calledCount++,
    };

    const normalKeys = ['s', 'S', 'p', 'P', 'n', 'N', 'a', 'b', '1', ' '];
    for (const k of normalKeys) {
      const handled = handleKeyboardEvent(
        {
          key: k,
          ctrlKey: false,
          metaKey: false,
          preventDefault: () => {
            preventCount++;
          },
        },
        fullHandlers
      );
      assert(!handled, `Key "${k}" without modifiers must NOT be handled`);
      assertEqual(calledCount, 0, `Key "${k}" must not trigger any handler`);
      assertEqual(preventCount, 0, `Key "${k}" must not call preventDefault`);
    }
  });

  runSyncTest('3.7 Unmapped Ctrl/Cmd combos (Ctrl+C, Ctrl+V, Ctrl+Z) pass through unhandled', () => {
    let calledCount = 0;
    let preventCount = 0;

    const fullHandlers: ShortcutHandlers = {
      onSave: () => calledCount++,
      onPrint: () => calledCount++,
      onNewDraft: () => calledCount++,
      onEscape: () => calledCount++,
    };

    const unmappedCombos = ['c', 'v', 'x', 'z', 'a', 'r', 'f', 'F12'];
    for (const k of unmappedCombos) {
      const handled = handleKeyboardEvent(
        {
          key: k,
          ctrlKey: true,
          metaKey: false,
          preventDefault: () => {
            preventCount++;
          },
        },
        fullHandlers
      );
      assert(!handled, `Ctrl+${k} must NOT be intercepted`);
      assertEqual(calledCount, 0);
      assertEqual(preventCount, 0);
    }
  });

  runSyncTest('3.8 Edge case inputs: undefined key, empty key, non-standard event objects', () => {
    assertEqual(matchKeyboardShortcut({ key: '' }), null);
    assertEqual(matchKeyboardShortcut({ key: undefined as any }), null);
    assertEqual(matchKeyboardShortcut({} as any), null);
  });

  // ============================================================================
  // PART 4: ROLLING BACKUP RETENTION STRESS (50 SEQUENTIAL ENTRIES -> EXACTLY 20)
  // ============================================================================
  console.log('\n--- 4. Rolling Backup Retention Stress (50 Entries Pruned to 20) ---');

  runSyncTest('4.1 Generating 50 sequential backup entries and pruning strictly retains the 20 newest', () => {
    // Generate 50 sequential backups from t=1000 to t=50000
    const totalEntries = 50;
    const mockBackups = Array.from({ length: totalEntries }, (_, i) => {
      const id = i + 1;
      const padId = String(id).padStart(2, '0');
      return {
        filename: `pipecal-backup-2026-10-16-1000${padId}.json`,
        mtimeMs: 1000000 + id * 1000, // item 1 is oldest (1001000), item 50 is newest (1050000)
        size: 5000 + id * 100,
      };
    });

    assertEqual(mockBackups.length, 50, 'Must have created 50 mock backups');

    // Prune with retention limit 20
    const toDelete = pruneBackups(mockBackups, 20);

    // Assert exactly 30 entries selected for deletion
    assertEqual(toDelete.length, 30, 'Pruning 50 backups with limit 20 must return exactly 30 items to delete');

    // Collect retained entries
    const deletedFilenames = new Set(toDelete.map((b) => b.filename));
    const retainedEntries = mockBackups.filter((b) => !deletedFilenames.has(b.filename));

    assertEqual(retainedEntries.length, 20, 'Retained backups count must be strictly 20');

    // Verify that the 30 deleted files are strictly the 30 oldest (id 1 to 30)
    for (let id = 1; id <= 30; id++) {
      const expectedOldFilename = `pipecal-backup-2026-10-16-1000${String(id).padStart(2, '0')}.json`;
      assert(
        deletedFilenames.has(expectedOldFilename),
        `Old file ${expectedOldFilename} (id ${id}) MUST be marked for deletion`
      );
    }

    // Verify that the 20 retained files are strictly the 20 newest (id 31 to 50)
    for (let id = 31; id <= 50; id++) {
      const expectedNewFilename = `pipecal-backup-2026-10-16-1000${String(id).padStart(2, '0')}.json`;
      assert(
        !deletedFilenames.has(expectedNewFilename),
        `New file ${expectedNewFilename} (id ${id}) MUST be retained`
      );
    }

    // Verify retained entries are strictly greater in timestamp than any deleted entry
    const minRetainedMtime = Math.min(...retainedEntries.map((b) => b.mtimeMs));
    const maxDeletedMtime = Math.max(...toDelete.map((b) => b.mtimeMs));
    assert(
      minRetainedMtime > maxDeletedMtime,
      `All retained backups (${minRetainedMtime}) must be strictly newer than all deleted backups (${maxDeletedMtime})`
    );
  });

  runSyncTest('4.2 Shuffled order input resilience: pruneBackups strictly sorts and preserves 20 newest', () => {
    const totalEntries = 50;
    const mockBackups = Array.from({ length: totalEntries }, (_, i) => ({
      filename: `pipecal-backup-shuffled-${i + 1}.json`,
      mtimeMs: 1000000 + (i + 1) * 500,
    }));

    // Deterministically shuffle array
    const shuffled = [...mockBackups].sort((a, b) => {
      return ((a.mtimeMs * 37) % 101) - ((b.mtimeMs * 37) % 101);
    });

    const toDelete = pruneBackups(shuffled, 20);
    assertEqual(toDelete.length, 30);

    const deletedSet = new Set(toDelete.map((b) => b.filename));
    const retained = shuffled.filter((b) => !deletedSet.has(b.filename));
    assertEqual(retained.length, 20);

    // Newest 20 should all have mtimeMs > the 30th item
    const sortedOriginal = [...mockBackups].sort((a, b) => b.mtimeMs - a.mtimeMs);
    const expectedRetainedFilenames = new Set(sortedOriginal.slice(0, 20).map((b) => b.filename));

    for (const item of retained) {
      assert(
        expectedRetainedFilenames.has(item.filename),
        `Retained item ${item.filename} must belong to top 20 newest`
      );
    }
  });

  runSyncTest('4.3 Identical mtime collision resilience: pruneBackups cleanly retains 20 items', () => {
    const collisions = Array.from({ length: 50 }, (_, i) => ({
      filename: `pipecal-backup-collision-${i + 1}.json`,
      mtimeMs: 5000000, // all have exact same timestamp
    }));

    const toDelete = pruneBackups(collisions, 20);
    assertEqual(toDelete.length, 30, 'Must prune 30 items even if all timestamps are identical');
    const remaining = collisions.filter((c) => !toDelete.includes(c));
    assertEqual(remaining.length, 20, 'Must retain 20 items');
  });

  runSyncTest('4.4 Boundary test cases: 0, 1, 19, 20, 21 items', () => {
    assertEqual(pruneBackups([], 20).length, 0);

    const one = [{ filename: '1.json', mtimeMs: 100 }];
    assertEqual(pruneBackups(one, 20).length, 0);

    const nineteen = Array.from({ length: 19 }, (_, i) => ({ filename: `${i}.json`, mtimeMs: i }));
    assertEqual(pruneBackups(nineteen, 20).length, 0);

    const twenty = Array.from({ length: 20 }, (_, i) => ({ filename: `${i}.json`, mtimeMs: i }));
    assertEqual(pruneBackups(twenty, 20).length, 0);

    const twentyOne = Array.from({ length: 21 }, (_, i) => ({ filename: `${i}.json`, mtimeMs: i }));
    const pruned21 = pruneBackups(twentyOne, 20);
    assertEqual(pruned21.length, 1);
    assertEqual(pruned21[0].filename, '0.json', 'The oldest file (index 0 with mtimeMs 0) must be pruned');
  });

  await runAsyncTest('4.5 Real filesystem stress simulation: 50 physical JSON backup files pruned down to 20', async () => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pipecal-stress-test-'));

    try {
      const fileEntries: { filename: string; filePath: string; mtimeMs: number }[] = [];

      // Create 50 actual JSON backup files on disk
      for (let i = 1; i <= 50; i++) {
        const pad = String(i).padStart(2, '0');
        const filename = `pipecal-backup-2026-10-16-1200${pad}.json`;
        const filePath = path.join(tempDir, filename);
        const content = JSON.stringify({ version: '2.0.0', index: i, timestamp: Date.now() });

        fs.writeFileSync(filePath, content, 'utf-8');

        // Set explicit mtime to ensure exact millisecond control
        const targetTime = 1700000000 + i * 10;
        fs.utimesSync(filePath, targetTime, targetTime);

        const stat = fs.statSync(filePath);
        fileEntries.push({ filename, filePath, mtimeMs: stat.mtimeMs });
      }

      assertEqual(fs.readdirSync(tempDir).length, 50, 'Initial temp directory must contain 50 files');

      // Execute exact Electron main process pruning algorithm
      const diskBackups = fs
        .readdirSync(tempDir)
        .filter((f) => f.startsWith('pipecal-backup-') && f.endsWith('.json'))
        .map((f) => {
          const fullPath = path.join(tempDir, f);
          const stat = fs.statSync(fullPath);
          return { filename: f, filePath: fullPath, mtimeMs: stat.mtimeMs };
        })
        .sort((a, b) => b.mtimeMs - a.mtimeMs);

      assertEqual(diskBackups.length, 50);

      if (diskBackups.length > BACKUP_RETENTION_LIMIT) {
        const toPrune = diskBackups.slice(BACKUP_RETENTION_LIMIT);
        assertEqual(toPrune.length, 30);

        for (const item of toPrune) {
          if (fs.existsSync(item.filePath)) {
            fs.unlinkSync(item.filePath);
          }
        }
      }

      // Verify remaining files on disk
      const remainingOnDisk = fs.readdirSync(tempDir);
      assertEqual(
        remainingOnDisk.length,
        20,
        `After pruning, exactly 20 files must remain on disk (got ${remainingOnDisk.length})`
      );

      // Verify that the remaining 20 files correspond to index 31 through 50
      for (let i = 31; i <= 50; i++) {
        const pad = String(i).padStart(2, '0');
        const expectedFile = `pipecal-backup-2026-10-16-1200${pad}.json`;
        assert(
          remainingOnDisk.includes(expectedFile),
          `Newer file ${expectedFile} must exist on disk after pruning`
        );
      }

      // Verify that indices 1 through 30 were unlinked
      for (let i = 1; i <= 30; i++) {
        const pad = String(i).padStart(2, '0');
        const deletedFile = `pipecal-backup-2026-10-16-1200${pad}.json`;
        assert(
          !remainingOnDisk.includes(deletedFile),
          `Older file ${deletedFile} must have been deleted from disk`
        );
      }
    } finally {
      // Cleanup temporary directory
      if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    }
  });

  await runAsyncTest('4.6 StorageService.triggerDiskBackup gracefully handles non-Electron environment', async () => {
    // When window is undefined or electronAPI is undefined, must return false without throwing
    const origWindow = (globalThis as any).window;
    try {
      (globalThis as any).window = undefined;
      const res1 = await StorageService.triggerDiskBackup();
      assertEqual(res1, false);

      (globalThis as any).window = {};
      const res2 = await StorageService.triggerDiskBackup();
      assertEqual(res2, false);

      (globalThis as any).window = { electronAPI: {} };
      const res3 = await StorageService.triggerDiskBackup();
      assertEqual(res3, false);
    } finally {
      (globalThis as any).window = origWindow;
    }
  });

  // ============================================================================
  // FINAL HARNESS VERDICT REPORT
  // ============================================================================
  console.log('\n================================================================');
  console.log(`CHALLENGER STRESS TESTS SUMMARY:`);
  console.log(`  PASSED: ${passedCount}`);
  console.log(`  FAILED: ${failedCount}`);

  if (failedCount > 0) {
    console.log('\nFAILURES IDENTIFIED:');
    failureDetails.forEach((f) => console.log(`  ✗ ${f}`));
    console.log('================================================================\n');
    console.log('VERDICT: REQUEST_CHANGES');
    throw new Error(`${failedCount} challenger offline stress tests failed`);
  } else {
    console.log('================================================================\n');
    console.log(`ALL CHALLENGER EMPIRICAL STRESS TESTS PASSED CLEANLY (${passedCount}/${passedCount})!`);
    console.log('VERDICT: APPROVE');
  }
  });
});
