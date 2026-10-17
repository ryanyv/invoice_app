import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';

describe('CHALLENGER 2 ADVERSARIAL TEST SUITE: M6 INTEGRATION & DESKTOP PACKAGING', () => {
  const rootDir = path.resolve(__dirname, '../../..');
  const pkgPath = path.join(rootDir, 'package.json');
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'));

  // -------------------------------------------------------------
  // SUITE 1: Git Branching & Merge Integrity
  // -------------------------------------------------------------
  describe('1. Git Integrity & Branch Ancestry', () => {
    it('1.1 main matches origin/main and contains the base ancestor', () => {
      const diffOutput = execSync('git diff main origin/main', { cwd: rootDir, encoding: 'utf-8' }).trim();
      expect(diffOutput).toBe('');

      const mainSha = execSync('git rev-parse main', { cwd: rootDir, encoding: 'utf-8' }).trim();
      const originMainSha = execSync('git rev-parse origin/main', { cwd: rootDir, encoding: 'utf-8' }).trim();
      expect(mainSha).toBe(originMainSha);
      expect(() => {
        execSync('git merge-base --is-ancestor f2366479a5de3e99220c3b04f32d92bdc2e51f78 main', { cwd: rootDir });
      }).not.toThrow();
    });

    it('1.2 all 5 feature branches are cleanly integrated into staging', () => {
      const featureBranches = [
        'feat/invoicing-workflow',
        'feat/pe-domain-calculator',
        'feat/print-export-sharing',
        'feat/offline-reliability',
        'feat/test-suite',
      ];

      for (const branch of featureBranches) {
        // Check merge-base ancestry: branch must be an ancestor of staging
        let isAncestor = false;
        try {
          execSync(`git merge-base --is-ancestor ${branch} staging`, { cwd: rootDir });
          isAncestor = true;
        } catch {
          isAncestor = false;
        }
        expect(isAncestor, `Branch ${branch} must be merged into staging`).toBe(true);

        // Check unmerged commits: git log branch ^staging must be empty
        const unmerged = execSync(`git log ${branch} ^staging --oneline`, { cwd: rootDir, encoding: 'utf-8' }).trim();
        expect(unmerged, `Branch ${branch} has unmerged commits`).toBe('');

        // Also check origin remote counterparts
        const originBranch = `origin/${branch}`;
        const originUnmerged = execSync(`git log ${originBranch} ^staging --oneline`, { cwd: rootDir, encoding: 'utf-8' }).trim();
        expect(originUnmerged, `Remote branch ${originBranch} has unmerged commits into staging`).toBe('');
      }
    });

    it('1.3 staging is in sync with origin/staging', () => {
      const stagingDiff = execSync('git diff staging origin/staging', { cwd: rootDir, encoding: 'utf-8' }).trim();
      expect(stagingDiff).toBe('');

      const stagingSha = execSync('git rev-parse staging', { cwd: rootDir, encoding: 'utf-8' }).trim();
      const originStagingSha = execSync('git rev-parse origin/staging', { cwd: rootDir, encoding: 'utf-8' }).trim();
      expect(stagingSha).toBe(originStagingSha);
    });
  });

  // -------------------------------------------------------------
  // SUITE 2: Package Configuration & electron-builder
  // -------------------------------------------------------------
  describe('2. Package Configuration & electron-builder Specs', () => {
    it('2.1 package.json specifies correct build metadata and appId', () => {
      expect(pkg.name).toBe('pipecal-invoice-pro');
      expect(pkg.version).toBe('2.1.0');
      expect(pkg.main).toBe('electron/main.cjs');
      expect(pkg.build).toBeDefined();
      expect(pkg.build.appId).toBe('com.polygharb.pipecal');
      expect(pkg.build.productName).toBe('Pipecal Pro');
      expect(pkg.build.directories.output).toBe('release');
    });

    it('2.2 package.json build.files explicitly packages dist, electron, and public', () => {
      const files = pkg.build.files;
      expect(Array.isArray(files)).toBe(true);
      expect(files).toContain('dist/**/*');
      expect(files).toContain('electron/**/*');
      expect(files).toContain('public/**/*');
    });

    it('2.3 package.json build configurations for mac and win targets are defined', () => {
      expect(pkg.build.mac).toBeDefined();
      expect(pkg.build.mac.target).toContain('dmg');
      expect(pkg.build.mac.target).toContain('dir');
      expect(pkg.build.mac.identity).toBeNull();

      expect(pkg.build.toolsets?.nsis).toBe('1.2.1');
      expect(pkg.build.win).toBeDefined();
      expect(pkg.build.win.target).toContain('nsis');
      expect(pkg.build.win.target).toContain('portable');
    });

    it('2.4 build and test scripts are properly configured', () => {
      expect(pkg.scripts.test).toBe('vitest run');
      expect(pkg.scripts.build).toBe('tsc && vite build');
      expect(pkg.scripts['electron:build:mac']).toContain('electron-builder --mac');
      expect(pkg.scripts['electron:build:win']).toContain('electron-builder --win --x64 --arm64');
    });
  });

  // -------------------------------------------------------------
  // SUITE 3: Complete Offline Typography & Zero Google Fonts Leak
  // -------------------------------------------------------------
  describe('3. Offline Fonts & Absence of External Google Fonts', () => {
    it('3.1 dist/index.html exists and contains zero Google Fonts references', () => {
      const distHtmlPath = path.join(rootDir, 'dist/index.html');
      expect(fs.existsSync(distHtmlPath)).toBe(true);
      const html = fs.readFileSync(distHtmlPath, 'utf-8');

      expect(html.includes('fonts.googleapis.com')).toBe(false);
      expect(html.includes('fonts.gstatic.com')).toBe(false);
      expect(html.includes('http://')).toBe(false);
      expect(html.includes('https://')).toBe(false);
    });

    it('3.2 index.html source contains zero Google Fonts references', () => {
      const indexHtmlPath = path.join(rootDir, 'index.html');
      const html = fs.readFileSync(indexHtmlPath, 'utf-8');
      expect(html.includes('fonts.googleapis.com')).toBe(false);
      expect(html.includes('fonts.gstatic.com')).toBe(false);
    });

    it('3.3 dist/assets/ CSS and JS bundles contain zero Google Fonts references', () => {
      const assetsDir = path.join(rootDir, 'dist/assets');
      expect(fs.existsSync(assetsDir)).toBe(true);
      const files = fs.readdirSync(assetsDir);
      expect(files.length).toBeGreaterThan(0);

      for (const file of files) {
        if (file.endsWith('.js') || file.endsWith('.css')) {
          const content = fs.readFileSync(path.join(assetsDir, file), 'utf-8');
          expect(content.includes('fonts.googleapis.com'), `${file} must not contain fonts.googleapis.com`).toBe(false);
          expect(content.includes('fonts.gstatic.com'), `${file} must not contain fonts.gstatic.com`).toBe(false);
        }
      }
    });

    it('3.4 public/fonts/ and dist/fonts/ contain 4 valid woff2 font files with wOF2 magic header', () => {
      const fontFiles = [
        'Vazirmatn[wght].woff2',
        'vazirmatn-arabic-wght-normal.woff2',
        'vazirmatn-latin-ext-wght-normal.woff2',
        'vazirmatn-latin-wght-normal.woff2',
      ];

      for (const font of fontFiles) {
        const publicFont = path.join(rootDir, 'public/fonts', font);
        const distFont = path.join(rootDir, 'dist/fonts', font);

        expect(fs.existsSync(publicFont), `public/fonts/${font} missing`).toBe(true);
        expect(fs.existsSync(distFont), `dist/fonts/${font} missing`).toBe(true);

        const pubBuf = fs.readFileSync(publicFont);
        const distBuf = fs.readFileSync(distFont);

        // Minimum file size check (>20KB)
        expect(pubBuf.length).toBeGreaterThan(20000);
        expect(distBuf.length).toBeGreaterThan(20000);

        // Magic header check: 'wOF2' (0x77 0x4F 0x46 0x32)
        const magicPub = pubBuf.subarray(0, 4).toString('ascii');
        const magicDist = distBuf.subarray(0, 4).toString('ascii');
        expect(magicPub).toBe('wOF2');
        expect(magicDist).toBe('wOF2');
      }
    });
  });

  // -------------------------------------------------------------
  // SUITE 4: Packaged Application Bundle Content & Asar Verification
  // -------------------------------------------------------------
  describe('4. Packaged Application Bundles & Asar Verification', () => {
    it('4.1 macOS arm64 DMG and .app bundle exist and satisfy size requirements', () => {
      const dmgPath = path.join(rootDir, `release/Pipecal Pro-${pkg.version}-arm64.dmg`);
      const appPath = path.join(rootDir, 'release/mac-arm64/Pipecal Pro.app');
      const macBinary = path.join(appPath, 'Contents/MacOS/Pipecal Pro');

      expect(fs.existsSync(dmgPath)).toBe(true);
      expect(fs.existsSync(appPath)).toBe(true);
      expect(fs.existsSync(macBinary)).toBe(true);

      const dmgStat = fs.statSync(dmgPath);
      expect(dmgStat.size).toBeGreaterThan(100 * 1024 * 1024); // > 100MB

      // Mach-O binary inspection
      const fileInfo = execSync(`file "${macBinary}"`, { encoding: 'utf-8' });
      expect(fileInfo).toContain('Mach-O 64-bit executable arm64');
    });

    it('4.2 Windows Setup installer, Portable exe, win-unpacked (x64), and win-arm64-unpacked exist', () => {
      const setupPath = path.join(rootDir, `release/Pipecal Pro Setup ${pkg.version}.exe`);
      const portablePath = path.join(rootDir, `release/Pipecal Pro ${pkg.version}.exe`);
      const unpackedExe = path.join(rootDir, 'release/win-unpacked/Pipecal Pro.exe');
      const unpackedArm64Exe = path.join(rootDir, 'release/win-arm64-unpacked/Pipecal Pro.exe');

      expect(fs.existsSync(setupPath)).toBe(true);
      expect(fs.existsSync(portablePath)).toBe(true);
      expect(fs.existsSync(unpackedExe)).toBe(true);
      expect(fs.existsSync(unpackedArm64Exe)).toBe(true);

      expect(fs.statSync(setupPath).size).toBeGreaterThan(200 * 1024 * 1024);
      expect(fs.statSync(portablePath).size).toBeGreaterThan(200 * 1024 * 1024);
      expect(fs.statSync(unpackedExe).size).toBeGreaterThan(200 * 1024 * 1024);
      expect(fs.statSync(unpackedArm64Exe).size).toBeGreaterThan(200 * 1024 * 1024);

      // Verify PE binary format
      const setupFileInfo = execSync(`file "${setupPath}"`, { encoding: 'utf-8' });
      expect(setupFileInfo).toContain('PE32 executable');

      const unpackedFileInfo = execSync(`file "${unpackedExe}"`, { encoding: 'utf-8' });
      expect(unpackedFileInfo).toContain('PE32+ executable (GUI) x86-64');

      const unpackedArm64FileInfo = execSync(`file "${unpackedArm64Exe}"`, { encoding: 'utf-8' });
      expect(unpackedArm64FileInfo).toContain('PE32+ executable (GUI) Aarch64');
    });

    it('4.3 macOS app.asar bundle contains dist/, electron/, and public/ offline fonts', () => {
      const asarPath = path.join(rootDir, 'release/mac-arm64/Pipecal Pro.app/Contents/Resources/app.asar');
      expect(fs.existsSync(asarPath)).toBe(true);

      const asarList = execSync(`npx asar list "${asarPath}"`, { encoding: 'utf-8' });
      const paths = asarList.split('\n');

      expect(paths).toContain('/dist/index.html');
      expect(paths).toContain('/electron/main.cjs');
      expect(paths).toContain('/electron/preload.cjs');
      expect(paths).toContain('/public/fonts/Vazirmatn[wght].woff2');
      expect(paths).toContain('/public/fonts/vazirmatn-arabic-wght-normal.woff2');
      expect(paths).toContain('/dist/fonts/Vazirmatn[wght].woff2');
      expect(paths).toContain('/dist/fonts/vazirmatn-arabic-wght-normal.woff2');
    });

    it('4.4 Windows win-unpacked app.asar contains dist/, electron/, and public/ offline fonts', () => {
      const asarPath = path.join(rootDir, 'release/win-unpacked/resources/app.asar');
      expect(fs.existsSync(asarPath)).toBe(true);

      const asarList = execSync(`npx asar list "${asarPath}"`, { encoding: 'utf-8' });
      const paths = asarList.split('\n');

      expect(paths).toContain('/dist/index.html');
      expect(paths).toContain('/electron/main.cjs');
      expect(paths).toContain('/electron/preload.cjs');
      expect(paths).toContain('/public/fonts/Vazirmatn[wght].woff2');
      expect(paths).toContain('/public/fonts/vazirmatn-arabic-wght-normal.woff2');
      expect(paths).toContain('/dist/fonts/Vazirmatn[wght].woff2');
      expect(paths).toContain('/dist/fonts/vazirmatn-arabic-wght-normal.woff2');
    });
  });

  // -------------------------------------------------------------
  // SUITE 5: Electron Main Process & Preload IPC Contract
  // -------------------------------------------------------------
  describe('5. Electron Main & Preload IPC Contract', () => {
    it('5.1 electron/main.cjs defines all required IPC handlers', () => {
      const mainPath = path.join(rootDir, 'electron/main.cjs');
      const content = fs.readFileSync(mainPath, 'utf-8');

      expect(content).toContain("ipcMain.handle('save-pdf'");
      expect(content).toContain("ipcMain.handle('export-file'");
      expect(content).toContain("ipcMain.handle('import-file'");
      expect(content).toContain("ipcMain.handle('save-backup'");
      expect(content).toContain("ipcMain.handle('list-backups'");
      expect(content).toContain("ipcMain.handle('load-backup'");
      expect(content).toContain("path.join(__dirname, '../dist/index.html')");
    });

    it('5.2 electron/preload.cjs exposes complete electronAPI interface', () => {
      const preloadPath = path.join(rootDir, 'electron/preload.cjs');
      const content = fs.readFileSync(preloadPath, 'utf-8');

      expect(content).toContain("contextBridge.exposeInMainWorld('electronAPI'");
      expect(content).toContain('savePdf:');
      expect(content).toContain('exportFile:');
      expect(content).toContain('importFile:');
      expect(content).toContain('saveBackup:');
      expect(content).toContain('listBackups:');
      expect(content).toContain('loadBackup:');
      expect(content).toContain('isElectron: true');
    });
  });
});
