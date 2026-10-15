const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1300,
    height: 880,
    minWidth: 1024,
    minHeight: 700,
    title: 'سیستم صدور پیش‌فاکتور و مدیریت قیمت (پلی غرب)',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
    },
    show: false,
    backgroundColor: '#0f172a',
  });

  mainWindow.webContents.on('did-fail-load', (e, code, desc, url) => {
    console.error('[did-fail-load]', code, desc, url);
  });
  mainWindow.webContents.on('console-message', (e, level, message, line, sourceId) => {
    console.log(`[Renderer Console] [${level}] ${message} (${sourceId}:${line})`);
  });

  mainWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'F12' || (input.key.toLowerCase() === 'i' && (input.control || input.meta) && (input.shift || input.alt))) {
      mainWindow.webContents.toggleDevTools();
    }
  });

  const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged;
  console.log('[Main Process] isDev:', isDev, 'app.isPackaged:', app.isPackaged, '__dirname:', __dirname);

  if (isDev) {
    const devUrl = process.env.VITE_DEV_SERVER_URL || 'http://localhost:5173';
    console.log('[Main Process] Loading devUrl:', devUrl);
    const loadUrlWithRetry = () => {
      mainWindow.loadURL(devUrl).catch(() => {
        setTimeout(loadUrlWithRetry, 500);
      });
    };
    loadUrlWithRetry();
  } else {
    const filePath = path.join(__dirname, '../dist/index.html');
    console.log('[Main Process] Loading file:', filePath);
    mainWindow.loadFile(filePath);
  }

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// IPC handlers for saving PDF and files
ipcMain.handle('save-pdf', async (event, { defaultName }) => {
  if (!mainWindow) return { success: false, error: 'Window not found' };

  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'ذخیره پیش‌فاکتور PDF',
    defaultPath: defaultName || 'invoice.pdf',
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });

  if (canceled || !filePath) {
    return { success: false, canceled: true };
  }

  try {
    const pdfData = await mainWindow.webContents.printToPDF({
      pageSize: 'A4',
      landscape: true,
      printBackground: true,
      margins: { marginType: 'custom', top: 0.3, bottom: 0.3, left: 0.3, right: 0.3 },
    });

    fs.writeFileSync(filePath, pdfData);
    return { success: true, filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('export-file', async (event, { defaultName, content, extensions = ['json'] }) => {
  if (!mainWindow) return { success: false, error: 'Window not found' };

  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'ذخیره فایل',
    defaultPath: defaultName,
    filters: [{ name: 'Data File', extensions }],
  });

  if (canceled || !filePath) return { success: false, canceled: true };

  try {
    fs.writeFileSync(filePath, content, 'utf-8');
    return { success: true, filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('import-file', async (event, { extensions = ['json', 'csv'] }) => {
  if (!mainWindow) return { success: false, error: 'Window not found' };

  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: 'انتخاب فایل جهت بازیابی',
    properties: ['openFile'],
    filters: [{ name: 'Data Files', extensions }],
  });

  if (canceled || filePaths.length === 0) return { success: false, canceled: true };

  try {
    const content = fs.readFileSync(filePaths[0], 'utf-8');
    return { success: true, content, filePath: filePaths[0] };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
