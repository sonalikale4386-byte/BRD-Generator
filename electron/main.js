'use strict';

const { app, BrowserWindow, shell, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const http = require('http');
const net  = require('net');
const fs   = require('fs');

let mainWindow = null;
let serverPort = 3978;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function findFreePort(start) {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.unref();
    srv.on('error', () => findFreePort(start + 1).then(resolve, reject));
    srv.listen(start, () => {
      const { port } = srv.address();
      srv.close(() => resolve(port));
    });
  });
}

function waitForServer(port, retries = 40) {
  return new Promise((resolve, reject) => {
    const attempt = (n) => {
      if (n <= 0) return reject(new Error('Server did not start in time'));
      setTimeout(() => {
        http.get(`http://localhost:${port}/health`, (res) => {
          if (res.statusCode === 200) resolve();
          else attempt(n - 1);
        }).on('error', () => attempt(n - 1));
      }, 500);
    };
    attempt(retries);
  });
}

// ─── App root (real filesystem path, works with asar:false) ──────────────────

function getAppRoot() {
  return app.isPackaged
    ? path.join(process.resourcesPath, 'app')
    : path.join(__dirname, '..');
}

// ─── Configure writable paths before starting the server ─────────────────────

function configurePaths() {
  const outputDir = path.join(
    app.getPath('documents'),
    'Synoptek BRD Generator',
    'output'
  );
  fs.mkdirSync(outputDir, { recursive: true });
  process.env.OUTPUT_DIR = outputDir;

  const userEnvPath = path.join(app.getPath('userData'), '.env');
  const bundledEnv  = path.join(getAppRoot(), '.env');
  if (fs.existsSync(bundledEnv)) {
    fs.copyFileSync(bundledEnv, userEnvPath);
  } else if (!fs.existsSync(userEnvPath)) {
    fs.writeFileSync(userEnvPath,
      '# Synoptek BRD Generator — Configuration\n\n' +
      'CLAUDE_API_KEY=\n' +
      'AZURE_TENANT_ID=\n' +
      'AZURE_CLIENT_ID=\n' +
      'AZURE_CLIENT_SECRET=\n'
    );
  }

  process.env.APP_ENV_PATH = userEnvPath;
}

// ─── Start the Express server in-process ─────────────────────────────────────

function startServer(port) {
  process.env.PORT = String(port);
  const serverPath = path.join(getAppRoot(), 'bot', 'server.js');
  require(serverPath);
}

// ─── Auto-updater ─────────────────────────────────────────────────────────────

function setupAutoUpdater() {
  if (!app.isPackaged) return;

  autoUpdater.autoDownload    = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('update-available', () => {
    dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Available',
      message: 'A new version of Synoptek BRD Generator is available.\nIt will download in the background and install when you close the app.',
      buttons: ['OK'],
    });
  });

  autoUpdater.on('update-downloaded', () => {
    dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Ready',
      message: 'Update downloaded. The app will restart to apply the update.',
      buttons: ['Restart Now', 'Later'],
    }).then(({ response }) => {
      if (response === 0) autoUpdater.quitAndInstall();
    });
  });

  autoUpdater.on('error', (err) => {
    console.error('Auto-updater error:', err.message);
  });

  autoUpdater.checkForUpdatesAndNotify();
}

// ─── BrowserWindow ───────────────────────────────────────────────────────────

function createWindow(port) {
  const iconPath = path.join(__dirname, 'icon.png');

  mainWindow = new BrowserWindow({
    width: 1140,
    height: 820,
    minWidth: 820,
    minHeight: 600,
    title: 'Synoptek BRD Generator',
    ...(fs.existsSync(iconPath) ? { icon: iconPath } : {}),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
    show: false,
    backgroundColor: '#0a1628',
  });

  mainWindow.setMenu(null);
  mainWindow.loadURL(`http://localhost:${port}`);

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    mainWindow.focus();
    setupAutoUpdater();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ─── Boot ─────────────────────────────────────────────────────────────────────

app.whenReady().then(async () => {
  try {
    configurePaths();
    serverPort = await findFreePort(3978);
    startServer(serverPort);
    await waitForServer(serverPort);
    createWindow(serverPort);
  } catch (err) {
    dialog.showErrorBox('Startup Error', `Failed to start BRD Generator:\n${err.message}`);
    app.quit();
  }
});

app.on('window-all-closed', () => app.quit());

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow(serverPort);
});
