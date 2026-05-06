'use strict';

const { app, BrowserWindow, shell, dialog } = require('electron'); // dialog kept for startup errors
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

// ─── App root — works with both asar:true and asar:false ─────────────────────

function getAppRoot() {
  if (!app.isPackaged) return path.join(__dirname, '..');
  // asar:true  → files live inside app.asar (Electron patches fs/require for it)
  // asar:false → files live in resources/app directory
  const asarPath  = path.join(process.resourcesPath, 'app.asar');
  const plainPath = path.join(process.resourcesPath, 'app');
  return fs.existsSync(plainPath) ? plainPath : asarPath;
}

// ─── Configure writable paths before starting the server ─────────────────────

function configurePaths() {
  // Keep folder name as 'Synoptek BRD Generator' to preserve existing session history.
  // Migrate from old path if a rename already created a separate folder.
  const oldOutputDir = path.join(app.getPath('documents'), 'Synoptek BRD Generator',    'output');
  const newOutputDir = path.join(app.getPath('documents'), 'Synoptek CE BRD Generator', 'output');

  let outputDir = oldOutputDir;
  // If only the new folder exists (fresh install after rename), migrate back to old name
  if (!fs.existsSync(oldOutputDir) && fs.existsSync(newOutputDir)) {
    try { fs.renameSync(path.dirname(newOutputDir), path.dirname(oldOutputDir)); } catch {}
  }
  fs.mkdirSync(outputDir, { recursive: true });
  process.env.OUTPUT_DIR = outputDir;

  const userEnvPath = path.join(app.getPath('userData'), '.env');
  const bundledEnv  = path.join(getAppRoot(), '.env');

  // Always copy bundled .env → userData so API key updates propagate on every app start
  // Use readFileSync+writeFileSync (copyFileSync doesn't work from inside asar)
  if (fs.existsSync(bundledEnv)) {
    fs.writeFileSync(userEnvPath, fs.readFileSync(bundledEnv));
  } else if (!fs.existsSync(userEnvPath)) {
    fs.writeFileSync(userEnvPath,
      '# Synoptek CE BRD Generator — Configuration\n\n' +
      '# Claude API (https://console.anthropic.com)\n' +
      'CLAUDE_API_KEY=\n\n' +
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

  let autoUpdater;
  try {
    autoUpdater = require('electron-updater').autoUpdater;
  } catch (e) {
    console.warn('electron-updater not available:', e.message);
    return;
  }

  autoUpdater.autoDownload         = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger               = null;

  autoUpdater.on('update-available', (info) => {
    console.log(`🔄 Update available (v${info.version}) — downloading silently`);
    // Show a non-blocking banner in the web UI
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.executeJavaScript(`
        (function() {
          if (document.getElementById('_upd_banner')) return;
          const b = document.createElement('div');
          b.id = '_upd_banner';
          b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;background:#1d4ed8;color:#fff;text-align:center;padding:8px 16px;font-size:13px;font-family:sans-serif';
          b.textContent = 'Downloading update v${info.version}…';
          document.body.appendChild(b);
        })()
      `).catch(() => {});
    }
  });

  autoUpdater.on('update-not-available', () => console.log('✅ App is up to date'));

  // Update fully downloaded — restart automatically, no user click required
  autoUpdater.on('update-downloaded', (info) => {
    console.log(`✅ Update v${info.version} downloaded — restarting automatically in 5 s`);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.executeJavaScript(`
        (function() {
          const b = document.getElementById('_upd_banner') || document.createElement('div');
          b.id = '_upd_banner';
          b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;background:#16a34a;color:#fff;text-align:center;padding:8px 16px;font-size:13px;font-family:sans-serif';
          b.textContent = 'Update v${info.version} ready — restarting in 5 seconds…';
          if (!b.parentNode) document.body.appendChild(b);
          let s = 4;
          const t = setInterval(() => {
            b.textContent = 'Update v${info.version} ready — restarting in ' + s-- + ' second' + (s === 0 ? '' : 's') + '…';
            if (s < 0) clearInterval(t);
          }, 1000);
        })()
      `).catch(() => {});
    }
    // true = silent NSIS install (no install UI), true = relaunch after install
    setTimeout(() => autoUpdater.quitAndInstall(true, true), 5000);
  });

  autoUpdater.on('error', (err) => {
    console.error('Auto-updater error:', err.message);
  });

  autoUpdater.checkForUpdates();
}

// ─── BrowserWindow ───────────────────────────────────────────────────────────

function createWindow(port) {
  const iconPath = path.join(__dirname, 'icon.png');

  mainWindow = new BrowserWindow({
    width: 1140,
    height: 820,
    minWidth: 820,
    minHeight: 600,
    title: 'Synoptek CE BRD Generator',
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
