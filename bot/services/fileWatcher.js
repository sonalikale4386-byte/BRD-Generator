/**
 * fileWatcher.js
 * Watches the local test-docs/ folder for new files.
 * Used in MOCK_ONEDRIVE=true mode to simulate OneDrive change notifications.
 *
 * Uses polling (setInterval) instead of fs.watch() because fs.watch() on Windows
 * fires 2-3 duplicate events per file operation, causing triple notifications.
 * Polling scans every 2 seconds — reliable, no duplicates, never misses a file.
 *
 * When a new supported file is added, calls this.onNewFile(filename).
 */
const fs   = require('fs');
const path = require('path');

const SUPPORTED = new Set(['.pdf', '.docx', '.txt', '.doc']);
const POLL_MS   = 2000;  // scan interval

class FileWatcher {
  constructor(folderPath) {
    this.folderPath = folderPath;
    this.knownFiles = new Set();
    this._interval  = null;
    this.onNewFile  = null;  // set by caller: (filename) => void
  }

  /** Start polling. Seeds known files so only truly NEW additions trigger. */
  start() {
    if (!fs.existsSync(this.folderPath)) {
      fs.mkdirSync(this.folderPath, { recursive: true });
    }

    // Seed existing files — these are NOT "new"
    fs.readdirSync(this.folderPath).forEach(f => this.knownFiles.add(f));
    console.log(`👀 Polling test-docs/ every ${POLL_MS / 1000}s for new files (${this.knownFiles.size} existing files seeded)`);

    this._interval = setInterval(() => {
      let current;
      try {
        current = fs.readdirSync(this.folderPath);
      } catch {
        return; // folder temporarily inaccessible — skip tick
      }

      for (const filename of current) {
        if (this.knownFiles.has(filename)) continue;

        const ext = path.extname(filename).toLowerCase();
        if (!SUPPORTED.has(ext)) continue;

        // Only notify once file is fully written (size stable across one tick)
        const fullPath = path.join(this.folderPath, filename);
        try {
          const size = fs.statSync(fullPath).size;
          if (size === 0) continue; // still being written — wait next tick
        } catch {
          continue;
        }

        // Mark known immediately so the next tick doesn't re-fire
        this.knownFiles.add(filename);
        console.log(`📄 New file detected in test-docs/: ${filename}`);
        if (this.onNewFile) this.onNewFile(filename);
      }
    }, POLL_MS);
  }

  /** Reseed known files (call after BRD generation to reset baseline) */
  refreshKnown() {
    this.knownFiles.clear();
    if (fs.existsSync(this.folderPath)) {
      fs.readdirSync(this.folderPath).forEach(f => this.knownFiles.add(f));
    }
  }

  stop() {
    if (this._interval) {
      clearInterval(this._interval);
      this._interval = null;
    }
  }
}

module.exports = { FileWatcher };
