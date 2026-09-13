const fs = require('node:fs');
const path = require('node:path');
const { app, BrowserWindow } = require('electron');

const WINDOW_STATE_FILE = path.join(app.getPath('userData'), 'spreadsheet-window-state.json');
const DEFAULT_BOUNDS = { width: 900, height: 700 };

let spreadsheetWindow = null;

function loadWindowState() {
  try {
    const raw = fs.readFileSync(WINDOW_STATE_FILE, 'utf-8');
    return { ...DEFAULT_BOUNDS, ...JSON.parse(raw) };
  } catch {
    return DEFAULT_BOUNDS;
  }
}

function saveWindowState(bounds) {
  try {
    fs.writeFileSync(WINDOW_STATE_FILE, JSON.stringify(bounds));
  } catch {
    // ウィンドウ状態の保存に失敗しても致命的ではないため無視する
  }
}

function sendInitData(win, payload) {
  win.webContents.send('spreadsheet:init', payload);
}

function openSpreadsheetWindow(payload) {
  if (spreadsheetWindow && !spreadsheetWindow.isDestroyed()) {
    spreadsheetWindow.focus();
    sendInitData(spreadsheetWindow, payload);
    return;
  }

  const bounds = loadWindowState();
  spreadsheetWindow = new BrowserWindow({
    ...bounds,
    backgroundColor: '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload-spreadsheet.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  spreadsheetWindow.loadFile(path.join(__dirname, '..', 'renderer', 'spreadsheet', 'index.html'));

  spreadsheetWindow.webContents.once('did-finish-load', () => {
    sendInitData(spreadsheetWindow, payload);
  });

  spreadsheetWindow.on('close', () => {
    if (!spreadsheetWindow) return;
    saveWindowState(spreadsheetWindow.getBounds());
  });

  spreadsheetWindow.on('closed', () => {
    spreadsheetWindow = null;
  });
}

module.exports = { openSpreadsheetWindow };
