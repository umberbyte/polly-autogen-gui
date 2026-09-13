const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { PollyClient, DescribeVoicesCommand } = require('@aws-sdk/client-polly');
const { runConversion } = require('./polly');
const {
  parseTalkScriptSheet,
  resolveUniqueFolder,
  loadRowsFromTextFolder,
  saveRowsToTextFolder,
} = require('./excelImport');
const { openSpreadsheetWindow } = require('./spreadsheetWindow');
const { loadVoiceSettings, saveVoiceSettings } = require('./voiceSettings');
const { checkMp3Status, anyMp3Exists, synthesizeRowToMp3 } = require('./spreadsheetAudio');
const { createWorkingCopy, embedAudioIntoPptx } = require('./pptxEmbed');

const DEFAULT_VOICE_SETTINGS = {
  languageCode: 'ja-JP',
  engine: 'neural',
  voiceId: 'Tomoko',
  outputFormat: 'mp3',
};

function getEffectiveVoiceSettings() {
  return loadVoiceSettings() || DEFAULT_VOICE_SETTINGS;
}

async function attachMp3Status(workingFolder, prefix, rows) {
  const status = await checkMp3Status(workingFolder, prefix, rows);
  return rows.map((row) => ({ ...row, hasMp3: Boolean(status[row.pageNumber]) }));
}

function createWindow() {
  const win = new BrowserWindow({
    width: 880,
    height: 1000,
    minWidth: 720,
    minHeight: 560,
    backgroundColor: '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
  return win;
}

ipcMain.handle('dialog:select-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const result = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [{ name: 'テキストファイル', extensions: ['txt'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('dialog:select-directory', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const result = await dialog.showOpenDialog(win, {
    properties: ['openDirectory'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('dialog:select-excel-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const result = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [{ name: 'Excelファイル', extensions: ['xlsx', 'xls'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('dialog:select-working-folder', async (event, { defaultPath } = {}) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const result = await dialog.showOpenDialog(win, {
    properties: ['openDirectory'],
    defaultPath,
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('polly:list-voices', async (event, { region } = {}) => {
  const pollyClient = new PollyClient({ region: region || undefined });

  const voices = [];
  let nextToken;
  do {
    // eslint-disable-next-line no-await-in-loop
    const response = await pollyClient.send(new DescribeVoicesCommand({ NextToken: nextToken }));
    voices.push(...(response.Voices || []));
    nextToken = response.NextToken;
  } while (nextToken);

  return voices.map((voice) => ({
    id: voice.Id,
    name: voice.Name,
    gender: voice.Gender,
    languageCode: voice.LanguageCode,
    languageName: voice.LanguageName,
    engines: voice.SupportedEngines || [],
  }));
});

ipcMain.handle('settings:load-voice-settings', () => loadVoiceSettings());

ipcMain.handle('settings:save-voice-settings', (event, settings) => {
  saveVoiceSettings(settings);
  return { ok: true };
});

ipcMain.handle('polly:convert', async (event, payload) => {
  const { options } = payload;
  const pollyClient = new PollyClient({
    region: options.region || undefined,
  });

  const onProgress = (progress) => {
    event.sender.send('polly:progress', progress);
  };

  await runConversion(payload, pollyClient, onProgress);
});

ipcMain.handle('excel:start-import', async (event, { excelPath }) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const parsedPath = path.parse(excelPath);
  const baseFolder = path.join(parsedPath.dir, parsedPath.name);

  let proposedFolder;
  try {
    proposedFolder = await resolveUniqueFolder(baseFolder);
    await fs.mkdir(proposedFolder, { recursive: true });
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }

  await dialog.showMessageBox(win, {
    type: 'info',
    message: '作業フォルダを選択してください',
    buttons: ['OK'],
  });

  const result = await dialog.showOpenDialog(win, {
    properties: ['openDirectory'],
    defaultPath: proposedFolder,
  });
  if (result.canceled || result.filePaths.length === 0) {
    return { ok: false, canceled: true };
  }
  const workingFolder = result.filePaths[0];

  let rows;
  try {
    rows = await parseTalkScriptSheet(excelPath);
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }

  const prefix = path.basename(workingFolder);
  const rowsWithMp3Status = await attachMp3Status(workingFolder, prefix, rows);
  openSpreadsheetWindow({
    workingFolder,
    prefix,
    rows: rowsWithMp3Status,
    warnings: [],
    voiceSettings: getEffectiveVoiceSettings(),
  });
  return { ok: true };
});

ipcMain.handle('workfolder:start-review', async (event, { folderPath }) => {
  const { workingFolder, rows, warnings } = await loadRowsFromTextFolder(folderPath);
  const prefix = path.basename(workingFolder);
  const rowsWithMp3Status = await attachMp3Status(workingFolder, prefix, rows);
  openSpreadsheetWindow({
    workingFolder,
    prefix,
    rows: rowsWithMp3Status,
    warnings,
    voiceSettings: getEffectiveVoiceSettings(),
  });
  return { ok: true, warnings };
});

ipcMain.handle('spreadsheet:save', async (event, { workingFolder, prefix, rows }) => {
  const count = await saveRowsToTextFolder(workingFolder, prefix, rows);
  return { ok: true, count };
});

ipcMain.handle(
  'spreadsheet:generate-one',
  async (event, { workingFolder, prefix, rows, pageNumber, voiceSettings }) => {
    await saveRowsToTextFolder(workingFolder, prefix, rows);

    const row = rows.find((r) => r.pageNumber === pageNumber);
    if (!row) return { ok: false, pageNumber, error: '対象の行が見つかりません' };

    try {
      const pollyClient = new PollyClient({});
      const outPath = await synthesizeRowToMp3(pollyClient, workingFolder, prefix, row, voiceSettings);
      return { ok: true, pageNumber, outPath };
    } catch (error) {
      return { ok: false, pageNumber, error: error.message || String(error) };
    }
  },
);

ipcMain.handle(
  'spreadsheet:batch-generate',
  async (event, { workingFolder, prefix, rows, voiceSettings }) => {
    const win = BrowserWindow.fromWebContents(event.sender);

    await saveRowsToTextFolder(workingFolder, prefix, rows);

    if (await anyMp3Exists(workingFolder, prefix, rows)) {
      const result = await dialog.showMessageBox(win, {
        type: 'warning',
        message: '既存のmp3ファイルを上書きします。よろしいですか?',
        buttons: ['OK', '中止'],
        defaultId: 0,
        cancelId: 1,
      });
      if (result.response !== 0) return { ok: false, canceled: true };
    }

    const pollyClient = new PollyClient({});
    const total = rows.length;
    event.sender.send('spreadsheet:batch-progress', { type: 'start', total });

    let stopped = false;
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i];
      event.sender.send('spreadsheet:batch-progress', {
        type: 'row-start',
        index: i + 1,
        total,
        pageNumber: row.pageNumber,
      });
      try {
        // eslint-disable-next-line no-await-in-loop
        await synthesizeRowToMp3(pollyClient, workingFolder, prefix, row, voiceSettings);
        event.sender.send('spreadsheet:batch-progress', {
          type: 'row-done',
          index: i + 1,
          total,
          pageNumber: row.pageNumber,
        });
      } catch (error) {
        const message = error.message || String(error);
        event.sender.send('spreadsheet:batch-progress', {
          type: 'row-error',
          index: i + 1,
          total,
          pageNumber: row.pageNumber,
          error: message,
        });

        // eslint-disable-next-line no-await-in-loop
        const choice = await dialog.showMessageBox(win, {
          type: 'error',
          message: `ページ${row.pageNumber}の音声生成でエラーが発生しました`,
          detail: message,
          buttons: ['続行', '中止'],
          defaultId: 0,
          cancelId: 1,
        });
        if (choice.response !== 0) {
          stopped = true;
          break;
        }
      }
    }

    event.sender.send('spreadsheet:batch-progress', { type: 'done', total, stopped });
    return { ok: true, stopped };
  },
);

ipcMain.handle('dialog:select-pptx-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const result = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [{ name: 'PowerPointファイル', extensions: ['pptx'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('pptx:prepare-working-copy', async (event, { pptxPath }) => {
  try {
    const workingPath = await createWorkingCopy(pptxPath);
    return { ok: true, workingPath, fileName: path.basename(workingPath) };
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }
});

ipcMain.handle('pptx:embed-audio', async (event, { workingPath, mp3Folder }) => {
  try {
    const result = await embedAudioIntoPptx(workingPath, mp3Folder, (progress) => {
      event.sender.send('pptx:embed-progress', progress);
    });
    return { ok: true, ...result };
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
