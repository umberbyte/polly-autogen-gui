const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
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
const {
  readDefaultCredentials,
  saveDefaultCredentials,
  ensureCredentialsFileExists,
} = require('./awsCredentials');

const AWS_REGION = 'ap-northeast-1';
const BATCH_CONCURRENCY = 10;

const activeBatchCancels = new Map();

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
    height: 1200,
    minWidth: 720,
    minHeight: 650,
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

ipcMain.handle('polly:list-voices', async () => {
  const pollyClient = new PollyClient({ region: AWS_REGION });

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

ipcMain.handle('settings:read-aws-credentials', () => readDefaultCredentials());

ipcMain.handle('settings:save-aws-credentials', async (event, credentials) => {
  await saveDefaultCredentials(credentials);
  return { ok: true };
});

ipcMain.handle('settings:open-aws-credentials-file', async () => {
  const filePath = await ensureCredentialsFileExists();
  const error = await shell.openPath(filePath);
  return { ok: !error, error: error || null };
});

ipcMain.handle('settings:load-voice-settings', () => loadVoiceSettings());

ipcMain.handle('settings:save-voice-settings', (event, settings) => {
  saveVoiceSettings(settings);
  return { ok: true };
});

ipcMain.handle('polly:convert', async (event, payload) => {
  const pollyClient = new PollyClient({
    region: AWS_REGION,
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

  if (proposedFolder !== baseFolder) {
    await dialog.showMessageBox(win, {
      type: 'info',
      message: `既に同じファイルで作業フォルダが生成されているため、${path.basename(proposedFolder)}フォルダを作成しました。`,
      buttons: ['OK'],
    });
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
      const pollyClient = new PollyClient({ region: AWS_REGION });
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

    const pollyClient = new PollyClient({ region: AWS_REGION, maxAttempts: 5 });
    const total = rows.length;
    const workerCount = Math.min(BATCH_CONCURRENCY, rows.length);
    event.sender.send('spreadsheet:batch-progress', {
      type: 'start',
      total,
      active: 0,
      maxConcurrency: workerCount,
    });

    let stopped = false;
    let completed = 0;
    let errorCount = 0;
    let active = 0;
    let nextIndex = 0;
    let dialogChain = Promise.resolve();

    activeBatchCancels.set(event.sender.id, () => {
      stopped = true;
    });

    async function worker() {
      for (;;) {
        if (stopped || nextIndex >= rows.length) return;
        const row = rows[nextIndex];
        nextIndex += 1;

        active += 1;
        event.sender.send('spreadsheet:batch-progress', {
          type: 'row-start',
          completed,
          total,
          pageNumber: row.pageNumber,
          active,
          maxConcurrency: workerCount,
        });

        try {
          // eslint-disable-next-line no-await-in-loop
          await synthesizeRowToMp3(pollyClient, workingFolder, prefix, row, voiceSettings);
          active -= 1;
          completed += 1;
          event.sender.send('spreadsheet:batch-progress', {
            type: 'row-done',
            completed,
            total,
            pageNumber: row.pageNumber,
            active,
            maxConcurrency: workerCount,
          });
        } catch (error) {
          active -= 1;
          completed += 1;
          errorCount += 1;
          const message = error.message || String(error);
          event.sender.send('spreadsheet:batch-progress', {
            type: 'row-error',
            completed,
            total,
            pageNumber: row.pageNumber,
            error: message,
            active,
            maxConcurrency: workerCount,
          });

          dialogChain = dialogChain.then(async () => {
            if (stopped) return;
            const choice = await dialog.showMessageBox(win, {
              type: 'error',
              message: `ページ${row.pageNumber}の音声生成でエラーが発生しました`,
              detail: message,
              buttons: ['続行', '中止'],
              defaultId: 0,
              cancelId: 1,
            });
            if (choice.response !== 0) stopped = true;
          });
        }
      }
    }

    try {
      await Promise.all(Array.from({ length: workerCount }, () => worker()));
      await dialogChain;
    } finally {
      activeBatchCancels.delete(event.sender.id);
    }

    event.sender.send('spreadsheet:batch-progress', {
      type: 'done',
      total,
      completed,
      errorCount,
      stopped,
      active: 0,
      maxConcurrency: workerCount,
    });
    return { ok: true, stopped };
  },
);

ipcMain.on('spreadsheet:batch-cancel', (event) => {
  const cancel = activeBatchCancels.get(event.sender.id);
  if (cancel) cancel();
});

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
    return {
      ok: true,
      workingPath,
      fileName: path.basename(workingPath),
      originalFileName: path.basename(pptxPath),
    };
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }
});

ipcMain.handle('shell:show-item-in-folder', (event, filePath) => {
  shell.showItemInFolder(filePath);
  return { ok: true };
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
