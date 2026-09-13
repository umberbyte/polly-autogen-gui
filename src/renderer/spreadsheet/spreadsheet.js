const LANGUAGE_NAMES = {
  'ar-AE': 'アラビア語(UAE)',
  arb: 'アラビア語(標準)',
  'ca-ES': 'カタルーニャ語',
  'cmn-CN': '中国語(北京語)',
  'cs-CZ': 'チェコ語',
  'cy-GB': 'ウェールズ語',
  'da-DK': 'デンマーク語',
  'de-AT': 'ドイツ語(オーストリア)',
  'de-CH': 'ドイツ語(スイス)',
  'de-DE': 'ドイツ語',
  'en-AU': '英語(オーストラリア)',
  'en-GB': '英語(英国)',
  'en-GB-WLS': '英語(ウェールズ)',
  'en-IE': '英語(アイルランド)',
  'en-IN': '英語(インド)',
  'en-NZ': '英語(ニュージーランド)',
  'en-SG': '英語(シンガポール)',
  'en-US': '英語(米国)',
  'en-ZA': '英語(南アフリカ)',
  'es-ES': 'スペイン語',
  'es-MX': 'スペイン語(メキシコ)',
  'es-US': 'スペイン語(米国)',
  'fi-FI': 'フィンランド語',
  'fr-BE': 'フランス語(ベルギー)',
  'fr-CA': 'フランス語(カナダ)',
  'fr-FR': 'フランス語',
  'is-IS': 'アイスランド語',
  'it-IT': 'イタリア語',
  'ja-JP': '日本語',
  'ko-KR': '韓国語',
  'nb-NO': 'ノルウェー語',
  'nl-BE': 'オランダ語(ベルギー)',
  'nl-NL': 'オランダ語',
  'pl-PL': 'ポーランド語',
  'pt-BR': 'ポルトガル語(ブラジル)',
  'pt-PT': 'ポルトガル語',
  'ro-RO': 'ルーマニア語',
  'ru-RU': 'ロシア語',
  'sv-SE': 'スウェーデン語',
  'tr-TR': 'トルコ語',
  'yue-CN': '広東語',
};

const state = {
  workingFolder: null,
  prefix: null,
  rows: [],
  editingIndex: null,
  voiceSettings: null,
};

const voiceSettingsLabel = document.getElementById('voice-settings-label');
const workingFolderLabel = document.getElementById('working-folder-label');
const statusEl = document.getElementById('status');
const tableBody = document.getElementById('script-table-body');
const saveBtn = document.getElementById('save-btn');
const batchGenerateBtn = document.getElementById('batch-generate-btn');
const batchCancelBtn = document.getElementById('batch-cancel-btn');
const batchProgressEl = document.getElementById('batch-progress');
const editDialog = document.getElementById('edit-dialog');
const editTextarea = document.getElementById('edit-textarea');
const editCancelBtn = document.getElementById('edit-cancel-btn');
const editOkBtn = document.getElementById('edit-ok-btn');
const audioPlayer = document.getElementById('audio-player');

function setStatus(message) {
  statusEl.textContent = message;
}

function renderVoiceSettingsLabel() {
  const settings = state.voiceSettings;
  if (!settings) {
    voiceSettingsLabel.textContent = '';
    return;
  }
  const languageName = LANGUAGE_NAMES[settings.languageCode] || settings.languageCode;
  voiceSettingsLabel.textContent =
    `現在の Polly 出力設定:${languageName}(${settings.languageCode})、` +
    `${settings.engine}、${settings.voiceId}`;
}

function toFileUrl(filePath) {
  let normalized = filePath.replace(/\\/g, '/');
  if (!normalized.startsWith('/')) normalized = `/${normalized}`;
  return `file://${encodeURI(normalized)}`;
}

function mp3PathFor(row) {
  return `${state.workingFolder}/mp3/${state.prefix}_${row.pageNumber}.mp3`;
}

function renderRows() {
  tableBody.innerHTML = '';
  state.rows.forEach((row, index) => {
    const tr = document.createElement('tr');

    const pageCell = document.createElement('td');
    pageCell.className = 'page-number-col';
    pageCell.textContent = row.pageNumber;
    tr.appendChild(pageCell);

    const scriptCell = document.createElement('td');
    scriptCell.className = 'script-cell';
    scriptCell.textContent = row.script;
    scriptCell.addEventListener('click', () => openEditDialog(index));
    tr.appendChild(scriptCell);

    const actionsCell = document.createElement('td');
    actionsCell.className = 'actions-col';

    const playBtn = document.createElement('button');
    playBtn.textContent = '再生';
    playBtn.disabled = !row.hasMp3;
    playBtn.addEventListener('click', () => {
      audioPlayer.src = toFileUrl(mp3PathFor(row));
      audioPlayer.play();
    });
    actionsCell.appendChild(playBtn);

    const generateBtn = document.createElement('button');
    generateBtn.textContent = '生成';
    generateBtn.addEventListener('click', () => generateOneRow(index));
    actionsCell.appendChild(generateBtn);

    tr.appendChild(actionsCell);

    tableBody.appendChild(tr);
  });
}

function openEditDialog(index) {
  state.editingIndex = index;
  editTextarea.value = state.rows[index].script;
  editDialog.showModal();
  editTextarea.focus();
}

editOkBtn.addEventListener('click', () => {
  if (state.editingIndex !== null) {
    state.rows[state.editingIndex].script = editTextarea.value;
    renderRows();
  }
});

editCancelBtn.addEventListener('click', () => {
  editDialog.close();
});

saveBtn.addEventListener('click', async () => {
  setStatus('保存中...');
  try {
    const result = await window.spreadsheetAPI.save({
      workingFolder: state.workingFolder,
      prefix: state.prefix,
      rows: state.rows,
    });
    setStatus(`保存しました(${result.count}件)`);
  } catch (error) {
    setStatus(`保存に失敗しました: ${error.message || error}`);
  }
});

async function generateOneRow(index) {
  const row = state.rows[index];
  setStatus(`ページ${row.pageNumber}を音声化中...`);
  try {
    const result = await window.spreadsheetAPI.generateOne({
      workingFolder: state.workingFolder,
      prefix: state.prefix,
      rows: state.rows,
      pageNumber: row.pageNumber,
      voiceSettings: state.voiceSettings,
    });
    if (result.ok) {
      row.hasMp3 = true;
      renderRows();
      setStatus(`ページ${row.pageNumber}の音声を生成しました`);
    } else {
      setStatus(`ページ${row.pageNumber}の音声生成に失敗しました: ${result.error}`);
    }
  } catch (error) {
    setStatus(`ページ${row.pageNumber}の音声生成に失敗しました: ${error.message || error}`);
  }
}

batchGenerateBtn.addEventListener('click', async () => {
  batchGenerateBtn.disabled = true;
  batchCancelBtn.hidden = false;
  batchCancelBtn.disabled = false;
  batchProgressEl.textContent = '';
  setStatus('一括音声出力を開始します...');
  try {
    const result = await window.spreadsheetAPI.batchGenerate({
      workingFolder: state.workingFolder,
      prefix: state.prefix,
      rows: state.rows,
      voiceSettings: state.voiceSettings,
    });
    if (!result.ok && result.canceled) {
      setStatus('一括音声出力を中止しました');
    } else if (!result.ok) {
      setStatus(`一括音声出力に失敗しました: ${result.error}`);
    }
  } catch (error) {
    setStatus(`一括音声出力に失敗しました: ${error.message || error}`);
  }
  batchGenerateBtn.disabled = false;
  batchCancelBtn.hidden = true;
});

batchCancelBtn.addEventListener('click', () => {
  batchCancelBtn.disabled = true;
  setStatus('中止要求を送信しました。処理中のリクエストの完了を待っています...');
  window.spreadsheetAPI.cancelBatchGenerate();
});

function formatBatchProgress(progress) {
  const completed = progress.completed ?? 0;
  const active = progress.active ?? 0;
  const max = progress.maxConcurrency ?? 0;
  return `[${completed}/${progress.total}] (スレッド ${active}/${max})`;
}

window.spreadsheetAPI.onBatchProgress((progress) => {
  switch (progress.type) {
    case 'start':
      batchProgressEl.textContent = formatBatchProgress(progress);
      setStatus(`一括音声出力中...(最大${progress.maxConcurrency}件同時処理)`);
      break;
    case 'row-start':
      batchProgressEl.textContent = formatBatchProgress(progress);
      setStatus(`ページ${progress.pageNumber}を音声化中...`);
      break;
    case 'row-done': {
      const row = state.rows.find((r) => r.pageNumber === progress.pageNumber);
      if (row) row.hasMp3 = true;
      renderRows();
      batchProgressEl.textContent = formatBatchProgress(progress);
      setStatus(`ページ${progress.pageNumber}完了(${progress.completed}/${progress.total})`);
      break;
    }
    case 'row-error':
      batchProgressEl.textContent = formatBatchProgress(progress);
      setStatus(`ページ${progress.pageNumber}でエラー: ${progress.error}`);
      break;
    case 'done':
      batchProgressEl.textContent = formatBatchProgress(progress);
      setStatus(
        progress.stopped
          ? `一括音声出力を中止しました(${progress.completed}/${progress.total}件処理、エラー${progress.errorCount}件)`
          : progress.errorCount > 0
            ? `一括音声出力が完了しました(${progress.total}件中${progress.errorCount}件エラー)`
            : `一括音声出力が完了しました(${progress.total}件)`,
      );
      break;
    default:
      break;
  }
});

window.spreadsheetAPI.onInit((data) => {
  state.workingFolder = data.workingFolder;
  state.prefix = data.prefix;
  state.rows = data.rows;
  state.voiceSettings = data.voiceSettings;
  workingFolderLabel.textContent = state.workingFolder;
  renderVoiceSettingsLabel();
  renderRows();

  if (data.warnings && data.warnings.length > 0) {
    setStatus(`警告: ${data.warnings.join(' / ')}`);
  } else {
    setStatus(`${state.rows.length}件を読み込みました`);
  }
});
