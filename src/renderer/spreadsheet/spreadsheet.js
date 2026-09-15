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
const ssmlContextMenu = document.getElementById('ssml-context-menu');
const ssmlRateInput = document.getElementById('ssml-rate-input');
const ssmlRateApplyBtn = document.getElementById('ssml-rate-apply');
const ssmlRateRemoveBtn = document.getElementById('ssml-rate-remove');

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

// ---- SSML (<prosody rate="N%">...</prosody>) helpers ----
// Storage format (row.script / .txt files) keeps raw, human-readable text with
// literal <prosody> tags inserted around rate-adjusted spans. XML-escaping for
// actual synthesis happens only at send time (see main/polly.js).

const PROSODY_PATTERN = /<prosody rate="(\d+)%">([\s\S]*?)<\/prosody>/g;

function escapeHtml(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function scriptToEditorHtml(script) {
  let html = '';
  let lastIndex = 0;
  let match;
  PROSODY_PATTERN.lastIndex = 0;
  // eslint-disable-next-line no-cond-assign
  while ((match = PROSODY_PATTERN.exec(script)) !== null) {
    html += escapeHtml(script.slice(lastIndex, match.index));
    const rate = match[1];
    html += `<span class="ssml-rate" data-rate="${rate}" title="再生速度 ${rate}%">${escapeHtml(match[2])}</span>`;
    lastIndex = PROSODY_PATTERN.lastIndex;
  }
  html += escapeHtml(script.slice(lastIndex));
  return html;
}

function serializeEditorNode(node) {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent;
  if (node.nodeType !== Node.ELEMENT_NODE) return '';
  if (node.tagName === 'BR') return '\n';
  if (node.classList && node.classList.contains('ssml-rate')) {
    const rate = node.dataset.rate;
    const inner = Array.from(node.childNodes).map(serializeEditorNode).join('');
    return `<prosody rate="${rate}%">${inner}</prosody>`;
  }
  const inner = Array.from(node.childNodes).map(serializeEditorNode).join('');
  return node.tagName === 'DIV' || node.tagName === 'P' ? `${inner}\n` : inner;
}

function serializeEditorContent() {
  const text = Array.from(editTextarea.childNodes).map(serializeEditorNode).join('');
  return text.replace(/\n$/, '');
}

let ssmlMenuTarget = null; // { mode: 'apply', range: Range } | { mode: 'edit', span: HTMLElement }

function hideSsmlContextMenu() {
  ssmlContextMenu.hidden = true;
  ssmlMenuTarget = null;
}

function showSsmlContextMenu(x, y) {
  ssmlContextMenu.hidden = false;
  const menuRect = ssmlContextMenu.getBoundingClientRect();
  const maxX = window.innerWidth - menuRect.width - 8;
  const maxY = window.innerHeight - menuRect.height - 8;
  ssmlContextMenu.style.left = `${Math.max(8, Math.min(x, maxX))}px`;
  ssmlContextMenu.style.top = `${Math.max(8, Math.min(y, maxY))}px`;
}

function applyRateToRange(range, ratePercent) {
  const span = document.createElement('span');
  span.className = 'ssml-rate';
  span.dataset.rate = String(ratePercent);
  span.title = `再生速度 ${ratePercent}%`;
  const fragment = range.extractContents();
  span.appendChild(fragment);
  range.insertNode(span);
  editTextarea.normalize();
}

function removeSpanTag(span) {
  const parent = span.parentNode;
  while (span.firstChild) parent.insertBefore(span.firstChild, span);
  parent.removeChild(span);
  editTextarea.normalize();
}

const WORD_BOUNDARY_PATTERN = /[\s。、！？.,!?\n]/;

function getRangeAtPoint(x, y) {
  let range = null;
  if (document.caretRangeFromPoint) {
    range = document.caretRangeFromPoint(x, y);
  } else if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(x, y);
    if (pos) {
      range = document.createRange();
      range.setStart(pos.offsetNode, pos.offset);
      range.collapse(true);
    }
  }
  if (!range || !editTextarea.contains(range.startContainer)) return null;
  return range;
}

const WORD_EXPAND_LIMIT = 20; // max chars to expand in each direction; Japanese text often has no
// whitespace/punctuation boundaries at all, so without a cap this could swallow the entire script.

function expandRangeToWord(range) {
  const container = range.startContainer;
  if (container.nodeType !== Node.TEXT_NODE) return null;
  const text = container.textContent;
  let start = range.startOffset;
  let end = range.startOffset;
  let steps = 0;
  while (start > 0 && !WORD_BOUNDARY_PATTERN.test(text[start - 1]) && steps < WORD_EXPAND_LIMIT) {
    start -= 1;
    steps += 1;
  }
  steps = 0;
  while (end < text.length && !WORD_BOUNDARY_PATTERN.test(text[end]) && steps < WORD_EXPAND_LIMIT) {
    end += 1;
    steps += 1;
  }
  if (start === end) return null;
  const wordRange = document.createRange();
  wordRange.setStart(container, start);
  wordRange.setEnd(container, end);
  return wordRange;
}

editTextarea.addEventListener('contextmenu', (event) => {
  event.preventDefault();

  const spanTarget = event.target.closest ? event.target.closest('.ssml-rate') : null;
  if (spanTarget) {
    ssmlMenuTarget = { mode: 'edit', span: spanTarget };
    ssmlRateInput.value = spanTarget.dataset.rate;
    ssmlRateRemoveBtn.hidden = false;
    showSsmlContextMenu(event.clientX, event.clientY);
    return;
  }

  const selection = window.getSelection();
  let range = null;
  if (selection && !selection.isCollapsed && selection.rangeCount > 0) {
    const selRange = selection.getRangeAt(0);
    if (editTextarea.contains(selRange.commonAncestorContainer)) {
      range = selRange.cloneRange();
    }
  }
  if (!range) {
    const pointRange = getRangeAtPoint(event.clientX, event.clientY);
    range = pointRange ? expandRangeToWord(pointRange) : null;
  }

  if (!range) return;
  ssmlMenuTarget = { mode: 'apply', range };
  ssmlRateInput.value = '';
  ssmlRateRemoveBtn.hidden = true;
  showSsmlContextMenu(event.clientX, event.clientY);
});

ssmlContextMenu.querySelectorAll('.ssml-menu-presets button').forEach((btn) => {
  btn.addEventListener('click', () => {
    applySsmlMenuRate(Number(btn.dataset.rate));
  });
});

ssmlRateApplyBtn.addEventListener('click', () => {
  const rate = Number(ssmlRateInput.value);
  if (!Number.isFinite(rate) || rate <= 0) return;
  applySsmlMenuRate(rate);
});

ssmlRateRemoveBtn.addEventListener('click', () => {
  if (ssmlMenuTarget && ssmlMenuTarget.mode === 'edit') {
    removeSpanTag(ssmlMenuTarget.span);
  }
  hideSsmlContextMenu();
});

function applySsmlMenuRate(rate) {
  if (!ssmlMenuTarget) return;
  if (ssmlMenuTarget.mode === 'apply') {
    applyRateToRange(ssmlMenuTarget.range, rate);
  } else if (ssmlMenuTarget.mode === 'edit') {
    ssmlMenuTarget.span.dataset.rate = String(rate);
    ssmlMenuTarget.span.title = `再生速度 ${rate}%`;
  }
  hideSsmlContextMenu();
}

document.addEventListener('mousedown', (event) => {
  if (!ssmlContextMenu.hidden && !ssmlContextMenu.contains(event.target)) {
    hideSsmlContextMenu();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && !ssmlContextMenu.hidden) {
    hideSsmlContextMenu();
  }
});

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
    scriptCell.innerHTML = scriptToEditorHtml(row.script);
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

    const saveRowBtn = document.createElement('button');
    saveRowBtn.textContent = '保存';
    saveRowBtn.addEventListener('click', () => saveOneRow(index));
    actionsCell.appendChild(saveRowBtn);

    tr.appendChild(actionsCell);

    tableBody.appendChild(tr);
  });
}

function openEditDialog(index) {
  state.editingIndex = index;
  editTextarea.innerHTML = scriptToEditorHtml(state.rows[index].script);
  editDialog.showModal();
  editTextarea.focus();
  try {
    document.execCommand('defaultParagraphSeparator', false, 'br');
  } catch {
    // ignore if unsupported
  }
}

editOkBtn.addEventListener('click', () => {
  if (state.editingIndex !== null) {
    state.rows[state.editingIndex].script = serializeEditorContent();
    renderRows();
  }
});

editCancelBtn.addEventListener('click', () => {
  hideSsmlContextMenu();
  editDialog.close();
});

async function saveOneRow(index) {
  const row = state.rows[index];
  setStatus(`ページ${row.pageNumber}を保存中...`);
  try {
    await window.spreadsheetAPI.save({
      workingFolder: state.workingFolder,
      prefix: state.prefix,
      rows: [row],
    });
    setStatus(`ページ${row.pageNumber}を保存しました`);
  } catch (error) {
    setStatus(`ページ${row.pageNumber}の保存に失敗しました: ${error.message || error}`);
  }
}

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
