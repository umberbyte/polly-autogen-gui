const state = {
  mode: null,
  targetPath: null,
  voices: [],
  savedVoiceSettings: null,
  pptxWorkingPath: null,
  mp3Folder: null,
};

const targetPathInput = document.getElementById('target-path');
const convertBtn = document.getElementById('convert-btn');
const reloadVoicesBtn = document.getElementById('reload-voices-btn');
const saveVoiceSettingsBtn = document.getElementById('save-voice-settings-btn');
const overallProgress = document.getElementById('overall-progress');
const fileList = document.getElementById('file-list');
const logEl = document.getElementById('log');
const languageCodeSelect = document.getElementById('language-code');
const voiceIdSelect = document.getElementById('voice-id');
const engineSelect = document.getElementById('engine');
const outputFormatSelect = document.getElementById('output-format');

function appendLog(message) {
  logEl.textContent += `${message}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function updateConvertBtnState() {
  convertBtn.disabled = !state.targetPath || state.voices.length === 0;
}

function setTarget(mode, targetPath) {
  state.mode = mode;
  state.targetPath = targetPath;
  targetPathInput.value = targetPath || '';
  updateConvertBtnState();
}

document.getElementById('select-file-btn').addEventListener('click', async () => {
  const filePath = await window.pollyAutogen.selectFile();
  if (filePath) setTarget('file', filePath);
});

document.getElementById('select-directory-btn').addEventListener('click', async () => {
  const dirPath = await window.pollyAutogen.selectDirectory();
  if (dirPath) setTarget('directory', dirPath);
});

document.getElementById('select-excel-btn').addEventListener('click', async () => {
  const excelPath = await window.pollyAutogen.selectExcelFile();
  if (!excelPath) return;
  const result = await window.pollyAutogen.startExcelImport(excelPath);
  if (!result.ok && !result.canceled) {
    appendLog(`Excel取り込みに失敗しました: ${result.error}`);
  }
});

document.getElementById('open-workfolder-btn').addEventListener('click', async () => {
  const folderPath = await window.pollyAutogen.selectWorkingFolder();
  if (!folderPath) return;
  const result = await window.pollyAutogen.startWorkFolderReview(folderPath);
  if (!result.ok) {
    appendLog(`作業フォルダの読み込みに失敗しました: ${result.error}`);
  } else if (result.warnings && result.warnings.length > 0) {
    result.warnings.forEach((warning) => appendLog(`警告: ${warning}`));
  }
});

const awsAccountSettingsBtn = document.getElementById('aws-account-settings-btn');
const awsCredentialsEditBtn = document.getElementById('aws-credentials-edit-btn');
const awsCredentialsDialog = document.getElementById('aws-credentials-dialog');
const awsAccessKeyInput = document.getElementById('aws-access-key-input');
const awsSecretKeyInput = document.getElementById('aws-secret-key-input');
const awsCredentialsCancelBtn = document.getElementById('aws-credentials-cancel-btn');
const awsCredentialsSaveBtn = document.getElementById('aws-credentials-save-btn');

awsAccountSettingsBtn.addEventListener('click', async () => {
  const current = await window.pollyAutogen.readAwsCredentials();
  awsAccessKeyInput.value = (current && current.accessKeyId) || '';
  awsSecretKeyInput.value = (current && current.secretAccessKey) || '';
  awsCredentialsDialog.showModal();
});

awsCredentialsCancelBtn.addEventListener('click', () => {
  awsCredentialsDialog.close();
});

awsCredentialsSaveBtn.addEventListener('click', async () => {
  await window.pollyAutogen.saveAwsCredentials({
    accessKeyId: awsAccessKeyInput.value.trim(),
    secretAccessKey: awsSecretKeyInput.value.trim(),
  });
  appendLog('AWSアカウントの設定を保存しました');
});

awsCredentialsEditBtn.addEventListener('click', async () => {
  const result = await window.pollyAutogen.openAwsCredentialsFile();
  if (!result.ok) appendLog(`設定ファイルを開けませんでした: ${result.error}`);
});

const pptxFileLabel = document.getElementById('pptx-file-label');
const mp3FolderLabel = document.getElementById('mp3-folder-label');
const embedAudioBtn = document.getElementById('embed-audio-btn');
const embedProgressEl = document.getElementById('embed-progress');

function updateEmbedBtnState() {
  embedAudioBtn.disabled = !(state.pptxWorkingPath && state.mp3Folder);
}

document.getElementById('select-pptx-btn').addEventListener('click', async () => {
  const pptxPath = await window.pollyAutogen.selectPptxFile();
  if (!pptxPath) return;
  const result = await window.pollyAutogen.preparePptxWorkingCopy(pptxPath);
  if (!result.ok) {
    appendLog(`PowerPointの処理用コピー作成に失敗しました: ${result.error}`);
    return;
  }
  state.pptxWorkingPath = result.workingPath;
  pptxFileLabel.textContent = `元ファイル: ${result.originalFileName}\n出力先ファイル: ${result.fileName}`;
  appendLog(`処理用コピーを作成しました: ${result.workingPath}`);
  updateEmbedBtnState();
});

document.getElementById('select-mp3-folder-btn').addEventListener('click', async () => {
  const folder = await window.pollyAutogen.selectDirectory();
  if (!folder) return;
  state.mp3Folder = folder;
  mp3FolderLabel.textContent = folder;
  updateEmbedBtnState();
});

embedAudioBtn.addEventListener('click', async () => {
  embedAudioBtn.disabled = true;
  embedProgressEl.textContent = '';
  appendLog('PowerPointへの音声埋め込みを開始します...');
  try {
    const result = await window.pollyAutogen.embedAudioToPptx({
      workingPath: state.pptxWorkingPath,
      mp3Folder: state.mp3Folder,
    });
    if (!result.ok) {
      appendLog(`音声埋め込みに失敗しました: ${result.error}`);
    } else {
      appendLog(`音声埋め込みが完了しました(${result.embeddedCount}件): ${state.pptxWorkingPath}`);
      (result.warnings || []).forEach((warning) => appendLog(`警告: ${warning}`));
      window.pollyAutogen.showItemInFolder(state.pptxWorkingPath);
    }
  } catch (error) {
    appendLog(`音声埋め込みに失敗しました: ${error.message || error}`);
  }
  updateEmbedBtnState();
});

window.pollyAutogen.onPptxEmbedProgress((progress) => {
  switch (progress.type) {
    case 'start':
      embedProgressEl.textContent = `[0/${progress.total}]`;
      break;
    case 'slide-start':
    case 'slide-done':
    case 'slide-skip':
    case 'slide-error':
      embedProgressEl.textContent = `[${progress.index}/${progress.total}]`;
      break;
    case 'done':
      embedProgressEl.textContent = '処理完了';
      break;
    default:
      break;
  }
});

function fillSelect(select, options, preferredValue) {
  select.innerHTML = '';
  for (const opt of options) {
    const el = document.createElement('option');
    el.value = opt.value;
    el.textContent = opt.label;
    select.appendChild(el);
  }
  if (options.some((opt) => opt.value === preferredValue)) {
    select.value = preferredValue;
  } else if (options.length > 0) {
    select.value = options[0].value;
  }
}

function updateSettingsMismatchIndicators() {
  const saved = state.savedVoiceSettings;
  const checks = [
    [languageCodeSelect, 'languageCode'],
    [engineSelect, 'engine'],
    [voiceIdSelect, 'voiceId'],
    [outputFormatSelect, 'outputFormat'],
  ];
  for (const [select, key] of checks) {
    select.classList.toggle('mismatch', Boolean(saved) && select.value !== saved[key]);
  }
}

function populateVoiceOptions() {
  const voicesForLanguageAndEngine = state.voices
    .filter(
      (v) => v.languageCode === languageCodeSelect.value && v.engines.includes(engineSelect.value),
    )
    .sort((a, b) => a.id.localeCompare(b.id));
  fillSelect(
    voiceIdSelect,
    voicesForLanguageAndEngine.map((v) => ({ value: v.id, label: `${v.id}(${v.gender})` })),
    'Tomoko',
  );
  updateSettingsMismatchIndicators();
}

function populateEngineOptions() {
  const engineSet = new Set();
  for (const v of state.voices) {
    if (v.languageCode !== languageCodeSelect.value) continue;
    v.engines.forEach((e) => engineSet.add(e));
  }
  const engines = [...engineSet].sort();
  fillSelect(
    engineSelect,
    engines.map((e) => ({ value: e, label: e })),
    'neural',
  );
  populateVoiceOptions();
}

function populateLanguageOptions() {
  const languageMap = new Map();
  for (const v of state.voices) {
    if (!languageMap.has(v.languageCode)) languageMap.set(v.languageCode, v.languageName);
  }
  const languages = [...languageMap.entries()]
    .map(([code, name]) => ({ value: code, label: `${name}(${code})` }))
    .sort((a, b) => a.value.localeCompare(b.value));
  fillSelect(languageCodeSelect, languages, 'ja-JP');
  populateEngineOptions();
}

async function loadVoices() {
  reloadVoicesBtn.disabled = true;
  appendLog('音声一覧を取得中...');
  try {
    state.voices = await window.pollyAutogen.listVoices();
    populateLanguageOptions();
    appendLog(`音声一覧を取得しました(${state.voices.length}件)`);
  } catch (error) {
    state.voices = [];
    appendLog(`音声一覧の取得に失敗しました: ${error.message || error}`);
  }
  updateConvertBtnState();
  reloadVoicesBtn.disabled = false;
}

function applySettingsToForm(settings) {
  if ([...languageCodeSelect.options].some((opt) => opt.value === settings.languageCode)) {
    languageCodeSelect.value = settings.languageCode;
  }
  populateEngineOptions();

  if ([...engineSelect.options].some((opt) => opt.value === settings.engine)) {
    engineSelect.value = settings.engine;
  }
  populateVoiceOptions();

  if ([...voiceIdSelect.options].some((opt) => opt.value === settings.voiceId)) {
    voiceIdSelect.value = settings.voiceId;
  }
  if ([...outputFormatSelect.options].some((opt) => opt.value === settings.outputFormat)) {
    outputFormatSelect.value = settings.outputFormat;
  }
}

async function applySavedOrDefaultVoiceSettings() {
  const saved = await window.pollyAutogen.loadVoiceSettings();
  if (saved) {
    state.savedVoiceSettings = saved;
    applySettingsToForm(saved);
  } else {
    // 設定ファイルが存在しない/空の場合は、現在のデフォルト選択(ja-JP/neural/Tomoko/mp3)を保存する
    const defaults = {
      languageCode: languageCodeSelect.value,
      engine: engineSelect.value,
      voiceId: voiceIdSelect.value,
      outputFormat: outputFormatSelect.value,
    };
    await window.pollyAutogen.saveVoiceSettings(defaults);
    state.savedVoiceSettings = defaults;
  }
  updateSettingsMismatchIndicators();
}

languageCodeSelect.addEventListener('change', populateEngineOptions);
engineSelect.addEventListener('change', populateVoiceOptions);
outputFormatSelect.addEventListener('change', updateSettingsMismatchIndicators);
reloadVoicesBtn.addEventListener('click', loadVoices);

saveVoiceSettingsBtn.addEventListener('click', async () => {
  const settings = {
    languageCode: languageCodeSelect.value,
    engine: engineSelect.value,
    voiceId: voiceIdSelect.value,
    outputFormat: outputFormatSelect.value,
  };
  await window.pollyAutogen.saveVoiceSettings(settings);
  state.savedVoiceSettings = settings;
  updateSettingsMismatchIndicators();
  appendLog('現在の音声設定を保存しました');
});

loadVoices().then(applySavedOrDefaultVoiceSettings);

function getOptions() {
  return {
    voiceId: voiceIdSelect.value,
    engine: engineSelect.value,
    languageCode: languageCodeSelect.value,
    outputFormat: outputFormatSelect.value,
  };
}

const fileStatusEls = new Map();

function resetFileList(total) {
  fileList.innerHTML = '';
  fileStatusEls.clear();
  overallProgress.textContent = total ? `[0/${total}]` : '';
}

function getOrCreateFileRow(fileName) {
  if (fileStatusEls.has(fileName)) return fileStatusEls.get(fileName);
  const li = document.createElement('li');
  li.textContent = fileName;
  fileList.appendChild(li);
  fileStatusEls.set(fileName, li);
  return li;
}

window.pollyAutogen.onProgress((progress) => {
  switch (progress.type) {
    case 'start':
      resetFileList(progress.total);
      appendLog(`変換を開始します(対象: ${progress.total}件)`);
      break;
    case 'file-start': {
      const row = getOrCreateFileRow(progress.fileName);
      row.textContent = `変換中: ${progress.fileName}`;
      overallProgress.textContent = `[${progress.index}/${progress.total}]`;
      break;
    }
    case 'file-done': {
      const row = getOrCreateFileRow(progress.fileName);
      row.textContent = `完了: ${progress.fileName}`;
      appendLog(`完了: ${progress.fileName} -> ${progress.outPath}`);
      break;
    }
    case 'file-error': {
      const row = getOrCreateFileRow(progress.fileName);
      row.textContent = `エラー: ${progress.fileName}`;
      appendLog(`エラー: ${progress.fileName}: ${progress.error}`);
      break;
    }
    case 'done':
      appendLog('すべての処理が完了しました');
      updateConvertBtnState();
      break;
    default:
      break;
  }
});

convertBtn.addEventListener('click', async () => {
  if (!state.targetPath) return;

  convertBtn.disabled = true;
  logEl.textContent = '';

  try {
    await window.pollyAutogen.convert({
      mode: state.mode,
      targetPath: state.targetPath,
      options: getOptions(),
    });
  } catch (error) {
    appendLog(`実行エラー: ${error.message || error}`);
    convertBtn.disabled = false;
  }
});
