const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');

const SETTINGS_FILE = path.join(app.getPath('userData'), 'voice-settings.json');

function loadVoiceSettings() {
  try {
    const raw = fs.readFileSync(SETTINGS_FILE, 'utf-8');
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function saveVoiceSettings(settings) {
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(settings));
}

module.exports = { loadVoiceSettings, saveVoiceSettings };
