const fs = require('node:fs/promises');
const path = require('node:path');
const { synthesizeToAudio } = require('./polly');

function mp3PathFor(workingFolder, prefix, pageNumber) {
  return path.join(workingFolder, 'mp3', `${prefix}_${pageNumber}.mp3`);
}

async function checkMp3Status(workingFolder, prefix, rows) {
  const status = {};
  for (const row of rows) {
    const mp3Path = mp3PathFor(workingFolder, prefix, row.pageNumber);
    // eslint-disable-next-line no-await-in-loop
    status[row.pageNumber] = await fs
      .access(mp3Path)
      .then(() => true)
      .catch(() => false);
  }
  return status;
}

async function anyMp3Exists(workingFolder, prefix, rows) {
  const status = await checkMp3Status(workingFolder, prefix, rows);
  return Object.values(status).some(Boolean);
}

async function synthesizeRowToMp3(pollyClient, workingFolder, prefix, row, voiceSettings) {
  const audio = await synthesizeToAudio(pollyClient, row.script, {
    voiceId: voiceSettings.voiceId,
    engine: voiceSettings.engine,
    languageCode: voiceSettings.languageCode,
    outputFormat: 'mp3',
  });

  const mp3Dir = path.join(workingFolder, 'mp3');
  await fs.mkdir(mp3Dir, { recursive: true });
  const outPath = mp3PathFor(workingFolder, prefix, row.pageNumber);
  await fs.writeFile(outPath, audio);
  return outPath;
}

module.exports = { mp3PathFor, checkMp3Status, anyMp3Exists, synthesizeRowToMp3 };
