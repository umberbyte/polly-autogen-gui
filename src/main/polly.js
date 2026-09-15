const fs = require('node:fs/promises');
const path = require('node:path');
const { SynthesizeSpeechCommand } = require('@aws-sdk/client-polly');

const MAX_CHUNK_LENGTH = 2900;
const BOUNDARY_CHARS = ['。', '！', '？', '\n'];

// <prosody rate="N%">...</prosody> spans inserted via the spreadsheet editor.
// Storage (row.script / .txt files) keeps this raw and human-readable; only
// at synthesis time do we XML-escape the surrounding plain text and wrap it
// in <speak> for Polly's SSML input.
const PROSODY_PATTERN = /<prosody rate="(\d+)%">([\s\S]*?)<\/prosody>/g;

function containsSsml(text) {
  return /<prosody rate="\d+%">/.test(text);
}

function escapeXmlText(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function buildSsmlBody(rawText) {
  let result = '';
  let lastIndex = 0;
  let match;
  PROSODY_PATTERN.lastIndex = 0;
  // eslint-disable-next-line no-cond-assign
  while ((match = PROSODY_PATTERN.exec(rawText)) !== null) {
    result += escapeXmlText(rawText.slice(lastIndex, match.index));
    result += `<prosody rate="${match[1]}%">${escapeXmlText(match[2])}</prosody>`;
    lastIndex = PROSODY_PATTERN.lastIndex;
  }
  result += escapeXmlText(rawText.slice(lastIndex));
  return result;
}

const OUTPUT_EXTENSIONS = {
  mp3: 'mp3',
  ogg_vorbis: 'ogg',
  pcm: 'pcm',
};

function splitText(text, maxLength = MAX_CHUNK_LENGTH) {
  if (text.length <= maxLength) return [text];

  const chunks = [];
  let rest = text;

  while (rest.length > maxLength) {
    let cutAt = -1;
    for (let i = maxLength; i > 0; i -= 1) {
      if (BOUNDARY_CHARS.includes(rest[i - 1])) {
        cutAt = i;
        break;
      }
    }
    if (cutAt === -1) cutAt = maxLength;

    chunks.push(rest.slice(0, cutAt));
    rest = rest.slice(cutAt);
  }

  if (rest.length > 0) chunks.push(rest);

  return chunks;
}

async function listTxtFiles(dirPath) {
  const entries = await fs.readdir(dirPath, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === '.txt')
    .map((entry) => path.join(dirPath, entry.name))
    .sort();
}

async function streamToBuffer(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

async function synthesizeChunk(pollyClient, text, options) {
  const command = new SynthesizeSpeechCommand({
    Text: text,
    TextType: options.textType === 'ssml' ? 'ssml' : 'text',
    VoiceId: options.voiceId,
    Engine: options.engine,
    LanguageCode: options.languageCode,
    OutputFormat: options.outputFormat,
  });
  const response = await pollyClient.send(command);
  return streamToBuffer(response.AudioStream);
}

async function synthesizeToAudio(pollyClient, text, options) {
  if (containsSsml(text)) {
    const ssml = `<speak>${buildSsmlBody(text)}</speak>`;
    return synthesizeChunk(pollyClient, ssml, { ...options, textType: 'ssml' });
  }

  const chunks = splitText(text);
  const buffers = [];
  for (const chunk of chunks) {
    // eslint-disable-next-line no-await-in-loop
    buffers.push(await synthesizeChunk(pollyClient, chunk, options));
  }
  return Buffer.concat(buffers);
}

async function convertTextFile(pollyClient, filePath, options) {
  const text = await fs.readFile(filePath, 'utf-8');
  const audio = await synthesizeToAudio(pollyClient, text, options);

  const dir = path.dirname(filePath);
  const baseName = path.basename(filePath, path.extname(filePath));
  const ext = OUTPUT_EXTENSIONS[options.outputFormat] || 'mp3';
  const outDir = path.join(dir, 'mp3');
  const outPath = path.join(outDir, `${baseName}.${ext}`);

  await fs.mkdir(outDir, { recursive: true });
  await fs.writeFile(outPath, audio);

  return outPath;
}

async function resolveTargetFiles(mode, targetPath) {
  if (mode === 'directory') {
    return listTxtFiles(targetPath);
  }
  return [targetPath];
}

async function runConversion({ mode, targetPath, options }, pollyClient, onProgress) {
  const files = await resolveTargetFiles(mode, targetPath);
  const total = files.length;

  onProgress({ type: 'start', total });

  for (let i = 0; i < files.length; i += 1) {
    const filePath = files[i];
    const fileName = path.basename(filePath);
    onProgress({ type: 'file-start', index: i + 1, total, fileName });

    try {
      // eslint-disable-next-line no-await-in-loop
      const outPath = await convertTextFile(pollyClient, filePath, options);
      onProgress({ type: 'file-done', index: i + 1, total, fileName, outPath });
    } catch (error) {
      onProgress({
        type: 'file-error',
        index: i + 1,
        total,
        fileName,
        error: error.message || String(error),
      });
    }
  }

  onProgress({ type: 'done', total });
}

module.exports = {
  splitText,
  listTxtFiles,
  synthesizeToAudio,
  convertTextFile,
  runConversion,
  OUTPUT_EXTENSIONS,
  containsSsml,
  buildSsmlBody,
};
