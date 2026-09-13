const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

function getCredentialsPath() {
  return path.join(os.homedir(), '.aws', 'credentials');
}

function findDefaultSection(content) {
  const headerMatch = /^\[default\]\s*$/m.exec(content);
  if (!headerMatch) return null;

  const start = headerMatch.index;
  const afterHeader = start + headerMatch[0].length;
  const nextHeaderMatch = /^\[.*\]\s*$/m.exec(content.slice(afterHeader));
  const end = nextHeaderMatch ? afterHeader + nextHeaderMatch.index : content.length;

  return { start, afterHeader, end };
}

async function readDefaultCredentials() {
  let content;
  try {
    content = await fs.readFile(getCredentialsPath(), 'utf-8');
  } catch {
    return { accessKeyId: '', secretAccessKey: '' };
  }

  const section = findDefaultSection(content);
  if (!section) return { accessKeyId: '', secretAccessKey: '' };

  const body = content.slice(section.afterHeader, section.end);
  const accessKeyMatch = /^[ \t]*aws_access_key_id\s*=\s*(.*)$/m.exec(body);
  const secretKeyMatch = /^[ \t]*aws_secret_access_key\s*=\s*(.*)$/m.exec(body);

  return {
    accessKeyId: accessKeyMatch ? accessKeyMatch[1].trim() : '',
    secretAccessKey: secretKeyMatch ? secretKeyMatch[1].trim() : '',
  };
}

function upsertLine(body, key, value) {
  // 行頭の空白/タブのみ許容する([ \t]、\sは使わない)。\s は改行にもマッチするため、
  // 前の行との区切りの改行まで一致に含めてしまい、置換時にその改行を消してしまうバグになる。
  const lineRegex = new RegExp(`^[ \\t]*${key}\\s*=.*$`, 'm');
  const line = `${key} = ${value}`;
  if (lineRegex.test(body)) {
    return body.replace(lineRegex, line);
  }
  const trimmed = body.replace(/\s*$/, '');
  return `${trimmed}${trimmed ? '\n' : ''}${line}\n`;
}

async function saveDefaultCredentials({ accessKeyId, secretAccessKey }) {
  const filePath = getCredentialsPath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  let content = '';
  try {
    content = await fs.readFile(filePath, 'utf-8');
  } catch {
    content = '';
  }

  const section = findDefaultSection(content);
  if (!section) {
    const trimmed = content.replace(/\s*$/, '');
    content = `${trimmed}${trimmed ? '\n\n' : ''}[default]\naws_access_key_id = ${accessKeyId}\naws_secret_access_key = ${secretAccessKey}\n`;
  } else {
    let body = content.slice(section.afterHeader, section.end);
    body = upsertLine(body, 'aws_access_key_id', accessKeyId);
    body = upsertLine(body, 'aws_secret_access_key', secretAccessKey);
    content = content.slice(0, section.afterHeader) + body + content.slice(section.end);
  }

  await fs.writeFile(filePath, content, { mode: 0o600 });
  await fs.chmod(filePath, 0o600);
}

async function ensureCredentialsFileExists() {
  const filePath = getCredentialsPath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  try {
    await fs.access(filePath);
  } catch {
    await fs.writeFile(filePath, '[default]\naws_access_key_id = \naws_secret_access_key = \n', {
      mode: 0o600,
    });
  }
  return filePath;
}

module.exports = {
  getCredentialsPath,
  readDefaultCredentials,
  saveDefaultCredentials,
  ensureCredentialsFileExists,
};
