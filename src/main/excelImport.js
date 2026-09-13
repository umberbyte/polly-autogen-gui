const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const TALK_SCRIPT_SHEET_NAME = 'トークスクリプト';
const PAGE_NUMBER_COLUMN = 1; // A
const SCRIPT_COLUMN = 6; // F

function cellValueToString(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join('');
  }
  if (typeof value === 'object' && value.text !== undefined) {
    return String(value.text);
  }
  return String(value).trim();
}

async function parseTalkScriptSheet(excelPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  const sheet = workbook.getWorksheet(TALK_SCRIPT_SHEET_NAME);
  if (!sheet) {
    throw new Error(`シート「${TALK_SCRIPT_SHEET_NAME}」が見つかりません`);
  }

  const rows = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const pageNumber = cellValueToString(row.getCell(PAGE_NUMBER_COLUMN).value).trim();
    const script = cellValueToString(row.getCell(SCRIPT_COLUMN).value).trim();
    if (!/^\d+$/.test(pageNumber) || !script) continue;
    rows.push({ pageNumber, script });
  }

  return rows;
}

async function pathExists(targetPath) {
  try {
    await fs.stat(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function resolveUniqueFolder(baseFolder) {
  if (!(await pathExists(baseFolder))) return baseFolder;

  let suffix = 2;
  let candidate = `${baseFolder} (${suffix})`;
  // eslint-disable-next-line no-await-in-loop
  while (await pathExists(candidate)) {
    suffix += 1;
    candidate = `${baseFolder} (${suffix})`;
  }
  return candidate;
}

async function resolveWorkingFolder(selectedFolder) {
  const directTextDir = path.join(selectedFolder, 'text');
  try {
    const stat = await fs.stat(directTextDir);
    if (stat.isDirectory()) {
      return { workingFolder: selectedFolder, textDir: directTextDir };
    }
  } catch {
    // <selectedFolder>/text が無い場合は下のケースにフォールスルー
  }

  if (path.basename(selectedFolder) === 'text') {
    // 「text」フォルダ自体を選んでしまった場合、一つ上を作業フォルダとして扱う
    return { workingFolder: path.dirname(selectedFolder), textDir: selectedFolder };
  }

  return { workingFolder: selectedFolder, textDir: directTextDir };
}

async function loadRowsFromTextFolder(selectedFolder) {
  const { workingFolder, textDir } = await resolveWorkingFolder(selectedFolder);

  let entries;
  try {
    entries = await fs.readdir(textDir, { withFileTypes: true });
  } catch (error) {
    if (error.code === 'ENOENT') {
      return {
        workingFolder,
        rows: [],
        warnings: [`「text」フォルダが見つかりません: ${textDir}(作業フォルダの指定が正しいか、保存済みか確認してください)`],
      };
    }
    return {
      workingFolder,
      rows: [],
      warnings: [`「text」フォルダの読み込みに失敗しました: ${error.message || error}`],
    };
  }

  const rows = [];
  const warnings = [];
  const txtFiles = entries.filter(
    (entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === '.txt',
  );

  if (txtFiles.length === 0) {
    warnings.push(`「text」フォルダに.txtファイルが見つかりませんでした: ${textDir}`);
  }

  for (const entry of txtFiles) {
    const baseName = path.basename(entry.name, path.extname(entry.name));
    const lastUnderscore = baseName.lastIndexOf('_');
    if (lastUnderscore === -1) {
      warnings.push(`ファイル名からページ番号を判別できません: ${entry.name}`);
      continue;
    }

    const pageNumber = baseName.slice(lastUnderscore + 1);
    // eslint-disable-next-line no-await-in-loop
    const script = await fs.readFile(path.join(textDir, entry.name), 'utf-8');
    rows.push({ pageNumber, script });
  }

  rows.sort((a, b) => {
    const numA = Number(a.pageNumber);
    const numB = Number(b.pageNumber);
    if (!Number.isNaN(numA) && !Number.isNaN(numB)) return numA - numB;
    return a.pageNumber.localeCompare(b.pageNumber);
  });

  return { workingFolder, rows, warnings };
}

async function saveRowsToTextFolder(workingFolder, prefix, rows) {
  const textDir = path.join(workingFolder, 'text');
  await fs.mkdir(textDir, { recursive: true });

  for (const row of rows) {
    const filePath = path.join(textDir, `${prefix}_${row.pageNumber}.txt`);
    // eslint-disable-next-line no-await-in-loop
    await fs.writeFile(filePath, row.script);
  }

  return rows.length;
}

module.exports = {
  TALK_SCRIPT_SHEET_NAME,
  parseTalkScriptSheet,
  resolveUniqueFolder,
  loadRowsFromTextFolder,
  saveRowsToTextFolder,
};
