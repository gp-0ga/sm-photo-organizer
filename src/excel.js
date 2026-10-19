import * as XLSX from "xlsx";
import { parseAssetsFromAssetListRows, parseAssetsFromRows, parseInspectionSheetRows } from "./domain.js";

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });
}

export async function readAssetsFromWorkbook(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellFormula: false, cellDates: false });
  const sheet = workbook.Sheets.No11;
  if (!sheet) {
    throw new Error("`No11`シートが見つかりません。");
  }
  return parseAssetsFromRows(sheetRows(sheet));
}

export async function readAssetsFromAssetListWorkbook(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellFormula: false, cellDates: false });
  let lastError;
  for (const sheetName of workbook.SheetNames) {
    try {
      return parseAssetsFromAssetListRows(sheetRows(workbook.Sheets[sheetName]));
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError ?? new Error("資産リストを読み取れませんでした。");
}

export async function readInspectionWorkbooks(files, onProgress = () => {}) {
  const inspections = [];
  const failures = [];
  const excelFiles = [...files].filter((file) => /\.xlsx$/i.test(file.name) && !file.name.startsWith("~$"));
  for (let index = 0; index < excelFiles.length; index += 1) {
    const file = excelFiles[index];
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellFormula: false, cellDates: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error("シートがありません。");
      inspections.push(parseInspectionSheetRows(sheetRows(sheet), file.name));
    } catch (error) {
      failures.push({ fileName: file.name, message: error.message ?? String(error) });
    }
    onProgress(index + 1, excelFiles.length);
  }
  if (!excelFiles.length) throw new Error("選択したフォルダーに.xlsx形式の点検表がありません。");
  return { inspections, failures, totalFiles: excelFiles.length };
}
