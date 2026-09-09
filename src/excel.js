import * as XLSX from "xlsx";
import { parseAssetsFromRows } from "./domain.js";

export async function readAssetsFromWorkbook(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellFormula: false, cellDates: false });
  const sheet = workbook.Sheets.No11;
  if (!sheet) {
    throw new Error("`No11`シートが見つかりません。");
  }
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });
  return parseAssetsFromRows(rows);
}
