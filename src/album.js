import * as XLSX from "xlsx";

export async function inspectPhotoAlbumTemplate(file, assets) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellFormula: false, cellDates: false });
  if (!workbook.Sheets.No11) throw new Error("写真帳に`No11`シートが見つかりません。");

  const photoSheets = workbook.SheetNames.filter((name) => /^\d+_\d+$/.test(name));
  if (!photoSheets.length) throw new Error("写真帳シート（例：01_01）が見つかりません。");

  const mappedAssets = new Map();
  for (const sheetName of photoSheets) {
    const sheet = workbook.Sheets[sheetName];
    const value = sheet?.AB4?.v;
    const assetNumber = String(value ?? "").trim();
    if (assetNumber) {
      const range = XLSX.utils.decode_range(sheet["!ref"] ?? "A1:A1");
      const items = [];
      for (let row = 21; row <= range.e.r + 1; row += 14) {
        const itemNumber = String(sheet[`AP${row}`]?.v ?? "").trim();
        const inspectionItem = String(sheet[`AQ${row}`]?.v ?? "").trim();
        if (itemNumber) items.push({ itemNumber, inspectionItem });
      }
      mappedAssets.set(assetNumber, { sheetName, items });
    }
  }

  const missing = assets.filter((asset) => !mappedAssets.has(asset.assetNumber));
  if (missing.length) {
    const preview = missing.slice(0, 3).map((asset) => asset.assetNumber).join("、");
    throw new Error(`健全度判定表の資産と写真帳が一致しません。写真帳にない資産：${preview}${missing.length > 3 ? "ほか" : ""}`);
  }

  for (const asset of assets) {
    const template = mappedAssets.get(asset.assetNumber);
    const expected = new Set(asset.items.map((item) => item.itemNumber));
    const actual = new Set(template.items.map((item) => item.itemNumber));
    const missingItems = [...expected].filter((itemNumber) => !actual.has(itemNumber));
    const unexpected = template.items.filter((item) => !expected.has(item.itemNumber) && item.inspectionItem !== "経過年数");
    if (missingItems.length || unexpected.length) {
      throw new Error(`${asset.assetNumber}の確認項目が健全度判定表と写真帳で一致しません。`);
    }
  }

  return { sheetCount: photoSheets.length, mappedAssets };
}

export function albumEntries(assets, photos) {
  const assetMap = new Map(assets.map((asset) => [asset.assetNumber, asset]));
  const entries = [];
  const counts = new Map();

  for (const photo of photos) {
    if (photo.excluded || !photo.destination) continue;
    if (photo.reviewRequired || photo.qrReadError) throw new Error(`${photo.file.name}は要確認です。確認済みにしてから写真帳を作成してください。`);
    const asset = assetMap.get(photo.assetNumber);
    if (!asset) throw new Error(`${photo.file.name}の資産分類が不正です。`);

    let destinationType;
    let itemNumber = null;
    let limit;
    let countKey;
    if (photo.destination === "全景") {
      destinationType = "full";
      limit = 1;
      countKey = `${asset.assetNumber}/full`;
    } else {
      const item = asset.items.find((candidate) => candidate.folderName === photo.destination);
      if (!item) throw new Error(`${photo.file.name}の写真帳分類先が不正です。`);
      destinationType = "item";
      itemNumber = item.itemNumber;
      limit = 4;
      countKey = `${asset.assetNumber}/${itemNumber}`;
    }

    const count = (counts.get(countKey) ?? 0) + 1;
    counts.set(countKey, count);
    if (count > limit) {
      const label = destinationType === "full" ? "全景" : photo.destination;
      throw new Error(`${asset.assetNumber}の「${label}」は最大${limit}枚です。`);
    }
    entries.push({ photo, assetNumber: asset.assetNumber, destinationType, itemNumber, slotIndex: count - 1 });
  }

  if (!entries.length) throw new Error("写真帳に貼る写真が選ばれていません。");
  return entries;
}

async function postBinary(url, body) {
  const response = await fetch(url, { method: "POST", headers: { "Content-Type": "application/octet-stream" }, body });
  if (!response.ok) throw new Error(await response.text() || "ファイルの受け渡しに失敗しました。");
}

export async function createPhotoAlbum({ healthWorkbook, albumTemplate, assets, photos, compress, onProgress = () => {} }) {
  const entries = albumEntries(assets, photos);
  const sessionId = crypto.randomUUID();
  const base = `/api/session/${sessionId}`;

  onProgress(0, entries.length + 3, "健全度判定表を準備中");
  await postBinary(`${base}/health`, healthWorkbook);
  onProgress(1, entries.length + 3, "写真帳様式を準備中");
  await postBinary(`${base}/album`, albumTemplate);

  const manifestEntries = [];
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    const fileKey = String(index + 1).padStart(4, "0");
    const blob = await compress(entry.photo.file);
    await postBinary(`${base}/photo/${fileKey}`, blob);
    manifestEntries.push({
      fileKey,
      sourceName: entry.photo.file.name,
      assetNumber: entry.assetNumber,
      destinationType: entry.destinationType,
      itemNumber: entry.itemNumber,
      slotIndex: entry.slotIndex,
    });
    onProgress(index + 2, entries.length + 3, `${entry.photo.file.name}を準備中`);
  }

  const outputName = albumTemplate.name.replace(/\.xlsx$/i, "_写真貼付済.xlsx");
  const response = await fetch(`${base}/build`, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify({
      outputName,
      assets: assets.map((asset) => ({
        assetNumber: asset.assetNumber,
        itemNumbers: asset.items.map((item) => item.itemNumber),
      })),
      photos: manifestEntries,
    }),
  });
  if (!response.ok) {
    let message = "写真帳を作成できませんでした。";
    try { message = (await response.json()).message ?? message; } catch { message = await response.text() || message; }
    if (/Excel\.Application|not installed|could not be started|Class not registered/i.test(message)) {
      message = "Microsoft Excelを起動できません。Excelがインストールされた会社PCで実行してください。";
    }
    throw new Error(message);
  }

  const output = await response.blob();
  const url = URL.createObjectURL(output);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = outputName;
  anchor.click();
  URL.revokeObjectURL(url);
  onProgress(entries.length + 3, entries.length + 3, "写真帳を作成しました");
  return { outputName, photoCount: entries.length };
}
