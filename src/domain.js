export const REQUIRED_HEADERS = ["資産番号", "項目番号", "資産名称", "点検項目"];
export const EXCLUDED_INSPECTION_ITEMS = new Set(["経過年数"]);

const WINDOWS_FORBIDDEN = /[<>:"/\\|?*\u0000-\u001F]/g;
const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i;

export function sanitizeFolderSegment(value) {
  let safe = String(value ?? "")
    .normalize("NFC")
    .replace(WINDOWS_FORBIDDEN, "_")
    .replace(/[. ]+$/g, "")
    .trim();
  if (!safe) safe = "名称未設定";
  if (WINDOWS_RESERVED.test(safe)) safe = `_${safe}`;
  return safe;
}

export function naturalCompare(left, right) {
  return String(left).localeCompare(String(right), "ja", { numeric: true, sensitivity: "base" });
}

export function findHeaderRow(rows) {
  for (let index = 0; index < Math.min(rows.length, 30); index += 1) {
    const values = rows[index].map((value) => String(value ?? "").trim());
    if (REQUIRED_HEADERS.every((header) => values.includes(header))) {
      return { index, values };
    }
  }
  throw new Error("No11シートに必要な見出し（資産番号・項目番号・資産名称・点検項目）が見つかりません。");
}

export function parseAssetsFromRows(rows) {
  const header = findHeaderRow(rows);
  const columns = Object.fromEntries(REQUIRED_HEADERS.map((name) => [name, header.values.indexOf(name)]));
  const assets = [];
  const byNumber = new Map();

  for (const row of rows.slice(header.index + 1)) {
    const assetNumber = String(row[columns["資産番号"]] ?? "").trim();
    const itemNumber = String(row[columns["項目番号"]] ?? "").trim();
    const assetName = String(row[columns["資産名称"]] ?? "").trim();
    const inspectionItem = String(row[columns["点検項目"]] ?? "").trim();

    if (!assetNumber || !assetName) continue;

    let asset = byNumber.get(assetNumber);
    if (!asset) {
      asset = {
        assetNumber,
        assetName,
        folderName: `${sanitizeFolderSegment(assetNumber)}_${sanitizeFolderSegment(assetName)}`,
        items: [],
      };
      byNumber.set(assetNumber, asset);
      assets.push(asset);
    }

    if (
      itemNumber &&
      inspectionItem &&
      !EXCLUDED_INSPECTION_ITEMS.has(inspectionItem) &&
      !asset.items.some((item) => item.itemNumber === itemNumber)
    ) {
      asset.items.push({
        itemNumber,
        inspectionItem,
        folderName: `${sanitizeFolderSegment(itemNumber)}_${sanitizeFolderSegment(inspectionItem)}`,
      });
    }
  }

  if (assets.length === 0) {
    throw new Error("No11シートから資産を読み取れませんでした。");
  }

  assets.sort((a, b) => naturalCompare(a.assetNumber, b.assetNumber));
  return assets;
}

export function markerPayload(assetNumber) {
  return `SM-ASSET|1|${assetNumber}`;
}

export const OTHER_ASSET_NUMBER = "__OTHER__";
export const OTHER_ASSET = {
  assetNumber: OTHER_ASSET_NUMBER,
  assetName: "その他",
  folderName: "その他",
  items: [],
};

export const SPECIAL_MARKERS = [
  { type: "review", label: "要確認", payload: "SM-REVIEW|1" },
  { type: "other", label: "その他", payload: "SM-OTHER|1" },
  { type: "unknown", label: "未確認01", payload: "SM-UNKNOWN|1|01" },
  { type: "unknown", label: "未確認02", payload: "SM-UNKNOWN|1|02" },
  { type: "unknown", label: "未確認03", payload: "SM-UNKNOWN|1|03" },
  { type: "end", label: "調査終了", payload: "SM-END|1" },
];

export function parseMarkerPayload(value) {
  const normalized = String(value ?? "").trim();
  const asset = /^SM-ASSET\|1\|(.+)$/.exec(normalized);
  if (asset) return { type: "asset", assetNumber: asset[1] };
  if (normalized === "SM-REVIEW|1") return { type: "review" };
  if (normalized === "SM-OTHER|1") return { type: "other" };
  const unknown = /^SM-UNKNOWN\|1\|(.+)$/.exec(normalized);
  if (unknown) return { type: "unknown", unknownId: unknown[1] };
  if (normalized === "SM-END|1") return { type: "end" };
  return null;
}

export function classifyDecodedEntries(entries, validAssetNumbers) {
  const valid = new Set(validAssetNumbers);
  const photos = [];
  const markers = [];
  let currentAssetNumber = null;
  let currentUnknownId = null;
  let currentSegment = [];

  for (const entry of entries) {
    const marker = parseMarkerPayload(entry.decodedValue);
    if (marker) {
      markers.push({ ...entry, marker });
      if (marker.type === "asset") {
        currentAssetNumber = valid.has(marker.assetNumber) ? marker.assetNumber : null;
        currentUnknownId = null;
        currentSegment = [];
      } else if (marker.type === "unknown") {
        currentAssetNumber = null;
        currentUnknownId = marker.unknownId;
        currentSegment = [];
      } else if (marker.type === "other") {
        currentAssetNumber = OTHER_ASSET_NUMBER;
        currentUnknownId = null;
        currentSegment = [];
      } else if (marker.type === "review") {
        for (const photo of currentSegment) photo.reviewRequired = true;
      } else if (marker.type === "end") {
        currentAssetNumber = null;
        currentUnknownId = null;
        currentSegment = [];
      }
      continue;
    }

    const photo = {
      id: entry.id,
      file: entry.file,
      assetNumber: currentAssetNumber,
      unknownId: currentUnknownId,
      destination: "",
      excluded: false,
      reviewRequired: false,
      qrReadError: entry.qrReadError ?? false,
    };
    photos.push(photo);
    currentSegment.push(photo);
  }

  return { photos, markers };
}
