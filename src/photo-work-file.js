const WORK_FILE_KIND = "asset-photo-organizer-work";
const WORK_FILE_VERSION = 1;

function fileToDataUrl(file) {
  if (!file) return null;
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({
      name: file.name,
      type: file.type,
      lastModified: file.lastModified,
      dataUrl: reader.result,
    });
    reader.onerror = () => reject(reader.error ?? new Error(`${file.name}を読み込めませんでした。`));
    reader.readAsDataURL(file);
  });
}

function dataUrlToFile(record) {
  if (!record?.dataUrl || typeof record.dataUrl !== "string") return null;
  const comma = record.dataUrl.indexOf(",");
  if (comma < 0) return null;
  const bytes = atob(record.dataUrl.slice(comma + 1));
  const buffer = new Uint8Array(bytes.length);
  for (let index = 0; index < bytes.length; index += 1) buffer[index] = bytes.charCodeAt(index);
  return new File([buffer], record.name || "保存ファイル", {
    type: record.type || "application/octet-stream",
    lastModified: record.lastModified || Date.now(),
  });
}

function safeFileName(value) {
  return String(value || "写真整理").replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").trim() || "写真整理";
}

export async function createPhotoWorkFile({ name, assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId }) {
  const photoMetadata = photos.map(({ file, ...photo }) => ({ ...photo, fileName: file?.name || "" }));
  const payload = {
    kind: WORK_FILE_KIND,
    version: WORK_FILE_VERSION,
    exportedAt: new Date().toISOString(),
    name: name || "",
    assets,
    healthWorkbookFile: null,
    albumTemplateFile: null,
    photoTargetKb: String(photoTargetKb ?? "200"),
    bookmarkPhotoId: bookmarkPhotoId ?? null,
    photos: photoMetadata,
  };
  return {
    blob: new Blob([JSON.stringify(payload)], { type: "application/json" }),
    fileName: `${safeFileName(name || "写真整理")}_作業ファイル.json`,
  };
}

export async function readPhotoWorkFile(file) {
  let payload;
  try {
    payload = JSON.parse(await file.text());
  } catch {
    throw new Error("作業ファイルを読み込めません。JSON形式のファイルを選択してください。");
  }
  if (payload?.kind !== WORK_FILE_KIND || payload.version !== WORK_FILE_VERSION) {
    throw new Error("この写真整理MVPで作成した作業ファイルではありません。");
  }
  const photos = (payload.photos ?? []).filter((photo) => photo?.fileName).map((photo) => ({ ...photo, file: null }));
  if (!Array.isArray(payload.assets) || !payload.assets.length || !photos.length) {
    throw new Error("作業ファイルに資産または写真がありません。");
  }
  return {
    ...payload,
    photos,
    healthWorkbookFile: null,
    albumTemplateFile: null,
  };
}
