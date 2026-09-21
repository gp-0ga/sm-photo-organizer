async function ensureDirectory(parent, name) {
  return parent.getDirectoryHandle(name, { create: true });
}

export function supportsDirectoryPicker() {
  return typeof window.showDirectoryPicker === "function";
}

export async function createAssetFolders(assets) {
  if (!supportsDirectoryPicker()) {
    throw new Error("このブラウザではフォルダ作成に対応していません。Microsoft EdgeまたはChromeを使用してください。");
  }
  const root = await window.showDirectoryPicker({ mode: "readwrite", id: "sm-photo-output" });
  const created = [];
  for (const asset of [...assets, OTHER_ASSET]) {
    const assetDir = await ensureDirectory(root, asset.folderName);
    await ensureDirectory(assetDir, "全景");
    for (const item of asset.items) {
      await ensureDirectory(assetDir, item.folderName);
    }
    created.push(asset.folderName);
  }
  return { rootName: root.name, created };
}

async function writeBlob(directory, fileName, blob) {
  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    let writable = null;
    try {
      const handle = await directory.getFileHandle(fileName, { create: true });
      writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return;
    } catch (error) {
      lastError = error;
      try { await writable?.abort(); } catch { /* 保存途中の一時状態を破棄 */ }
      if (attempt < 2) await new Promise((resolve) => setTimeout(resolve, 120 * (attempt + 1)));
    }
  }
  const message = String(lastError?.message ?? lastError ?? "");
  if (/state cached|changed since it was read from disk|状態が変わ/i.test(message)) {
    throw new Error("保存先フォルダの状態が変わりました。OneDrive同期中の場合は同期完了を待ち、同じフォルダを選び直して再実行してください。");
  }
  throw lastError;
}

function timestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

export async function exportOrganizedPhotos(assets, photos, compress, fileNameFor, onProgress = () => {}) {
  if (!supportsDirectoryPicker()) throw new Error("Microsoft EdgeまたはChromeを使用してください。");
  const unresolved = photos.filter((photo) => !photo.excluded && (!photo.assetNumber || photo.reviewRequired || photo.qrReadError));
  if (unresolved.length) {
    throw new Error(`未分類・要確認・読込失敗の写真が${unresolved.length}枚あります。すべて確認してから出力してください。`);
  }
  const selected = photos.filter((photo) => !photo.excluded && photo.assetNumber);
  for (const asset of assets) {
    if (asset.siteAbsent) continue;
    const assetPhotos = selected.filter((photo) => photo.assetNumber === asset.assetNumber);
    const fullCount = assetPhotos.filter((photo) => photoDestinations(photo).includes("全景")).length;
    if (fullCount !== 1) {
      throw new Error(`${asset.assetNumber}の「全景」は必ず1枚選択してください（現在${fullCount}枚）。`);
    }
    for (const item of asset.items) {
      const count = assetPhotos.filter((photo) => photoDestinations(photo).includes(item.folderName)).length;
      if (count > 4) throw new Error(`${asset.assetNumber}の「${item.folderName}」は最大4枚です。`);
    }
  }

  const selectedRoot = await window.showDirectoryPicker({ mode: "readwrite", id: "sm-photo-export" });
  const output = await ensureDirectory(selectedRoot, `写真整理出力_${timestamp()}`);
  const directories = new Map();
  for (const asset of [...assets, OTHER_ASSET]) {
    const assetDir = await ensureDirectory(output, asset.folderName);
    const children = new Map();
    children.set("全景", await ensureDirectory(assetDir, "全景"));
    for (const item of asset.items) children.set(item.folderName, await ensureDirectory(assetDir, item.folderName));
    directories.set(asset.assetNumber, { assetDir, children });
  }

  const usedNames = new Map();
  const active = photos.filter((photo) => !photo.excluded);
  const jobs = active.map((photo) => {
    const dirs = directories.get(photo.assetNumber);
    const baseName = fileNameFor(photo.file.name);
    const key = `${photo.assetNumber}/${baseName.toLowerCase()}`;
    const occurrence = (usedNames.get(key) ?? 0) + 1;
    usedNames.set(key, occurrence);
    const fileName = occurrence === 1 ? baseName : baseName.replace(/\.jpg$/i, `_${occurrence}.jpg`);
    return { photo, dirs, fileName };
  });
  let nextIndex = 0;
  let completed = 0;
  const worker = async () => {
    while (true) {
      const index = nextIndex;
      nextIndex += 1;
      if (index >= jobs.length) return;
      const { photo, dirs, fileName } = jobs[index];
      const blob = await compress(photo.file);
      await writeBlob(dirs.assetDir, fileName, blob);
      for (const destinationName of photoDestinations(photo)) {
        const destination = dirs.children.get(destinationName);
        if (!destination) throw new Error(`${fileName} の写真帳分類先が不正です。`);
        await writeBlob(destination, fileName, blob);
      }
      completed += 1;
      onProgress(completed, active.length, fileName);
    }
  };
  await Promise.all(Array.from({ length: Math.min(2, jobs.length) }, () => worker()));
  return { outputName: output.name, photoCount: active.length };
}
import { OTHER_ASSET, photoDestinations } from "./domain.js";
