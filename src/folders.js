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
  const handle = await directory.getFileHandle(fileName, { create: true });
  const writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();
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
  for (let index = 0; index < active.length; index += 1) {
    const photo = active[index];
    const dirs = directories.get(photo.assetNumber);
    const baseName = fileNameFor(photo.file.name);
    const key = `${photo.assetNumber}/${baseName.toLowerCase()}`;
    const occurrence = (usedNames.get(key) ?? 0) + 1;
    usedNames.set(key, occurrence);
    const fileName = occurrence === 1 ? baseName : baseName.replace(/\.jpg$/i, `_${occurrence}.jpg`);
    const blob = await compress(photo.file);
    await writeBlob(dirs.assetDir, fileName, blob);
    if (photo.destination) {
      const destination = dirs.children.get(photo.destination);
      if (!destination) throw new Error(`${fileName} の写真帳分類先が不正です。`);
      await writeBlob(destination, fileName, blob);
    }
    onProgress(index + 1, active.length, fileName);
  }
  return { outputName: output.name, photoCount: active.length };
}
import { OTHER_ASSET } from "./domain.js";
