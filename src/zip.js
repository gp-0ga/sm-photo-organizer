const encoder = new TextEncoder();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date) {
  const year = Math.max(1980, date.getFullYear());
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

function localHeader(name, bytes, checksum, modified) {
  const header = new Uint8Array(30 + name.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0x0800, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, modified.time, true);
  view.setUint16(12, modified.date, true);
  view.setUint32(14, checksum, true);
  view.setUint32(18, bytes.length, true);
  view.setUint32(22, bytes.length, true);
  view.setUint16(26, name.length, true);
  header.set(name, 30);
  return header;
}

function centralHeader(name, bytes, checksum, modified, offset) {
  const header = new Uint8Array(46 + name.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0x0800, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, modified.time, true);
  view.setUint16(14, modified.date, true);
  view.setUint32(16, checksum, true);
  view.setUint32(20, bytes.length, true);
  view.setUint32(24, bytes.length, true);
  view.setUint16(28, name.length, true);
  view.setUint32(42, offset, true);
  header.set(name, 46);
  return header;
}

export async function createZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.path.replaceAll("\\", "/"));
    const bytes = new Uint8Array(await entry.blob.arrayBuffer());
    const checksum = crc32(bytes);
    const modified = dosDateTime(entry.modified ?? new Date());
    const header = localHeader(name, bytes, checksum, modified);
    localParts.push(header, bytes);
    centralParts.push(centralHeader(name, bytes, checksum, modified, offset));
    offset += header.length + bytes.length;
  }

  const centralSize = centralParts.reduce((total, part) => total + part.length, 0);
  const end = new Uint8Array(22);
  const view = new DataView(end.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(8, entries.length, true);
  view.setUint16(10, entries.length, true);
  view.setUint32(12, centralSize, true);
  view.setUint32(16, offset, true);
  return new Blob([...localParts, ...centralParts, end], { type: "application/zip" });
}

export async function createAssetPhotoZip(assets, photos) {
  const assetByNumber = new Map([...assets, OTHER_ASSET].map((asset) => [asset.assetNumber, asset]));
  const captured = photos.filter((photo) => photo.source === "camera" && photo.assetNumber && !photo.excluded);
  if (!captured.length) throw new Error("保存する撮影写真がありません。");
  const entries = captured.map((photo) => {
    const asset = assetByNumber.get(photo.assetNumber);
    if (!asset) throw new Error(`${photo.assetNumber} の資産情報が見つかりません。`);
    return {
      path: `${asset.folderName}/${photo.file.name}`,
      blob: photo.file,
      modified: new Date(photo.file.lastModified || Date.now()),
    };
  });
  return { blob: await createZip(entries), photoCount: entries.length };
}

export function zipFileName(date = new Date()) {
  const day = [date.getFullYear(), String(date.getMonth() + 1).padStart(2, "0"), String(date.getDate()).padStart(2, "0")].join("");
  const time = [String(date.getHours()).padStart(2, "0"), String(date.getMinutes()).padStart(2, "0")].join("");
  return `資産写真_${day}_${time}.zip`;
}
import { OTHER_ASSET } from "./domain.js";
