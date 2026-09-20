import jsQR from "jsqr";
import { classifyDecodedEntries, naturalCompare } from "./domain.js";

function canvasToBlob(canvas, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("JPEG変換に失敗しました。")), "image/jpeg", quality);
  });
}

async function drawFile(file, maxSide) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
  const width = Math.max(1, Math.round(bitmap.width * scale));
  const height = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(bitmap, 0, 0, width, height);
  bitmap.close?.();
  return { canvas, context };
}

export async function decodeQrFromPhoto(file) {
  const { canvas, context } = await drawFile(file, 1400);
  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const result = jsQR(image.data, image.width, image.height, { inversionAttempts: "attemptBoth" });
  return result?.data ?? null;
}

export async function analyzePhotoFiles(fileList, assets, onProgress = () => {}) {
  const files = [...fileList]
    .filter((file) => file.type.startsWith("image/") || /\.(jpe?g|png|heic|heif)$/i.test(file.name))
    .sort((left, right) => naturalCompare(left.name, right.name));
  const decoded = new Array(files.length);
  let nextIndex = 0;
  let completed = 0;
  const worker = async () => {
    while (true) {
      const index = nextIndex++;
      if (index >= files.length) return;
      const file = files[index];
      let decodedValue = null;
      let qrReadError = false;
      try {
        decodedValue = await decodeQrFromPhoto(file);
      } catch {
        qrReadError = true;
      }
      // 並列処理の完了順ではなく、元のファイル順で必ず保持する。
      decoded[index] = { id: `${index}-${file.name}`, file, decodedValue, qrReadError };
      completed += 1;
      onProgress(completed, files.length, file.name);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  };
  await Promise.all(Array.from({ length: Math.min(2, files.length) }, () => worker()));
  return classifyDecodedEntries(decoded, assets.map((asset) => asset.assetNumber));
}

export async function compressToJpeg(file, maxBytes = 200 * 1024) {
  const bitmap = await createImageBitmap(file);
  let scale = Math.min(1, 3000 / Math.max(bitmap.width, bitmap.height));
  let lastBlob = null;

  for (let resize = 0; resize < 8; resize += 1) {
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    canvas.getContext("2d").drawImage(bitmap, 0, 0, width, height);

    for (const quality of [0.88, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.22]) {
      lastBlob = await canvasToBlob(canvas, quality);
      if (lastBlob.size <= maxBytes) {
        bitmap.close?.();
        return lastBlob;
      }
    }
    scale *= 0.82;
  }

  bitmap.close?.();
  if (!lastBlob || lastBlob.size > maxBytes) {
    throw new Error(`${file.name} を指定容量（${Math.round(maxBytes / 1024)}KB）以下へ圧縮できませんでした。`);
  }
  return lastBlob;
}

export function targetBytesFromKilobytes(value) {
  const kilobytes = Number(value);
  if (!Number.isFinite(kilobytes) || kilobytes < 50 || kilobytes > 5000) {
    throw new Error("写真容量は50～5000KBで指定してください。");
  }
  return Math.round(kilobytes * 1024);
}

export function jpegFileName(fileName) {
  return fileName.replace(/\.[^.]+$/, "") + ".jpg";
}
