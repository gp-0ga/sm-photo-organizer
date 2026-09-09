function canvasToBlob(canvas, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => blob ? resolve(blob) : reject(new Error("JPEG変換に失敗しました。")),
      "image/jpeg",
      quality,
    );
  });
}

export function cameraFileName(assetNumber, capturedAt = new Date()) {
  const stamp = [
    capturedAt.getFullYear(),
    String(capturedAt.getMonth() + 1).padStart(2, "0"),
    String(capturedAt.getDate()).padStart(2, "0"),
    "_",
    String(capturedAt.getHours()).padStart(2, "0"),
    String(capturedAt.getMinutes()).padStart(2, "0"),
    String(capturedAt.getSeconds()).padStart(2, "0"),
    "_",
    String(capturedAt.getMilliseconds()).padStart(3, "0"),
  ].join("");
  return `${assetNumber}_${stamp}.jpg`;
}

export function cropForZoom(width, height, zoom) {
  const safeZoom = Math.max(1, Number(zoom) || 1);
  const sourceWidth = width / safeZoom;
  const sourceHeight = height / safeZoom;
  return {
    x: (width - sourceWidth) / 2,
    y: (height - sourceHeight) / 2,
    width: sourceWidth,
    height: sourceHeight,
  };
}

export async function captureVideoFrame(video, { zoom = 1, maxSide = 2560, quality = 0.8 } = {}) {
  if (!video.videoWidth || !video.videoHeight) {
    throw new Error("カメラ映像を取得できません。");
  }

  const crop = cropForZoom(video.videoWidth, video.videoHeight, zoom);
  const scale = Math.min(1, maxSide / Math.max(crop.width, crop.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(crop.width * scale));
  canvas.height = Math.max(1, Math.round(crop.height * scale));
  canvas.getContext("2d").drawImage(
    video,
    crop.x,
    crop.y,
    crop.width,
    crop.height,
    0,
    0,
    canvas.width,
    canvas.height,
  );
  const blob = await canvasToBlob(canvas, quality);
  return { blob, width: canvas.width, height: canvas.height };
}

export async function openRearCamera() {
  if (!navigator.mediaDevices?.getUserMedia) {
    throw new Error("このブラウザはWebカメラに対応していません。");
  }
  return navigator.mediaDevices.getUserMedia({
    audio: false,
    video: {
      facingMode: { ideal: "environment" },
      width: { ideal: 3840 },
      height: { ideal: 2160 },
    },
  });
}

export function stopCamera(stream) {
  for (const track of stream?.getTracks?.() ?? []) track.stop();
}
