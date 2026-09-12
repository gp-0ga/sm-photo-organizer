import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { cameraFileName, captureVideoFrame, openRearCamera, stopCamera } from "./camera.js";
import { createAssetPhotoZip, zipFileName } from "./zip.js";

const excelInput = document.querySelector("#capture-excel-input");
const excelStatus = document.querySelector("#capture-excel-status");
const assetSelect = document.querySelector("#capture-asset");
const workspace = document.querySelector("#capture-workspace");
const preview = document.querySelector("#capture-preview");
const placeholder = document.querySelector("#capture-placeholder");
const current = document.querySelector("#capture-current");
const startButton = document.querySelector("#capture-start");
const shutterButton = document.querySelector("#capture-shutter");
const stopButton = document.querySelector("#capture-stop");
const zoomInput = document.querySelector("#capture-zoom");
const zoomValue = document.querySelector("#capture-zoom-value");
const captureStatus = document.querySelector("#capture-status");
const counts = document.querySelector("#capture-counts");
const saveButton = document.querySelector("#capture-save-zip");
const saveStatus = document.querySelector("#capture-save-status");

let assets = [];
let photos = [];
let stream = null;
let captureInProgress = false;

function setStatus(element, message, tone = "neutral") {
  element.className = `status ${tone}`;
  element.textContent = message;
}

function selectedAsset() {
  return assets.find((asset) => asset.assetNumber === assetSelect.value) ?? null;
}

function updateCurrent() {
  const asset = selectedAsset();
  current.textContent = asset
    ? `撮影先：${asset.assetNumber} ${asset.assetName}`
    : "撮影する資産を選択してください";
  shutterButton.disabled = !stream || !asset || captureInProgress;
}

function updateCounts() {
  saveButton.disabled = photos.length === 0;
  if (!photos.length) {
    counts.innerHTML = "<p>まだ撮影していません。</p>";
    return;
  }
  const byAsset = new Map();
  for (const photo of photos) byAsset.set(photo.assetNumber, (byAsset.get(photo.assetNumber) ?? 0) + 1);
  counts.replaceChildren();
  for (const [assetNumber, count] of byAsset.entries()) {
    const asset = assets.find((item) => item.assetNumber === assetNumber);
    const row = document.createElement("div");
    const label = document.createElement("strong");
    const value = document.createElement("span");
    label.textContent = `${assetNumber} ${asset?.assetName ?? ""}`;
    value.textContent = `${count}枚`;
    row.append(label, value);
    counts.append(row);
  }
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  setStatus(excelStatus, `${file.name} を端末内で読み込んでいます…`, "working");
  try {
    assets = await readAssetsFromWorkbook(file);
    assetSelect.replaceChildren(new Option("資産を選択", ""));
    for (const asset of assets) assetSelect.add(new Option(`${asset.assetNumber} ${asset.assetName}`, asset.assetNumber));
    assetSelect.disabled = false;
    workspace.hidden = false;
    setStatus(excelStatus, `${assets.length}資産を読み込みました。Excelは外部送信していません。`, "success");
    updateCurrent();
  } catch (error) {
    assets = [];
    assetSelect.disabled = true;
    workspace.hidden = true;
    setStatus(excelStatus, error.message ?? String(error), "error");
  }
});

assetSelect.addEventListener("change", () => {
  updateCurrent();
  const asset = selectedAsset();
  if (asset) setStatus(captureStatus, `${asset.assetNumber} ${asset.assetName} へ振り分けます。`, "success");
});

zoomInput.addEventListener("input", () => {
  const zoom = Number(zoomInput.value);
  zoomValue.textContent = `${zoom.toFixed(1)}×`;
  preview.style.transform = `scale(${zoom})`;
});

startButton.addEventListener("click", async () => {
  startButton.disabled = true;
  setStatus(captureStatus, "カメラの使用を許可してください。", "working");
  try {
    stream = await openRearCamera();
    preview.srcObject = stream;
    await preview.play();
    placeholder.hidden = true;
    stopButton.disabled = false;
    updateCurrent();
    setStatus(captureStatus, "カメラを開きました。", "success");
  } catch (error) {
    startButton.disabled = false;
    setStatus(captureStatus, error.message ?? String(error), "error");
  }
});

stopButton.addEventListener("click", () => {
  stopCamera(stream);
  stream = null;
  preview.srcObject = null;
  preview.style.transform = "scale(1)";
  placeholder.hidden = false;
  startButton.disabled = false;
  stopButton.disabled = true;
  zoomInput.value = "1";
  zoomValue.textContent = "1.0×";
  updateCurrent();
  setStatus(captureStatus, "カメラを閉じました。", "neutral");
});

shutterButton.addEventListener("click", async () => {
  const asset = selectedAsset();
  if (!stream || !asset || captureInProgress) return;
  captureInProgress = true;
  updateCurrent();
  setStatus(captureStatus, `${asset.assetNumber} ${asset.assetName} へ保存中…`, "working");
  try {
    const capturedAt = new Date();
    const { blob, width, height } = await captureVideoFrame(preview, {
      zoom: Number(zoomInput.value), maxSide: 2560, quality: 0.8,
    });
    const file = new File([blob], cameraFileName(asset.assetNumber, capturedAt), {
      type: "image/jpeg", lastModified: capturedAt.getTime(),
    });
    photos.push({ source: "camera", assetNumber: asset.assetNumber, excluded: false, file });
    updateCounts();
    if (navigator.vibrate) navigator.vibrate(35);
    setStatus(captureStatus, `${asset.assetNumber} ${asset.assetName} へ保存しました（${width}×${height}・${(blob.size / 1024 / 1024).toFixed(2)}MB）。`, "success");
  } catch (error) {
    setStatus(captureStatus, error.message ?? String(error), "error");
  } finally {
    captureInProgress = false;
    updateCurrent();
  }
});

saveButton.addEventListener("click", async () => {
  saveButton.disabled = true;
  setStatus(saveStatus, "資産別ZIPを作成中…", "working");
  try {
    const { blob, photoCount } = await createAssetPhotoZip(assets, photos);
    const name = zipFileName();
    const file = new File([blob], name, { type: "application/zip" });
    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({ files: [file], title: name });
      setStatus(saveStatus, `${photoCount}枚のZIPを共有画面へ渡しました。`, "success");
    } else {
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = name;
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      setStatus(saveStatus, `${photoCount}枚のZIPをダウンロードしました。`, "success");
    }
  } catch (error) {
    if (error?.name === "AbortError") setStatus(saveStatus, "保存をキャンセルしました。", "neutral");
    else setStatus(saveStatus, error.message ?? String(error), "error");
  } finally {
    updateCounts();
  }
});

window.addEventListener("pagehide", () => stopCamera(stream));
