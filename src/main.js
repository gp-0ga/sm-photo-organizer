import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { createAssetQr } from "./qr.js";
import { createQr } from "./qr.js";
import { createAssetFolders, exportOrganizedPhotos } from "./folders.js";
import { SPECIAL_MARKERS } from "./domain.js";
import { analyzePhotoFiles, compressToJpeg, jpegFileName } from "./photos.js";
import { albumEntries, createPhotoAlbum, inspectPhotoAlbumTemplate } from "./album.js";
import { cameraFileName, captureVideoFrame, openRearCamera, stopCamera } from "./camera.js";
import { createAssetPhotoZip, zipFileName } from "./zip.js";

const excelInput = document.querySelector("#excel-input");
const excelStatus = document.querySelector("#excel-status");
const folderStatus = document.querySelector("#folder-status");
const createFoldersButton = document.querySelector("#create-folders");
const printMarkersButton = document.querySelector("#print-markers");
const assetList = document.querySelector("#asset-list");
const markerList = document.querySelector("#marker-list");
const assetTemplate = document.querySelector("#asset-template");
const assetsEmpty = document.querySelector("#assets-empty");
const summary = document.querySelector("#summary");
const photoInput = document.querySelector("#photo-input");
const photoStatus = document.querySelector("#photo-status");
const photoSummary = document.querySelector("#photo-summary");
const photoList = document.querySelector("#photo-list");
const photoTemplate = document.querySelector("#photo-template");
const bulkTools = document.querySelector("#bulk-tools");
const selectAllPhotos = document.querySelector("#select-all-photos");
const bulkAsset = document.querySelector("#bulk-asset");
const applyBulkAsset = document.querySelector("#apply-bulk-asset");
const clearReview = document.querySelector("#clear-review");
const exportPhotosButton = document.querySelector("#export-photos");
const exportStatus = document.querySelector("#export-status");
const albumInput = document.querySelector("#album-input");
const albumTemplateStatus = document.querySelector("#album-template-status");
const albumStatus = document.querySelector("#album-status");
const createAlbumButton = document.querySelector("#create-album");
const cameraAsset = document.querySelector("#camera-asset");
const cameraWorkspace = document.querySelector("#camera-workspace");
const cameraPreview = document.querySelector("#camera-preview");
const cameraPlaceholder = document.querySelector("#camera-placeholder");
const cameraCurrent = document.querySelector("#camera-current");
const cameraStart = document.querySelector("#camera-start");
const cameraShutter = document.querySelector("#camera-shutter");
const cameraStop = document.querySelector("#camera-stop");
const cameraZoom = document.querySelector("#camera-zoom");
const cameraZoomValue = document.querySelector("#camera-zoom-value");
const cameraStatus = document.querySelector("#camera-status");
const cameraCounts = document.querySelector("#camera-counts");
const cameraSaveZip = document.querySelector("#camera-save-zip");
const cameraSaveStatus = document.querySelector("#camera-save-status");

let assets = [];
let photos = [];
let markerCount = 0;
let healthWorkbookFile = null;
let albumTemplateFile = null;
const thumbnailUrls = new Set();
let cameraStream = null;
let captureInProgress = false;

function setStatus(element, message, tone = "neutral") {
  element.hidden = false;
  element.className = `status ${tone}`;
  element.textContent = message;
}

function renderSummary() {
  const items = assets.reduce((total, asset) => total + asset.items.length, 0);
  summary.hidden = false;
  summary.innerHTML = `
    <div><strong>${assets.length}</strong><span>資産</span></div>
    <div><strong>${items}</strong><span>写真用点検項目</span></div>
    <div><strong>${assets.length + items}</strong><span>作成するサブフォルダ</span></div>
  `;
}

function renderAssets() {
  assetList.replaceChildren();
  for (const asset of assets) {
    const node = assetTemplate.content.cloneNode(true);
    node.querySelector(".asset-number").textContent = asset.assetNumber;
    node.querySelector(".asset-name").textContent = asset.assetName;
    node.querySelector(".item-count").textContent = `${asset.items.length + 1}フォルダ`;
    const items = node.querySelector(".item-list");
    const fullView = document.createElement("li");
    fullView.textContent = "全景";
    items.append(fullView);
    for (const item of asset.items) {
      const li = document.createElement("li");
      li.textContent = item.folderName;
      items.append(li);
    }
    assetList.append(node);
  }
  assetsEmpty.hidden = true;
  renderSummary();
}

function renderCameraAssets() {
  cameraAsset.innerHTML = `<option value="">資産を選択</option>${assets
    .map((asset) => `<option value="${asset.assetNumber}">${asset.assetNumber} ${asset.assetName}</option>`)
    .join("")}`;
  cameraAsset.disabled = false;
  cameraWorkspace.hidden = false;
  updateCameraCurrent();
  updateCameraCounts();
}

function currentCameraAsset() {
  return assets.find((asset) => asset.assetNumber === cameraAsset.value) ?? null;
}

function updateCameraCurrent() {
  const asset = currentCameraAsset();
  cameraCurrent.textContent = asset
    ? `撮影先：${asset.assetNumber} ${asset.assetName}`
    : "撮影する資産を選択してください";
  cameraShutter.disabled = !cameraStream || !asset || captureInProgress;
}

function updateCameraCounts() {
  const captured = photos.filter((photo) => photo.source === "camera");
  cameraSaveZip.disabled = captured.length === 0;
  if (!captured.length) {
    cameraCounts.innerHTML = `<p>まだ撮影していません。</p>`;
    return;
  }
  const counts = new Map();
  for (const photo of captured) counts.set(photo.assetNumber, (counts.get(photo.assetNumber) ?? 0) + 1);
  cameraCounts.innerHTML = [...counts.entries()]
    .map(([assetNumber, count]) => {
      const asset = assets.find((item) => item.assetNumber === assetNumber);
      return `<div><strong>${assetNumber} ${asset?.assetName ?? ""}</strong><span>${count}枚</span></div>`;
    })
    .join("");
}

async function renderMarkers() {
  markerList.replaceChildren();
  for (const asset of assets) {
    const card = document.createElement("article");
    card.className = "marker-card";
    const image = document.createElement("img");
    image.alt = `${asset.assetNumber} ${asset.assetName} 資産切替QR`;
    image.src = await createAssetQr(asset.assetNumber);
    const number = document.createElement("strong");
    number.textContent = asset.assetNumber;
    const name = document.createElement("span");
    name.textContent = asset.assetName;
    card.append(image, number, name);
    markerList.append(card);
  }
  for (const special of SPECIAL_MARKERS) {
    const card = document.createElement("article");
    card.className = `marker-card special ${special.type}`;
    const image = document.createElement("img");
    image.alt = `${special.label} QR`;
    image.src = await createQr(special.payload);
    const label = document.createElement("strong");
    label.textContent = special.label;
    const note = document.createElement("span");
    note.textContent = special.type === "review" ? "後でPC確認" : "特殊マーカー";
    card.append(image, label, note);
    markerList.append(card);
  }
}

function assetOptions(selectedValue = "") {
  const options = [{ value: "", label: "未分類" }];
  for (const asset of assets) options.push({ value: asset.assetNumber, label: `${asset.assetNumber} ${asset.assetName}` });
  return options.map(({ value, label }) => `<option value="${value}"${value === selectedValue ? " selected" : ""}>${label}</option>`).join("");
}

function destinationOptions(photo) {
  const asset = assets.find((item) => item.assetNumber === photo.assetNumber);
  const options = [{ value: "", label: "写真帳には使わない" }];
  if (asset) {
    options.push({ value: "全景", label: "全景" });
    for (const item of asset.items) options.push({ value: item.folderName, label: item.folderName });
  }
  return options.map(({ value, label }) => `<option value="${value}"${value === photo.destination ? " selected" : ""}>${label}</option>`).join("");
}

function updatePhotoSummary() {
  const active = photos.filter((photo) => !photo.excluded);
  const unclassified = active.filter((photo) => !photo.assetNumber).length;
  const review = active.filter((photo) => photo.reviewRequired || photo.qrReadError).length;
  photoSummary.hidden = false;
  photoSummary.innerHTML = `
    <div><strong>${photos.length}</strong><span>写真</span></div>
    <div><strong>${unclassified}</strong><span>未分類</span></div>
    <div><strong>${review}</strong><span>要確認・読込失敗</span></div>
  `;
  exportPhotosButton.disabled = photos.length === 0 || unclassified > 0 || review > 0;
  updateAlbumReadiness();
}

function updateAlbumReadiness() {
  createAlbumButton.disabled = true;
  if (!healthWorkbookFile) {
    setStatus(albumStatus, "先に健全度判定表を読み込んでください。", "neutral");
    return;
  }
  if (!albumTemplateFile) {
    setStatus(albumStatus, "写真帳様式を選択してください。", "neutral");
    return;
  }
  try {
    const entries = albumEntries(assets, photos);
    createAlbumButton.disabled = false;
    setStatus(albumStatus, `${entries.length}枚を写真帳へ貼り付けます。`, "success");
  } catch (error) {
    const message = error.message ?? String(error);
    const tone = message.includes("選ばれていません") ? "neutral" : "error";
    setStatus(albumStatus, message, tone);
  }
}

function renderPhotos() {
  for (const url of thumbnailUrls) URL.revokeObjectURL(url);
  thumbnailUrls.clear();
  photoList.replaceChildren();
  for (const photo of photos) {
    const node = photoTemplate.content.cloneNode(true);
    const card = node.querySelector(".photo-card");
    card.dataset.photoId = photo.id;
    if (photo.reviewRequired || photo.qrReadError) card.classList.add("needs-review");
    const thumbnail = node.querySelector(".photo-thumb");
    const url = URL.createObjectURL(photo.file);
    thumbnailUrls.add(url);
    thumbnail.src = url;
    thumbnail.alt = photo.file.name;
    node.querySelector(".photo-name").textContent = photo.file.name;
    const assetSelect = node.querySelector(".photo-asset");
    assetSelect.innerHTML = assetOptions(photo.assetNumber ?? "");
    const destinationSelect = node.querySelector(".photo-destination");
    destinationSelect.innerHTML = destinationOptions(photo);
    node.querySelector(".photo-exclude").checked = photo.excluded;
    node.querySelector(".review-badge").hidden = !(photo.reviewRequired || photo.qrReadError);

    assetSelect.addEventListener("change", () => {
      photo.assetNumber = assetSelect.value || null;
      photo.destination = "";
      destinationSelect.innerHTML = destinationOptions(photo);
      updatePhotoSummary();
    });
    destinationSelect.addEventListener("change", () => {
      photo.destination = destinationSelect.value;
      updateAlbumReadiness();
    });
    node.querySelector(".photo-exclude").addEventListener("change", (event) => {
      photo.excluded = event.target.checked;
      card.classList.toggle("excluded", photo.excluded);
      updatePhotoSummary();
    });
    photoList.append(node);
  }
  bulkTools.hidden = photos.length === 0;
  bulkAsset.innerHTML = assetOptions();
  updatePhotoSummary();
}

function selectedPhotoIds() {
  return new Set([...document.querySelectorAll(".photo-card")]
    .filter((card) => card.querySelector(".photo-select").checked)
    .map((card) => card.dataset.photoId));
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  createFoldersButton.disabled = true;
  printMarkersButton.disabled = true;
  setStatus(excelStatus, `${file.name} を読み込んでいます…`, "working");
  try {
    assets = await readAssetsFromWorkbook(file);
    healthWorkbookFile = file;
    albumTemplateFile = null;
    albumInput.value = "";
    renderAssets();
    renderCameraAssets();
    await renderMarkers();
    createFoldersButton.disabled = false;
    printMarkersButton.disabled = false;
    photoInput.disabled = false;
    albumInput.disabled = false;
    setStatus(photoStatus, "写真を端末1台分ずつ選択してください。", "neutral");
    setStatus(albumTemplateStatus, "この健全度判定表と対になる写真帳様式を選択してください。", "neutral");
    setStatus(excelStatus, `${assets.length}資産を読み込みました。元のExcelは変更していません。`, "success");
    updateAlbumReadiness();
  } catch (error) {
    assets = [];
    healthWorkbookFile = null;
    albumTemplateFile = null;
    albumInput.disabled = true;
    assetList.replaceChildren();
    markerList.replaceChildren();
    summary.hidden = true;
    assetsEmpty.hidden = false;
    cameraAsset.disabled = true;
    cameraWorkspace.hidden = true;
    setStatus(excelStatus, error.message ?? String(error), "error");
  }
});

cameraAsset.addEventListener("change", () => {
  updateCameraCurrent();
  const asset = currentCameraAsset();
  if (asset) setStatus(cameraStatus, `${asset.assetNumber} ${asset.assetName} へ振り分けます。`, "success");
});

cameraZoom.addEventListener("input", () => {
  const zoom = Number(cameraZoom.value);
  cameraZoomValue.textContent = `${zoom.toFixed(1)}×`;
  cameraPreview.style.transform = `scale(${zoom})`;
});

cameraStart.addEventListener("click", async () => {
  cameraStart.disabled = true;
  setStatus(cameraStatus, "カメラの使用を許可してください。", "working");
  try {
    cameraStream = await openRearCamera();
    cameraPreview.srcObject = cameraStream;
    await cameraPreview.play();
    cameraPlaceholder.hidden = true;
    cameraStop.disabled = false;
    updateCameraCurrent();
    setStatus(cameraStatus, "カメラを開きました。資産を確認して撮影してください。", "success");
  } catch (error) {
    cameraStart.disabled = false;
    setStatus(cameraStatus, error.message ?? String(error), "error");
  }
});

cameraStop.addEventListener("click", () => {
  stopCamera(cameraStream);
  cameraStream = null;
  cameraPreview.srcObject = null;
  cameraPreview.style.transform = "scale(1)";
  cameraPlaceholder.hidden = false;
  cameraStart.disabled = false;
  cameraStop.disabled = true;
  cameraZoom.value = "1";
  cameraZoomValue.textContent = "1.0×";
  updateCameraCurrent();
  setStatus(cameraStatus, "カメラを閉じました。", "neutral");
});

cameraShutter.addEventListener("click", async () => {
  const asset = currentCameraAsset();
  if (!cameraStream || !asset || captureInProgress) return;
  captureInProgress = true;
  updateCameraCurrent();
  setStatus(cameraStatus, `${asset.assetNumber} ${asset.assetName} へ保存中…`, "working");
  try {
    const capturedAt = new Date();
    const { blob, width, height } = await captureVideoFrame(cameraPreview, {
      zoom: Number(cameraZoom.value),
      maxSide: 2560,
      quality: 0.8,
    });
    const file = new File([blob], cameraFileName(asset.assetNumber, capturedAt), {
      type: "image/jpeg",
      lastModified: capturedAt.getTime(),
    });
    photos.push({
      id: `camera-${capturedAt.getTime()}-${photos.length}`,
      file,
      assetNumber: asset.assetNumber,
      unknownId: null,
      destination: "",
      excluded: false,
      reviewRequired: false,
      qrReadError: false,
      source: "camera",
    });
    renderPhotos();
    updateCameraCounts();
    if (navigator.vibrate) navigator.vibrate(35);
    setStatus(
      cameraStatus,
      `${asset.assetNumber} ${asset.assetName} へ保存しました（${width}×${height}・${(blob.size / 1024 / 1024).toFixed(2)}MB）。`,
      "success",
    );
  } catch (error) {
    setStatus(cameraStatus, error.message ?? String(error), "error");
  } finally {
    captureInProgress = false;
    updateCameraCurrent();
  }
});

cameraSaveZip.addEventListener("click", async () => {
  cameraSaveZip.disabled = true;
  setStatus(cameraSaveStatus, "資産別ZIPを作成中…", "working");
  try {
    const { blob, photoCount } = await createAssetPhotoZip(assets, photos);
    const name = zipFileName();
    const file = new File([blob], name, { type: "application/zip" });

    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({ files: [file], title: name });
      setStatus(cameraSaveStatus, `${photoCount}枚の資産別ZIPを共有画面へ渡しました。`, "success");
    } else {
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = name;
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      setStatus(cameraSaveStatus, `${photoCount}枚の資産別ZIPをダウンロードしました。`, "success");
    }
  } catch (error) {
    if (error?.name === "AbortError") setStatus(cameraSaveStatus, "保存をキャンセルしました。", "neutral");
    else setStatus(cameraSaveStatus, error.message ?? String(error), "error");
  } finally {
    updateCameraCounts();
  }
});

window.addEventListener("pagehide", () => stopCamera(cameraStream));

albumInput.addEventListener("change", async () => {
  const file = albumInput.files?.[0];
  albumTemplateFile = null;
  if (!file) {
    updateAlbumReadiness();
    return;
  }
  setStatus(albumTemplateStatus, `${file.name}を確認しています…`, "working");
  try {
    const result = await inspectPhotoAlbumTemplate(file, assets);
    albumTemplateFile = file;
    setStatus(albumTemplateStatus, `${result.sheetCount}枚の写真帳シートを確認しました。`, "success");
  } catch (error) {
    albumInput.value = "";
    setStatus(albumTemplateStatus, error.message ?? String(error), "error");
  }
  updateAlbumReadiness();
});

createFoldersButton.addEventListener("click", async () => {
  createFoldersButton.disabled = true;
  setStatus(folderStatus, "保存先を選択してください。既存ファイルは上書きしません。", "working");
  try {
    const result = await createAssetFolders(assets);
    setStatus(folderStatus, `${result.rootName} に ${result.created.length} 個の資産フォルダを作成しました。`, "success");
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus(folderStatus, "フォルダ作成をキャンセルしました。", "neutral");
    } else {
      setStatus(folderStatus, error.message ?? String(error), "error");
    }
  } finally {
    createFoldersButton.disabled = false;
  }
});

printMarkersButton.addEventListener("click", () => window.print());

photoInput.addEventListener("change", async () => {
  const files = photoInput.files;
  if (!files?.length) return;
  photoInput.disabled = true;
  setStatus(photoStatus, `${files.length}ファイルのQRを確認しています…`, "working");
  try {
    const result = await analyzePhotoFiles(files, assets, (done, total, name) => {
      setStatus(photoStatus, `${done}/${total} ${name} を確認中`, "working");
    });
    photos = result.photos;
    markerCount = result.markers.length;
    renderPhotos();
    setStatus(photoStatus, `${photos.length}枚の写真と${markerCount}枚のマーカーを読み取りました。`, "success");
  } catch (error) {
    setStatus(photoStatus, error.message ?? String(error), "error");
  } finally {
    photoInput.disabled = false;
  }
});

selectAllPhotos.addEventListener("change", () => {
  for (const checkbox of document.querySelectorAll(".photo-select")) checkbox.checked = selectAllPhotos.checked;
});

applyBulkAsset.addEventListener("click", () => {
  const ids = selectedPhotoIds();
  if (!ids.size || !bulkAsset.value) return;
  for (const photo of photos) {
    if (ids.has(photo.id)) {
      photo.assetNumber = bulkAsset.value;
      photo.destination = "";
    }
  }
  renderPhotos();
});

clearReview.addEventListener("click", () => {
  const ids = selectedPhotoIds();
  for (const photo of photos) {
    if (ids.has(photo.id)) {
      photo.reviewRequired = false;
      photo.qrReadError = false;
    }
  }
  renderPhotos();
});

exportPhotosButton.addEventListener("click", async () => {
  exportPhotosButton.disabled = true;
  setStatus(exportStatus, "保存先を選択してください。新しい出力フォルダを作成します。", "working");
  try {
    const result = await exportOrganizedPhotos(
      assets,
      photos,
      compressToJpeg,
      jpegFileName,
      (done, total, name) => setStatus(exportStatus, `${done}/${total} ${name} を出力中`, "working"),
    );
    setStatus(exportStatus, `${result.outputName} に${result.photoCount}枚を出力しました。原本は変更していません。`, "success");
  } catch (error) {
    if (error?.name === "AbortError") setStatus(exportStatus, "出力をキャンセルしました。", "neutral");
    else setStatus(exportStatus, error.message ?? String(error), "error");
  } finally {
    updatePhotoSummary();
  }
});

createAlbumButton.addEventListener("click", async () => {
  createAlbumButton.disabled = true;
  setStatus(albumStatus, "写真帳を準備しています…", "working");
  try {
    const result = await createPhotoAlbum({
      healthWorkbook: healthWorkbookFile,
      albumTemplate: albumTemplateFile,
      assets,
      photos,
      compress: compressToJpeg,
      onProgress: (done, total, name) => setStatus(albumStatus, `${done}/${total} ${name}`, "working"),
    });
    setStatus(albumStatus, `${result.outputName}を作成しました。${result.photoCount}枚を貼り付けています。`, "success");
  } catch (error) {
    setStatus(albumStatus, error.message ?? String(error), "error");
  } finally {
    try {
      albumEntries(assets, photos);
      createAlbumButton.disabled = !(healthWorkbookFile && albumTemplateFile);
    } catch {
      createAlbumButton.disabled = true;
    }
  }
});
