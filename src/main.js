import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { createAssetFolders, exportOrganizedPhotos } from "./folders.js";
import { OTHER_ASSET, OTHER_ASSET_NUMBER, photoDestinations } from "./domain.js";
import { analyzePhotoFiles, compressToJpeg, jpegFileName, targetBytesFromKilobytes } from "./photos.js";
import { albumEntries, createPhotoAlbum, inspectPhotoAlbumTemplate } from "./album.js";
import { cameraFileName, captureVideoFrame, openRearCamera, stopCamera } from "./camera.js";
import { createAssetPhotoZip, zipFileName } from "./zip.js";
import { deletePhotoSnapshot, listPhotoSnapshots, loadPhotoSession, savePhotoSession, savePhotoSnapshot } from "./photo-session-storage.js";
import { createPhotoWorkFile, readPhotoWorkFile } from "./photo-work-file.js";

const hostedOrganizer = /\/organize(?:\.html)?$/i.test(window.location.pathname);
if (hostedOrganizer) document.body.classList.add("hosted-organizer");

const excelInput = document.querySelector("#excel-input");
const excelFileName = document.querySelector("#excel-file-name");
const excelStatus = document.querySelector("#excel-status");
const folderStatus = document.querySelector("#folder-status");
const createFoldersButton = document.querySelector("#create-folders");
const addAssetButton = document.querySelector("#add-asset");
const assetList = document.querySelector("#asset-list");
const assetTemplate = document.querySelector("#asset-template");
const assetsEmpty = document.querySelector("#assets-empty");
const excelPanel = document.querySelector("#excel-panel");
const assetsPanel = document.querySelector("#assets-panel");
const toggleExcelPanel = document.querySelector("#toggle-excel-panel");
const toggleAssetsPanel = document.querySelector("#toggle-assets-panel");
const summary = document.querySelector("#summary");
const photoInput = document.querySelector("#photo-input");
const photoFileName = document.querySelector("#photo-file-name");
const photoStatus = document.querySelector("#photo-status");
const sessionStatus = document.querySelector("#session-status");
const sessionName = document.querySelector("#session-name");
const saveSessionSnapshotButton = document.querySelector("#save-session-snapshot");
const savedSessionSelect = document.querySelector("#saved-session-select");
const loadSessionSnapshotButton = document.querySelector("#load-session-snapshot");
const deleteSessionSnapshotButton = document.querySelector("#delete-session-snapshot");
const jumpToBookmarkButton = document.querySelector("#jump-to-bookmark");
const exportWorkFileButton = document.querySelector("#export-work-file");
const chooseWorkFolderButton = document.querySelector("#choose-work-folder");
const importWorkFileButton = document.querySelector("#import-work-file");
const importWorkFileInput = document.querySelector("#import-work-file-input");
const photoSummary = document.querySelector("#photo-summary");
const photoDestinationAlert = document.querySelector("#photo-destination-alert");
const photoList = document.querySelector("#photo-list");
const gridDensity6Button = document.querySelector("#grid-density-6");
const gridDensity3Button = document.querySelector("#grid-density-3");
const photoViewMode = document.querySelector("#photo-view-mode");
const photoFilterAssetSelect = document.querySelector("#photo-filter-asset");
const photoFilterKeyword = document.querySelector("#photo-filter-keyword");
const photoFilterStatus = document.querySelector("#photo-filter-status");
const photoTemplate = document.querySelector("#photo-template");
const bulkTools = document.querySelector("#bulk-tools");
const bulkToolsSentinel = document.querySelector("#bulk-tools-sentinel");
const selectAllPhotos = document.querySelector("#select-all-photos");
const bulkSelectionStatus = document.querySelector("#bulk-selection-status");
const bulkAsset = document.querySelector("#bulk-asset");
const applyBulkAsset = document.querySelector("#apply-bulk-asset");
const clearReview = document.querySelector("#clear-review");
const exportPhotosButton = document.querySelector("#export-photos");
const exportStatus = document.querySelector("#export-status");
const exportToolbarStatus = document.querySelector("#export-toolbar-status");
const photoTargetKb = document.querySelector("#photo-target-kb");
const albumInput = document.querySelector("#album-input");
const albumTemplateStatus = document.querySelector("#album-template-status");
const albumStatus = document.querySelector("#album-status");
const createAlbumButton = document.querySelector("#create-album");
const preflightDialog = document.querySelector("#preflight-dialog");
const preflightList = document.querySelector("#preflight-list");
const showAssetsTabButton = document.querySelector("#show-assets-tab");
const workspaceTabs = [...document.querySelectorAll("[data-workspace-tab]")];
const workspacePanels = [...document.querySelectorAll("[data-workspace-panel]")];
const workspaceLayout = document.querySelector(".workspace-layout");
const workspaceSide = document.querySelector(".workspace-side");
const hideSidePanelButton = document.querySelector("#hide-side-panel");
const showSidePanelButton = document.querySelector("#show-side-panel");
const editDestinationOrderButton = document.querySelector("#edit-destination-order");
const orderDialog = document.querySelector("#order-dialog");
const orderDestinationSelect = document.querySelector("#order-destination-select");
const orderList = document.querySelector("#order-list");
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
let bookmarkPhotoId = null;
let healthWorkbookFile = null;
let albumTemplateFile = null;
let pendingWorkPhotoMetadata = null;
let pendingWorkBookmarkId = null;
const thumbnailUrls = new Set();
let cameraStream = null;
let captureInProgress = false;
let saveTimer = null;
let saveInProgress = false;
let saveQueued = false;
const selectedPhotoIdsState = new Set();
const PHOTO_RENDER_CHUNK_SIZE = 80;
let photoRenderQueue = [];
let photoRenderCursor = 0;
let photoRenderObserver = null;

function setupPanelToggle(panel, button, storageKey) {
  if (!panel || !button) return;
  const update = (collapsed) => {
    panel.classList.toggle("panel-collapsed", collapsed);
    button.textContent = collapsed ? "展開" : "折りたたむ";
    button.setAttribute("aria-expanded", String(!collapsed));
  };
  let collapsed = false;
  try { collapsed = localStorage.getItem(storageKey) === "1"; } catch { /* ignore */ }
  update(collapsed);
  button.addEventListener("click", () => {
    collapsed = !collapsed;
    update(collapsed);
    try { localStorage.setItem(storageKey, collapsed ? "1" : "0"); } catch { /* ignore */ }
  });
}

setupPanelToggle(excelPanel, toggleExcelPanel, "asset-photo-excel-panel-collapsed");
setupPanelToggle(assetsPanel, toggleAssetsPanel, "asset-photo-assets-panel-collapsed");

function setStatus(element, message, tone = "neutral") {
  element.hidden = false;
  element.className = `status ${tone}`;
  element.textContent = message;
}

function setExportStatus(message, tone = "neutral") {
  setStatus(exportStatus, message, tone);
  setStatus(exportToolbarStatus, message, tone);
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
    const warning = node.querySelector(".asset-warning");
    const warningLabels = [];
    if (asset.notInExcel) warningLabels.push("Excelに計上なし");
    if (asset.siteAbsent) warningLabels.push("現地なし／対象外");
    if (warningLabels.length) {
      warning.hidden = false;
      warning.textContent = warningLabels.join("・");
    }
    const siteAbsent = node.querySelector(".asset-site-absent");
    siteAbsent.checked = Boolean(asset.siteAbsent);
    siteAbsent.addEventListener("change", () => {
      asset.siteAbsent = siteAbsent.checked;
      renderAssets();
      updatePhotoDestinationAlert();
      scheduleSessionSave();
    });
    node.querySelector(".asset-edit").dataset.assetNumber = asset.assetNumber;
    node.querySelector(".asset-delete").dataset.assetNumber = asset.assetNumber;
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
  addAssetButton.disabled = false;
  renderSummary();
  bulkTools.hidden = assets.length === 0 && photos.length === 0;
  renderPhotoFilterOptions();
}

const UNCLASSIFIED_FILTER_VALUE = "__unclassified__";

function renderPhotoFilterOptions() {
  const previousValue = photoFilterAssetSelect.value;
  const options = [
    { value: "", label: "すべて表示" },
    { value: UNCLASSIFIED_FILTER_VALUE, label: "未分類" },
    ...assets.map((asset) => ({ value: asset.assetNumber, label: `${asset.assetNumber} ${asset.assetName}` })),
    { value: OTHER_ASSET_NUMBER, label: "その他（資産不明）" },
  ];
  photoFilterAssetSelect.innerHTML = options
    .map(({ value, label }) => `<option value="${value}">${label}</option>`)
    .join("");
  photoFilterAssetSelect.value = options.some((option) => option.value === previousValue) ? previousValue : "";
}

function matchesPhotoFilter(photo) {
  if (photoViewMode.value === "needs" && photo.assetNumber && !photo.reviewRequired && !photo.qrReadError) return false;
  const filterValue = photoFilterAssetSelect.value;
  if (filterValue === UNCLASSIFIED_FILTER_VALUE && photo.assetNumber) return false;
  if (filterValue && filterValue !== UNCLASSIFIED_FILTER_VALUE && photo.assetNumber !== filterValue) return false;
  return matchesPhotoKeyword(photo);
}

function matchesPhotoKeyword(photo) {
  const keyword = photoFilterKeyword?.value.trim().toLocaleLowerCase() || "";
  if (!keyword) return true;
  const asset = assets.find((item) => item.assetNumber === photo.assetNumber);
  const haystack = [photo.file.name, photo.assetNumber, asset?.assetName, ...photoDestinations(photo)].filter(Boolean).join(" ").toLocaleLowerCase();
  return haystack.includes(keyword);
}

function applyPhotoFilter() {
  let visibleCount = 0;
  for (const card of document.querySelectorAll(".photo-card")) {
    const photo = photos.find((item) => item.id === card.dataset.photoId);
    const matches = !photo || matchesPhotoFilter(photo);
    card.hidden = !matches;
    if (matches) visibleCount += 1;
  }
  const matchingCount = photos.filter(matchesPhotoFilter).length;
  const hasFilter = photoFilterAssetSelect.value || photoFilterKeyword?.value.trim();
  photoFilterStatus.textContent = hasFilter ? `${matchingCount}枚を表示中（全${photos.length}枚）` : "";
  for (const heading of photoList.querySelectorAll(".photo-group-heading")) {
    heading.hidden = ![...photoList.querySelectorAll(".photo-card")]
      .some((card) => card.dataset.groupKey === heading.dataset.groupKey && !card.hidden);
  }
  updateBulkControls();
}

function unclassifyPhotosForAssetNumbers(assetNumbers) {
  const removed = new Set(assetNumbers);
  for (const photo of photos) {
    if (!removed.has(photo.assetNumber)) continue;
    photo.assetNumber = null;
    setPhotoDestinations(photo, []);
    photo.reviewRequired = true;
  }
}

function mergeWorkbookAssets(nextAssets) {
  const nextNumbers = new Set(nextAssets.map((asset) => asset.assetNumber));
  const missingOldAssets = assets
    .filter((asset) => !nextNumbers.has(asset.assetNumber))
    .map((asset) => ({ ...asset, notInExcel: true }));
  unclassifyPhotosForAssetNumbers(missingOldAssets.map((asset) => asset.assetNumber));
  return [...nextAssets.map((asset) => ({ ...asset, notInExcel: false })), ...missingOldAssets];
}

function renderCameraAssets() {
  cameraAsset.innerHTML = `<option value="">資産を選択</option>${[...assets, OTHER_ASSET]
    .map((asset) => `<option value="${asset.assetNumber}">${asset.assetNumber} ${asset.assetName}</option>`)
    .join("")}`;
  cameraAsset.disabled = false;
  cameraWorkspace.hidden = false;
  updateCameraCurrent();
  updateCameraCounts();
}

function currentCameraAsset() {
  return [...assets, OTHER_ASSET].find((asset) => asset.assetNumber === cameraAsset.value) ?? null;
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
      const asset = [...assets, OTHER_ASSET].find((item) => item.assetNumber === assetNumber);
      return `<div><strong>${assetNumber} ${asset?.assetName ?? ""}</strong><span>${count}枚</span></div>`;
    })
    .join("");
}

function assetOptions(selectedValue = "") {
  const options = [{ value: "", label: "未分類" }];
  for (const asset of assets) options.push({ value: asset.assetNumber, label: `${asset.assetNumber} ${asset.assetName}` });
  options.push({ value: OTHER_ASSET_NUMBER, label: "その他（資産不明）" });
  return options.map(({ value, label }) => `<option value="${value}"${value === selectedValue ? " selected" : ""}>${label}</option>`).join("");
}

function destinationOptions(photo) {
  if (photo.assetNumber === OTHER_ASSET_NUMBER) return `<option value="">写真帳には使わない</option>`;
  const asset = assets.find((item) => item.assetNumber === photo.assetNumber);
  const selected = new Set(photoDestinations(photo));
  const options = [{ value: "", label: "写真帳には使わない" }];
  if (asset) {
    options.push({ value: "全景", label: "全景" });
    for (const item of asset.items) options.push({ value: item.folderName, label: item.folderName });
  }
  return options.filter(({ value }) => value === "" || selected.has(value)).map(({ value, label }) => `<span class="photo-destination-chip">${label}</span>`).join("");
}

function destinationChoices(photo) {
  if (photo.assetNumber === OTHER_ASSET_NUMBER) return "";
  const asset = assets.find((item) => item.assetNumber === photo.assetNumber);
  if (!asset) return "";
  const selected = new Set(photoDestinations(photo));
  const options = [{ value: "", label: "写真帳には使わない" }, ...asset.items.map((item) => ({ value: item.folderName, label: item.folderName }))];
  return `<label class="photo-full-choice"><input type="checkbox" class="photo-full-check"${selected.has("全景") ? " checked" : ""} /> 全景</label>
    <label class="photo-item-choice">サブフォルダー<select class="photo-item-select">${options.map(({ value, label }) => `<option value="${value}"${selected.has(value) ? " selected" : ""}>${label}</option>`).join("")}</select></label>`;
}

function destinationLimit(destination) {
  return destination === "全景" ? 1 : 4;
}

function destinationCount(photo, destination) {
  return photos.filter((candidate) =>
    candidate.id !== photo.id &&
    !candidate.excluded &&
    candidate.assetNumber === photo.assetNumber &&
    photoDestinations(candidate).includes(destination),
  ).length;
}

function setPhotoDestinations(photo, destinations) {
  photo.destinations = [...new Set(destinations.filter(Boolean))];
  photo.destination = photo.destinations[0] ?? "";
  photo.destinationOrder ??= {};
  photo.destinations.forEach((destination) => {
    if (!Number.isFinite(photo.destinationOrder[destination])) photo.destinationOrder[destination] = photos.indexOf(photo);
  });
}

function normalizePhotoDestinations(photo) {
  const current = photoDestinations(photo);
  const full = current.includes("全景") ? ["全景"] : [];
  const item = current.find((destination) => destination !== "全景");
  setPhotoDestinations(photo, [...full, item || ""]);
  return photo;
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
  updatePhotoDestinationAlert();
  exportPhotosButton.disabled = photos.length === 0 || unclassified > 0 || review > 0;
  updateAlbumReadiness();
}

function updatePhotoDestinationAlert() {
  if (!photoDestinationAlert) return;
  const issues = destinationIssues();
  if (!photos.length) {
    photoDestinationAlert.hidden = true;
    photoDestinationAlert.textContent = "";
    return;
  }
  if (!issues.length) {
    photoDestinationAlert.hidden = false;
    photoDestinationAlert.className = "photo-destination-alert is-ok";
    photoDestinationAlert.textContent = "確認済み：全景・サブフォルダーの枚数に問題ありません。";
    return;
  }
  const fullIssues = issues.filter((issue) => issue.title.includes("全景"));
  const otherIssues = issues.filter((issue) => !issue.title.includes("全景"));
  const labels = [
    ...fullIssues.slice(0, 4).map((issue) => `${issue.title}（${issue.detail}）`),
    ...otherIssues.slice(0, 2).map((issue) => `${issue.title}（${issue.detail}）`),
  ];
  const remaining = issues.length - labels.length;
  photoDestinationAlert.hidden = false;
  photoDestinationAlert.className = "photo-destination-alert";
  photoDestinationAlert.textContent = `出力前に確認：${labels.join("、")}${remaining > 0 ? `、ほか${remaining}件` : ""}`;
}

function destinationIssues() {
  const issues = [];
  const active = photos.filter((photo) => !photo.excluded);
  const unresolved = active.filter((photo) => !photo.assetNumber || photo.reviewRequired || photo.qrReadError);
  if (unresolved.length) issues.push({ title: "未分類・要確認の写真", detail: `${unresolved.length}枚（例：${unresolved.slice(0, 3).map((photo) => photo.file.name).join("、")}）` });
  for (const asset of assets) {
    if (asset.siteAbsent) continue;
    const assetPhotos = active.filter((photo) => photo.assetNumber === asset.assetNumber);
    const fullPhotos = assetPhotos.filter((photo) => photoDestinations(photo).includes("全景"));
    if (fullPhotos.length !== 1) issues.push({ title: `${asset.assetNumber} ${asset.assetName}：全景`, detail: `${fullPhotos.length}枚／必須1枚` });
    for (const item of asset.items) {
      const itemPhotos = assetPhotos.filter((photo) => photoDestinations(photo).includes(item.folderName));
      if (itemPhotos.length > 4) issues.push({ title: `${asset.assetNumber}：${item.folderName}`, detail: `${itemPhotos.length}枚／最大4枚` });
    }
  }
  return issues;
}

function showDestinationIssues() {
  const issues = destinationIssues();
  if (!issues.length) return false;
  preflightList.replaceChildren(...issues.map((issue) => {
    const item = document.createElement("div");
    item.className = "preflight-item";
    item.innerHTML = `<strong>${issue.title}</strong><span>${issue.detail}</span>`;
    return item;
  }));
  if (typeof preflightDialog.showModal === "function") preflightDialog.showModal();
  else setStatus(exportStatus, issues.map((issue) => `${issue.title}：${issue.detail}`).join(" / "), "error");
  return true;
}

function destinationGroups() {
  const groups = [];
  for (const asset of assets) {
    groups.push({ key: `${asset.assetNumber}|全景`, label: `${asset.assetNumber} ${asset.assetName} / 全景`, assetNumber: asset.assetNumber, destination: "全景" });
    for (const item of asset.items) groups.push({ key: `${asset.assetNumber}|${item.folderName}`, label: `${asset.assetNumber} ${asset.assetName} / ${item.folderName}`, assetNumber: asset.assetNumber, destination: item.folderName });
  }
  return groups.filter((group) => photos.some((photo) => !photo.excluded && photo.assetNumber === group.assetNumber && photoDestinations(photo).includes(group.destination)));
}

function renderDestinationOrder() {
  const group = destinationGroups().find((item) => item.key === orderDestinationSelect.value);
  orderList.replaceChildren();
  if (!group) return;
  const selected = photos
    .filter((photo) => !photo.excluded && photo.assetNumber === group.assetNumber && photoDestinations(photo).includes(group.destination))
    .sort((left, right) => (left.destinationOrder?.[group.destination] ?? photos.indexOf(left)) - (right.destinationOrder?.[group.destination] ?? photos.indexOf(right)));
  selected.forEach((photo, index) => {
    const row = document.createElement("div");
    row.className = "order-row";
    row.dataset.photoId = photo.id;
    row.innerHTML = `<span class="order-row-number">${index + 1}</span><span class="order-row-name">${photo.file.name}</span><button type="button" class="secondary order-up" ${index === 0 ? "disabled" : ""}>上へ</button><button type="button" class="secondary order-down" ${index === selected.length - 1 ? "disabled" : ""}>下へ</button>`;
    orderList.append(row);
  });
}

function openDestinationOrderEditor() {
  const groups = destinationGroups();
  orderDestinationSelect.replaceChildren(...groups.map((group) => new Option(group.label, group.key)));
  renderDestinationOrder();
  if (typeof orderDialog.showModal === "function") orderDialog.showModal();
}

function moveDestinationOrder(direction, photoId) {
  const group = destinationGroups().find((item) => item.key === orderDestinationSelect.value);
  if (!group) return;
  const selected = photos
    .filter((photo) => !photo.excluded && photo.assetNumber === group.assetNumber && photoDestinations(photo).includes(group.destination))
    .sort((left, right) => (left.destinationOrder?.[group.destination] ?? photos.indexOf(left)) - (right.destinationOrder?.[group.destination] ?? photos.indexOf(right)));
  const index = selected.findIndex((photo) => photo.id === photoId);
  const next = index + direction;
  if (index < 0 || next < 0 || next >= selected.length) return;
  [selected[index], selected[next]] = [selected[next], selected[index]];
  selected.forEach((photo, position) => {
    photo.destinationOrder ??= {};
    photo.destinationOrder[group.destination] = position;
  });
  scheduleSessionSave();
  renderDestinationOrder();
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

function formatSavedAt(value) {
  if (!value) return "";
  return new Intl.DateTimeFormat("ja-JP", { dateStyle: "short", timeStyle: "short" }).format(new Date(value));
}

function setSessionStatus(message, tone = "") {
  sessionStatus.className = `session-status${tone ? ` ${tone}` : ""}`;
  sessionStatus.textContent = message;
}


async function saveSessionNow() {
  if (!assets.length && !photos.length && !healthWorkbookFile && !albumTemplateFile) return;
  if (saveInProgress) {
    saveQueued = true;
    return;
  }
  saveInProgress = true;
  try {
    const savedAt = await savePhotoSession({
      assets,
      photos,
      healthWorkbookFile,
      albumTemplateFile,
      photoTargetKb: photoTargetKb.value,
      bookmarkPhotoId,
    });
    setSessionStatus(`自動保存済み：${formatSavedAt(savedAt)}`, "saved");
  } catch (error) {
    setSessionStatus(`自動保存できませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    saveInProgress = false;
    if (saveQueued) {
      saveQueued = false;
      scheduleSessionSave();
    }
  }
}

function scheduleSessionSave() {
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => { void saveSessionNow(); }, 350);
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

let workFolderHandle = null;

async function chooseWorkFolder() {
  if (typeof window.showDirectoryPicker !== "function") {
    setSessionStatus("このブラウザでは保存先フォルダを直接指定できません。ダウンロード後に移動してください。", "error");
    return null;
  }
  try {
    workFolderHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    setSessionStatus(`保存先を「${workFolderHandle.name}」に変更しました。`, "saved");
    return workFolderHandle;
  } catch (error) {
    if (error?.name !== "AbortError") setSessionStatus(`保存先を変更できませんでした：${error.message ?? String(error)}`, "error");
    return null;
  }
}

async function saveWorkFileToFolder(folderHandle, blob, fileName) {
  const fileHandle = await folderHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  try {
    await writable.write(blob);
  } finally {
    await writable.close();
  }
}

function applyLoadedSession(saved, message = "") {
    assets = saved.assets ?? [];
    photos = (saved.photos ?? []).map((photo) => normalizePhotoDestinations(photo));
    healthWorkbookFile = saved.healthWorkbookFile;
    albumTemplateFile = saved.albumTemplateFile;
    bookmarkPhotoId = photos.some((photo) => photo.id === saved.bookmarkPhotoId) ? saved.bookmarkPhotoId : null;
    excelFileName.textContent = healthWorkbookFile?.name || "未選択";
    photoFileName.textContent = photos.length ? `${photos.length}枚` : "未選択";
    if (saved.photoTargetKb) photoTargetKb.value = saved.photoTargetKb;
    renderAssets();
    renderCameraAssets();
    renderPhotos();
    if (bookmarkPhotoId) scrollToBookmark();
    createFoldersButton.disabled = assets.length === 0;
    photoInput.disabled = assets.length === 0;
    albumInput.disabled = assets.length === 0;
    setStatus(excelStatus, `${assets.length}資産を復元しました。`, "success");
    setStatus(photoStatus, `${photos.length}枚を復元しました。`, "success");
    setStatus(albumTemplateStatus, albumTemplateFile ? "写真帳様式も復元しました。" : "写真帳様式を選択してください。", "neutral");
    setSessionStatus(message || `作業を復元しました（${formatSavedAt(saved.savedAt)}）。`, "saved");
    updateAlbumReadiness();
}

async function refreshSnapshotList() {
  const snapshots = await listPhotoSnapshots();
  savedSessionSelect.replaceChildren(new Option(snapshots.length ? "保存済み作業を選択" : "保存済み作業はありません", ""));
  for (const snapshot of snapshots) {
    const label = `${snapshot.name || "名称未設定"}（${formatSavedAt(snapshot.savedAt)}・${snapshot.photos?.length ?? 0}枚）`;
    savedSessionSelect.add(new Option(label, snapshot.id));
  }
  loadSessionSnapshotButton.disabled = snapshots.length === 0;
  deleteSessionSnapshotButton.disabled = snapshots.length === 0;
}

async function restoreSession() {
  try {
    const saved = await loadPhotoSession();
    if (!saved || (!saved.assets?.length && !saved.photos?.length)) {
      await refreshSnapshotList();
      return;
    }
    applyLoadedSession(saved, `前回の作業を復元しました（${formatSavedAt(saved.savedAt)}）。`);
    await refreshSnapshotList();
  } catch (error) {
    setSessionStatus(`前回の作業を復元できませんでした：${error.message ?? String(error)}`, "error");
  }
}

function createPhotoCard(photo) {
  const node = photoTemplate.content.cloneNode(true);
  const card = node.querySelector(".photo-card");
  card.dataset.photoId = photo.id;
  node.querySelector(".photo-select").checked = selectedPhotoIdsState.has(photo.id);
  if (photo.reviewRequired || photo.qrReadError) card.classList.add("needs-review");
  const thumbnail = node.querySelector(".photo-thumb");
  const url = URL.createObjectURL(photo.file);
  thumbnailUrls.add(url);
  thumbnail.src = url;
  thumbnail.alt = photo.file.name;
  node.querySelector(".photo-name").textContent = photo.file.name;
  const photoDate = photo.capturedAt || photo.file.lastModified;
  node.querySelector(".photo-date").textContent = photoDate ? `撮影日時：${new Intl.DateTimeFormat("ja-JP", { dateStyle: "short", timeStyle: "short" }).format(new Date(photoDate))}` : "撮影日時：不明";
  const assetSelect = node.querySelector(".photo-asset");
  assetSelect.innerHTML = assetOptions(photo.assetNumber ?? "");
  const destinationButton = node.querySelector(".photo-destination");
  const destinationOptionsPanel = node.querySelector(".photo-destination-options");
  const renderDestinationChooser = () => {
    const destinations = photoDestinations(photo);
    destinationButton.textContent = destinations.length ? destinations.join("・") : "分類なし";
    destinationOptionsPanel.innerHTML = destinationChoices(photo);
    destinationButton.hidden = true;
    destinationOptionsPanel.dataset.open = "true";
    destinationOptionsPanel.hidden = false;
  };
  renderDestinationChooser();
  node.querySelector(".photo-exclude").checked = photo.excluded;
  node.querySelector(".review-badge").hidden = !(photo.reviewRequired || photo.qrReadError);
  const bookmarkButton = node.querySelector(".photo-bookmark");
  const isBookmarked = photo.id === bookmarkPhotoId;
  card.classList.toggle("bookmarked", isBookmarked);
  bookmarkButton.title = isBookmarked ? "ここまで確認済み（解除）" : "ここまで確認";
  bookmarkButton.setAttribute("aria-pressed", String(isBookmarked));
  bookmarkButton.addEventListener("click", () => {
    const previousBookmarkId = bookmarkPhotoId;
    bookmarkPhotoId = bookmarkPhotoId === photo.id ? null : photo.id;
    if (previousBookmarkId && previousBookmarkId !== photo.id) {
      const previousCard = photoList.querySelector(`.photo-card[data-photo-id="${previousBookmarkId}"]`);
      if (previousCard) {
        previousCard.classList.remove("bookmarked");
        const previousButton = previousCard.querySelector(".photo-bookmark");
        previousButton.title = "ここまで確認";
        previousButton.setAttribute("aria-pressed", "false");
      }
    }
    const marked = bookmarkPhotoId === photo.id;
    card.classList.toggle("bookmarked", marked);
    bookmarkButton.title = marked ? "ここまで確認済み（解除）" : "ここまで確認";
    bookmarkButton.setAttribute("aria-pressed", String(marked));
    jumpToBookmarkButton.disabled = !bookmarkPhotoId;
    scheduleSessionSave();
  });

  assetSelect.addEventListener("change", () => {
    photo.assetNumber = assetSelect.value || null;
    setPhotoDestinations(photo, []);
    renderDestinationChooser();
    updatePhotoSummary();
    scheduleSessionSave();
    applyPhotoFilter();
  });
  destinationButton.addEventListener("click", () => {
    destinationOptionsPanel.dataset.open = destinationOptionsPanel.hidden ? "true" : "false";
    destinationOptionsPanel.hidden = destinationOptionsPanel.dataset.open !== "true";
  });
  destinationOptionsPanel.addEventListener("change", (event) => {
    if (!event.target.matches(".photo-full-check, .photo-item-select")) return;
    const full = destinationOptionsPanel.querySelector(".photo-full-check")?.checked;
    const item = destinationOptionsPanel.querySelector(".photo-item-select")?.value || "";
    if (full && destinationCount(photo, "全景") >= destinationLimit("全景") && !photoDestinations(photo).includes("全景") && !photo.excluded) {
      event.target.checked = false;
      setStatus(albumStatus, "全景は最大1枚です。", "error");
      return;
    }
    if (item && destinationCount(photo, item) >= destinationLimit(item) && !photoDestinations(photo).includes(item) && !photo.excluded) {
      event.target.value = "";
      setStatus(albumStatus, `${item}は最大4枚です。`, "error");
      return;
    }
    setPhotoDestinations(photo, [full ? "全景" : "", item]);
    renderDestinationChooser();
    updatePhotoDestinationAlert();
    updateAlbumReadiness();
    scheduleSessionSave();
  });
  node.querySelector(".photo-exclude").addEventListener("change", (event) => {
    photo.excluded = event.target.checked;
    card.classList.toggle("excluded", photo.excluded);
    updatePhotoSummary();
    scheduleSessionSave();
  });
  node.querySelector(".photo-select").addEventListener("change", (event) => {
    if (event.target.checked) selectedPhotoIdsState.add(photo.id);
    else selectedPhotoIdsState.delete(photo.id);
    updateBulkControls();
  });
  return node;
}

function appendPhotoRenderChunk() {
  const sentinel = photoList.querySelector(".photo-render-sentinel");
  sentinel?.remove();
  const fragment = document.createDocumentFragment();
  const chunkSize = photoRenderObserver ? PHOTO_RENDER_CHUNK_SIZE : photoRenderQueue.length;
  const end = Math.min(photoRenderCursor + chunkSize, photoRenderQueue.length);
  for (; photoRenderCursor < end; photoRenderCursor += 1) {
    const entry = photoRenderQueue[photoRenderCursor];
    if (entry.type === "heading") {
      const heading = document.createElement("div");
      heading.className = "photo-group-heading";
      heading.dataset.groupKey = entry.groupKey;
      heading.innerHTML = `<strong>${entry.label}</strong><span>${entry.count}枚${entry.needs ? `・要確認 ${entry.needs}枚` : ""}</span>`;
      fragment.append(heading);
      continue;
    }
    const card = createPhotoCard(entry.photo);
    card.querySelector(".photo-card").dataset.groupKey = entry.groupKey || "";
    fragment.append(card);
  }
  photoList.append(fragment);
  if (photoRenderCursor < photoRenderQueue.length) {
    const next = document.createElement("div");
    next.className = "photo-render-sentinel";
    photoList.append(next);
    photoRenderObserver?.observe(next);
  }
  applyPhotoFilter();
}

function renderPhotos() {
  for (const url of thumbnailUrls) URL.revokeObjectURL(url);
  thumbnailUrls.clear();
  photoRenderObserver?.disconnect();
  photoList.replaceChildren();
  let displayPhotos = photos.slice();
  if (photoViewMode.value === "asset") {
    const assetOrder = new Map([...assets, OTHER_ASSET].map((asset, index) => [asset.assetNumber, index]));
    displayPhotos.sort((left, right) => {
      const leftGroup = left.assetNumber ?? "__unclassified__";
      const rightGroup = right.assetNumber ?? "__unclassified__";
      const groupCompare = (assetOrder.get(leftGroup) ?? assets.length + 1) - (assetOrder.get(rightGroup) ?? assets.length + 1);
      return groupCompare || photos.indexOf(left) - photos.indexOf(right);
    });
  }
  displayPhotos = displayPhotos.filter(matchesPhotoFilter);
  photoRenderQueue = [];
  let previousGroupKey = null;
  for (const photo of displayPhotos) {
    if (photoViewMode.value === "asset") {
      const groupKey = photo.assetNumber ?? "__unclassified__";
      if (groupKey !== previousGroupKey) {
        const asset = [...assets, OTHER_ASSET].find((item) => item.assetNumber === photo.assetNumber);
        const groupPhotos = displayPhotos.filter((item) => (item.assetNumber ?? "__unclassified__") === groupKey);
        const needs = groupPhotos.filter((item) => !item.assetNumber || item.reviewRequired || item.qrReadError).length;
        photoRenderQueue.push({ type: "heading", groupKey, label: asset ? `${asset.assetNumber} ${asset.assetName}` : "未分類", count: groupPhotos.length, needs });
        previousGroupKey = groupKey;
      }
      photoRenderQueue.push({ type: "photo", photo, groupKey });
    } else {
      photoRenderQueue.push({ type: "photo", photo, groupKey: "" });
    }
  }
  if (!photoRenderObserver && "IntersectionObserver" in window) {
    photoRenderObserver = new IntersectionObserver(([entry]) => {
      if (entry.isIntersecting) appendPhotoRenderChunk();
    }, { rootMargin: "600px 0px" });
  }
  photoRenderCursor = 0;
  appendPhotoRenderChunk();
  bulkTools.hidden = assets.length === 0 && photos.length === 0;
  bulkAsset.innerHTML = assetOptions();
  selectAllPhotos.checked = false;
  selectAllPhotos.indeterminate = false;
  updatePhotoSummary();
  updateBulkControls();
  jumpToBookmarkButton.disabled = !bookmarkPhotoId;
  applyPhotoFilter();
}

photoFilterAssetSelect.addEventListener("change", renderPhotos);
photoFilterKeyword?.addEventListener("input", renderPhotos);
photoViewMode.addEventListener("change", renderPhotos);

function selectWorkspaceTab(name) {
  for (const tab of workspaceTabs) {
    const active = tab.dataset.workspaceTab === name;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  }
  for (const panel of workspacePanels) panel.hidden = panel.dataset.workspacePanel !== name;
}

for (const tab of workspaceTabs) tab.addEventListener("click", () => selectWorkspaceTab(tab.dataset.workspaceTab));
if (showAssetsTabButton) showAssetsTabButton.addEventListener("click", () => selectWorkspaceTab("assets"));

function setSidePanelCollapsed(collapsed) {
  workspaceLayout?.classList.toggle("side-panel-collapsed", collapsed);
  if (workspaceSide) workspaceSide.setAttribute("aria-hidden", String(collapsed));
  if (hideSidePanelButton) hideSidePanelButton.hidden = collapsed;
  if (showSidePanelButton) showSidePanelButton.hidden = !collapsed;
}

hideSidePanelButton?.addEventListener("click", () => setSidePanelCollapsed(true));
showSidePanelButton?.addEventListener("click", () => setSidePanelCollapsed(false));
editDestinationOrderButton.addEventListener("click", openDestinationOrderEditor);
orderDestinationSelect.addEventListener("change", renderDestinationOrder);
orderList.addEventListener("click", (event) => {
  const button = event.target.closest("button");
  const row = event.target.closest(".order-row");
  if (!button || !row) return;
  moveDestinationOrder(button.classList.contains("order-up") ? -1 : 1, row.dataset.photoId);
});

function scrollToBookmark() {
  if (!bookmarkPhotoId) return;
  const card = photoList.querySelector(`.photo-card[data-photo-id="${bookmarkPhotoId}"]`);
  if (!card) return;
  card.scrollIntoView({ behavior: "smooth", block: "center" });
  card.classList.add("bookmark-flash");
  window.setTimeout(() => card.classList.remove("bookmark-flash"), 1200);
}

if (bulkToolsSentinel && "IntersectionObserver" in window) {
  new IntersectionObserver(
    ([entry]) => bulkTools.classList.toggle("is-stuck", !entry.isIntersecting),
    { threshold: 1 },
  ).observe(bulkToolsSentinel);
}

const BULK_TOOLS_COLLAPSE_KEY = "bulk-tools-collapsed";
const bulkToolsCollapseButton = document.querySelector("#bulk-tools-collapse");

function applyBulkToolsCollapsed(collapsed) {
  bulkTools.classList.toggle("collapsed", collapsed);
  bulkToolsCollapseButton.setAttribute("aria-expanded", String(!collapsed));
  bulkToolsCollapseButton.textContent = collapsed ? "▼" : "▲";
  bulkToolsCollapseButton.title = collapsed ? "ツールバーを開く" : "ツールバーを折りたたむ";
  try {
    localStorage.setItem(BULK_TOOLS_COLLAPSE_KEY, collapsed ? "1" : "0");
  } catch {
    // 折りたたみ設定を保存できない場合も表示は継続する
  }
}

bulkToolsCollapseButton.addEventListener("click", () => {
  applyBulkToolsCollapsed(!bulkTools.classList.contains("collapsed"));
});

try {
  applyBulkToolsCollapsed(localStorage.getItem(BULK_TOOLS_COLLAPSE_KEY) === "1");
} catch {
  applyBulkToolsCollapsed(false);
}

function visiblePhotoCards() {
  return [...document.querySelectorAll(".photo-card:not([hidden])")];
}

function selectedPhotoIds() {
  const visibleIds = new Set(visiblePhotoCards().map((card) => card.dataset.photoId));
  return new Set([...selectedPhotoIdsState].filter((id) => visibleIds.has(id)));
}

function updateBulkControls() {
  const visibleIds = photos.filter(matchesPhotoFilter).map((photo) => photo.id);
  const selectedCount = visibleIds.filter((id) => selectedPhotoIdsState.has(id)).length;
  const total = visibleIds.length;
  bulkSelectionStatus.textContent = `${selectedCount}枚選択中`;
  selectAllPhotos.indeterminate = selectedCount > 0 && selectedCount < total;
  selectAllPhotos.checked = total > 0 && selectedCount === total;
  bulkAsset.disabled = selectedCount === 0;
  applyBulkAsset.disabled = selectedCount === 0 || !bulkAsset.value;
  clearReview.disabled = selectedCount === 0;
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  excelFileName.textContent = file.name;
  createFoldersButton.disabled = true;
  setStatus(excelStatus, `${file.name} を読み込んでいます…`, "working");
  try {
    const workbookAssets = await readAssetsFromWorkbook(file);
    const merged = assets.length > 0;
    assets = merged ? mergeWorkbookAssets(workbookAssets) : workbookAssets;
    healthWorkbookFile = file;
    albumTemplateFile = null;
    albumInput.value = "";
    renderAssets();
    renderCameraAssets();
    if (merged && photos.length) renderPhotos();
    createFoldersButton.disabled = false;
    photoInput.disabled = false;
    albumInput.disabled = false;
    setStatus(photoStatus, "写真を選択してください。", "neutral");
    setStatus(albumTemplateStatus, "この健全度判定表と対になる写真帳様式を選択してください。", "neutral");
    setStatus(
      excelStatus,
      merged
        ? `${assets.length}資産に更新しました。既存の写真の振り分けは維持されます。元のExcelは変更していません。`
        : `${assets.length}資産を読み込みました。元のExcelは変更していません。`,
      "success",
    );
    updateAlbumReadiness();
    scheduleSessionSave();
  } catch (error) {
    assets = [];
    addAssetButton.disabled = true;
    healthWorkbookFile = null;
    albumTemplateFile = null;
    albumInput.disabled = true;
    assetList.replaceChildren();
    summary.hidden = true;
    assetsEmpty.hidden = false;
    cameraAsset.disabled = true;
    cameraWorkspace.hidden = true;
    setStatus(excelStatus, error.message ?? String(error), "error");
  }
});

addAssetButton.addEventListener("click", () => {
  const assetNumber = window.prompt("追加する資産番号を入力してください。")?.trim();
  if (!assetNumber) return;
  if (assets.some((asset) => asset.assetNumber === assetNumber)) {
    setStatus(folderStatus, "同じ資産番号がすでにあります。", "error");
    return;
  }
  const assetName = window.prompt("資産名を入力してください。")?.trim();
  if (!assetName) return;
  assets.push({
    assetNumber,
    assetName,
    folderName: `${assetNumber}_${assetName}`,
    items: [],
    notInExcel: true,
  });
  assets.sort((a, b) => String(a.assetNumber).localeCompare(String(b.assetNumber), "ja", { numeric: true }));
  renderAssets();
  renderCameraAssets();
  scheduleSessionSave();
});

assetList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-asset-number]");
  if (!button) return;
  const assetNumber = button.dataset.assetNumber;
  const index = assets.findIndex((asset) => asset.assetNumber === assetNumber);
  if (index < 0) return;
  if (button.classList.contains("asset-delete")) {
    if (!window.confirm(`「${assets[index].assetName}」を削除しますか？分類済み写真は未分類に戻ります。`)) return;
    unclassifyPhotosForAssetNumbers([assetNumber]);
    assets.splice(index, 1);
  } else {
    const asset = assets[index];
    const nextNumber = window.prompt("資産番号を入力してください。", asset.assetNumber)?.trim();
    if (!nextNumber || (nextNumber !== asset.assetNumber && assets.some((item) => item.assetNumber === nextNumber))) return;
    const nextName = window.prompt("資産名を入力してください。", asset.assetName)?.trim();
    if (!nextName) return;
    if (nextNumber !== asset.assetNumber) unclassifyPhotosForAssetNumbers([asset.assetNumber]);
    assets[index] = { ...asset, assetNumber: nextNumber, assetName: nextName, folderName: `${nextNumber}_${nextName}`, notInExcel: true };
  }
  renderAssets();
  renderCameraAssets();
  renderPhotos();
  scheduleSessionSave();
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
      destinations: [],
      destinationOrder: {},
      excluded: false,
      reviewRequired: false,
      qrReadError: false,
      source: "camera",
      capturedAt: capturedAt.toISOString(),
    });
    // 撮影時は既存カードを作り直さず、新しい1枚だけを追加する。
    // 大量の写真を扱う現場で、既存サムネイルの再読込を避けるため。
    photoList.append(createPhotoCard(photos.at(-1)));
    bulkTools.hidden = false;
    bulkAsset.innerHTML = assetOptions();
    updatePhotoSummary();
    updateBulkControls();
    jumpToBookmarkButton.disabled = !bookmarkPhotoId;
    scheduleSessionSave();
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

window.addEventListener("pagehide", () => {
  stopCamera(cameraStream);
  void saveSessionNow();
});

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
  scheduleSessionSave();
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

photoInput.addEventListener("change", async () => {
  const files = photoInput.files;
  if (!files?.length) return;
  photoFileName.textContent = `${files.length}枚を選択中`;
  photoInput.disabled = true;
  setStatus(photoStatus, `${files.length}ファイルのQRを確認しています…`, "working");
  try {
    const result = await analyzePhotoFiles(files, assets, (done, total, name) => {
      setStatus(photoStatus, `${done}/${total} ${name} を確認中`, "working");
    });
    photos = result.photos;
    if (pendingWorkPhotoMetadata?.length) {
      const byName = new Map(pendingWorkPhotoMetadata.map((photo) => [photo.fileName, photo]));
      const importedBookmark = pendingWorkPhotoMetadata.find((photo) => photo.id === pendingWorkBookmarkId)?.fileName;
      photos = photos.map((photo) => {
        const saved = byName.get(photo.file.name);
        return saved ? normalizePhotoDestinations({ ...photo, assetNumber: saved.assetNumber ?? null, destination: saved.destination || saved.destinations?.[0] || "", destinations: Array.isArray(saved.destinations) ? saved.destinations : (saved.destination ? [saved.destination] : []), destinationOrder: saved.destinationOrder || {}, excluded: Boolean(saved.excluded), reviewRequired: Boolean(saved.reviewRequired), qrReadError: Boolean(saved.qrReadError) }) : photo;
      });
      bookmarkPhotoId = importedBookmark ? photos.find((photo) => photo.file.name === importedBookmark)?.id ?? null : null;
      pendingWorkPhotoMetadata = null;
      pendingWorkBookmarkId = null;
    }
    markerCount = result.markers.length;
    bookmarkPhotoId = null;
    renderPhotos();
    photoFileName.textContent = `${photos.length}枚`;
    setStatus(photoStatus, `${photos.length}枚の写真と${markerCount}枚のマーカーを読み取りました。`, "success");
  } catch (error) {
    setStatus(photoStatus, error.message ?? String(error), "error");
  } finally {
    photoInput.disabled = false;
    scheduleSessionSave();
  }
});

selectAllPhotos.addEventListener("change", () => {
  for (const photo of photos) {
    if (!matchesPhotoFilter(photo)) continue;
    if (selectAllPhotos.checked) selectedPhotoIdsState.add(photo.id);
    else selectedPhotoIdsState.delete(photo.id);
  }
  for (const card of document.querySelectorAll(".photo-card")) {
    card.querySelector(".photo-select").checked = selectedPhotoIdsState.has(card.dataset.photoId);
  }
  updateBulkControls();
});

bulkAsset.addEventListener("change", updateBulkControls);

applyBulkAsset.addEventListener("click", () => {
  const ids = selectedPhotoIds();
  if (!ids.size || !bulkAsset.value) return;
  for (const photo of photos) {
    if (ids.has(photo.id)) {
      photo.assetNumber = bulkAsset.value;
      setPhotoDestinations(photo, []);
      const card = photoList.querySelector(`.photo-card[data-photo-id="${photo.id}"]`);
      if (!card) continue;
      card.querySelector(".photo-asset").value = photo.assetNumber;
      card.querySelector(".photo-destination").textContent = "分類なし";
      const destinationPanel = card.querySelector(".photo-destination-options");
      destinationPanel.innerHTML = destinationChoices(photo);
      destinationPanel.dataset.open = "true";
      destinationPanel.hidden = false;
    }
  }
  updatePhotoSummary();
  updateBulkControls();
  scheduleSessionSave();
  applyPhotoFilter();
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
  scheduleSessionSave();
});

exportPhotosButton.addEventListener("click", async () => {
  exportPhotosButton.disabled = true;
  try {
    if (showDestinationIssues()) return;
    const targetBytes = targetBytesFromKilobytes(photoTargetKb.value);
    setExportStatus(`保存先を選択してください。1枚${photoTargetKb.value}KB以下で新しい出力フォルダを作成します。`, "working");
    const result = await exportOrganizedPhotos(
      assets,
      photos,
      (file) => compressToJpeg(file, targetBytes),
      jpegFileName,
      (done, total, name) => setExportStatus(`${done}/${total} ${name} を出力中`, "working"),
    );
    setExportStatus(`${result.outputName} に${result.photoCount}枚を${photoTargetKb.value}KB以下で出力しました。元写真は変更していません。`, "success");
  } catch (error) {
    if (error?.name === "AbortError") setExportStatus("出力をキャンセルしました。", "neutral");
    else setExportStatus(error.message ?? String(error), "error");
  } finally {
    updatePhotoSummary();
  }
});

createAlbumButton.addEventListener("click", async () => {
  createAlbumButton.disabled = true;
  setStatus(albumStatus, "写真帳を準備しています…", "working");
  try {
    if (showDestinationIssues()) return;
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

photoTargetKb.addEventListener("change", scheduleSessionSave);

saveSessionSnapshotButton.addEventListener("click", async () => {
  if (!assets.length && !photos.length) {
    setSessionStatus("保存する作業がありません。先にExcelまたは写真を読み込んでください。", "error");
    return;
  }
  saveSessionSnapshotButton.disabled = true;
  try {
    const name = sessionName.value.trim() || `写真整理_${new Date().toLocaleDateString("ja-JP")}`;
    const snapshot = await savePhotoSnapshot({
      name,
      assets,
      photos,
      healthWorkbookFile,
      albumTemplateFile,
      photoTargetKb: photoTargetKb.value,
      bookmarkPhotoId,
    });
    sessionName.value = name;
    setSessionStatus(`「${name}」として保存しました。後から一覧から開けます。`, "saved");
    await refreshSnapshotList();
    const option = [...savedSessionSelect.options].find((item) => item.value === snapshot.id);
    if (option) savedSessionSelect.value = snapshot.id;
  } catch (error) {
    setSessionStatus(`作業を保存できませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    saveSessionSnapshotButton.disabled = false;
  }
});

jumpToBookmarkButton.addEventListener("click", scrollToBookmark);

const GRID_DENSITY_KEY = "photo-grid-density";

function applyGridDensity(density) {
  photoList.classList.toggle("grid-density-3", density === "3");
  gridDensity6Button.setAttribute("aria-pressed", String(density !== "3"));
  gridDensity3Button.setAttribute("aria-pressed", String(density === "3"));
  try {
    localStorage.setItem(GRID_DENSITY_KEY, density);
  } catch {
    // 表示設定を保存できない場合も一覧表示は継続する
  }
}

gridDensity6Button.addEventListener("click", () => applyGridDensity("6"));
gridDensity3Button.addEventListener("click", () => applyGridDensity("3"));

try {
  applyGridDensity(localStorage.getItem(GRID_DENSITY_KEY) === "3" ? "3" : "6");
} catch {
  applyGridDensity("6");
}

chooseWorkFolderButton.addEventListener("click", () => { void chooseWorkFolder(); });

exportWorkFileButton.addEventListener("click", async () => {
  if (!assets.length || !photos.length) {
    setSessionStatus("先にExcelと写真を読み込んでください。", "error");
    return;
  }
  exportWorkFileButton.disabled = true;
  setSessionStatus("別PC引き継ぎ用の作業ファイルを作成中…", "working");
  try {
    const result = await createPhotoWorkFile({
      name: sessionName.value.trim() || `写真整理_${new Date().toLocaleDateString("ja-JP")}`,
      assets,
      photos,
      healthWorkbookFile,
      albumTemplateFile,
      photoTargetKb: photoTargetKb.value,
      bookmarkPhotoId,
    });
    if (workFolderHandle && typeof workFolderHandle.getFileHandle === "function") {
      await saveWorkFileToFolder(workFolderHandle, result.blob, result.fileName);
      setSessionStatus(`${result.fileName}を「${workFolderHandle.name}」に保存しました。別PCで「作業ファイル読込」から開けます。`, "saved");
    } else {
      downloadBlob(result.blob, result.fileName);
      setSessionStatus(`${result.fileName}をダウンロードしました。保存先を指定する場合は「保存先変更」を押してください。`, "saved");
    }
  } catch (error) {
    setSessionStatus(`作業ファイルを作成できませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    exportWorkFileButton.disabled = false;
  }
});

importWorkFileButton.addEventListener("click", () => importWorkFileInput.click());
importWorkFileInput.addEventListener("change", async () => {
  const file = importWorkFileInput.files?.[0];
  importWorkFileInput.value = "";
  if (!file) return;
  importWorkFileButton.disabled = true;
  setSessionStatus(`${file.name}を読み込んでいます…`, "working");
  try {
    const saved = await readPhotoWorkFile(file);
    assets = saved.assets ?? [];
    photos = [];
    healthWorkbookFile = null;
    albumTemplateFile = null;
    pendingWorkPhotoMetadata = saved.photos ?? [];
    pendingWorkBookmarkId = saved.bookmarkPhotoId ?? null;
    renderAssets();
    renderCameraAssets();
    renderPhotos();
    photoInput.disabled = false;
    albumInput.disabled = false;
    setStatus(photoStatus, `作業ファイルを読み込みました。同じ写真を選択すると分類を復元します。`, "success");
    setSessionStatus(`「${saved.name || file.name}」を読み込みました。写真を再選択してください。`, "saved");
    sessionName.value = saved.name || "";
    await refreshSnapshotList();
  } catch (error) {
    setSessionStatus(`作業ファイルを読み込めませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    importWorkFileButton.disabled = false;
  }
});

loadSessionSnapshotButton.addEventListener("click", async () => {
  const id = savedSessionSelect.value;
  if (!id) return;
  loadSessionSnapshotButton.disabled = true;
  try {
    const saved = await loadPhotoSession(id);
    if (!saved) throw new Error("保存済み作業が見つかりません。");
    applyLoadedSession(saved, `「${saved.name || "名称未設定"}」を開きました。修正後は再度「この作業を保存」を押してください。`);
    sessionName.value = saved.name || "";
    scheduleSessionSave();
  } catch (error) {
    setSessionStatus(`保存済み作業を開けませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    await refreshSnapshotList();
  }
});

deleteSessionSnapshotButton.addEventListener("click", async () => {
  const id = savedSessionSelect.value;
  if (!id) return;
  const option = [...savedSessionSelect.options].find((item) => item.value === id);
  if (!window.confirm(`「${option?.textContent ?? "選択中の保存済み作業"}」を削除します。元に戻せません。よろしいですか？`)) return;
  deleteSessionSnapshotButton.disabled = true;
  try {
    await deletePhotoSnapshot(id);
    setSessionStatus("保存済み作業を削除しました。", "saved");
  } catch (error) {
    setSessionStatus(`保存済み作業を削除できませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    await refreshSnapshotList();
  }
});

void restoreSession();
