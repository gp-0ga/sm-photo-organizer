import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { createAssetFolders, exportOrganizedPhotos } from "./folders.js";
import { OTHER_ASSET, OTHER_ASSET_NUMBER, photoDestinations } from "./domain.js";
import { analyzePhotoFiles, compressToJpeg, jpegFileName, targetBytesFromKilobytes } from "./photos.js";
import { albumEntries, createPhotoAlbum, inspectPhotoAlbumTemplate } from "./album.js";
import { cameraFileName, captureVideoFrame, openRearCamera, stopCamera } from "./camera.js";
import { createAssetPhotoZip, createZip, zipFileName } from "./zip.js";
import { deletePhotoSnapshot, listPhotoSnapshots, loadPhotoSession, savePhotoSession, savePhotoSnapshot } from "./photo-session-storage.js";
import { createPhotoWorkFile, readPhotoWorkFile } from "./photo-work-file.js";
import photoBookBat from "../scripts/写真帳作成.bat?raw";
import photoBookRunner from "../scripts/photo-book-runner.ps1?raw";
import photoBookBuilder from "../scripts/build-photo-album.ps1?raw";

const hostedOrganizer = /\/organize(?:\.html)?$/i.test(window.location.pathname);
if (hostedOrganizer) document.body.classList.add("hosted-organizer");
const isLocalPhotoBookApp = ["127.0.0.1", "localhost", "[::1]"].includes(window.location.hostname);

const excelInput = document.querySelector("#excel-input");
const excelFileName = document.querySelector("#excel-file-name");
function setSelectedFileName(element, text) {
  element.hidden = !text;
  element.textContent = text || "";
}
const excelStatus = document.querySelector("#excel-status");
const folderStatus = document.querySelector("#folder-status");
const createFoldersButton = document.querySelector("#create-folders");
const addAssetButton = document.querySelector("#add-asset");
const assetBulkTools = document.querySelector("#asset-bulk-tools");
const selectAllAssets = document.querySelector("#select-all-assets");
const assetBulkStatus = document.querySelector("#asset-bulk-status");
const assetBulkSiteAbsent = document.querySelector("#asset-bulk-site-absent");
const assetBulkSitePresent = document.querySelector("#asset-bulk-site-present");
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
const photoSortFilter = document.querySelector("#photo-sort-filter");
const photoFilterKeyword = document.querySelector("#photo-filter-keyword");
const photoFilterStatus = document.querySelector("#photo-filter-status");
const photoTemplate = document.querySelector("#photo-template");
const bulkTools = document.querySelector("#bulk-tools");
const photoToolsEmpty = document.querySelector("#photo-tools-empty");
const photoToolsEmptyCopy = document.querySelector("#photo-tools-empty-copy");
const bulkToolsSentinel = document.querySelector("#bulk-tools-sentinel");
const selectAllPhotos = document.querySelector("#select-all-photos");
const bulkSelectionStatus = document.querySelector("#bulk-selection-status");
const bulkAsset = document.querySelector("#bulk-asset");
const applyBulkAsset = document.querySelector("#apply-bulk-asset");
const clearReview = document.querySelector("#clear-review");
const clearAllReview = document.querySelector("#clear-all-review");
const bulkExcludeToggle = document.querySelector("#bulk-exclude-toggle");
const bulkExcludeToggleLabel = document.querySelector("#bulk-exclude-toggle-label");
const exportPhotosButton = document.querySelector("#export-photos");
const exportStatus = document.querySelector("#export-status");
const exportToolbarStatus = document.querySelector("#export-toolbar-status");
const photoActionStatus = document.querySelector("#photo-action-status");
const photoActionStatusCopy = document.querySelector("#photo-action-status-copy");
const undoPhotoActionButton = document.querySelector("#undo-photo-action");
const photoTargetKb = document.querySelector("#photo-target-kb");
const albumInput = document.querySelector("#album-input");
const albumTemplateStatus = document.querySelector("#album-template-status");
const albumStatus = document.querySelector("#album-status");
const createAlbumButton = document.querySelector("#create-album");
const exportAlbumKitButton = document.querySelector("#export-album-kit");
const preflightDialog = document.querySelector("#preflight-dialog");
const preflightList = document.querySelector("#preflight-list");
const outputPreviewDialog = document.querySelector("#output-preview-dialog");
const outputPreviewList = document.querySelector("#output-preview-list");
const outputPreviewTitle = document.querySelector("#output-preview-title");
const outputPreviewDescription = document.querySelector("#output-preview-description");
const outputPreviewConfirm = document.querySelector("#output-preview-confirm");
const viewSelectionSummaryButton = document.querySelector("#view-selection-summary");
const showAssetsTabButton = document.querySelector("#show-assets-tab");
const workspaceTabs = [...document.querySelectorAll("[data-workspace-tab]")];
const workspacePanels = [...document.querySelectorAll("[data-workspace-panel]")];
const workspacePanes = [...document.querySelectorAll(".workspace-pane[data-pane]")];
const paneToggleButtons = [...document.querySelectorAll("[data-pane-toggle]")];
const assetEditorDialog = document.querySelector("#asset-editor-dialog");
const assetEditorForm = document.querySelector("#asset-editor-form");
const assetEditorHeading = document.querySelector("#asset-editor-heading");
const assetEditorNumber = document.querySelector("#asset-editor-number");
const assetEditorName = document.querySelector("#asset-editor-name");
const assetEditorError = document.querySelector("#asset-editor-error");
const workspaceToolsPane = document.querySelector('.workspace-pane[data-pane="tools"]');
const workspaceToolsRail = document.querySelector('.pane-toggle-rail[data-pane-toggle="tools"]');
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

function photoBookLocalOnlyMessage() {
  return "写真帳作成は、ExcelをPC内で操作するためローカル版で実行します。作業用フォルダの start-mvp.bat を開き、http://127.0.0.1:8765/ で作業してください。";
}

let assets = [];
let photos = [];
let markerCount = 0;
let bookmarkPhotoId = null;
let healthWorkbookFile = null;
let albumTemplateFile = null;
let pendingWorkPhotoMetadata = null;
let pendingWorkBookmarkId = null;
const thumbnailUrls = new Set();
const outputPreviewUrls = new Set();
let cameraStream = null;
let captureInProgress = false;
let saveTimer = null;
let saveInProgress = false;
let saveQueued = false;
const selectedPhotoIdsState = new Set();
const selectedAssetNumbers = new Set();
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

let undoPhotoAction = null;

function showPhotoActionStatus(message, undo = null) {
  undoPhotoAction = undo;
  photoActionStatusCopy.textContent = message;
  undoPhotoActionButton.hidden = typeof undo !== "function";
  photoActionStatus.hidden = false;
}

function hidePhotoActionStatus() {
  undoPhotoAction = null;
  photoActionStatus.hidden = true;
}

undoPhotoActionButton.addEventListener("click", () => {
  const undo = undoPhotoAction;
  if (!undo) return;
  undoPhotoAction = null;
  undo();
  photoActionStatusCopy.textContent = "直前の操作を元に戻しました。";
  undoPhotoActionButton.hidden = true;
});

function openAssetEditor({ title, assetNumber = "", assetName = "", originalNumber = "" }) {
  assetEditorHeading.textContent = title;
  assetEditorNumber.value = assetNumber;
  assetEditorName.value = assetName;
  assetEditorError.hidden = true;
  assetEditorError.textContent = "";
  assetEditorDialog.showModal();
  assetEditorNumber.focus();

  return new Promise((resolve) => {
    const cancelButton = assetEditorForm.querySelector('button[value="cancel"]');
    const onCancel = () => assetEditorDialog.close("cancel");
    const onSubmit = (event) => {
      event.preventDefault();
      const nextNumber = assetEditorNumber.value.trim();
      const nextName = assetEditorName.value.trim();
      if (!nextNumber || !nextName) {
        assetEditorError.textContent = "資産番号と資産名を入力してください。";
        assetEditorError.hidden = false;
        return;
      }
      if (assets.some((asset) => asset.assetNumber === nextNumber && asset.assetNumber !== originalNumber)) {
        assetEditorError.textContent = "同じ資産番号がすでにあります。";
        assetEditorError.hidden = false;
        assetEditorNumber.focus();
        return;
      }
      assetEditorDialog.close("save");
    };
    const onClose = () => {
      assetEditorForm.removeEventListener("submit", onSubmit);
      cancelButton.removeEventListener("click", onCancel);
      assetEditorDialog.removeEventListener("close", onClose);
      resolve(assetEditorDialog.returnValue === "save"
        ? { assetNumber: assetEditorNumber.value.trim(), assetName: assetEditorName.value.trim() }
        : null);
    };
    assetEditorForm.addEventListener("submit", onSubmit);
    cancelButton.addEventListener("click", onCancel);
    assetEditorDialog.addEventListener("close", onClose);
  });
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
    node.querySelector(".asset-select-label").addEventListener("click", (event) => event.stopPropagation());
    node.querySelector(".asset-site-status").addEventListener("click", (event) => event.stopPropagation());
    const assetSelect = node.querySelector(".asset-select");
    assetSelect.checked = selectedAssetNumbers.has(asset.assetNumber);
    assetSelect.addEventListener("change", () => {
      if (assetSelect.checked) selectedAssetNumbers.add(asset.assetNumber);
      else selectedAssetNumbers.delete(asset.assetNumber);
      updateAssetBulkControls();
    });
    const siteAbsent = node.querySelector(".asset-site-absent");
    siteAbsent.checked = Boolean(asset.siteAbsent);
    siteAbsent.addEventListener("change", () => {
      asset.siteAbsent = siteAbsent.checked;
      warning.hidden = !(asset.notInExcel || asset.siteAbsent);
      warning.textContent = [asset.notInExcel && "Excelに計上なし", asset.siteAbsent && "現地なし／対象外"].filter(Boolean).join("・");
      updatePhotoSummary();
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
  updatePhotoToolsVisibility();
  renderPhotoFilterOptions();
  const currentAssetNumbers = new Set(assets.map((asset) => asset.assetNumber));
  for (const assetNumber of [...selectedAssetNumbers]) {
    if (!currentAssetNumbers.has(assetNumber)) selectedAssetNumbers.delete(assetNumber);
  }
  assetBulkTools.hidden = assets.length === 0;
  updateAssetBulkControls();
}

function updateAssetBulkControls() {
  const total = assets.length;
  const selectedCount = selectedAssetNumbers.size;
  assetBulkStatus.textContent = `${selectedCount}件選択中`;
  selectAllAssets.indeterminate = selectedCount > 0 && selectedCount < total;
  selectAllAssets.checked = total > 0 && selectedCount === total;
  assetBulkSiteAbsent.disabled = selectedCount === 0;
  assetBulkSitePresent.disabled = selectedCount === 0;
}

function setSiteAbsentForSelection(siteAbsent) {
  if (!selectedAssetNumbers.size) return;
  for (const asset of assets) {
    if (!selectedAssetNumbers.has(asset.assetNumber)) continue;
    asset.siteAbsent = siteAbsent;
  }
  renderAssets();
  updatePhotoSummary();
  scheduleSessionSave();
}

selectAllAssets.addEventListener("change", () => {
  selectedAssetNumbers.clear();
  if (selectAllAssets.checked) {
    for (const asset of assets) selectedAssetNumbers.add(asset.assetNumber);
  }
  for (const card of assetList.querySelectorAll(".asset-card")) {
    card.querySelector(".asset-select").checked = selectedAssetNumbers.has(card.querySelector(".asset-number").textContent);
  }
  updateAssetBulkControls();
});

assetBulkSiteAbsent.addEventListener("click", () => setSiteAbsentForSelection(true));
assetBulkSitePresent.addEventListener("click", () => setSiteAbsentForSelection(false));

const UNCLASSIFIED_FILTER_VALUE = "filter:unclassified";
const NEEDS_REVIEW_FILTER_VALUE = "filter:needs";
const ORDER_CAPTURE_VALUE = "order:capture";
const ORDER_ASSET_VALUE = "order:asset";

function currentSortFilterValue() {
  return photoSortFilter.value || ORDER_CAPTURE_VALUE;
}

function currentOrderMode() {
  const value = currentSortFilterValue();
  return value === ORDER_ASSET_VALUE ? "asset" : "capture";
}

function currentFilterValue() {
  const value = currentSortFilterValue();
  return value.startsWith("filter:") ? value : "";
}

function renderPhotoFilterOptions() {
  const previousValue = photoSortFilter.value;
  const orderOptions = [
    { value: ORDER_CAPTURE_VALUE, label: "撮影順（すべて表示）" },
    { value: ORDER_ASSET_VALUE, label: "資産ごとに並べる（すべて表示）" },
  ];
  const filterOptions = [
    { value: UNCLASSIFIED_FILTER_VALUE, label: "未分類のみ" },
    { value: NEEDS_REVIEW_FILTER_VALUE, label: "要確認・読込失敗のみ" },
    ...assets.map((asset) => ({ value: `filter:${asset.assetNumber}`, label: `${asset.assetNumber} ${asset.assetName} のみ` })),
    { value: `filter:${OTHER_ASSET_NUMBER}`, label: "その他（資産不明）のみ" },
  ];
  photoSortFilter.innerHTML = `
    <optgroup label="並び替え（全件表示）">${orderOptions.map(({ value, label }) => `<option value="${value}">${label}</option>`).join("")}</optgroup>
    <optgroup label="絞り込み">${filterOptions.map(({ value, label }) => `<option value="${value}">${label}</option>`).join("")}</optgroup>
  `;
  const allValues = [...orderOptions, ...filterOptions].map((option) => option.value);
  photoSortFilter.value = allValues.includes(previousValue) ? previousValue : ORDER_CAPTURE_VALUE;
}

function updatePhotoToolsVisibility() {
  const hasPhotos = photos.length > 0;
  bulkTools.hidden = !hasPhotos;
  if (photoToolsEmpty) photoToolsEmpty.hidden = hasPhotos;
  if (photoToolsEmptyCopy && !hasPhotos) {
    photoToolsEmptyCopy.textContent = assets.length
      ? "撮影写真を読み込むと、絞り込み・一括変更・確認済み操作が表示されます。"
      : "健全度判定表と撮影写真を読み込むと、絞り込み・一括変更・確認済み操作が表示されます。";
  }
}

function renderPhotoEmptyState() {
  const empty = document.createElement("div");
  empty.className = "photo-empty-state";
  empty.innerHTML = assets.length
    ? "<strong>撮影写真を読み込んでください</strong><span>写真を選択すると、QRマーカーをもとに資産ごとの整理を始められます。</span>"
    : "<strong>写真整理はここに表示されます</strong><span>左の「読み込み」から健全度判定表、次に撮影写真を選択してください。</span>";
  photoList.replaceChildren(empty);
}

function matchesPhotoFilter(photo) {
  const filterValue = currentFilterValue();
  if (filterValue === UNCLASSIFIED_FILTER_VALUE) {
    if (photo.assetNumber) return false;
  } else if (filterValue === NEEDS_REVIEW_FILTER_VALUE) {
    if (!photo.reviewRequired && !photo.qrReadError) return false;
  } else if (filterValue) {
    const assetNumber = filterValue.slice("filter:".length);
    if (photo.assetNumber !== assetNumber) return false;
  }
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
  const hasFilter = currentFilterValue() || photoFilterKeyword?.value.trim();
  photoFilterStatus.textContent = hasFilter ? `${matchingCount}枚を表示中（全${photos.length}枚）` : "";
  const visibleReviewCount = photos.filter((photo) => matchesPhotoFilter(photo) && (photo.reviewRequired || photo.qrReadError)).length;
  clearAllReview.hidden = visibleReviewCount === 0;
  clearAllReview.textContent = `表示中の要確認をまとめて解除（${visibleReviewCount}枚）`;
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
  return `<label class="photo-item-choice">サブフォルダー<select class="photo-item-select">${options.map(({ value, label }) => `<option value="${value}"${selected.has(value) ? " selected" : ""}>${label}</option>`).join("")}</select></label>`;
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
    ...fullIssues.slice(0, 2).map((issue) => `${issue.title}（${issue.detail}）`),
    ...otherIssues.slice(0, 1).map((issue) => `${issue.title}（${issue.detail}）`),
  ];
  const remaining = issues.length - labels.length;
  photoDestinationAlert.hidden = false;
  photoDestinationAlert.className = "photo-destination-alert";
  photoDestinationAlert.textContent = `出力前に確認：${labels.join("、")}${remaining > 0 ? `、ほか${remaining}件。下の「選択状況一覧」で確認できます。` : ""}`;
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

function clearOutputPreviewUrls() {
  for (const url of outputPreviewUrls) URL.revokeObjectURL(url);
  outputPreviewUrls.clear();
}

function openOutputPreview({ forExport = true } = {}) {
  if (!outputPreviewDialog || !outputPreviewList || typeof outputPreviewDialog.showModal !== "function") return Promise.resolve(true);
  if (outputPreviewTitle) outputPreviewTitle.textContent = forExport ? "出力内容の確認" : "選択状況一覧";
  if (outputPreviewDescription) outputPreviewDescription.textContent = forExport
    ? "資産ごとに、全景・サブフォルダーへ入る写真を確認してください。"
    : "健全度判定表の点検結果1・2を入力するときも、この一覧を確認できます。";
  if (outputPreviewConfirm) outputPreviewConfirm.hidden = !forExport;
  clearOutputPreviewUrls();
  outputPreviewList.replaceChildren();
  const groups = [];
  for (const asset of assets) {
    if (asset.siteAbsent) continue;
    const destinations = ["全景", ...asset.items.map((item) => item.folderName)];
    for (const destination of destinations) {
      const selected = photos.filter((photo) => {
        if (photo.excluded || photo.assetNumber !== asset.assetNumber) return false;
        return photoDestinations(photo).includes(destination);
      });
      if (selected.length) groups.push({ asset, destination, photos: selected });
    }
  }
  for (const group of groups) {
    const section = document.createElement("section");
    section.className = "output-preview-group";
    const heading = document.createElement("h3");
    heading.innerHTML = `${group.asset.assetNumber} ${group.asset.assetName} / ${group.destination}<span>${group.photos.length}枚</span>`;
    section.append(heading);
    const grid = document.createElement("div");
    grid.className = "output-preview-grid";
    for (const photo of group.photos) {
      const figure = document.createElement("figure");
      figure.className = "output-preview-item";
      const image = document.createElement("img");
      const url = URL.createObjectURL(photo.file);
      outputPreviewUrls.add(url);
      image.src = url;
      image.alt = photo.file.name;
      image.loading = "lazy";
      const caption = document.createElement("figcaption");
      caption.textContent = photo.file.name;
      figure.append(image, caption);
      grid.append(figure);
    }
    section.append(grid);
    outputPreviewList.append(section);
  }
  outputPreviewDialog.showModal();
  return new Promise((resolve) => {
    const finish = () => {
      outputPreviewDialog.removeEventListener("close", finish);
      const confirmed = outputPreviewDialog.returnValue === "confirm";
      clearOutputPreviewUrls();
      resolve(confirmed);
    };
    outputPreviewDialog.addEventListener("close", finish);
  });
}

viewSelectionSummaryButton?.addEventListener("click", () => {
  void openOutputPreview({ forExport: false });
});

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
  exportAlbumKitButton.disabled = true;
  let entries;
  try {
    entries = albumEntries(assets, photos);
    exportAlbumKitButton.disabled = false;
  } catch (error) {
    const message = error.message ?? String(error);
    const tone = message.includes("選ばれていません") ? "neutral" : "error";
    setStatus(albumStatus, message, tone);
    return;
  }
  if (!isLocalPhotoBookApp) {
    setStatus(albumStatus, `${entries.length}枚を写真帳へ貼り付ける指示を保存できます。`, "success");
    return;
  }
  if (!healthWorkbookFile) {
    setStatus(albumStatus, "先に健全度判定表を読み込んでください。", "neutral");
    return;
  }
  if (!albumTemplateFile) {
    setStatus(albumStatus, "写真帳様式を選択してください。", "neutral");
    return;
  }
  try {
    createAlbumButton.disabled = false;
    setStatus(albumStatus, `${entries.length}枚を写真帳へ貼り付けます。`, "success");
  } catch (error) {
    const message = error.message ?? String(error);
    const tone = message.includes("選ばれていません") ? "neutral" : "error";
    setStatus(albumStatus, message, tone);
  }
}

function photoBookKitFileName() {
  // ZIPをWindows標準の展開機能で開いても文字化けしないよう、配布ファイル名はASCIIに統一する。
  return "photo-book-kit.zip";
}

async function createPhotoBookKit() {
  const entries = albumEntries(assets, photos);
  const instruction = {
    kind: "asset-photo-album-instruction",
    version: 1,
    exportedAt: new Date().toISOString(),
    assets: assets.map((asset) => ({
      assetNumber: asset.assetNumber,
      itemNumbers: asset.items.map((item) => item.itemNumber),
    })),
    photos: entries.map((entry) => ({
      fileName: entry.photo.file.name,
      assetNumber: entry.assetNumber,
      destinationType: entry.destinationType,
      itemNumber: entry.itemNumber,
      slotIndex: entry.slotIndex,
    })),
  };
  const guide = [
    "資産写真整理MVP 写真帳作成セット",
    "",
    "1. ZIPを展開します。",
    "2. 「START-PHOTO-BOOK.bat」をダブルクリックします。",
    "3. 写真帳Excel、元写真フォルダ、出力先フォルダを順に選びます。",
    "4. 元の写真帳Excelは変更せず、出力先に写真だけを貼り付けたコピーを作成します。",
    "",
    "注意：このセットでは点検結果1・2を転記しません。Microsoft ExcelがインストールされたWindows PCで実行してください。",
    "元写真はJPG、JPEG、PNGに対応します。同名写真が複数ある場合は整理してから実行してください。",
  ].join("\r\n");
  const blob = await createZip([
    { path: "START-PHOTO-BOOK.bat", blob: new Blob([photoBookBat], { type: "text/plain;charset=us-ascii" }) },
    { path: "photo-book-runner.ps1", blob: new Blob([photoBookRunner], { type: "text/plain;charset=utf-8" }) },
    { path: "build-photo-album.ps1", blob: new Blob([photoBookBuilder], { type: "text/plain;charset=utf-8" }) },
    { path: "photo-book-instruction.json", blob: new Blob([JSON.stringify(instruction, null, 2)], { type: "application/json" }) },
    // UTF-8 BOMを付け、Windowsのメモ帳でも日本語の案内文を正しく開けるようにする。
    { path: "README.txt", blob: new Blob(["\uFEFF", guide], { type: "text/plain;charset=utf-8" }) },
  ]);
  return { blob, entryCount: entries.length };
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
    setSelectedFileName(excelFileName, healthWorkbookFile?.name || "");
    setSelectedFileName(photoFileName, "");
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
  const destinationOptionsPanel = node.querySelector(".photo-destination-options");
  const fullBadge = node.querySelector(".photo-full-badge");
  const renderDestinationChooser = () => {
    const destinations = photoDestinations(photo);
    destinationOptionsPanel.innerHTML = destinationChoices(photo);
    const hasAsset = Boolean(assets.find((item) => item.assetNumber === photo.assetNumber));
    const isFull = destinations.includes("全景");
    fullBadge.hidden = !hasAsset;
    fullBadge.classList.toggle("active", isFull);
    fullBadge.setAttribute("aria-pressed", String(isFull));
    fullBadge.title = isFull ? "全景を解除" : "全景として使う";
  };
  renderDestinationChooser();
  node.querySelector(".photo-exclude").checked = photo.excluded;
  const reviewBadge = node.querySelector(".review-badge");
  reviewBadge.hidden = !(photo.reviewRequired || photo.qrReadError);
  reviewBadge.addEventListener("click", () => {
    clearPhotoReviews([photo]);
  });
  const bookmarkButton = node.querySelector(".photo-bookmark");
  const isBookmarked = photo.id === bookmarkPhotoId;
  card.classList.toggle("bookmarked", isBookmarked);
  bookmarkButton.title = isBookmarked ? "ここまで確認済み（解除）" : "ここまで確認";
  bookmarkButton.setAttribute("aria-label", isBookmarked ? "しおりを解除" : "ここまで確認済みにする");
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
        previousButton.setAttribute("aria-label", "ここまで確認済みにする");
        previousButton.setAttribute("aria-pressed", "false");
      }
    }
    const marked = bookmarkPhotoId === photo.id;
    card.classList.toggle("bookmarked", marked);
    bookmarkButton.title = marked ? "ここまで確認済み（解除）" : "ここまで確認";
    bookmarkButton.setAttribute("aria-label", marked ? "しおりを解除" : "ここまで確認済みにする");
    bookmarkButton.setAttribute("aria-pressed", String(marked));
    jumpToBookmarkButton.disabled = !bookmarkPhotoId;
    scheduleSessionSave();
  });

  assetSelect.addEventListener("change", () => {
    const previousAssetNumber = photo.assetNumber;
    const previousDestinations = photoDestinations(photo);
    const nextAssetNumber = assetSelect.value || null;
    if (nextAssetNumber === previousAssetNumber) return;
    photo.assetNumber = nextAssetNumber;
    setPhotoDestinations(photo, []);
    renderDestinationChooser();
    updatePhotoSummary();
    scheduleSessionSave();
    applyPhotoFilter();
    const cleared = previousDestinations.filter(Boolean).length > 0;
    showPhotoActionStatus(
      `${photo.file.name}の資産を変更しました。${cleared ? "写真帳分類は解除されました。" : ""}`,
      () => {
        photo.assetNumber = previousAssetNumber;
        setPhotoDestinations(photo, previousDestinations);
        renderPhotos();
        scheduleSessionSave();
      },
    );
  });
  destinationOptionsPanel.addEventListener("change", (event) => {
    if (!event.target.matches(".photo-item-select")) return;
    const full = photoDestinations(photo).includes("全景");
    const item = event.target.value || "";
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
  fullBadge.addEventListener("click", () => {
    const currentlyFull = photoDestinations(photo).includes("全景");
    if (!currentlyFull && destinationCount(photo, "全景") >= destinationLimit("全景") && !photo.excluded) {
      setStatus(albumStatus, "全景は最大1枚です。", "error");
      return;
    }
    const item = destinationOptionsPanel.querySelector(".photo-item-select")?.value || "";
    setPhotoDestinations(photo, [currentlyFull ? "" : "全景", item]);
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
  const hasActiveFilter = Boolean(currentFilterValue() || photoFilterKeyword?.value.trim());
  const chunkSize = photoRenderObserver && !hasActiveFilter ? PHOTO_RENDER_CHUNK_SIZE : photoRenderQueue.length;
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
  if (!photos.length) {
    renderPhotoEmptyState();
    updatePhotoToolsVisibility();
    updatePhotoSummary();
    updateBulkControls();
    jumpToBookmarkButton.disabled = true;
    return;
  }
  let displayPhotos = photos.slice();
  if (currentOrderMode() === "asset") {
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
    if (currentOrderMode() === "asset") {
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
  updatePhotoToolsVisibility();
  bulkAsset.innerHTML = assetOptions();
  selectAllPhotos.checked = false;
  selectAllPhotos.indeterminate = false;
  updatePhotoSummary();
  updateBulkControls();
  jumpToBookmarkButton.disabled = !bookmarkPhotoId;
  applyPhotoFilter();
}

photoSortFilter.addEventListener("change", renderPhotos);
photoFilterKeyword?.addEventListener("input", renderPhotos);

function selectWorkspaceTab(name) {
  for (const tab of workspaceTabs) {
    const active = tab.dataset.workspaceTab === name;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  }
  for (const panel of workspacePanels) panel.hidden = panel.dataset.workspacePanel !== name;
  syncToolsPaneForWorkspace(name);
}

for (const tab of workspaceTabs) tab.addEventListener("click", () => selectWorkspaceTab(tab.dataset.workspaceTab));
if (showAssetsTabButton) showAssetsTabButton.addEventListener("click", () => selectWorkspaceTab("assets"));

const PANE_COLLAPSE_KEY_PREFIX = "workspace-pane-collapsed-";

function syncToolsPaneForWorkspace(workspaceName) {
  if (!workspaceToolsPane || !workspaceToolsRail) return;
  const isPhotoWorkspace = workspaceName === "photos";
  const isCollapsed = workspaceToolsPane.classList.contains("pane-collapsed");
  workspaceToolsPane.hidden = !isPhotoWorkspace;
  workspaceToolsRail.hidden = !isPhotoWorkspace || !isCollapsed;
  workspaceToolsRail.classList.toggle("visible", isPhotoWorkspace && isCollapsed);
}

function setPaneCollapsed(paneName, collapsed) {
  const pane = workspacePanes.find((element) => element.dataset.pane === paneName);
  if (pane) {
    pane.classList.toggle("pane-collapsed", collapsed);
    pane.setAttribute("aria-hidden", String(collapsed));
  }
  for (const button of paneToggleButtons) {
    if (button.dataset.paneToggle !== paneName) continue;
    if (button.classList.contains("pane-toggle-rail")) {
      button.hidden = !collapsed;
      button.classList.toggle("visible", collapsed);
    } else {
      button.setAttribute("aria-expanded", String(!collapsed));
    }
  }
  try {
    localStorage.setItem(`${PANE_COLLAPSE_KEY_PREFIX}${paneName}`, collapsed ? "1" : "0");
  } catch {
    // 開閉状態を保存できない場合も表示は継続する
  }
  if (paneName === "tools") {
    const activeWorkspace = workspaceTabs.find((tab) => tab.classList.contains("active"))?.dataset.workspaceTab;
    syncToolsPaneForWorkspace(activeWorkspace);
  }
}

for (const button of paneToggleButtons) {
  const paneName = button.dataset.paneToggle;
  const collapseOnClick = !button.classList.contains("pane-toggle-rail");
  button.addEventListener("click", () => setPaneCollapsed(paneName, collapseOnClick));
}

for (const pane of workspacePanes) {
  let restored = false;
  try {
    restored = localStorage.getItem(`${PANE_COLLAPSE_KEY_PREFIX}${pane.dataset.pane}`) === "1";
  } catch {
    restored = false;
  }
  setPaneCollapsed(pane.dataset.pane, restored);
}
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

function visiblePhotoCards() {
  return [...document.querySelectorAll(".photo-card:not([hidden])")];
}

function selectedPhotoIds() {
  const visibleIds = new Set(visiblePhotoCards().map((card) => card.dataset.photoId));
  return new Set([...selectedPhotoIdsState].filter((id) => visibleIds.has(id)));
}

function updateBulkControls() {
  const visibleIds = photos.filter(matchesPhotoFilter).map((photo) => photo.id);
  const visibleIdSet = new Set(visibleIds);
  const selectedCount = visibleIds.filter((id) => selectedPhotoIdsState.has(id)).length;
  const selectedPhotos = photos.filter((photo) => visibleIdSet.has(photo.id) && selectedPhotoIdsState.has(photo.id));
  const total = visibleIds.length;
  bulkSelectionStatus.textContent = `${selectedCount}枚選択中`;
  selectAllPhotos.indeterminate = selectedCount > 0 && selectedCount < total;
  selectAllPhotos.checked = total > 0 && selectedCount === total;
  bulkAsset.disabled = selectedCount === 0;
  applyBulkAsset.disabled = selectedCount === 0 || !bulkAsset.value;
  clearReview.disabled = !selectedPhotos.some((photo) => photo.reviewRequired || photo.qrReadError);
  bulkExcludeToggle.disabled = selectedCount === 0;
  if (selectedCount > 0) {
    const excludedCount = selectedPhotos.filter((photo) => photo.excluded).length;
    bulkExcludeToggle.indeterminate = excludedCount > 0 && excludedCount < selectedPhotos.length;
    bulkExcludeToggle.checked = excludedCount === selectedPhotos.length;
  } else {
    bulkExcludeToggle.indeterminate = false;
    bulkExcludeToggle.checked = false;
  }
  bulkExcludeToggleLabel.textContent = bulkExcludeToggle.checked ? "選択写真の除外を解除" : "選択写真を除外にする";
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  setSelectedFileName(excelFileName, file.name);
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

addAssetButton.addEventListener("click", async () => {
  const result = await openAssetEditor({ title: "資産を追加" });
  if (!result) return;
  const { assetNumber, assetName } = result;
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

assetList.addEventListener("click", async (event) => {
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
    const result = await openAssetEditor({
      title: "資産の番号・名称を変更",
      assetNumber: asset.assetNumber,
      assetName: asset.assetName,
      originalNumber: asset.assetNumber,
    });
    if (!result) return;
    const { assetNumber: nextNumber, assetName: nextName } = result;
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
    updatePhotoToolsVisibility();
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
  setSelectedFileName(photoFileName, `${files.length}枚を選択中`);
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
    setSelectedFileName(photoFileName, "");
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

function cardsByPhotoId() {
  const map = new Map();
  for (const card of document.querySelectorAll(".photo-card")) map.set(card.dataset.photoId, card);
  return map;
}

function clearPhotoReviews(targetPhotos) {
  const affected = targetPhotos.filter((photo) => photo.reviewRequired || photo.qrReadError);
  if (!affected.length) return;
  const previous = affected.map((photo) => ({
    id: photo.id,
    reviewRequired: photo.reviewRequired,
    qrReadError: photo.qrReadError,
  }));
  for (const photo of affected) {
    photo.reviewRequired = false;
    photo.qrReadError = false;
  }
  renderPhotos();
  scheduleSessionSave();
  showPhotoActionStatus(`${affected.length}枚を確認済みにしました。`, () => {
    for (const saved of previous) {
      const photo = photos.find((item) => item.id === saved.id);
      if (!photo) continue;
      photo.reviewRequired = saved.reviewRequired;
      photo.qrReadError = saved.qrReadError;
    }
    renderPhotos();
    scheduleSessionSave();
  });
}

applyBulkAsset.addEventListener("click", () => {
  const ids = selectedPhotoIds();
  if (!ids.size || !bulkAsset.value) return;
  const previous = photos
    .filter((photo) => ids.has(photo.id) && photo.assetNumber !== bulkAsset.value)
    .map((photo) => ({ id: photo.id, assetNumber: photo.assetNumber, destinations: photoDestinations(photo) }));
  if (!previous.length) {
    showPhotoActionStatus("選択写真には、すでに同じ資産が設定されています。");
    return;
  }
  const cards = cardsByPhotoId();
  for (const photo of photos) {
    if (ids.has(photo.id) && photo.assetNumber !== bulkAsset.value) {
      photo.assetNumber = bulkAsset.value;
      setPhotoDestinations(photo, []);
      const card = cards.get(photo.id);
      if (!card) continue;
      card.querySelector(".photo-asset").value = photo.assetNumber;
      card.querySelector(".photo-destination-options").innerHTML = destinationChoices(photo);
      const fullBadge = card.querySelector(".photo-full-badge");
      fullBadge.hidden = !assets.find((item) => item.assetNumber === photo.assetNumber);
      fullBadge.classList.remove("active");
      fullBadge.setAttribute("aria-pressed", "false");
      fullBadge.title = "全景として使う";
    }
  }
  updatePhotoSummary();
  updateBulkControls();
  scheduleSessionSave();
  applyPhotoFilter();
  const cleared = previous.some((item) => item.destinations.filter(Boolean).length > 0);
  showPhotoActionStatus(
    `${previous.length}枚の資産を変更しました。${cleared ? "写真帳分類は解除されました。" : ""}`,
    () => {
      for (const saved of previous) {
        const photo = photos.find((item) => item.id === saved.id);
        if (!photo) continue;
        photo.assetNumber = saved.assetNumber;
        setPhotoDestinations(photo, saved.destinations);
      }
      renderPhotos();
      scheduleSessionSave();
    },
  );
});

clearReview.addEventListener("click", () => {
  const ids = selectedPhotoIds();
  clearPhotoReviews(photos.filter((photo) => ids.has(photo.id)));
});

clearAllReview.addEventListener("click", () => {
  clearPhotoReviews(photos.filter(matchesPhotoFilter));
});

function setExcludedForSelection(excluded) {
  const ids = selectedPhotoIds();
  if (!ids.size) return;
  const cards = cardsByPhotoId();
  for (const photo of photos) {
    if (!ids.has(photo.id)) continue;
    photo.excluded = excluded;
    const card = cards.get(photo.id);
    if (!card) continue;
    card.querySelector(".photo-exclude").checked = excluded;
    card.classList.toggle("excluded", excluded);
  }
  updatePhotoSummary();
  updateBulkControls();
  scheduleSessionSave();
}

bulkExcludeToggle.addEventListener("change", () => setExcludedForSelection(bulkExcludeToggle.checked));

exportPhotosButton.addEventListener("click", async () => {
  exportPhotosButton.disabled = true;
  try {
    if (showDestinationIssues()) return;
    const confirmed = await openOutputPreview();
    if (!confirmed) return;
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

exportAlbumKitButton.addEventListener("click", async () => {
  exportAlbumKitButton.disabled = true;
  try {
    if (showDestinationIssues()) return;
    setStatus(albumStatus, "写真帳作成セットを準備しています…", "working");
    const result = await createPhotoBookKit();
    downloadBlob(result.blob, photoBookKitFileName());
    setStatus(albumStatus, `写真帳作成セットを保存しました。${result.entryCount}枚を貼り付けます。`, "success");
  } catch (error) {
    setStatus(albumStatus, error.message ?? String(error), "error");
  } finally {
    updateAlbumReadiness();
  }
});

createAlbumButton.addEventListener("click", async () => {
  if (!isLocalPhotoBookApp) {
    setStatus(albumStatus, photoBookLocalOnlyMessage(), "error");
    return;
  }
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
    const deletedIndex = savedSessionSelect.selectedIndex;
    savedSessionSelect.querySelector(`option[value="${CSS.escape(id)}"]`)?.remove();
    if (savedSessionSelect.options.length <= 1) {
      savedSessionSelect.options[0].textContent = "保存済み作業はありません";
      loadSessionSnapshotButton.disabled = true;
      deleteSessionSnapshotButton.disabled = true;
    } else {
      const nextIndex = Math.min(Math.max(deletedIndex, 1), savedSessionSelect.options.length - 1);
      savedSessionSelect.selectedIndex = nextIndex;
      loadSessionSnapshotButton.disabled = false;
      deleteSessionSnapshotButton.disabled = false;
    }
    setSessionStatus("保存済み作業を削除しました。", "saved");
  } catch (error) {
    setSessionStatus(`保存済み作業を削除できませんでした：${error.message ?? String(error)}`, "error");
  } finally {
    void refreshSnapshotList();
  }
});

if (!isLocalPhotoBookApp) {
  createAlbumButton.title = "写真帳作成はPC用ローカル版で実行します";
  setStatus(albumStatus, "写真帳作成はPC用ローカル版で実行します。", "neutral");
}

void restoreSession();
