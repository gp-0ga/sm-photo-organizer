import * as XLSX from "xlsx";
import "./drawing.css";

const pageKinds = ["平面図", "立面図", "屋上平面図", "断面図", "矩計図", "面積表", "使わない"];
const kindOptions = {
  area: ["外壁面積", "防水面積"],
  opening: ["開口控除"],
  line: ["縦目地", "横目地", "打継シール", "サッシ周りシール", "防水目地", "建物周長", "その他"],
};

const state = {
  pages: [],
  activePageId: null,
  mode: "area",
  currentPoints: [],
  selectedId: null,
  zoom: 1,
  drawType: "polygon",
  interaction: "draw",
  calibrating: false,
  pointer: null,
  dragPreview: null,
  suppressClick: false,
};

const canvas = document.querySelector("#drawing-canvas");
const ctx = canvas.getContext("2d");
const canvasWrap = document.querySelector("#canvas-wrap");
const fileInput = document.querySelector("#file-input");
const pageList = document.querySelector("#page-list");
const workName = document.querySelector("#work-name");
const saveWork = document.querySelector("#save-work");
const loadWork = document.querySelector("#load-work");
const deleteWork = document.querySelector("#delete-work");
const savedWorkList = document.querySelector("#saved-work-list");
const pageKind = document.querySelector("#page-kind");
const scaleDenominator = document.querySelector("#scale-denominator");
const confirmScale = document.querySelector("#confirm-scale");
const calibrateScale = document.querySelector("#calibrate-scale");
const knownLength = document.querySelector("#known-length");
const applySameScale = document.querySelector("#apply-same-scale");
const applyPlanElevation = document.querySelector("#apply-plan-elevation");
const manualCorrection = document.querySelector("#manual-correction");
const setCorrection = document.querySelector("#set-correction");
const scaleNote = document.querySelector("#scale-note");
const quantityKind = document.querySelector("#quantity-kind");
const selectEdit = document.querySelector("#select-edit");
const drawPolygon = document.querySelector("#draw-polygon");
const drawRect = document.querySelector("#draw-rect");
const buildingInput = document.querySelector("#building-input");
const elevationInput = document.querySelector("#elevation-input");
const finishInput = document.querySelector("#finish-input");
const memoInput = document.querySelector("#memo-input");
const quantityList = document.querySelector("#quantity-list");
const selectedEditor = document.querySelector("#selected-editor");
const editBuilding = document.querySelector("#edit-building");
const editElevation = document.querySelector("#edit-elevation");
const editFinish = document.querySelector("#edit-finish");
const editMemo = document.querySelector("#edit-memo");
const editWidth = document.querySelector("#edit-width");
const editHeight = document.querySelector("#edit-height");
const editArea = document.querySelector("#edit-area");
const editLength = document.querySelector("#edit-length");
const applySelected = document.querySelector("#apply-selected");
const finishShape = document.querySelector("#finish-shape");
const undoPoint = document.querySelector("#undo-point");
const deleteSelected = document.querySelector("#delete-selected");
const exportExcel = document.querySelector("#export-excel");
const totals = document.querySelector("#totals");
const zoomOut = document.querySelector("#zoom-out");
const zoomIn = document.querySelector("#zoom-in");
const zoomReadout = document.querySelector("#zoom-readout");
const cursorReadout = document.querySelector("#cursor-readout");
const prevPage = document.querySelector("#prev-page");
const nextPage = document.querySelector("#next-page");
const pageReadout = document.querySelector("#page-readout");
const actionStatus = document.querySelector("#action-status");
const scalePresetButtons = [...document.querySelectorAll(".scale-presets button")];
const STORAGE_KEY = "drawingQuantityMvpWorks";
const DRAFT_KEY = "draft";
const DB_NAME = "drawingQuantityMvp";
const DB_VERSION = 1;
const WORK_STORE = "works";
let autosaveTimer = null;

function activePage() {
  return state.pages.find((page) => page.id === state.activePageId) ?? null;
}

function activePageIndex() {
  return state.pages.findIndex((page) => page.id === state.activePageId);
}

function selectPageByIndex(index) {
  if (index < 0 || index >= state.pages.length) return;
  state.activePageId = state.pages[index].id;
  state.currentPoints = [];
  state.selectedId = null;
  fitActivePageToView();
  syncControls();
  renderPageList();
  renderQuantityList();
  draw();
  scheduleAutosave();
}

function setStatus(message, type = "") {
  actionStatus.textContent = message;
  actionStatus.className = `action-status${type ? ` ${type}` : ""}`;
}

function setLoading(isLoading) {
  document.body.classList.toggle("is-loading", isLoading);
  fileInput.disabled = isLoading;
}

function openWorkDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(WORK_STORE)) {
        db.createObjectStore(WORK_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function withWorkStore(mode, callback) {
  const db = await openWorkDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(WORK_STORE, mode);
    const store = transaction.objectStore(WORK_STORE);
    const result = callback(store);
    transaction.oncomplete = () => {
      db.close();
      resolve(result);
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getSavedWorks() {
  return withWorkStore("readonly", (store) => requestToPromise(store.getAll()));
}

async function putSavedWork(work) {
  return withWorkStore("readwrite", (store) => requestToPromise(store.put(work)));
}

async function getSavedWork(id) {
  return withWorkStore("readonly", (store) => requestToPromise(store.get(id)));
}

async function deleteSavedWork(id) {
  return withWorkStore("readwrite", (store) => requestToPromise(store.delete(id)));
}

function serializeCurrentWork(name = workName.value.trim() || "自動保存") {
  return {
    id: "draft",
    name,
    savedAt: new Date().toISOString(),
    activePageId: state.activePageId,
    pages: state.pages.map(serializePage),
  };
}

function scheduleAutosave() {
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(async () => {
    if (state.pages.length === 0) return;
    try {
      await putSavedWork(serializeCurrentWork());
    } catch {
      setStatus("自動保存の容量を超えました。Excel出力または作業保存で退避してください。");
    }
  }, 350);
}

async function renderSavedWorks() {
  const entries = (await getSavedWorks())
    .filter((work) => work.id !== DRAFT_KEY)
    .sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || ""));
  savedWorkList.replaceChildren(
    ...entries.map((work) => {
      const option = document.createElement("option");
      option.value = work.id;
      option.textContent = `${work.name} / ${new Date(work.savedAt).toLocaleString("ja-JP")}`;
      return option;
    }),
  );
  loadWork.disabled = entries.length === 0;
  deleteWork.disabled = entries.length === 0;
}

function serializePage(page) {
  return {
    id: page.id,
    name: page.name,
    fileType: page.fileType,
    url: page.url,
    kind: page.kind,
    unsupported: page.unsupported,
    scaleDenominator: page.scaleDenominator,
    dpi: page.dpi,
    scaleCorrection: page.scaleCorrection,
    calibrationMeasured: page.calibrationMeasured,
    scaleConfirmed: page.scaleConfirmed,
    shapes: page.shapes,
  };
}

function pageFromSaved(saved) {
  return new Promise((resolve) => {
    if (saved.unsupported || !saved.url) {
      resolve({ ...saved, image: null, shapes: saved.shapes || [] });
      return;
    }
    const image = new Image();
    image.onload = () => resolve({ ...saved, image, shapes: saved.shapes || [] });
    image.onerror = () => resolve({ ...saved, unsupported: true, image: null, shapes: saved.shapes || [] });
    image.src = saved.url;
  });
}

async function saveCurrentWork() {
  if (state.pages.length === 0) {
    setStatus("保存する図面がありません。");
    return;
  }
  const name = workName.value.trim() || `図面数量拾い_${new Date().toLocaleString("ja-JP")}`;
  const id = crypto.randomUUID();
  try {
    await putSavedWork({ ...serializeCurrentWork(name), id });
  } catch {
    setStatus("ブラウザ内保存に失敗しました。不要な保存済み作業を削除するか、Excel出力で退避してください。");
    return;
  }
  await renderSavedWorks();
  savedWorkList.value = id;
  setStatus(`${name} をこのPCのブラウザ内に保存しました。`, "success");
}

async function loadSelectedWork() {
  const selected = await getSavedWork(savedWorkList.value);
  if (!selected) return;
  if (state.pages.length > 0 && !confirm("現在の作業を保存済み作業で置き換えて読み込みます。よろしいですか？")) return;
  setLoading(true);
  setStatus(`${selected.name} を読み込んでいます。`);
  try {
    state.pages = await Promise.all((selected.pages || []).map(pageFromSaved));
    state.activePageId = selected.activePageId || state.pages[0]?.id || null;
    workName.value = selected.name;
    state.currentPoints = [];
    state.selectedId = null;
    fitActivePageToView();
    syncControls();
    renderPageList();
    renderQuantityList();
    draw();
    setStatus(`${selected.name} を復元しました。`, "success");
  } finally {
    setLoading(false);
  }
}

async function loadDraftWork() {
  const selected = await getSavedWork("draft");
  if (!selected?.pages?.length) return false;
  if (!confirm(`前回の途中作業があります。\n保存時刻: ${new Date(selected.savedAt).toLocaleString("ja-JP")}\n読み込んでよろしいですか？`)) return false;
  setLoading(true);
  try {
    state.pages = await Promise.all(selected.pages.map(pageFromSaved));
    state.activePageId = selected.activePageId || state.pages[0]?.id || null;
    workName.value = selected.name === "自動保存" ? "" : selected.name;
    state.currentPoints = [];
    state.selectedId = null;
    fitActivePageToView();
    syncControls();
    renderPageList();
    renderQuantityList();
    draw();
    setStatus(`前回の途中作業を自動復元しました。保存時刻: ${new Date(selected.savedAt).toLocaleString("ja-JP")}`, "success");
    return true;
  } finally {
    setLoading(false);
  }
}

async function deleteSelectedWork() {
  const selected = await getSavedWork(savedWorkList.value);
  if (!selected) return;
  await deleteSavedWork(savedWorkList.value);
  await renderSavedWorks();
  setStatus(`${selected.name} を保存済み作業から削除しました。`);
}

function fitActivePageToView() {
  const page = activePage();
  if (!page?.image || !canvasWrap) return;
  const availableWidth = Math.max(canvasWrap.clientWidth - 32, 320);
  const availableHeight = Math.max(canvasWrap.clientHeight - 32, 240);
  const fitZoom = Math.min(availableWidth / page.image.naturalWidth, availableHeight / page.image.naturalHeight);
  const nextZoom = Math.min(4, Math.max(0.25, fitZoom));
  state.zoom = nextZoom;
  updateZoomReadout();
}

function updateZoomReadout() {
  zoomReadout.textContent = `${Math.round(state.zoom * 100)}%`;
}

function setZoom(nextZoom) {
  const before = state.zoom;
  state.zoom = Math.min(4, Math.max(0.25, nextZoom));
  updateZoomReadout();
  draw();
  return state.zoom !== before;
}

const PDF_RENDER_DPI = 160;

function mmPerPixel(page) {
  if (!page.scaleConfirmed) return null;
  return (25.4 * page.scaleDenominator) / page.dpi;
}

function metric(page, pixels) {
  const scale = mmPerPixel(page);
  return scale ? (pixels * scale * (page.scaleCorrection || 1)) / 1000 : 0;
}

function metricWithoutCorrection(page, pixels) {
  const scale = mmPerPixel(page);
  return scale ? (pixels * scale) / 1000 : 0;
}

function polygonArea(points) {
  if (points.length < 3) return 0;
  let sum = 0;
  for (let i = 0; i < points.length; i += 1) {
    const next = points[(i + 1) % points.length];
    sum += points[i].x * next.y - next.x * points[i].y;
  }
  return Math.abs(sum) / 2;
}

function lineLength(points, closed = false) {
  if (points.length < 2) return 0;
  let total = 0;
  const limit = closed ? points.length : points.length - 1;
  for (let i = 0; i < limit; i += 1) {
    const next = points[(i + 1) % points.length];
    total += Math.hypot(next.x - points[i].x, next.y - points[i].y);
  }
  return total;
}

function distanceToSegment(point, a, b) {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const lengthSquared = dx * dx + dy * dy;
  if (lengthSquared === 0) return Math.hypot(point.x - a.x, point.y - a.y);
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSquared));
  return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

function pointInPolygon(point, polygon) {
  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const pi = polygon[i];
    const pj = polygon[j];
    const intersects = pi.y > point.y !== pj.y > point.y && point.x < ((pj.x - pi.x) * (point.y - pi.y)) / (pj.y - pi.y) + pi.x;
    if (intersects) inside = !inside;
  }
  return inside;
}

function shapeQuantity(page, shape) {
  const scale = mmPerPixel(page);
  if (!scale) return { width: 0, height: 0, length: 0, area: 0 };
  const adopted = shape.adopted || {};
  if (shape.mode === "line") {
    const closed = shape.kind === "建物周長";
    const measuredLength = metric(page, lineLength(shape.points, closed));
    return { width: 0, height: 0, length: adopted.length || measuredLength, area: 0, measuredLength };
  }
  const xs = shape.points.map((point) => point.x);
  const ys = shape.points.map((point) => point.y);
  const width = metric(page, Math.max(...xs) - Math.min(...xs));
  const height = metric(page, Math.max(...ys) - Math.min(...ys));
  const measuredArea = polygonArea(shape.points) * ((scale * (page.scaleCorrection || 1)) / 1000) ** 2;
  const boundingArea = width * height;
  const isRectangular = shape.points.length === 4 && boundingArea > 0 && Math.abs(measuredArea - boundingArea) <= boundingArea * 0.01;
  const adoptedWidth = adopted.width || width;
  const adoptedHeight = adopted.height || height;
  const area = isRectangular ? adoptedWidth * adoptedHeight : adopted.area || measuredArea;
  return {
    width: adoptedWidth,
    height: adoptedHeight,
    length: 0,
    area,
    perimeter: metric(page, lineLength(shape.points, true)),
    measuredWidth: width,
    measuredHeight: height,
    measuredArea,
    measuredPerimeter: metric(page, lineLength(shape.points, true)),
    isRectangular,
  };
}

function formatNumber(value, digits = 2) {
  return Number.isFinite(value) ? value.toFixed(digits) : "0.00";
}

function updateKindOptions() {
  quantityKind.replaceChildren(
    ...kindOptions[state.mode].map((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      return option;
    }),
  );
}

function imagePageFromSource({ name, fileType, src, kind = "平面図" }) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      resolve({
        id: crypto.randomUUID(),
        name,
        fileType,
        url: src,
        image,
        kind,
        unsupported: false,
        scaleDenominator: 100,
        dpi: PDF_RENDER_DPI,
        scaleCorrection: 1,
        calibrationMeasured: null,
        scaleConfirmed: false,
        shapes: [],
      });
    };
    image.src = src;
  });
}

async function loadPdfFile(file) {
  const response = await fetch(`/api/render-pdf?dpi=${PDF_RENDER_DPI}`, {
    method: "POST",
    headers: { "Content-Type": "application/pdf" },
    body: await file.arrayBuffer(),
  });
  const result = await response.json();
  if (!response.ok) {
    throw new Error(result.error || "PDFの画像化に失敗しました。");
  }
  return Promise.all(
    result.pages.map((page) =>
      imagePageFromSource({
        name: `${file.name} p.${page.pageNumber}`,
        fileType: "application/pdf",
        src: page.dataUrl,
        kind: "平面図",
      }),
    ),
  );
}

function unsupportedPage(file) {
  return {
    id: crypto.randomUUID(),
    name: file.name,
    fileType: file.type || "unknown",
    kind: "使わない",
    unsupported: true,
    scaleDenominator: 100,
    dpi: PDF_RENDER_DPI,
    scaleCorrection: 1,
    calibrationMeasured: null,
    scaleConfirmed: false,
    shapes: [],
  };
}

async function loadInputFile(file) {
  const lowerName = file.name.toLowerCase();
  if (file.type === "application/pdf" || lowerName.endsWith(".pdf")) {
    return loadPdfFile(file);
  }
  if (file.type.startsWith("image/") && !lowerName.endsWith(".tif") && !lowerName.endsWith(".tiff")) {
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
    return [await imagePageFromSource({ name: file.name, fileType: file.type, src: dataUrl })];
  }
  return [unsupportedPage(file)];
}

async function loadFiles(files) {
  const groups = [];
  for (const file of files) {
    groups.push(await loadInputFile(file));
  }
  return groups.flat();
}

async function handleFiles(files) {
  state.pages.forEach((page) => page.url && URL.revokeObjectURL(page.url));
  pageList.textContent = "読み込み中...";
  setLoading(true);
  setStatus("PDFをローカルで画像化しています。ページ数が多い場合は少し待ちます。");
  try {
    state.pages = await loadFiles([...files]);
    setStatus(`${state.pages.length}ページを読み込みました。左のカードまたは前/次で確認できます。`, "success");
  } catch (error) {
    pageList.textContent = "";
    alert(error.message);
    state.pages = [];
    setStatus("読み込みに失敗しました。PDFを画像化してから再度試してください。");
  } finally {
    setLoading(false);
  }
  state.activePageId = state.pages.find((page) => !page.unsupported)?.id ?? state.pages[0]?.id ?? null;
  fitActivePageToView();
  syncControls();
  renderPageList();
  renderQuantityList();
  draw();
  scheduleAutosave();
}

function syncControls() {
  const page = activePage();
  const disabled = !page || page.unsupported;
  pageKind.disabled = !page;
  scaleDenominator.disabled = disabled;
  confirmScale.disabled = disabled;
  calibrateScale.disabled = disabled;
  const canApplyCorrection = !disabled && page.scaleConfirmed && page.scaleCorrection !== 1;
  applySameScale.disabled = !canApplyCorrection;
  applyPlanElevation.disabled = !canApplyCorrection;
  manualCorrection.disabled = disabled;
  setCorrection.disabled = disabled;
  finishShape.disabled = disabled;
  exportExcel.disabled = state.pages.every((item) => item.shapes.length === 0);
  const index = activePageIndex();
  prevPage.disabled = index <= 0;
  nextPage.disabled = index < 0 || index >= state.pages.length - 1;
  pageReadout.textContent = index >= 0 ? `${index + 1} / ${state.pages.length}` : "0 / 0";
  if (!page) {
    scaleNote.textContent = "図面未選択";
    return;
  }
  pageKind.value = page.kind;
  scaleDenominator.value = page.scaleDenominator;
  manualCorrection.value = formatNumber(page.scaleCorrection || 1, 3);
  scalePresetButtons.forEach((button) => {
    button.classList.toggle("selected", button.dataset.scale === String(page.scaleDenominator));
  });
  scaleNote.textContent = page.unsupported
    ? "この形式はMVPでは直接表示できません。PDF/TIFFをPNG/JPGに変換して読み込んでください。"
      : page.scaleConfirmed
        ? `縮尺確定: 1/${page.scaleDenominator}${page.scaleCorrection !== 1 ? ` / 実寸補正 x${formatNumber(page.scaleCorrection, 3)}` : ""}。次は入力モードを選んで図面上をクリックできます。`
        : "図面に書かれた縮尺を入力して確定してください。DPIは内部固定値で自動処理します。";
  renderTotals();
}

function renderPageList() {
  pageList.replaceChildren(
    ...state.pages.map((page, index) => {
      const row = document.createElement("div");
      row.className = `page-item${page.id === state.activePageId ? " active" : ""}`;
      row.tabIndex = 0;
      row.setAttribute("role", "button");
      row.setAttribute("aria-label", `${index + 1}ページ目を表示`);
      const text = document.createElement("div");
      const scaleText = page.scaleConfirmed ? `縮尺 1/${page.scaleDenominator} 確定` : "縮尺未確定";
      const calibrationText = page.scaleCorrection !== 1 ? `<span class="badge calibrated">補正x${formatNumber(page.scaleCorrection, 2)}</span>` : "";
      text.innerHTML = `
        <div class="page-name">${index + 1}. ${page.name}</div>
        <div class="page-meta">
          <span class="badge kind">${page.kind}</span>
          <span class="badge ${page.scaleConfirmed ? "confirmed" : "unconfirmed"}">${scaleText}</span>
          ${calibrationText}
          <span class="badge">${page.shapes.length}件</span>
        </div>`;
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = "表示";
      row.addEventListener("click", () => selectPageByIndex(index));
      row.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          selectPageByIndex(index);
        }
      });
      row.append(text, button);
      return row;
    }),
  );
}

function allShapeEntries() {
  return state.pages.flatMap((page) => page.shapes.map((shape, index) => ({ page, shape, index })));
}

function selectedEntry() {
  return allShapeEntries().find((entry) => entry.shape.id === state.selectedId) || null;
}

function quantityLabel(page, shape) {
  const q = shapeQuantity(page, shape);
  if (shape.mode === "line") return `${formatNumber(q.length)}m`;
  const perimeter = shape.mode === "opening" ? ` / 周長 ${formatNumber(q.perimeter)}m` : "";
  return `${formatNumber(q.width)}m x ${formatNumber(q.height)}m = ${formatNumber(q.area)}m2${perimeter}`;
}

function renderQuantityList() {
  const entries = allShapeEntries();
  if (entries.length === 0) {
    quantityList.innerHTML = '<p class="small-note">確定済み数量はまだありません。</p>';
  } else {
    quantityList.replaceChildren(
      ...entries.map(({ page, shape, index }) => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = `quantity-item${shape.id === state.selectedId ? " selected" : ""}`;
        item.innerHTML = `<strong>${shape.kind} ${quantityLabel(page, shape)}</strong><span>${page.kind} / ${shape.building || "-"} / ${shape.elevation || "-"} / ${shape.finish || "-"} / No.${index + 1}</span>`;
        item.addEventListener("click", () => {
          state.selectedId = shape.id;
          state.activePageId = page.id;
          state.interaction = "select";
          fitActivePageToView();
          syncControls();
          renderAllPanels();
          draw();
        });
        return item;
      }),
    );
  }
  renderSelectedEditor();
}

function renderSelectedEditor() {
  const entry = selectedEntry();
  selectedEditor.hidden = !entry;
  if (!entry) return;
  const q = shapeQuantity(entry.page, entry.shape);
  editBuilding.value = entry.shape.building || "";
  editElevation.value = entry.shape.elevation || "";
  editFinish.value = entry.shape.finish || "";
  editMemo.value = entry.shape.memo || "";
  editWidth.value = entry.shape.mode === "line" ? "" : formatNumber(q.width, 3);
  editHeight.value = entry.shape.mode === "line" ? "" : formatNumber(q.height, 3);
  editArea.value = entry.shape.mode === "line" ? "" : formatNumber(q.area, 3);
  editArea.disabled = entry.shape.mode !== "line" && q.isRectangular;
  editArea.title = editArea.disabled ? "四角形は採用幅x採用高さから自動計算します。" : "";
  editLength.value = entry.shape.mode === "line" ? formatNumber(q.length, 3) : "";
}

function refreshAreaPreview() {
  const entry = selectedEntry();
  if (!entry || entry.shape.mode === "line") return;
  const q = shapeQuantity(entry.page, entry.shape);
  if (!q.isRectangular) return;
  const width = Number(editWidth.value);
  const height = Number(editHeight.value);
  if (Number.isFinite(width) && width > 0 && Number.isFinite(height) && height > 0) {
    editArea.value = formatNumber(width * height, 3);
  }
}

function renderAllPanels() {
  renderPageList();
  renderTotals();
  renderQuantityList();
}

function applySelectedEdits() {
  const entry = selectedEntry();
  if (!entry) return;
  const shape = entry.shape;
  shape.building = editBuilding.value.trim();
  shape.elevation = editElevation.value.trim();
  shape.finish = editFinish.value.trim();
  shape.memo = editMemo.value.trim();
  shape.adopted = shape.adopted || {};
  if (shape.mode === "line") {
    const length = Number(editLength.value);
    shape.adopted.length = Number.isFinite(length) && length > 0 ? length : undefined;
  } else {
    const width = Number(editWidth.value);
    const height = Number(editHeight.value);
    shape.adopted.width = Number.isFinite(width) && width > 0 ? width : undefined;
    shape.adopted.height = Number.isFinite(height) && height > 0 ? height : undefined;
    if (shapeQuantity(entry.page, shape).isRectangular) {
      shape.adopted.area = undefined;
    } else {
      const area = Number(editArea.value);
      shape.adopted.area = Number.isFinite(area) && area > 0 ? area : undefined;
    }
  }
  setStatus("選択内容を更新しました。", "success");
  renderAllPanels();
  draw();
  scheduleAutosave();
}

function applyActiveCorrection(predicate, label) {
  const source = activePage();
  if (!source || !source.scaleConfirmed || source.scaleCorrection === 1) return;
  let count = 0;
  state.pages.forEach((page) => {
    if (page.id === source.id || page.unsupported || !predicate(page, source)) return;
    page.scaleDenominator = source.scaleDenominator;
    page.dpi = source.dpi;
    page.scaleCorrection = source.scaleCorrection;
    page.calibrationMeasured = source.calibrationMeasured;
    page.scaleConfirmed = true;
    count += 1;
  });
  setStatus(`${label}に補正係数 x${formatNumber(source.scaleCorrection, 3)} を適用しました。対象${count}ページ。`, "success");
  renderAllPanels();
  syncControls();
  scheduleAutosave();
}

function drawPoint(point, color = "#176b4d") {
  ctx.beginPath();
  ctx.arc(point.x, point.y, 5 / state.zoom, 0, Math.PI * 2);
  ctx.fillStyle = color;
  ctx.fill();
}

function drawPolyline(points, closed, color, fill = null) {
  if (!points.length) return;
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  points.slice(1).forEach((point) => ctx.lineTo(point.x, point.y));
  if (closed) ctx.closePath();
  if (fill) {
    ctx.fillStyle = fill;
    ctx.fill();
  }
  ctx.strokeStyle = color;
  ctx.lineWidth = 2 / state.zoom;
  ctx.stroke();
  points.forEach((point) => drawPoint(point, color));
}

function draw() {
  const page = activePage();
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (!page || page.unsupported || !page.image) {
    canvas.width = 1200;
    canvas.height = 800;
    canvas.style.width = "1200px";
    canvas.style.height = "800px";
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#667782";
    ctx.font = "24px sans-serif";
    ctx.fillText("図面画像を読み込んでください", 60, 90);
    return;
  }
  canvas.width = page.image.naturalWidth;
  canvas.height = page.image.naturalHeight;
  canvas.style.width = `${canvas.width * state.zoom}px`;
  canvas.style.height = `${canvas.height * state.zoom}px`;
  canvas.style.marginRight = "0";
  canvas.style.marginBottom = "0";
  ctx.drawImage(page.image, 0, 0);
  page.shapes.forEach((shape) => {
    const selected = shape.id === state.selectedId;
    const color = selected ? "#d4482f" : shape.mode === "opening" ? "#276db5" : shape.mode === "line" ? "#8158a8" : "#176b4d";
    const fill = shape.mode === "line" ? null : selected ? "rgb(212 72 47 / 18%)" : "rgb(23 107 77 / 14%)";
    drawPolyline(shape.points, shape.mode !== "line", color, fill);
  });
  drawPolyline(state.currentPoints, false, "#c47717");
  if (state.dragPreview) {
    drawPolyline(state.dragPreview, state.dragPreview.length > 2, "#c47717", "rgb(196 119 23 / 16%)");
  }
}

function canvasPoint(event) {
  const rect = canvas.getBoundingClientRect();
  const point = {
    x: (event.clientX - rect.left) / state.zoom,
    y: (event.clientY - rect.top) / state.zoom,
  };
  if ((event.shiftKey || event.ctrlKey) && state.currentPoints.length > 0) {
    return snapOrthogonal(state.currentPoints.at(-1), point);
  }
  return point;
}

function snapOrthogonal(origin, point) {
  const dx = Math.abs(point.x - origin.x);
  const dy = Math.abs(point.y - origin.y);
  if (dx >= dy) {
    return { x: point.x, y: origin.y };
  }
  return { x: origin.x, y: point.y };
}

function ensureScaleForCalibration(page) {
  page.scaleDenominator = Number(scaleDenominator.value) || page.scaleDenominator || 100;
  page.dpi = PDF_RENDER_DPI;
  page.scaleConfirmed = true;
}

function hitShape(point) {
  const page = activePage();
  if (!page) return null;
  return [...page.shapes].reverse().find((shape) => {
    const vertexHit = shape.points.some((shapePoint) => Math.hypot(shapePoint.x - point.x, shapePoint.y - point.y) < 10 / state.zoom);
    if (vertexHit) return true;
    if (shape.mode === "line") {
      return shape.points.some((shapePoint, index) => {
        const next = shape.points[index + 1];
        return next ? distanceToSegment(point, shapePoint, next) < 8 / state.zoom : false;
      });
    }
    return pointInPolygon(point, shape.points);
  });
}

function hitVertex(point) {
  const page = activePage();
  if (!page) return null;
  const shapes = [...page.shapes].reverse();
  for (const shape of shapes) {
    const pointIndex = shape.points.findIndex((shapePoint) => Math.hypot(shapePoint.x - point.x, shapePoint.y - point.y) < 12 / state.zoom);
    if (pointIndex >= 0) {
      return { shape, pointIndex };
    }
  }
  return null;
}

function addShape(points) {
  const page = activePage();
  if (!page || !page.scaleConfirmed || points.length < 2) return;
  if ((state.mode === "area" || state.mode === "opening") && points.length < 3) return;
  page.shapes.push({
    id: crypto.randomUUID(),
    mode: state.mode,
    kind: quantityKind.value,
    building: buildingInput.value.trim(),
    elevation: elevationInput.value.trim(),
    finish: finishInput.value.trim(),
    memo: memoInput.value.trim(),
    points: points.map((point) => ({ ...point })),
    source: "人確定",
    createdAt: new Date().toLocaleString("ja-JP"),
  });
  state.currentPoints = [];
  state.selectedId = page.shapes.at(-1).id;
  setStatus(`${quantityKind.value}を確定しました。${quantityLabel(page, page.shapes.at(-1))}`, "success");
  renderAllPanels();
  syncControls();
  draw();
  scheduleAutosave();
}

function rectFromTwoPoints(points) {
  const [a, b] = points;
  return [{ x: a.x, y: a.y }, { x: b.x, y: a.y }, { x: b.x, y: b.y }, { x: a.x, y: b.y }];
}

function isRectangleDragMode() {
  return state.interaction === "draw" && (state.mode === "opening" || (state.mode === "area" && state.drawType === "rect"));
}

function tryAutoFinish() {
  if (state.mode === "opening" && state.currentPoints.length === 2) {
    addShape(rectFromTwoPoints(state.currentPoints));
  }
  if (state.mode === "area" && state.drawType === "rect" && state.currentPoints.length === 2) {
    addShape(rectFromTwoPoints(state.currentPoints));
  }
  if (state.mode === "line" && state.currentPoints.length === 2 && quantityKind.value !== "建物周長") {
    addShape(state.currentPoints);
  }
}

function renderTotals() {
  const rows = [];
  state.pages.forEach((page) => {
    page.shapes.forEach((shape) => {
      const q = shapeQuantity(page, shape);
      rows.push({ kind: shape.kind, area: q.area, length: q.length });
    });
  });
  const byKind = new Map();
  rows.forEach((row) => {
    const current = byKind.get(row.kind) ?? { area: 0, length: 0 };
    current.area += row.area;
    current.length += row.length;
    byKind.set(row.kind, current);
  });
  totals.replaceChildren(
    ...[...byKind.entries()].map(([kind, value]) => {
      const row = document.createElement("div");
      row.className = "total-line";
      const unitValue = value.area > 0 ? `${formatNumber(value.area)} m2` : `${formatNumber(value.length)} m`;
      row.innerHTML = `<span>${kind}</span><strong>${unitValue}</strong>`;
      return row;
    }),
  );
  const wallArea = rows.filter((row) => row.kind === "外壁面積").reduce((sum, row) => sum + row.area, 0);
  const openingArea = rows.filter((row) => row.kind === "開口控除").reduce((sum, row) => sum + row.area, 0);
  if (wallArea > 0 || openingArea > 0) {
    const netArea = wallArea - openingArea;
    const openingRate = wallArea > 0 ? (openingArea / wallArea) * 100 : 0;
    const net = document.createElement("div");
    net.className = "total-line";
    net.innerHTML = `<span>外壁差引</span><strong>${formatNumber(netArea)} m2</strong>`;
    const rate = document.createElement("div");
    rate.className = "total-line";
    rate.innerHTML = `<span>開口率</span><strong>${formatNumber(openingRate, 1)}%</strong>`;
    totals.append(net, rate);
  }
}

function addSummary(summary, keys, row) {
  const key = keys.map((name) => row[name] || "-").join("\t");
  const current = summary.get(key) || Object.fromEntries(keys.map((name) => [name, row[name] || "-"]));
  current.分類 = row.分類;
  current.面積m2 = (current.面積m2 || 0) + (row.面積m2 || 0);
  current.延長m = (current.延長m || 0) + (row.延長m || 0);
  current.周長m = (current.周長m || 0) + (row.周長m || 0);
  summary.set(key, current);
}

function elevationDirection(value = "") {
  const text = String(value);
  if (text.includes("北")) return "北";
  if (text.includes("東")) return "東";
  if (text.includes("南")) return "南";
  if (text.includes("西")) return "西";
  if (text.includes("屋上")) return "屋上";
  return "";
}

function directionRank(value = "") {
  const direction = elevationDirection(value) || String(value);
  return { 北: 0, 東: 1, 南: 2, 西: 3, 屋上: 4 }[direction] ?? 9;
}

function categoryRank(value = "") {
  return {
    外壁面積: 0,
    開口控除: 1,
    縦目地: 2,
    横目地: 3,
    打継シール: 4,
    サッシ周りシール: 5,
    防水面積: 6,
    防水目地: 7,
    建物周長: 8,
  }[value] ?? 99;
}

function compareText(a = "", b = "") {
  return String(a || "").localeCompare(String(b || ""), "ja", { numeric: true, sensitivity: "base" });
}

function compareOutputRows(a, b) {
  return (
    compareText(a.棟, b.棟) ||
    directionRank(a.立面方位 || a.立面範囲) - directionRank(b.立面方位 || b.立面範囲) ||
    compareText(a.立面範囲, b.立面範囲) ||
    compareText(a.仕上, b.仕上) ||
    categoryRank(a.分類) - categoryRank(b.分類) ||
    compareText(a.図面種別, b.図面種別) ||
    compareText(a.図面名, b.図面名) ||
    compareText(a.対象ID, b.対象ID)
  );
}

function summaryRank(value = "") {
  return {
    棟別: 0,
    仕上別: 1,
    棟_仕上別: 2,
    立面方位別: 3,
    立面範囲別: 4,
    外壁開口率: 5,
    建具周長合計: 6,
    足場概算: 7,
  }[value] ?? 99;
}

function buildSummaryRows(quantityRows) {
  const rows = [];
  const groups = [
    { title: "棟別", keys: ["棟", "分類"] },
    { title: "仕上別", keys: ["仕上", "分類"] },
    { title: "棟_仕上別", keys: ["棟", "仕上", "分類"] },
    { title: "立面方位別", keys: ["棟", "立面方位", "仕上", "分類"] },
    { title: "立面範囲別", keys: ["棟", "立面範囲", "仕上", "分類"] },
  ];
  groups.forEach((group) => {
    const summary = new Map();
    quantityRows.filter((row) => row.入力種別 !== "概算").forEach((row) => addSummary(summary, group.keys, row));
    summary.forEach((value) => rows.push({ 集計区分: group.title, ...value }));
  });
  const wallArea = quantityRows.filter((row) => row.分類 === "外壁面積").reduce((sum, row) => sum + (row.面積m2 || 0), 0);
  const openingArea = quantityRows.filter((row) => row.分類 === "開口控除").reduce((sum, row) => sum + (row.面積m2 || 0), 0);
  const openingPerimeter = quantityRows.filter((row) => row.分類 === "開口控除").reduce((sum, row) => sum + (row.周長m || 0), 0);
  rows.push({
    集計区分: "外壁開口率",
    分類: "開口率",
    面積m2: wallArea > 0 ? (openingArea / wallArea) * 100 : 0,
    備考: `開口控除 ${formatNumber(openingArea)}m2 / 外壁面積 ${formatNumber(wallArea)}m2`,
  });
  rows.push({
    集計区分: "建具周長合計",
    分類: "開口控除",
    周長m: openingPerimeter,
    備考: "開口控除(建具)の周長合計。額縁・シーリングなど建具周り施工数量の目安。",
  });
  const buildingPerimeter = quantityRows.filter((row) => row.分類 === "建物周長").reduce((sum, row) => sum + (row.延長m || 0), 0);
  if (wallArea > 0) {
    rows.push({
      集計区分: "足場概算",
      分類: "外部足場概算面積",
      面積m2: wallArea,
      備考: "外壁面積を足場概算数量として転用。正確な足場計画数量ではありません。",
    });
  }
  if (buildingPerimeter > 0) {
    rows.push({
      集計区分: "足場概算",
      分類: "安全手すり概算延長",
      延長m: buildingPerimeter,
      備考: "建物周長を安全手すり概算数量として転用。正確な足場計画数量ではありません。",
    });
  }
  return rows.sort((a, b) => summaryRank(a.集計区分) - summaryRank(b.集計区分) || compareOutputRows(a, b));
}

function exportWorkbook() {
  const quantityRows = [];
  const evidenceRows = [];
  state.pages.forEach((page) => {
    page.shapes.forEach((shape, index) => {
      const q = shapeQuantity(page, shape);
      const base = {
        図面名: page.name,
        図面種別: page.kind,
        対象ID: shape.id,
        棟: shape.building,
        立面方位: elevationDirection(shape.elevation),
        立面範囲: shape.elevation,
        仕上: shape.finish,
        分類: shape.kind,
        入力種別: shape.mode === "area" ? "面積" : shape.mode === "opening" ? "開口" : "線数量",
        実測幅m: q.measuredWidth,
        実測高さm: q.measuredHeight,
        実測延長m: q.measuredLength,
        実測面積m2: q.measuredArea,
        実測周長m: q.measuredPerimeter,
        幅m: q.width,
        高さm: q.height,
        延長m: q.length,
        面積m2: q.area,
        周長m: q.perimeter,
        手修正: shape.adopted && Object.values(shape.adopted).some((value) => value) ? "あり" : "",
        縮尺: `1/${page.scaleDenominator}`,
        実寸補正係数: page.scaleCorrection,
        補正確認m: page.calibrationMeasured,
        画像DPI: page.dpi,
        備考: shape.memo,
      };
      quantityRows.push(base);
      evidenceRows.push({
        図面名: page.name,
        図面種別: page.kind,
        対象ID: shape.id,
        棟: shape.building,
        立面方位: elevationDirection(shape.elevation),
        立面範囲: shape.elevation,
        仕上: shape.finish,
        分類: shape.kind,
        入力種別: shape.mode === "area" ? "面積" : shape.mode === "opening" ? "開口" : "線数量",
        採用幅m: q.width,
        採用高さm: q.height,
        採用延長m: q.length,
        採用面積m2: q.area,
        採用周長m: q.perimeter,
        根拠番号: index + 1,
        AI候補人確定: shape.source,
        縮尺確定: page.scaleConfirmed ? "確定" : "未確定",
        実寸補正係数: page.scaleCorrection,
        補正確認m: page.calibrationMeasured,
        概算区分: ["外部足場概算面積", "安全手すり概算延長"].includes(shape.kind) ? "概算" : "",
        座標: shape.points.map((point) => `${Math.round(point.x)},${Math.round(point.y)}`).join(" / "),
        承認日時: shape.createdAt,
        備考: shape.memo,
      });
    });
  });
  quantityRows.sort(compareOutputRows);
  evidenceRows.sort(compareOutputRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(quantityRows), "数量表");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(evidenceRows), "根拠一覧");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(buildSummaryRows(quantityRows)), "集計");
  XLSX.writeFile(workbook, `図面数量拾い_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

fileInput.addEventListener("change", (event) => handleFiles(event.target.files));
saveWork.addEventListener("click", saveCurrentWork);
loadWork.addEventListener("click", loadSelectedWork);
deleteWork.addEventListener("click", deleteSelectedWork);
applySelected.addEventListener("click", applySelectedEdits);
editWidth.addEventListener("input", refreshAreaPreview);
editHeight.addEventListener("input", refreshAreaPreview);
pageKind.addEventListener("change", () => {
  const page = activePage();
  if (!page) return;
  page.kind = pageKind.value;
  renderAllPanels();
  scheduleAutosave();
});
confirmScale.addEventListener("click", () => {
  const page = activePage();
  if (!page) return;
  page.scaleDenominator = Number(scaleDenominator.value) || 100;
  page.dpi = PDF_RENDER_DPI;
  page.scaleCorrection = page.scaleCorrection || 1;
  page.scaleConfirmed = true;
  confirmScale.textContent = "確定しました";
  setTimeout(() => {
    confirmScale.textContent = "この縮尺を確定";
  }, 900);
  setStatus(`${page.name} を 1/${page.scaleDenominator} で確定しました。`, "success");
  renderAllPanels();
  syncControls();
  scheduleAutosave();
});
calibrateScale.addEventListener("click", () => {
  const page = activePage();
  if (!page) return;
  state.calibrating = true;
  state.currentPoints = [];
  setStatus("補正開始。寸法線の1点目をクリックしてください。間違えたら1点戻すで取り消せます。");
  draw();
});
applySameScale.addEventListener("click", () => {
  applyActiveCorrection(
    (page, source) => page.scaleDenominator === source.scaleDenominator,
    `同じ縮尺1/${activePage()?.scaleDenominator ?? ""}の図面`,
  );
});
applyPlanElevation.addEventListener("click", () => {
  applyActiveCorrection(
    (page) => page.kind === "平面図" || page.kind === "立面図",
    "平面図・立面図",
  );
});
setCorrection.addEventListener("click", () => {
  const page = activePage();
  const correction = Number(manualCorrection.value);
  if (!page || !Number.isFinite(correction) || correction <= 0) {
    setStatus("補正係数は0より大きい数値を入れてください。");
    return;
  }
  page.scaleCorrection = correction;
  page.scaleConfirmed = true;
  setStatus(`補正係数 x${formatNumber(correction, 3)} を採用しました。`, "success");
  renderAllPanels();
  syncControls();
  draw();
  scheduleAutosave();
});
document.querySelectorAll(".mode-button").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".mode-button").forEach((item) => item.classList.remove("selected"));
    button.classList.add("selected");
    state.mode = button.dataset.mode;
    state.interaction = "draw";
    state.currentPoints = [];
    state.calibrating = false;
    selectEdit.classList.remove("selected");
    updateKindOptions();
    draw();
  });
});
selectEdit.addEventListener("click", () => {
  state.interaction = "select";
  state.currentPoints = [];
  state.calibrating = false;
  selectEdit.classList.add("selected");
  drawPolygon.classList.remove("selected");
  drawRect.classList.remove("selected");
  setStatus("選択/編集モードです。図形をクリックして選択、角点をドラッグして修正、選択削除で削除できます。");
  draw();
});
drawPolygon.addEventListener("click", () => {
  state.interaction = "draw";
  state.drawType = "polygon";
  selectEdit.classList.remove("selected");
  drawPolygon.classList.add("selected");
  drawRect.classList.remove("selected");
});
drawRect.addEventListener("click", () => {
  state.interaction = "draw";
  state.drawType = "rect";
  selectEdit.classList.remove("selected");
  drawRect.classList.add("selected");
  drawPolygon.classList.remove("selected");
});
scalePresetButtons.forEach((button) => {
  button.addEventListener("click", () => {
    scaleDenominator.value = button.dataset.scale;
    scalePresetButtons.forEach((item) => item.classList.toggle("selected", item === button));
  });
});
canvas.addEventListener("pointerdown", (event) => {
  const page = activePage();
  if (!page || page.unsupported) return;
  const point = canvasPoint(event);
  const vertex = state.interaction === "select" && !state.calibrating && state.currentPoints.length === 0 ? hitVertex(point) : null;
  const canDrawRect = page.scaleConfirmed && !state.calibrating && isRectangleDragMode();
  state.pointer = {
    startClientX: event.clientX,
    startClientY: event.clientY,
    lastClientX: event.clientX,
    lastClientY: event.clientY,
    startPoint: point,
    type: vertex ? "vertex" : canDrawRect ? "rect" : "pan",
    vertex,
    moved: false,
  };
  if (vertex) {
    state.selectedId = vertex.shape.id;
  }
  canvas.setPointerCapture?.(event.pointerId);
});
canvas.addEventListener("pointermove", (event) => {
  if (!state.pointer) return;
  const dx = event.clientX - state.pointer.lastClientX;
  const dy = event.clientY - state.pointer.lastClientY;
  const movedTotal = Math.hypot(event.clientX - state.pointer.startClientX, event.clientY - state.pointer.startClientY);
  if (movedTotal > 4) state.pointer.moved = true;
  if (state.pointer.type === "rect") {
    state.dragPreview = rectFromTwoPoints([state.pointer.startPoint, canvasPoint(event)]);
    draw();
  } else if (state.pointer.type === "vertex" && state.pointer.vertex) {
    state.pointer.vertex.shape.points[state.pointer.vertex.pointIndex] = canvasPoint(event);
    draw();
  } else if (state.pointer.moved) {
    canvasWrap.classList.add("is-panning");
    canvasWrap.scrollLeft -= dx;
    canvasWrap.scrollTop -= dy;
  }
  state.pointer.lastClientX = event.clientX;
  state.pointer.lastClientY = event.clientY;
});
canvas.addEventListener("pointerup", (event) => {
  if (!state.pointer) return;
  const pointer = state.pointer;
  state.pointer = null;
  canvas.releasePointerCapture?.(event.pointerId);
  canvasWrap.classList.remove("is-panning");
  if (pointer.moved) {
    state.suppressClick = true;
    setTimeout(() => {
      state.suppressClick = false;
    }, 0);
  }
  if (pointer.type === "rect" && pointer.moved) {
    const endPoint = canvasPoint(event);
    state.dragPreview = null;
    addShape(rectFromTwoPoints([pointer.startPoint, endPoint]));
    return;
  }
  if (pointer.type === "vertex" && pointer.moved) {
    state.dragPreview = null;
    renderAllPanels();
    draw();
    scheduleAutosave();
    return;
  }
  state.dragPreview = null;
  draw();
});
canvas.addEventListener("pointercancel", () => {
  state.pointer = null;
  state.dragPreview = null;
  canvasWrap.classList.remove("is-panning");
  draw();
});
canvas.addEventListener("click", (event) => {
  if (state.suppressClick) return;
  const page = activePage();
  if (!page || page.unsupported) return;
  const point = canvasPoint(event);
  if (state.calibrating) {
    ensureScaleForCalibration(page);
    state.currentPoints.push(point);
    if (state.currentPoints.length === 1) {
      setStatus("1点目を記録しました。寸法線の反対側をクリックしてください。");
      draw();
      return;
    }
    if (state.currentPoints.length === 2) {
      const measured = metricWithoutCorrection(page, lineLength(state.currentPoints, false));
      const actual = Number(knownLength.value);
      if (actual > 0 && measured > 0) {
        page.scaleCorrection = actual / measured;
        page.calibrationMeasured = actual;
        manualCorrection.value = formatNumber(page.scaleCorrection, 3);
        state.currentPoints = [];
        state.calibrating = false;
        setStatus(`計測${formatNumber(measured, 3)}m / 実寸${formatNumber(actual, 3)}m。補正係数 x${formatNumber(page.scaleCorrection, 3)}。`, "success");
        renderAllPanels();
        syncControls();
        draw();
        scheduleAutosave();
        return;
      }
      state.currentPoints = [];
      setStatus("実寸mを入力してから、もう一度寸法線の両端をクリックしてください。");
      draw();
      return;
    }
  }
  const hit = state.interaction === "select" ? hitShape(point) : null;
  if (hit && state.currentPoints.length === 0) {
    state.selectedId = hit.id;
    renderQuantityList();
    draw();
    return;
  }
  if (state.interaction === "select") {
    state.selectedId = null;
    draw();
    return;
  }
  if (!page.scaleConfirmed) return;
  state.currentPoints.push(point);
  tryAutoFinish();
  draw();
});
canvas.addEventListener("mousemove", (event) => {
  const point = canvasPoint(event);
  cursorReadout.textContent = `x:${Math.round(point.x)} y:${Math.round(point.y)}`;
});
finishShape.addEventListener("click", () => addShape(state.currentPoints));
undoPoint.addEventListener("click", () => {
  state.currentPoints.pop();
  if (state.calibrating) {
    setStatus(state.currentPoints.length === 0 ? "補正用の点を取り消しました。寸法線の1点目をクリックしてください。" : "1点戻しました。寸法線の反対側をクリックしてください。");
  }
  draw();
});
deleteSelected.addEventListener("click", () => {
  const page = activePage();
  if (!page || !state.selectedId) return;
  page.shapes = page.shapes.filter((shape) => shape.id !== state.selectedId);
  state.selectedId = null;
  renderAllPanels();
  syncControls();
  draw();
  scheduleAutosave();
});
zoomOut.addEventListener("click", () => setZoom(state.zoom - 0.1));
zoomIn.addEventListener("click", () => setZoom(state.zoom + 0.1));
prevPage.addEventListener("click", () => selectPageByIndex(activePageIndex() - 1));
nextPage.addEventListener("click", () => selectPageByIndex(activePageIndex() + 1));
exportExcel.addEventListener("click", exportWorkbook);

updateKindOptions();
renderSavedWorks();
updateZoomReadout();
syncControls();
renderQuantityList();
draw();
loadDraftWork().catch(() => setStatus("前回の途中作業を復元できませんでした。保存済み作業から読み込んでください。"));
