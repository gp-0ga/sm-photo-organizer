import "./compat.js";
import "./styles.css";
import { readAssetsFromAssetListWorkbook, readAssetsFromWorkbook, readInspectionWorkbooks } from "./excel.js";
import { matchInspectionAssets } from "./domain.js";
import { createAssetQr } from "./qr.js";
import { clearMarkerSelection, loadMarkerAssets, loadMarkerSelection, saveMarkerAssets, saveMarkerSelection } from "./marker-storage.js";
import { parseMarkerAssetsFromTextItems } from "./pdf-marker.js";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import pdfjsWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

const excelInput = document.querySelector("#marker-excel-input");
const healthExcelInput = document.querySelector("#marker-health-excel-input");
const inspectionFolderInput = document.querySelector("#marker-inspection-folder-input");
const healthImportPanel = document.querySelector("#marker-health-import");
const assetListImportPanel = document.querySelector("#marker-asset-list-import");
const importModeInputs = document.querySelectorAll('input[name="marker-import-mode"]');
const pdfInput = document.querySelector("#marker-pdf-input");
const excelStatus = document.querySelector("#marker-excel-status");
const viewPanel = document.querySelector("#marker-view-panel");
const printPanel = document.querySelector("#marker-print-panel");
const assetSearch = document.querySelector("#marker-asset-search");
const assetSelect = document.querySelector("#marker-asset-select");
const displayQr = document.querySelector("#marker-display-qr");
const displayNumber = document.querySelector("#marker-display-number");
const displayName = document.querySelector("#marker-display-name");
const previousButton = document.querySelector("#marker-previous");
const nextButton = document.querySelector("#marker-next");
const position = document.querySelector("#marker-position");
const printButton = document.querySelector("#marker-print");
const cardList = document.querySelector("#marker-card-list");
const changeListButton = document.querySelector("#marker-change-list");
const browserNotice = document.querySelector("#marker-browser-notice");
const roomPanel = document.querySelector("#room-panel");
const roomBuildingInput = document.querySelector("#room-building-input");
const roomNameInput = document.querySelector("#room-name-input");
const roomFloorSelect = document.querySelector("#room-floor-select");
const roomWallSelect = document.querySelector("#room-wall-select");
const roomCeilingSelect = document.querySelector("#room-ceiling-select");
const roomSaveButton = document.querySelector("#room-save");
const roomSourceInput = document.querySelector("#room-source-input");
const roomSourceStatus = document.querySelector("#room-source-status");
const roomCandidateList = document.querySelector("#room-candidate-list");
const roomList = document.querySelector("#room-list");

let assets = [];
let assetListAssets = [];
let currentIndex = 0;
const qrUrls = new Map();
let cardsRendered = false;
const roomsStorageKey = "asset-marker.rooms.v1";
let rooms = loadRooms();

function loadRooms() {
  try {
    const stored = JSON.parse(localStorage.getItem(roomsStorageKey) || "[]");
    return Array.isArray(stored) ? stored.filter((room) => room && room.id && room.name) : [];
  } catch {
    return [];
  }
}

function saveRooms() {
  localStorage.setItem(roomsStorageKey, JSON.stringify(rooms));
}

async function readPdfText(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  const document = await loadingTask.promise;
  const items = [];
  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const content = await page.getTextContent();
    items.push(...content.items.map((item) => item.str ?? ""), "\n");
  }
  await loadingTask.destroy();
  return items;
}

async function readMarkerAssetsFromPdf(file) {
  const assets = parseMarkerAssetsFromTextItems(await readPdfText(file));
  if (!assets.length) throw new Error("このPDFから資産情報を読み取れませんでした。PCアプリの「全マーカーを印刷」で作成したPDFを選択してください。");
  return assets;
}

function roomCandidatesFromText(text) {
  const seen = new Set();
  return (Array.isArray(text) ? text : [text]).flatMap((value) => String(value ?? "").split(/[\r\n]+/))
    .map((line) => line.replace(/[()（）［］【】]/g, " ").replace(/\s+/g, " ").trim())
    .filter((line) => line.length >= 2 && line.length <= 40)
    .filter((line) => /(室|ホール|廊下|便所|倉庫|機械|事務|会議|玄関|階段|食堂|更衣|浴室|休憩)/.test(line))
    .filter((line) => { if (seen.has(line)) return false; seen.add(line); return true; })
    .slice(0, 80);
}

function setStatus(message, tone = "neutral") {
  excelStatus.className = `status ${tone}`;
  excelStatus.textContent = message;
}

function selectedImportMode() {
  return document.querySelector('input[name="marker-import-mode"]:checked')?.value ?? "health";
}

function resetImportForMode() {
  const healthMode = selectedImportMode() === "health";
  healthImportPanel.hidden = !healthMode;
  assetListImportPanel.hidden = healthMode;
  healthExcelInput.value = "";
  excelInput.value = "";
  inspectionFolderInput.value = "";
  inspectionFolderInput.disabled = true;
  assetListAssets = [];
  viewPanel.hidden = true;
  printPanel.hidden = true;
  document.body.classList.remove("marker-ready");
  setStatus(healthMode ? "健全度判定表を選択してください。" : "最初に資産リストを選択してください。", "neutral");
}

async function qrFor(asset) {
  if (!qrUrls.has(asset.assetNumber)) qrUrls.set(asset.assetNumber, await createAssetQr(asset.assetNumber));
  return qrUrls.get(asset.assetNumber);
}

async function showAsset(index) {
  if (!assets.length) return;
  currentIndex = Math.max(0, Math.min(index, assets.length - 1));
  const asset = assets[currentIndex];
  try { saveMarkerSelection(localStorage, asset.assetNumber); } catch { /* 保存できなくても表示は続ける */ }
  assetSelect.value = asset.assetNumber;
  displayQr.src = await qrFor(asset);
  displayQr.alt = `${asset.assetNumber} ${asset.assetName} 資産切替QR`;
  displayNumber.textContent = asset.assetNumber;
  displayName.textContent = asset.assetName;
  displayName.classList.toggle("long", asset.assetName.length > 22);
  displayName.classList.toggle("very-long", asset.assetName.length > 38);
  position.textContent = `${currentIndex + 1} / ${assets.length}`;
  previousButton.disabled = currentIndex === 0;
  nextButton.disabled = currentIndex === assets.length - 1;
  document.querySelectorAll(".marker-card.selected").forEach((card) => card.classList.remove("selected"));
  document.querySelector(`[data-asset-number="${CSS.escape(asset.assetNumber)}"]`)?.classList.add("selected");
}

function renderAssetOptions(query = "") {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const matches = assets.filter((asset) => {
    if (!normalizedQuery) return true;
    return `${asset.assetNumber} ${asset.assetName}`.toLocaleLowerCase().includes(normalizedQuery);
  });
  const selected = assets[currentIndex]?.assetNumber;
  assetSelect.replaceChildren();
  for (const asset of matches) {
    assetSelect.add(new Option(`${asset.assetNumber} ${asset.assetName}`, asset.assetNumber));
  }
  assetSelect.disabled = matches.length === 0;
  if (!matches.length) return;
  const next = matches.some((asset) => asset.assetNumber === selected) ? selected : matches[0].assetNumber;
  assetSelect.value = next;
}

async function renderMarkerCards() {
  if (cardsRendered) return;
  cardList.replaceChildren();
  for (let index = 0; index < assets.length; index += 1) {
    const asset = assets[index];
    const card = document.createElement("button");
    card.type = "button";
    card.className = "marker-card marker-card-button";
    card.dataset.assetNumber = asset.assetNumber;
    const start = document.createElement("b");
    start.className = "marker-card-start";
    start.textContent = "ここからこの資産";
    const image = document.createElement("img");
    image.src = await qrFor(asset);
    image.alt = `${asset.assetNumber} ${asset.assetName} 資産切替QR`;
    const number = document.createElement("strong");
    number.textContent = asset.assetNumber;
    const name = document.createElement("span");
    name.textContent = asset.assetName;
    const payload = document.createElement("small");
    payload.className = "marker-payload";
    payload.textContent = `SM-ASSET-NAME|1|${asset.assetNumber}|${asset.assetName}`;
    card.append(start, image, number, name, payload);
    card.addEventListener("click", () => {
      showAsset(index);
      viewPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    cardList.append(card);
  }
  cardsRendered = true;
}

function roomAssetOptions(select, selectedValue = "") {
  select.replaceChildren(new Option("未選択", ""));
  for (const asset of assets) {
    select.add(new Option(`${asset.assetNumber} ${asset.assetName}`, asset.assetNumber));
  }
  select.value = selectedValue || "";
}

function showAssetByNumber(assetNumber) {
  const index = assets.findIndex((asset) => asset.assetNumber === assetNumber);
  if (index >= 0) {
    void showAsset(index);
    viewPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function renderRoomInputs() {
  roomAssetOptions(roomFloorSelect);
  roomAssetOptions(roomWallSelect);
  roomAssetOptions(roomCeilingSelect);
}

function renderRooms() {
  roomList.replaceChildren();
  for (const room of rooms) {
    const row = document.createElement("div");
    row.className = "room-row";
    const text = document.createElement("span");
    const title = document.createElement("strong");
    title.textContent = `${room.building ? `${room.building} / ` : ""}${room.name}`;
    const details = document.createElement("small");
    details.className = "room-assets";
    details.textContent = `床: ${room.floor || "未選択"}　壁: ${room.wall || "未選択"}　天井: ${room.ceiling || "未選択"}`;
    text.append(title, document.createElement("br"), details);
    const actions = document.createElement("div");
    for (const [label, assetNumber] of [["床", room.floor], ["壁", room.wall], ["天井", room.ceiling]]) {
      if (!assetNumber) continue;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "secondary";
      button.textContent = label;
      button.addEventListener("click", () => showAssetByNumber(assetNumber));
      actions.append(button);
    }
    row.append(text, actions);
    roomList.append(row);
  }
}

function clearRoomInputs() {
  roomBuildingInput.value = "";
  roomNameInput.value = "";
  roomFloorSelect.value = "";
  roomWallSelect.value = "";
  roomCeilingSelect.value = "";
}

async function readRoomCandidates(file) {
  const lowerName = file.name.toLocaleLowerCase();
  if (file.type === "application/pdf" || lowerName.endsWith(".pdf")) {
    const text = await readPdfText(file);
    return { candidates: roomCandidatesFromText(text), scanned: !roomCandidatesFromText(text).length };
  }
  if (typeof window.TextDetector === "function") {
    const bitmap = await createImageBitmap(file);
    const detected = await new window.TextDetector().detect(bitmap);
    bitmap.close?.();
    return { candidates: detected.map((item) => item.rawValue).filter(Boolean), scanned: false };
  }
  return { candidates: [], scanned: true };
}

function renderRoomCandidates(candidates) {
  roomCandidateList.replaceChildren();
  for (const candidate of candidates) {
    const row = document.createElement("div");
    row.className = "room-candidate";
    const label = document.createElement("span");
    label.textContent = candidate;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "secondary";
    button.textContent = "部屋名に使用";
    button.addEventListener("click", () => { roomNameInput.value = candidate; roomNameInput.focus(); });
    row.append(label, button);
    roomCandidateList.append(row);
  }
}

async function useAssets(nextAssets, message, tone = "success") {
  assets = nextAssets;
  qrUrls.clear();
  cardsRendered = false;
  cardList.replaceChildren();
  assetSearch.value = "";
  renderAssetOptions();
  renderRoomInputs();
  renderRooms();
  roomPanel.hidden = false;
  viewPanel.hidden = false;
  printPanel.hidden = true;
  document.body.classList.add("marker-ready");
  const storedSelection = loadMarkerSelection(localStorage);
  const restoredIndex = Math.max(0, assets.findIndex((asset) => asset.assetNumber === storedSelection));
  await showAsset(restoredIndex);
  setStatus(message, tone);
}

roomSaveButton.addEventListener("click", () => {
  const name = roomNameInput.value.trim();
  if (!name) {
    roomSourceStatus.className = "status error";
    roomSourceStatus.textContent = "部屋名を入力してください。";
    return;
  }
  const room = {
    id: `${roomBuildingInput.value.trim()}::${name}`,
    building: roomBuildingInput.value.trim(),
    name,
    floor: roomFloorSelect.value,
    wall: roomWallSelect.value,
    ceiling: roomCeilingSelect.value,
  };
  const existing = rooms.findIndex((item) => item.id === room.id);
  if (existing >= 0) rooms[existing] = room;
  else rooms.push(room);
  saveRooms();
  renderRooms();
  clearRoomInputs();
  roomSourceStatus.className = "status success";
  roomSourceStatus.textContent = `「${room.building ? `${room.building} / ` : ""}${room.name}」を登録しました。`;
});

roomSourceInput.addEventListener("change", async () => {
  const file = roomSourceInput.files?.[0];
  if (!file) return;
  roomSourceStatus.className = "status working";
  roomSourceStatus.textContent = `${file.name}から部屋名候補を読み込んでいます…`;
  try {
    const result = await readRoomCandidates(file);
    renderRoomCandidates(result.candidates);
    if (result.candidates.length) {
      roomSourceStatus.className = "status success";
      roomSourceStatus.textContent = `${result.candidates.length}件の候補を表示しました。内容を確認して部屋名に使用してください。`;
    } else if (result.scanned) {
      roomSourceStatus.className = "status neutral";
      roomSourceStatus.textContent = "文字情報を読み取れませんでした。スキャン画像はこの画面で部屋名を手入力してください。";
    }
  } catch (error) {
    roomSourceStatus.className = "status error";
    roomSourceStatus.textContent = error.message ?? String(error);
  }
});

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  inspectionFolderInput.disabled = true;
  inspectionFolderInput.value = "";
  setStatus(`${file.name}を読み込んでいます…`, "working");
  try {
    assetListAssets = await readAssetsFromAssetListWorkbook(file);
    inspectionFolderInput.disabled = false;
    setStatus(`${assetListAssets.length}資産を読み取りました。次に、点検表フォルダーを選択してください。`, "success");
  } catch (error) {
    assetListAssets = [];
    assets = [];
    document.body.classList.remove("marker-ready");
    cardList.replaceChildren();
    setStatus(error.message ?? String(error), "error");
  }
});

healthExcelInput.addEventListener("change", async () => {
  const file = healthExcelInput.files?.[0];
  if (!file) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  setStatus(`${file.name}のNo11を読み込んでいます…`, "working");
  try {
    const loadedAssets = await readAssetsFromWorkbook(file);
    saveMarkerAssets(localStorage, loadedAssets);
    clearMarkerSelection(localStorage);
    await useAssets(loadedAssets, `${loadedAssets.length}資産を健全度判定表から読み込みました。Excelは外部へ送信していません。`);
  } catch (error) {
    assets = [];
    document.body.classList.remove("marker-ready");
    cardList.replaceChildren();
    setStatus(error.message ?? String(error), "error");
  }
});

for (const input of importModeInputs) input.addEventListener("change", resetImportForMode);

inspectionFolderInput.addEventListener("change", async () => {
  const files = inspectionFolderInput.files;
  if (!files?.length || !assetListAssets.length) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  setStatus(`${files.length}ファイルを確認しています…`, "working");
  try {
    const loaded = await readInspectionWorkbooks(files, (done, total) => {
      if (done === total || done % 20 === 0) setStatus(`点検表を確認しています… ${done} / ${total}`, "working");
    });
    const matched = matchInspectionAssets(assetListAssets, loaded.inspections);
    if (!matched.assets.length) throw new Error("資産リストと一致する点検表がありませんでした。別の点検表フォルダーを選択してください。");
    const warningCount = loaded.failures.length + matched.duplicates.length + matched.unmatched.length + matched.nameMismatches.length;
    const notes = [];
    if (loaded.failures.length) notes.push(`読取不可 ${loaded.failures.length}件`);
    if (matched.duplicates.length) notes.push(`重複 ${matched.duplicates.length}件`);
    if (matched.unmatched.length) notes.push(`資産リストにない番号 ${matched.unmatched.length}件`);
    if (matched.nameMismatches.length) notes.push(`名称相違 ${matched.nameMismatches.length}件`);
    saveMarkerAssets(localStorage, matched.assets);
    clearMarkerSelection(localStorage);
    const message = warningCount
      ? `${matched.assets.length}資産を読み込みました。確認事項: ${notes.join("、")}。該当ファイルを除いて表示します。`
      : `${matched.assets.length}資産を点検表フォルダーから読み込みました。Excelは外部へ送信していません。`;
    await useAssets(matched.assets, message, warningCount ? "neutral" : "success");
  } catch (error) {
    assets = [];
    document.body.classList.remove("marker-ready");
    cardList.replaceChildren();
    setStatus(error.message ?? String(error), "error");
  }
});

pdfInput.addEventListener("change", async () => {
  const file = pdfInput.files?.[0];
  if (!file) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  setStatus(`${file.name}から資産情報を読み込んでいます…`, "working");
  try {
    const loadedAssets = await readMarkerAssetsFromPdf(file);
    saveMarkerAssets(localStorage, loadedAssets);
    clearMarkerSelection(localStorage);
    await useAssets(loadedAssets, `${loadedAssets.length}資産をPDFから読み込みました。圏外でもマーカーを表示できます。`);
  } catch (error) {
    document.body.classList.remove("marker-ready");
    const where = String(error?.stack ?? "").split("\n").slice(0, 3).join(" / ").slice(0, 300);
    setStatus(`${error.message ?? String(error)}${where ? `（詳細: ${where}）` : ""}`, "error");
  }
});

assetSelect.addEventListener("change", () => {
  const index = assets.findIndex((asset) => asset.assetNumber === assetSelect.value);
  if (index >= 0) showAsset(index);
});
assetSearch.addEventListener("input", () => {
  const before = assets[currentIndex]?.assetNumber;
  renderAssetOptions(assetSearch.value);
  const next = assets.findIndex((asset) => asset.assetNumber === assetSelect.value);
  if (next >= 0 && assets[next].assetNumber !== before) showAsset(next);
});
previousButton.addEventListener("click", () => showAsset(currentIndex - 1));
nextButton.addEventListener("click", () => showAsset(currentIndex + 1));
printButton.addEventListener("click", async () => {
  printButton.disabled = true;
  setStatus("印刷用マーカーを準備しています…", "working");
  try {
    await renderMarkerCards();
    printPanel.hidden = false;
    setStatus(`${assets.length}資産の印刷用マーカーを準備しました。`, "success");
    window.print();
  } finally {
    printButton.disabled = false;
  }
});

changeListButton.addEventListener("click", () => {
  document.body.classList.remove("marker-ready");
  viewPanel.hidden = true;
  printPanel.hidden = true;
  excelInput.value = "";
  healthExcelInput.value = "";
  inspectionFolderInput.value = "";
  inspectionFolderInput.disabled = true;
  assetListAssets = [];
  const nextPrompt = selectedImportMode() === "health" ? "健全度判定表" : "資産リスト";
  setStatus(`新しい${nextPrompt}を選択してください。保存済み一覧は、新しい一覧を読み込むまで残ります。`, "neutral");
  document.querySelector(".marker-load-panel")?.scrollIntoView({ block: "start" });
});

const restoredAssets = loadMarkerAssets(localStorage);
if (/CriOS/i.test(navigator.userAgent) && !restoredAssets) {
  browserNotice.hidden = false;
  browserNotice.textContent = "Chromeで開いています。Safariの保存内容は共有されません。このChromeでExcelを読み込むか、Safariで開いてください。";
}
if (restoredAssets) {
  useAssets(restoredAssets, `${restoredAssets.length}資産をこのブラウザから復元しました。圏外でもマーカーを表示できます。`);
}

if ("serviceWorker" in navigator && location.protocol === "https:") {
  navigator.serviceWorker.register("/sw.js").catch(() => {});
}
