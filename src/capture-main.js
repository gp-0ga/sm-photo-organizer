import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { createAssetQr } from "./qr.js";
import { clearMarkerSelection, loadMarkerAssets, loadMarkerSelection, saveMarkerAssets, saveMarkerSelection } from "./marker-storage.js";

const excelInput = document.querySelector("#marker-excel-input");
const pdfInput = document.querySelector("#marker-pdf-input");
const excelStatus = document.querySelector("#marker-excel-status");
const viewPanel = document.querySelector("#marker-view-panel");
const printPanel = document.querySelector("#marker-print-panel");
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

let assets = [];
let currentIndex = 0;
const qrUrls = new Map();
let cardsRendered = false;

function decodePdfLiteral(value) {
  return value.replace(/\\([\\()nrt])/g, (_, code) => ({ "n": "\n", "r": "\r", "t": "\t" }[code] ?? code));
}

async function readMarkerAssetsFromPdf(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const chunks = [];
  const decoder = new TextDecoder("latin1");
  const raw = decoder.decode(bytes);
  const streamPattern = /stream\r?\n([\s\S]*?)\r?\nendstream/g;
  let match;
  while ((match = streamPattern.exec(raw))) {
    const streamBytes = bytes.slice(match.index + match[0].indexOf("\n") + 1, match.index + match[0].length - "endstream".length);
    try {
      const writer = new Blob([streamBytes]).stream().pipeThrough(new DecompressionStream("deflate"));
      const reader = writer.pipeThrough(new TextDecoderStream("latin1")).getReader();
      let text = "";
      for (;;) { const next = await reader.read(); if (next.done) break; text += next.value; }
      chunks.push(text);
    } catch {
      chunks.push(match[1]);
    }
  }
  const text = chunks.join("\n") + "\n" + raw;
  const assets = [];
  const seen = new Set();
  const pattern = /SM-ASSET-NAME\|1\|([^|()\\]+)\|([^()\\]*(?:\\.[^()\\]*)*)/g;
  while ((match = pattern.exec(text))) {
    const assetNumber = decodePdfLiteral(match[1]).trim();
    const assetName = decodePdfLiteral(match[2]).trim();
    if (assetNumber && assetName && !seen.has(assetNumber)) { seen.add(assetNumber); assets.push({ assetNumber, assetName, items: [] }); }
  }
  if (!assets.length) throw new Error("このPDFから資産情報を読み取れませんでした。PCアプリの「全マーカーを印刷」で作成したPDFを選択してください。");
  return assets;
}

function setStatus(message, tone = "neutral") {
  excelStatus.className = `status ${tone}`;
  excelStatus.textContent = message;
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

async function useAssets(nextAssets, message) {
  assets = nextAssets;
  qrUrls.clear();
  cardsRendered = false;
  cardList.replaceChildren();
  assetSelect.replaceChildren();
  for (const asset of assets) assetSelect.add(new Option(`${asset.assetNumber} ${asset.assetName}`, asset.assetNumber));
  viewPanel.hidden = false;
  printPanel.hidden = true;
  document.body.classList.add("marker-ready");
  const storedSelection = loadMarkerSelection(localStorage);
  const restoredIndex = Math.max(0, assets.findIndex((asset) => asset.assetNumber === storedSelection));
  await showAsset(restoredIndex);
  setStatus(message, "success");
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  setStatus(`${file.name}を読み込んでいます…`, "working");
  try {
    const loadedAssets = await readAssetsFromWorkbook(file);
    assets = loadedAssets;
    saveMarkerAssets(localStorage, loadedAssets);
    clearMarkerSelection(localStorage);
    await useAssets(loadedAssets, `${loadedAssets.length}資産をこのブラウザ内に保存しました。Excelは外部へ送信していません。`);
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
    setStatus(error.message ?? String(error), "error");
  }
});

assetSelect.addEventListener("change", () => {
  const index = assets.findIndex((asset) => asset.assetNumber === assetSelect.value);
  if (index >= 0) showAsset(index);
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
  setStatus("新しいExcelを選択してください。保存済み一覧は、新しい一覧を読み込むまで残ります。", "neutral");
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
