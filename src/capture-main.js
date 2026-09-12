import "./styles.css";
import { readAssetsFromWorkbook } from "./excel.js";
import { createAssetQr } from "./qr.js";

const excelInput = document.querySelector("#marker-excel-input");
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

let assets = [];
let currentIndex = 0;
const qrUrls = new Map();

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
  assetSelect.value = asset.assetNumber;
  displayQr.src = await qrFor(asset);
  displayQr.alt = `${asset.assetNumber} ${asset.assetName} 資産切替QR`;
  displayNumber.textContent = asset.assetNumber;
  displayName.textContent = asset.assetName;
  position.textContent = `${currentIndex + 1} / ${assets.length}`;
  previousButton.disabled = currentIndex === 0;
  nextButton.disabled = currentIndex === assets.length - 1;
  document.querySelectorAll(".marker-card.selected").forEach((card) => card.classList.remove("selected"));
  document.querySelector(`[data-asset-number="${CSS.escape(asset.assetNumber)}"]`)?.classList.add("selected");
}

async function renderMarkerCards() {
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
    card.append(start, image, number, name);
    card.addEventListener("click", () => {
      showAsset(index);
      viewPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    cardList.append(card);
  }
}

excelInput.addEventListener("change", async () => {
  const file = excelInput.files?.[0];
  if (!file) return;
  viewPanel.hidden = true;
  printPanel.hidden = true;
  setStatus(`${file.name}を読み込んでいます…`, "working");
  try {
    assets = await readAssetsFromWorkbook(file);
    qrUrls.clear();
    assetSelect.replaceChildren();
    for (const asset of assets) assetSelect.add(new Option(`${asset.assetNumber} ${asset.assetName}`, asset.assetNumber));
    await renderMarkerCards();
    viewPanel.hidden = false;
    printPanel.hidden = false;
    await showAsset(0);
    setStatus(`${assets.length}資産のマーカーを作成しました。Excelは外部へ送信していません。`, "success");
  } catch (error) {
    assets = [];
    cardList.replaceChildren();
    setStatus(error.message ?? String(error), "error");
  }
});

assetSelect.addEventListener("change", () => {
  const index = assets.findIndex((asset) => asset.assetNumber === assetSelect.value);
  if (index >= 0) showAsset(index);
});
previousButton.addEventListener("click", () => showAsset(currentIndex - 1));
nextButton.addEventListener("click", () => showAsset(currentIndex + 1));
printButton.addEventListener("click", () => window.print());
