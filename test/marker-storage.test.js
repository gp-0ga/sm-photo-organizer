import test from "node:test";
import assert from "node:assert/strict";
import { clearMarkerSelection, loadMarkerAssets, loadMarkerSelection, saveMarkerAssets, saveMarkerSelection } from "../src/marker-storage.js";

function memoryStorage() {
  const values = new Map();
  return {
    getItem: (key) => values.get(key) ?? null,
    setItem: (key, value) => values.set(key, String(value)),
    removeItem: (key) => values.delete(key),
  };
}

test("資産番号と資産名を端末内保存から復元する", () => {
  const storage = memoryStorage();
  const assets = [{ assetNumber: "12001_01", assetName: "ポンプ設備", items: ["保存しない項目"] }];
  saveMarkerAssets(storage, assets, new Date("2026-09-12T00:00:00Z"));
  const restored = loadMarkerAssets(storage);
  assert.equal(restored[0].assetNumber, "12001_01");
  assert.equal(restored[0].assetName, "ポンプ設備");
  assert.equal("items" in restored[0], false);
});

test("最後に選択した資産を保存して消去できる", () => {
  const storage = memoryStorage();
  saveMarkerSelection(storage, "12002_01");
  assert.equal(loadMarkerSelection(storage), "12002_01");
  clearMarkerSelection(storage);
  assert.equal(loadMarkerSelection(storage), null);
});

test("壊れた保存データは復元しない", () => {
  const storage = memoryStorage();
  storage.setItem("asset-marker.assets.v1", "not-json");
  assert.equal(loadMarkerAssets(storage), null);
});
