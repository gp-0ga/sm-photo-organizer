import test from "node:test";
import assert from "node:assert/strict";
import { classifyDecodedEntries, matchInspectionAssets, parseAssetsFromAssetListRows, parseAssetsFromRows, parseInspectionSheetRows, parseMarkerPayload, markerPayload, sanitizeFolderSegment, OTHER_ASSET_NUMBER } from "../src/domain.js";

test("No11の行から資産を重複排除し、経過年数を除外する", () => {
  const rows = [
    ["", "資産番号", "項目番号", "資産名称", "点検項目"],
    ["", "12001_01", "12001_01_1", "外装壁_管理棟", "損傷"],
    ["", "12001_01", "12001_01_2", "外装壁_管理棟", "経過年数"],
    ["", "12002_01", "12002_01_1", "躯体_管理棟", "ひび割れ"],
  ];
  const assets = parseAssetsFromRows(rows);
  assert.equal(assets.length, 2);
  assert.equal(assets[0].folderName, "12001_01_外装壁_管理棟");
  assert.deepEqual(assets[0].items.map((item) => item.folderName), ["12001_01_1_損傷"]);
});

test("Windowsで使えない文字を置換する", () => {
  assert.equal(sanitizeFolderSegment('A/B:C*D?'), "A_B_C_D_");
});

test("資産リストの改行入り見出しから資産を読み取る", () => {
  const rows = [
    ["", "リスト\nNO", "", "", "", "", "", "資産名称", "", "", "", "", "", "", "", "", "調査対象"],
    ["", 1200070, "", "", "", "", "", "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室", "", "", "", "", "", "", "", "", "○"],
  ];
  const assets = parseAssetsFromAssetListRows(rows);
  assert.equal(assets[0].assetNumber, "1200070");
  assert.equal(assets[0].assetName, "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室");
  assert.equal(assets[0].inspectionTarget, true);
});

test("点検表を資産番号で資産リストに対応付ける", () => {
  const rows = [
    ["リスト番号", "", "資産名称"],
    [1200070, "", "内装床"],
  ];
  const inspection = parseInspectionSheetRows(rows, "1200070.xlsx");
  const result = matchInspectionAssets([
    { assetNumber: "1200070", assetName: "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室", items: [] },
    { assetNumber: "1200080", assetName: "内装天井_管理棟", items: [] },
  ], [inspection, inspection, { assetNumber: "9999990", assetName: "対象外", sourceName: "extra.xlsx" }]);
  assert.deepEqual(result.assets.map((asset) => asset.assetNumber), ["1200070"]);
  assert.equal(result.assets[0].assetName, "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室");
  assert.equal(result.duplicates.length, 1);
  assert.equal(result.unmatched.length, 1);
  assert.equal(result.nameMismatches.length, 1);
});

test("資産マーカーを往復変換する", () => {
  const payload = markerPayload("12001_01");
  assert.deepEqual(parseMarkerPayload(payload), { type: "asset", assetNumber: "12001_01" });
});

test("要確認マーカーは直前区間だけを要確認にする", () => {
  const fake = (name) => ({ name });
  const result = classifyDecodedEntries([
    { id: "m1", file: fake("marker.jpg"), decodedValue: markerPayload("12001_01") },
    { id: "p1", file: fake("1.jpg"), decodedValue: null },
    { id: "p2", file: fake("2.jpg"), decodedValue: null },
    { id: "r1", file: fake("review.jpg"), decodedValue: "SM-REVIEW|1" },
    { id: "m2", file: fake("marker2.jpg"), decodedValue: markerPayload("12002_01") },
    { id: "p3", file: fake("3.jpg"), decodedValue: null },
  ], ["12001_01", "12002_01"]);
  assert.equal(result.photos.length, 3);
  assert.deepEqual(result.photos.map((photo) => photo.reviewRequired), [true, true, false]);
  assert.deepEqual(result.photos.map((photo) => photo.assetNumber), ["12001_01", "12001_01", "12002_01"]);
});

test("その他マーカーは資産不明写真をその他へ分類する", () => {
  const fake = (name) => ({ name });
  const result = classifyDecodedEntries([
    { id: "o1", file: fake("other.jpg"), decodedValue: "SM-OTHER|1" },
    { id: "p1", file: fake("1.jpg"), decodedValue: null },
  ], ["12001_01"]);
  assert.deepEqual(parseMarkerPayload("SM-OTHER|1"), { type: "other" });
  assert.equal(result.photos[0].assetNumber, OTHER_ASSET_NUMBER);
});
