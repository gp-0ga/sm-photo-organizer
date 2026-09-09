import test from "node:test";
import assert from "node:assert/strict";
import { classifyDecodedEntries, parseAssetsFromRows, parseMarkerPayload, markerPayload, sanitizeFolderSegment } from "../src/domain.js";

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
