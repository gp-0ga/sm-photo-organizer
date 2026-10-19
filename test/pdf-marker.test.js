import test from "node:test";
import assert from "node:assert/strict";
import { markerAssetPrintPayload, parseMarkerAssetsFromTextItems } from "../src/pdf-marker.js";

test("印刷PDFのマーカー文字から資産情報を読み取る", () => {
  assert.deepEqual(parseMarkerAssetsFromTextItems([
    "ここからこの資産",
    "SM-ASSET-NAME|1|1200070_1_1|機械設備",
    "SM-ASSET-NAME | 1 | 1200080_1_1 | 電気設備",
  ]), [
    { assetNumber: "1200070_1_1", assetName: "機械設備", items: [] },
    { assetNumber: "1200080_1_1", assetName: "電気設備", items: [] },
  ]);
});

test("PDFで分割されたマーカー文字を結合して読み取る", () => {
  assert.deepEqual(parseMarkerAssetsFromTextItems([
    "SM-ASSET-NAME", "|", "1", "|", "1200070_1_1", "|", "機械設備",
  ]), [
    { assetNumber: "1200070_1_1", assetName: "機械設備", items: [] },
  ]);
});

test("新様式の完全な資産名称をPDFで分割されても復元する", () => {
  const payload = markerAssetPrintPayload("1200070", "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室");
  const pieces = payload.match(/.{1,12}/g);
  assert.deepEqual(parseMarkerAssetsFromTextItems(pieces), [
    { assetNumber: "1200070", assetName: "内装床_磁器質ﾀｲﾙ_管理棟_1F_風除室", items: [] },
  ]);
});
