import test from "node:test";
import assert from "node:assert/strict";
import { createAssetPhotoZip, zipFileName } from "../src/zip.js";

test("撮影写真を資産フォルダ名入りZIPにする", async () => {
  const file = new File([new Uint8Array([1, 2, 3])], "01_test.jpg", { type: "image/jpeg", lastModified: 0 });
  const result = await createAssetPhotoZip(
    [{ assetNumber: "01", folderName: "01_橋梁" }],
    [{ source: "camera", assetNumber: "01", excluded: false, file }],
  );
  const bytes = new Uint8Array(await result.blob.arrayBuffer());
  const text = new TextDecoder().decode(bytes);
  assert.equal(result.photoCount, 1);
  assert.equal(new DataView(bytes.buffer).getUint32(0, true), 0x04034b50);
  assert.match(text, /01_橋梁\/01_test\.jpg/);
});

test("ZIP名に作成日時を入れる", () => {
  assert.equal(zipFileName(new Date(2026, 8, 10, 9, 7)), "資産写真_20260910_0907.zip");
});
