import test from "node:test";
import assert from "node:assert/strict";
import { cameraFileName, cropForZoom } from "../src/camera.js";

test("資産番号と撮影時刻から重複しにくいJPEG名を作る", () => {
  const date = new Date(2026, 8, 10, 9, 7, 5, 42);
  assert.equal(cameraFileName("01", date), "01_20260910_090705_042.jpg");
});

test("デジタルズームは中央部分を切り出す", () => {
  assert.deepEqual(cropForZoom(4000, 3000, 2), {
    x: 1000,
    y: 750,
    width: 2000,
    height: 1500,
  });
});
