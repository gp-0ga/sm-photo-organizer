import test from "node:test";
import assert from "node:assert/strict";
import { albumEntries } from "../src/album.js";

const assets = [{
  assetNumber: "12001_01",
  assetName: "外装壁",
  items: [{ itemNumber: "12001_01_1", folderName: "12001_01_1_損傷" }],
}];

const photo = (name, destination) => ({ file: { name }, assetNumber: "12001_01", destination, excluded: false });

test("全景1枚と確認項目4枚を写真帳配置へ変換する", () => {
  const entries = albumEntries(assets, [
    photo("full.jpg", "全景"),
    photo("1.jpg", "12001_01_1_損傷"),
    photo("2.jpg", "12001_01_1_損傷"),
    photo("3.jpg", "12001_01_1_損傷"),
    photo("4.jpg", "12001_01_1_損傷"),
  ]);
  assert.deepEqual(entries.map(({ destinationType, slotIndex }) => [destinationType, slotIndex]), [
    ["full", 0], ["item", 0], ["item", 1], ["item", 2], ["item", 3],
  ]);
});

test("全景2枚は拒否する", () => {
  assert.throws(() => albumEntries(assets, [photo("1.jpg", "全景"), photo("2.jpg", "全景")]), /最大1枚/);
});

test("確認項目5枚は拒否する", () => {
  assert.throws(() => albumEntries(assets, Array.from({ length: 5 }, (_, index) => photo(`${index}.jpg`, "12001_01_1_損傷"))), /最大4枚/);
});
