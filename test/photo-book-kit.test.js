import test from "node:test";
import assert from "node:assert/strict";
import { photoBookKitArchiveEntries } from "../src/photo-book-kit.js";

test("写真帳作成セットに健全度判定表と実行ファイルを含める", () => {
  const healthWorkbook = new Blob(["health"]);
  const entries = photoBookKitArchiveEntries({
    healthWorkbook,
    instruction: { kind: "asset-photo-album-instruction", version: 1 },
    guide: "guide",
    photoBookBat: "bat",
    photoBookRunner: "runner",
    photoBookBuilder: "builder",
  });

  assert.deepEqual(entries.map((entry) => entry.path), [
    "START-PHOTO-BOOK.bat",
    "photo-book-runner.ps1",
    "build-photo-album.ps1",
    "photo-book-instruction.json",
    "health.xlsx",
    "README.txt",
  ]);
  assert.equal(entries.find((entry) => entry.path === "health.xlsx").blob, healthWorkbook);
});

test("健全度判定表なしでは写真帳作成セットを作らない", () => {
  assert.throws(() => photoBookKitArchiveEntries({
    healthWorkbook: null,
    instruction: {},
    guide: "",
    photoBookBat: "",
    photoBookRunner: "",
    photoBookBuilder: "",
  }), /健全度判定表/);
});
