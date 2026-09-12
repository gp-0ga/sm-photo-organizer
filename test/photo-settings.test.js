import test from "node:test";
import assert from "node:assert/strict";
import { targetBytesFromKilobytes } from "../src/photos.js";

test("写真の指定容量をKBからバイトへ変換する", () => {
  assert.equal(targetBytesFromKilobytes(200), 204800);
  assert.equal(targetBytesFromKilobytes("750"), 768000);
});

test("写真容量の範囲外を拒否する", () => {
  assert.throws(() => targetBytesFromKilobytes(49), /50～5000KB/);
  assert.throws(() => targetBytesFromKilobytes(5001), /50～5000KB/);
});
