import test from "node:test";
import assert from "node:assert/strict";
import { exportOrganizedPhotos } from "../src/folders.js";

class MemoryDirectory {
  constructor(name) {
    this.name = name;
    this.directories = new Map();
    this.files = new Map();
  }

  async getDirectoryHandle(name) {
    if (!this.directories.has(name)) this.directories.set(name, new MemoryDirectory(name));
    return this.directories.get(name);
  }

  async getFileHandle(name) {
    const directory = this;
    return {
      async createWritable() {
        return {
          async write(blob) { directory.files.set(name, blob); },
          async close() {},
        };
      },
    };
  }
}

test("全写真を資産直下へ保存し、採用写真だけサブフォルダへ複写する", async () => {
  const root = new MemoryDirectory("root");
  global.window = { showDirectoryPicker: async () => root };
  const assets = [{
    assetNumber: "12001_01",
    folderName: "12001_01_外装壁",
    items: [{ folderName: "12001_01_1_損傷" }],
  }];
  const photos = [
    { file: { name: "IMG_0001.HEIC" }, assetNumber: "12001_01", destination: "全景", excluded: false, reviewRequired: false, qrReadError: false },
    { file: { name: "IMG_0002.JPG" }, assetNumber: "12001_01", destination: "", excluded: false, reviewRequired: false, qrReadError: false },
    { file: { name: "IMG_0003.JPG" }, assetNumber: "12001_01", destination: "", excluded: true, reviewRequired: false, qrReadError: false },
  ];
  const result = await exportOrganizedPhotos(
    assets,
    photos,
    async () => new Blob(["jpeg"], { type: "image/jpeg" }),
    (name) => name.replace(/\.[^.]+$/, ".jpg"),
  );
  assert.equal(result.photoCount, 2);
  const output = [...root.directories.values()][0];
  const asset = output.directories.get("12001_01_外装壁");
  assert.deepEqual([...asset.files.keys()], ["IMG_0001.jpg", "IMG_0002.jpg"]);
  assert.deepEqual([...asset.directories.get("全景").files.keys()], ["IMG_0001.jpg"]);
  assert.equal(asset.directories.get("12001_01_1_損傷").files.size, 0);
});
