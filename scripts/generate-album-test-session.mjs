import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import QRCode from "qrcode";
import { readAssetsFromWorkbook } from "../src/excel.js";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sessionRoot = path.join(projectRoot, ".tmp_album_test");
const photoRoot = path.join(sessionRoot, "photos");
await fs.mkdir(photoRoot, { recursive: true });

const healthSource = "C:/Users/gpt0g/Desktop/102_健全度判定表_A_○○町_R08.xlsx";
const albumSource = "C:/Users/gpt0g/Desktop/202_写真帳_A_○○町_R08.xlsx";
await fs.copyFile(healthSource, path.join(sessionRoot, "health.xlsx"));
await fs.copyFile(albumSource, path.join(sessionRoot, "album.xlsx"));

const healthBytes = await fs.readFile(healthSource);
const assets = await readAssetsFromWorkbook(new File([healthBytes], "health.xlsx"));
const photos = [];
for (let index = 0; index < 5; index += 1) {
  const fileKey = String(index + 1).padStart(4, "0");
  await QRCode.toFile(path.join(photoRoot, `${fileKey}.jpg`), `PHOTO-ALBUM-TEST-${fileKey}`, { width: 900, margin: 3 });
  photos.push({
    fileKey,
    sourceName: `${fileKey}.jpg`,
    assetNumber: assets[0].assetNumber,
    destinationType: index === 0 ? "full" : "item",
    itemNumber: index === 0 ? null : assets[0].items[0].itemNumber,
    slotIndex: index === 0 ? 0 : index - 1,
  });
}

const manifest = {
  outputName: "photo-album-test.xlsx",
  assets: assets.map((asset) => ({ assetNumber: asset.assetNumber, itemNumbers: asset.items.map((item) => item.itemNumber) })),
  photos,
};
await fs.writeFile(path.join(sessionRoot, "manifest.json"), JSON.stringify(manifest, null, 2), "utf8");
console.log(sessionRoot);
