import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import QRCode from "qrcode";

const output = new URL("../test/fixtures/photos/", import.meta.url);
await fs.mkdir(output, { recursive: true });

await QRCode.toFile(fileURLToPath(new URL("IMG_0001.png", output)), "SM-ASSET|1|12001_01", { errorCorrectionLevel: "H", width: 720, margin: 4 });
await QRCode.toFile(fileURLToPath(new URL("IMG_0004.png", output)), "SM-REVIEW|1", { errorCorrectionLevel: "H", width: 720, margin: 4 });
await QRCode.toFile(fileURLToPath(new URL("IMG_0005.png", output)), "SM-ASSET|1|12002_01", { errorCorrectionLevel: "H", width: 720, margin: 4 });

for (const [name, value] of [["IMG_0002.png", "TEST-PHOTO-1"], ["IMG_0003.png", "TEST-PHOTO-2"], ["IMG_0006.png", "TEST-PHOTO-3"]]) {
  await QRCode.toFile(fileURLToPath(new URL(name, output)), value, { errorCorrectionLevel: "H", width: 720, margin: 4 });
}

console.log(new URL(".", output).pathname);
