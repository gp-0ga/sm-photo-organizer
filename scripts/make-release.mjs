import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const releaseRoot = path.join(projectRoot, "release", "資産写真整理MVP");
const scriptRoot = path.join(releaseRoot, "scripts");
const distRoot = path.join(releaseRoot, "dist");
await fs.mkdir(scriptRoot, { recursive: true });
await fs.mkdir(distRoot, { recursive: true });

await Promise.all([
  fs.copyFile(path.join(projectRoot, "start-mvp.bat"), path.join(releaseRoot, "start-mvp.bat")),
  fs.copyFile(path.join(projectRoot, "README.md"), path.join(releaseRoot, "README.md")),
  fs.copyFile(path.join(projectRoot, "dist", "index.html"), path.join(distRoot, "index.html")),
  fs.copyFile(path.join(projectRoot, "scripts", "serve-mvp.ps1"), path.join(scriptRoot, "serve-mvp.ps1")),
  fs.copyFile(path.join(projectRoot, "scripts", "build-photo-album.ps1"), path.join(scriptRoot, "build-photo-album.ps1")),
]);

console.log(`Created release at ${releaseRoot}`);
