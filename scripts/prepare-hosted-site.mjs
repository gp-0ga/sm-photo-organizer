import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.join(projectRoot, "build");
await fs.rm(outputRoot, { recursive: true, force: true });
await fs.mkdir(outputRoot, { recursive: true });
await fs.copyFile(path.join(projectRoot, "dist", "capture.html"), path.join(outputRoot, "index.html"));
await fs.cp(path.join(projectRoot, "dist", "assets"), path.join(outputRoot, "assets"), { recursive: true });
console.log("Created hosted marker build/index.html with assets");
