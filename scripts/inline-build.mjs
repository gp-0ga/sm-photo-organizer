import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const distRoot = path.join(projectRoot, "dist");
const indexPath = path.join(distRoot, "index.html");

let html = await fs.readFile(indexPath, "utf8");

for (const match of [...html.matchAll(/<link rel="stylesheet" crossorigin href="([^"]+)">/g)]) {
  const filePath = path.join(distRoot, match[1].replace(/^\//, ""));
  const css = await fs.readFile(filePath, "utf8");
  html = html.replace(match[0], `<style>${css}</style>`);
}

for (const match of [...html.matchAll(/<script type="module" crossorigin src="([^"]+)"><\/script>/g)]) {
  const filePath = path.join(distRoot, match[1].replace(/^\//, ""));
  const script = await fs.readFile(filePath, "utf8");
  const encoded = Buffer.from(script, "utf8").toString("base64");
  const loader = `eval(new TextDecoder().decode(Uint8Array.from(atob("${encoded}"), character => character.charCodeAt(0))))`;
  html = html.replace(match[0], () => `<script type="module">${loader}<\/script>`);
}

await fs.writeFile(indexPath, html, "utf8");
await fs.rm(path.join(distRoot, "assets"), { recursive: true, force: true });

console.log("Created standalone dist/index.html");
