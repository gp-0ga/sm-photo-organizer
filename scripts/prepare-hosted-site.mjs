import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.join(projectRoot, "build");
await fs.rm(outputRoot, { recursive: true, force: true });
await fs.mkdir(outputRoot, { recursive: true });
await fs.copyFile(path.join(projectRoot, "dist", "capture.html"), path.join(outputRoot, "index.html"));
await fs.copyFile(path.join(projectRoot, "dist", "index.html"), path.join(outputRoot, "organize.html"));
await fs.cp(path.join(projectRoot, "dist", "assets"), path.join(outputRoot, "assets"), { recursive: true });
for (const fileName of ["manifest.webmanifest", "icon.svg"]) {
  await fs.copyFile(path.join(projectRoot, "dist", fileName), path.join(outputRoot, fileName));
}
const assetFiles = await fs.readdir(path.join(outputRoot, "assets"));
const cacheFiles = ["./", "./index.html", "./organize.html", "./manifest.webmanifest", "./icon.svg", ...assetFiles.map((file) => `./assets/${file}`)];
const cacheName = `asset-marker-${Date.now()}`;
const serviceWorker = `const CACHE=${JSON.stringify(cacheName)};
const FILES=${JSON.stringify(cacheFiles)};
self.addEventListener("install",(event)=>{event.waitUntil(caches.open(CACHE).then((cache)=>cache.addAll(FILES)).then(()=>self.skipWaiting()));});
self.addEventListener("activate",(event)=>{event.waitUntil(caches.keys().then((keys)=>Promise.all(keys.filter((key)=>key!==CACHE).map((key)=>caches.delete(key)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",(event)=>{if(event.request.method!=="GET")return;event.respondWith(fetch(event.request).then((response)=>{const copy=response.clone();caches.open(CACHE).then((cache)=>cache.put(event.request,copy));return response;}).catch(async()=>await caches.match(event.request)||await caches.match("./index.html")));});
`;
await fs.writeFile(path.join(outputRoot, "sw.js"), serviceWorker);
const hosting = JSON.parse(await fs.readFile(path.join(projectRoot, ".openai", "hosting.json"), "utf8"));
hosting.static = { ...hosting.static, directory: "dist" };
await fs.mkdir(path.join(outputRoot, ".openai"), { recursive: true });
await fs.writeFile(path.join(outputRoot, ".openai", "hosting.json"), `${JSON.stringify(hosting, null, 2)}\n`);
console.log("Created hosted marker build with offline assets");
