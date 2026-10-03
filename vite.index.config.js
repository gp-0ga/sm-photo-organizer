import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    outDir: "index-dist",
    emptyOutDir: true,
    rollupOptions: { input: { index: resolve(import.meta.dirname, "index.html") } },
  },
});
