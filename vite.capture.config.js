import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    outDir: "capture-dist",
    emptyOutDir: true,
    rollupOptions: { input: { capture: resolve(import.meta.dirname, "capture.html") } },
  },
});
