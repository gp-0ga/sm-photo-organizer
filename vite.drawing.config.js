import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    outDir: "drawing-dist",
    emptyOutDir: true,
    rollupOptions: {
      input: { drawing: resolve(import.meta.dirname, "drawing.html") },
    },
  },
});
