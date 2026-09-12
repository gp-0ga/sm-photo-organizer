import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        index: resolve(import.meta.dirname, "index.html"),
        capture: resolve(import.meta.dirname, "capture.html"),
      },
    },
  },
});
