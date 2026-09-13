import { defineConfig } from "vite";
import { resolve } from "node:path";
import { randomUUID } from "node:crypto";
import { execFile } from "node:child_process";
import { mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

function execFileAsync(command, args) {
  return new Promise((resolvePromise, reject) => {
    execFile(command, args, { windowsHide: true, maxBuffer: 20 * 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        error.message = `${error.message}\n${stderr}`;
        reject(error);
        return;
      }
      resolvePromise({ stdout, stderr });
    });
  });
}

function popplerExe(name) {
  const bundled = process.env.USERPROFILE
    ? join(process.env.USERPROFILE, ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "native", "poppler", "Library", "bin", `${name}.exe`)
    : name;
  return process.platform === "win32" ? bundled : name;
}

function readRequestBody(req) {
  return new Promise((resolvePromise, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolvePromise(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function sendJson(res, statusCode, body) {
  res.statusCode = statusCode;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(body));
}

function pdfRenderPlugin() {
  return {
    name: "local-pdf-render",
    configureServer(server) {
      server.middlewares.use("/api/render-pdf", async (req, res) => {
        if (req.method !== "POST") {
          sendJson(res, 405, { error: "POSTのみ対応しています。" });
          return;
        }
        const workDir = join(tmpdir(), `drawing-pdf-${randomUUID()}`);
        try {
          const dpi = Math.min(Math.max(Number(new URL(req.url, "http://localhost").searchParams.get("dpi")) || 160, 72), 300);
          await mkdir(workDir, { recursive: true });
          const inputPath = join(workDir, "input.pdf");
          const outputPrefix = join(workDir, "page");
          await writeFile(inputPath, await readRequestBody(req));
          const info = await execFileAsync(popplerExe("pdfinfo"), [inputPath]);
          const totalPages = Number(info.stdout.match(/^Pages:\s+(\d+)/m)?.[1]) || 1;
          const pageLimit = Math.min(totalPages, 30);
          await execFileAsync(popplerExe("pdftoppm"), ["-png", "-r", String(dpi), "-f", "1", "-l", String(pageLimit), inputPath, outputPrefix]);
          const files = (await readdir(workDir))
            .filter((name) => /^page-\d+\.png$/.test(name))
            .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
          const pages = await Promise.all(
            files.map(async (name) => ({
              pageNumber: Number(name.match(/\d+/)[0]),
              dataUrl: `data:image/png;base64,${(await readFile(join(workDir, name))).toString("base64")}`,
            })),
          );
          sendJson(res, 200, { pages, totalPages, renderedPages: pages.length, truncated: totalPages > pageLimit });
        } catch (error) {
          sendJson(res, 500, { error: error.message });
        } finally {
          await rm(workDir, { recursive: true, force: true });
        }
      });
    },
  };
}

export default defineConfig({
  plugins: [pdfRenderPlugin()],
  server: {
    host: "127.0.0.1",
    port: 5173,
  },
  build: {
    rollupOptions: {
      input: {
        index: resolve(import.meta.dirname, "index.html"),
        capture: resolve(import.meta.dirname, "capture.html"),
        drawing: resolve(import.meta.dirname, "drawing.html"),
      },
    },
  },
});
