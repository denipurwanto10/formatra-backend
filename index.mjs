// Formatra — conversion backend
// ----------------------------------------------------------------------------
// A small HTTP server that hands high-fidelity document conversions off to
// real rendering/parsing engines instead of doing them in the browser. It
// only handles three routes:
//
//   POST /api/convert/word-to-pdf   (.doc/.docx -> .pdf)   [LibreOffice]
//   POST /api/convert/ppt-to-pdf    (.ppt/.pptx -> .pdf)   [LibreOffice]
//   POST /api/convert/pdf-to-word   (.pdf       -> .docx)  [pdf2docx, falls
//                                                            back to
//                                                            LibreOffice]
//
// Design notes:
//  - Uses only Node's built-in `http` for everything except PDF -> Word,
//    which shells out to a small Python script (pdf_to_docx.py) using the
//    `pdf2docx` library. `node server/index.mjs` is enough to run the
//    LibreOffice-only routes (Node 18+); PDF -> Word additionally needs
//    Python 3 with `pdf2docx` installed (see server/README.md).
//  - Each LibreOffice request gets its own temp working directory AND its
//    own isolated LibreOffice user profile (`-env:UserInstallation=...`).
//    This is required for correctness, not just cleanliness: LibreOffice
//    headless treats the profile directory as a single-writer lock, so two
//    conversions sharing a profile can corrupt or block each other. A shared
//    profile is the most common cause of "works once then hangs"
//    LibreOffice-headless bugs.
//  - A small semaphore caps how many conversion child processes (soffice OR
//    python/pdf2docx) run at once (MAX_CONCURRENT_JOBS) so a burst of
//    uploads can't fork-bomb the host.
//  - Every temp file/dir is removed in a `finally` block, so failed/timed out
//    jobs don't leak disk space.
//  - This process trusts nothing from the client except the file bytes: the
//    client-supplied filename is only ever used to read an extension against
//    a whitelist, never as a filesystem path.
//  - PDF -> Word engine choice: `pdf2docx` reconstructs the PDF as a real
//    layout tree (positioned text blocks, table grids, images) and rebuilds
//    the .docx from that structure, which keeps position/spacing/alignment/
//    columns/tables much closer to the source than LibreOffice's
//    writer_pdf_import filter (which re-flows the PDF into ordinary Writer
//    paragraphs). If pdf2docx isn't available or fails on a given file, this
//    route automatically falls back to the original LibreOffice path for
//    that request — Word -> PDF and PPT -> PDF are untouched either way.
// ----------------------------------------------------------------------------

import http from "node:http";
import { spawn } from "node:child_process";
import { mkdtemp, mkdir, rm, readFile, stat } from "node:fs/promises";
import { existsSync, createWriteStream } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { pathToFileURL, fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ---- Configuration (override via environment variables) --------------------
const PORT = Number(process.env.PORT || 8787);
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";
// 50MB was tight for real-world decks/scans with lots of embedded images —
// raised to a more realistic default; still fully overridable via env.
const MAX_FILE_SIZE_MB = Number(process.env.MAX_FILE_SIZE_MB || 150);
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;
// Base timeout for small files, plus extra time scaled to the upload size —
// a 90s flat cap was routinely too short for large/complex documents (many
// pages, big embedded images, heavy tables), causing conversions that were
// actually still progressing to be killed and reported as failures.
const SOFFICE_TIMEOUT_BASE_MS = Number(process.env.SOFFICE_TIMEOUT_MS || 90_000);
const SOFFICE_TIMEOUT_PER_MB_MS = Number(process.env.SOFFICE_TIMEOUT_PER_MB_MS || 2_500);
const SOFFICE_TIMEOUT_MAX_MS = Number(process.env.SOFFICE_TIMEOUT_MAX_MS || 600_000);
const MAX_CONCURRENT_JOBS = Number(process.env.MAX_CONCURRENT_JOBS || 2);

/** Scales the conversion timeout with input size so large files get
 *  proportionally more time instead of hitting a flat cap tuned for small
 *  ones. Clamped so a single pathological upload can't hang a slot forever. */
function timeoutForSize(sizeBytes) {
  const sizeMb = sizeBytes / (1024 * 1024);
  const scaled = SOFFICE_TIMEOUT_BASE_MS + sizeMb * SOFFICE_TIMEOUT_PER_MB_MS;
  return Math.min(SOFFICE_TIMEOUT_MAX_MS, Math.max(SOFFICE_TIMEOUT_BASE_MS, scaled));
}

// Default to the binary name on PATH (correct for the documented Docker/Linux
// deployment, where `apt-get install libreoffice` puts `soffice` on PATH).
// A hardcoded Windows install path here previously broke Word->PDF/PPT->PDF
// on every non-Windows deployment unless SOFFICE_BIN was manually overridden
// — silently contradicting the README's documented default of "soffice".
// Override via the SOFFICE_BIN env var to force a specific path, e.g.
// SOFFICE_BIN="C:\Program Files\LibreOffice\program\soffice.exe".
//
// On Windows specifically, the official LibreOffice installer does NOT add
// soffice.exe to PATH, so plain "soffice" reliably fails with
// "spawn soffice ENOENT" unless the user manually sets SOFFICE_BIN every
// time. To avoid that trap, if no SOFFICE_BIN is set and we're on win32, we
// probe the handful of paths the installer actually uses and pick the first
// one that exists — silently falling back to plain "soffice" (and letting
// the normal ENOENT error explain what to do) if none of them do either.
function resolveSofficeBin() {
  if (process.env.SOFFICE_BIN) return process.env.SOFFICE_BIN;
  if (process.platform !== "win32") return "soffice";

  const candidates = [
    path.join(process.env["ProgramFiles"] || "C:\\Program Files", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["LOCALAPPDATA"] || "", "Programs", "LibreOffice", "program", "soffice.exe"),
  ];
  const found = candidates.find((p) => p && existsSync(p));
  return found || "soffice";
}
const SOFFICE_BIN = resolveSofficeBin();

// PDF -> Word engine. "pdf2docx" (default) uses the layout-preserving Python
// converter with automatic fallback to LibreOffice on failure; set
// PDF_TO_WORD_ENGINE=libreoffice to disable pdf2docx entirely and restore the
// exact previous behavior for this route.
const PDF2DOCX_ENABLED = process.env.PDF_TO_WORD_ENGINE !== "libreoffice";
const PYTHON_BIN = process.env.PYTHON_BIN || "python3";
const PDF2DOCX_SCRIPT = path.join(__dirname, "pdf_to_docx.py");
const PDF2DOCX_TIMEOUT_BASE_MS = Number(process.env.PDF2DOCX_TIMEOUT_MS || 120_000);
const PDF2DOCX_TIMEOUT_PER_MB_MS = Number(process.env.PDF2DOCX_TIMEOUT_PER_MB_MS || 4_000);
const PDF2DOCX_TIMEOUT_MAX_MS = Number(process.env.PDF2DOCX_TIMEOUT_MAX_MS || 600_000);
function pdf2docxTimeoutForSize(sizeBytes) {
  const sizeMb = sizeBytes / (1024 * 1024);
  const scaled = PDF2DOCX_TIMEOUT_BASE_MS + sizeMb * PDF2DOCX_TIMEOUT_PER_MB_MS;
  return Math.min(PDF2DOCX_TIMEOUT_MAX_MS, Math.max(PDF2DOCX_TIMEOUT_BASE_MS, scaled));
}

const ROUTES = {
  "/api/convert/word-to-pdf": {
    accept: [".doc", ".docx", ".rtf", ".odt"],
    outExt: ".pdf",
    outMime: "application/pdf",
    engine: "libreoffice",
    args: (inputPath, outDir) => ["--convert-to", "pdf", "--outdir", outDir, inputPath],
  },
  "/api/convert/ppt-to-pdf": {
    accept: [".ppt", ".pptx", ".odp"],
    outExt: ".pdf",
    outMime: "application/pdf",
    engine: "libreoffice",
    args: (inputPath, outDir) => ["--convert-to", "pdf", "--outdir", outDir, inputPath],
  },
  "/api/convert/pdf-to-word": {
    accept: [".pdf"],
    outExt: ".docx",
    outMime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    // Primary engine is pdf2docx (see runConversion). These LibreOffice args
    // are kept as the automatic fallback path for this same route, used only
    // if pdf2docx is disabled, missing, or fails on a particular file.
    engine: "pdf2docx",
    // Explicit import filter: without it, some LibreOffice builds try to
    // guess the filter from content and occasionally misdetect a PDF as a
    // generic "text" import, dropping layout entirely.
    args: (inputPath, outDir) => [
      "--infilter=writer_pdf_import",
      "--convert-to",
      "docx:MS Word 2007 XML",
      "--outdir",
      outDir,
      inputPath,
    ],
  },
};

// ---- Tiny concurrency semaphore --------------------------------------------
let active = 0;
const queue = [];
function acquireSlot() {
  if (active < MAX_CONCURRENT_JOBS) {
    active += 1;
    return Promise.resolve();
  }
  return new Promise((resolve) => queue.push(resolve));
}
function releaseSlot() {
  active = Math.max(0, active - 1);
  const next = queue.shift();
  if (next) {
    active += 1;
    next();
  }
}

// ---- Helpers ----------------------------------------------------------------
function sendJson(res, status, body) {
  const data = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(data),
  });
  res.end(data);
}

function applyCors(res) {
  res.setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Filename");
}

function safeExt(filename) {
  const ext = path.extname(String(filename || "")).toLowerCase();
  return /^\.[a-z0-9]{2,5}$/.test(ext) ? ext : "";
}

/** Streams the request body straight to disk (instead of buffering the whole
 *  upload in memory first) so large PDFs/PPTX/DOCX don't multiply the
 *  server's memory footprint under concurrent uploads. Still aborts early
 *  if the stream exceeds the size cap, and cleans up the partial file. */
function streamBodyToFile(req, destPath) {
  return new Promise((resolve, reject) => {
    let total = 0;
    let settled = false;
    const out = createWriteStream(destPath);

    const fail = (err) => {
      if (settled) return;
      settled = true;
      // Stop feeding the write stream and drop it, but deliberately do NOT
      // destroy `req` here — in Node's http server, req and res share the
      // same underlying socket, and destroying req tears the socket down
      // immediately. That was silently turning "file too large" into a
      // broken connection with no JSON error ever reaching the client
      // instead of the intended 413 response. The caller sends the JSON
      // error over that same still-open socket; the request is fully
      // drained/closed afterward instead.
      try {
        req.unpipe(out);
      } catch {
        /* already unpiped */
      }
      out.destroy();
      reject(err);
    };

    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > MAX_FILE_SIZE_BYTES) {
        fail(
          Object.assign(new Error(`File exceeds the ${MAX_FILE_SIZE_MB}MB limit`), { statusCode: 413 })
        );
      }
    });
    req.on("error", (err) => fail(Object.assign(err, { statusCode: 400 })));
    out.on("error", (err) => fail(Object.assign(err, { statusCode: 500 })));
    out.on("finish", () => {
      if (settled) return;
      settled = true;
      resolve(total);
    });

    req.pipe(out);
  });
}

/** Runs soffice with an isolated profile + timeout. Resolves with the produced file's path.
 *  Unchanged from the original implementation — this is still exactly what
 *  word-to-pdf and ppt-to-pdf use, and what pdf-to-word falls back to. */
async function runLibreOfficeConversion(route, inputPath, workDir, timeoutMs) {
  const outDir = path.join(workDir, "out");
  await mkdir(outDir, { recursive: true });
  const profileDir = path.join(workDir, "profile");

  const args = [
    `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
    "--headless",
    "--nologo",
    "--nofirststartwizard",
    "--nodefault",
    "--norestore",
    ...route.args(inputPath, outDir),
  ];

  const env = { ...process.env, SAL_USE_VCLPLUGIN: "svp" };

  await new Promise((resolve, reject) => {
    const child = spawn(SOFFICE_BIN, args, { env });
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(
        Object.assign(new Error("Conversion timed out (file may be too large or complex)"), {
          statusCode: 504,
        })
      );
    }, timeoutMs);

    child.stderr?.on("data", (d) => {
      stderr += d.toString();
    });
    child.on("error", (err) => {
      clearTimeout(timer);
      if (err.code === "ENOENT") {
        reject(
          Object.assign(
            new Error(
              `LibreOffice (soffice) tidak ditemukan di "${SOFFICE_BIN}". Pastikan LibreOffice terpasang, lalu set ` +
                `variabel lingkungan SOFFICE_BIN ke path lengkap soffice.exe/soffice sebelum menjalankan server, ` +
                `misalnya (Windows, PowerShell): $env:SOFFICE_BIN="C:\\Program Files\\LibreOffice\\program\\soffice.exe"`
            ),
            { statusCode: 500 }
          )
        );
        return;
      }
      reject(Object.assign(err, { statusCode: 500 }));
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) return resolve();
      // A few failure modes are common enough (and confusing enough from a
      // raw soffice exit code) to worth a plain-language message instead of
      // just surfacing the raw stderr tail.
      const tail = stderr.slice(-500);
      if (/password|encrypted|wrong password/i.test(tail)) {
        return reject(
          Object.assign(new Error("File dilindungi kata sandi. Hapus proteksi terlebih dahulu sebelum dikonversi."), {
            statusCode: 422,
          })
        );
      }
      reject(
        Object.assign(new Error(`soffice exited with code ${code}: ${tail}`), {
          statusCode: 500,
        })
      );
    });
  });

  const producedName = path.basename(inputPath, path.extname(inputPath)) + route.outExt;
  const producedPath = path.join(outDir, producedName);
  if (!existsSync(producedPath)) {
    throw Object.assign(new Error("Conversion finished but no output file was produced"), {
      statusCode: 500,
    });
  }
  return producedPath;
}

/** Runs the pdf2docx Python script with a timeout. Resolves with the produced file's path. */
async function runPdf2DocxConversion(inputPath, workDir, timeoutMs) {
  const outDir = path.join(workDir, "out");
  await mkdir(outDir, { recursive: true });
  const outputPath = path.join(outDir, path.basename(inputPath, path.extname(inputPath)) + ".docx");

  await new Promise((resolve, reject) => {
    const child = spawn(PYTHON_BIN, [PDF2DOCX_SCRIPT, inputPath, outputPath]);
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(
        Object.assign(new Error("pdf2docx conversion timed out (file may be too large or complex)"), {
          statusCode: 504,
        })
      );
    }, timeoutMs);

    child.stderr?.on("data", (d) => {
      stderr += d.toString();
    });
    child.on("error", (err) => {
      // e.g. ENOENT if python3 isn't installed at all
      clearTimeout(timer);
      reject(Object.assign(err, { statusCode: 500 }));
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else
        reject(
          Object.assign(new Error(`pdf2docx exited with code ${code}: ${stderr.slice(-800)}`), {
            statusCode: 500,
          })
        );
    });
  });

  if (!existsSync(outputPath)) {
    throw Object.assign(new Error("pdf2docx finished but no output file was produced"), {
      statusCode: 500,
    });
  }
  return outputPath;
}

/** Picks the engine for a route. Only pdf-to-word has more than one option;
 *  word-to-pdf and ppt-to-pdf always take the exact same LibreOffice path
 *  they always have. */
async function runConversion(route, inputPath, workDir, sizeBytes) {
  if (route.engine === "pdf2docx") {
    if (PDF2DOCX_ENABLED) {
      try {
        return await runPdf2DocxConversion(inputPath, workDir, pdf2docxTimeoutForSize(sizeBytes));
      } catch (err) {
        console.error(
          `[formatra-server] pdf2docx failed (${err.message}); falling back to LibreOffice for this request`
        );
      }
    }
    return runLibreOfficeConversion(route, inputPath, workDir, timeoutForSize(sizeBytes));
  }
  return runLibreOfficeConversion(route, inputPath, workDir, timeoutForSize(sizeBytes));
}

async function handleConvert(req, res, route) {
  const filenameHeader = req.headers["x-filename"];
  let originalName = "input";
  try {
    originalName = decodeURIComponent(String(filenameHeader || "input"));
  } catch {
    /* keep default */
  }
  const ext = safeExt(originalName);
  if (!ext || !route.accept.includes(ext)) {
    return sendJson(res, 415, {
      error: `Unsupported file type${ext ? ` "${ext}"` : ""}. Expected one of: ${route.accept.join(", ")}`,
    });
  }

  const workDir = await mkdtemp(path.join(os.tmpdir(), "pdftk-convert-"));
  const inputPath = path.join(workDir, `input-${crypto.randomUUID()}${ext}`);

  let inputSize = 0;
  try {
    // Streamed straight to disk rather than buffered fully in memory first —
    // keeps concurrent large-file uploads from multiplying RSS.
    inputSize = await streamBodyToFile(req, inputPath);
  } catch (err) {
    // No concurrency slot acquired yet at this point, so nothing to release.
    rm(workDir, { recursive: true, force: true }).catch(() => {});
    sendJson(res, err.statusCode || 400, { error: err.message });
    // Now that the response is flushed, stop the client from continuing to
    // stream a (likely still-large) request body we're no longer reading.
    if (!req.destroyed) req.destroy();
    return;
  }
  if (!inputSize) {
    rm(workDir, { recursive: true, force: true }).catch(() => {});
    return sendJson(res, 400, { error: "Empty file body" });
  }

  await acquireSlot();
  try {
    const outputPath = await runConversion(route, inputPath, workDir, inputSize);
    const outputBuffer = await readFile(outputPath);
    const outStat = await stat(outputPath);
    res.writeHead(200, {
      "Content-Type": route.outMime,
      "Content-Length": outStat.size,
      "Content-Disposition": `attachment; filename="converted${route.outExt}"`,
    });
    res.end(outputBuffer);
  } catch (err) {
    sendJson(res, err.statusCode || 500, { error: err.message || "Conversion failed" });
  } finally {
    releaseSlot();
    // Always clean up, even if the response above already failed.
    rm(workDir, { recursive: true, force: true }).catch(() => {});
  }
}

const server = http.createServer(async (req, res) => {
  applyCors(res);

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    return res.end();
  }

  if (req.method === "GET" && req.url === "/api/health") {
    return sendJson(res, 200, {
      ok: true,
      maxFileSizeMB: MAX_FILE_SIZE_MB,
      maxConcurrentJobs: MAX_CONCURRENT_JOBS,
      sofficeBin: SOFFICE_BIN,
      pdf2docxEnabled: PDF2DOCX_ENABLED,
    });
  }

  const route = req.method === "POST" ? ROUTES[req.url] : null;
  if (!route) {
    return sendJson(res, 404, { error: "Not found" });
  }

  try {
    await handleConvert(req, res, route);
  } catch (err) {
    // Last-resort guard so a bug above never leaves a request hanging.
    if (!res.headersSent) sendJson(res, 500, { error: err.message || "Internal error" });
  }
});

server.listen(PORT, () => {
  console.log(`[formatra-server] listening on http://localhost:${PORT}`);
  console.log(`[formatra-server] soffice binary: ${SOFFICE_BIN}`);
  console.log(
    `[formatra-server] max concurrent jobs: ${MAX_CONCURRENT_JOBS}, base timeout: ${SOFFICE_TIMEOUT_BASE_MS}ms (scales with file size, max ${SOFFICE_TIMEOUT_MAX_MS}ms), max file size: ${MAX_FILE_SIZE_MB}MB`
  );
});
