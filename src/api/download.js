/**
 * Download/preview API — serves conversion results produced by upload.js's
 * /convert handler and stored in storage/resultStore. Exposes the generated
 * reveal.js HTML for in-browser preview, a self-contained offline ZIP bundle
 * (HTML + reveal.js assets + media rewritten to relative paths), and
 * individual media files by id. All routes 404 when the result id is unknown
 * or expired (resultStore is in-memory and does not persist across restarts).
 */

const express = require("express");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { getResult } = require("../storage/resultStore");

const router = express.Router();

// reveal.js files we bundle into the offline ZIP.
// Paths are relative to node_modules/reveal.js/dist/.
// Keep this list in sync with the <link>/<script> tags emitted by
// src/generator/revealjs/html.js (TC-03: reveal.js v4.6+).
const REVEALJS_BUNDLE_FILES = [
  "reset.css",
  "reveal.css",
  "reveal.js",
  "theme/white.css",
];

const REVEALJS_CDN_PREFIX = "https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/";

/**
 * Rewrite a generated HTML document so it can run offline from a ZIP.
 *
 * Two substitutions:
 *   1. /api/v1/media/{id}/foo.png  ->  assets/foo.png   (US-05 AC-3)
 *   2. CDN reveal.js URLs          ->  reveal/...        (offline demo safety)
 *
 * Exposed for unit testing.
 *
 * @param {string} html - HTML as stored in resultStore (preview-mode URLs)
 * @param {string} resultId - the result ID used in the /api/v1/media/{id}/ URLs
 * @returns {string} - HTML with relative paths suitable for the ZIP bundle
 */
function rewriteHtmlForBundle(html, resultId) {
  // split/join = literal global replace; no RegExp construction needed,
  // so no metacharacter escaping to worry about either.
  return html
    .split(`/api/v1/media/${resultId}/`).join("assets/")
    .split(REVEALJS_CDN_PREFIX).join("reveal/");
}

/**
 * Load the reveal.js distribution files we need to bundle.
 * Resolved via require.resolve so it works regardless of the cwd
 * the server is started from.
 *
 * @returns {Map<string, Buffer>} - mapping from zip-relative path to file bytes
 */
function loadRevealJsBundle() {
  // Direct path avoids Jest's module resolver intercepting require.resolve.
  // download.js is at src/api/download.js, so ../../node_modules is the project root.
  const distDir = path.resolve(__dirname, "..", "..", "node_modules", "reveal.js", "dist");

  const files = new Map();
  for (const rel of REVEALJS_BUNDLE_FILES) {
    const abs = path.join(distDir, rel);
    files.set(`reveal/${rel}`, fs.readFileSync(abs));
  }
  return files;
}

/**
 * @openapi
 * /preview/{id}:
 *   get:
 *     summary: In-browser preview of a conversion result
 *     description: >
 *       In-browser preview (FR-15, NFR-01). Returns the stored HTML as-is;
 *       its embedded media URLs are served by GET /media/{id}/{filename}.
 *     parameters:
 *       - name: id
 *         in: path
 *         required: true
 *         schema: { type: string, format: uuid }
 *         description: result_id returned by POST /convert.
 *     responses:
 *       200:
 *         description: The reveal.js HTML document.
 *         content:
 *           text/html: {}
 *       404:
 *         description: Unknown or expired result_id (error_code RESULT_NOT_FOUND).
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 error_code: { type: string, example: RESULT_NOT_FOUND }
 *                 message: { type: string }
 */
router.get("/preview/:id", (req, res) => {
  const result = getResult(req.params.id);

  if (!result) {
    return res.status(404).json({
      error_code: "RESULT_NOT_FOUND",
      message: "Conversion result not found or expired.",
    });
  }

  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.send(result.html);
});

/**
 * @openapi
 * /result/{id}:
 *   get:
 *     summary: Download an offline-usable ZIP bundle
 *     description: >
 *       Offline-usable ZIP bundle (FR-15, US-05 AC-3). The ZIP contains
 *       /index.html (rewritten so all asset references are relative),
 *       /assets/* (every media file from the conversion), and /reveal/*
 *       (the reveal.js distribution files), so the bundle runs standalone
 *       with no network access.
 *     parameters:
 *       - name: id
 *         in: path
 *         required: true
 *         schema: { type: string, format: uuid }
 *         description: result_id returned by POST /convert.
 *     responses:
 *       200:
 *         description: The ZIP bundle, as an attachment.
 *         content:
 *           application/zip: {}
 *       404:
 *         description: Unknown or expired result_id (error_code RESULT_NOT_FOUND).
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 error_code: { type: string, example: RESULT_NOT_FOUND }
 *                 message: { type: string }
 */
router.get("/result/:id", async (req, res, next) => {
  try {
    const result = getResult(req.params.id);

    if (!result) {
      return res.status(404).json({
        error_code: "RESULT_NOT_FOUND",
        message: "Conversion result not found or expired.",
      });
    }

    const zip = new JSZip();

    // 1. Rewritten HTML at the ZIP root
    const bundleHtml = rewriteHtmlForBundle(result.html, req.params.id);
    zip.file("index.html", bundleHtml);

    // 2. Media files under /assets/
    //    result.media[].bundlePath is "media/<filename>" — strip prefix
    for (const mediaFile of result.media || []) {
      const filename = mediaFile.bundlePath.replace(/^media\//, "");
      zip.file(`assets/${filename}`, mediaFile.bytes);
    }

    // 3. reveal.js distribution under /reveal/
    const revealFiles = loadRevealJsBundle();
    for (const [zipPath, bytes] of revealFiles) {
      zip.file(zipPath, bytes);
    }

    const buffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });

    const baseName = (result.filename || "presentation").replace(/\.pptx$/i, "");

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${baseName}.zip"`
    );
    res.setHeader("Content-Length", buffer.length);
    res.send(buffer);
  } catch (err) {
    next(err);
  }
});

/**
 * @openapi
 * /media/{id}/{filename}:
 *   get:
 *     summary: Fetch one media file from a conversion result
 *     description: Serves media files referenced by the preview endpoint's HTML.
 *     parameters:
 *       - name: id
 *         in: path
 *         required: true
 *         schema: { type: string, format: uuid }
 *         description: result_id returned by POST /convert.
 *       - name: filename
 *         in: path
 *         required: true
 *         schema: { type: string }
 *         description: Media filename as referenced in the preview HTML.
 *     responses:
 *       200:
 *         description: >
 *           The media file. Content-Type is derived from the file extension
 *           (png/jpg/jpeg/gif/svg map to their image/* type; anything else
 *           falls back to application/octet-stream).
 *         content:
 *           image/*: {}
 *           application/octet-stream: {}
 *       404:
 *         description: >
 *           error_code RESULT_NOT_FOUND when id is unknown/expired, or
 *           MEDIA_NOT_FOUND when id is valid but filename isn't part of
 *           that result.
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 error_code: { type: string, enum: [RESULT_NOT_FOUND, MEDIA_NOT_FOUND] }
 *                 message: { type: string }
 */
router.get("/media/:id/:filename", (req, res) => {
  const result = getResult(req.params.id);

  if (!result) {
    return res.status(404).json({
      error_code: "RESULT_NOT_FOUND",
      message: "Conversion result not found or expired.",
    });
  }

  const media = result.media || [];

  const mediaFile = media.find((item) => {
    return item.bundlePath === `media/${req.params.filename}`;
  });

  if (!mediaFile) {
    return res.status(404).json({
      error_code: "MEDIA_NOT_FOUND",
      message: "Media file not found.",
    });
  }

  const extension = req.params.filename.split(".").pop().toLowerCase();

  const mimeTypes = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
  };

  res.setHeader(
    "Content-Type",
    mimeTypes[extension] || "application/octet-stream"
  );

  res.send(mediaFile.bytes);
});

module.exports = router;
module.exports.rewriteHtmlForBundle = rewriteHtmlForBundle;
