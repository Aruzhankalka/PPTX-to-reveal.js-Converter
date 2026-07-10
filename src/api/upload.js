/**
 * Upload API — the single entry point into the whole conversion pipeline
 * (sanitize -> parser/pptx -> generator/revealjs). Accepts a multipart PPTX
 * upload, validates it's a real PPTX (ZIP + required parts) before doing any
 * real work, and stores the {html, ir, media, warnings} result in
 * storage/resultStore for download.js's preview/result/media routes to
 * serve. Bad-input errors are returned inline as 400 JSON; unexpected
 * failures are forwarded to errorHandler.js (500).
 */

const express = require('express');
const multer = require('multer');
const JSZip = require('jszip');
const { parsePptx } = require('../parser/pptx');
const { generate } = require('../generator/revealjs');
const crypto = require("crypto");
const { saveResult } = require("../storage/resultStore");
const { sanitize } = require('../security/sanitizer');

const router = express.Router();

// Multer config: store uploads in memory (small files, stateless per spec §2.1)
// Limit: 50 MB — generous enough for the 50-slide reference deck per NFR-02
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50 MB
    files: 1,
  },
});

/**
 * Verify that an already-opened zip has the minimum shape of a real PPTX:
 *   - [Content_Types].xml
 *   - ppt/presentation.xml
 * Zip-validity itself (is this even a ZIP archive) is checked by the caller
 * before this runs, since opening the zip is now a shared, one-time step.
 * @param {JSZip} zip - an opened JSZip instance
 * @returns {{valid: boolean, reason: (string|undefined)}} reason is only set when valid is false
 */
function validatePptxStructure(zip) {
  if (!zip.file('[Content_Types].xml')) {
    return { valid: false, reason: 'Missing [Content_Types].xml — not a valid PPTX' };
  }

  if (!zip.file('ppt/presentation.xml')) {
    return { valid: false, reason: 'Missing ppt/presentation.xml — not a valid PPTX' };
  }

  return { valid: true };
}

/**
 * @openapi
 * /convert:
 *   post:
 *     summary: Convert a PPTX file to a reveal.js presentation
 *     description: >
 *       Forward direction: PPTX -> reveal.js (FR-01, FR-02). Accepts one
 *       multipart file (field name "file", .pptx extension required, max
 *       50 MB). The uploaded ZIP is validated, sanitized (VBA macros and
 *       script-bearing SVG media removed), parsed, and rendered; the result
 *       is cached server-side (30 minute TTL) and referenced by result_id
 *       for the preview/result/media endpoints below.
 *     requestBody:
 *       required: true
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               file:
 *                 type: string
 *                 format: binary
 *                 description: A .pptx file, max 50 MB.
 *             required: [file]
 *     responses:
 *       200:
 *         description: Conversion succeeded.
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 result_id: { type: string, format: uuid }
 *                 preview_url: { type: string, example: "/api/v1/preview/{result_id}" }
 *                 download_url: { type: string, example: "/api/v1/result/{result_id}" }
 *                 warnings: { type: array, items: { type: string } }
 *                 statistics:
 *                   type: object
 *                   properties:
 *                     slide_count: { type: integer }
 *       400:
 *         description: >
 *           Bad input. error_code is one of NO_FILE (no file field), FILE_TOO_LARGE
 *           (>50 MB, from multer), INVALID_UPLOAD (wrong field name or >1 file),
 *           INVALID_EXTENSION (filename doesn't end in .pptx), or INVALID_PPTX
 *           (not a valid ZIP, or missing required PPTX parts).
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 error_code: { type: string }
 *                 message: { type: string }
 *       500:
 *         description: Unexpected server error (error_code INTERNAL_ERROR).
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 error_code: { type: string, example: INTERNAL_ERROR }
 *                 message: { type: string }
 */
router.post('/convert', upload.single('file'), async (req, res, next) => {
  try {
    // Multer puts the file on req.file when storage:memoryStorage is used
    if (!req.file) {
      return res.status(400).json({
        error_code: 'NO_FILE',
        message: 'No file was uploaded. Send a multipart/form-data request with field "file".',
      });
    }

    // Quick extension check — fast rejection before we open the zip
    if (!req.file.originalname.toLowerCase().endsWith('.pptx')) {
      return res.status(400).json({
        error_code: 'INVALID_EXTENSION',
        message: `Expected a .pptx file but received "${req.file.originalname}".`,
      });
    }

    // Open the zip once and reuse the same instance for structural validation,
    // sanitization, and parsing — these used to each open (and sanitize also
    // re-serialized) their own independent copy of the same archive.
    let zip;
    try {
      zip = await JSZip.loadAsync(req.file.buffer);
    } catch {
      return res.status(400).json({
        error_code: 'INVALID_PPTX',
        message: 'File is not a valid ZIP archive',
      });
    }

    // Real structural validation
    const check = validatePptxStructure(zip);
    if (!check.valid) {
      return res.status(400).json({
        error_code: 'INVALID_PPTX',
        message: check.reason,
      });
    }

    // Sanitize the zip in place before parsing (NFR-08)
    await sanitize(zip);

    // Parse PPTX -> IR
    let ir;
    let media;
    let parseWarnings;
    try {
      const result = await parsePptx(zip, { filename: req.file.originalname });
      ir = result.ir;
      media = result.media;
      parseWarnings = result.warnings;
    } catch (err) {
      if (err.code === 'INVALID_PPTX') {
        return res.status(400).json({
          error_code: 'INVALID_PPTX',
          message: err.message,
        });
      }
      throw err; // unexpected -> centralized error handler -> 500
    }

    // Generate reveal.js HTML
    const { html, warnings: genWarnings } = generate(ir);


    let finalHtml = html;

// Store generated result for preview and download endpoints
    const resultId = crypto.randomUUID();

    finalHtml = finalHtml.replace(
      /src="media\/([^"]+)"/g,
      `src="/api/v1/media/${resultId}/$1"`
    );

    const warnings = [...(parseWarnings || []), ...(genWarnings || [])];

    saveResult(resultId, {
      html: finalHtml,
      ir,
      media,
      filename: req.file.originalname,
      warnings
    });

    res.json({
      result_id: resultId,
      preview_url: `/api/v1/preview/${resultId}`,
      download_url: `/api/v1/result/${resultId}`,
      warnings,
      statistics: {
        slide_count: ir.slideset?.slides?.length || 0
      }
    });


  } catch (err) {
    // Pass to centralized error handler (we'll add this next)
    next(err);
  }
});

module.exports = router;
