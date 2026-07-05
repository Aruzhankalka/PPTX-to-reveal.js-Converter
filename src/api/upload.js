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
 * @returns {{ valid: boolean, reason?: string }}
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
 * POST /api/v1/convert
 * Forward direction: PPTX -> reveal.js (FR-01, FR-02)
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
    } catch (err) {
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
