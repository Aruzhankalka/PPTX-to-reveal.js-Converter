const express = require('express');
const multer = require('multer');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const { parsePptx } = require('../parser/pptx');
const { generate } = require('../generator/revealjs');
const crypto = require("crypto");
const { saveResult } = require("../storage/resultStore");
const { sanitize } = require('../security/sanitizer');

const router = express.Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024, files: 1 },
});

async function validatePptxStructure(buffer) {
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (err) {
    return { valid: false, reason: 'File is not a valid ZIP archive' };
  }
  if (!zip.file('[Content_Types].xml')) {
    return { valid: false, reason: 'Missing [Content_Types].xml — not a valid PPTX' };
  }
  if (!zip.file('ppt/presentation.xml')) {
    return { valid: false, reason: 'Missing ppt/presentation.xml — not a valid PPTX' };
  }
  return { valid: true };
}

router.post('/convert', upload.single('file'), async (req, res, next) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        error_code: 'NO_FILE',
        message: 'No file was uploaded.',
      });
    }

    if (!req.file.originalname.toLowerCase().endsWith('.pptx')) {
      return res.status(400).json({
        error_code: 'INVALID_EXTENSION',
        message: `Expected a .pptx file but received "${req.file.originalname}".`,
      });
    }

    const check = await validatePptxStructure(req.file.buffer);
    if (!check.valid) {
      return res.status(400).json({
        error_code: 'INVALID_PPTX',
        message: check.reason,
      });
    }

    // Sanitize the buffer before parsing (NFR-08)
    const cleanBuffer = await sanitize(req.file.buffer);

    // Parse PPTX -> IR
    let ir;
    let media;
    let parseWarnings;
    try {
      const result = await parsePptx(cleanBuffer, { filename: req.file.originalname });
      ir = result.ir;
      media = result.media;
      console.log("Media:", media);
      parseWarnings = result.warnings;
    } catch (err) {
      if (err.code === 'INVALID_PPTX') {
        return res.status(400).json({
          error_code: 'INVALID_PPTX',
          message: err.message,
        });
      }
      throw err;
    }

    const { html, warnings: genWarnings } = generate(ir);
    let finalHtml = html;
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
      statistics: { slide_count: ir.slideset?.slides?.length || 0 }
    });

  } catch (err) {
    next(err);
  }
});

router.get('/preview/fixture', (req, res, next) => {
  try {
    const fixturePath = path.join(__dirname, '..', '..', 'tests', 'fixtures', 'minimal-ir.json');
    const ir = JSON.parse(fs.readFileSync(fixturePath, 'utf8'));
    const { html } = generate(ir);
    res.set('Content-Type', 'text/html; charset=utf-8');
    res.send(html);
  } catch (err) {
    next(err);
  }
});

module.exports = router;