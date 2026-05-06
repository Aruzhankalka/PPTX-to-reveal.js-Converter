const express = require("express");
const { getResult } = require("../storage/resultStore");

const router = express.Router();

router.get("/preview/:id", (req, res) => {
  const result = getResult(req.params.id);

  if (!result) {
    return res.status(404).json({
      error_code: "RESULT_NOT_FOUND",
      message: "Conversion result not found or expired."
    });
  }

  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.send(result.html);
});

router.get("/result/:id", (req, res) => {
  const result = getResult(req.params.id);

  if (!result) {
    return res.status(404).json({
      error_code: "RESULT_NOT_FOUND",
      message: "Conversion result not found or expired."
    });
  }

  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${result.filename || "presentation"}.html"`
  );

  res.send(result.html);
});


router.get("/media/:id/:filename", (req, res) => {
  const result = getResult(req.params.id);

  if (!result) {
    return res.status(404).json({
      error_code: "RESULT_NOT_FOUND",
      message: "Conversion result not found or expired."
    });
  }

  const media = result.media || [];

  const mediaFile = media.find((item) => {
    return item.bundlePath === `media/${req.params.filename}`;
  });

  if (!mediaFile) {
    return res.status(404).json({
      error_code: "MEDIA_NOT_FOUND",
      message: "Media file not found."
    });
  }

  const extension = req.params.filename.split(".").pop().toLowerCase();

  const mimeTypes = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml"
  };

  res.setHeader(
    "Content-Type",
    mimeTypes[extension] || "application/octet-stream"
  );

  res.send(mediaFile.bytes);
});

module.exports = router;
