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

module.exports = router;
