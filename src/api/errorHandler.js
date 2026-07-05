/**
 * Centralized Express error middleware.
 * Per Specification §4.4 — never leaks stack traces, always returns
 * a consistent JSON error body, keeps the process alive (NFR-04).
 */
// The unused 4th parameter is required: Express identifies error middleware
// by function arity (4 args), so `_next` must stay.
function errorHandler(err, req, res, _next) {
  // Log full details server-side for debugging
  console.error(`[${new Date().toISOString()}] Error on ${req.method} ${req.path}:`, err);

  // Multer-specific error handling
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({
      error_code: 'FILE_TOO_LARGE',
      message: 'File exceeds the 50 MB upload limit.',
    });
  }
  if (err.code === 'LIMIT_FILE_COUNT' || err.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(400).json({
      error_code: 'INVALID_UPLOAD',
      message: 'Upload only one file with form field name "file".',
    });
  }

  // Generic fallback — never expose internals to the client
  return res.status(500).json({
    error_code: 'INTERNAL_ERROR',
    message: 'An unexpected error occurred. Please try again.',
  });
}

module.exports = errorHandler;