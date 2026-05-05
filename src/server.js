const express = require('express');
const path = require('path');
const uploadRouter = require('./api/upload');
const errorHandler = require('./api/errorHandler');

const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'web')));

// Health check
app.get('/api/v1/health', (req, res) => {
  res.json({
    status: 'ok',
    nodeVersion: process.version,
    converterVersion: '0.1.0',
  });
});

// Conversion endpoints under /api/v1
app.use('/api/v1', uploadRouter);

// 404 for unknown routes — must come AFTER all real routes
app.use((req, res) => {
  res.status(404).json({
    error_code: 'NOT_FOUND',
    message: `${req.method} ${req.path} is not a known endpoint.`,
  });
});

// Centralized error handler — must be LAST
app.use(errorHandler);

app.listen(PORT, () => {
  console.log(`Converter running on http://localhost:${PORT}`);
});
