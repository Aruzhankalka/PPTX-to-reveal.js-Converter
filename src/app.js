const express = require('express');
const path = require('path');
const uploadRouter = require('./api/upload');
const downloadRouter = require('./api/download');
const errorHandler = require('./api/errorHandler');

const app = express();

app.use(express.static(path.join(__dirname, 'web')));

app.get('/api/v1/health', (req, res) => {
  res.json({
    status: 'ok',
    nodeVersion: process.version,
    converterVersion: '0.1.0',
  });
});

app.use('/api/v1', uploadRouter);
app.use('/api/v1', downloadRouter);

app.use((req, res) => {
  res.status(404).json({
    error_code: 'NOT_FOUND',
    message: `${req.method} ${req.path} is not a known endpoint.`,
  });
});

app.use(errorHandler);

module.exports = app;
