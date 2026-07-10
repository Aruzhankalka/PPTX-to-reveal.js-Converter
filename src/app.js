/**
 * Express app assembly — wires the static web UI (src/web), the health
 * check, the upload/download API routers, the JSON 404 fallback, and the
 * centralized error handler, in the order that determines request handling
 * precedence. Exported (not started) so tests can mount it with supertest
 * without binding a port; server.js is the only thing that calls .listen().
 */

const express = require('express');
const path = require('path');
const swaggerUi = require('swagger-ui-express');
const uploadRouter = require('./api/upload');
const downloadRouter = require('./api/download');
const errorHandler = require('./api/errorHandler');
const openapiSpec = require('./api/openapi');

const app = express();

app.use(express.static(path.join(__dirname, 'web')));

/**
 * @openapi
 * /health:
 *   get:
 *     summary: Health check
 *     description: Liveness probe — always returns 200 while the process is up.
 *     responses:
 *       200:
 *         description: Service is running.
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 status: { type: string, example: ok }
 *                 nodeVersion: { type: string, example: v18.20.0 }
 *                 converterVersion: { type: string, example: 0.1.0 }
 */
app.get('/api/v1/health', (req, res) => {
  res.json({
    status: 'ok',
    nodeVersion: process.version,
    converterVersion: '0.1.0',
  });
});

// R5/R6: interactive OpenAPI docs UI. The full machine-readable spec is
// written to disk by `npm run openapi:export` (scripts/export-openapi.js);
// this route serves the same in-memory spec object for browsing.
app.use('/api/v1/docs', swaggerUi.serve, swaggerUi.setup(openapiSpec));

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
