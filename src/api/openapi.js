'use strict';

/**
 * OpenAPI spec builder (R5/R6) — scans the @openapi JSDoc blocks in
 * app.js/upload.js/download.js via swagger-jsdoc and assembles the full
 * spec object. Used both to mount swagger-ui-express at /api/v1/docs
 * (app.js) and by scripts/export-openapi.js to write openapi.json to disk.
 */

const swaggerJsdoc = require('swagger-jsdoc');
const path = require('path');

const options = {
  definition: {
    openapi: '3.0.3',
    info: {
      title: 'PPTX to reveal.js Converter API',
      version: '0.1.0',
      description:
        'Converts an uploaded PowerPoint (.pptx) file into a reveal.js HTML ' +
        'presentation, and serves the result for in-browser preview, offline ' +
        'ZIP download, or individual media retrieval.',
    },
    servers: [{ url: '/api/v1' }],
  },
  apis: [
    path.join(__dirname, 'upload.js'),
    path.join(__dirname, 'download.js'),
    path.join(__dirname, '..', 'app.js'),
  ],
};

const openapiSpec = swaggerJsdoc(options);

module.exports = openapiSpec;
