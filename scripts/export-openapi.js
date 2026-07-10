'use strict';

/**
 * Writes the OpenAPI spec (src/api/openapi.js) to disk as openapi.json —
 * the R5 machine-readable artifact, generated rather than hand-maintained
 * so it can never drift from the @openapi JSDoc blocks it's built from.
 * Run via `npm run openapi:export`; also requireable so tests can reuse
 * the same write logic against a scratch path.
 */

const fs = require('fs');
const path = require('path');

const DEFAULT_OUT_PATH = path.join(__dirname, '..', 'openapi.json');

/**
 * @param {string} [outPath] - destination file path
 * @returns {string} the path written to
 */
function exportOpenapiJson(outPath = DEFAULT_OUT_PATH) {
  const openapiSpec = require('../src/api/openapi');
  fs.writeFileSync(outPath, JSON.stringify(openapiSpec, null, 2) + '\n');
  return outPath;
}

if (require.main === module) {
  const written = exportOpenapiJson();
  console.log(`OpenAPI spec written to ${written}`);
}

module.exports = { exportOpenapiJson, DEFAULT_OUT_PATH };
