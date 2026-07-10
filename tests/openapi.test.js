'use strict';

const request = require('supertest');
const fs = require('fs');
const os = require('os');
const path = require('path');
const app = require('../src/app');
const { exportOpenapiJson } = require('../scripts/export-openapi');

// Cheap guard against undocumented future endpoints (R5): every route
// mounted in src/app.js must have a corresponding @openapi block that
// swagger-jsdoc picks up, in both the live docs UI and the exported spec.
const EXPECTED_PATHS = ['/convert', '/preview/{id}', '/result/{id}', '/media/{id}/{filename}', '/health'];

describe('OpenAPI documentation coverage', () => {
  test('GET /api/v1/docs (Swagger UI) serves the page and embeds all 5 documented paths', async () => {
    const page = await request(app).get('/api/v1/docs/');
    expect(page.status).toBe(200);

    // swagger-ui-express embeds the actual spec JSON in a companion script
    // (swagger-ui-init.js) rather than inline in the HTML page itself.
    const initScript = await request(app).get('/api/v1/docs/swagger-ui-init.js');
    expect(initScript.status).toBe(200);
    for (const p of EXPECTED_PATHS) {
      expect(initScript.text).toContain(p);
    }
  });

  test('exported openapi.json contains all 5 documented paths', () => {
    const outPath = path.join(os.tmpdir(), `openapi-test-${Date.now()}.json`);
    exportOpenapiJson(outPath);
    try {
      const spec = JSON.parse(fs.readFileSync(outPath, 'utf8'));
      const paths = Object.keys(spec.paths || {});
      for (const p of EXPECTED_PATHS) {
        expect(paths).toContain(p);
      }
    } finally {
      fs.unlinkSync(outPath);
    }
  });
});
