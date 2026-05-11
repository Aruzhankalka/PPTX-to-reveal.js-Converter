const request = require('supertest');
const express = require('express');
const JSZip = require('jszip');
const uploadRouter = require('../src/api/upload');
const errorHandler = require('../src/api/errorHandler');

function buildApp() {
  const app = express();
  app.use('/api/v1', uploadRouter);
  app.use(errorHandler);
  return app;
}

async function buildFakePptx() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0"?><presentation/>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('POST /api/v1/convert', () => {
  test('rejects request with no file (FR-01)', async () => {
    const res = await request(buildApp()).post('/api/v1/convert');
    expect(res.status).toBe(400);
    expect(res.body.error_code).toBe('NO_FILE');
  });

  test('rejects file with non-pptx extension (FR-02)', async () => {
    const res = await request(buildApp())
      .post('/api/v1/convert')
      .attach('file', Buffer.from('hello'), 'notes.txt');
    expect(res.status).toBe(400);
    expect(res.body.error_code).toBe('INVALID_EXTENSION');
  });

  test('rejects file with .pptx extension but invalid contents (FR-02, FR-18)', async () => {
    const res = await request(buildApp())
      .post('/api/v1/convert')
      .attach('file', Buffer.from('not a real zip'), 'fake.pptx');
    expect(res.status).toBe(400);
    expect(res.body.error_code).toBe('INVALID_PPTX');
  });

  test('rejects .pptx with valid zip but missing presentation.xml (FR-02)', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    const buf = await zip.generateAsync({ type: 'nodebuffer' });

    const res = await request(buildApp())
      .post('/api/v1/convert')
      .attach('file', buf, 'broken.pptx');
    expect(res.status).toBe(400);
    expect(res.body.error_code).toBe('INVALID_PPTX');
  });

  test('accepts a structurally valid .pptx and returns a result_id (FR-01)', async () => {
    const buf = await buildFakePptx();
    const res = await request(buildApp())
      .post('/api/v1/convert')
      .attach('file', buf, 'good.pptx');
    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toMatch(/application\/json/);
    expect(res.body).toHaveProperty('result_id');
    expect(res.body).toHaveProperty('preview_url');
    expect(res.body).toHaveProperty('download_url');
    expect(res.body.statistics.slide_count).toBe(0);
  });
});
