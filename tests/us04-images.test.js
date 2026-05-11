'use strict';

const fs = require('fs');
const path = require('path');
const request = require('supertest');
const express = require('express');
const uploadRouter = require('../src/api/upload');
const downloadRouter = require('../src/api/download');
const errorHandler = require('../src/api/errorHandler');

const app = express();
app.use('/api/v1', uploadRouter);
app.use('/api/v1', downloadRouter);
app.use(errorHandler);

const SAMPLE_PPTX = path.join(__dirname, 'fixtures', 'sample.pptx');

describe('US-04 — Images are preserved in the reveal.js output', () => {

  test('sample.pptx fixture exists and is a valid file', () => {
    expect(fs.existsSync(SAMPLE_PPTX)).toBe(true);
    const stats = fs.statSync(SAMPLE_PPTX);
    expect(stats.size).toBeGreaterThan(0);
  });

  test('uploading sample.pptx returns HTTP 200 with a result_id', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('result_id');
  }, 30000);

  test('generated HTML contains img tags for images in the PPTX', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    expect(res.status).toBe(200);
    const resultId = res.body.result_id;
    expect(resultId).toBeDefined();
    const preview = await request(app)
      .get(`/api/v1/preview/${resultId}`);
    expect(preview.status).toBe(200);
    expect(preview.text).toContain('<img');
  }, 30000);

  test('image paths in generated HTML are not empty', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    expect(res.status).toBe(200);
    const preview = await request(app)
      .get(`/api/v1/preview/${res.body.result_id}`);
    const imgMatches = preview.text.match(/<img[^>]+src="([^"]+)"/g) || [];
    expect(imgMatches.length).toBeGreaterThan(0);
  }, 30000);

  test('all image src paths in the output use the API media endpoint', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    expect(res.status).toBe(200);
    const preview = await request(app)
      .get(`/api/v1/preview/${res.body.result_id}`);
    const imgSrcs = [...preview.text.matchAll(/<img[^>]+src="([^"]+)"/g)]
      .map(m => m[1]);
    expect(imgSrcs.length).toBeGreaterThan(0);
    for (const src of imgSrcs) {
      expect(src).toMatch(/^\/api\/v1\/media\//);
    }
  }, 30000);

});