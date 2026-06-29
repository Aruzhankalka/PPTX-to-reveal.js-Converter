'use strict';

const request = require('supertest');
const app = require('../src/app');
const path = require('path');

const SAMPLE_PPTX = path.join(__dirname, 'fixtures', 'sample.pptx');

describe('QA FR-08 — Web Fonts', () => {

  let resultId;
  let previewHtml;

  beforeAll(async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    resultId = res.body.result_id;

    const preview = await request(app)
      .get(`/api/v1/preview/${resultId}`);
    previewHtml = preview.text;
  }, 30000);

  test('should convert PPTX successfully', async () => {
    expect(resultId).toBeDefined();
    expect(previewHtml).toBeDefined();
  });

  test('should contain font-related CSS in the output', () => {
    const hasFontFamily = previewHtml.includes('font-family') ||
                          previewHtml.includes('font-face') ||
                          previewHtml.includes('font-size');
    expect(hasFontFamily).toBe(true);
  });

  test('should not contain undefined font references', () => {
    expect(previewHtml).not.toContain('font-family: undefined');
    expect(previewHtml).not.toContain('font-size: undefined');
  });

  test('should contain reveal.js presentation with text elements', () => {
    expect(previewHtml).toContain('<section');
    expect(previewHtml).toContain('reveal');
  });

  test('should return a valid ZIP download', async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`);
    expect(res.status).toBe(200);
  }, 15000);

});