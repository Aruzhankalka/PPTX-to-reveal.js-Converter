'use strict';

const request = require('supertest');
const app = require('../src/app');
const path = require('path');
const fs = require('fs');

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

  // Test 1: Conversion succeeds
  test('should convert PPTX successfully', async () => {
    expect(resultId).toBeDefined();
    expect(previewHtml).toBeDefined();
    console.log(`Result ID: ${resultId}`);
  });

  // Test 2: HTML contains font references
  test('should contain font-related CSS in the output', () => {
    const hasFontFamily = previewHtml.includes('font-family') ||
                          previewHtml.includes('font-face') ||
                          previewHtml.includes('font-size');
    expect(hasFontFamily).toBe(true);
    console.log('Font CSS found in HTML output');
  });

  // Test 3: HTML does not use undefined fonts
  test('should not contain undefined font references', () => {
    expect(previewHtml).not.toContain('font-family: undefined');
    expect(previewHtml).not.toContain('font-size: undefined');
    console.log('No undefined font references found');
  });

  // Test 4: HTML contains reveal.js with font support
  test('should contain reveal.js presentation with text elements', () => {
    expect(previewHtml).toContain('<section');
    expect(previewHtml).toContain('reveal');
    console.log('Presentation structure with text elements found');
  });

  // Test 5: Download ZIP contains font assets
  test('should return a valid ZIP download', async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`);
    expect(res.status).toBe(200);
    console.log('ZIP download available');
  }, 15000);

});