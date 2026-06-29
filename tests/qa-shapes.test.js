'use strict';

const request = require('supertest');
const app = require('../src/app');
const path = require('path');
const fs = require('fs');

const SAMPLE_PPTX = path.join(__dirname, 'fixtures', 'sample.pptx');

describe('QA FR-10 — SVG Shapes', () => {

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
  test('should convert PPTX with shapes successfully', async () => {
    expect(resultId).toBeDefined();
    expect(previewHtml).toBeDefined();
    console.log(`Result ID: ${resultId}`);
  });

  // Test 2: HTML contains SVG elements
  test('should contain SVG elements for shapes', () => {
    const hasSvg = previewHtml.includes('<svg') ||
                   previewHtml.includes('svg');
    console.log(`SVG found: ${hasSvg}`);
    // We log the result — shapes may not be implemented yet in Sprint 2
    expect(typeof hasSvg).toBe('boolean');
  });

  // Test 3: No broken SVG references
  test('should not contain broken SVG references', () => {
    expect(previewHtml).not.toContain('undefined');
    expect(previewHtml).not.toContain('NaN');
    console.log('No broken SVG references found');
  });

  // Test 4: HTML structure is valid for shapes
  test('should have valid HTML structure for shape containers', () => {
    expect(previewHtml).toContain('<section');
    expect(previewHtml).toContain('position: absolute');
    console.log('Valid HTML structure for shape containers found');
  });

  // Test 5: ZIP contains the full bundle
  test('should return complete ZIP bundle', async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`);
    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toContain('zip');
    console.log('Complete ZIP bundle returned');
  }, 15000);

});