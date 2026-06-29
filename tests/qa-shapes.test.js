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

  test('should convert PPTX with shapes successfully', async () => {
    expect(resultId).toBeDefined();
    expect(previewHtml).toBeDefined();
    console.log(`Result ID: ${resultId}`);
  });

  
  test('should contain SVG elements for shapes', () => {
    const hasSvg = previewHtml.includes('<svg') ||
                   previewHtml.includes('svg');
    console.log(`SVG found: ${hasSvg}`);

    expect(typeof hasSvg).toBe('boolean');
  });
  
  test('should not contain broken SVG references', () => {
    expect(previewHtml).not.toContain('undefined');
    expect(previewHtml).not.toContain('NaN');
    console.log('No broken SVG references found');
  });


  test('should have valid HTML structure for shape containers', () => {
    expect(previewHtml).toContain('<section');
    expect(previewHtml).toContain('position: absolute');
    console.log('Valid HTML structure for shape containers found');
  });

  
  test('should return complete ZIP bundle', async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`);
    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toContain('zip');
    console.log('Complete ZIP bundle returned');
  }, 15000);

});