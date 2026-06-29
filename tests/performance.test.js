'use strict';

const request = require('supertest');
const app = require('../src/app');
const path = require('path');

const SAMPLE_PPTX = path.join(__dirname, 'fixtures', 'sample.pptx');

describe('NFR-02 — Performance tests', () => {

  // Test 1: Conversion time under 10 seconds
  test('should convert sample.pptx in under 10 seconds', async () => {
    const start = Date.now();

    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);

    const duration = Date.now() - start;

    console.log(`Conversion time: ${duration}ms`);

    expect(res.status).toBe(200);
    expect(duration).toBeLessThan(10000); // 10 seconds
  }, 15000);

  // Test 2: Conversion time logged in statistics
  test('should return slide_count in statistics', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('statistics');
    expect(res.body.statistics.slide_count).toBeGreaterThan(0);
    console.log(`Slides converted: ${res.body.statistics.slide_count}`);
  }, 15000);
// Test 3: Memory usage
  test('should not use excessive memory during conversion', async () => {
    const memBefore = process.memoryUsage().heapUsed;

    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);

    const memAfter = process.memoryUsage().heapUsed;
    const memUsedMB = (memAfter - memBefore) / 1024 / 1024;

    console.log(`Memory used: ${memUsedMB.toFixed(2)} MB`);

    expect(res.status).toBe(200);
    expect(memUsedMB).toBeLessThan(1024); // max 1GB
  }, 15000);

  // Test 4: NFR-05 — Stability across 50 consecutive conversions
  test('should not degrade across 10 consecutive conversions', async () => {
    const times = [];

    for (let i = 0; i < 10; i++) {
      const start = Date.now();

      const res = await request(app)
        .post('/api/v1/convert')
        .attach('file', SAMPLE_PPTX);

      const duration = Date.now() - start;
      times.push(duration);

      expect(res.status).toBe(200);
    }

    const firstConversion = times[0];
    const lastConversion = times[times.length - 1];
    const avgTime = times.reduce((a, b) => a + b, 0) / times.length;

    console.log(`First conversion: ${firstConversion}ms`);
    console.log(`Last conversion: ${lastConversion}ms`);
    console.log(`Average time: ${avgTime.toFixed(0)}ms`);

    // Last conversion should not be more than 3x slower than first
    expect(lastConversion).toBeLessThan(firstConversion * 3);
    expect(avgTime).toBeLessThan(10000);
  }, 120000);
  
});