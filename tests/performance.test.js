'use strict';

const request = require('supertest');
const app = require('../src/app');
const path = require('path');

const SAMPLE_PPTX = path.join(__dirname, 'fixtures', 'sample.pptx');

describe('NFR-02 — Performance tests', () => {

  test('should convert sample.pptx in under 10 seconds', async () => {
    const start = Date.now();
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    const duration = Date.now() - start;

    console.log(`Conversion time: ${duration}ms`);
    expect(res.status).toBe(200);
    expect(duration).toBeLessThan(10000);
  }, 15000);

  test('should return slide_count in statistics', async () => {
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('statistics');
    expect(res.body.statistics.slide_count).toBeGreaterThan(0);
    console.log(`Slides converted: ${res.body.statistics.slide_count}`);
  }, 15000);

  test('should not use excessive memory during conversion', async () => {
    const memBefore = process.memoryUsage().heapUsed;
    const res = await request(app)
      .post('/api/v1/convert')
      .attach('file', SAMPLE_PPTX);
    const memAfter = process.memoryUsage().heapUsed;
    const memUsedMB = (memAfter - memBefore) / 1024 / 1024;

    console.log(`Memory used: ${memUsedMB.toFixed(2)} MB`);
    expect(res.status).toBe(200);
    expect(memUsedMB).toBeLessThan(1024);
  }, 15000);

  test('should not degrade across 10 consecutive conversions', async () => {
    const times = [];

    for (let i = 0; i < 10; i++) {
      const start = Date.now();
      const res = await request(app)
        .post('/api/v1/convert')
        .attach('file', SAMPLE_PPTX);
      times.push(Date.now() - start);
      expect(res.status).toBe(200);
    }

    const first = times[0];
    const last = times[times.length - 1];
    const avg = times.reduce((a, b) => a + b, 0) / times.length;

    console.log(`First: ${first}ms | Last: ${last}ms | Avg: ${avg.toFixed(0)}ms`);
    expect(last).toBeLessThan(first * 3);
    expect(avg).toBeLessThan(10000);
  }, 120000);

});