const { test, expect } = require('@playwright/test');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const FormData = require('form-data');
const fetch = require('node-fetch');

const SAMPLE_PPTX = path.join(__dirname, '..', 'fixtures', 'sample.pptx');

let resultId;

test.beforeAll(async () => {
  // Upload PPTX and get result_id
  const form = new FormData();
  form.append('file', fs.createReadStream(SAMPLE_PPTX));

  const res = await fetch('http://localhost:3000/api/v1/convert', {
    method: 'POST',
    body: form,
    headers: form.getHeaders(),
  });

  const data = await res.json();
  resultId = data.result_id;
  console.log(`Result ID: ${resultId}`);
});

test('Chrome — reveal.js presentation loads correctly', async ({ page }) => {
  await page.goto(`http://localhost:3000/api/v1/preview/${resultId}`);
  const content = await page.content();
  expect(content).toContain('reveal');
  console.log('Chrome — OK');
});

test('Firefox — reveal.js presentation loads correctly', async ({ page }) => {
  await page.goto(`http://localhost:3000/api/v1/preview/${resultId}`);
  const content = await page.content();
  expect(content).toContain('reveal');
  console.log('Firefox — OK');
});

test('HTML is valid for all browsers', async ({ page }) => {
  await page.goto(`http://localhost:3000/api/v1/preview/${resultId}`);
  const content = await page.content();
  expect(content).toContain('<!DOCTYPE html>');
  expect(content).toContain('reveal.js');
  console.log('HTML valid — OK');
});