'use strict';

const JSZip = require('jszip');
const { generatePptx } = require('../src/generator/pptx/index');

const minimalIR = {
  slideset: {
    slides: [
      {
        title: 'Hello World',
        contents: []
      }
    ]
  }
};

const fullIR = {
  slideset: {
    slides: [
      {
        title: 'Slide 1 — Text only',
        contents: [
          {
            type: 'text',
            x: 100, y: 100, width: 800, height: 100,
            paragraphs: [
              { runs: [{ text: 'This is a test paragraph' }] }
            ]
          }
        ]
      },
      {
        title: 'Slide 2 — Empty',
        contents: []
      },
      {
        title: 'Slide 3 — Multiple paragraphs',
        contents: [
          {
            type: 'text',
            x: 50, y: 50, width: 600, height: 200,
            paragraphs: [
              { runs: [{ text: 'First paragraph' }] },
              { runs: [{ text: 'Second paragraph' }] }
            ]
          }
        ]
      }
    ]
  }
};

describe('Generator IR → PPTX (FR-16)', () => {

  test('should return a Buffer', async () => {
    const { buffer } = await generatePptx(minimalIR);
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);
  });

  test('should produce a valid ZIP archive', async () => {
    const { buffer } = await generatePptx(minimalIR);
    const zip = await JSZip.loadAsync(buffer);
    expect(zip).toBeDefined();
  });

  test('should contain required PPTX structure files', async () => {
    const { buffer } = await generatePptx(minimalIR);
    const zip = await JSZip.loadAsync(buffer);
    const files = Object.keys(zip.files);

    expect(files).toContain('[Content_Types].xml');
    expect(files).toContain('_rels/.rels');
    expect(files).toContain('ppt/presentation.xml');
  });

  test('should generate correct number of slide files', async () => {
    const { buffer } = await generatePptx(fullIR);
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files)
      .filter(f => f.match(/ppt\/slides\/slide\d+\.xml/));
    expect(slideFiles.length).toBe(3);
  });

  test('should include title text in slide XML', async () => {
    const { buffer } = await generatePptx(minimalIR);
    const zip = await JSZip.loadAsync(buffer);
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(slideXml).toContain('Hello World');
  });

  test('should include paragraph text in slide XML', async () => {
    const { buffer } = await generatePptx(fullIR);
    const zip = await JSZip.loadAsync(buffer);
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(slideXml).toContain('This is a test paragraph');
  });

  test('should return no warnings for valid IR', async () => {
    const { warnings } = await generatePptx(fullIR);
    expect(warnings.length).toBe(0);
  });

  test('should handle empty IR gracefully', async () => {
    const emptyIR = { slideset: { slides: [] } };
    const { buffer, warnings } = await generatePptx(emptyIR);
    expect(buffer).toBeInstanceOf(Buffer);
    expect(warnings.length).toBeGreaterThan(0);
  });

});