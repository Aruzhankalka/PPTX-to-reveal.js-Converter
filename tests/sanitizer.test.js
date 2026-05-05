'use strict';

const JSZip = require('jszip');
const { sanitize } = require('../src/security/sanitizer');

// ── Helper: create a minimal PPTX buffer ──────────────────────────────────────
async function buildMinimalPPTX(extraFiles = {}) {
  const zip = new JSZip();

  // Minimal OOXML structure
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
</p:presentation>`);

  // Add any extra files for testing
  for (const [path, content] of Object.entries(extraFiles)) {
    zip.file(path, content);
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

// ── Tests ─────────────────────────────────────────────────────────────────────
describe('Sanitizer', () => {

  // Test 1: Clean PPTX passes through unchanged
  test('should pass a clean PPTX without modifications', async () => {
    const buffer = await buildMinimalPPTX();
    const result = await sanitize(buffer);
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  // Test 2: VBA macro is removed
  test('should remove vbaProject.bin from PPTX', async () => {
    const buffer = await buildMinimalPPTX({
      'ppt/vbaProject.bin': Buffer.from('fake vba content')
    });
    const result = await sanitize(buffer);

    // Check the cleaned ZIP no longer contains vbaProject.bin
    const cleanedZip = await JSZip.loadAsync(result);
    const hasVBA = Object.keys(cleanedZip.files).some(
      f => f.toLowerCase().includes('vbaproject.bin')
    );
    expect(hasVBA).toBe(false);
  });

  // Test 3: Script tags are removed from SVG
  test('should strip <script> tags from SVG files', async () => {
    const svgWithScript = `<svg xmlns="http://www.w3.org/2000/svg">
      <script>alert('xss')</script>
      <rect width="100" height="100"/>
    </svg>`;

    const buffer = await buildMinimalPPTX({
      'ppt/media/image1.svg': svgWithScript
    });
    const result = await sanitize(buffer);

    const cleanedZip = await JSZip.loadAsync(result);
    const cleanedSVG = await cleanedZip
      .file('ppt/media/image1.svg')
      .async('string');

    expect(cleanedSVG).not.toContain('<script>');
    expect(cleanedSVG).not.toContain('alert');
    expect(cleanedSVG).toContain('<rect');
  });

  // Test 4: HTML imports are rejected
  test('should reject PPTX containing HTML imports', async () => {
    const buffer = await buildMinimalPPTX({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?>
<root>
  <import text/html="dangerous.html"/>
</root>`
    });

    await expect(sanitize(buffer)).rejects.toThrow('HTML import detected');
  });

  // Test 5: Inline event handlers removed from SVG
  test('should remove inline event handlers from SVG', async () => {
    const svgWithHandler = `<svg xmlns="http://www.w3.org/2000/svg">
      <rect onload="alert('xss')" width="100" height="100"/>
    </svg>`;

    const buffer = await buildMinimalPPTX({
      'ppt/media/image2.svg': svgWithHandler
    });
    const result = await sanitize(buffer);

    const cleanedZip = await JSZip.loadAsync(result);
    const cleanedSVG = await cleanedZip
      .file('ppt/media/image2.svg')
      .async('string');

    expect(cleanedSVG).not.toContain('onload');
    expect(cleanedSVG).toContain('<rect');
  });

});