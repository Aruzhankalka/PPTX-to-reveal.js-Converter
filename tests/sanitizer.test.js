'use strict';

const JSZip = require('jszip');
const { sanitize } = require('../src/security/sanitizer');

function buildMinimalPPTX(extraFiles = {}) {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
</p:presentation>`);

  for (const [path, content] of Object.entries(extraFiles)) {
    zip.file(path, content);
  }

  return zip;
}

describe('Sanitizer', () => {

  test('should pass a clean PPTX without modifications', async () => {
    const zip = buildMinimalPPTX();
    await sanitize(zip);
    expect(zip.file('ppt/presentation.xml')).not.toBeNull();
  });

  test('should remove vbaProject.bin from PPTX', async () => {
    const zip = buildMinimalPPTX({
      'ppt/vbaProject.bin': Buffer.from('fake vba content')
    });
    await sanitize(zip);
    const hasVBA = Object.keys(zip.files).some(
      f => f.toLowerCase().includes('vbaproject.bin')
    );
    expect(hasVBA).toBe(false);
  });

  test('should strip <script> tags from SVG files', async () => {
    const svgWithScript = `<svg xmlns="http://www.w3.org/2000/svg">
      <script>alert('xss')</script>
      <rect width="100" height="100"/>
    </svg>`;

    const zip = buildMinimalPPTX({
      'ppt/media/image1.svg': svgWithScript
    });
    await sanitize(zip);
    const cleanedSVG = await zip.file('ppt/media/image1.svg').async('string');

    expect(cleanedSVG).not.toContain('<script>');
    expect(cleanedSVG).not.toContain('alert');
    expect(cleanedSVG).toContain('<rect');
  });

  test('should reject PPTX containing HTML imports', async () => {
    const zip = buildMinimalPPTX({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?>
<root>
  <import text/html="dangerous.html"/>
</root>`
    });
    await expect(sanitize(zip)).rejects.toThrow('HTML import detected');
  });

  test('should remove inline event handlers from SVG', async () => {
    const svgWithHandler = `<svg xmlns="http://www.w3.org/2000/svg">
      <rect onload="alert('xss')" width="100" height="100"/>
    </svg>`;

    const zip = buildMinimalPPTX({
      'ppt/media/image2.svg': svgWithHandler
    });
    await sanitize(zip);
    const cleanedSVG = await zip.file('ppt/media/image2.svg').async('string');

    expect(cleanedSVG).not.toContain('onload');
    expect(cleanedSVG).toContain('<rect');
  });

});
