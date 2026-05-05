const JSZip = require('jszip');
const { parsePptx } = require('../src/parser/pptx');

/**
 * Build a minimal but valid .pptx in memory for testing.
 * This avoids needing a real .pptx fixture file in the repo.
 */
async function buildTestPptx({ slideCount = 1, withImage = false } = {}) {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);

  zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  // presentation.xml with N slides
  let sldIds = '';
  for (let i = 1; i <= slideCount; i++) {
    sldIds += `<p:sldId id="${256 + i}" r:id="rId${i}"/>`;
  }
  zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>${sldIds}</p:sldIdLst>
</p:presentation>`);

  // presentation.xml.rels
  let presRels = '';
  for (let i = 1; i <= slideCount; i++) {
    presRels += `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
  }
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${presRels}
</Relationships>`);

  // Slide files
  for (let i = 1; i <= slideCount; i++) {
    const picXml = withImage && i === 1 ? `
      <p:pic>
        <p:blipFill><a:blip r:embed="rId10"/></p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="1905000" y="1905000"/>
            <a:ext cx="4762500" cy="3571875"/>
          </a:xfrm>
        </p:spPr>
      </p:pic>` : '';

    zip.file(`ppt/slides/slide${i}.xml`, `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="685800"/>
            <a:ext cx="7315200" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="4400"/>
              <a:t>Slide ${i} title</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:r>
              <a:rPr i="1"/>
              <a:t>Body text on slide ${i}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>${picXml}
    </p:spTree>
  </p:cSld>
</p:sld>`);

    if (withImage && i === 1) {
      zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`);
      // Tiny 1x1 PNG bytes
      const png = Buffer.from('89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D4944415478DA63F8FFFFFF3F0005FE02FEDCCC59E70000000049454E44AE426082', 'hex');
      zip.file('ppt/media/image1.png', png);
    }
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('parsePptx', () => {
  test('produces valid IR for a single-slide pptx (FR-03, FR-04)', async () => {
    const buf = await buildTestPptx({ slideCount: 1 });
    const { ir } = await parsePptx(buf, { filename: 'one.pptx' });

    expect(ir.slideset.slides).toHaveLength(1);
    expect(ir.slideset.filename).toBe('one.pptx');
  });

  test('preserves slide order across multiple slides (FR-04)', async () => {
    const buf = await buildTestPptx({ slideCount: 3 });
    const { ir } = await parsePptx(buf);

    expect(ir.slideset.slides).toHaveLength(3);
    const titles = ir.slideset.slides.map((s) => s.title);
    expect(titles).toEqual(['Slide 1 title', 'Slide 2 title', 'Slide 3 title']);
  });

  test('extracts text formatting — bold and italic (FR-06 subset)', async () => {
    const buf = await buildTestPptx({ slideCount: 1 });
    const { ir } = await parsePptx(buf);

    const paragraphs = ir.slideset.slides[0].contents.text[0].paragraphs;
    expect(paragraphs[0].runs[0].formatting.weight).toBe('bold');
    expect(paragraphs[1].runs[0].formatting.italics).toBe(true);
  });

  test('extracts an image and its bytes (FR-09)', async () => {
    const buf = await buildTestPptx({ slideCount: 1, withImage: true });
    const { ir, media } = await parsePptx(buf);

    expect(ir.slideset.slides[0].contents.media).toHaveLength(1);
    expect(ir.slideset.slides[0].contents.media[0]['media-type']).toBe('image');
    expect(ir.slideset.slides[0].contents.media[0]['file-link']).toMatch(/^media\/image1\.png$/);
    expect(media).toHaveLength(1);
    expect(media[0].bytes).toBeInstanceOf(Buffer);
  });

  test('rejects a non-zip buffer with INVALID_PPTX', async () => {
    await expect(parsePptx(Buffer.from('not a zip'))).rejects.toMatchObject({
      code: 'INVALID_PPTX',
    });
  });

  test('rejects a zip without presentation.xml', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'hi');
    const buf = await zip.generateAsync({ type: 'nodebuffer' });

    await expect(parsePptx(buf)).rejects.toMatchObject({
      code: 'INVALID_PPTX',
    });
  });

  test('produces IR that passes schema validation (TC-05)', async () => {
    const buf = await buildTestPptx({ slideCount: 2, withImage: true });
    // If this didn't throw, it passed validation inside parsePptx
    const { ir } = await parsePptx(buf);
    expect(ir).toBeDefined();
  });
});