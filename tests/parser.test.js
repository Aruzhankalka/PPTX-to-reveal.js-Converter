const JSZip = require('jszip');
const { parsePptx } = require('../src/parser/pptx');
const { paragraphToIr, runToIr } = require('../src/parser/pptx/text');

/**
 * Build a minimal but valid .pptx in memory for testing.
 * This avoids needing a real .pptx fixture file in the repo.
 */
async function buildTestPptx({ slideCount = 1, withImage = false, withSchemeClr = false } = {}) {
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
  <p:sldSz cx="9144000" cy="5143500"/>
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
    const schemeClrXml = withSchemeClr && i === 1 ? `
          <a:p>
            <a:r>
              <a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr>
              <a:t>themed text</a:t>
            </a:r>
          </a:p>` : '';

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
          <a:p>
            <a:r>
              <a:rPr u="sng"/>
              <a:t>Underlined text on slide ${i}</a:t>
            </a:r>
          </a:p>${schemeClrXml}
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

  test('extracts text formatting — underline u="sng" from run rPr (FR-06)', async () => {
    const buf = await buildTestPptx({ slideCount: 1 });
    const { ir } = await parsePptx(buf);

    const paragraphs = ir.slideset.slides[0].contents.text[0].paragraphs;
    expect(paragraphs[2].runs[0].formatting['text-decoration']).toBe('underline');
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

  test('extracts slide canvas dimensions from <p:sldSz> (FR-07)', async () => {
    const buf = await buildTestPptx({ slideCount: 1 });
    const { ir } = await parsePptx(buf);
    expect(ir.slideset.master.slideWidth).toBe(960);
    expect(ir.slideset.master.slideHeight).toBe(540);
  });

  test('assigns z-index to text blocks in spTree order (FR-13)', async () => {
    const buf = await buildTestPptx({ slideCount: 1, withImage: true });
    const { ir } = await parsePptx(buf);
    const text = ir.slideset.slides[0].contents.text;
    const media = ir.slideset.slides[0].contents.media;
    expect(text[0]['z-index']).toBe(0);
    // media z-index starts after all text blocks
    expect(media[0]['z-index']).toBeGreaterThanOrEqual(text.length);
  });

  test('resolves schemeClr to CSS variable reference (FR-06)', async () => {
    const buf = await buildTestPptx({ slideCount: 1, withSchemeClr: true });
    const { ir } = await parsePptx(buf);
    const paragraphs = ir.slideset.slides[0].contents.text[0].paragraphs;
    const themedRun = paragraphs[3].runs[0]; // fourth paragraph added by withSchemeClr
    expect(themedRun.formatting.color).toBe('var(--theme-accent1)');
  });
});

// ---------------------------------------------------------------------------
// Helper: minimal PPTX with a single body placeholder whose XML we control
// ---------------------------------------------------------------------------
async function buildPptxWithBodyShape(bodySpXml) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
  zip.file('_rels/.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);
  zip.file('ppt/presentation.xml', `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>${bodySpXml}</p:spTree></p:cSld>
</p:sld>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

// Minimal body <p:sp> with controlled bodyPr and a single paragraph
function makeBodySpXml({ normAutofitXml = '', paraXml = '' } = {}) {
  const bodyPrContent = normAutofitXml
    ? `<a:bodyPr>${normAutofitXml}</a:bodyPr>`
    : `<a:bodyPr/>`;
  return `<p:sp>
    <p:nvSpPr>
      <p:cNvSpPr/>
      <p:nvPr><p:ph type="body"/></p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="100000" y="100000"/><a:ext cx="5000000" cy="3000000"/></a:xfrm>
    </p:spPr>
    <p:txBody>
      ${bodyPrContent}<a:lstStyle/>
      <a:p>${paraXml}<a:r><a:rPr sz="2400"/><a:t>text</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>`;
}

// ---------------------------------------------------------------------------
// normAutofit — fontScale and lnSpcReduction applied at parse time
// ---------------------------------------------------------------------------
describe('normAutofit — fontScale reduces run sizes', () => {
  test('fontScale=62500 scales a 24pt run to 15pt (62.5% of nominal)', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ normAutofitXml: '<a:normAutofit fontScale="62500"/>' })
    );
    const { ir } = await parsePptx(buf);
    const run = ir.slideset.slides[0].contents.text[0].paragraphs[0].runs[0];
    expect(run.formatting.size).toBe('15pt');
  });

  test('fontScale=100000 (100%) leaves run sizes unchanged', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ normAutofitXml: '<a:normAutofit fontScale="100000"/>' })
    );
    const { ir } = await parsePptx(buf);
    const run = ir.slideset.slides[0].contents.text[0].paragraphs[0].runs[0];
    expect(run.formatting.size).toBe('24pt');
  });
});

describe('normAutofit — lnSpcReduction reduces line-height', () => {
  test('lnSpcReduction=20000 multiplies explicit line-height 1.5 by 0.8 → 1.2', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({
        normAutofitXml: '<a:normAutofit lnSpcReduction="20000"/>',
        paraXml: '<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>',
      })
    );
    const { ir } = await parsePptx(buf);
    const para = ir.slideset.slides[0].contents.text[0].paragraphs[0];
    expect(para.formatting['line-spacing']).toBe('1.2');
  });

  test('lnSpcReduction=20000 reduces from default 1.0 → 0.8 when no explicit lnSpc', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ normAutofitXml: '<a:normAutofit lnSpcReduction="20000"/>' })
    );
    const { ir } = await parsePptx(buf);
    const para = ir.slideset.slides[0].contents.text[0].paragraphs[0];
    expect(para.formatting['line-spacing']).toBe('0.8');
  });
});

// ---------------------------------------------------------------------------
// <a:lnSpc> — paragraph line spacing parsing
// ---------------------------------------------------------------------------
describe('<a:lnSpc> paragraph line spacing', () => {
  test('spcPct val=150000 → line-spacing "1.5" (unitless CSS)', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ paraXml: '<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>' })
    );
    const { ir } = await parsePptx(buf);
    const para = ir.slideset.slides[0].contents.text[0].paragraphs[0];
    expect(para.formatting['line-spacing']).toBe('1.5');
  });

  test('spcPct val=100000 → line-spacing "1" (100% = unitless 1.0)', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ paraXml: '<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>' })
    );
    const { ir } = await parsePptx(buf);
    const para = ir.slideset.slides[0].contents.text[0].paragraphs[0];
    expect(para.formatting['line-spacing']).toBe('1');
  });

  test('spcPts val=2400 → line-spacing "24pt" (exact points)', async () => {
    const buf = await buildPptxWithBodyShape(
      makeBodySpXml({ paraXml: '<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc></a:pPr>' })
    );
    const { ir } = await parsePptx(buf);
    const para = ir.slideset.slides[0].contents.text[0].paragraphs[0];
    expect(para.formatting['line-spacing']).toBe('24pt');
  });
});

// ---------------------------------------------------------------------------
// txStyles spacing fallback — unit-tested directly via paragraphToIr
// ---------------------------------------------------------------------------
describe('paragraphToIr — txStyles spacing fallback', () => {
  const txStyles = {
    body: {
      1: { size: '22pt', lineSpacing: '0.9', spaceBefore: '10pt', spaceAfter: '0pt' },
    },
  };

  test('paragraph with no explicit lnSpc inherits lineSpacing from txStyles', () => {
    const para = paragraphToIr({}, 0, null, 'body', txStyles);
    expect(para.formatting['line-spacing']).toBe('0.9');
  });

  test('paragraph with no explicit spcBef inherits spaceBefore from txStyles', () => {
    const para = paragraphToIr({}, 0, null, 'body', txStyles);
    expect(para.formatting['space-before']).toBe('10pt');
  });

  test('explicit <a:lnSpc> on the paragraph takes precedence over txStyles', () => {
    const aP = { 'a:pPr': { 'a:lnSpc': { 'a:spcPct': { '@_val': '200000' } } } };
    const para = paragraphToIr(aP, 0, null, 'body', txStyles);
    expect(para.formatting['line-spacing']).toBe('2');
  });

  test('null txStyles leaves paragraph formatting driven only by its own pPr', () => {
    const para = paragraphToIr({}, 0, null, 'body', null);
    expect(para.formatting).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// <a:fld> field elements are collected as runs in paragraphToIr
// ---------------------------------------------------------------------------

describe('paragraphToIr — <a:fld> field elements collected as runs', () => {
  test('slide number field text is included as a run', () => {
    const aP = { 'a:fld': { '@_type': 'slidenum', '@_id': '{G1}', 'a:t': '6' } };
    const para = paragraphToIr(aP, 0, null, 'sldNum', null);
    expect(para.runs.some((r) => r.text === '6')).toBe(true);
  });

  test('date field text is included as a run', () => {
    const aP = { 'a:fld': { '@_type': 'datetime', '@_id': '{G2}', 'a:t': '01/15/2026' } };
    const para = paragraphToIr(aP, 0, null, 'dt', null);
    expect(para.runs.some((r) => r.text === '01/15/2026')).toBe(true);
  });

  test('paragraph with both <a:r> and <a:fld> collects text from both', () => {
    const aP = {
      'a:r':  { 'a:t': 'prefix' },
      'a:fld': { '@_type': 'slidenum', 'a:t': '3' },
    };
    const para = paragraphToIr(aP, 0, null, null, null);
    const allText = para.runs.map((r) => r.text);
    expect(allText).toContain('prefix');
    expect(allText).toContain('3');
  });
});

// ---------------------------------------------------------------------------
// ftr/sldNum/dt placeholder types do NOT inherit txStyles.body size/spacing
// ---------------------------------------------------------------------------

describe('paragraphToIr — ftr/sldNum/dt do not inherit txStyles.body size/spacing', () => {
  const txStyles = {
    body: { 1: { size: '22pt', lineSpacing: '0.9', spaceBefore: '10pt', spaceAfter: '0pt' } },
  };

  test('ftr run has no size when txStyles would otherwise supply 22pt', () => {
    const para = paragraphToIr({ 'a:r': { 'a:t': 'footer' } }, 0, null, 'ftr', txStyles);
    const size = para.runs[0].formatting && para.runs[0].formatting.size;
    expect(size).toBeUndefined();
  });

  test('ftr paragraph has no spacing from txStyles.body', () => {
    const para = paragraphToIr({}, 0, null, 'ftr', txStyles);
    expect(para.formatting).toBeUndefined();
  });

  test('sldNum run has no size from txStyles.body', () => {
    const para = paragraphToIr({ 'a:r': { 'a:t': '6' } }, 0, null, 'sldNum', txStyles);
    const size = para.runs[0].formatting && para.runs[0].formatting.size;
    expect(size).toBeUndefined();
  });

  test('dt run has no size from txStyles.body', () => {
    const para = paragraphToIr({ 'a:r': { 'a:t': '01/01' } }, 0, null, 'dt', txStyles);
    const size = para.runs[0].formatting && para.runs[0].formatting.size;
    expect(size).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Bold/italic/underline inheritance from <a:defRPr> (FR-06)
// ---------------------------------------------------------------------------

describe('paragraphToIr — inherits bold/italic/underline from defRPr cascade', () => {
  test('bold from paragraph <a:defRPr> is inherited by runs with no explicit rPr', () => {
    const para = {
      'a:pPr': { 'a:defRPr': { '@_b': '1' } },
      'a:r': [
        { 'a:t': 'run without rPr' },
        { 'a:rPr': { '@_sz': '2400' }, 'a:t': 'run with size only' },
      ],
    };
    const result = paragraphToIr(para, 0, null, 'body', null);
    expect(result.runs[0].formatting.weight).toBe('bold');
    expect(result.runs[1].formatting.weight).toBe('bold');
  });

  test('italic from paragraph <a:defRPr> is inherited', () => {
    const para = {
      'a:pPr': { 'a:defRPr': { '@_i': '1' } },
      'a:r': [{ 'a:t': 'italic run' }],
    };
    const result = paragraphToIr(para, 0, null, 'body', null);
    expect(result.runs[0].formatting.italics).toBe(true);
  });

  test('underline from paragraph <a:defRPr> is inherited', () => {
    const para = {
      'a:pPr': { 'a:defRPr': { '@_u': 'sng' } },
      'a:r': [{ 'a:t': 'underlined run' }],
    };
    const result = paragraphToIr(para, 0, null, 'body', null);
    expect(result.runs[0].formatting['text-decoration']).toBe('underline');
  });

  test('run with b="0" is not bold even when paraDefRPr has b="1"', () => {
    const para = {
      'a:pPr': { 'a:defRPr': { '@_b': '1' } },
      'a:r': [{ 'a:rPr': { '@_b': '0' }, 'a:t': 'explicitly not bold' }],
    };
    const result = paragraphToIr(para, 0, null, 'body', null);
    expect(result.runs[0].formatting).not.toHaveProperty('weight');
  });

  test('bold from lstStyle <a:defRPr> is inherited when paraDefRPr is absent', () => {
    const lstStyle = { 'a:lvl1pPr': { 'a:defRPr': { '@_b': '1' } } };
    const para = { 'a:r': [{ 'a:t': 'lstStyle bold' }] };
    const result = paragraphToIr(para, 0, lstStyle, 'body', null);
    expect(result.runs[0].formatting.weight).toBe('bold');
  });

  test('paraDefRPr bold takes precedence over lstStyle bold when both set', () => {
    const lstStyle = { 'a:lvl1pPr': { 'a:defRPr': { '@_b': '1' } } };
    const para = {
      'a:pPr': { 'a:defRPr': { '@_b': '1', '@_i': '1' } },
      'a:r': [{ 'a:t': 'both bold and italic' }],
    };
    const result = paragraphToIr(para, 0, lstStyle, 'body', null);
    expect(result.runs[0].formatting.weight).toBe('bold');
    expect(result.runs[0].formatting.italics).toBe(true);
  });

  test('b="true" string value is recognised as bold (OOXML ST_TextBooleanType)', () => {
    const para = {
      'a:pPr': { 'a:defRPr': { '@_b': 'true' } },
      'a:r': [{ 'a:t': 'bold via true string' }],
    };
    const result = paragraphToIr(para, 0, null, 'body', null);
    expect(result.runs[0].formatting.weight).toBe('bold');
  });

  test('underline from txStyles entry is inherited by runs when no defRPr overrides it', () => {
    const txStyles = { body: { 1: { underline: true } } };
    const para = { 'a:r': [{ 'a:t': 'underlined from master' }] };
    const result = paragraphToIr(para, 0, null, 'body', txStyles);
    expect(result.runs[0].formatting['text-decoration']).toBe('underline');
  });

  test('lstStyle u="none" clears txStyles underline (explicit override)', () => {
    const txStyles = { body: { 1: { underline: true } } };
    const lstStyle = { 'a:lvl1pPr': { 'a:defRPr': { '@_u': 'none' } } };
    const para = { 'a:r': [{ 'a:t': 'not underlined' }] };
    const result = paragraphToIr(para, 0, lstStyle, 'body', txStyles);
    expect(result.runs[0].formatting).not.toHaveProperty('text-decoration');
  });

  test('txStyles bold is cleared by paraDefRPr b="0"', () => {
    const txStyles = { body: { 1: { bold: true } } };
    const para = {
      'a:pPr': { 'a:defRPr': { '@_b': '0' } },
      'a:r': [{ 'a:t': 'not bold' }],
    };
    const result = paragraphToIr(para, 0, null, 'body', txStyles);
    expect(result.runs[0].formatting).not.toHaveProperty('weight');
  });

  test('txStyles italic is not applied when phType is ftr', () => {
    const txStyles = { body: { 1: { italic: true } } };
    const para = { 'a:r': [{ 'a:t': 'footer text' }] };
    const result = paragraphToIr(para, 0, null, 'ftr', txStyles);
    expect(result.runs[0].formatting?.italics).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// schemeClr alias normalization in extractRunFormatting (tx1→dk1, bg1→lt1)
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// runToIr — underline directly from run's own <a:rPr> (FR-06)
// ---------------------------------------------------------------------------
describe('runToIr — underline from run rPr (FR-06)', () => {
  test('u="sng" produces text-decoration: underline', () => {
    const run = { 'a:rPr': { '@_u': 'sng' }, 'a:t': 'underlined' };
    const result = runToIr(run, null, null);
    expect(result.formatting['text-decoration']).toBe('underline');
  });

  test('u="dbl" also produces text-decoration: underline', () => {
    const run = { 'a:rPr': { '@_u': 'dbl' }, 'a:t': 'double underline' };
    const result = runToIr(run, null, null);
    expect(result.formatting['text-decoration']).toBe('underline');
  });

  test('u="none" produces no text-decoration', () => {
    const run = { 'a:rPr': { '@_u': 'none' }, 'a:t': 'not underlined' };
    const result = runToIr(run, null, null);
    expect(result.formatting?.['text-decoration']).toBeUndefined();
  });

  test('bold and underline together do not interfere', () => {
    const run = { 'a:rPr': { '@_b': '1', '@_u': 'sng' }, 'a:t': 'bold+underline' };
    const result = runToIr(run, null, null);
    expect(result.formatting.weight).toBe('bold');
    expect(result.formatting['text-decoration']).toBe('underline');
  });
});

// ---------------------------------------------------------------------------

describe('runToIr — schemeClr aliases are normalized to theme map keys', () => {
  function makeRunWithScheme(val) {
    return { 'a:rPr': { 'a:solidFill': { 'a:schemeClr': { '@_val': val } } }, 'a:t': 'x' };
  }

  test('tx1 is normalized to var(--theme-dk1)', () => {
    const run = runToIr(makeRunWithScheme('tx1'), null);
    expect(run.formatting.color).toBe('var(--theme-dk1)');
  });

  test('tx2 is normalized to var(--theme-dk2)', () => {
    const run = runToIr(makeRunWithScheme('tx2'), null);
    expect(run.formatting.color).toBe('var(--theme-dk2)');
  });

  test('bg1 is normalized to var(--theme-lt1)', () => {
    const run = runToIr(makeRunWithScheme('bg1'), null);
    expect(run.formatting.color).toBe('var(--theme-lt1)');
  });

  test('bg2 is normalized to var(--theme-lt2)', () => {
    const run = runToIr(makeRunWithScheme('bg2'), null);
    expect(run.formatting.color).toBe('var(--theme-lt2)');
  });

  test('non-alias scheme colors pass through unchanged', () => {
    const run = runToIr(makeRunWithScheme('accent1'), null);
    expect(run.formatting.color).toBe('var(--theme-accent1)');
  });

  test('dk1 passes through unchanged (already the canonical key)', () => {
    const run = runToIr(makeRunWithScheme('dk1'), null);
    expect(run.formatting.color).toBe('var(--theme-dk1)');
  });
});

// ---------------------------------------------------------------------------
// Empty paragraph preservation — blank-line spacing (FR: preserve spacing)
// ---------------------------------------------------------------------------

describe('paragraphToIr — empty paragraph preservation', () => {
  // (a) An empty <a:p> (no <a:r>) is preserved as { runs: [] }, not filtered out.
  test('preserves an empty <a:p> as runs:[] with size from <a:endParaRPr @_sz>', () => {
    // Simulate <a:p><a:endParaRPr sz="2800"/></a:p> (28pt, no runs)
    const aP = { 'a:endParaRPr': { '@_sz': '2800' } };
    const para = paragraphToIr(aP, 0, null, null, null, null, null, true);
    expect(para.runs).toHaveLength(0);
    expect(para.formatting && para.formatting.size).toBe('28pt');
  });

  test('empty <a:p> with no @_sz falls back to lstStyle size', () => {
    const lstStyle = {
      'a:lvl1pPr': { 'a:defRPr': { '@_sz': '3200' } }, // 32pt from lstStyle
    };
    const aP = { 'a:endParaRPr': {} }; // no @_sz
    const para = paragraphToIr(aP, 0, lstStyle, null, null, null, null, true);
    expect(para.runs).toHaveLength(0);
    expect(para.formatting && para.formatting.size).toBe('32pt');
  });

  test('empty <a:p> with no size source leaves formatting.size absent', () => {
    const aP = {}; // no endParaRPr, no lstStyle, no txStyles
    const para = paragraphToIr(aP, 0, null, null, null, null, null, true);
    expect(para.runs).toHaveLength(0);
    expect(para.formatting && para.formatting.size).toBeUndefined();
  });
});