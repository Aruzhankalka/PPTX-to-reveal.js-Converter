'use strict';

const JSZip = require('jszip');
const { extractXfrm, collectPlaceholders, loadLayoutGeometry, lookupGeo, collectLayoutMedia } =
  require('../src/parser/pptx/layouts');
const { parsePptx } = require('../src/parser/pptx');

// ---------------------------------------------------------------------------
// Test constants — standard widescreen PPTX dimensions
// ---------------------------------------------------------------------------
const SLIDE_W_EMU  = 9144000;
const SLIDE_H_EMU  = 5143500;

// Footer position typical of a "Title and Content" layout
const FTR_X_EMU  = 457200;   //  48 px
const FTR_Y_EMU  = 4940550;  // 519 px
const FTR_CX_EMU = 2286000;  // 240 px
const FTR_CY_EMU =  457200;  //  48 px

// EMU→px conversion must match units.js (÷ 9525)
const toP = (emu) => Math.round(emu / 9525);

// ---------------------------------------------------------------------------
// Helpers: raw XML fragments used across several tests
// ---------------------------------------------------------------------------

function phSpXml({ idx, type, withXfrm, x, y, cx, cy, anchor, normAutofit } = {}) {
  const phAttrs = [
    type  ? `type="${type}"` : '',
    idx   !== undefined ? `idx="${idx}"` : '',
  ].filter(Boolean).join(' ');

  const xfrmXml = withXfrm
    ? `<p:spPr>
        <a:xfrm>
          <a:off x="${x ?? FTR_X_EMU}" y="${y ?? FTR_Y_EMU}"/>
          <a:ext cx="${cx ?? FTR_CX_EMU}" cy="${cy ?? FTR_CY_EMU}"/>
        </a:xfrm>
      </p:spPr>`
    : `<p:spPr/>`; // no <a:xfrm> — position must be resolved from layout/master

  // anchor is the PPTX text-body vertical anchor: "t", "ctr", or "b"
  const bodyPrAttr = anchor ? ` anchor="${anchor}"` : '';
  // normAutofit: { fontScale, lnSpcReduction } both as integer 1000ths-of-percent values
  const normAutofitXml = normAutofit
    ? `<a:normAutofit fontScale="${normAutofit.fontScale}"${normAutofit.lnSpcReduction ? ` lnSpcReduction="${normAutofit.lnSpcReduction}"` : ''}/>`
    : '';
  const bodyPrXml = normAutofitXml
    ? `<a:bodyPr${bodyPrAttr}>${normAutofitXml}</a:bodyPr>`
    : `<a:bodyPr${bodyPrAttr}/>`;

  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:nvSpPr>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph ${phAttrs}/></p:nvPr>
    </p:nvSpPr>
    ${xfrmXml}
    <p:txBody>${bodyPrXml}<a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
  </p:sp>`;
}

// Wrap an array of <p:sp> fragments in a minimal spTree node parsed by fast-xml-parser
function makeSpTree(spXmls) {
  const { XMLParser } = require('fast-xml-parser');
  const parser = new XMLParser({
    ignoreAttributes: false, attributeNamePrefix: '@_',
    parseTagValue: false, parseAttributeValue: false,
    trimValues: false, removeNSPrefix: false,
  });
  const xml = `<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    ${spXmls.join('\n')}
  </p:spTree>`;
  return parser.parse(xml)['p:spTree'];
}

// ---------------------------------------------------------------------------
// extractXfrm
// ---------------------------------------------------------------------------

describe('extractXfrm', () => {
  function parseSp(xml) {
    const { XMLParser } = require('fast-xml-parser');
    const p = new XMLParser({
      ignoreAttributes: false, attributeNamePrefix: '@_',
      parseTagValue: false, parseAttributeValue: false,
      trimValues: false, removeNSPrefix: false,
    });
    return p.parse(xml)['p:sp'];
  }

  test('returns geo with position/width/height from <a:xfrm>', () => {
    const sp = parseSp(phSpXml({ withXfrm: true,
      x: FTR_X_EMU, y: FTR_Y_EMU, cx: FTR_CX_EMU, cy: FTR_CY_EMU }));
    const geo = extractXfrm(sp);
    expect(geo).not.toBeNull();
    expect(geo.position.x).toBe(toP(FTR_X_EMU));
    expect(geo.position.y).toBe(toP(FTR_Y_EMU));
    expect(geo.width).toBe(toP(FTR_CX_EMU));
    expect(geo.height).toBe(toP(FTR_CY_EMU));
  });

  test('returns null when <p:spPr> has no <a:xfrm>', () => {
    const sp = parseSp(phSpXml({ withXfrm: false }));
    expect(extractXfrm(sp)).toBeNull();
  });

  test('does not set rotation when @_rot is absent', () => {
    const sp = parseSp(phSpXml({ withXfrm: true }));
    expect(extractXfrm(sp).rotation).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// collectPlaceholders
// ---------------------------------------------------------------------------

describe('collectPlaceholders', () => {
  test('indexes by idx and type for a placeholder with xfrm', () => {
    const spTree = makeSpTree([phSpXml({ type: 'ftr', idx: 11, withXfrm: true })]);
    const { byIdx, byType } = collectPlaceholders(spTree);
    expect(byIdx.has(11)).toBe(true);
    expect(byType.has('ftr')).toBe(true);
    expect(byIdx.get(11).position.y).toBe(toP(FTR_Y_EMU));
  });

  test('placeholder without xfrm is NOT indexed', () => {
    const spTree = makeSpTree([phSpXml({ type: 'ftr', idx: 11, withXfrm: false })]);
    const { byIdx, byType } = collectPlaceholders(spTree);
    expect(byIdx.has(11)).toBe(false);
    expect(byType.has('ftr')).toBe(false);
  });

  test('absent idx defaults to 0', () => {
    // No idx attribute → title placeholder at idx=0 per OOXML §19.2.1.27
    const spTree = makeSpTree([phSpXml({ type: 'title', withXfrm: true })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.has(0)).toBe(true);
  });

  test('multiple placeholders all indexed independently', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'title', idx: 0,  withXfrm: true,  x: 100,  y: 200,  cx: 300, cy: 100 }),
      phSpXml({ type: 'ftr',   idx: 11, withXfrm: true,  x: FTR_X_EMU, y: FTR_Y_EMU,
                                                          cx: FTR_CX_EMU, cy: FTR_CY_EMU }),
    ]);
    const { byIdx, byType } = collectPlaceholders(spTree);
    expect(byIdx.size).toBe(2);
    expect(byType.size).toBe(2);
    expect(byIdx.has(0)).toBe(true);
    expect(byIdx.has(11)).toBe(true);
  });

  test('empty spTree returns empty maps', () => {
    const spTree = makeSpTree([]);
    const { byIdx, byType } = collectPlaceholders(spTree);
    expect(byIdx.size).toBe(0);
    expect(byType.size).toBe(0);
  });

  test('textAnchor is extracted from <a:bodyPr anchor="ctr">', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'ftr', idx: 11, withXfrm: true, anchor: 'ctr' }),
    ]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).textAnchor).toBe('ctr');
  });

  test('textAnchor is "b" when <a:bodyPr anchor="b">', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'ftr', idx: 11, withXfrm: true, anchor: 'b' }),
    ]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).textAnchor).toBe('b');
  });

  test('textAnchor is undefined when <a:bodyPr> has no anchor attribute', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'ftr', idx: 11, withXfrm: true }),  // no anchor
    ]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).textAnchor).toBeUndefined();
  });

  test('normAutofit fontScale=62500 and lnSpcReduction=20000 are stored in geo (as fractions)', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'body', idx: 1, withXfrm: true,
                normAutofit: { fontScale: 62500, lnSpcReduction: 20000 } }),
    ]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(1).normAutofit).toBeDefined();
    // fontScale: 62500 / 100000 = 0.625
    expect(byIdx.get(1).normAutofit.fontScale).toBeCloseTo(0.625);
    // lnSpcRed: 20000 / 100000 = 0.2
    expect(byIdx.get(1).normAutofit.lnSpcRed).toBeCloseTo(0.2);
  });

  test('placeholder without <a:normAutofit> has undefined geo.normAutofit', () => {
    const spTree = makeSpTree([
      phSpXml({ type: 'body', idx: 1, withXfrm: true }),  // no normAutofit
    ]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(1).normAutofit).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// lookupGeo
// ---------------------------------------------------------------------------

describe('lookupGeo', () => {
  const geo11 = { position: { x: 48, y: 519 }, width: 240, height: 48 };
  const geoFtr = { position: { x: 0, y: 520 }, width: 300, height: 40 };

  let geoMap;
  beforeEach(() => {
    geoMap = { byIdx: new Map([[11, geo11]]), byType: new Map([['ftr', geoFtr]]) };
  });

  test('returns byIdx result when idx matches', () => {
    expect(lookupGeo(geoMap, 11, 'ftr')).toBe(geo11);
  });

  test('falls through to byType when idx is not in map', () => {
    expect(lookupGeo(geoMap, 99, 'ftr')).toBe(geoFtr);
  });

  test('returns null when neither idx nor type match', () => {
    expect(lookupGeo(geoMap, 99, 'body')).toBeNull();
  });

  test('handles null phType gracefully', () => {
    expect(lookupGeo(geoMap, 99, null)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// loadLayoutGeometry (async, uses in-memory JSZip)
// ---------------------------------------------------------------------------

// Build the layout XML with a footer placeholder that has an explicit xfrm
function makeLayoutXml(withFooterXfrm = true, extraSpXml = '', ftrAnchor = '') {
  const bodyPrAttr = ftrAnchor ? ` anchor="${ftrAnchor}"` : '';
  const ftrSp = withFooterXfrm
    ? `<p:sp>
        <p:nvSpPr>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${FTR_X_EMU}" y="${FTR_Y_EMU}"/>
            <a:ext cx="${FTR_CX_EMU}" cy="${FTR_CY_EMU}"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody><a:bodyPr${bodyPrAttr}/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
      </p:sp>`
    : '';
  return `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             type="titleContent">
  <p:cSld name="Title and Content">
    <p:spTree>
      ${ftrSp}
      ${extraSpXml}
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;
}

function makeMasterXml(withTitleXfrm = true) {
  const titleSp = withTitleXfrm
    ? `<p:sp>
        <p:nvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="685800"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>`
    : '';
  return `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Office Theme">
    <p:spTree>
      ${titleSp}
    </p:spTree>
  </p:cSld>
</p:sldMaster>`;
}

async function buildZipWithLayout({ layoutHasFtr = true, masterHasTitle = true } = {}) {
  const zip = new JSZip();
  zip.file('ppt/slideLayouts/slideLayout3.xml', makeLayoutXml(layoutHasFtr));
  zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', makeMasterXml(masterHasTitle));
  return zip;
}

describe('loadLayoutGeometry', () => {
  test('returns empty maps when layoutPath is null', async () => {
    const zip = new JSZip();
    const { byIdx, byType } = await loadLayoutGeometry(zip, null);
    expect(byIdx.size).toBe(0);
    expect(byType.size).toBe(0);
  });

  test('returns empty maps when layout file is missing from ZIP', async () => {
    const zip = new JSZip(); // no files
    const { byIdx, byType } = await loadLayoutGeometry(zip, 'ppt/slideLayouts/missing.xml');
    expect(byIdx.size).toBe(0);
    expect(byType.size).toBe(0);
  });

  test('reads footer geometry from layout (byIdx and byType)', async () => {
    const zip = await buildZipWithLayout({ layoutHasFtr: true });
    const { byIdx, byType } = await loadLayoutGeometry(zip,
      'ppt/slideLayouts/slideLayout3.xml');
    expect(byIdx.has(11)).toBe(true);
    expect(byType.has('ftr')).toBe(true);
    expect(byIdx.get(11).position.x).toBe(toP(FTR_X_EMU));
    expect(byIdx.get(11).position.y).toBe(toP(FTR_Y_EMU));
    expect(byIdx.get(11).width).toBe(toP(FTR_CX_EMU));
    expect(byIdx.get(11).height).toBe(toP(FTR_CY_EMU));
  });

  test('falls back to master for placeholder absent from layout', async () => {
    // Layout has no footer, but master has title — title should come from master
    const zip = await buildZipWithLayout({ layoutHasFtr: false, masterHasTitle: true });
    const { byType } = await loadLayoutGeometry(zip,
      'ppt/slideLayouts/slideLayout3.xml');
    expect(byType.has('title')).toBe(true);
    expect(byType.get('title').position.x).toBe(toP(914400));
  });

  test('layout entry takes precedence over master entry for the same type', async () => {
    // Layout has title at x=100; master also has title but at x=914400
    const layoutWithTitle = makeLayoutXml(false,
      `<p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="100" y="200"/>
            <a:ext cx="300" cy="100"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>`);
    const zip = new JSZip();
    zip.file('ppt/slideLayouts/slideLayout3.xml', layoutWithTitle);
    zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
    zip.file('ppt/slideMasters/slideMaster1.xml', makeMasterXml(true));

    const { byType } = await loadLayoutGeometry(zip, 'ppt/slideLayouts/slideLayout3.xml');
    // x=100 EMU → emuToPx(100) = Math.round(100/9525) = 0 (rounds to 0 at that scale)
    // The important thing is that the value came from the layout (x=100) not master (x=914400)
    expect(byType.get('title').position.x).not.toBe(toP(914400));
  });

  test('gracefully handles missing layout rels (no master fallback)', async () => {
    const zip = new JSZip();
    zip.file('ppt/slideLayouts/slideLayout3.xml', makeLayoutXml(true));
    // No _rels file
    const { byIdx } = await loadLayoutGeometry(zip, 'ppt/slideLayouts/slideLayout3.xml');
    // Layout geometry still loaded even without rels
    expect(byIdx.has(11)).toBe(true);
  });

  test('textAnchor="ctr" in layout footer is stored in the geometry object', async () => {
    const zip = new JSZip();
    zip.file('ppt/slideLayouts/slideLayout3.xml', makeLayoutXml(true, '', 'ctr'));
    zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
    zip.file('ppt/slideMasters/slideMaster1.xml', makeMasterXml(false));
    const { byIdx } = await loadLayoutGeometry(zip, 'ppt/slideLayouts/slideLayout3.xml');
    expect(byIdx.get(11).textAnchor).toBe('ctr');
  });
});

// ---------------------------------------------------------------------------
// End-to-end: parsePptx resolves layout placeholder geometry (FR-11)
// ---------------------------------------------------------------------------

/**
 * Build a PPTX where:
 * - slide1 has a footer placeholder with NO <a:xfrm> (must inherit from layout)
 * - layout provides the real xfrm for that placeholder
 * After parsing the footer text block should carry the layout's coordinates.
 */
async function buildPptxWithInheritedFooter() {
  const zip = new JSZip();

  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
  zip.file('_rels/.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file('ppt/presentation.xml',
    `<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="${SLIDE_W_EMU}" cy="${SLIDE_H_EMU}"/>
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`);

  zip.file('ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"        Target="slides/slide1.xml"/>
  <Relationship Id="rIdM"  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>`);

  // Slide 1: title with explicit xfrm + footer placeholder with NO xfrm
  zip.file('ppt/slides/slide1.xml',
    `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <!-- Title — has explicit xfrm, should use its own coordinates -->
    <p:sp>
      <p:nvSpPr>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="title"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="457200" y="274638"/>
          <a:ext cx="8229600" cy="1143000"/>
        </a:xfrm>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>My Slide Title</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <!-- Footer — NO <a:xfrm>: must be resolved from layout -->
    <p:sp>
      <p:nvSpPr>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>footer text</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`);

  // Slide rels → layout
  zip.file('ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout3.xml"/>
</Relationships>`);

  // Layout: provides footer xfrm
  zip.file('ppt/slideLayouts/slideLayout3.xml', makeLayoutXml(true));

  // Layout rels → master
  zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  // Master (minimal — provides no extra placeholder geometry needed for this test)
  zip.file('ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Office Theme"><p:spTree/></p:cSld>
</p:sldMaster>`);

  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout3.xml"/>
</Relationships>`);

  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('parsePptx — FR-11 placeholder geometry inheritance', () => {
  let ir;

  beforeAll(async () => {
    const buf = await buildPptxWithInheritedFooter();
    ({ ir } = await parsePptx(buf, { filename: 'layout-test.pptx' }));
  });

  test('slide is parsed successfully', () => {
    expect(ir.slideset.slides).toHaveLength(1);
  });

  test('title has its own explicit xfrm coordinates (unaffected by fix)', () => {
    const texts = ir.slideset.slides[0].contents.text;
    const title = texts.find((b) => b.paragraphs[0].runs[0].text === 'My Slide Title');
    expect(title).toBeDefined();
    expect(title.position.x).toBe(toP(457200));
    expect(title.position.y).toBe(toP(274638));
  });

  test('footer text block gets position from layout (not CSS fallback)', () => {
    const texts = ir.slideset.slides[0].contents.text;
    const footer = texts.find((b) => b.paragraphs[0].runs[0].text === 'footer text');
    expect(footer).toBeDefined();
    // Real position resolved from layout
    expect(footer.position).toBeDefined();
    expect(footer.position.x).toBe(toP(FTR_X_EMU));
    expect(footer.position.y).toBe(toP(FTR_Y_EMU));
    expect(footer.width).toBe(toP(FTR_CX_EMU));
    expect(footer.height).toBe(toP(FTR_CY_EMU));
  });

  test('footer text block has NO footer-placement flag when real position is set', () => {
    const texts = ir.slideset.slides[0].contents.text;
    const footer = texts.find((b) => b.paragraphs[0].runs[0].text === 'footer text');
    expect(footer['footer-placement']).toBeUndefined();
  });

  test('IR still passes schema validation after geometry fix', () => {
    const { validate } = require('../src/ir/validator');
    const { valid } = validate(ir);
    expect(valid).toBe(true);
  });

  test('footer text block gets text-anchor="t" regardless of layout anchor (top-anchor override)', () => {
    // Footer placeholders use anchor="ctr" in many templates (inc. the test layout
    // which has no explicit anchor = defaults to no CSS flex). After the FR-11 fix
    // the parser overrides footer text-anchor to "t" so the generator top-anchors
    // the text.  This keeps the readable cap/ascender region visible inside the
    // small footer box instead of centering and clipping the letter middles.
    const texts = ir.slideset.slides[0].contents.text;
    const footer = texts.find((b) => b.paragraphs[0].runs[0].text === 'footer text');
    expect(footer['text-anchor']).toBe('t');
  });

  test('footer position.y uses the same EMU→px scale as slideHeight (scale-once)', () => {
    // Both element positions and slide dimensions are derived by dividing EMU
    // values by the single EMU_PER_PX constant (9525).  This test asserts the
    // scale is applied consistently: footer top (y) must be a number computed
    // with the same divisor as slideHeightPx so they can be compared directly.
    const { EMU_PER_PX } = require('../src/parser/pptx/units');

    const slideH = ir.slideset.master.slideHeight;
    expect(slideH).toBe(Math.round(SLIDE_H_EMU / EMU_PER_PX)); // 540

    const texts  = ir.slideset.slides[0].contents.text;
    const footer = texts.find((b) => b.paragraphs[0].runs[0].text === 'footer text');
    expect(footer.position.y).toBe(Math.round(FTR_Y_EMU / EMU_PER_PX)); // 519

    // 519 < 540 — footer top is within the slide bounds; overflow:visible lets
    // its bottom edge (519 + 48 = 567) render without being clipped.
    expect(footer.position.y).toBeLessThan(slideH);
  });
});

// ---------------------------------------------------------------------------
// collectPlaceholders — defaultFontSize from placeholder lstStyle
// ---------------------------------------------------------------------------

function ftrWithLstStyleXml(sz) {
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:nvSpPr>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="${FTR_X_EMU}" y="${FTR_Y_EMU}"/>
        <a:ext cx="${FTR_CX_EMU}" cy="${FTR_CY_EMU}"/>
      </a:xfrm>
    </p:spPr>
    <p:txBody>
      <a:bodyPr/>
      <a:lstStyle><a:lvl1pPr><a:defRPr sz="${sz}"/></a:lvl1pPr></a:lstStyle>
      <a:p><a:endParaRPr/></a:p>
    </p:txBody>
  </p:sp>`;
}

describe('collectPlaceholders — defaultFontSize from lstStyle', () => {
  test('sz=900 → defaultFontSize "9pt"', () => {
    const spTree = makeSpTree([ftrWithLstStyleXml(900)]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultFontSize).toBe('9pt');
  });

  test('sz=1800 → defaultFontSize "18pt"', () => {
    const spTree = makeSpTree([ftrWithLstStyleXml(1800)]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultFontSize).toBe('18pt');
  });

  test('absent lstStyle defRPr sz leaves defaultFontSize undefined', () => {
    const spTree = makeSpTree([phSpXml({ type: 'ftr', idx: 11, withXfrm: true })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultFontSize).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// collectPlaceholders — defaultColor from placeholder lstStyle solidFill
// ---------------------------------------------------------------------------

function ftrWithColorXml({ srgb, scheme } = {}) {
  const fillXml = srgb
    ? `<a:solidFill><a:srgbClr val="${srgb}"/></a:solidFill>`
    : scheme
      ? `<a:solidFill><a:schemeClr val="${scheme}"/></a:solidFill>`
      : '';
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:nvSpPr>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="${FTR_X_EMU}" y="${FTR_Y_EMU}"/>
        <a:ext cx="${FTR_CX_EMU}" cy="${FTR_CY_EMU}"/>
      </a:xfrm>
    </p:spPr>
    <p:txBody>
      <a:bodyPr/>
      <a:lstStyle><a:lvl1pPr><a:defRPr>${fillXml}</a:defRPr></a:lvl1pPr></a:lstStyle>
      <a:p><a:endParaRPr/></a:p>
    </p:txBody>
  </p:sp>`;
}

describe('collectPlaceholders — defaultColor from lstStyle solidFill', () => {
  test('srgbClr stores #rrggbb color in defaultColor', () => {
    const spTree = makeSpTree([ftrWithColorXml({ srgb: 'FF0000' })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultColor).toBe('#FF0000');
  });

  test('schemeClr tx1 is normalized to var(--theme-dk1)', () => {
    const spTree = makeSpTree([ftrWithColorXml({ scheme: 'tx1' })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultColor).toBe('var(--theme-dk1)');
  });

  test('schemeClr accent1 passes through as var(--theme-accent1)', () => {
    const spTree = makeSpTree([ftrWithColorXml({ scheme: 'accent1' })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultColor).toBe('var(--theme-accent1)');
  });

  test('absent solidFill leaves defaultColor undefined', () => {
    const spTree = makeSpTree([phSpXml({ type: 'ftr', idx: 11, withXfrm: true })]);
    const { byIdx } = collectPlaceholders(spTree);
    expect(byIdx.get(11).defaultColor).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// collectLayoutMedia — layout + master <p:pic> collection
// ---------------------------------------------------------------------------

const PIC_X_EMU  = 9000000;   // ~945px (top-right corner)
const PIC_Y_EMU  =  342900;   //  ~36px
const PIC_CX_EMU = 1371600;   // ~144px
const PIC_CY_EMU =  685800;   //  ~72px

function picXml(rId) {
  return `<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:nvPicPr>
      <p:cNvPr id="10" name="Logo"/>
      <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
      <p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="${rId}"/>
      <a:stretch><a:fillRect/></a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm>
        <a:off x="${PIC_X_EMU}" y="${PIC_Y_EMU}"/>
        <a:ext cx="${PIC_CX_EMU}" cy="${PIC_CY_EMU}"/>
      </a:xfrm>
      <a:prstGeom prst="rect"/>
    </p:spPr>
  </p:pic>`;
}

async function buildZipWithPics() {
  const zip = new JSZip();

  zip.file('ppt/slideLayouts/slideLayout3.xml',
    `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Test"><p:spTree>${picXml('rIdL1')}</p:spTree></p:cSld>
</p:sldLayout>`);

  zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-logo.png"/>
  <Relationship Id="rId1"  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  zip.file('ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Office Theme"><p:spTree>${picXml('rIdM1')}</p:spTree></p:cSld>
</p:sldMaster>`);

  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdM1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/master-logo.png"/>
</Relationships>`);

  return zip;
}

// Layout has no pics; master has one pic — used to test that master pics are
// preserved when they don't overlap with any layout pic.
async function buildZipWithMasterOnlyPic() {
  const zip = new JSZip();

  zip.file('ppt/slideLayouts/slideLayout3.xml',
    `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Test"><p:spTree/></p:cSld>
</p:sldLayout>`);

  zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  zip.file('ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Office Theme"><p:spTree>${picXml('rIdM1')}</p:spTree></p:cSld>
</p:sldMaster>`);

  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdM1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/master-logo.png"/>
</Relationships>`);

  return zip;
}

describe('collectLayoutMedia', () => {
  test('returns empty arrays when layoutPath is null', async () => {
    const { layoutMedia, masterMedia } = await collectLayoutMedia(new JSZip(), null);
    expect(layoutMedia).toEqual([]);
    expect(masterMedia).toEqual([]);
  });

  test('layout pic has correct position (EMU→px)', async () => {
    const zip = await buildZipWithPics();
    const { layoutMedia } = await collectLayoutMedia(zip, 'ppt/slideLayouts/slideLayout3.xml');
    expect(layoutMedia).toHaveLength(1);
    expect(layoutMedia[0].position.x).toBe(toP(PIC_X_EMU));
    expect(layoutMedia[0].position.y).toBe(toP(PIC_Y_EMU));
    expect(layoutMedia[0].width).toBe(toP(PIC_CX_EMU));
    expect(layoutMedia[0].height).toBe(toP(PIC_CY_EMU));
  });

  test('layout pic file-link resolves to the correct zip path', async () => {
    const zip = await buildZipWithPics();
    const { layoutMedia } = await collectLayoutMedia(zip, 'ppt/slideLayouts/slideLayout3.xml');
    expect(layoutMedia[0]['file-link']).toBe('ppt/media/layout-logo.png');
  });

  test('master pic at same position as layout pic is deduplicated (prevents doubled logos)', async () => {
    const zip = await buildZipWithPics();
    const { masterMedia } = await collectLayoutMedia(zip, 'ppt/slideLayouts/slideLayout3.xml');
    // Layout has a pic at the same position — master pic is filtered to avoid doubling it.
    expect(masterMedia).toHaveLength(0);
  });

  test('master-only pic (no layout overlap) has correct position', async () => {
    const zip = await buildZipWithMasterOnlyPic();
    const { masterMedia } = await collectLayoutMedia(zip, 'ppt/slideLayouts/slideLayout3.xml');
    expect(masterMedia).toHaveLength(1);
    expect(masterMedia[0].position.x).toBe(toP(PIC_X_EMU));
  });

  test('master-only pic (no layout overlap) file-link resolves to the correct zip path', async () => {
    const zip = await buildZipWithMasterOnlyPic();
    const { masterMedia } = await collectLayoutMedia(zip, 'ppt/slideLayouts/slideLayout3.xml');
    expect(masterMedia[0]['file-link']).toBe('ppt/media/master-logo.png');
  });

  test('returns empty arrays when layout has no pics and master has no pics', async () => {
    const zip = await buildZipWithLayout({ layoutHasFtr: true });
    const { layoutMedia, masterMedia } = await collectLayoutMedia(zip,
      'ppt/slideLayouts/slideLayout3.xml');
    expect(layoutMedia).toEqual([]);
    expect(masterMedia).toEqual([]);
  });
});
