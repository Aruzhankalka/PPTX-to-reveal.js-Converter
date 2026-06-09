const JSZip = require('jszip');
const { parseLayoutXml, parseMaster } = require('../src/parser/pptx/master');
const { parsePptx } = require('../src/parser/pptx');

// ---------------------------------------------------------------------------
// parseLayoutXml — pure unit tests (no ZIP)
// ---------------------------------------------------------------------------
describe('parseLayoutXml', () => {
  const LAYOUT_XML = `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="obj">
  <p:cSld name="Title and Content"/>
</p:sldLayout>`;

  test('extracts name and type', () => {
    const layout = parseLayoutXml(LAYOUT_XML, 'ppt/slideLayouts/slideLayout2.xml');
    expect(layout.name).toBe('Title and Content');
    expect(layout.type).toBe('obj');
  });

  test('uses layoutPath as id', () => {
    const layout = parseLayoutXml(LAYOUT_XML, 'ppt/slideLayouts/slideLayout2.xml');
    expect(layout.id).toBe('ppt/slideLayouts/slideLayout2.xml');
  });

  test('returns null name when cSld has no name attribute', () => {
    const xml = `<p:sldLayout xmlns:p="x" type="blank"><p:cSld/></p:sldLayout>`;
    expect(parseLayoutXml(xml, 'x.xml').name).toBeNull();
  });

  test('returns null type when type attribute is absent', () => {
    const xml = `<p:sldLayout xmlns:p="x"><p:cSld name="Blank"/></p:sldLayout>`;
    expect(parseLayoutXml(xml, 'x.xml').type).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Helper: build a minimal PPTX in memory that includes a master + layouts
// ---------------------------------------------------------------------------
async function buildPptxWithMaster() {
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
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rIdM1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`);

  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"        Target="slides/slide1.xml"/>
  <Relationship Id="rIdM1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>`);

  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Office Theme"/>
</p:sldMaster>`);

  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"       Target="../theme/theme1.xml"/>
</Relationships>`);

  zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld name="Title Slide"/>
</p:sldLayout>`);

  zip.file('ppt/slideLayouts/slideLayout2.xml', `<?xml version="1.0"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj">
  <p:cSld name="Title and Content"/>
</p:sldLayout>`);

  zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window"     lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`);

  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`);

  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>`);

  return zip.generateAsync({ type: 'nodebuffer' });
}

// ---------------------------------------------------------------------------
// parseMaster — integration tests against the in-memory ZIP
// ---------------------------------------------------------------------------
describe('parseMaster', () => {
  let zip;
  let result;

  beforeAll(async () => {
    const buf = await buildPptxWithMaster();
    const JSZipLib = require('jszip');
    zip = await JSZipLib.loadAsync(buf);
    result = await parseMaster(zip);
  });

  test('returns a non-null result when a master is present', () => {
    expect(result).not.toBeNull();
  });

  test('extracts master name', () => {
    expect(result.masterName).toBe('Office Theme');
  });

  test('extracts theme via the master rels (not the fallback path)', () => {
    expect(result.theme).not.toBeNull();
    expect(result.theme.colors.accent1).toBe('#4472C4');
    expect(result.theme.fonts.major).toBe('Calibri Light');
  });

  test('enumerates layouts in rels order', () => {
    expect(result.layouts).toHaveLength(2);
    expect(result.layouts[0]).toMatchObject({ type: 'title', name: 'Title Slide' });
    expect(result.layouts[1]).toMatchObject({ type: 'obj',   name: 'Title and Content' });
  });

  test('layout ids are ZIP paths', () => {
    expect(result.layouts[0].id).toBe('ppt/slideLayouts/slideLayout1.xml');
    expect(result.layouts[1].id).toBe('ppt/slideLayouts/slideLayout2.xml');
  });

  test('returns null when presentation has no master relationship', async () => {
    const emptyZip = new JSZip();
    emptyZip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
    expect(await parseMaster(emptyZip)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// parseTxStyles — per-level entries include line-spacing and paragraph spacing
// ---------------------------------------------------------------------------
describe('parseTxStyles — per-level entries include spacing', () => {
  let txStyles;

  const MASTER_WITH_STYLES = `<?xml version="1.0"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Test Master"/>
  <p:txStyles>
    <p:bodyStyle>
      <a:lvl1pPr>
        <a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
        <a:spcBef><a:spcPts val="1000"/></a:spcBef>
        <a:spcAft><a:spcPts val="0"/></a:spcAft>
        <a:defRPr sz="2200"/>
      </a:lvl1pPr>
      <a:lvl2pPr>
        <a:lnSpc><a:spcPts val="1440"/></a:lnSpc>
        <a:spcBef><a:spcPct val="50000"/></a:spcBef>
        <a:defRPr sz="1800"/>
      </a:lvl2pPr>
    </p:bodyStyle>
  </p:txStyles>
</p:sldMaster>`;

  beforeAll(async () => {
    const JSZipLib = require('jszip');
    const zip = new JSZipLib();
    zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdM1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>`);
    zip.file('ppt/slideMasters/slideMaster1.xml', MASTER_WITH_STYLES);
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
    const result = await parseMaster(zip);
    txStyles = result.txStyles;
  });

  test('body level 1 has font size from defRPr sz (2200 → 22pt)', () => {
    expect(txStyles.body[1].size).toBe('22pt');
  });

  test('body level 1 has lineSpacing from spcPct (90000 → "0.9")', () => {
    expect(txStyles.body[1].lineSpacing).toBe('0.9');
  });

  test('body level 1 has spaceBefore from spcPts (1000 → "10pt")', () => {
    expect(txStyles.body[1].spaceBefore).toBe('10pt');
  });

  test('body level 1 has spaceAfter from spcPts (0 → "0pt")', () => {
    expect(txStyles.body[1].spaceAfter).toBe('0pt');
  });

  test('body level 2 has lineSpacing from spcPts (1440 → "14.4pt")', () => {
    expect(txStyles.body[2].lineSpacing).toBe('14.4pt');
  });

  test('body level 2 has spaceBefore from spcPct (50000 → "0.5em")', () => {
    expect(txStyles.body[2].spaceBefore).toBe('0.5em');
  });

  test('absent txStyles section returns empty entries object', () => {
    expect(txStyles.title).toEqual({});
  });
});

// ---------------------------------------------------------------------------
// End-to-end: parsePptx wires layout-id onto slides and populates slideset
// ---------------------------------------------------------------------------
describe('parsePptx master + layout wiring (FR-11, FR-12)', () => {
  let ir;

  beforeAll(async () => {
    const buf = await buildPptxWithMaster();
    ({ ir } = await parsePptx(buf, { filename: 'test.pptx' }));
  });

  test('slideset.master.theme contains colour slots', () => {
    expect(ir.slideset.master.theme.colors.dk1).toBe('#000000');
    expect(ir.slideset.master.theme.colors.accent1).toBe('#4472C4');
  });

  test('slideset.master.theme contains font names', () => {
    expect(ir.slideset.master.theme.fonts.minor).toBe('Calibri');
  });

  test('slideset.layouts lists all layouts', () => {
    expect(ir.slideset.layouts).toHaveLength(2);
  });

  test('slide carries layout-id pointing to the correct layout', () => {
    expect(ir.slideset.slides[0]['layout-id']).toBe('ppt/slideLayouts/slideLayout2.xml');
  });

  test('IR passes schema validation', () => {
    const { validate } = require('../src/ir/validator');
    const { valid, errors } = validate(ir);
    expect(valid).toBe(true);
  });
});
