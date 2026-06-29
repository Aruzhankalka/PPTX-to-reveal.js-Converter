const JSZip = require('jszip');
const { parseFonts, makeFontId, guessFallback } = require('../src/parser/pptx/fonts');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const PPTX_NS  = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS_NS  = 'http://schemas.openxmlformats.org/package/2006/relationships';
const FONT_REL = `${R_NS}/font`;
const FT_REL   = `${R_NS}/fontTable`;

function makePresRels(withFontTable = true) {
  const rel = withFontTable
    ? `<Relationship Id="rId1" Type="${FT_REL}" Target="fontTable.xml"/>`
    : '';
  return `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">${rel}</Relationships>`;
}

function makeFontTableXml(entries = '') {
  return `<?xml version="1.0"?>
<p:fontTbl xmlns:p="${PPTX_NS}" xmlns:r="${R_NS}">
  ${entries}
</p:fontTbl>`;
}

function makeFontTableRels(rels = '') {
  return `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">${rels}</Relationships>`;
}

/**
 * Build a minimal in-memory zip for font extraction tests.
 *
 * @param {object} options
 * @param {string}  options.fontTableEntries  - raw XML string of <p:font> entries
 * @param {string}  options.fontTableRels     - raw XML string of <Relationship> entries
 * @param {Map}     options.fontFiles         - Map<zipPath, Buffer> for embedded font bytes
 * @param {boolean} options.noFontTable       - omit fontTable.xml entirely
 * @param {boolean} options.noPresRel         - omit the fontTable rel from presentation.xml.rels
 */
async function buildZip({
  fontTableEntries = '',
  fontTableRels    = '',
  fontFiles        = new Map(),
  noFontTable      = false,
  noPresRel        = false,
} = {}) {
  const zip = new JSZip();

  zip.file('ppt/_rels/presentation.xml.rels', makePresRels(!noPresRel));

  if (!noFontTable) {
    zip.file('ppt/fontTable.xml', makeFontTableXml(fontTableEntries));
    if (fontTableRels) {
      zip.file('ppt/_rels/fontTable.xml.rels', makeFontTableRels(fontTableRels));
    }
  }

  for (const [path, bytes] of fontFiles) {
    zip.file(path, bytes);
  }

  return zip;
}

// ---------------------------------------------------------------------------
// Unit tests — makeFontId
// ---------------------------------------------------------------------------

describe('makeFontId', () => {
  test('simple family produces lowercase kebab id', () => {
    expect(makeFontId('Arial', 400, 'normal')).toBe('arial-400-normal');
  });

  test('spaces in family name become hyphens', () => {
    expect(makeFontId('Times New Roman', 700, 'italic')).toBe('times-new-roman-700-italic');
  });

  test('dots and special chars are replaced by hyphens', () => {
    expect(makeFontId('Open.Sans', 400, 'normal')).toBe('open-sans-400-normal');
  });

  test('result matches schema pattern ^[a-z0-9-]+$', () => {
    const id = makeFontId('Comic Sans MS', 700, 'normal');
    expect(id).toMatch(/^[a-z0-9-]+$/);
  });
});

// ---------------------------------------------------------------------------
// Unit tests — guessFallback
// ---------------------------------------------------------------------------

describe('guessFallback', () => {
  test('monospace families get monospace stack', () => {
    expect(guessFallback('Courier New')).toContain('monospace');
    expect(guessFallback('Source Code Pro')).toContain('monospace');
  });

  test('serif families get serif stack', () => {
    expect(guessFallback('Times New Roman')).toContain('serif');
    expect(guessFallback('Georgia')).toContain('serif');
  });

  test('unknown families get sans-serif stack', () => {
    expect(guessFallback('Calibri')).toContain('sans-serif');
    expect(guessFallback('Arial')).toContain('sans-serif');
  });

  test('fallback string is always non-empty', () => {
    expect(guessFallback('SomeRandomFont').length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// Integration tests — parseFonts
// ---------------------------------------------------------------------------

describe('parseFonts', () => {
  test('returns empty arrays when no fontTable.xml exists', async () => {
    const zip = await buildZip({ noFontTable: true });
    const { fonts, fontBytes } = await parseFonts(zip);
    expect(fonts).toEqual([]);
    expect(fontBytes).toEqual([]);
  });

  test('returns empty arrays when fontTable has no <p:font> entries', async () => {
    const zip = await buildZip({ fontTableEntries: '' });
    const { fonts, fontBytes } = await parseFonts(zip);
    expect(fonts).toEqual([]);
    expect(fontBytes).toEqual([]);
  });

  test('non-embedded font produces a missing entry', async () => {
    const zip = await buildZip({
      fontTableEntries: '<p:font typeface="Arial"/>',
    });
    const { fonts, fontBytes } = await parseFonts(zip);

    expect(fonts).toHaveLength(1);
    expect(fontBytes).toHaveLength(0);

    const entry = fonts[0];
    expect(entry.family).toBe('Arial');
    expect(entry.source).toBe('missing');
    expect(entry['font-file']).toBeNull();
    expect(entry.metricsCompatible).toBe(false);
    expect(entry.fallback.length).toBeGreaterThan(0);
    expect(entry.warnings).toHaveLength(1);
  });

  test('missing entry satisfies required fontRef fields', async () => {
    const zip = await buildZip({
      fontTableEntries: '<p:font typeface="Calibri"/>',
    });
    const { fonts } = await parseFonts(zip);
    const entry = fonts[0];

    // All required fields present
    expect(typeof entry['font-id']).toBe('string');
    expect(typeof entry.family).toBe('string');
    expect(typeof entry.weight).toBe('number');
    expect(['normal', 'italic', 'oblique']).toContain(entry.style);
    expect(typeof entry.fallback).toBe('string');
    expect(typeof entry.subset).toBe('boolean');
    expect(typeof entry.metricsCompatible).toBe('boolean');
    expect(typeof entry.license.fsType).toBe('number');
    expect(typeof entry.license.embeddable).toBe('boolean');
    expect(Array.isArray(entry.warnings)).toBe(true);
  });

  test('embedded font produces an embedded entry with bytes', async () => {
    const dummyBytes = Buffer.from([0x00, 0x01, 0x00, 0x00]);
    const zip = await buildZip({
      fontTableEntries: `<p:font typeface="TestFont"><p:regular r:id="rId1"/></p:font>`,
      fontTableRels: `<Relationship Id="rId1" Type="${FONT_REL}" Target="fonts/testfont.fntdata"/>`,
      fontFiles: new Map([['ppt/fonts/testfont.fntdata', dummyBytes]]),
    });

    const { fonts, fontBytes } = await parseFonts(zip);

    expect(fonts).toHaveLength(1);
    expect(fontBytes).toHaveLength(1);

    const entry = fonts[0];
    expect(entry.family).toBe('TestFont');
    expect(entry.source).toBe('embedded');
    expect(entry.weight).toBe(400);
    expect(entry.style).toBe('normal');
    expect(entry['font-file']).toMatch(/^fonts\//);
    expect(entry.format).toBe('ttf');
    expect(entry.metricsCompatible).toBe(true);
    expect(entry.subset).toBe(true);

    expect(fontBytes[0].bytes).toEqual(dummyBytes);
    expect(fontBytes[0].bundlePath).toBe(entry['font-file']);
  });

  test('all four font variants (regular, bold, italic, boldItalic) are extracted', async () => {
    const fontFiles = new Map([
      ['ppt/fonts/f1.fntdata', Buffer.from([0x01])],
      ['ppt/fonts/f2.fntdata', Buffer.from([0x02])],
      ['ppt/fonts/f3.fntdata', Buffer.from([0x03])],
      ['ppt/fonts/f4.fntdata', Buffer.from([0x04])],
    ]);

    const zip = await buildZip({
      fontTableEntries: `
        <p:font typeface="MultiFont">
          <p:regular    r:id="rId1"/>
          <p:bold       r:id="rId2"/>
          <p:italic     r:id="rId3"/>
          <p:boldItalic r:id="rId4"/>
        </p:font>`,
      fontTableRels: `
        <Relationship Id="rId1" Type="${FONT_REL}" Target="fonts/f1.fntdata"/>
        <Relationship Id="rId2" Type="${FONT_REL}" Target="fonts/f2.fntdata"/>
        <Relationship Id="rId3" Type="${FONT_REL}" Target="fonts/f3.fntdata"/>
        <Relationship Id="rId4" Type="${FONT_REL}" Target="fonts/f4.fntdata"/>`,
      fontFiles,
    });

    const { fonts, fontBytes } = await parseFonts(zip);

    expect(fonts).toHaveLength(4);
    expect(fontBytes).toHaveLength(4);

    const variants = fonts.map((f) => `${f.weight}-${f.style}`).sort();
    expect(variants).toEqual(['400-italic', '400-normal', '700-italic', '700-normal']);
  });

  test('theme slot typefaces (+mj-lt, +mn-lt) are skipped', async () => {
    const zip = await buildZip({
      fontTableEntries: `
        <p:font typeface="+mj-lt"/>
        <p:font typeface="+mn-lt"/>
        <p:font typeface="Arial"/>`,
    });

    const { fonts } = await parseFonts(zip);
    // Only Arial should appear; theme slots are ignored
    expect(fonts).toHaveLength(1);
    expect(fonts[0].family).toBe('Arial');
  });

  test('duplicate family entries are deduplicated by id', async () => {
    const zip = await buildZip({
      fontTableEntries: `
        <p:font typeface="Arial"/>
        <p:font typeface="Arial"/>`,
    });

    const { fonts } = await parseFonts(zip);
    expect(fonts).toHaveLength(1);
  });

  test('works without a fontTable rel in presentation.xml.rels (fallback path)', async () => {
    const zip = new JSZip();
    // No rel pointing to fontTable — parseFonts falls back to ppt/fontTable.xml
    zip.file('ppt/_rels/presentation.xml.rels', makePresRels(false));
    zip.file('ppt/fontTable.xml', makeFontTableXml('<p:font typeface="Verdana"/>'));

    const { fonts } = await parseFonts(zip);
    expect(fonts).toHaveLength(1);
    expect(fonts[0].family).toBe('Verdana');
  });

  test('font id matches schema pattern ^[a-z0-9-]+$', async () => {
    const zip = await buildZip({
      fontTableEntries: '<p:font typeface="Times New Roman"/>',
    });
    const { fonts } = await parseFonts(zip);
    expect(fonts[0]['font-id']).toMatch(/^[a-z0-9-]+$/);
  });
});
