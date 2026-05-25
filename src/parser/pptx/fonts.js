const { readText, readBinary } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');

const TYPE_FONTTABLE = '/fonttable';
const TYPE_FONT      = '/font';

// Maps <p:regular> / <p:bold> / <p:italic> / <p:boldItalic> to CSS weight+style.
const VARIANT_MAP = [
  { el: 'p:regular',    weight: 400, style: 'normal' },
  { el: 'p:bold',       weight: 700, style: 'normal' },
  { el: 'p:italic',     weight: 400, style: 'italic' },
  { el: 'p:boldItalic', weight: 700, style: 'italic' },
];

/**
 * Build a stable fontRef.id from family, weight, style.
 * Schema pattern: ^[a-z0-9-]+$
 * Example: "Times New Roman", 700, "italic" → "times-new-roman-700-italic"
 */
function makeFontId(family, weight, style) {
  return [family, String(weight), style]
    .join('-')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

/**
 * Guess a CSS font-family fallback stack from the requested family name.
 */
function guessFallback(family) {
  const lower = family.toLowerCase();
  if (/mono|courier|code|console|typewriter/.test(lower)) {
    return '"Courier New", Courier, monospace';
  }
  if (/serif|roman|georgia|garamond|palatino|times|book antiqua|bookman/.test(lower)) {
    return '"Times New Roman", Times, serif';
  }
  return 'Arial, Helvetica, sans-serif';
}

/**
 * Extract embedded fonts from the PPTX and build the IR font registry (FR-08).
 *
 * Flow:
 *   ppt/_rels/presentation.xml.rels
 *     → fontTable rel → ppt/fontTable.xml
 *   ppt/_rels/fontTable.xml.rels
 *     → rId → font binary path inside zip (ppt/fonts/*.fntdata or *.odttf)
 *
 * Each <p:font> in the font table is either:
 *   - Non-embedded: no variant child elements → source='missing'
 *   - Embedded:     has <p:regular r:id="rId1"> etc. → source='embedded'
 *
 * PPTX embeds fonts as XOR-obfuscated OTF/TTF (.odttf / .fntdata). The bytes
 * are stored as-is with a warning; full deobfuscation is a Sprint 3 concern.
 *
 * @param {JSZip} zip
 * @returns {Promise<{
 *   fonts: object[],                                    // fontRef[] for ir.slideset.fonts
 *   fontBytes: Array<{ bundlePath: string, bytes: Buffer }>
 * }>}
 */
async function parseFonts(zip) {
  // 1. Locate fontTable.xml via presentation.xml.rels (fall back to convention)
  const presRelsXml = await readText(zip, 'ppt/_rels/presentation.xml.rels');
  const presRels = parseRelationships(presRelsXml);

  const fontTableRel = Object.values(presRels).find((r) =>
    r.type.toLowerCase().endsWith(TYPE_FONTTABLE)
  );
  const fontTablePath = fontTableRel
    ? resolveTarget('ppt', fontTableRel.target)
    : 'ppt/fontTable.xml';

  // 2. Parse fontTable.xml
  const fontTableXml = await readText(zip, fontTablePath);
  if (!fontTableXml) return { fonts: [], fontBytes: [] };

  const parsed = parseXml(fontTableXml);
  const fontTbl = parsed && parsed['p:fontTbl'];
  if (!fontTbl) return { fonts: [], fontBytes: [] };

  // 3. Load fontTable.xml.rels to resolve rId → zip path for embedded fonts
  const ftDir = fontTablePath.includes('/')
    ? fontTablePath.substring(0, fontTablePath.lastIndexOf('/'))
    : '';
  const ftName = fontTablePath.substring(fontTablePath.lastIndexOf('/') + 1);
  const ftRelsPath = `${ftDir}/_rels/${ftName}.rels`;
  const ftRelsXml = await readText(zip, ftRelsPath);
  const ftRels = parseRelationships(ftRelsXml);

  // Build rId → zip path for every font relationship
  const fontFileByRId = {};
  for (const [rId, rel] of Object.entries(ftRels)) {
    if (rel.type.toLowerCase().endsWith(TYPE_FONT)) {
      fontFileByRId[rId] = resolveTarget(ftDir, rel.target);
    }
  }

  // 4. Build fontRef entries, one per (family, weight, style) combination
  const fonts = [];
  const fontBytes = [];
  const seenIds = new Set();

  for (const fontEntry of asArray(fontTbl['p:font'])) {
    const family = fontEntry['@_typeface'];
    // Skip missing, empty, or theme-slot references (+mj-lt, +mn-lt, etc.)
    if (!family || family.startsWith('+')) continue;

    let embeddedAny = false;

    for (const { el, weight, style } of VARIANT_MAP) {
      const variantNode = fontEntry[el];
      if (!variantNode) continue;

      const rId = variantNode['@_r:id'];
      const zipPath = rId && fontFileByRId[rId];
      const id = makeFontId(family, weight, style);
      if (seenIds.has(id)) continue;

      if (zipPath) {
        const bytes = await readBinary(zip, zipPath);
        if (bytes) {
          seenIds.add(id);
          const bundlePath = `fonts/${id}.ttf`;
          fontBytes.push({ bundlePath, bytes });
          fonts.push({
            id,
            family,
            weight,
            style,
            source: 'embedded',
            file: bundlePath,
            format: 'ttf',
            fallback: guessFallback(family),
            subset: true, // PPTX typically subsets embedded fonts to used glyphs
            metricsCompatible: true,
            license: { fsType: 0, embeddable: true },
            warnings: [
              `Font '${family}' extracted from PPTX embed; ` +
              'bytes may be XOR-obfuscated (.odttf). Verify rendering in browser.',
            ],
          });
          embeddedAny = true;
        }
      }
    }

    if (!embeddedAny) {
      // Font referenced but not embedded: one missing entry for weight=400 normal
      const id = makeFontId(family, 400, 'normal');
      if (!seenIds.has(id)) {
        seenIds.add(id);
        fonts.push({
          id,
          family,
          weight: 400,
          style: 'normal',
          source: 'missing',
          file: null,
          fallback: guessFallback(family),
          subset: false,
          metricsCompatible: false,
          license: { fsType: 0, embeddable: false },
          warnings: [
            `Font '${family}' is not embedded in the PPTX; browser fallback will be used.`,
          ],
        });
      }
    }
  }

  return { fonts, fontBytes };
}

module.exports = { parseFonts, makeFontId, guessFallback };
