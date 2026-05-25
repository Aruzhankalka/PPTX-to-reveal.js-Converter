const { parseThemeXml } = require('../src/parser/pptx/theme');

// Minimal theme XML matching the Office default colour scheme
const OFFICE_THEME_XML = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window"     lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

describe('parseThemeXml', () => {
  test('returns null for null input', () => {
    expect(parseThemeXml(null)).toBeNull();
  });

  test('returns null when root element is not a:theme', () => {
    expect(parseThemeXml('<foo/>')).toBeNull();
  });

  test('extracts theme name', () => {
    const theme = parseThemeXml(OFFICE_THEME_XML);
    expect(theme.name).toBe('Office Theme');
  });

  test('resolves srgbClr slots to #RRGGBB', () => {
    const { colors } = parseThemeXml(OFFICE_THEME_XML);
    expect(colors.accent1).toBe('#4472C4');
    expect(colors.dk2).toBe('#44546A');
    expect(colors.hlink).toBe('#0563C1');
    expect(colors.folHlink).toBe('#954F72');
  });

  test('resolves sysClr slots using lastClr', () => {
    const { colors } = parseThemeXml(OFFICE_THEME_XML);
    expect(colors.dk1).toBe('#000000');
    expect(colors.lt1).toBe('#FFFFFF');
  });

  test('extracts all 12 colour slots', () => {
    const { colors } = parseThemeXml(OFFICE_THEME_XML);
    const slots = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
    for (const slot of slots) {
      expect(colors).toHaveProperty(slot);
    }
  });

  test('extracts major (heading) and minor (body) font', () => {
    const { fonts } = parseThemeXml(OFFICE_THEME_XML);
    expect(fonts.major).toBe('Calibri Light');
    expect(fonts.minor).toBe('Calibri');
  });

  test('omits fonts when fontScheme is absent', () => {
    const xml = OFFICE_THEME_XML.replace(/<a:fontScheme[\s\S]*?<\/a:fontScheme>/, '');
    const { fonts } = parseThemeXml(xml);
    expect(fonts.major).toBeUndefined();
    expect(fonts.minor).toBeUndefined();
  });

  test('omits colour slot when sysClr has no lastClr', () => {
    const xml = OFFICE_THEME_XML.replace(
      '<a:sysClr val="windowText" lastClr="000000"/>',
      '<a:sysClr val="windowText"/>',
    );
    const { colors } = parseThemeXml(xml);
    expect(colors.dk1).toBeUndefined();
  });

  test('theme missing themeElements returns null', () => {
    expect(parseThemeXml('<a:theme xmlns:a="x" name="t"/>')).toBeNull();
  });
});
