'use strict';

/**
 * Flat-CSS-string color resolver — used for formatting fields that want a
 * plain CSS color value (run/paragraph text color, layout placeholder
 * background), as opposed to shapes.js's structured shapeColor object
 * ({space, hex|ref}) used by shape fill/stroke.
 */

// OOXML scheme-color aliases not present in the theme colors map.
// tx1/tx2 are display-text aliases for dk1/dk2; bg1/bg2 are background aliases for lt1/lt2.
const SCHEME_ALIAS = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };

/**
 * Resolve an <a:solidFill> (or any node with the same srgbClr/schemeClr
 * children, e.g. <a:defRPr>'s fill) to a flat CSS color string.
 *
 * Used for the "flat CSS string" formatting fields — paragraph/run/master
 * formatting.color, layout placeholder background — as opposed to the
 * structured shapeColor object ({space,hex|ref}) used by shapes[].fill/stroke,
 * which is resolved separately by shapes.js's resolveColorNode.
 *
 * @param {object|undefined} fill - parsed <a:solidFill> node
 * @returns {string|null} '#RRGGBB', 'var(--theme-X)', or null when unresolvable
 */
function resolveSolidFillCss(fill) {
  if (!fill) return null;
  if (fill['a:srgbClr'] && fill['a:srgbClr']['@_val']) {
    return '#' + fill['a:srgbClr']['@_val'];
  }
  if (fill['a:schemeClr'] && fill['a:schemeClr']['@_val']) {
    const raw = fill['a:schemeClr']['@_val'];
    return `var(--theme-${SCHEME_ALIAS[raw] || raw})`;
  }
  return null;
}

module.exports = { SCHEME_ALIAS, resolveSolidFillCss };
