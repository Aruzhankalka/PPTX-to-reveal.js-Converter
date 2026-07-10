/**
 * EMU (English Metric Units) is OOXML's native coordinate unit.
 * 914,400 EMU = 1 inch = 96 px at 96 DPI.
 */
const EMU_PER_PX = 9525;

/**
 * Convert an EMU value to CSS px at 96 DPI (1 px = 9525 EMU), rounded to the
 * nearest integer px — geometry is not sub-pixel precise anywhere else in
 * the pipeline, so this rounding is final, not intermediate.
 *
 * @param {string|number|null} emu - raw EMU value (e.g. an <a:off x="..."> attribute)
 * @returns {number|null} whole CSS px, or null when emu is null/undefined/NaN
 */
function emuToPx(emu) {
  if (emu == null) return null;
  const n = Number(emu);
  if (Number.isNaN(n)) return null;
  return Math.round(n / EMU_PER_PX);
}

/**
 * PPTX stores rotations in 60,000ths of a degree.
 *
 * @param {string|number|null} rotAttr - raw <a:xfrm rot="..."> attribute value
 * @returns {number} degrees; 0 when rotAttr is absent or not a number
 */
function pptxRotationToDegrees(rotAttr) {
  if (rotAttr == null) return 0;
  const n = Number(rotAttr);
  if (Number.isNaN(n)) return 0;
  return n / 60000;
}

module.exports = { emuToPx, pptxRotationToDegrees, EMU_PER_PX };