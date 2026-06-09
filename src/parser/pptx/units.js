/**
 * EMU (English Metric Units) is OOXML's native coordinate unit.
 * 914,400 EMU = 1 inch = 96 px at 96 DPI.
 */
const EMU_PER_PX = 9525;

function emuToPx(emu) {
  if (emu == null) return null;
  const n = Number(emu);
  if (Number.isNaN(n)) return null;
  return Math.round(n / EMU_PER_PX);
}

/**
 * PPTX stores rotations in 60,000ths of a degree.
 */
function pptxRotationToDegrees(rotAttr) {
  if (rotAttr == null) return 0;
  const n = Number(rotAttr);
  if (Number.isNaN(n)) return 0;
  return n / 60000;
}

module.exports = { emuToPx, pptxRotationToDegrees, EMU_PER_PX };