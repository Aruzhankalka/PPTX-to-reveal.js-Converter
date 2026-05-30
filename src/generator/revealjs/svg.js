const { escapeHtml } = require("./escape");

function emuToPx(value) {
  if (typeof value !== "number") return 0;
  return Math.round(value / 9525);
}

function getShapeStyle(shape) {
  const x = emuToPx(shape.position?.x || 0);
  const y = emuToPx(shape.position?.y || 0);
  const width = emuToPx(shape.width || 0);
  const height = emuToPx(shape.height || 0);
  const rotation = shape.rotation || 0;
  const zIndex = shape["z-index"] || shape.zIndex || 0;

  return [
    "position:absolute",
    `left:${x}px`,
    `top:${y}px`,
    `width:${width}px`,
    `height:${height}px`,
    `transform:rotate(${rotation}deg)`,
    `z-index:${zIndex}`
  ].join(";");
}

function renderShapeElement(shape) {
  const type = shape.type || "rectangle";
  const fill = shape.fill?.color || "transparent";
  const stroke = shape.stroke?.color || "transparent";
  const strokeWidth = shape.stroke?.width || 0;

  switch (type) {
    case "ellipse":
    case "ellipsis":
      return `<ellipse cx="50%" cy="50%" rx="50%" ry="50%" fill="${escapeHtml(fill)}" stroke="${escapeHtml(stroke)}" stroke-width="${strokeWidth}" />`;

    case "line":
      return `<line x1="0" y1="50%" x2="100%" y2="50%" stroke="${escapeHtml(stroke)}" stroke-width="${strokeWidth || 2}" />`;

    case "triangle":
      return `<polygon points="50,0 100,100 0,100" fill="${escapeHtml(fill)}" stroke="${escapeHtml(stroke)}" stroke-width="${strokeWidth}" />`;

    case "rectangle":
    default:
      return `<rect x="0" y="0" width="100%" height="100%" fill="${escapeHtml(fill)}" stroke="${escapeHtml(stroke)}" stroke-width="${strokeWidth}" />`;
  }
}

function renderShape(shape) {
  const style = getShapeStyle(shape);
  const content = renderShapeElement(shape);

  return `<svg class="shape" style="${style}" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
  ${content}
</svg>`;
}

module.exports = {
  renderShape,
  emuToPx
};