'use strict';

const JSZip = require('jszip');

function escapeXml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function pxToEmu(px) {
  return Math.round((px || 0) * 9144);
}

function buildContentTypes(slideCount) {
  const slideOverrides = Array.from({ length: slideCount }, (_, i) =>
    `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
</Types>`;
}

function buildRootRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
}

function buildPresentationRels(slideCount) {
  const slideRels = Array.from({ length: slideCount }, (_, i) =>
    `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${slideRels}
</Relationships>`;
}

function buildSlideRels(images) {
  if (!images || images.length === 0) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
  }

  const imageRels = images.map((img, i) =>
    `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${img.filename}"/>`
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${imageRels}
</Relationships>`;
}

function buildPresentation(slideCount) {
  const sldIdLst = Array.from({ length: slideCount }, (_, i) =>
    `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldMasterIdLst/>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst>
    ${sldIdLst}
  </p:sldIdLst>
</p:presentation>`;
}

function buildImageShape(img, index) {
  const x = pxToEmu(img.x || 0);
  const y = pxToEmu(img.y || 0);
  const cx = pxToEmu(img.width || 300);
  const cy = pxToEmu(img.height || 200);

  return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${10 + index}" name="Image ${index + 1}"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId${index + 1}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;
}

function buildSlide(slide) {
  const warnings = [];
  const images = [];
  let spTree = '';

  if (slide.title) {
    const titleText = typeof slide.title === 'string'
      ? slide.title
      : (slide.title.paragraphs?.[0]?.runs?.[0]?.text || '');

    spTree += `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="1" name="Title"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:r><a:t>${escapeXml(titleText)}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;
  }

  if (slide.contents && Array.isArray(slide.contents)) {
    for (const element of slide.contents) {
      if (!element) continue;

      if (element.type === 'text' || element.paragraphs) {
        const paragraphs = element.paragraphs || [];
        const textXml = paragraphs.map(p => {
          const runs = (p.runs || []).map(r =>
            `<a:r><a:t>${escapeXml(r.text || '')}</a:t></a:r>`
          ).join('');
          return `<a:p>${runs || '<a:r><a:t></a:t></a:r>'}</a:p>`;
        }).join('');

        spTree += `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Content"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${pxToEmu(element.x || 0)}" y="${pxToEmu(element.y || 0)}"/>
      <a:ext cx="${pxToEmu(element.width || 800)}" cy="${pxToEmu(element.height || 100)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    ${textXml}
  </p:txBody>
</p:sp>`;
      }

      if (element.type === 'media' || element.src) {
        const filename = element.src
          ? element.src.split('/').pop()
          : `image${images.length + 1}.png`;
        images.push({
          filename,
          data: element.data,
          x: element.x,
          y: element.y,
          width: element.width,
          height: element.height
        });
        spTree += buildImageShape(element, images.length - 1);
      }
    }
  }

  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      ${spTree}
    </p:spTree>
  </p:cSld>
</p:sld>`;

  return { slideXml, images, warnings };
}

/**
 * Generate a .pptx buffer from an IR document.
 * Takes the intermediate representation and builds a valid OOXML ZIP archive.
 */
async function generatePptx(ir, media = {}) {
  const warnings = [];
  const zip = new JSZip();
  const slides = ir.slideset?.slides || ir.slides || [];

  if (slides.length === 0) {
    warnings.push('No slides found in IR document');
  }

  zip.file('[Content_Types].xml', buildContentTypes(slides.length));
  zip.file('_rels/.rels', buildRootRels());
  zip.file('ppt/presentation.xml', buildPresentation(slides.length));
  zip.file('ppt/_rels/presentation.xml.rels', buildPresentationRels(slides.length));

  for (let i = 0; i < slides.length; i++) {
    const { slideXml, images, warnings: slideWarnings } = buildSlide(slides[i]);
    warnings.push(...slideWarnings);

    zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, buildSlideRels(images));

    for (const img of images) {
      if (img.data) {
        zip.file(`ppt/media/${img.filename}`, img.data);
      } else if (media[img.filename]) {
        zip.file(`ppt/media/${img.filename}`, media[img.filename]);
      }
    }
  }

  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { buffer, warnings };
}

module.exports = { generatePptx };