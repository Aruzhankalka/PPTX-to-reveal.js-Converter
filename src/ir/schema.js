/**
 * Web-Slide Internal Data Format — Sprint 2 JSON Schema
 *
 * Sprint 1 features (unchanged):
 *   FR-01/02  File upload + format validation
 *   FR-03     Text extraction from slides
 *   FR-04     Slide order preservation
 *   FR-05     reveal.js HTML output
 *   FR-09     Image inclusion
 *   FR-15     In-browser viewing
 *
 * Sprint 2 additions (this file):
 *   FR-06  Full text formatting: bold, italic, underline, strikethrough,
 *          color, font family/size, line spacing, alignment, lists.
 *   FR-08  Font registry: document-wide FontRef entries with per-run RunFont
 *          pointers. Supports embedded (WOFF2) / substituted / missing fonts.
 *   FR-10  Shapes: rect, ellipse, line, arrow, polyline, polygon, callout,
 *          connector — each with fill, stroke, geometry, rotation, and an
 *          optional embedded text-frame (paragraphs array).
 *
 * Sprint 1 documents remain valid; all Sprint 2 additions are optional at the
 * schema level so the validator can serve both generations of IR documents.
 *
 * NOT covered (deferred):
 *   FR-11/12  Master/layouts + theme colors
 *   FR-13     Stacking order (z-index)
 *   FR-14     Animations
 *   Tables, groups, transitions, notes
 *
 * EXTENSIBILITY:
 *   `additionalProperties: true` on element-level objects lets future sprints
 *   add fields without breaking existing IR documents.
 *   `additionalProperties: false` is reserved for closed-set definitions
 *   (fontRef, runFont, position, license) where the contract is precise.
 *
 * AUTHORITY:
 *   Source of truth is /docs/web-slide-internal-data.example.json. Any
 *   mismatch must be raised in cross-group sync, not silently fixed here.
 */

const slideSchema = {
  type: 'object',
  required: ['contents'],
  properties: {
    // Sprint 1: plain string title. Sprint 2+ widens to paragraph object
    // when per-run formatting on titles is needed.
    title: {
      oneOf: [
        { type: 'string' },
        {
          type: 'object',
          properties: { content: { type: 'string' } },
          additionalProperties: true,
        },
      ],
    },
    'layout-id': {
      type: 'string',
      description: 'Reference to layouts[] entry. Optional until FR-11 (Sprint 3).',
    },
    hidden: { type: 'boolean' },
    contents: {
      type: 'object',
      required: [],
      properties: {
        text: {
          type: 'array',
          items: { $ref: '#/definitions/textBlock' },
        },
        media: {
          type: 'array',
          items: { $ref: '#/definitions/mediaItem' },
        },
        // FR-10 (Sprint 2): typed shape items
        shapes: {
          type: 'array',
          items: { $ref: '#/definitions/shape' },
        },
        tables: { type: 'array' },
        groups: { type: 'array' },
        animations: { type: 'array' },
        background: {},
        transition: { type: 'string' },
        notes: { type: 'string' },
      },
      additionalProperties: true,
    },
  },
  additionalProperties: true,
};

const sprint2Schema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  $id: 'https://hof-university.de/web-slide-internal-data/sprint2.schema.json',
  title: 'Web-Slide Internal Data Format (Sprint 2)',
  description:
    'IR schema for the PPTX -> reveal.js converter. ' +
    'Sprint 2 adds full text formatting (FR-06), font extraction (FR-08), ' +
    'and shapes (FR-10) to the Sprint 1 text + media base.',
  type: 'object',
  required: ['slideset'],
  additionalProperties: false,
  properties: {
    slideset: {
      type: 'object',
      required: ['slides'],
      additionalProperties: true,
      properties: {
        filename: { type: 'string' },
        title: { type: 'string' },
        author: { type: 'string' },
        'creation-date': {
          type: 'string',
          // Example file says "yyyy-mm-dd"; full ISO 8601 also accepted since
          // PPTX core.xml uses dateTime.
          pattern: '^\\d{4}-\\d{2}-\\d{2}(T.*)?$',
        },

        // FR-08 (Sprint 2): document-wide font registry
        fonts: {
          type: 'array',
          items: { $ref: '#/definitions/fontRef' },
          description:
            'One entry per (family, weight, style) combination, ' +
            'deduplicated across all text runs.',
        },

        // FR-11/12 (Sprint 3): master theme + color variables
        master: {
          type: 'object',
          additionalProperties: true,
        },

        // FR-11 (Sprint 3): slide layouts
        layouts: {
          type: 'array',
          items: { type: 'object', additionalProperties: true },
        },

        slides: {
          type: 'array',
          minItems: 0,
          items: slideSchema,
        },
      },
    },
  },

  // ---------------------------------------------------------------------------
  // Reusable definitions
  // ---------------------------------------------------------------------------
  definitions: {
    // -------------------------------------------------------------------------
    // Position (unchanged from Sprint 1)
    // -------------------------------------------------------------------------
    position: {
      type: 'object',
      required: ['x', 'y'],
      properties: {
        x: { type: 'number' },
        y: { type: 'number' },
      },
      additionalProperties: false,
    },

    // -------------------------------------------------------------------------
    // FR-08: Font registry — FontRef + RunFont
    // -------------------------------------------------------------------------

    fontRef: {
      type: 'object',
      required: [
        'id', 'family', 'weight', 'style', 'source',
        'fallback', 'subset', 'metricsCompatible', 'license', 'warnings',
      ],
      additionalProperties: false,
      properties: {
        id: {
          type: 'string',
          pattern: '^[a-z0-9-]+$',
          description:
            'Stable identifier referenced by text runs. ' +
            "Format: family-weight-style, lowercased, spaces/dots replaced by hyphens. " +
            "Example: 'open-sans-700-italic'.",
        },
        family: {
          type: 'string',
          minLength: 1,
          description: 'CSS font-family name. Matches <a:latin typeface=...> in PPTX.',
        },
        weight: {
          type: 'integer',
          minimum: 100,
          maximum: 900,
          description: 'CSS weight value read from OS/2 table. Defaults to 400 if absent.',
        },
        style: {
          enum: ['normal', 'italic', 'oblique'],
          description: "Read from the font's head/OS/2 tables.",
        },
        source: {
          enum: ['embedded', 'substituted', 'missing'],
          description:
            'embedded: extracted from /ppt/fonts/ inside the PPTX. ' +
            'substituted: original font unavailable; a replacement was used. ' +
            'missing: no file produced; generator must rely on CSS fallback only.',
        },
        file: {
          type: ['string', 'null'],
          description:
            'Relative path inside the output bundle, anchored at the reveal.js HTML. ' +
            'Null only when source=missing.',
        },
        format: {
          enum: ['woff2', 'woff', 'ttf'],
          description: 'Format of the file at `file`. Target format is woff2.',
        },
        fallback: {
          type: 'string',
          minLength: 1,
          description: 'CSS-ready fallback chain. Always non-empty.',
        },
        subset: {
          type: 'boolean',
          description:
            'True if the embedded font is a PPTX subset embed (only used glyphs). ' +
            'Surfaces as a warning because subsequent edits may show missing glyphs.',
        },
        metricsCompatible: {
          type: 'boolean',
          description:
            'True when glyph advance widths match the originally-requested font, ' +
            'so line breaks match PowerPoint. Always true for embedded, ' +
            'always false for missing. Drives the line-break fidelity report.',
        },
        license: {
          type: 'object',
          required: ['fsType', 'embeddable'],
          additionalProperties: false,
          properties: {
            fsType: {
              type: 'integer',
              description: 'Raw OS/2 fsType value from the font file. Stored for auditability.',
            },
            embeddable: {
              type: 'boolean',
              description:
                'Resolved embedding decision based on fsType. ' +
                'False blocks the font from being shipped.',
            },
          },
        },
        warnings: {
          type: 'array',
          items: { type: 'string' },
          description: 'Per-font human-readable notes. Bubbled into the API warnings array.',
        },
      },
      allOf: [
        {
          // When source is embedded or substituted, file and format must be present.
          if: { properties: { source: { enum: ['embedded', 'substituted'] } } },
          then: { required: ['file', 'format'] },
        },
        {
          if: { properties: { source: { const: 'missing' } } },
          then: {
            properties: {
              file: { type: 'null' },
              metricsCompatible: { const: false },
            },
          },
        },
        {
          if: { properties: { source: { const: 'embedded' } } },
          then: { properties: { metricsCompatible: { const: true } } },
        },
      ],
    },

    runFont: {
      type: 'object',
      required: ['ref', 'family', 'requestedWeight', 'requestedStyle'],
      additionalProperties: false,
      properties: {
        ref: {
          type: ['string', 'null'],
          description:
            "Pointer to a fontRef.id in the document's fonts array. " +
            'Null only when no entry could be created at all (indicates parser bug).',
        },
        family: {
          type: 'string',
          description:
            'Original family requested by PPTX. Retained even after resolution ' +
            'so the generator can render meaningful fallback CSS if ref is null.',
        },
        requestedWeight: {
          type: 'integer',
          minimum: 100,
          maximum: 900,
          description:
            "700 if <a:rPr b='1'/>, 400 otherwise. May differ from the resolved " +
            "fontRef.weight; generator falls back to CSS faux-bold in that case.",
        },
        requestedStyle: {
          enum: ['normal', 'italic', 'oblique'],
          description:
            'Style requested by the run. Generator falls back to CSS faux-italic ' +
            'when no matching italic variant is available.',
        },
      },
    },

    // -------------------------------------------------------------------------
    // FR-06 + FR-08: Text run (extended for Sprint 2)
    // -------------------------------------------------------------------------
    run: {
      type: 'object',
      required: ['text'],
      properties: {
        text: { type: 'string' },

        // FR-08: structured font reference (Sprint 2)
        font: { $ref: '#/definitions/runFont' },

        // FR-06: flat run-level formatting (Sprint 2)
        size: { type: 'number', description: 'Font size in points.' },
        color: { type: 'string', description: 'CSS color or theme variable.' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        strikethrough: { type: 'boolean' },

        // Sprint 1 formatting bag — retained for backward compatibility.
        // Sprint 1 parsers set formatting.weight / formatting.italics here;
        // Sprint 2 parsers prefer the flat fields above.
        formatting: {
          type: 'object',
          properties: {
            font: { type: 'string' },
            size: { type: 'string', description: 'CSS length, e.g. "14pt" or "24px".' },
            color: { type: 'string', description: 'CSS color or theme variable.' },
            weight: { type: 'string', enum: ['normal', 'bold'] },
            italics: { type: 'boolean' },
            'text-decoration': {
              type: 'string',
              enum: ['underline', 'strikethrough', 'none'],
            },
          },
          additionalProperties: true,
        },
        'super-sub-script': {
          type: 'string',
          enum: ['normal', 'super', 'sub'],
        },
        link: {
          type: 'object',
          properties: {
            href: { type: 'string' },
            target: { type: 'string' },
          },
          additionalProperties: true,
        },
      },
      additionalProperties: true,
    },

    // -------------------------------------------------------------------------
    // Paragraph (Sprint 1 structure preserved; FR-06 fields already covered)
    // -------------------------------------------------------------------------
    paragraph: {
      type: 'object',
      required: ['runs'],
      properties: {
        id: { type: 'string' },
        formatting: {
          type: 'object',
          properties: {
            font: { type: 'string' },
            size: { type: 'string' },
            color: { type: 'string' },
            weight: { type: 'string', enum: ['normal', 'bold'] },
            italics: { type: 'boolean' },
            'text-decoration': {
              type: 'string',
              enum: ['underline', 'strikethrough', 'none'],
            },
            // FR-06: paragraph-level formatting
            'line-spacing': { type: 'string' },
            'list-type': {
              type: 'string',
              enum: ['numbered', 'bullets', 'none'],
            },
            'indent-level': { type: 'integer', minimum: 0, maximum: 5 },
            margin: { type: 'string' },
            align: {
              type: 'string',
              enum: ['left', 'right', 'center', 'justify'],
            },
            'vertical-align': {
              type: 'string',
              enum: ['top', 'middle', 'bottom'],
            },
          },
          additionalProperties: true,
        },
        bullets: {},
        runs: {
          type: 'array',
          minItems: 1,
          items: { $ref: '#/definitions/run' },
        },
      },
      additionalProperties: true,
    },

    // -------------------------------------------------------------------------
    // Text block (unchanged from Sprint 1)
    // -------------------------------------------------------------------------
    textBlock: {
      type: 'object',
      required: ['paragraphs'],
      properties: {
        id: { type: 'string', description: 'Stable id used for animation refs (FR-14).' },
        'placeholder-id': { type: 'string' },
        position: { $ref: '#/definitions/position' },
        'pos-type': {
          type: 'string',
          enum: ['relative-to-placeholder', 'relative-to-slide'],
        },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
        overflow: {
          type: 'string',
          enum: ['auto-fit', 'shrink-on-overflow', 'overflow-visible', 'none'],
        },
        'z-index': { type: 'integer' },
        background: { type: 'string' },
        paragraphs: {
          type: 'array',
          minItems: 1,
          items: { $ref: '#/definitions/paragraph' },
        },
      },
      additionalProperties: true,
    },

    // -------------------------------------------------------------------------
    // Media (unchanged from Sprint 1)
    // -------------------------------------------------------------------------
    mediaItem: {
      type: 'object',
      required: ['file-link', 'media-type', 'position', 'width', 'height'],
      properties: {
        id: { type: 'string' },
        'file-link': {
          type: 'string',
          description: 'Path or URL relative to the output bundle.',
        },
        'media-type': {
          type: 'string',
          enum: ['image', 'video'],
        },
        position: { $ref: '#/definitions/position' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: { type: 'number' },
        'z-index': { type: 'integer' },
        scale: { type: 'number' },
        crop: {
          type: 'array',
          items: { type: 'number' },
          minItems: 4,
          maxItems: 4,
          description: '[top, right, bottom, left]',
        },
        effects: { type: 'object', additionalProperties: true },
        playback: { type: 'object', additionalProperties: true },
      },
      additionalProperties: true,
    },

    // -------------------------------------------------------------------------
    // FR-10: Shape sub-definitions
    // -------------------------------------------------------------------------

    shapeFill: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: ['solid', 'gradient', 'pattern', 'none'],
        },
        color: { type: 'string', description: 'CSS color or theme variable. Used when type=solid.' },
        stops: {
          type: 'array',
          items: {
            type: 'object',
            required: ['position', 'color'],
            properties: {
              position: { type: 'number', minimum: 0, maximum: 1 },
              color: { type: 'string' },
            },
            additionalProperties: false,
          },
          description: 'Gradient colour stops. Used when type=gradient.',
        },
        angle: {
          type: 'number',
          description: 'Gradient angle in degrees. Used when type=gradient.',
        },
      },
      additionalProperties: true,
    },

    shapeStroke: {
      type: 'object',
      properties: {
        color: { type: 'string' },
        width: { type: 'number', description: 'Stroke width in points.' },
        style: {
          type: 'string',
          enum: ['solid', 'dashed', 'dotted', 'none'],
        },
        cap: {
          type: 'string',
          enum: ['flat', 'round', 'square'],
        },
      },
      additionalProperties: true,
    },

    shapeGeometry: {
      type: 'object',
      description: 'Shape-specific geometry data (varies by shape type).',
      properties: {
        points: {
          type: 'array',
          items: { $ref: '#/definitions/position' },
          description: 'Vertex list for polyline / polygon / connector.',
        },
        startArrow: {
          type: 'string',
          enum: ['none', 'triangle', 'stealth', 'diamond', 'oval', 'open'],
          description: 'Arrowhead at the start of line / arrow / connector.',
        },
        endArrow: {
          type: 'string',
          enum: ['none', 'triangle', 'stealth', 'diamond', 'oval', 'open'],
          description: 'Arrowhead at the end of line / arrow / connector.',
        },
        calloutPoint: {
          $ref: '#/definitions/position',
          description: 'Tip of the callout tail.',
        },
      },
      additionalProperties: true,
    },

    // FR-10: Shape — top-level shape element
    shape: {
      type: 'object',
      required: ['type'],
      properties: {
        id: { type: 'string' },
        type: {
          type: 'string',
          enum: ['rect', 'ellipse', 'line', 'arrow', 'polyline', 'polygon', 'callout', 'connector'],
        },
        position: { $ref: '#/definitions/position' },
        width: { type: 'number' },
        height: { type: 'number' },
        rotation: {
          type: 'number',
          description: 'Clockwise rotation in degrees.',
        },
        'z-index': { type: 'integer' },
        fill: { $ref: '#/definitions/shapeFill' },
        stroke: { $ref: '#/definitions/shapeStroke' },
        geometry: { $ref: '#/definitions/shapeGeometry' },
        text: {
          type: 'array',
          items: { $ref: '#/definitions/paragraph' },
          description: 'Embedded text-frame paragraphs.',
        },
      },
      additionalProperties: true,
    },
  },
};

module.exports = sprint2Schema;
