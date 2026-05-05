/**
 * Web-Slide Internal Data Format — Sprint 1 JSON Schema
 *
 * This schema defines the minimal subset of the shared web-slide internal
 * data format required for the Sprint 1 MVP of the PPTX -> reveal.js converter.
 *
 * Sprint 1 covers (per Specification §8.1):
 *   - File upload + format validation (FR-01, FR-02)
 *   - Text extraction from slides (FR-03)
 *   - Slide order preservation (FR-04)
 *   - reveal.js HTML output (FR-05)
 *   - Image inclusion (FR-09)
 *   - In-browser viewing (FR-15)
 *
 * NOT covered in Sprint 1 (deferred to later sprints):
 *   - Master theme + color variables (FR-12, Sprint 2)
 *   - Master/slide layouts (FR-11, Sprint 2)
 *   - Shapes (FR-10, Sprint 2)
 *   - Stacking order via z-index (FR-13, Sprint 2)
 *   - Animations (FR-14, Sprint 2)
 *   - Tables, groups, transitions, notes (Sprint 2/3)
 *
 * EXTENSIBILITY:
 *   We use `additionalProperties: true` on element-level objects so that
 *   Sprint 2 fields (e.g. shapes, animations, master theme) can be added
 *   without breaking Sprint 1 documents already in fixtures or tests.
 *   We use `additionalProperties: false` only on the top-level document
 *   to catch typos in the contract itself.
 *
 * AUTHORITY:
 *   Source of truth is /docs/web-slide-internal-data.example.json. Any
 *   mismatch between this schema and the example file must be raised
 *   with the cross-group sync, not silently fixed here.
 */

const slideSchema = {
  type: 'object',
  required: ['contents'],
  properties: {
    // Per the example format, `title` is "a single paragraph"; we keep it
    // as a free-form string for Sprint 1 since the parser only emits plain
    // text titles. Will be widened to a paragraph object in Sprint 2 when
    // we support per-run formatting on titles.
    title: {
      oneOf: [
        { type: 'string' },
        {
          type: 'object',
          properties: {
            content: { type: 'string' },
          },
          additionalProperties: true,
        },
      ],
    },
    'layout-id': {
      type: 'string',
      description: 'Reference to layouts[] entry. Optional in Sprint 1 (no layout extraction).',
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
        // Sprint 2+ fields kept open so future IR docs validate:
        shapes: { type: 'array' },
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

const sprint1Schema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  $id: 'https://hof-university.de/web-slide-internal-data/sprint1.schema.json',
  title: 'Web-Slide Internal Data Format (Sprint 1 subset)',
  description:
    'Minimal IR schema for the PPTX -> reveal.js converter MVP. ' +
    'Validates documents containing slides with text and media only.',
  type: 'object',
  required: ['slideset'],
  additionalProperties: false,
  properties: {
    slideset: {
      type: 'object',
      required: ['slides'],
      additionalProperties: true,
      properties: {
        // -- Document metadata (FR-03 partial: filename always known) --
        filename: { type: 'string' },
        title: { type: 'string' },
        author: { type: 'string' },
        'creation-date': {
          type: 'string',
          // Example file says "yyyy-mm-dd"; we accept full ISO 8601 too
          // since PPTX core.xml uses dateTime. Confirm with cross-group sync.
          pattern: '^\\d{4}-\\d{2}-\\d{2}(T.*)?$',
        },

        // -- Fonts list (FR-08, Sprint 2) — declared for forward-compat --
        fonts: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              'font-id': { type: 'string' },
              'font-file': { type: 'string' },
            },
            additionalProperties: true,
          },
        },

        // -- Master section (FR-11, FR-12, Sprint 2) — optional in Sprint 1 --
        master: {
          type: 'object',
          additionalProperties: true,
        },

        // -- Layouts list (FR-11, Sprint 2) — optional in Sprint 1 --
        layouts: {
          type: 'array',
          items: { type: 'object', additionalProperties: true },
        },

        // -- Slides (Sprint 1 core) --
        slides: {
          type: 'array',
          minItems: 0,
          items: slideSchema,
        },
      },
    },
  },

  // ---------------------------------------------------------------------
  // Reusable definitions
  // ---------------------------------------------------------------------
  definitions: {
    // Position used by text blocks and media. Coordinates are stored in
    // the unit declared by slideset.master['dimension-units']; default px.
    // The PPTX parser will produce EMU-derived values converted to px,
    // per Specification §3.3.
    position: {
      type: 'object',
      required: ['x', 'y'],
      properties: {
        x: { type: 'number' },
        y: { type: 'number' },
      },
      additionalProperties: false,
    },

    // -- Run: equivalent to <span> in HTML. Carries inline formatting. --
    // Sprint 1 supports text + the formatting subset needed for FR-06
    // (bold, italics, underline, strikethrough, color, font, size).
    run: {
      type: 'object',
      required: ['text'],
      properties: {
        text: { type: 'string' },
        formatting: {
          type: 'object',
          properties: {
            font: { type: 'string' },
            size: { type: 'string', description: 'CSS length, e.g. "14pt" or "24px"' },
            color: { type: 'string', description: 'CSS color or theme variable' },
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

    // -- Paragraph: ordered list of runs with paragraph-level formatting. --
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

    // -- Text block: a positioned container of paragraphs (a PPTX text frame). --
    textBlock: {
      type: 'object',
      required: ['paragraphs'],
      properties: {
        id: { type: 'string', description: 'Stable id used for animation refs (Sprint 2)' },
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

    // -- Media: image or video (FR-09). Sprint 1 = image only; video --
    // -- accepted by schema but generator will emit static placeholder --
    // -- per OoS-02 of the Requirements Analysis.
    mediaItem: {
      type: 'object',
      required: ['file-link', 'media-type', 'position', 'width', 'height'],
      properties: {
        id: { type: 'string' },
        'file-link': {
          type: 'string',
          description: 'Path or URL relative to the output bundle',
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
  },
};

module.exports = sprint1Schema;
