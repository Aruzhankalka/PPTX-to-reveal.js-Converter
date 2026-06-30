/**
 * Web-Slide Internal Data Format — Sprint 2/3 JSON Schema
 *
 * Sprint 1 features (unchanged):
 *   FR-01/02  File upload + format validation
 *   FR-03     Text extraction from slides
 *   FR-04     Slide order preservation
 *   FR-05     reveal.js HTML output
 *   FR-09     Image inclusion
 *   FR-15     In-browser viewing
 *
 * Sprint 2 additions:
 *   FR-06  Full text formatting: bold, italic, underline, strikethrough,
 *          color, font family/size, line spacing, alignment, lists.
 *   FR-08  Font registry: document-wide FontRef entries with per-run RunFont
 *          pointers. Supports embedded (WOFF2) / substituted / missing fonts.
 *   FR-10  Shapes: rect, roundRect, ellipse, line, arrow, polyline, polygon,
 *          callout, connector, unknown — each with fill, stroke, geometry,
 *          rotation, and an optional embedded TextBlock.
 *   FR-14  Animations: per-slide animation array with trigger, effect, timing,
 *          and fidelity signal. targetId cross-reference validated at runtime.
 *
 * Sprint 1 documents remain valid; all Sprint 2/3 additions are optional at
 * the schema level so the validator can serve both generations of IR documents.
 *
 * NOT covered (deferred):
 *   FR-13     Stacking order (z-index) — handled via z field on shapes
 *   Tables    typed loosely (bare array); table-style/is-header/border unmapped
 *   Transitions, notes
 *
 * FR-11/12 (Master/layouts + theme colors) are implemented: master.theme,
 * master['color-theme'], master['slide-dimensions'], master['aspect-ratio'],
 * and master.formatting (presentation-wide default text style, from
 * presentation.xml's <p:defaultTextStyle>) are all populated. layouts[]
 * carries a `placeholders` array per layout (definitions.placeholder below),
 * built by master.js's buildLayoutPlaceholders from each layout's own
 * placeholder shapes.
 *
 * Groups (<p:grpSp>) are implemented: parseShapes() walks p:grpSp at any
 * nesting depth, emits one `group` entry per container (definitions.group
 * below), and corrects every descendant shape/picture's position+rotation
 * through the composed chOff/chExt scale + rotation transform so it lands in
 * absolute slide EMU rather than raw group-local coordinates.
 *
 * EXTENSIBILITY:
 *   `additionalProperties: true` on element-level objects lets future sprints
 *   add fields without breaking existing IR documents.
 *   `additionalProperties: false` is reserved for closed-set definitions
 *   (fontRef, runFont, position, license, shapePos, shapeColor) where the
 *   contract is precise.
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
        // FR-14 (Sprint 3): per-slide animation sequence
        animations: {
          type: 'array',
          items: { $ref: '#/definitions/animation' },
        },
        tables: { type: 'array' },
        // FR-10 (extended): groups[] — one entry per <p:grpSp>
        groups: {
          type: 'array',
          items: { $ref: '#/definitions/group' },
        },
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
  title: 'Web-Slide Internal Data Format (Sprint 2/3)',
  description:
    'IR schema for the PPTX -> reveal.js converter. ' +
    'Sprint 2 adds full text formatting (FR-06), font extraction (FR-08), ' +
    'and shapes (FR-10). Sprint 3 adds animations (FR-14).',
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
          properties: {
            // Spec's aspect-ratio is a closed enum — PPTX allows arbitrary
            // custom slide sizes, so the parser clamps to whichever of the
            // two is numerically closest (see nearestAspectRatio in
            // parser/pptx/index.js) rather than emitting e.g. "8:5".
            'aspect-ratio': { type: 'string', enum: ['16:9', '4:3'] },
            // Presentation-wide default text style ("Global preset!"), read
            // from <p:defaultTextStyle><a:lvl1pPr> in presentation.xml — the
            // top of the formatting inheritance chain (see
            // getDefaultTextStyle in parser/pptx/slides.js).
            formatting: { $ref: '#/definitions/formatting' },
          },
        },

        // FR-11 (Sprint 3): slide layouts
        layouts: {
          type: 'array',
          items: { $ref: '#/definitions/layout' },
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
    // Position (unchanged from Sprint 1 — used by textBlock / mediaItem)
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
    // FR-06: Formatting bag — shared shape for paragraph.formatting and
    // master.formatting (the spec's presentation-wide "Global preset!").
    // run.formatting carries the same field names but is defined separately
    // (it omits margin/list-type/vertical-align, which don't apply to a run).
    // -------------------------------------------------------------------------
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

    // -------------------------------------------------------------------------
    // FR-10/14: Shape geometry position — x, y, w, h in EMU (integers).
    // Required on every shape element, including unknown stubs, so embedded
    // text always has a bounding box.
    // -------------------------------------------------------------------------
    // shapePos kept for gradient stop + effect schema refs; shape geometry has
    // moved to position:{x,y} (px, definitions.position) + top-level width/height.
    shapePos: {
      type: 'object',
      required: ['x', 'y'],
      additionalProperties: true,
      properties: {
        x: { type: 'number' },
        y: { type: 'number' },
      },
    },

    // -------------------------------------------------------------------------
    // FR-10/14: Structured color — keeps theme references unresolved so
    // FR-12 can emit var(--accent1) downstream.  The generator converts to
    // CSS, not the IR.
    // -------------------------------------------------------------------------
    shapeColor: {
      oneOf: [
        {
          type: 'object',
          required: ['space', 'hex'],
          additionalProperties: false,
          properties: {
            space: { type: 'string', const: 'srgb' },
            hex: {
              type: 'string',
              pattern: '^[0-9A-Fa-f]{6}$',
              description: 'Six-digit uppercase hex, e.g. "4472C4".',
            },
            alpha: {
              type: 'integer',
              minimum: 0,
              maximum: 100,
              description: 'Opacity in percent (0 = fully transparent, 100 = fully opaque). Absent when fully opaque.',
            },
          },
        },
        {
          type: 'object',
          required: ['space', 'ref'],
          additionalProperties: false,
          properties: {
            space: { type: 'string', const: 'theme' },
            ref: {
              type: 'string',
              enum: [
                'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
                'text1', 'text2', 'bg1', 'bg2', 'link', 'linkVisited',
              ],
              description: 'Theme color slot name.',
            },
            alpha: {
              type: 'integer',
              minimum: 0,
              maximum: 100,
              description: 'Opacity in percent. Absent when fully opaque.',
            },
          },
        },
      ],
    },

    // -------------------------------------------------------------------------
    // FR-10 (extended): Gradient stop — one position+color pair in a gradient
    // -------------------------------------------------------------------------
    gradientStop: {
      type: 'object',
      required: ['pos', 'color'],
      additionalProperties: false,
      properties: {
        pos: {
          type: 'integer',
          minimum: 0,
          maximum: 100000,
          description: 'Stop position in PPTX units (0–100000 maps to 0%–100%).',
        },
        color: { $ref: '#/definitions/shapeColor' },
      },
    },

    // -------------------------------------------------------------------------
    // FR-10 (extended): Arrow end marker — used in stroke headEnd / tailEnd
    // -------------------------------------------------------------------------
    arrowEnd: {
      type: 'object',
      required: ['type'],
      additionalProperties: false,
      properties: {
        type: {
          type: 'string',
          enum: ['none', 'triangle', 'stealth', 'arrow', 'diamond', 'oval'],
        },
        width:  { type: 'string', enum: ['sm', 'med', 'lg'] },
        length: { type: 'string', enum: ['sm', 'med', 'lg'] },
      },
    },

    // -------------------------------------------------------------------------
    // FR-10: Shape fill — spec shape: {type, color?}
    // solid.color is a flat CSS string ('#RRGGBB', 'rgba(...)', 'var(--theme-X)').
    // Gradient keeps stops as structured shapeColor for SVG alpha precision.
    // -------------------------------------------------------------------------
    shapeFill: {
      oneOf: [
        {
          type: 'object',
          required: ['type'],
          additionalProperties: false,
          properties: {
            type: { type: 'string', const: 'none' },
          },
        },
        {
          type: 'object',
          required: ['type', 'color'],
          additionalProperties: false,
          properties: {
            type:  { type: 'string', const: 'solid' },
            color: { type: 'string', description: 'Flat CSS color: #RRGGBB or var(--theme-X).' },
            alpha: { type: 'number', minimum: 0, maximum: 1,
                     description: 'Opacity (0=transparent, 1=opaque). Absent when fully opaque. Generator emits as SVG fill-opacity.' },
          },
        },
        {
          type: 'object',
          required: ['type', 'kind', 'stops'],
          additionalProperties: false,
          properties: {
            type:  { type: 'string', const: 'gradient' },
            kind:  { type: 'string', enum: ['linear', 'radial'] },
            angle: {
              type: 'integer',
              description: 'Native PPTX angle units (1/60000 deg). Linear kind only.',
            },
            stops: {
              type: 'array',
              minItems: 2,
              items: { $ref: '#/definitions/gradientStop' },
            },
          },
        },
      ],
    },

    // -------------------------------------------------------------------------
    // FR-10: Shape stroke — spec shape: {type, color, width, style}
    // color is a flat CSS string; width is in CSS px; style matches spec enum.
    // headEnd/tailEnd are extensions (arrowheads, not in spec but kept for rendering).
    // -------------------------------------------------------------------------
    shapeStroke: {
      oneOf: [
        {
          type: 'object',
          required: ['type'],
          additionalProperties: false,
          properties: {
            type: { type: 'string', const: 'none' },
          },
        },
        {
          type: 'object',
          required: ['type', 'color', 'width', 'style'],
          additionalProperties: false,
          properties: {
            type:  { type: 'string', const: 'solid' },
            color: { type: 'string', description: 'Flat CSS color string.' },
            width: { type: 'number', description: 'Stroke width in CSS px.' },
            style: {
              type: 'string',
              enum: ['solid', 'dashed', 'pointed', 'none'],
              description: 'Dash pattern: solid (no dash), dashed, pointed (dotted), none.',
            },
            headEnd: { $ref: '#/definitions/arrowEnd' },
            tailEnd: { $ref: '#/definitions/arrowEnd' },
          },
        },
      ],
    },

    // -------------------------------------------------------------------------
    // FR-08: Font registry — FontRef + RunFont
    // -------------------------------------------------------------------------

    fontRef: {
      type: 'object',
      required: [
        'font-id', 'family', 'weight', 'style', 'source',
        'fallback', 'subset', 'metricsCompatible', 'license', 'warnings',
      ],
      additionalProperties: false,
      properties: {
        'font-id': {
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
        'font-file': {
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
          // When source is embedded or substituted, font-file and format must be present.
          if: { properties: { source: { enum: ['embedded', 'substituted'] } } },
          then: { required: ['font-file', 'format'] },
        },
        {
          if: { properties: { source: { const: 'missing' } } },
          then: {
            properties: {
              'font-file': { type: 'null' },
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
    //
    // Two variants are valid:
    //   Normal text run: { text: '...' [, formatting, link, ...] }
    //   Tab run:         { type: 'tab' }  — no text; marks a tab-stop position.
    //     Produced when <a:tab/> appears as a paragraph child or when <a:t>
    //     contains a literal U+0009 tab character (split at parse time).
    // -------------------------------------------------------------------------
    run: {
      type: 'object',
      // text is required for normal runs; tab runs use type:'tab' and omit text
      properties: {
        type: {
          type: 'string',
          description: '"tab" marks a tab-stop positioning run with no text content.',
        },
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
        formatting: { $ref: '#/definitions/formatting' },
        bullets: {},
        // Tab stop positions parsed from <a:pPr><a:tabLst><a:tab l=".." algn=".."/>
        tabStops: {
          type: 'array',
          items: {
            type: 'object',
            required: ['pos', 'align'],
            additionalProperties: false,
            properties: {
              pos:   { type: 'integer', description: 'Tab stop position in EMU from the left margin.' },
              align: { type: 'string',  enum: ['l', 'r', 'ctr', 'dec'], description: 'Tab stop alignment.' },
            },
          },
          description: 'Explicit tab stops from <a:pPr><a:tabLst>. Absent when no stops are defined.',
        },
        runs: {
          type: 'array',
          minItems: 0, // 0 allows empty paragraphs (blank lines between title lines)
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
        autoFit: {
          type: 'string',
          enum: ['none', 'norm', 'shape'],
          description:
            'PPTX <a:bodyPr> auto-fit mode. ' +
            'none → <a:noAutofit/> text is clipped at the box boundary. ' +
            'norm → <a:normAutofit/> font/spacing were scaled to fit (fontScale already applied by parser). ' +
            'shape → <a:spAutoFit/> the shape grew to contain its text. ' +
            'Absent when the PPTX did not specify a mode (PowerPoint treats this as none).',
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
    // FR-11: Placeholder — one layout slot (title/body/footer/...), as
    // produced by master.js's buildLayoutPlaceholders. position/width/height
    // are absent when the placeholder has no explicit <a:xfrm> of its own and
    // inherits geometry from the master (not an error — just nothing to report
    // at this level).
    // -------------------------------------------------------------------------
    placeholder: {
      type: 'object',
      required: ['placeholder-id'],
      additionalProperties: true,
      properties: {
        'placeholder-id': { type: 'string' },
        position: { $ref: '#/definitions/position' },
        width: { type: 'number' },
        height: { type: 'number' },
        padding: { type: 'string', description: 'CSS shorthand, e.g. "8px 91px 4px 91px" (top right bottom left).' },
        type: { type: 'string', enum: ['text', 'image', 'video', 'table', 'other'] },
        role: { type: 'string', enum: ['title', 'subtitle', 'body', 'footer', 'date', 'slide-number'] },
        background: { type: 'string' },
        formatting: { $ref: '#/definitions/formatting' },
      },
    },

    // -------------------------------------------------------------------------
    // FR-11: Layout — one entry per slide layout, with its placeholder slots.
    // -------------------------------------------------------------------------
    layout: {
      type: 'object',
      required: ['layout-id'],
      additionalProperties: true,
      properties: {
        'layout-id': { type: 'string' },
        name: { type: ['string', 'null'] },
        placeholders: {
          type: 'array',
          items: { $ref: '#/definitions/placeholder' },
        },
      },
    },

    // -------------------------------------------------------------------------
    // FR-10: Shape — typed shape element with EMU geometry
    //
    // position: { x, y, w, h } in EMU — REQUIRED including for unknown stubs
    //   so embedded text always has a bounding box.
    // rotation: integer in native PPTX rot units (1/60000 of a degree), 0 if
    //   absent.  The generator converts to CSS degrees, not the IR.
    // position: {x,y} in CSS px (definitions.position); width/height as top-level px.
    // type: spec vocabulary (rectangle|triangle|ellipsis|...|custom).
    // subtype: internal rendering vocabulary (rect|roundRect|ellipse|...) for generator.
    // fill.color / stroke.color: flat CSS strings (plain '#RRGGBB' or 'var(--theme-X)').
    // stroke.width: CSS px; stroke.style: solid|dashed|pointed|none.
    // paragraphs: embedded text directly on shape (spec); text-anchor/text-insets separate.
    // config: type-specific key-value bag (spec field).
    // -------------------------------------------------------------------------
    shape: {
      type: 'object',
      required: ['id', 'type', 'position', 'z-index'],
      additionalProperties: true,
      properties: {
        id: {
          type: 'string',
          description: 'Unique within the slide; used as animation targetId.',
        },
        type: {
          type: 'string',
          enum: ['rectangle', 'triangle', 'ellipsis', 'line', 'connector',
                 'polyline', 'polygon', 'callout', 'arrow', 'star',
                 'cloud', 'database', 'chevron', 'custom'],
          description: 'Spec vocabulary. Use subtype for generator dispatch.',
        },
        subtype: {
          type: 'string',
          description: 'Internal rendering type (rect|roundRect|ellipse|pentagon|arc|…). ' +
            'Generator switch dispatches on this, not type.',
        },
        supported: {
          type: 'boolean',
          description: 'false when the generator cannot fully render this shape.',
        },
        position: {
          $ref: '#/definitions/position',
          description: 'Top-left corner in CSS px (same convention as text/media/layouts).',
        },
        width:  { type: 'number', description: 'Shape width in CSS px.' },
        height: { type: 'number', description: 'Shape height in CSS px.' },
        rotation: {
          type: 'number',
          default: 0,
          description: 'Clockwise rotation in degrees.',
        },
        flipH: { type: 'boolean', default: false },
        flipV: { type: 'boolean', default: false },
        fill:   { $ref: '#/definitions/shapeFill' },
        stroke: { $ref: '#/definitions/shapeStroke' },
        config: {
          type: 'object',
          additionalProperties: true,
          description: 'Type-specific settings (spec field). Derived from adjustments[].',
        },
        paragraphs: {
          type: 'array',
          items: { $ref: '#/definitions/paragraph' },
          description: 'Embedded text — spec places paragraphs[] directly on the shape.',
        },
        'text-anchor': {
          type: 'string',
          enum: ['t', 'ctr', 'b'],
          description: 'Vertical text anchor inside the shape bounding box.',
        },
        'text-insets': {
          type: 'object',
          description: 'Body padding in EMU {l,r,t,b}.',
          additionalProperties: true,
        },
        adjustments: {
          type: 'array',
          items: {
            type: 'object',
            required: ['name', 'value'],
            additionalProperties: false,
            properties: {
              name:  { type: 'string' },
              value: { type: 'number' },
            },
          },
          description: 'Geometry guide overrides, e.g. [{name:"adj",value:16667}] for roundRect.',
        },
        effects: {
          type: 'object',
          additionalProperties: true,
          properties: {
            shadow: {
              type: 'object',
              required: ['mode', 'color', 'blurEmu', 'distanceEmu', 'directionAngle', 'alphaPct'],
              additionalProperties: false,
              properties: {
                mode:           { type: 'string', enum: ['outer', 'inner'] },
                color:          { $ref: '#/definitions/shapeColor' },
                blurEmu:        { type: 'integer', minimum: 0 },
                distanceEmu:    { type: 'integer', minimum: 0 },
                directionAngle: { type: 'integer' },
                alphaPct:       { type: 'integer', minimum: 0, maximum: 100 },
              },
            },
          },
        },
        customGeometry: {
          type: 'object',
          required: ['w', 'h', 'paths'],
          additionalProperties: false,
          properties: {
            w: { type: 'integer', description: 'Path coordinate-space width in EMU.' },
            h: { type: 'integer', description: 'Path coordinate-space height in EMU.' },
            paths: {
              type: 'array',
              items: {
                type: 'object',
                required: ['commands'],
                additionalProperties: false,
                properties: {
                  commands: {
                    type: 'array',
                    items: {
                      type: 'object',
                      required: ['op'],
                      additionalProperties: false,
                      properties: {
                        op: {
                          type: 'string',
                          enum: ['moveTo', 'lnTo', 'cubicBezTo', 'quadBezTo', 'arcTo', 'close'],
                        },
                        pts: {
                          type: 'array',
                          items: {
                            type: 'object',
                            required: ['x', 'y'],
                            additionalProperties: false,
                            properties: {
                              x: { type: 'integer' },
                              y: { type: 'integer' },
                            },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
        points: {
          type: 'array',
          items: {
            type: 'object',
            required: ['x', 'y'],
            additionalProperties: false,
            properties: {
              x: { type: 'integer', description: 'X coordinate in EMU.' },
              y: { type: 'integer', description: 'Y coordinate in EMU.' },
            },
          },
          description: 'Vertex list in EMU — required for polyline, polygon, connector.',
        },
        text: {
          $ref: '#/definitions/textBlock',
          description: 'Embedded text-frame. Position is the shape bounding box.',
        },
        'z-index': {
          type: 'integer',
          description: 'Z-index from spTree document order.',
        },
      },
      allOf: [
        {
          // polyline / polygon / connector require a vertex list (check subtype).
          if: {
            properties: {
              subtype: { enum: ['polyline', 'polygon', 'connector'] },
            },
          },
          then: { required: ['points'] },
        },
      ],
    },

    // -------------------------------------------------------------------------
    // FR-10 (extended): Group — one entry per <p:grpSp> container, at any
    // nesting depth.
    //
    // elements[]: ids of this group's DIRECT children only (shapes, media, or
    //   nested groups) — a membership list, not a z-order; each member still
    //   carries its own absolute z-index.
    // position: bounding box in absolute slide EMU. The parser pre-applies
    //   every ancestor group's chOff/chExt scale + rotation transform, so this
    //   (and every descendant shape/picture's own position) is already in
    //   slide coordinates — never raw group-local coordinates.
    // rotation: the group's own rotation composed with its ancestors', same
    //   convention as shape.rotation (degrees).
    // -------------------------------------------------------------------------
    group: {
      type: 'object',
      required: ['id', 'elements', 'position', 'z-index'],
      additionalProperties: true,
      properties: {
        id: {
          type: 'string',
          description: 'Unique within the slide.',
        },
        elements: {
          type: 'array',
          items: { type: 'string' },
          description: "ids of this group's direct children.",
        },
        position: {
          $ref: '#/definitions/position',
          description: 'Top-left corner in CSS px (ancestor group transforms already applied).',
        },
        width:  { type: 'number', description: 'Group width in CSS px.' },
        height: { type: 'number', description: 'Group height in CSS px.' },
        rotation: {
          type: 'number',
          default: 0,
          description: 'Clockwise rotation in degrees, composed with ancestor groups.',
        },
        'z-index': {
          type: 'integer',
          description: 'Z-index from spTree document order.',
        },
      },
    },

    // -------------------------------------------------------------------------
    // FR-14: Animation — per-element animation entry on a slide.
    //
    // targetId references an element id on the same slide (textBlock, shape, or
    // mediaItem).  Cross-document referential integrity is validated at runtime
    // by the validator's validateTargetIds() helper, not in JSON Schema.
    //
    // supported: false means the generator skips this entry and a warning was
    //   already pushed to the conversion warnings array (same pattern as
    //   fontRef.metricsCompatible).
    // -------------------------------------------------------------------------
    animation: {
      type: 'object',
      required: ['id', 'targetId', 'trigger', 'sequence', 'effect', 'timing', 'supported'],
      additionalProperties: true,
      properties: {
        id: { type: 'string', description: 'Unique within the slide.' },
        targetId: {
          type: 'string',
          description: 'id of the animated element on the same slide.',
        },
        trigger: {
          type: 'string',
          enum: ['onClick', 'withPrevious', 'afterPrevious'],
        },
        sequence: {
          type: 'integer',
          description: 'Zero-based sequence index within the slide.',
        },
        effect: {
          type: 'object',
          required: ['class', 'preset'],
          additionalProperties: false,
          properties: {
            class: {
              type: 'string',
              enum: ['entrance', 'emphasis', 'exit', 'motionPath'],
            },
            preset: {
              type: 'string',
              description: 'Human-readable preset name, e.g. "fade", "flyIn", "appear".',
            },
          },
        },
        timing: {
          type: 'object',
          required: ['delayMs', 'durationMs'],
          additionalProperties: false,
          properties: {
            delayMs: { type: 'integer', minimum: 0 },
            durationMs: { type: 'integer', minimum: 0 },
          },
        },
        supported: {
          type: 'boolean',
          description:
            'False when the generator cannot render this effect; a warning was ' +
            'pushed to the conversion warnings array. Analogous to ' +
            'fontRef.metricsCompatible.',
        },
      },
    },
  },
};

module.exports = sprint2Schema;
