# PPTX-to-reveal.js Converter

Interdisciplinary Software Development Project — Hof University

Converts PowerPoint (`.pptx`) presentations into self-contained [reveal.js](https://revealjs.com/) HTML presentations. Upload a `.pptx` through the web UI (or the HTTP API), preview the result in the browser, and download it as a portable bundle that runs offline.

## What it converts

The parser walks the OOXML structure (slides, layouts, masters, theme) and builds an intermediate representation (IR), which the generator renders to HTML/SVG/CSS. Supported today:

- **Text** — paragraphs, runs, formatting (bold/italic/underline, color, font, size), bullet/numbered lists, hyperlinks, tab stops.
- **Shapes** — presets (rectangles, ellipses, arrows, stars, callouts, chevrons, custom freeform geometry, etc.), solid and gradient fills, strokes with dash styles and arrowheads, drop shadow / inner shadow / glow / soft-edge effects, rotation and flipping, grouped shapes.
- **Theme** — color scheme resolved to CSS custom properties (`var(--theme-accentN)`), including `lumMod`/`lumOff`/`tint`/`shade` color variants, `fillRef`/`lnRef`/`effectRef` style inheritance.
- **Tables** — cell styles, borders, header rows.
- **Media** — images (with cropping), embedded fonts.
- **Layouts & masters** — placeholder inheritance (position, formatting, background) from slide → layout → master.
- **Animations** — entrance/exit/emphasis effects mapped to reveal.js-compatible output where possible.

Shape presets without a mapped renderer log a warning (`shape preset "X" not yet supported`); the shape still carries its parsed position/fill/stroke in the IR even if nothing renders for it visually. A few effects with no faithful web equivalent (3D rotation/bevel, pattern/texture fills) aren't parsed at all yet.

## Requirements

- Node.js **>= 18** (enforced via `engines` + `.npmrc`'s `engine-strict=true`; see `.nvmrc`)

## Setup

```bash
git clone <this repo>
cd pptx-revealjs-converter
npm install
```

## Running

```bash
npm run dev
```

Then open **http://localhost:3000/** and upload a `.pptx` file.

There's also a health check at `GET /api/v1/health`.

## API

All endpoints are under `/api/v1`.

| Method | Path | Description |
|---|---|---|
| `POST` | `/convert` | Upload a `.pptx` (multipart field `file`, max 50 MB). Returns `{ result_id, preview_url, download_url, warnings, statistics }`. |
| `GET` | `/preview/:id` | Render the converted result inline in the browser. |
| `GET` | `/result/:id` | Download the result as a ZIP (`index.html` + `assets/` media + a bundled `reveal/` distribution for offline use). |
| `GET` | `/media/:id/:filename` | Serve an individual media asset referenced by a preview. |
| `GET` | `/health` | Liveness check; returns Node/converter version. |
| `GET` | `/docs` | Interactive Swagger UI for this table (see below). |

Conversion results are held in memory for 30 minutes after upload, then discarded (see `src/storage/resultStore.js`) — there's no persistent storage or database.

Uploads are sanitized before parsing (VBA macro stripping, inline SVG script/handler removal, HTML-import rejection) and validated against a JSON Schema IR contract before rendering.

### Interactive API docs (OpenAPI)

Every endpoint above is annotated with `@openapi` JSDoc blocks (see `src/api/upload.js`, `src/api/download.js`, `src/app.js`), assembled by `swagger-jsdoc` into a single spec (`src/api/openapi.js`).

- Browse it live at **`/api/v1/docs`** (Swagger UI) while the server is running.
- Export the machine-readable spec to `openapi.json` at the repo root:
  ```bash
  npm run openapi:export
  ```
  (generated, gitignored — regenerate whenever the `@openapi` annotations change.)

## Testing

```bash
npm test           # Jest unit/integration tests
npm run lint       # ESLint (eslint:recommended + eslint-plugin-security)
npm run format     # Prettier — write
npm run format:check
```

Playwright-based end-to-end/cross-browser tests live under `tests/e2e/`.

### Code documentation

Every module has a JSDoc header and every exported function has `@param`/`@returns` tags. Generate browsable HTML docs (default JSDoc template) with:

```bash
npm run docs       # writes to docs/api/ (generated, gitignored)
```

## Project structure

```
src/
  api/                Express routes (upload, download, error handling)
  security/            Upload sanitization (NFR-08)
  parser/pptx/          .pptx → IR (slides, shapes, text, tables, theme, animations, media)
  ir/                  IR JSON Schema + validator
  generator/revealjs/    IR → reveal.js HTML/SVG/CSS
  storage/             In-memory result store (TTL-based)
  web/                 Static upload UI (served at "/")
tests/                 Jest unit/integration tests, Playwright e2e specs, fixtures
```

`src/generator/pptx/` contains an early, currently-unwired IR→PPTX (reverse conversion) module — not part of the forward conversion path used by the API today.
