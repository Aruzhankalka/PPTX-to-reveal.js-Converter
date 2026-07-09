# PPTX to reveal.js Converter

Interdisciplinary SW-Dev Project, Hochschule Hof, SS 2026

Converts PowerPoint files into reveal.js web presentations. Text, formatting, images,
shapes, master layouts and theme colors are preserved.

## What you need

Node.js 18 LTS. Check with `node --version`. If you don't have it, get it from nodejs.org.
Newer Node versions print dependency warnings during install but the app still runs.

## Getting it running

```
git clone https://github.com/Aruzhankalka/PPTX-to-reveal.js-Converter.git
cd PPTX-to-reveal.js-Converter
npm install
npm run dev
```

Open http://localhost:3000 and upload a .pptx file.

If you don't have one at hand, use `tests/fixtures/sample.pptx`.

There is also a setup script that does the same thing and checks your Node version first:
`setup.sh` on Linux and macOS, `setup.bat` on Windows.

## Tests

```
npm test
```

That runs the 390 Jest tests. The browser tests are separate and need the server running:

```
npm run dev             (one terminal)
npx playwright test     (another terminal)
```

The first time you run Playwright you also need `npx playwright install chromium firefox`.

## How it works

A PPTX file is parsed into an intermediate JSON format (we call it the IR), and the
generator turns that IR into reveal.js HTML. The same IR can be fed to the PPTX
generator to go back the other way. Uploads go through a sanitizer first, which strips
VBA macros, scripts inside SVG files and HTML imports.

```
src/
  api/            upload, preview and download endpoints
  parser/pptx/    reads .pptx, produces the IR
  ir/             JSON schema and validator for the IR
  generator/
    revealjs/     IR to reveal.js HTML
    pptx/         IR to .pptx
  security/       the sanitizer
  web/            frontend
tests/            Jest
tests/e2e/        Playwright
```

## If something breaks

`npm: command not found` means Node.js isn't in your PATH. Restart the terminal, or
reboot on Windows.

Images missing when you open a downloaded HTML file: use the "Open in browser" button
instead while the server is running. The images are served by the backend.

Playwright complaining about `@playwright/test`: run `npm install` again, then
`npx playwright install chromium firefox`.
