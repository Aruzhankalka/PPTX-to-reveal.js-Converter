const fs = require("fs");
const path = require("path");
const request = require("supertest");
const JSZip = require("jszip");
const app = require("../../src/app");

const FIXTURE_PATH = path.join(
  __dirname,
  "..",
  "fixtures",
  "sample-with-image.pptx"
);

describe("US-05 AC-3: GET /api/v1/result/:id returns an offline-usable ZIP", () => {
  let resultId;

  beforeAll(async () => {
    if (!fs.existsSync(FIXTURE_PATH)) {
      throw new Error(
        `Missing fixture at ${FIXTURE_PATH}. ` +
          `Create a small PPTX with at least one image and commit it.`
      );
    }

    const res = await request(app)
      .post("/api/v1/convert")
      .attach("file", FIXTURE_PATH)
      .expect(200);

    expect(res.body.result_id).toBeDefined();
    resultId = res.body.result_id;
  });

  test("responds with application/zip and a .zip filename", async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`)
      .expect(200)
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });

    expect(res.headers["content-type"]).toMatch(/application\/zip/);
    expect(res.headers["content-disposition"]).toMatch(/filename=".*\.zip"/);
    expect(res.body.length).toBeGreaterThan(0);
  });

  test("ZIP contains index.html with rewritten asset paths", async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`)
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });

    const zip = await JSZip.loadAsync(res.body);

    const indexEntry = zip.file("index.html");
    expect(indexEntry).not.toBeNull();

    const indexHtml = await indexEntry.async("string");

    expect(indexHtml).not.toMatch(/\/api\/v1\/media\//);
    expect(indexHtml).toMatch(/src="assets\/[^"]+"/);
    expect(indexHtml).not.toMatch(/cdn\.jsdelivr\.net/);
    expect(indexHtml).toMatch(/reveal\/reveal\.js/);
    expect(indexHtml).toMatch(/reveal\/reveal\.css/);
  });

  test("ZIP contains every media file referenced by index.html", async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`)
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });

    const zip = await JSZip.loadAsync(res.body);
    const indexHtml = await zip.file("index.html").async("string");

    const matches = [...indexHtml.matchAll(/src="(assets\/[^"]+)"/g)];
    expect(matches.length).toBeGreaterThan(0);

    for (const m of matches) {
      const assetPath = m[1];
      const entry = zip.file(assetPath);
      expect(entry).not.toBeNull();
      const bytes = await entry.async("nodebuffer");
      expect(bytes.length).toBeGreaterThan(0);
    }
  });

  test("ZIP contains bundled reveal.js distribution files", async () => {
    const res = await request(app)
      .get(`/api/v1/result/${resultId}`)
      .buffer(true)
      .parse((res, cb) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });

    const zip = await JSZip.loadAsync(res.body);

    for (const required of [
      "reveal/reveal.js",
      "reveal/reveal.css",
      "reveal/reset.css",
      "reveal/theme/white.css",
    ]) {
      expect(zip.file(required)).not.toBeNull();
    }
  });

  test("returns 404 with a structured error for an unknown result ID", async () => {
    const res = await request(app)
      .get("/api/v1/result/does-not-exist")
      .expect(404);

    expect(res.body.error_code).toBe("RESULT_NOT_FOUND");
    expect(res.body.message).toEqual(expect.any(String));
  });
});
