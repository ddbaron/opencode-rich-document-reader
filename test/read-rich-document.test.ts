import assert from "node:assert/strict";
import { after, before, describe, it } from "node:test";
import { readFile, rm, stat, symlink, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import type { OfficeParserAST } from "officeparser";
import { RichDocumentReaderPlugin } from "../src/index.ts";
import { READER_LIMITS } from "../src/limits.ts";
import { readRichDocument } from "../src/reader.ts";
import { createFixtures } from "./fixtures.ts";

let fixtures: Awaited<ReturnType<typeof createFixtures>>;
const extractedDirectories = new Set<string>();

function context() {
  const abort = new AbortController();
  return { directory: fixtures.root, worktree: fixtures.root, abort: abort.signal };
}

function rememberExtraction(result: { metadata?: { media?: Array<{ temporaryPath: string }> } }) {
  for (const item of result.metadata?.media ?? []) extractedDirectories.add(dirname(item.temporaryPath));
}

before(async () => {
  fixtures = await createFixtures();
});

after(async () => {
  await rm(fixtures.root, { recursive: true, force: true });
  await Promise.all(
    [...extractedDirectories].map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

describe("plugin registration", () => {
  it("registers exactly the read_rich_document agent tool", async () => {
    const hooks = await RichDocumentReaderPlugin({} as never);
    assert.deepEqual(Object.keys(hooks.tool ?? {}), ["read_rich_document"]);
    assert.equal(typeof hooks.tool?.read_rich_document.execute, "function");
  });
});

describe("read_rich_document", () => {
  it("extracts DOCX structure, media, and section association without default attachments", async () => {
    const result = await readRichDocument({ path: "structure.docx" }, context());
    rememberExtraction(result);

    assert.match(result.output, /# Project Overview/);
    assert.match(result.output, /First bullet/);
    assert.match(result.output, /\| Header \| Value \|/);
    assert.match(result.output, /\[Open reference\]\(https:\/\/example\.com\/reference\)/);
    assert.match(result.output, /`media-1`/);
    assert.match(result.output, /image1\.png/);
    assert.match(result.output, /image\/png/);
    assert.match(result.output, /## Section 1/);
    assert.match(result.output, /## Section 2/);
    assert.match(result.output, /Section: Project Overview/);
    assert.equal(result.attachments, undefined);
    assert.ok(result.metadata);
    assert.equal(result.metadata.media.length, 1);
    assert.equal(await stat(result.metadata.media[0].temporaryPath).then(() => true), true);
    assert.doesNotMatch(result.output, /iVBORw0KGgo/);
  });

  it("associates media after a DOCX section boundary with the physical section", async () => {
    const ast = {
      config: {},
      type: "docx",
      metadata: {},
      content: [
        {
          type: "heading",
          text: "First section",
          children: [{ type: "text", text: "First section" }],
          metadata: { level: 1 },
        },
        {
          type: "paragraph",
          rawContent: "<w:p><w:pPr><w:sectPr/></w:pPr></w:p>",
          children: [],
        },
        {
          type: "image",
          metadata: { attachmentName: "after-section.png" },
          children: [],
        },
      ],
      attachments: [
        {
          type: "image",
          name: "after-section.png",
          extension: "png",
          mimeType: "image/png",
          data: "iVBORw0KGgo=",
        },
      ],
      warnings: [],
      to: async () => ({ value: "after section", messages: [] }),
      toText: () => "",
    } as unknown as OfficeParserAST;
    const formats = new Map([
      [
        ".docx",
        {
          extension: ".docx",
          parserType: "docx" as const,
          parse: async () => ast,
        },
      ],
    ]);
    await writeFile(join(fixtures.root, "sectioned.docx"), Buffer.from("fixture"));

    const result = await readRichDocument(
      { path: "sectioned.docx" },
      context(),
      { formats },
    );
    rememberExtraction(result);

    assert.equal(result.metadata?.media[0].location, "Section 2 - Section: First section");
  });

  it("labels media from headers, footers, and slide masters", async () => {
    const node = (attachmentName: string, sectionNumber?: number) => ({
      type: "image",
      metadata: { attachmentName, ...(sectionNumber === undefined ? {} : { sectionNumber }) },
      children: [],
    });
    const ast = {
      config: {},
      type: "pptx",
      metadata: {},
      content: [],
      auxiliary: {
        headers: [node("header.png", 2)],
        footers: [node("footer.png")],
        slideMasters: [node("master.png")],
      },
      attachments: ["header.png", "footer.png", "master.png"].map((name) => ({
        type: "image" as const,
        name,
        extension: "png",
        mimeType: "image/png",
        data: "iVBORw0KGgo=",
      })),
      warnings: [],
      to: async () => ({ value: "auxiliary media", messages: [] }),
      toText: () => "",
    } as unknown as OfficeParserAST;
    const formats = new Map([
      [
        ".pptx",
        {
          extension: ".pptx",
          parserType: "pptx" as const,
          parse: async () => ast,
        },
      ],
    ]);
    await writeFile(join(fixtures.root, "auxiliary.pptx"), Buffer.from("fixture"));

    const result = await readRichDocument(
      { path: "auxiliary.pptx" },
      context(),
      { formats },
    );
    rememberExtraction(result);

    assert.deepEqual(
      result.metadata?.media.map((item) => item.location),
      ["Header - Section 2", "Footer", "Slide master"],
    );
  });

  it("extracts ODT structure and associates media with the nearest section", async () => {
    const result = await readRichDocument({ path: "structure.odt" }, context());
    rememberExtraction(result);

    assert.match(result.output, /# ODT Section/);
    assert.match(result.output, /ODT bullet/);
    assert.match(result.output, /ODT Header/);
    assert.match(result.output, /`media-1`/);
    assert.match(result.output, /Section: ODT Section/);
    assert.ok(result.metadata);
    assert.equal(result.metadata.format, "odt");
  });

  it("retains PPTX slide and speaker-note context while indexing its image", async () => {
    const result = await readRichDocument({ path: "slides.pptx" }, context());
    rememberExtraction(result);

    assert.match(result.output, /PPTX Slide One/);
    assert.match(result.output, /Slide bullet/);
    assert.match(result.output, /Speaker notes/);
    assert.match(result.output.split("## Embedded media", 1)[0], /## Slide 1/);
    assert.ok(result.metadata);
    assert.equal(result.metadata.format, "pptx");
    assert.match(result.metadata.media[0].location, /^Slide 1/);
    assert.equal(result.metadata.media.find((item) => item.type === "chart")?.mimeType, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
  });

  it("attaches only explicitly selected image media as a native file attachment", async () => {
    const result = await readRichDocument({ path: "structure.docx", media: ["media-1"] }, context());
    rememberExtraction(result);

    assert.equal(result.attachments?.length, 1);
    assert.equal(result.attachments?.[0].type, "file");
    assert.equal(result.attachments?.[0].mime, "image/png");
    assert.match(result.attachments?.[0].url ?? "", /^data:image\/png;base64,/);
    assert.match(result.attachments?.[0].filename ?? "", /^media-1\.png$/);
    assert.doesNotMatch(result.output, /iVBORw0KGgo/);
  });

  it("leaves the source document byte-for-byte unchanged", async () => {
    const beforeBytes = await readFile(fixtures.docx);
    const beforeStat = await stat(fixtures.docx);
    const result = await readRichDocument({ path: "structure.docx" }, context());
    rememberExtraction(result);
    const afterBytes = await readFile(fixtures.docx);
    const afterStat = await stat(fixtures.docx);

    assert.deepEqual(afterBytes, beforeBytes);
    assert.equal(afterStat.mtimeMs, beforeStat.mtimeMs);
  });

  it("rejects lexical path escapes and symlinks that resolve outside the project", async () => {
    const outside = join(dirname(fixtures.root), "outside.docx");
    const link = join(fixtures.root, "escaped.docx");
    await writeFile(outside, Buffer.from("outside"));
    await symlink(outside, link);

    await assert.rejects(
      () => readRichDocument({ path: "../outside.docx" }, context()),
      /escapes the current project/i,
    );
    await assert.rejects(
      () => readRichDocument({ path: "escaped.docx" }, context()),
      /symlink resolves outside/i,
    );

    await rm(link, { force: true });
    await rm(outside, { force: true });
  });

  it("rejects unsupported, missing, and malformed documents with useful errors", async () => {
    await writeFile(join(fixtures.root, "notes.pdf"), Buffer.from("not supported"));
    await writeFile(join(fixtures.root, "broken.docx"), Buffer.from("not a zip archive"));

    await assert.rejects(
      () => readRichDocument({ path: "notes.pdf" }, context()),
      /unsupported document extension/i,
    );
    await assert.rejects(
      () => readRichDocument({ path: "missing.docx" }, context()),
      /document not found/i,
    );
    await assert.rejects(
      () => readRichDocument({ path: "broken.docx" }, context()),
      /could not parse document|zip|corrupt|malformed/i,
    );
  });

  it("rejects an unknown media selector instead of attaching every image", async () => {
    await assert.rejects(
      () => readRichDocument({ path: "structure.docx", media: ["media-99"] }, context()),
      /unknown media selector/i,
    );
  });

  it("rejects image-labeled attachments without an image MIME type", async () => {
    const ast = {
      config: {},
      type: "docx",
      metadata: {},
      content: [{ type: "image", metadata: { attachmentName: "binary.bin" }, children: [] }],
      attachments: [
        {
          type: "image" as const,
          name: "binary.bin",
          extension: "bin",
          mimeType: "application/octet-stream",
          data: "AA==",
        },
      ],
      warnings: [],
      to: async () => ({ value: "binary", messages: [] }),
      toText: () => "",
    } as unknown as OfficeParserAST;
    const formats = new Map([
      [
        ".docx",
        {
          extension: ".docx",
          parserType: "docx" as const,
          parse: async () => ast,
        },
      ],
    ]);
    await writeFile(join(fixtures.root, "binary.docx"), Buffer.from("fixture"));

    await assert.rejects(
      () => readRichDocument({ path: "binary.docx", media: ["media-1"] }, context(), { formats }),
      /supported image/i,
    );
  });

  it("rejects parser-truncated tables before Markdown conversion", async () => {
    let converted = false;
    const warning = {
      type: "warning" as const,
      code: "TABLE_CELL_LIMIT_EXCEEDED",
      message: "Table cell materialization reached its limit.",
    };
    const ast = {
      config: {},
      type: "odt",
      metadata: {},
      content: [],
      attachments: [],
      warnings: [warning],
      to: async () => {
        converted = true;
        return { value: "should not convert", messages: [] };
      },
      toText: () => "",
    } as unknown as OfficeParserAST;
    const formats = new Map([
      [
        ".odt",
        {
          extension: ".odt",
          parserType: "odt" as const,
          parse: async () => ast,
        },
      ],
    ]);
    await writeFile(join(fixtures.root, "truncated.odt"), Buffer.from("fixture"));

    await assert.rejects(
      () => readRichDocument({ path: "truncated.odt" }, context(), { formats }),
      /table cell limit/i,
    );
    assert.equal(converted, false);
  });

  it("rejects oversized tables before conversion", async () => {
    const cell = { type: "cell" } as const;
    const cells = Array.from({ length: READER_LIMITS.maxTableCells + 1 }, () => cell);
    const ast = {
      config: {},
      type: "docx",
      metadata: {},
      content: [{ type: "table", children: [{ type: "row", children: cells }] }],
      attachments: [],
      warnings: [],
      to: async () => ({ value: "should not convert", messages: [] }),
      toText: () => "",
    } as unknown as OfficeParserAST;
    const formats = new Map([
      [
        ".docx",
        {
          extension: ".docx",
          parserType: "docx" as const,
          parse: async () => ast,
        },
      ],
    ]);
    await writeFile(join(fixtures.root, "oversized.docx"), Buffer.from("fixture"));

    await assert.rejects(
      () => readRichDocument({ path: "oversized.docx" }, context(), { formats }),
      /too many table cells/i,
    );
  });
});
