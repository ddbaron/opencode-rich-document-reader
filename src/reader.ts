import { constants } from "node:fs";
import { mkdtemp, open, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, relative } from "node:path";
import type { ToolAttachment } from "@opencode-ai/plugin";
import type { OfficeContentNode, OfficeParserAST } from "officeparser";
import { formatForExtension, formatRegistry } from "./registry.ts";
import { READER_LIMITS, parserLimits } from "./limits.ts";
import {
  MediaSelectionError,
  mediaTable,
  selectedAttachments,
  selectMedia,
  sectionBoundariesFor,
  writeMedia,
} from "./media.ts";
import { DocumentPathError, projectPaths, resolveDocumentPath } from "./path-safety.ts";
import type {
  MediaIndexEntry,
  MediaRecord,
  ReadRichDocumentArgs,
  ReaderToolContext,
  RichDocumentResultMetadata,
  RichDocumentToolResult,
  ReadRichDocumentDependencies,
  ParseIssue,
  DocumentIssue,
} from "./types.ts";

export class RichDocumentError extends Error {
  readonly code: string;

  constructor(code: string, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "RichDocumentError";
    this.code = code;
  }
}

function issueKey(issue: DocumentIssue): string {
  return `${issue.type}:${String(issue.code)}:${issue.message}`;
}

function uniqueIssues(issues: ParseIssue[]): DocumentIssue[] {
  const seen = new Set<string>();
  const result: DocumentIssue[] = [];
  for (const { issue } of issues) {
    const key = issueKey(issue);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(issue);
  }
  return result;
}

function warningSection(issues: DocumentIssue[]): string {
  const rows = issues.length
    ? issues.map(
        (issue) =>
          `- **${issue.type} ${String(issue.code)}** ${issue.message.replace(/[\r\n]+/g, " ")}`,
      )
    : ["- None reported."];
  return ["## Conversion warnings", "", ...rows].join("\n");
}

function stripInlineMediaData(markdown: string): string {
  return markdown.replace(
    /data:[^\s,;]+(?:;[^\s,;]+)*;base64,[A-Za-z0-9+/=_-]+/g,
    "[embedded media omitted; see the media index]",
  );
}

function completeMarkdown(markdown: string, records: MediaRecord[], issues: DocumentIssue[]): string {
  const body = stripInlineMediaData(markdown.trim()) || "_No readable text content was found._";
  return [body, mediaTable(records), warningSection(issues)].join("\n\n");
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return String(error);
}

function parserError(error: unknown, sourcePath: string): Error {
  if (error instanceof RichDocumentError) return error;
  if (error instanceof DocumentPathError) return error;

  const issue =
    error && typeof error === "object" && "officeIssue" in error && error.officeIssue
      ? error.officeIssue
      : undefined;
  const detail =
    issue && typeof issue === "object" && "message" in issue ? String(issue.message) : errorMessage(error);
  const code =
    issue && typeof issue === "object" && "code" in issue ? String(issue.code) : "PARSE_FAILED";
  return new RichDocumentError(code, `Could not parse document ${sourcePath}: ${detail}`, { cause: error });
}

function extractionError(error: unknown, sourcePath: string): RichDocumentError {
  if (error instanceof RichDocumentError) return error;
  return new RichDocumentError(
    "MEDIA_EXTRACTION_FAILED",
    `Could not extract document media from ${sourcePath}: ${errorMessage(error)}`,
    { cause: error },
  );
}

function conversionError(error: unknown, sourcePath: string): RichDocumentError {
  if (error instanceof RichDocumentError) return error;
  return new RichDocumentError(
    "CONVERSION_FAILED",
    `Could not convert document ${sourcePath} to Markdown: ${errorMessage(error)}`,
    { cause: error },
  );
}

function projectRootFor(context: ReaderToolContext): string {
  return context.worktree || context.directory;
}

function sourceLabel(documentInput: string, projectRoot: string, absolutePath: string): string {
  const relativePath = relative(projectRoot, absolutePath);
  return relativePath && !relativePath.startsWith("..") ? relativePath : documentInput;
}

function attachmentMetadata(records: MediaRecord[]): MediaIndexEntry[] {
  return records.map(({ entry }) => ({ ...entry }));
}

function toolAttachments(records: MediaRecord[]): ToolAttachment[] {
  return selectedAttachments(records);
}

async function readValidatedDocument(absolutePath: string, documentInput: string): Promise<Buffer> {
  try {
    const documentHandle = await open(absolutePath, constants.O_RDONLY | constants.O_NOFOLLOW);
    try {
      const documentStat = await documentHandle.stat();
      if (!documentStat.isFile()) {
        throw new RichDocumentError("NOT_A_FILE", `Document path is not a regular file: ${documentInput}`);
      }
      if (documentStat.size > READER_LIMITS.maxArchiveBytes) {
        throw new RichDocumentError(
          "ARCHIVE_SIZE_LIMIT",
          `Document archive is too large (${documentStat.size} bytes). The limit is ${READER_LIMITS.maxArchiveBytes} bytes.`,
        );
      }
      const input = await readFile(documentHandle);
      if (input.byteLength > READER_LIMITS.maxArchiveBytes) {
        throw new RichDocumentError(
          "ARCHIVE_SIZE_LIMIT",
          `Document archive is too large (${input.byteLength} bytes). The limit is ${READER_LIMITS.maxArchiveBytes} bytes.`,
        );
      }
      return input;
    } finally {
      await documentHandle.close().catch(() => undefined);
    }
  } catch (error) {
    if (error instanceof RichDocumentError) throw error;
    throw new RichDocumentError(
      "UNREADABLE_DOCUMENT",
      `Document is not readable: ${documentInput} (${errorMessage(error)})`,
      { cause: error },
    );
  }
}

function enforceTableCellLimit(ast: OfficeParserAST, sourcePath: string): void {
  const pending: OfficeContentNode[] = [...ast.content];
  for (const nodes of [ast.auxiliary?.headers, ast.auxiliary?.footers, ast.auxiliary?.slideMasters]) {
    if (nodes) {
      for (const node of nodes) pending.push(node);
    }
  }

  let cellCount = 0;
  while (pending.length) {
    const node = pending.pop();
    if (!node) continue;
    if (node.type === "cell") {
      cellCount += 1;
      if (cellCount > READER_LIMITS.maxTableCells) {
        throw new RichDocumentError(
          "TABLE_CELL_LIMIT",
          `Document ${sourcePath} contains too many table cells (${cellCount}). The limit is ${READER_LIMITS.maxTableCells}.`,
        );
      }
    }
    for (const nodes of [node.children, node.notes, node.comments]) {
      if (nodes) {
        for (const child of nodes) pending.push(child);
      }
    }
  }
}

function contextHeading(text: string): OfficeContentNode {
  return {
    type: "heading",
    text,
    children: [{ type: "text", text }],
    metadata: { level: 2 },
  };
}

function addDocumentContext(ast: OfficeParserAST): OfficeParserAST {
  const sectionBoundaries = sectionBoundariesFor(ast) ?? [];
  const hasMultipleSections = sectionBoundaries.some(
    (boundary, index) => boundary && index < ast.content.length - 1,
  );
  const content: OfficeContentNode[] = [];
  let sectionNumber = 1;
  let slideIndex = 0;

  if (hasMultipleSections) content.push(contextHeading(`Section ${sectionNumber}`));
  for (const [index, node] of ast.content.entries()) {
    if (node.type === "slide") {
      slideIndex += 1;
      content.push(contextHeading(`Slide ${node.metadata?.slideNumber ?? slideIndex}`));
    }
    content.push(node);
    if (hasMultipleSections && sectionBoundaries[index] && index < ast.content.length - 1) {
      sectionNumber += 1;
      content.push(contextHeading(`Section ${sectionNumber}`));
    }
  }

  return content.length === ast.content.length ? ast : { ...ast, content };
}

async function parseDocument(
  format: NonNullable<ReturnType<typeof formatForExtension>>,
  input: Buffer,
  context: ReaderToolContext,
  sourcePath: string,
  issues: ParseIssue[],
): Promise<OfficeParserAST> {
  try {
    return await format.parse(input, {
      ...parserLimits(),
      extractAttachments: true,
      ignoreComments: false,
      ignoreHeadersAndFooters: false,
      ignoreNotes: false,
      ignoreSlideMasters: false,
      ocr: false,
      onWarning: (issue) => issues.push({ issue, source: "parser" }),
      abortSignal: context.abort,
      includeRawContent: format.parserType === "docx",
    });
  } catch (error) {
    throw parserError(error, sourcePath);
  }
}

async function extractMedia(
  ast: OfficeParserAST,
  baseTemporaryDirectory: string,
  sourcePath: string,
): Promise<{ temporaryDirectory?: string; records: MediaRecord[] }> {
  let temporaryDirectory: string | undefined;
  try {
    if (ast.attachments.length) {
      temporaryDirectory = await mkdtemp(join(baseTemporaryDirectory, "opencode-rich-document-"));
    }
    const records = temporaryDirectory ? await writeMedia(ast, temporaryDirectory) : [];
    return { temporaryDirectory, records };
  } catch (error) {
    throw extractionError(error, sourcePath);
  }
}

function selectRequestedMedia(records: MediaRecord[], selectors: string[] | undefined): MediaRecord[] {
  try {
    return selectMedia(records, selectors);
  } catch (error) {
    if (error instanceof MediaSelectionError) {
      throw new RichDocumentError(error.code, error.message, { cause: error });
    }
    throw error;
  }
}

async function convertMarkdown(
  ast: OfficeParserAST,
  sourcePath: string,
  issues: ParseIssue[],
) {
  try {
    return await ast.to("md", {
      includeImages: false,
      includeCharts: false,
      onWarning: (issue) => issues.push({ issue, source: "conversion" }),
      mdConfig: {
        dialect: "github",
        fallbackToHtml: true,
      },
    });
  } catch (error) {
    throw conversionError(error, sourcePath);
  }
}

export async function readRichDocument(
  args: ReadRichDocumentArgs,
  context: ReaderToolContext,
  dependencies: ReadRichDocumentDependencies = {},
): Promise<RichDocumentToolResult> {
  const projectRoot = projectRootFor(context);
  const paths = projectPaths(context, args.path);
  const resolved = await resolveDocumentPath(paths.documentPath, paths.projectRoot);
  const format = formatForExtension(resolved.extension, dependencies.formats ?? formatRegistry);

  if (!format) {
    throw new RichDocumentError(
      "UNSUPPORTED_FORMAT",
      `Unsupported document extension "${resolved.extension || "(none)"}". Supported formats are .docx, .odt, and .pptx.`,
    );
  }
  if (resolved.size > READER_LIMITS.maxArchiveBytes) {
    throw new RichDocumentError(
      "ARCHIVE_SIZE_LIMIT",
      `Document archive is too large (${resolved.size} bytes). The limit is ${READER_LIMITS.maxArchiveBytes} bytes.`,
    );
  }

  const input = await readValidatedDocument(resolved.absolutePath, args.path);

  const issues: ParseIssue[] = [];
  const ast = await parseDocument(format, input, context, args.path, issues);
  enforceTableCellLimit(ast, args.path);

  let temporaryDirectory: string | undefined;
  let keepTemporaryDirectory = false;
  try {
    const extraction = await extractMedia(ast, dependencies.tempDirectory ?? tmpdir(), args.path);
    temporaryDirectory = extraction.temporaryDirectory;
    const selected = selectRequestedMedia(extraction.records, args.media);
    const conversion = await convertMarkdown(addDocumentContext(ast), args.path, issues);

    for (const issue of ast.warnings) issues.push({ issue, source: "parser" });
    for (const issue of conversion.messages) issues.push({ issue, source: "conversion" });

    const media = attachmentMetadata(extraction.records);
    const metadata: RichDocumentResultMetadata = {
      format: format.parserType,
      sourcePath: sourceLabel(args.path, projectRoot, resolved.absolutePath),
      media,
    };
    const attachments = toolAttachments(selected);
    const result: RichDocumentToolResult = {
      title: `Read ${sourceLabel(args.path, projectRoot, resolved.absolutePath)}`,
      output: completeMarkdown(String(conversion.value), extraction.records, uniqueIssues(issues)),
      metadata,
    };
    if (attachments.length) result.attachments = attachments;
    keepTemporaryDirectory = true;
    return result;
  } finally {
    // Successful reads intentionally keep the temporary directory available to
    // the agent. On failure, avoid leaving a partial extraction behind.
    if (temporaryDirectory && !keepTemporaryDirectory) {
      await rm(temporaryDirectory, { recursive: true, force: true }).catch(() => undefined);
    }
  }
}
