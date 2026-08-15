import { constants } from "node:fs";
import { mkdtemp, open, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join, relative } from "node:path";
import type { ToolAttachment } from "@opencode-ai/plugin";
import type { OfficeContentNode, OfficeParserAST } from "officeparser";
import { formatForExtension, formatRegistry } from "./registry.ts";
import { READER_LIMITS, parserLimits } from "./limits.ts";
import {
  MediaSelectionError,
  type MediaTableOptions,
  mediaTable,
  selectedAttachments,
  selectMedia,
  sectionBoundariesFor,
  writeMedia,
} from "./media.ts";
import { DurableExportError, writeDurableExport } from "./export.ts";
import { DocumentPathError, projectPaths, resolveDocumentPath, resolveExportDestination } from "./path-safety.ts";
import type {
  DurableExportMetadata,
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

function completeMarkdown(
  markdown: string,
  records: MediaRecord[],
  issues: DocumentIssue[],
  mediaOptions?: MediaTableOptions,
): string {
  const body = stripInlineMediaData(markdown.trim()) || "_No readable text content was found._";
  return [body, mediaTable(records, mediaOptions), warningSection(issues)].join("\n\n");
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return String(error);
}

function isAbortError(error: unknown): error is Error {
  return error instanceof Error && error.name === "AbortError";
}

function parserError(error: unknown, sourcePath: string): Error {
  if (error instanceof RichDocumentError) return error;
  if (error instanceof DocumentPathError) return error;
  if (isAbortError(error)) return error;

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

async function readValidatedDocument(
  absolutePath: string,
  documentInput: string,
  abortSignal: AbortSignal,
): Promise<Buffer> {
  try {
    abortSignal.throwIfAborted();
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
      const chunks: Buffer[] = [];
      let totalBytes = 0;
      while (totalBytes <= READER_LIMITS.maxArchiveBytes) {
        abortSignal.throwIfAborted();
        const chunkSize = Math.min(
          1024 * 1024,
          READER_LIMITS.maxArchiveBytes + 1 - totalBytes,
        );
        const chunk = Buffer.allocUnsafe(chunkSize);
        const { bytesRead } = await documentHandle.read(chunk, 0, chunkSize, totalBytes);
        if (!bytesRead) break;
        chunks.push(chunk.subarray(0, bytesRead));
        totalBytes += bytesRead;
        if (totalBytes > READER_LIMITS.maxArchiveBytes) {
          throw new RichDocumentError(
            "ARCHIVE_SIZE_LIMIT",
            `Document archive is too large (${totalBytes} bytes). The limit is ${READER_LIMITS.maxArchiveBytes} bytes.`,
          );
        }
      }
      abortSignal.throwIfAborted();
      return Buffer.concat(chunks, totalBytes);
    } finally {
      await documentHandle.close().catch(() => undefined);
    }
  } catch (error) {
    if (abortSignal.aborted) abortSignal.throwIfAborted();
    if (isAbortError(error)) throw error;
    if (error instanceof RichDocumentError) throw error;
    throw new RichDocumentError(
      "UNREADABLE_DOCUMENT",
      `Document is not readable: ${documentInput} (${errorMessage(error)})`,
      { cause: error },
    );
  }
}

function enforceTableCellLimit(ast: OfficeParserAST, sourcePath: string, issues: ParseIssue[]): void {
  const tableLimitWarning = [...issues.map(({ issue }) => issue), ...ast.warnings].find(
    (issue) => String(issue.code) === "TABLE_CELL_LIMIT_EXCEEDED",
  );
  if (tableLimitWarning) {
    throw new RichDocumentError(
      "TABLE_CELL_LIMIT",
      `Document ${sourcePath} exceeded the table cell limit during parsing. The limit is ${READER_LIMITS.maxTableCells}.`,
      { cause: tableLimitWarning },
    );
  }

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
    context.abort.throwIfAborted();
    const ast = await format.parse(input, {
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
    context.abort.throwIfAborted();
    return ast;
  } catch (error) {
    if (context.abort.aborted) context.abort.throwIfAborted();
    throw parserError(error, sourcePath);
  }
}

async function extractMedia(
  ast: OfficeParserAST,
  baseTemporaryDirectory: string,
  sourcePath: string,
  abortSignal: AbortSignal,
): Promise<{ temporaryDirectory?: string; records: MediaRecord[] }> {
  let temporaryDirectory: string | undefined;
  try {
    abortSignal.throwIfAborted();
    if (ast.attachments.length) {
      temporaryDirectory = await mkdtemp(join(baseTemporaryDirectory, "opencode-rich-document-"));
    }
    const records = temporaryDirectory ? await writeMedia(ast, temporaryDirectory, abortSignal) : [];
    abortSignal.throwIfAborted();
    return { temporaryDirectory, records };
  } catch (error) {
    if (temporaryDirectory) {
      await rm(temporaryDirectory, { recursive: true, force: true }).catch(() => undefined);
    }
    if (abortSignal.aborted) abortSignal.throwIfAborted();
    if (isAbortError(error)) throw error;
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
  abortSignal: AbortSignal,
) {
  try {
    abortSignal.throwIfAborted();
    const conversion = await ast.to("md", {
      includeImages: false,
      includeCharts: false,
      abortSignal,
      onWarning: (issue) => issues.push({ issue, source: "conversion" }),
      mdConfig: {
        dialect: "github",
        fallbackToHtml: true,
      },
    });
    abortSignal.throwIfAborted();
    return conversion;
  } catch (error) {
    if (abortSignal.aborted) abortSignal.throwIfAborted();
    if (isAbortError(error)) throw error;
    throw conversionError(error, sourcePath);
  }
}

export async function readRichDocument(
  args: ReadRichDocumentArgs,
  context: ReaderToolContext,
  dependencies: ReadRichDocumentDependencies = {},
): Promise<RichDocumentToolResult> {
  context.abort.throwIfAborted();
  const projectRoot = projectRootFor(context);
  const paths = projectPaths(context, args.path);
  const resolved = await resolveDocumentPath(paths.documentPath, paths.projectRoot);
  context.abort.throwIfAborted();
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

  const exportDestination = args.export
    ? await resolveExportDestination(args.export.destination, paths.projectRoot, resolved.absolutePath)
    : undefined;

  const input = await readValidatedDocument(resolved.absolutePath, args.path, context.abort);

  const issues: ParseIssue[] = [];
  const ast = await parseDocument(format, input, context, args.path, issues);
  context.abort.throwIfAborted();
  enforceTableCellLimit(ast, args.path, issues);

  let temporaryDirectory: string | undefined;
  let keepTemporaryDirectory = false;
  try {
    const extraction = await extractMedia(
      ast,
      dependencies.tempDirectory ?? tmpdir(),
      args.path,
      context.abort,
    );
    temporaryDirectory = extraction.temporaryDirectory;
    context.abort.throwIfAborted();
    const selected = selectRequestedMedia(extraction.records, args.media);
    context.abort.throwIfAborted();
    const contextualAst = addDocumentContext(ast);
    context.abort.throwIfAborted();
    const conversion = await convertMarkdown(contextualAst, args.path, issues, context.abort);
    context.abort.throwIfAborted();

    for (const issue of ast.warnings) issues.push({ issue, source: "parser" });
    for (const issue of conversion.messages) issues.push({ issue, source: "conversion" });

    const media = attachmentMetadata(extraction.records);
    const issuesForMarkdown = uniqueIssues(issues);
    const ephemeralMarkdown = completeMarkdown(String(conversion.value), extraction.records, issuesForMarkdown);
    let output = ephemeralMarkdown;
    let exportMetadata: DurableExportMetadata | undefined;
    if (exportDestination) {
      const exportedMarkdown = completeMarkdown(String(conversion.value), extraction.records, issuesForMarkdown, {
        pathHeading: "Exported path",
        pathFor: (entry) => `media/${basename(entry.temporaryPath)}`,
      });
      try {
        exportMetadata = await writeDurableExport({
          destination: exportDestination,
          sourcePath: sourceLabel(args.path, projectRoot, resolved.absolutePath),
          format: format.parserType,
          markdown: exportedMarkdown,
          records: extraction.records,
          abortSignal: context.abort,
        });
      } catch (error) {
        if (error instanceof DurableExportError) {
          throw new RichDocumentError(error.code, error.message, { cause: error });
        }
        throw error;
      }
      output = exportedMarkdown;
    }
    const metadata: RichDocumentResultMetadata = {
      format: format.parserType,
      sourcePath: sourceLabel(args.path, projectRoot, resolved.absolutePath),
      media,
      ...(exportMetadata ? { export: exportMetadata } : {}),
    };
    const attachments = toolAttachments(selected);
    const result: RichDocumentToolResult = {
      title: `Read ${sourceLabel(args.path, projectRoot, resolved.absolutePath)}`,
      output,
      metadata,
    };
    if (attachments.length) result.attachments = attachments;
    context.abort.throwIfAborted();
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
