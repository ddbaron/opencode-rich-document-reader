import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, relative } from "node:path";
import type { ToolAttachment } from "@opencode-ai/plugin";
import type { OfficeParserAST } from "officeparser";
import { formatForExtension, formatRegistry } from "./registry.ts";
import { READER_LIMITS, parserLimits } from "./limits.ts";
import {
  MediaSelectionError,
  mediaTable,
  selectedAttachments,
  selectMedia,
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

  let input: Buffer;
  try {
    input = await readFile(resolved.absolutePath);
  } catch (error) {
    throw new RichDocumentError(
      "UNREADABLE_DOCUMENT",
      `Document is not readable: ${args.path} (${errorMessage(error)})`,
      { cause: error },
    );
  }

  const issues: ParseIssue[] = [];
  const ast = await parseDocument(format, input, context, args.path, issues);

  let temporaryDirectory: string | undefined;
  let keepTemporaryDirectory = false;
  try {
    const extraction = await extractMedia(ast, dependencies.tempDirectory ?? tmpdir(), args.path);
    temporaryDirectory = extraction.temporaryDirectory;
    const selected = selectRequestedMedia(extraction.records, args.media);
    const conversion = await convertMarkdown(ast, args.path, issues);

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
