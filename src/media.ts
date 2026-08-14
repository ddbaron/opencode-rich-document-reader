import { mkdir, writeFile } from "node:fs/promises";
import { basename, extname, join } from "node:path";
import type { OfficeAttachment, OfficeContentNode, OfficeParserAST } from "officeparser";
import type { MediaIndexEntry, MediaRecord } from "./types.ts";

interface LocationContext {
  heading?: string;
  role?: string;
  sectionNumber?: number;
  slideNumber?: number;
  inNotes: boolean;
}

const OOXML_DOCUMENT_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const OOXML_CHART_MIME = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const SUPPORTED_IMAGE_MIME_TYPES = new Set([
  "image/jpeg",
  "image/png",
  "image/gif",
  "image/bmp",
  "image/tiff",
  "image/svg+xml",
]);

function textOf(node: OfficeContentNode): string {
  return (node.text ?? "").replace(/\s+/g, " ").trim();
}

function attachmentNameOf(node: OfficeContentNode): string | undefined {
  const metadata = node.metadata;
  if (metadata && "attachmentName" in metadata && typeof metadata.attachmentName === "string") {
    return metadata.attachmentName;
  }
  return undefined;
}

function locationFor(context: LocationContext): string {
  const parts: string[] = [];
  if (context.role) parts.push(context.role);
  if (context.slideNumber !== undefined) parts.push(`Slide ${context.slideNumber}`);
  if (context.sectionNumber !== undefined) parts.push(`Section ${context.sectionNumber}`);
  if (context.heading) parts.push(`Section: ${context.heading}`);
  if (context.inNotes) parts.push("Notes");
  return parts.join(" - ") || "Document body";
}

export function sectionBoundariesFor(ast: OfficeParserAST): boolean[] | undefined {
  if (ast.type !== "docx") return undefined;
  return ast.content.map((node) => /<w:sectPr(?:\s|\/?>)/.test(node.rawContent ?? ""));
}

function collectAttachmentLocations(
  nodes: OfficeContentNode[],
  locations: Map<string, string>,
  parent: LocationContext,
  sectionBoundaries?: readonly boolean[],
): void {
  let inherited = { ...parent };
  for (const [index, node] of nodes.entries()) {
    const metadata = node.metadata;
    const context: LocationContext = { ...inherited };

    if (metadata && "sectionNumber" in metadata && typeof metadata.sectionNumber === "number") {
      context.sectionNumber = metadata.sectionNumber;
    }
    if (node.type === "slide" && metadata && "slideNumber" in metadata) {
      context.slideNumber = metadata.slideNumber;
    }
    if (node.type === "note") {
      context.inNotes = true;
      if (metadata && "slideNumber" in metadata && metadata.slideNumber !== undefined) {
        context.slideNumber = metadata.slideNumber;
      }
    }
    if (node.type === "heading") {
      const heading = textOf(node);
      if (heading) context.heading = heading;
    }

    const attachmentName = attachmentNameOf(node);
    if (attachmentName && !locations.has(attachmentName)) {
      locations.set(attachmentName, locationFor(context));
    }

    if (node.children) collectAttachmentLocations(node.children, locations, context);
    if (node.notes) collectAttachmentLocations(node.notes, locations, { ...context, inNotes: true });
    if (node.comments) collectAttachmentLocations(node.comments, locations, context);

    // Headings establish the nearest semantic section for following blocks in
    // DOCX and ODT, whose AST has no standalone section node.
    if (node.type === "heading" && context.heading) inherited.heading = context.heading;
    if (node.type === "slide" && context.slideNumber !== undefined) {
      inherited.slideNumber = context.slideNumber;
    }
    if (sectionBoundaries?.[index] && index < nodes.length - 1) {
      inherited.sectionNumber = (inherited.sectionNumber ?? 1) + 1;
    }
  }
}

function safeExtension(attachment: OfficeAttachment): string {
  const fromParser = attachment.extension.replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
  if (fromParser) return `.${fromParser}`;

  const fromName = extname(attachment.name).replace(/[^a-zA-Z0-9.]/g, "").toLowerCase();
  return fromName || ".bin";
}

function mediaMimeType(attachment: OfficeAttachment): string {
  const mimeType = String(attachment.mimeType);
  if (attachment.type === "chart" && mimeType === OOXML_DOCUMENT_MIME) return OOXML_CHART_MIME;
  return mimeType;
}

export async function writeMedia(
  ast: OfficeParserAST,
  temporaryDirectory: string,
  abortSignal?: AbortSignal,
): Promise<MediaRecord[]> {
  abortSignal?.throwIfAborted();
  const locations = new Map<string, string>();
  const sectionBoundaries = sectionBoundariesFor(ast);
  const hasMultipleSections = sectionBoundaries?.some(
    (boundary, index) => boundary && index < ast.content.length - 1,
  ) ?? false;
  collectAttachmentLocations(
    ast.content,
    locations,
    { inNotes: false, ...(hasMultipleSections ? { sectionNumber: 1 } : {}) },
    sectionBoundaries,
  );
  if (ast.auxiliary) {
    collectAttachmentLocations(ast.auxiliary.headers ?? [], locations, {
      inNotes: false,
      role: "Header",
    });
    collectAttachmentLocations(ast.auxiliary.footers ?? [], locations, {
      inNotes: false,
      role: "Footer",
    });
    collectAttachmentLocations(ast.auxiliary.slideMasters ?? [], locations, {
      inNotes: false,
      role: "Slide master",
    });
  }

  abortSignal?.throwIfAborted();
  await mkdir(temporaryDirectory, { recursive: true, mode: 0o700 });
  const records: MediaRecord[] = [];

  for (const [index, attachment] of ast.attachments.entries()) {
    abortSignal?.throwIfAborted();
    const label = `media-${index + 1}`;
    const temporaryPath = join(temporaryDirectory, `${label}${safeExtension(attachment)}`);
    await writeFile(temporaryPath, Buffer.from(attachment.data, "base64"), { mode: 0o600 });
    const entry: MediaIndexEntry = {
      label,
      type: attachment.type,
      originalName: attachment.name,
      mimeType: mediaMimeType(attachment),
      temporaryPath,
      location: locations.get(attachment.name) ?? "Document body",
    };
    records.push({ entry, attachment });
  }

  abortSignal?.throwIfAborted();
  return records;
}

export function mediaTable(records: MediaRecord[]): string {
  const cell = (value: string) => value.replace(/\|/g, "\\|").replace(/[\r\n]+/g, " ");
  const rows = records.map(({ entry }) =>
    `| \`${cell(entry.label)}\` | ${cell(entry.type)} | ${cell(entry.originalName)} | ${cell(entry.mimeType)} | ${cell(entry.temporaryPath)} | ${cell(entry.location)} |`,
  );

  return [
    "## Embedded media",
    "",
    "| Label | Type | Original attachment | MIME type | Temporary path | Location |",
    "| --- | --- | --- | --- | --- | --- |",
    ...(rows.length ? rows : ["| _none_ |  |  |  |  |  |"]),
  ].join("\n");
}

export class MediaSelectionError extends Error {
  readonly code = "MEDIA_SELECTOR_INVALID";

  constructor(message: string) {
    super(message);
    this.name = "MediaSelectionError";
  }
}

function isSupportedImageMimeType(mimeType: string): boolean {
  return SUPPORTED_IMAGE_MIME_TYPES.has(mimeType.toLowerCase());
}

export function selectMedia(records: MediaRecord[], selectors: string[] | undefined): MediaRecord[] {
  if (!selectors?.length) return [];

  const selected: MediaRecord[] = [];
  for (const selector of selectors) {
    const record = records.find(
      ({ entry }) => entry.label === selector || entry.originalName === selector,
    );
    if (!record) {
      const available = records.map(({ entry }) => entry.label).join(", ") || "none";
      throw new MediaSelectionError(`Unknown media selector "${selector}". Available media labels: ${available}.`);
    }
    if (record.attachment.type !== "image" || !isSupportedImageMimeType(record.entry.mimeType)) {
      throw new MediaSelectionError(
        `Media selector "${selector}" does not name a supported image and cannot be attached.`,
      );
    }
    if (!selected.some((item) => item.entry.label === record.entry.label)) selected.push(record);
  }
  return selected;
}

export function selectedAttachments(records: MediaRecord[]) {
  return records.map(({ entry, attachment }) => ({
    type: "file" as const,
    mime: entry.mimeType,
    url: `data:${entry.mimeType};base64,${attachment.data}`,
    filename: basename(entry.temporaryPath),
  }));
}
