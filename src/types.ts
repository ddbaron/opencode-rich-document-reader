import type { ToolContext, ToolResult } from "@opencode-ai/plugin";
import type {
  OfficeAttachment,
  OfficeContentNode,
  OfficeParserAST,
  OfficeParserConfig,
  SupportedFileType,
} from "officeparser";

export interface DocumentIssue {
  type: "warning" | "info" | "error" | string;
  code: string;
  message: string;
}

export interface ReadRichDocumentArgs {
  path: string;
  media?: string[];
}

export type ReaderToolContext = Pick<ToolContext, "directory" | "worktree" | "abort">;

export interface RichDocumentFormat {
  extension: string;
  parserType: SupportedFileType;
  parse(input: Uint8Array, config: OfficeParserConfig): Promise<OfficeParserAST>;
}

export interface MediaIndexEntry {
  label: string;
  type: OfficeAttachment["type"];
  originalName: string;
  mimeType: string;
  temporaryPath: string;
  location: string;
}

export interface RichDocumentResultMetadata {
  format: SupportedFileType;
  sourcePath: string;
  media: MediaIndexEntry[];
}

export type RichDocumentToolResult = Omit<Extract<ToolResult, { output: string }>, "metadata"> & {
  metadata?: RichDocumentResultMetadata;
};

export interface MediaRecord {
  entry: MediaIndexEntry;
  attachment: OfficeAttachment;
}

export interface ParseIssue {
  issue: DocumentIssue;
  source: "parser" | "conversion";
}

export interface ResolvedDocument {
  absolutePath: string;
  extension: string;
  size: number;
}

export interface ProjectPaths {
  projectRoot: string;
  documentPath: string;
}

export interface ReadRichDocumentDependencies {
  formats?: ReadonlyMap<string, RichDocumentFormat>;
  tempDirectory?: string;
}
