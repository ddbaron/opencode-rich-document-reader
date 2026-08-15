import { access, lstat, realpath, stat } from "node:fs/promises";
import { constants } from "node:fs";
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep, win32 } from "node:path";
import type { ProjectPaths, ResolvedDocument, ResolvedExportDestination } from "./types.ts";

export class DocumentPathError extends Error {
  readonly code:
    | "INVALID_PATH"
    | "PROJECT_UNAVAILABLE"
    | "MISSING_DOCUMENT"
    | "PATH_ESCAPE"
    | "SYMLINK_ESCAPE"
    | "NOT_A_FILE"
    | "UNREADABLE_DOCUMENT";

  constructor(code: DocumentPathError["code"], message: string) {
    super(message);
    this.name = "DocumentPathError";
    this.code = code;
  }
}

export class ExportPathError extends Error {
  readonly code:
    | "INVALID_EXPORT_DESTINATION"
    | "EXPORT_PATH_ESCAPE"
    | "EXPORT_SYMLINK_ESCAPE"
    | "EXPORT_DESTINATION_UNREADABLE";

  constructor(code: ExportPathError["code"], message: string) {
    super(message);
    this.name = "ExportPathError";
    this.code = code;
  }
}

function isWithin(root: string, candidate: string): boolean {
  const distance = relative(root, candidate);
  return distance === "" || (distance !== ".." && !distance.startsWith(`..${sep}`) && !isAbsolute(distance));
}

async function canonicalProjectRoot(root: string): Promise<string> {
  try {
    return await realpath(resolve(root));
  } catch {
    throw new DocumentPathError("PROJECT_UNAVAILABLE", `Current project is not available: ${root}`);
  }
}

function exportStem(sourcePath: string): string {
  const extension = extname(sourcePath);
  const stem = basename(sourcePath, extension);
  return stem || basename(sourcePath);
}

function pathErrorCode(error: unknown): string | undefined {
  return error && typeof error === "object" && "code" in error ? String(error.code) : undefined;
}

async function assertExportPathInsideProject(projectRoot: string, candidate: string): Promise<void> {
  let current = candidate;
  while (true) {
    try {
      const canonical = await realpath(current);
      if (!isWithin(projectRoot, canonical)) {
        throw new ExportPathError(
          "EXPORT_SYMLINK_ESCAPE",
          `Export destination symlink escapes the current project: ${candidate}`,
        );
      }
      return;
    } catch (error) {
      if (error instanceof ExportPathError) throw error;
      if (pathErrorCode(error) !== "ENOENT" && pathErrorCode(error) !== "ENOTDIR") {
        throw new ExportPathError(
          "EXPORT_DESTINATION_UNREADABLE",
          `Export destination cannot be inspected: ${candidate}`,
        );
      }
      const parent = dirname(current);
      if (parent === current) {
        throw new ExportPathError(
          "EXPORT_DESTINATION_UNREADABLE",
          `Export destination cannot be inspected: ${candidate}`,
        );
      }
      current = parent;
    }
  }
}

export async function resolveExportDestination(
  destination: string | undefined,
  rootInput: string,
  sourcePath: string,
): Promise<ResolvedExportDestination> {
  if (destination !== undefined && (!destination.trim() || destination.includes("\0"))) {
    throw new ExportPathError("INVALID_EXPORT_DESTINATION", "Export destination must be a non-empty path.");
  }

  const projectRoot = await canonicalProjectRoot(rootInput);
  const candidate =
    destination === undefined
      ? join(dirname(sourcePath), `${exportStem(sourcePath)}.export`)
      : resolve(projectRoot, destination);

  if (destination !== undefined && (isAbsolute(destination) || win32.isAbsolute(destination))) {
    throw new ExportPathError(
      "EXPORT_PATH_ESCAPE",
      `Export destination must be project-relative, not absolute: ${destination}`,
    );
  }
  if (!isWithin(projectRoot, candidate)) {
    throw new ExportPathError(
      "EXPORT_PATH_ESCAPE",
      `Export destination escapes the current project: ${destination ?? candidate}`,
    );
  }

  await assertExportPathInsideProject(projectRoot, candidate);
  const relativePath = relative(projectRoot, candidate).split(sep).join("/") || ".";
  return { projectRoot, absolutePath: candidate, relativePath };
}

export async function resolveDocumentPath(documentInput: string, rootInput: string): Promise<ResolvedDocument> {
  if (!documentInput.trim() || documentInput.includes("\0")) {
    throw new DocumentPathError("INVALID_PATH", "A non-empty document path is required.");
  }

  const projectRoot = await canonicalProjectRoot(rootInput);
  const lexicalDocumentPath = resolve(projectRoot, documentInput);

  if (!isWithin(projectRoot, lexicalDocumentPath)) {
    throw new DocumentPathError(
      "PATH_ESCAPE",
      `Document path escapes the current project: ${documentInput}`,
    );
  }

  try {
    await lstat(lexicalDocumentPath);
  } catch (error) {
    const code = error && typeof error === "object" && "code" in error ? error.code : undefined;
    if (code === "ENOENT" || code === "ENOTDIR") {
      throw new DocumentPathError("MISSING_DOCUMENT", `Document not found: ${documentInput}`);
    }
    throw new DocumentPathError("UNREADABLE_DOCUMENT", `Document cannot be inspected: ${documentInput}`);
  }

  let documentPath: string;
  try {
    documentPath = await realpath(lexicalDocumentPath);
  } catch (error) {
    const code = error && typeof error === "object" && "code" in error ? error.code : undefined;
    if (code === "ENOENT" || code === "ENOTDIR") {
      throw new DocumentPathError("MISSING_DOCUMENT", `Document not found: ${documentInput}`);
    }
    throw new DocumentPathError("UNREADABLE_DOCUMENT", `Document cannot be resolved: ${documentInput}`);
  }

  if (!isWithin(projectRoot, documentPath)) {
    throw new DocumentPathError(
      "SYMLINK_ESCAPE",
      `Document symlink resolves outside the current project: ${documentInput}`,
    );
  }

  let documentStat;
  try {
    documentStat = await stat(documentPath);
    if (!documentStat.isFile()) {
      throw new DocumentPathError("NOT_A_FILE", `Document path is not a regular file: ${documentInput}`);
    }
    await access(documentPath, constants.R_OK);
  } catch (error) {
    if (error instanceof DocumentPathError) throw error;
    throw new DocumentPathError("UNREADABLE_DOCUMENT", `Document is not readable: ${documentInput}`);
  }

  return {
    absolutePath: documentPath,
    extension: extname(documentPath).toLowerCase(),
    size: documentStat.size,
  };
}

export function projectPaths(context: { directory: string; worktree: string }, documentInput: string): ProjectPaths {
  const projectRoot = context.worktree || context.directory;
  return { projectRoot, documentPath: documentInput };
}
