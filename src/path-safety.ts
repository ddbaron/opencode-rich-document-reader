import { access, lstat, realpath, stat } from "node:fs/promises";
import { constants } from "node:fs";
import { extname, isAbsolute, relative, resolve, sep } from "node:path";
import type { ProjectPaths, ResolvedDocument } from "./types.ts";

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
