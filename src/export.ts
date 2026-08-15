import { copyFile, mkdir, rm, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join } from "node:path";
import type {
  DurableExportMedia,
  DurableExportMetadata,
  MediaRecord,
  ResolvedExportDestination,
} from "./types.ts";

export class DurableExportError extends Error {
  readonly code: "EXPORT_EXISTS" | "EXPORT_FAILED";

  constructor(code: DurableExportError["code"], message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "DurableExportError";
    this.code = code;
  }
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return String(error);
}

function fileStem(sourcePath: string): string {
  const stem = basename(sourcePath, extname(sourcePath));
  return stem || basename(sourcePath);
}

export interface DurableExportOptions {
  destination: ResolvedExportDestination;
  sourcePath: string;
  format: string;
  markdown: string;
  records: readonly MediaRecord[];
  abortSignal: AbortSignal;
}

export async function writeDurableExport({
  destination,
  sourcePath,
  format,
  markdown,
  records,
  abortSignal,
}: DurableExportOptions): Promise<DurableExportMetadata> {
  const directoryPath = destination.absolutePath;
  const mediaDirectoryPath = join(directoryPath, "media");
  const markdownPath = join(directoryPath, `${fileStem(sourcePath)}.md`);
  const manifestPath = join(directoryPath, "manifest.json");
  let createdDirectory = false;

  try {
    abortSignal.throwIfAborted();
    await mkdir(dirname(directoryPath), { recursive: true, mode: 0o700 });
    await mkdir(directoryPath, { recursive: false, mode: 0o700 });
    createdDirectory = true;
    await mkdir(mediaDirectoryPath, { mode: 0o700 });

    const media: DurableExportMedia[] = [];
    for (const { entry } of records) {
      abortSignal.throwIfAborted();
      const fileName = basename(entry.temporaryPath);
      const path = join(mediaDirectoryPath, fileName);
      await copyFile(entry.temporaryPath, path);
      media.push({
        label: entry.label,
        type: entry.type,
        originalName: entry.originalName,
        mimeType: entry.mimeType,
        location: entry.location,
        path,
        relativePath: `media/${fileName}`,
      });
    }

    abortSignal.throwIfAborted();
    await writeFile(markdownPath, markdown, { encoding: "utf8", mode: 0o600 });
    const metadata: DurableExportMetadata = {
      directoryPath,
      markdownPath,
      mediaDirectoryPath,
      manifestPath,
      media,
    };
    await writeFile(
      manifestPath,
      `${JSON.stringify(
        {
          version: 1,
          sourcePath,
          format,
          exportDirectory: directoryPath,
          exportDirectoryRelativePath: destination.relativePath,
          markdownPath,
          markdownRelativePath: basename(markdownPath),
          mediaDirectoryPath,
          mediaDirectoryRelativePath: "media",
          manifestPath,
          manifestRelativePath: "manifest.json",
          media,
        },
        null,
        2,
      )}\n`,
      { encoding: "utf8", mode: 0o600 },
    );
    abortSignal.throwIfAborted();
    return metadata;
  } catch (error) {
    if (createdDirectory) {
      await rm(directoryPath, { recursive: true, force: true }).catch(() => undefined);
    }
    if (abortSignal.aborted) abortSignal.throwIfAborted();
    if (error instanceof DurableExportError) throw error;
    const code = error && typeof error === "object" && "code" in error ? error.code : undefined;
    if (code === "EEXIST") {
      throw new DurableExportError(
        "EXPORT_EXISTS",
        `Export destination already exists and was not overwritten: ${directoryPath}`,
        { cause: error },
      );
    }
    throw new DurableExportError(
      "EXPORT_FAILED",
      `Could not write durable export to ${directoryPath}: ${errorMessage(error)}`,
      { cause: error },
    );
  }
}
