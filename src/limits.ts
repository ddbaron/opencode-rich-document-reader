import type { OfficeParserConfig } from "officeparser";

/** Limits are deliberately conservative because documents are supplied to an agent. */
export const READER_LIMITS = {
  maxArchiveBytes: 64 * 1024 * 1024,
  maxUncompressedBytes: 128 * 1024 * 1024,
  maxZipEntries: 4096,
  maxTableCells: 250_000,
} as const;

export function parserLimits(): Pick<OfficeParserConfig, "decompressionLimits"> {
  return {
    decompressionLimits: {
      maxUncompressedBytes: READER_LIMITS.maxUncompressedBytes,
      maxZipEntries: READER_LIMITS.maxZipEntries,
      maxTableCells: READER_LIMITS.maxTableCells,
    },
  };
}
