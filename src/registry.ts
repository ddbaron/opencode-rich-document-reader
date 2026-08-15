import * as officeParser from "officeparser";
import type { OfficeParserConfig, SupportedFileType } from "officeparser";
import type { RichDocumentFormat } from "./types.ts";

type ParseOffice = RichDocumentFormat["parse"];

interface OfficeParserModuleValue {
  OfficeParser?: unknown;
  default?: unknown;
  parseOffice?: unknown;
}

interface ResolvedParser {
  parseOffice: ParseOffice;
  receiver: unknown;
}

function moduleValue(value: unknown): OfficeParserModuleValue | undefined {
  if ((typeof value !== "object" && typeof value !== "function") || value === null) return undefined;
  return value as OfficeParserModuleValue;
}

// Bun's CJS interop can leave the named OfficeParser binding undefined while
// retaining the default class and top-level parseOffice export.
function resolveOfficeParser(value: unknown, seen = new Set<object>()): ResolvedParser | undefined {
  const candidate = moduleValue(value);
  if (!candidate) return undefined;

  const reference = value as object;
  if (seen.has(reference)) return undefined;
  seen.add(reference);

  for (const nested of [candidate.OfficeParser, candidate.default]) {
    const resolved = resolveOfficeParser(nested, seen);
    if (resolved) return resolved;
  }

  if (typeof candidate.parseOffice === "function") {
    return { parseOffice: candidate.parseOffice as ParseOffice, receiver: value };
  }
  return undefined;
}

function parserFor(module: unknown): ResolvedParser {
  const parser = resolveOfficeParser(module);
  if (!parser) throw new TypeError("officeparser does not expose a parseOffice function");
  return parser;
}

export function createOfficeFormat(
  parserType: SupportedFileType,
  parserModule: unknown = officeParser,
): RichDocumentFormat {
  return {
    extension: `.${parserType}`,
    parserType,
    parse(input: Uint8Array, config: OfficeParserConfig) {
      const parser = parserFor(parserModule);
      return parser.parseOffice.call(parser.receiver, input, { ...config, fileType: parserType });
    },
  };
}

/**
 * The format registry is the only format-selection surface used by the tool.
 * Adding a parser-backed format does not change the OpenCode tool arguments or
 * result shape.
 */
export const formatRegistry: ReadonlyMap<string, RichDocumentFormat> = new Map(
  ["docx", "odt", "pptx"].map((format) => {
    const parserType = format as SupportedFileType;
    return [`.${format}`, createOfficeFormat(parserType)] as const;
  }),
);

export function formatForExtension(extension: string, formats = formatRegistry): RichDocumentFormat | undefined {
  return formats.get(extension.toLowerCase());
}
