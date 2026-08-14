import { OfficeParser } from "officeparser";
import type { OfficeParserConfig, SupportedFileType } from "officeparser";
import type { RichDocumentFormat } from "./types.ts";

function officeFormat(parserType: SupportedFileType): RichDocumentFormat {
  return {
    extension: `.${parserType}`,
    parserType,
    parse(input: Uint8Array, config: OfficeParserConfig) {
      return OfficeParser.parseOffice(input, { ...config, fileType: parserType });
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
    return [`.${format}`, officeFormat(parserType)] as const;
  }),
);

export function formatForExtension(extension: string, formats = formatRegistry): RichDocumentFormat | undefined {
  return formats.get(extension.toLowerCase());
}
