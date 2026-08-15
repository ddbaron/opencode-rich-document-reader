import { type Plugin, tool } from "@opencode-ai/plugin";
import { readRichDocument } from "./reader.ts";

export const RichDocumentReaderPlugin: Plugin = async () => ({
  tool: {
    read_rich_document: tool({
      description:
        "Read a DOCX, ODT, or PPTX file as structure-preserving Markdown and list its extracted media. Select image labels only when vision inspection is needed. Pass export: {} to save a durable sibling export, or export.destination for another project-relative directory.",
      args: {
        path: tool.schema.string().describe("Path to a DOCX, ODT, or PPTX inside the current project."),
        media: tool.schema
          .array(tool.schema.string())
          .optional()
          .describe("Optional media labels from the embedded media index to attach as images."),
        export: tool.schema
          .object({
            destination: tool.schema
              .string()
              .optional()
              .describe("Optional project-relative destination directory; omit for a sibling .export directory."),
          })
          .optional()
          .describe(
            "Explicitly persist the Markdown, all extracted media, and a manifest. Omit to keep the current temporary behavior.",
          ),
      },
      async execute(args, context) {
        return readRichDocument(args, context);
      },
    }),
  },
});

export default RichDocumentReaderPlugin;
