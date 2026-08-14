import { type Plugin, tool } from "@opencode-ai/plugin";
import { readRichDocument } from "./reader.ts";

export const RichDocumentReaderPlugin: Plugin = async () => ({
  tool: {
    read_rich_document: tool({
      description:
        "Read a DOCX, ODT, or PPTX file as structure-preserving Markdown and list its extracted media. Select image labels only when vision inspection is needed.",
      args: {
        path: tool.schema.string().describe("Path to a DOCX, ODT, or PPTX inside the current project."),
        media: tool.schema
          .array(tool.schema.string())
          .optional()
          .describe("Optional media labels from the embedded media index to attach as images."),
      },
      async execute(args, context) {
        return readRichDocument(args, context);
      },
    }),
  },
});

export default RichDocumentReaderPlugin;
