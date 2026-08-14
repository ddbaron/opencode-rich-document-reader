# OpenCode Rich Document Reader

An OpenCode plugin that lets coding and review agents inspect `.docx`, `.odt`, and `.pptx` documents as structured Markdown with embedded media references.

## Install

Install the public GitHub plugin with OpenCode:

```sh
opencode plugin github:ddbaron/opencode-rich-document-reader
```

Restart OpenCode after installation if the current session does not list `read_rich_document`.

OpenCode installs the plugin's `@opencode-ai/plugin` and `officeparser` dependencies automatically.

## Use

Ask the agent to call `read_rich_document` with a document path inside the current project:

```json
{"path":"docs/architecture.docx"}
```

The tool returns structure-preserving Markdown for headings, lists, tables, links, sections, slide context, notes, and supported document content. It also returns a structured media index in tool metadata.

It appends an ordered embedded-media index containing each attachment's source name, MIME type, document section, and isolated temporary path.

Use a media label from the index when vision inspection is relevant. The selected image is returned as a native OpenCode file attachment, while its isolated temporary path remains available for follow-up tools.

Media is not attached by default, which prevents documents with many screenshots from inflating the model context.

To attach selected image media directly to the next model turn, call the tool again with a returned media label:

```json
{"path":"docs/architecture.docx","media":["media-1"]}
```

Selectable image media supports `image/jpeg`, `image/png`, `image/gif`, `image/bmp`, `image/tiff`, `image/svg+xml`, and `image/webp` attachments.

## Safety and boundaries

The source document is read without modification.

Paths must resolve inside the current project, and symlinks that escape it are rejected.

Extracted files are written beneath a unique system temporary directory rather than beside the source document.

The parser applies bounded archive, entry-count, and table-cell limits.

Unsupported extensions, missing files, unreadable paths, malformed archives, and missing dependencies produce explicit errors.

PDFs and spreadsheets are intentionally outside this initial interface.

## Development

Install dependencies and run the checks:

```sh
npm ci
npm run check
npm test
```

The test suite creates minimal DOCX, ODT, and PPTX fixtures, verifies structure and media extraction, and exercises path and malformed-document errors.

## License

MIT.
