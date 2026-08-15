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

For calls without `export`, it appends an ordered embedded-media index containing each attachment's source name, MIME type, document section, and isolated temporary path.

Use a media label from the index when vision inspection is relevant. The selected image is returned as a native OpenCode file attachment, while its isolated temporary path remains available for follow-up tools.

Media is not attached by default, which prevents documents with many screenshots from inflating the model context.

To attach selected image media directly to the next model turn, call the tool again with a returned media label:

```json
{"path":"docs/architecture.docx","media":["media-1"]}
```

Selectable image media supports `image/jpeg`, `image/png`, `image/gif`, `image/bmp`, `image/tiff`, `image/svg+xml`, and `image/webp` attachments.

### Durable export (opt-in)

Calls without `export` keep the existing ephemeral behavior. To explicitly save a durable copy beside the source document, pass an empty export request:

```json
{"path":"docs/architecture.docx","export":{}}
```

For `docs/architecture.docx`, the default sibling export directory is `docs/architecture.export/`. It contains `architecture.md`, a `media/` directory containing every extracted attachment, and `manifest.json`. The exported Markdown preserves the document structure and uses relative paths such as `media/media-1.png` in its media index.

Callers may choose another durable directory with a project-relative destination:

```json
{"path":"docs/architecture.docx","export":{"destination":"artifacts/architecture"}}
```

The destination is an output directory relative to the current project. Missing parent directories are created, but an existing destination is rejected and never overwritten; choose a new destination or remove the old export explicitly before retrying. Absolute destinations and paths that escape the project, including escaping symlinks, are rejected.

The result metadata includes `metadata.export.directoryPath`, `markdownPath`, `mediaDirectoryPath`, `manifestPath`, and durable paths for each exported media item. The manifest is JSON and repeats those paths and the source-relative path so later turns can use the saved artifacts without depending on the temporary extraction directory. Export is never automatic, and the source document is not changed.

## Safety and boundaries

The source document is read without modification.

Paths must resolve inside the current project, and symlinks that escape it are rejected.

Without an explicit `export` request, extracted files are written beneath a unique system temporary directory rather than beside the source document, and the existing temporary media index is returned. Successful reads keep that directory available for the current agent workflow; failed reads clean up partial extraction. An explicit export additionally writes the durable directory described above and does not change the non-export contract.

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

The test suite creates minimal DOCX, ODT, and PPTX fixtures, verifies structure and media extraction, exercises durable export contents and path/collision behavior, and covers malformed-document errors.

## License

MIT.
