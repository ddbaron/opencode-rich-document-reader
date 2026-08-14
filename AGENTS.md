# Project agent memory

This file is the project's committed home for project-intrinsic agent knowledge: build, test, release, architecture, and sharp-edge notes that should travel with the code.

- Add durable project-specific notes here as they are discovered through real work.

## Project notes

- The published package entry point and OpenCode tool registration live in `src/index.ts`; the reader, safety checks, format registry, and media extraction are split across `src/`.
- Run `npm run check && npm test` for the local typecheck and executable fixture suite. The DOCX, ODT, and PPTX fixtures are generated as real ZIP archives by `test/fixtures.ts`.

## Maintaining this file

Keep this file for knowledge useful to almost every future agent session in this project.
Do not repeat what the codebase already shows; point to the authoritative file or command instead.
Prefer rewriting or pruning existing entries over appending new ones.
When updating this file, preserve this bar for all agents and keep entries concise.
