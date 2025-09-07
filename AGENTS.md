# Repository Guidelines

## Project Structure & Module Organization
- `code.gas.js`: Single entry file for Google Apps Script (GAS). Exposes runnable functions at bottom (e.g., `onOpen`, `runCreateEstimate`, `testCore`).
- `types/*.d.ts`: Local type helpers for GAS Advanced Services (Sheets) and utilities.
- `tsconfig.json`: Strict type-check for JS via `// @ts-check` with no emit.
- `coding-guide.md`: Detailed coding standards used by this project.
- `package.json`, `pnpm-lock.yaml`, `mise.toml`: Tooling metadata; no bundler or transpile step.

## Build, Test, and Development Commands
- `mise install`: Sync toolchain (Node 22, pnpm 10).
- `pnpm i`: Install dev dependencies (types, TS checker).
- `pnpm exec tsc --noEmit`: Type-check `code.gas.js` with strict settings.
- Tests run in GAS: open the Apps Script editor and run `testCore`, `testTemplateCore`, or specific helpers like `testPoMembersCore`. Execution logs appear in `Logger`.

## Coding Style & Naming Conventions
- Use `const` + arrow functions; avoid `var`. One action per function.
- Entry points at file end as top-level `const` functions so GAS can discover them.
- Logging via `Logger.log` only. Wrap with `logInfo/logWarn/logError`.
- 2-space indent; semicolons required; target line length ≤ 120.
- Add JSDoc for params/returns; keep `// @ts-check` on; prefer `UpperCamel` for types, `lowerCamel` for functions, `SCREAMING_SNAKE` for constants.

## Testing Guidelines
- No Node test runner. Register tests in the in-file `tests` array and expose wrappers (e.g., `testSampleCore`).
- Run named tests via `runTestByName("name")` or batch with `runTestsByNames([..])` (see bottom of `code.gas.js`).
- Keep tests side-effect free unless explicitly marked; prefer separating pure calc from GAS I/O.

## Commit & Pull Request Guidelines
- Follow Conventional Commits when possible (`feat:`, `fix:`), otherwise concise, imperative subject. Reference issues/PRs.
- PRs should include: purpose, scope, before/after notes, and how to verify (which GAS test functions to run). Add screenshots of Sheets/Form changes when relevant and paste key `Logger` lines.

## Security & Configuration
- Do not commit secrets or spreadsheet IDs tied to private data. Centralize such values in a `CONFIG` block and document required scopes.
- Enable required Advanced Services/APIs in GAS (Sheets API). Keep least privilege triggers and share artifacts with minimal access.
