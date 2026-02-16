# AGENTS.md

## Purpose
Keep this repository focused on a reliable VS Code extension that lets users browse, preview, save-as, and restore OneDrive file versions from the editor UI.

## Product Scope
- Extension ID: `onedrive-versions`
- Runtime: VS Code extension host (Node/TypeScript)
- Primary file: `src/extension.ts`
- UI surface: editor title menu and command palette commands
- Backend API: Microsoft Graph OneDrive item versions API

## Non-Negotiable Behaviors
- Only show OneDrive version actions when active file is in a detected OneDrive local root.
- Version list must be sorted newest to oldest by `lastModifiedDateTime`.
- `Previous Version` moves to older history; `Next Version` moves toward newer history.
- `Save Version As...` must never overwrite without explicit user choice in save dialog.
- `Restore Selected Version` must warn before writing bytes to active local file.
- Never silently discard unsaved local edits before restore.

## Engineering Standards
- Language: TypeScript, strict mode enabled.
- Keep logic in small functions with explicit error messages.
- Avoid adding dependencies unless clearly justified.
- Preserve cross-platform path handling (Windows/macOS/Linux).
- Prefer ASCII content in docs and source unless project already requires otherwise.

## Required Validation for Changes
- Run: `npm run compile`
- If command/UX behavior changes, update:
  - `README.md`
  - `CHANGELOG.md`
  - `MEMORY.md` (Current State + Next Steps)
- If settings or commands change, ensure `package.json` `contributes` section is updated.

## Documentation Discipline
- Keep `README.md` user-focused (install, setup, usage, settings, limitations).
- Keep `MEMORY.md` execution-focused (current implementation, gaps, risks, and resume steps).
- Update `MEMORY.md` in the same change set as code edits that alter behavior.

## Git Workflow
- Commit small, coherent changes.
- Use clear commit messages describing functional outcome.
- Do not commit secrets, tokens, or local machine paths that are not examples.
- Keep `.gitignore` aligned with Node + VS Code extension artifacts.

## Near-Term Roadmap
- Improve binary preview handling (open as readonly temp file for non-text content).
- Add automated tests for mapping resolution and version index navigation behavior.
- Add packaging/publishing notes for Marketplace and GitHub releases.
