# MEMORY.md

## Project
- Name: `onedrive-versions`
- Goal: VS Code extension for browsing and restoring OneDrive document versions from within the editor.
- Status date: 2026-02-16

## Current State
- Extension scaffold is complete and compiles successfully.
- Commands implemented:
  - `onedriveVersions.pickVersion`
  - `onedriveVersions.previousVersion`
  - `onedriveVersions.nextVersion`
  - `onedriveVersions.saveAsVersion`
  - `onedriveVersions.restoreVersion`
- Editor title submenu is contributed and conditionally shown based on context keys:
  - `oneDriveVersions.active`
  - `oneDriveVersions.hasVersions`
- OneDrive detection sources:
  - User settings mappings: `onedriveVersions.mappings`
  - Environment: `OneDrive`, `OneDriveCommercial`, `OneDriveConsumer`
- Graph auth uses VS Code Microsoft auth provider with scopes:
  - `User.Read`
  - `Files.ReadWrite.All`
- Content preview provider:
  - Scheme: `onedrive-version`
  - Decodes fetched bytes as UTF-8 text
  - Shows placeholder text for likely binary content
- Local git repo is initialized with first commit on `main`.
- Remote configured:
  - `origin = https://github.com/NichUK/onedrive-versions.git`
- Push attempt failed because GitHub repository does not yet exist.

## Files of Interest
- `src/extension.ts`: core extension logic
- `package.json`: command/menu/settings contributions
- `README.md`: user docs and setup guidance
- `CHANGELOG.md`: release notes
- `.gitignore`: ignore policy for build artifacts
- `AGENTS.md`: project guardrails and workflow

## Known Gaps / Risks
- No automated tests yet.
- Binary version preview is text fallback only (not binary-aware diff/view).
- Restore currently writes local file bytes; cloud “restore” is achieved via sync upload rather than a dedicated Graph restore endpoint action.
- GitHub repo creation/publish is blocked until repo exists or authenticated GitHub API/CLI access is available.

## Resume Checklist
1. Run `npm run compile` to verify baseline.
2. Launch extension host with `F5` and validate commands on a real OneDrive-synced file.
3. Confirm mapping behavior for at least:
   - default `/me/drive` case
   - explicit `driveId` + `remoteRoot` case
4. Add tests around:
   - path mapping normalization
   - version ordering
   - previous/next index clamping
5. If behavior changes, update `README.md`, `CHANGELOG.md`, and this file.

## Publishing Notes
- Intended public GitHub repo: `NichUK/onedrive-versions`
- If GitHub CLI is unavailable, create repo via web UI, then:
  - `git init`
  - `git add .`
  - `git commit -m "Initial commit"`
  - `git branch -M main`
  - `git remote add origin https://github.com/NichUK/onedrive-versions.git`
  - `git push -u origin main`
