# MEMORY.md

## Project
- Name: `onedrive-versions`
- Goal: VS Code extension for browsing and restoring OneDrive document versions from within the editor.
- Status date: 2026-02-16

## Current State
- Extension scaffold is complete and compiles successfully.
- Commands implemented:
  - `onedriveVersions.connectAccount`
  - `onedriveVersions.openSetupGuide`
  - `onedriveVersions.pickVersion`
  - `onedriveVersions.previousVersion`
  - `onedriveVersions.nextVersion`
  - `onedriveVersions.saveAsVersion`
  - `onedriveVersions.restoreVersion`
- Editor title submenu is contributed and conditionally shown based on context keys:
  - `oneDriveVersions.active`
  - `oneDriveVersions.hasVersions`
- Editor title now uses an always-visible `Pick Version` command button, enabled only when `oneDriveVersions.active` is true.
- Active-file resolution now treats `onedrive-version:` preview documents as their source local file so version controls remain enabled while browsing previews.
- Version content retrieval now falls back to `/items/{id}/content` when Graph rejects `/versions/{id}/content` for current-version IDs.
- Version previews now open in diff mode (`onedrive-version` vs current local file).
- Status bar badge shows selected OneDrive version timestamp and remains clickable to re-open picker.
- OneDrive detection sources:
  - User settings mappings: `onedriveVersions.mappings`
  - Environment: `OneDrive`, `OneDriveCommercial`, `OneDriveConsumer`
  - Windows registry metadata from `HKCU\\Software\\SyncEngines\\Providers\\OneDrive` (`MountPoint`, `UrlNamespace`, `FullRemotePath`)
  - Fallback path inference from local folder segment names matching `OneDrive` / `OneDrive - <Org>`
- Graph item resolution strategy:
  - Try `/me/drive/root:{path}` first
  - On `itemNotFound`, fallback to iterate `/me/drives` and resolve the same path in each drive
  - Also retry with progressively trimmed leading path segments for mount-root mismatches
  - If still unresolved and registry URL metadata exists, resolve item via Graph `/shares/{encodedUrl}/driveItem`
  - If `/shares` is blocked (`accessDenied`), fallback to URL-prefix matching against drive `webUrl` and resolve via `/drives/{id}/root:/...`
- Graph auth uses VS Code Microsoft auth provider with scopes:
  - `Files.Read.All`
- Additional auth mode available:
  - Device code via MSAL (`onedriveVersions.auth.mode = deviceCode`) using user-provided Entra app `clientId`
  - Device-code flow now auto-opens verification URL and copies user code to clipboard
- Onboarding UX implemented:
  - First-run prompt when device-code mode has no `clientId`
  - Actionable auth-error prompts (switch auth mode, open settings, open setup guide)
  - Background auto-load uses non-interactive auth checks to avoid surprise sign-in prompts
- Production defaults now ship with:
  - `onedriveVersions.auth.mode = deviceCode`
  - `onedriveVersions.auth.clientId = 6bb315fa-774e-4147-8e0c-2afd44ffb86e`
  - `onedriveVersions.auth.tenantId = organizations`
- Content preview provider:
  - Scheme: `onedrive-version`
  - Decodes fetched bytes as UTF-8 text
  - Shows placeholder text for likely binary content
- Local git repo is initialized with first commit on `main`.
- GitHub repo created and published (public): `https://github.com/NichUK/onedrive-versions`
- Remote configured:
  - `origin = git@github.com:NichUK/onedrive-versions.git`

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
- Tenant policy may block VS Code first-party Graph auth (`AADSTS65002`); use device-code auth mode in that case.
- Unit tests added for resolver helpers under `src/test`.

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
- Current publish flow used:
  - Create repo via GitHub API using stored Git credentials
  - Push `main` via SSH remote
- If GitHub CLI is unavailable, alternative is web UI + push:
  - `git init`
  - `git add .`
  - `git commit -m "Initial commit"`
  - `git branch -M main`
  - `git remote add origin https://github.com/NichUK/onedrive-versions.git`
  - `git push -u origin main`
