# OneDrive Versions (VS Code Extension)

This extension adds OneDrive version navigation for files opened from OneDrive-synced folders.

## What it does

- Detects when the active file is inside a OneDrive folder.
- Adds an `OneDrive Versions` dropdown to the editor title bar.
- Lets you:
  - Step to an older version (`Previous Version`)
  - Step to a newer version (`Next Version`)
  - Pick any version from a list
  - Save a selected version to another file (`Save Version As...`)
  - Restore the selected version as the current local file (`Restore Selected Version`)

## Requirements

- VS Code signed in to Microsoft account(s) used for OneDrive.
- OneDrive-synced files available locally.
- Microsoft Graph delegated permission used by this extension:
  - `Files.Read`

## Authentication Modes

- `onedriveVersions.auth.mode = "vscode"`
  - Uses VS Code Microsoft authentication provider.
- `onedriveVersions.auth.mode = "deviceCode"` (production default)
  - Uses MSAL device-code sign-in with your own Entra app registration.
  - Automatically opens the verification URL in your browser and copies the device code to clipboard.
  - Required settings:
    - `onedriveVersions.auth.clientId`
    - Optional `onedriveVersions.auth.tenantId` (default `organizations`)

If your tenant blocks VS Code first-party auth with `AADSTS65002`, switch to `deviceCode` mode.

## First-Run Onboarding

Use command palette:
- `OneDrive: Connect Microsoft Account`
- `OneDrive: Open Setup Guide`

If you see tenant auth error `AADSTS65002`, use `Switch Auth Mode` in the error prompt, then complete device-code setup.
Background auto-load does not trigger interactive sign-in; sign-in prompts are shown when you explicitly run version/account commands.

## Settings

- `onedriveVersions.autoLoadVersions` (default: `true`)
  - Automatically loads versions for active OneDrive files.
- `onedriveVersions.mappings` (default: `[]`)
  - Optional mapping entries:
    - `localRoot` (required): local OneDrive sync root.
    - `driveId` (optional): specific Graph drive ID.
    - `remoteRoot` (optional, default `/`): subpath root in that drive.

Example:

```json
"onedriveVersions.mappings": [
  {
    "localRoot": "C:\\Users\\you\\OneDrive - Contoso",
    "driveId": "b!abc123...",
    "remoteRoot": "/Shared/Engineering"
  }
]
```

Device-code auth example:

```json
"onedriveVersions.auth.mode": "deviceCode",
"onedriveVersions.auth.clientId": "00000000-0000-0000-0000-000000000000",
"onedriveVersions.auth.tenantId": "contoso.onmicrosoft.com"
```

Recommended production defaults (already shipped in this repo):

```json
"onedriveVersions.auth.mode": "deviceCode",
"onedriveVersions.auth.clientId": "6bb315fa-774e-4147-8e0c-2afd44ffb86e",
"onedriveVersions.auth.tenantId": "organizations"
```

For external tenants, first run may still require tenant admin consent to this app's `Files.Read` permission.

## Notes

- Preview works best for text files. Binary versions will show a placeholder message in the preview tab.
- `Restore Selected Version` writes bytes to the local file. OneDrive sync then uploads it as the current cloud version.
- If OneDrive environment variables are unavailable, the extension also tries to infer a local OneDrive root from folder names like `OneDrive` or `OneDrive - <Org>`.
- On Windows, the extension also reads OneDrive sync mount points from `HKCU\\Software\\SyncEngines\\Providers\\OneDrive`.
- If a file path is not found in `/me/drive`, the extension falls back to searching your accessible drives (`/me/drives`) for the same relative path.
- For synced SharePoint/library mounts, it also retries with trimmed leading path segments when resolving the remote item path.
- For synced library folders, it also uses OneDrive registry URL metadata (`FullRemotePath`/`UrlNamespace`) and Graph `/shares/{encodedUrl}` resolution.
