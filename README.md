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
- Consent to Microsoft Graph scopes used by this extension:
  - `Files.Read`

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

## Notes

- Preview works best for text files. Binary versions will show a placeholder message in the preview tab.
- `Restore Selected Version` writes bytes to the local file. OneDrive sync then uploads it as the current cloud version.
- If OneDrive environment variables are unavailable, the extension also tries to infer a local OneDrive root from folder names like `OneDrive` or `OneDrive - <Org>`.
