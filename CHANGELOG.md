# Changelog

## 0.0.1

- Initial extension scaffold.
- Added OneDrive file detection for active editor files.
- Added version loading from Microsoft Graph.
- Added editor-title dropdown menu with version actions.
- Added preview, save-as, and restore workflows.
- Added fallback local root inference for OneDrive folder names when env-based detection is unavailable.
- Reduced Microsoft Graph delegated scope request to `Files.Read` to improve tenant compatibility.
- Added `deviceCode` authentication mode using MSAL + Entra app registration for tenants that block VS Code first-party Graph auth.
