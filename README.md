# TeamSpy

TeamSpy is a standalone Electron desktop app for internal use on macOS and Windows. It signs in with Microsoft Entra, reads Microsoft Graph data locally, and shows a configurable Teams directory table with:

- sortable headers
- per-column filters
- drag-and-drop column reordering
- hide/show column controls
- Teams chat and call actions
- local-only execution with no hosting requirement

## What It Uses

- Electron for the desktop shell and local backend
- React + Vite for the renderer UI
- `@azure/msal-node` for desktop sign-in
- Microsoft Graph for users, presence, photos, manager data, and mailbox time zones

## Microsoft Entra Setup

Create a Microsoft Entra app registration for desktop sign-in.

1. Register a new application.
2. Set supported account types to match the organizations allowed to sign in.
3. Under `Authentication`, add the platform `Mobile and desktop applications`.
4. Add this redirect URI:

```text
http://localhost
```

5. Add these delegated Microsoft Graph permissions:

- `User.Read`
- `User.Read.All`
- `Presence.Read.All`
- `ProfilePhoto.Read.All`
- `MailboxSettings.Read`
- `AuditLog.Read.All`

6. Grant admin consent for the tenant.

## App Configuration

TeamSpy no longer asks end users for tenant or app-registration settings.

The authentication model is now:

- the app is built with a publisher-supplied Microsoft Entra public client ID
- the tenant is resolved from the user's Microsoft 365 login via the `organizations` authority

For distributed builds, set the value in:

```text
electron/publisher-config.ts
```

Required Entra setup:

- Supported account types should allow the organizational accounts you want to sign in
- Authentication platform: `Mobile and desktop applications`
- Redirect URI: `http://localhost`

## Run It

```bash
pnpm install
pnpm dev
```

## Build It

Desktop build artifacts:

```bash
pnpm build
```

Package a macOS app bundle:

```bash
pnpm dist:mac
```

Package Windows installers/artifacts:

```bash
pnpm dist:win
```

Notes:

- macOS packaging is best run on macOS
- Windows packaging is best run on Windows
- the app code itself is cross-platform; packaging still depends on the host build environment

## Behavior Notes

- `Last Seen` uses Entra sign-in activity when available. It is not true Teams last-active history.
- `Work Location` currently falls back to `officeLocation` because Microsoft Graph does not expose the dynamic Teams work-location signal in this flow.
- User photos and mailbox settings degrade gracefully when tenant policy or licenses do not allow access.
- Teams actions open Microsoft-supported deep links from the desktop app.
- The desktop shell uses a mac-native title bar on macOS and a standard native window frame on Windows.

## Project Layout

```text
electron/           Electron main process, preload bridge, auth, Graph calls
src/                React renderer
src/shared/         Shared desktop API and data types
```
