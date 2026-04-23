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

Create a **single-tenant** app registration in Microsoft Entra ID.

1. Register a new application.
2. Set supported account types to `Accounts in this organizational directory only`.
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

Record:

- Application (client) ID
- Directory (tenant) ID

## App Configuration

TeamSpy no longer requires manual `.env` setup for normal use.

When the app starts, open **Settings** and enter:

- your Microsoft Entra application (client) ID
- your tenant ID

TeamSpy stores those values in its local app profile on the machine and uses them for future launches.

Optional developer fallback:

- `.env.local`
- `.env`

Those are still supported if you want to preseed settings during development, but the intended user workflow is in-app configuration.

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
