# Distributing the desktop app / デスクトップアプリの配布

## Overview / 概要

| | macOS | Windows |
|---|---|---|
| Artifacts / 配布物 | `.dmg` (universal), `.zip` (for auto-update) | NSIS installer (x64 / arm64), `.zip` |
| Signing / 署名 | Developer ID Application + Hardened Runtime | Code-signing certificate (OV/EV) |
| Notarization / 公証 | notarytool (App Store Connect API key) | — |
| Built by / ビルド | `.github/workflows/desktop-release.yml` on `macos-latest` | same, on `windows-latest` |

Signing and notarization run only when the secrets below are set; without them, unsigned builds are produced
(macOS users then need right-click → Open the first time, Windows users see a SmartScreen warning).

署名・公証は下の Secrets があるときだけ行います。無ければ署名なしの配布物を作ります (macOS では初回に右クリック→開く、
Windows では SmartScreen の警告が必要になります)。

## Secrets / 設定する Secrets

Repository → Settings → Secrets and variables → Actions:

| Name | What / 内容 |
|---|---|
| `MAC_CSC_LINK` | "Developer ID Application" certificate exported as `.p12`, base64-encoded (`base64 -i cert.p12 \| pbcopy`) |
| `MAC_CSC_KEY_PASSWORD` | Password of that `.p12` |
| `APPLE_API_KEY_P8` | Contents of the App Store Connect API key (`AuthKey_XXXX.p8`, role "Developer") for notarization |
| `APPLE_API_KEY_ID` | Key ID of that key |
| `APPLE_API_ISSUER` | Issuer ID (App Store Connect → Users and Access → Integrations) |
| `WIN_CSC_LINK` | Windows code-signing certificate (`.pfx`), base64-encoded |
| `WIN_CSC_KEY_PASSWORD` | Password of that `.pfx` |

And a repository **variable** `WIN_PUBLISHER_NAME`: the certificate's subject CN exactly (e.g. `Hiroya Kubo`). When
set, auto-update on Windows refuses installers not signed by that publisher.

Notes:
- An EV certificate on a hardware token (or a cloud HSM such as Azure Trusted Signing) can't be exported as `.pfx`;
  in that case change the Windows signing step (electron-builder supports `win.azureSignOptions` / a custom `sign` hook).
- Apple ID + app-specific password also works instead of an API key (`APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD`,
  `APPLE_TEAM_ID`), but the API key is recommended for CI.

## Releasing / リリースの手順

1. Bump `desktop/package.json` `version` (e.g. `0.2.0`) and merge it to `main`.
2. Tag and push: `git tag desktop-v0.2.0 && git push origin desktop-v0.2.0`.
3. The workflow builds both platforms (each job only builds: `--publish never`), then one `release` job creates a
   single **draft** release for the tag and attaches every artifact and the `latest*.yml` files (their file names
   rewritten to absolute download URLs of this release by `desktop/scripts/update-feed.mjs`).
   Check it (download, open on a clean machine), write the notes, and publish it. **Untick "Set as the latest
   release"**: the repository's latest release should stay the library's `vX.Y.Z`.
4. Publishing it runs `.github/workflows/desktop-update-feed.yml`, which copies its `latest*.yml` into the
   **`desktop-latest`** pre-release (created the first time). That is where installed apps look for updates.

1. `desktop/package.json` の `version` を上げて `main` に入れる。
2. タグ `desktop-vX.Y.Z` を push する。
3. ワークフローが両方の OS の配布物を作り (各 OS のジョブは作るだけ)、最後のジョブがそのタグの**下書き**のリリースを
   1つ作って、配布物と `latest*.yml` (配布物の名前をこのリリースの絶対 URL に書き換えたもの) を添付する。中身を確かめ、
   リリースノートを書いて公開する。**「Set as the latest release」は外す** (リポジトリの Latest はライブラリの `vX.Y.Z` のまま)。
4. 公開すると `desktop-update-feed.yml` が、その `latest*.yml` を自動更新の参照先のプレリリース **`desktop-latest`** に置く。

Running the workflow manually (Actions → Desktop release → Run workflow) builds without publishing and keeps the
artifacts on the run, which is useful for testing signing.

Library releases (npm) are separate and keep using plain `vX.Y.Z` tags.

## Auto-update / 自動更新

- Off by default. Users turn it on in Settings ("Check GitHub Releases for a new version at startup").
  既定では無効。利用者が設定で有効にしたときだけ確認する。
- `electron-updater` in the main process reads `latest-mac.yml` / `latest.yml` from the **`desktop-latest`**
  pre-release (`https://github.com/kubohiroya/docx-revision-analyzer/releases/download/desktop-latest/…`, the
  `generic` provider in `electron-builder.config.cjs`), downloads the update from the `desktop-vX.Y.Z` release those
  files point to, and installs it on quit. It sends no document or analysis data; the app's own window remains
  blocked from the network.
- Why not the `github` provider: it reads the repository's **latest** release, which is the library's `vX.Y.Z`
  (no `latest*.yml` there), so updates would never be found. A pre-release is never the latest release, so
  `desktop-latest` keeps the two apart.
- Artifact names contain no spaces (`artifactName` in the config): GitHub replaces spaces with dots on upload, which
  would no longer match the names in `latest*.yml`.
- Verification: the download's SHA-512 must match the release's `latest*.yml`; on macOS, Squirrel.Mac also
  requires the update to be signed with the same Developer ID as the running app (so auto-update only works for
  signed builds); on Windows, with `WIN_PUBLISHER_NAME` set, the installer's signature must be from that publisher.
- Only published releases are offered: `desktop-latest` is updated when a `desktop-v*` release is published, not
  while it is a draft. To re-point it (e.g. roll back), run "Desktop update feed" manually with a published tag.

## Opening files from OneDrive / SharePoint / OneDrive・SharePoint から開く

"Open from URL" signs the user in to Microsoft and reads the file with Microsoft Graph. This needs an application
registered once in Microsoft Entra ID (free; no client secret, since the app is a public client):

1. [Entra admin center](https://entra.microsoft.com) → App registrations → New registration
   - Name: `Docx Revision Analyzer`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
     (for both university/company and personal OneDrive; choose "any organizational directory" only to exclude
     personal accounts)
   - Redirect URI: platform **Public client/native (mobile & desktop)**, `http://localhost`
2. API permissions → Add → Microsoft Graph → Delegated → `Files.Read.All` (and `offline_access`, `User.Read`).
   These normally don't need admin consent, but some organizations block user consent; then an administrator has to
   grant consent (or users see "Need admin approval").
3. Copy the **Application (client) ID** into `desktop/microsoft.json` (see `microsoft.example.json`); it is bundled
   into the packaged app. A client ID is not a secret.

Users (or organizations with their own registration) can also enter a client ID in Settings → Microsoft account →
Application (client) ID; `DRA_MS_CLIENT_ID` overrides both for development.

What the app does: the sign-in runs in the user's browser (authorization code + PKCE with a localhost redirect); the
token cache is stored in the app's data folder, encrypted with the OS keychain (Electron `safeStorage`). The main
process talks only to `login.microsoftonline.com`, `graph.microsoft.com` and the download URL Graph returns; the
document is read into memory and never stored or sent anywhere. The app window itself remains blocked from the
network.

「URL から開く」は Microsoft にサインインして Microsoft Graph でファイルを読みます。Microsoft Entra ID へのアプリ登録
(無料、シークレット不要) が一度必要です。登録したアプリのクライアント ID を `desktop/microsoft.json` に書くと配布物に
組み込まれます。組織によっては、管理者の同意が必要です。

## Local builds / 手元でのビルド

```bash
cd desktop
CSC_IDENTITY_AUTO_DISCOVERY=false npm run dist -- --mac dir --arm64   # unsigned .app in ../release/desktop/mac/
```
