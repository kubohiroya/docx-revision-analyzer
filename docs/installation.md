# Installation

English | [日本語](./installation.ja.md) · [Back to the README](../README.md)

How to get and set up the npm package, the single executables and the drag-and-drop apps. To build from source, see the [developer guide](development.md).

**Contents**

- [From npm (Node.js 18+)](#from-npm-nodejs-18)
- [Single executables (no Node.js needed)](#single-executables-no-nodejs-needed)
- [Finder drag-and-drop app (macOS only)](#finder-drag-and-drop-app-macos-only)
- [Windows drag-and-drop (no wrapper app needed)](#windows-drag-and-drop-no-wrapper-app-needed)
- [Desktop app](#desktop-app)

## From npm (Node.js 18+)

```bash
npm install -g docx-revision-analyzer
```

This installs five commands: `docx-revision-chart`, `docx-revision-flow`, `docx-ai-suspicion-score`, `docx-revision-snapshot`, and `docx-revision-versions`.

## Single executables (no Node.js needed)

Download the file for your OS / CPU from [GitHub Releases](https://github.com/kubohiroya/docx-revision-analyzer/releases) and unzip it.

| File | Contents |
|---|---|
| `docx-revision-<chart\|flow>-<os>-<cpu>.zip`, `docx-ai-suspicion-score-<os>-<cpu>.zip` | Standalone executable for each tool (`linux-x64` / `linux-arm64` / `macos-x64` / `macos-arm64` / `windows-x64`) |
| `docx-revision-<chart\|flow>-macos-arm64.app.zip` | Drag-and-drop app for macOS (Apple Silicon): drop `.docx` files onto its icon in Finder and the SVG is written next to them |

The drag-and-drop apps are only ad-hoc signed, so the first time, right-click the app in Finder and choose "Open".
On Windows you can drop `.docx` files straight onto the `.exe`.

For building from source, making the single executables and drag-and-drop apps, the test fixtures, and the directory
layout, see the [developer guide](development.md).

## Finder drag-and-drop app (macOS only)

Beyond the command line, `docx-revision-chart` can be packaged as a small
macOS app (a "droplet") that you can drop one or more `.docx` files onto
directly in Finder — no Terminal needed. For each file dropped, it writes
the chart next to it, named `<filename>-<that file's last-modified time>.svg`
(e.g. `report-20260919-143000.svg`), and reports success or failure with a
native notification/dialog instead of console output.

```bash
npm run build:mac-app
```

This builds the Bun binary and wraps it into `dist-bin/docx-revision-chart.app`
using `osacompile` (an AppleScript compiler that ships with every Mac — no
extra install beyond Bun itself). Move or copy that `.app` anywhere convenient
(e.g. your Applications folder or the Dock) and drop `.docx` files onto its
icon.

`npm run build:mac-app:flow` builds the same kind of droplet for
`docx-revision-flow` (`dist-bin/docx-revision-flow.app`). It covers the
whole editing period with default settings and writes
`<file name>-flow-<that file's last-modified time>.svg`.

Why this needs a wrapper at all: macOS Finder only delivers dropped files to
proper application bundles (via an Apple Event), never directly as command-line
arguments to a bare, unbundled executable — so the raw Bun binary by itself
can't be a drop target. `scripts/make-mac-droplet.sh` generates a minimal
AppleScript application whose `on open` handler shells out to the bundled
binary with `--drop`, which is what produces the timestamped filename.

A few things worth knowing:

- **macOS only**, and the build script must be run in an actual Terminal on
  your Mac — it won't run inside a sandboxed/Linux environment.
- The app is **ad-hoc code-signed** (`codesign -s -`) by the build script,
  which is enough to run it locally on Apple Silicon. It's not notarized, so
  on first launch Gatekeeper may still warn that "the developer cannot be
  verified" — right-click the app and choose **Open** once to get past that.
- Once built, the `.app` is fully self-contained; it does not need Bun,
  Node.js, or this repository to keep working.
- This has been built and reasoned through against current documentation on
  the technique, but hasn't been verified end-to-end on an actual Mac by the
  author of this feature — please open an issue if dropping a file doesn't
  behave as described.

## Windows drag-and-drop (no wrapper app needed)

Windows doesn't need a wrapper app the way macOS does: Explorer launches a
`.exe` directly with the dropped file path(s) as plain command-line
arguments, so the Windows binary produced by `npm run build:binary:all`
(the `bun-windows-x64` target) can be dropped onto directly, with no
`osacompile`-style packaging step.

- Drop one **or several** `.docx` files onto `docx-revision-chart.exe` at
  once — Explorer passes them all as separate arguments to a single process
  invocation, and each is processed independently (a failure on one file
  doesn't stop the rest). This is exactly the multi-file support added to
  `docx-revision-chart` for this purpose (see [Usage](usage.md)).
- Use `--drop` the same way as on macOS, so each output is named
  `<filename>-<that file's last-modified time>.svg` instead of the plain
  `<filename>.svg` default.
- Console-subsystem executables launched by double-click or drag-drop on
  Windows open (and then immediately close) a console window when the
  process exits, so there's nothing to read there. To give some feedback
  without a terminal, `--drop` on Windows also pops up a native message box
  (via a bundled PowerShell call) summarizing which files succeeded or
  failed.
- Like the macOS droplet, **this has been implemented and reasoned through
  but not verified on actual Windows hardware** — please open an issue if
  it doesn't behave as described.

```powershell
# Cross-compile from macOS/Linux (or run natively on Windows with Bun installed)
npm run build:binary:all
# -> dist-bin/docx-revision-chart-windows-x64.exe (exact name depends on build-binary.sh)
```

## Desktop app

The desktop app (Electron), for people who aren't used to the command line, doesn't have an installer yet. To run it
from source, see [desktop/README.md](../desktop/README.md); to package it, see
[desktop/DISTRIBUTION.md](../desktop/DISTRIBUTION.md). Its features are described in [Usage](usage.md#desktop-app-preview).
