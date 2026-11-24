# Installation

English | [日本語](./installation.ja.md) · [Back to the README](../README.md)

How to get and set up the npm package, the standalone CLI and the drag-and-drop apps. To build from source, see the [developer guide](development.md).

Get the unified CLI **v2.0.0** as a standalone executable from [GitHub Releases](https://github.com/kubohiroya/docx-revision-analyzer/releases/tag/v2.0.0). npm publication is separate. For source builds, see the [developer guide](development.md).

**Contents**

- [From npm (Node.js 18+)](#from-npm-nodejs-18)
- [Standalone CLI (no Node.js needed)](#standalone-cli-no-nodejs-needed)
- [Finder drag-and-drop app (macOS only)](#finder-drag-and-drop-app-macos-only)
- [Windows](#windows)
- [Migrating from the separate commands](#migrating-from-the-separate-commands)
- [Desktop app](#desktop-app)

## From npm (Node.js 18+)

The published npm v1.2.0 uses the separate commands. npm publication of this release is separate; use the standalone binaries below for the unified CLI.

```bash
npm install -g docx-revision-analyzer@1.2.0
```

That legacy package provides commands such as `docx-revision-chart`. See the [migration table](#migrating-from-the-separate-commands) when moving to v2.0.0.

## Standalone CLI (no Node.js needed)

Download **one** `docx-revision-analyzer-<os>-<cpu>.zip` for your OS / CPU from [CLI v2.0.0](https://github.com/kubohiroya/docx-revision-analyzer/releases/tag/v2.0.0), then unzip it:

| Platform | File to choose | Executable inside |
|---|---|---|
| Linux x64 / ARM64 | `docx-revision-analyzer-linux-x64.zip` / `-linux-arm64.zip` | `docx-revision-analyzer` |
| Mac Intel / Apple Silicon | `docx-revision-analyzer-macos-x64.zip` / `-macos-arm64.zip` | `docx-revision-analyzer` |
| Windows x64 / ARM64 | `docx-revision-analyzer-windows-x64.zip` / `-windows-arm64.zip` | `docx-revision-analyzer.exe` |

Each ZIP contains one executable with all five subcommands, plus LICENSE. On macOS/Linux, use `chmod +x docx-revision-analyzer` if needed, then run `./docx-revision-analyzer flow report.docx`. Windows examples are below.

Compare the ZIP's SHA-256 with its line in [SHA256SUMS.txt](https://github.com/kubohiroya/docx-revision-analyzer/releases/download/v2.0.0/SHA256SUMS.txt). These builds are not Developer ID / Authenticode signed or notarized, so your OS may warn on first launch.

The separate macOS droplets keep their names, `docx-revision-chart.app` and `docx-revision-flow.app`. They call the shared CLI with `chart` or `flow` and remain self-contained. They are ad-hoc signed; on first launch, right-click the app in Finder and choose "Open".

For building from source, making the standalone CLI and drag-and-drop apps, the test fixtures, and the directory
layout, see the [developer guide](development.md).

## Finder drag-and-drop app (macOS only)

Beyond the command line, `docx-revision-analyzer chart` can be packaged as a small
macOS app (a "droplet") that you can drop one or more `.docx` files onto
directly in Finder — no Terminal needed. For each file dropped, it writes
the chart next to it, named `<filename>-<that file's last-modified time>.svg`
(e.g. `report-20260919-143000.svg`), and reports success or failure with a
native notification/dialog instead of console output.

```bash
npm run build:droplet
```

This builds the shared Bun CLI and wraps it into
`release/droplet/mac-<arch>/docx-revision-chart.app` (`<arch>` is `arm64` or `x64`)
using `osacompile` (an AppleScript compiler that ships with every Mac — no
extra install beyond Bun itself). Move or copy that `.app` anywhere convenient
(e.g. your Applications folder or the Dock) and drop `.docx` files onto its
icon.

The same command also builds a droplet for
`docx-revision-analyzer flow` (`release/droplet/mac-<arch>/docx-revision-flow.app`). It covers the
whole editing period with default settings and writes
`<file name>-flow-<that file's last-modified time>.svg`.

Why this needs a wrapper at all: macOS Finder only delivers dropped files to
proper application bundles (via an Apple Event), never directly as command-line
arguments to a bare, unbundled executable — so the raw Bun binary by itself
can't be a drop target. `scripts/make-mac-droplet.sh` generates a minimal
AppleScript application whose `on open` handler shells out to the bundled
binary with `chart --drop` or `flow --drop`, which is what produces the timestamped filename.

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

## Windows

Build from source with `npm run build:cli` on Windows, or cross-compile with `npm run build:cli:all-platforms`. The Windows output is `release/cli/win-<arch>/docx-revision-analyzer.exe` (`x64` / `arm64`).

```powershell
.\docx-revision-analyzer.exe --help
.\docx-revision-analyzer.exe chart .\report.docx
.\docx-revision-analyzer.exe flow .\report.docx .\report2.docx
```

The CLI requires a subcommand, so dropping a document directly onto the `.exe` does not select an analysis. For a graphical drag-and-drop interface, use the [desktop app](#desktop-app). When invoking `chart` or `flow` with `--drop` from a launcher, the timestamped output names and Windows result dialogs remain available.

## Migrating from the separate commands

Both npm and the standalone build expose one command. Replace existing invocations as follows; subcommand options and output formats remain unchanged. Old executable names and alias links are not generated.

| Previous command | New invocation |
|---|---|
| `docx-revision-chart` | `docx-revision-analyzer chart` |
| `docx-revision-flow` | `docx-revision-analyzer flow` |
| `docx-ai-suspicion-score` | `docx-revision-analyzer score` |
| `docx-revision-snapshot` | `docx-revision-analyzer snapshot` |
| `docx-revision-versions` | `docx-revision-analyzer versions` |

Existing settings filenames (such as `docx-revision-chart.yml`) and the analysis JSON's `tool` identifiers remain unchanged. The build option `--cli=...` has been removed; every build includes all five subcommands.

## Desktop app

The desktop app (Electron, preview) has installers in [desktop-v0.2.0](https://github.com/kubohiroya/docx-revision-analyzer/releases/tag/desktop-v0.2.0). Choose the `.dmg` for Mac or `Setup-0.2.0.exe` for Windows. These are separate from the CLI.
To run from source, see [desktop/README.md](../desktop/README.md); to package it, see [desktop/DISTRIBUTION.md](../desktop/DISTRIBUTION.md). Its features are described in [Usage](usage.md#desktop-app-preview).
