# Installing and building docx-revision-analyzer

English | [日本語](./INSTALL.ja.md)

For usage, see [README.md](./README.md). If you only want to use the published package or executables, the
"Installation" section of the README is enough; this document covers building from source and packaging.

---

## Setup

```bash
git clone https://github.com/kubohiroya/docx-revision-analyzer.git
cd docx-revision-analyzer
npm install
npm run build
```

The project uses **npm** (`package-lock.json`, `"packageManager": "npm@…"`; CI runs `npm ci`). pnpm also works for
local builds (`pnpm install`, `pnpm run build`), but don't commit `pnpm-lock.yaml` (it's gitignored);
`pnpm-workspace.yaml` only allows esbuild's install script for pnpm.

`npm run build` (`pnpm run build`) builds **everything this machine can build** in one go and prints a summary:

| Step | Output | When |
|---|---|---|
| `lib` | Library and CLIs → `dist/` | always |
| `binaries` | Single executables `docx-revision-chart` / `docx-revision-flow` / `docx-ai-suspicion-score` → `dist-bin/` | if [Bun](https://bun.sh) is installed |
| `droplets` | macOS droplets `docx-revision-chart.app` / `docx-revision-flow.app` → `dist-bin/` | on macOS with Bun |
| `desktop` | Desktop app for this OS/CPU (unsigned) → `desktop/release/` (installs `desktop/` dependencies the first time) | always |

Steps whose tools are missing are skipped with the reason. Options (after `--`):

```bash
npm run build -- --skip=desktop          # everything except the desktop app
npm run build -- --only=lib,binaries     # only these steps
npm run build -- --all-platforms         # cross-build the executables for Linux/macOS/Windows × x64/arm64
npm run build -- --installers            # also build the desktop installers (dmg/zip/NSIS; slow)
```

`npm run build:lib` only compiles the library (what `prepublishOnly` and CI use). The individual commands below
(`build:binary:flow`, `build:mac-app:flow`, …) still work.

Requires Node.js 18+ (developed and tested on Node.js 22). After building,
run `dist/cli/chart.js` / `dist/cli/flow.js` / `dist/cli/score.js` directly with `node` (see
[README.md](./README.md) for usage). Running `npm link` exposes them globally as `docx-revision-chart` /
`docx-revision-flow` / `docx-ai-suspicion-score`.

---

## Building single-binary executables with Bun

If [Bun](https://bun.sh) is installed, `docx-revision-chart` can be compiled
into a single executable that needs neither Node.js nor `node_modules` to run
(JSZip, fast-xml-parser, commander, etc. are all bundled into the binary).
Bun compiles the TypeScript source (`src/cli/chart.ts`) directly, so there's
no need to run `npm run build:lib` first.

```bash
# Install Bun if you don't have it
curl -fsSL https://bun.sh/install | bash

# Build one binary for the current OS/CPU -> dist-bin/docx-revision-chart
npm run build:binary

# Cross-build for the major OS/CPU combinations (Linux/macOS/Windows × x64/arm64)
npm run build:binary:all
```

The resulting binary can be run and distributed as-is (no Node.js install
needed):

```bash
./dist-bin/docx-revision-chart fixtures/multi-session.docx -p 12 -o out.svg
```

This wraps `scripts/build-binary.sh` under the hood. To build a binary for
`docx-ai-suspicion-score` instead, run `npm run build:binary:score` or
`bash scripts/build-binary.sh score` (add `--all` for the same cross-build).
For `docx-revision-flow`, use `npm run build:binary:flow`
(`bash scripts/build-binary.sh flow`).

> Binaries cross-compiled for other OSes (`--all`) can't be smoke-tested on
> the machine that built them — verify them on the target OS before
> distributing.

---

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

---

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
  `docx-revision-chart` for this purpose (see the CLI usage in [README.md](./README.md)).
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

---

## Publishing to npm (maintainers)

```bash
npm version <patch|minor|major>   # bump the version in package.json and tag it
npm publish --dry-run              # check what gets published (prepublishOnly rebuilds dist/)
npm publish
```

Only `dist/`, `README.md` / `README.ja.md`, and `LICENSE` are published (the `files` field in `package.json`).
`prepublishOnly` deletes `dist/` and rebuilds it, so stale build output can't slip in.

## Attaching single executables to a GitHub Release (maintainers)

```bash
npm run build:binary:all            # docx-revision-chart
bash scripts/build-binary.sh flow --all
bash scripts/build-binary.sh score --all
npm run build:mac-app && npm run build:mac-app:flow   # macOS only
```

Zip the files in `dist-bin/` and attach them to the release.

---

## Test fixtures

`scripts/makeFixtures.ts` generates synthetic `.docx` files with Track
Changes history, so both tools can be exercised without a real
revision-tracked Word document on hand.

```bash
npm run fixtures
# internally: ts-node scripts/makeFixtures.ts (runs the TS directly, no build needed)
```

- `fixtures/chart-demo.docx`: three writing sessions over two days (4.5-hour
  and 17-hour idle gaps) with one bulk paste and sporadic deletions of a draft (the `docx-revision-chart`
  example in this README)
- `fixtures/natural-writing.docx`: simulates ~37 minutes of gradual, organic
  typing (score: 0 / low)
- `fixtures/suspicious-paste.docx`: simulates a bit of typing, then 295
  characters inserted in a single one-second burst — i.e. a simulated
  paste of externally-authored text (score: 93 / very_high)
- `fixtures/multi-session.docx`: simulates writing spread across 3 days, with
  29-hour and 20.5-hour idle gaps in between (for exercising the `-p` option)
- `fixtures/flow-demo.docx`: a multi-paragraph document with headings and a
  figure, edited in three sessions — typing by hand; pasting three paragraphs
  at once and touching them up; replacing and deleting paragraphs, making small
  fixes and adding a figure; and a fourth session reordering paragraphs by cut +
  paste (recorded as a move) and by copy + paste + delete (for `docx-revision-flow`; sample output in
  `fixtures/flow-demo.svg`)

Sample output for each is bundled under `fixtures/*.svg` (`*.png` versions are included for quick visual
inspection). The files with plain names were rendered in English (`--lang en`) and the `*.ja.svg` / `*.ja.png`
ones in Japanese (`--lang ja`); README.md and README.ja.md use them respectively.
`fixtures/revision-states.svg`, `revision-states-onedrive.svg` and their `.ja.svg` / `.png` versions are the hand-drawn state diagrams in the README, not
tool output; regenerate them with `npm run diagrams` (`scripts/make-revision-states.mjs`; the PNGs need Inkscape) rather than `npm run fixtures`.

---

## Directory layout

Follows a standard npm package layout. `src/` holds the TypeScript sources;
`dist/` is the build output (JS + type declarations, gitignored, included in
the published package via `files`); `dist-bin/` is where Bun writes the
single-binary build.

```
docx-revision-analyzer/
├── package.json          Package manifest (bin, files, license: MIT, etc.)
├── tsconfig.json
├── LICENSE                Full text of the MIT license
├── README.md              English manual (this file)
├── README.ja.md           Japanese manual
├── .gitignore
├── src/
│   ├── index.ts           Library entry point for programmatic use
│   ├── cli/
│   │   ├── chart.ts       docx-revision-chart CLI
│   │   ├── flow.ts        docx-revision-flow CLI
│   │   ├── config.ts      Loads the settings file (<tool name>.yml)
│   │   ├── common.ts      Shared CLI code (--preserve-history prompts, multi-file processing and output)
│   │   └── score.ts       docx-ai-suspicion-score CLI
│   └── lib/
│       ├── docxRevisions.ts  Shared library: extracts revision events from a .docx
│       ├── i18n.ts           Display-language detection and messages (English / Japanese)
│       ├── insertionKinds.ts Classifies insertions (bulk / fine-grained / moved); shared by chart and flow
│       ├── historySettings.ts Checks/rewrites the Track Changes and personal-information settings (--preserve-history)
│       ├── timeBuckets.ts    Aggregates events into time buckets (for the chart)
│       ├── sessions.ts       Splits events into sessions by idle gap, for -p
│       ├── svgChart.ts       SVG rendering (single-chart and session-split variants)
│       ├── docxLayout.ts     For flow: splits the body into paragraphs and revision-tagged pieces; reads page setup
│       ├── flow.ts           For flow: start/end document reconstruction, schematic pagination, paragraph intensities and change classification
│       ├── flowSvg.ts        For flow: SVG rendering
│       ├── suspicionScore.ts The AI-misuse suspicion score algorithm
│       └── filenames.ts      For --drop: builds the last-modified-time-based output filename
├── scripts/
│   ├── build-binary.sh       Bun single-binary build script
│   ├── make-mac-droplet.sh   Builds the macOS Finder droplet (.app)
│   └── makeFixtures.ts       Generates the synthetic test .docx files
├── fixtures/              Pre-generated test .docx files and sample outputs
├── dist/                  Output of `npm run build:lib` (gitignored)
└── dist-bin/              Output of `npm run build:binary` (gitignored)
```
