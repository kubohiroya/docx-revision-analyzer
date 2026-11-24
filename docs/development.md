# Developer guide

English | [日本語](./development.ja.md) · [Back to the README](../README.md)

Building from source, making the single executables and drag-and-drop apps, publishing, test fixtures, and the directory layout.

**Contents**

- [Setup](#setup)
- [Building single-binary executables with Bun](#building-single-binary-executables-with-bun)
- [Publishing to npm (maintainers)](#publishing-to-npm-maintainers)
- [Attaching single executables to a GitHub Release (maintainers)](#attaching-single-executables-to-a-github-release-maintainers)
- [Test fixtures](#test-fixtures)
- [Directory layout](#directory-layout)
- [Contributing](#contributing)

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

All build output goes to two places: `dist/` (the npm package) and `release/` (everything you distribute),
grouped by kind and then by platform. `<os>` is `mac` / `win` / `linux` and `<arch>` is `x64` / `arm64`
(the names electron-builder uses).

| Step | Output | When |
|---|---|---|
| `lib` | Library and CLIs (JavaScript) → `dist/` | always |
| `cli` | One CLI `docx-revision-analyzer` → `release/cli/<os>-<arch>/` | if [Bun](https://bun.sh) is installed |
| `droplet` | macOS droplets `docx-revision-chart.app` / `docx-revision-flow.app` → `release/droplet/mac-<arch>/` | on macOS with Bun |
| `desktop` | Desktop app for this OS/CPU (unsigned) → `release/desktop/<os>/` (installs `desktop/` dependencies the first time) | always |

```
dist/                         npm package (tsc output)
release/
├── cli/
│   ├── mac-arm64/            docx-revision-analyzer
│   ├── mac-x64/
│   ├── linux-x64/
│   ├── linux-arm64/
│   ├── win-x64/              docx-revision-analyzer.exe
│   └── win-arm64/            docx-revision-analyzer.exe
├── droplet/
│   └── mac-arm64/            docx-revision-chart.app, docx-revision-flow.app
└── desktop/
    ├── mac/                  *.dmg, *.zip, mac-arm64/ or mac-universal/ (the .app)
    └── win/                  *Setup*.exe, *.zip, win-unpacked/, win-arm64-unpacked/
```

Steps whose tools are missing are skipped with the reason. Each step also has its own script:

| Script | Does |
|---|---|
| `npm run build` | every step this machine can do |
| `npm run build:lib` | only compile the library (what `prepublishOnly` and CI use) |
| `npm run build:cli` | one CLI executable for this OS/CPU |
| `npm run build:cli:all-platforms` | one CLI for linux-x64, linux-arm64, mac-x64, mac-arm64, win-x64, win-arm64 |
| `npm run build:droplet` | macOS droplets |
| `npm run build:desktop` | desktop app for this OS/CPU (unpacked, no installer) |
| `npm run build:desktop:installers` | desktop app installers (dmg/zip/NSIS; slow) |
| `npm run clean` | delete `dist/` and `release/` |

Options (after `--`, for `npm run build` and the scripts above):

```bash
npm run build -- --skip=desktop          # everything except the desktop app
npm run build -- --only=lib,cli          # only these steps
npm run build -- --all-platforms         # cross-build the executables for every platform
npm run build -- --installers            # also build the desktop installers
```

Requires Node.js 18+ (developed and tested on Node.js 22). After building,
run `node dist/cli/index.js flow report.docx` directly (see [Usage](usage.md)).
`npm link` exposes the single public command `docx-revision-analyzer`. No aliases or alias links are provided.
See [Migration](installation.md#migrating-from-the-separate-commands) for the old command names.

For CLI changes, run `npm run test:cli`. It exercises all five subcommands, help, language selection,
settings, and argument boundaries. To run the same checks against a compiled executable on this machine:

```bash
DRA_CLI_BINARY="$PWD/release/cli/mac-arm64/docx-revision-analyzer" npm run test:cli
```

Replace the path with the executable built for your host OS / CPU.

## Building single-binary executables with Bun

If [Bun](https://bun.sh) is installed, `docx-revision-analyzer` can be compiled
into a single executable that needs neither Node.js nor `node_modules` to run
(JSZip, fast-xml-parser, commander, etc. are all bundled into the binary).
Bun compiles the TypeScript source (`src/cli/index.ts`) directly, so there's
no need to run `npm run build:lib` first.

```bash
# Install Bun if you don't have it
curl -fsSL https://bun.sh/install | bash

# Build one CLI containing all five subcommands for the current OS/CPU -> release/cli/<os>-<arch>/
npm run build:cli

# Cross-build for the major OS/CPU combinations -> release/cli/{linux,mac,win}-{x64,arm64}/
npm run build:cli:all-platforms
```

The resulting binary can be run and distributed as-is (no Node.js install
needed):

```bash
./release/cli/mac-arm64/docx-revision-analyzer chart fixtures/multi-session.docx -p 12 -o out.svg
```

> Binaries cross-compiled for other OSes (`--all-platforms`) can't be smoke-tested on
> the machine that built them — verify them on the target OS before
> distributing.

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
npm run build:cli:all-platforms     # one CLI executable per supported OS/CPU
npm run build:droplet               # macOS only
```

Package the single CLI per OS/CPU as `docx-revision-analyzer-<os>-<cpu>.zip`. Archive OS names are `linux` / `macos` / `windows` (build folders use `linux` / `mac` / `win`). Zip the macOS `.app` bundles separately. Attach these to the next release; do not replace published assets.

### Automated CLI release

The `CLI release` workflow validates six native hosts: Linux/macOS/Windows × x64/ARM64. On a `vX.Y.Z` tag matching `package.json`, it runs the full root suite, tests each compiled CLI, validates ZIP architectures and checksums, checks the desktop module, and creates a **draft** release. It uses Node.js 22 and Bun 1.4.2. Workflow/package-script PRs also exercise native builds and packaging.

1. Update the version, release notes in `docs/releases/vX.Y.Z.md`, and download links through a PR.
2. After CI and merge, tag that merge commit and push the tag.
3. Wait for all jobs, inspect the draft's six CLI ZIPs, two Apple Silicon droplets and SHA256SUMS.txt, then publish the approved release as Latest.
4. Confirm the public links. npm publication remains separate.

`node scripts/package-cli.mjs` packages the host CLI; add `--include-droplets` on mac-arm64 after building droplets. `node scripts/verify-cli-release.mjs release/assets --complete` verifies the complete asset set and writes SHA256SUMS.txt. Each CLI ZIP contains one executable and LICENSE.

## Test fixtures

`scripts/makeFixtures.ts` generates synthetic `.docx` files with Track
Changes history, so both tools can be exercised without a real
revision-tracked Word document on hand.

```bash
npm run fixtures
# internally: ts-node scripts/makeFixtures.ts (runs the TS directly, no build needed)
```

- `fixtures/chart-demo.docx`: three writing sessions over two days (4.5-hour
  and 17-hour idle gaps) with one bulk paste and sporadic deletions of a draft (the `docx-revision-analyzer chart`
  example in [Usage](usage.md))
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
  paste (recorded as a move) and by copy + paste + delete (for `docx-revision-analyzer flow`; sample output in
  `fixtures/flow-demo.svg`)
- `fixtures/snapshot-demo/sprint{1,2,3}/thesis.docx`: a thesis handed in at three sprints. Before each one, the
  teacher accepted all changes and handed it back; sprint 3 includes text written elsewhere and pasted in, and a
  paragraph typed with Track Changes off (for checking `docx-revision-analyzer snapshot`). The sample outputs
  `fixtures/snapshot-through.svg` (the chart across all submissions) and `fixtures/tamper-warning.svg` (sprint 3's
  chart, with the tampering warning) are made like this:

  ```bash
  for n in 1 2 3; do node dist/cli/index.js snapshot /tmp/archive fixtures/snapshot-demo/sprint$n --lang en; done
  # /tmp/archive/thesis/thesis-through-tampered.svg -> fixtures/snapshot-through.svg
  # /tmp/archive/thesis/thesis-s03-tampered.svg     -> fixtures/tamper-warning.svg  (.ja.svg with --lang ja)
  ```

Sample output for each is bundled under `fixtures/*.svg` (`*.png` versions are included for quick visual
inspection). The files with plain names were rendered in English (`--lang en`) and the `*.ja.svg` / `*.ja.png`
ones in Japanese (`--lang ja`); the English and Japanese READMEs and documents use them respectively.
`fixtures/revision-states.svg`, `revision-states-onedrive.svg` and their `.ja.svg` / `.png` versions are the hand-drawn state diagrams in the README and the operations guide, not
tool output; regenerate them with `npm run diagrams` (`scripts/make-revision-states.mjs`; the PNGs need Inkscape) rather than `npm run fixtures`.

## Directory layout

Follows a standard npm package layout. `src/` holds the TypeScript sources;
`dist/` is the build output (JS + type declarations, gitignored, included in
the published package via `files`); `release/` holds everything built for
distribution (single executables, droplets and the desktop app; see above).

```
docx-revision-analyzer/
├── package.json          Package manifest (bin, files, license: MIT, etc.)
├── tsconfig.json
├── LICENSE                Full text of the MIT license
├── README.md / README.ja.md   Overview (English / Japanese)
├── docs/                  Documents (installation, usage, operations-guide, advanced, library, development; each .ja.md is the Japanese version)
├── src/
│   ├── index.ts           Entry point when importing as a library (Node.js)
│   ├── core.ts            The core without file system access (for browsers; docx-revision-analyzer/core)
│   ├── cli/
│   │   ├── index.ts       public CLI entry point (five subcommands)
│   │   ├── chart.ts       chart subcommand
│   │   ├── flow.ts        flow subcommand
│   │   ├── score.ts       score subcommand
│   │   ├── snapshot.ts    snapshot subcommand (archiving and charting submissions over time)
│   │   ├── versions.ts    versions subcommand (writing out the document at earlier points in time)
│   │   ├── config.ts      Settings files (<tool>.yml)
│   │   └── common.ts      Shared CLI logic (--preserve-history, --lock, --template, multiple files and result output)
│   ├── lib/
│   │   ├── docxRevisions.ts  Extracts tracked-change events from a .docx
│   │   ├── docxLayout.ts     Splits the body into paragraphs and tracked fragments; reads page size etc.
│   │   ├── revisionPositions.ts Where each edit happened (in the final document and at edit time)
│   │   ├── insertionKinds.ts Classifies insertions (bulk, fine-grained, move/reorder)
│   │   ├── insertionWindows.ts Insertion windows and their features (--json)
│   │   ├── insertionRules.ts Highlight rules (--rules)
│   │   ├── classifiers.ts    The classifier pipeline
│   │   ├── categories.ts     Color categories of the figures
│   │   ├── figureTargets.ts  Keys for parts of the figures and annotations (--annotations)
│   │   ├── timeBuckets.ts    Time-bucket aggregation (chart)
│   │   ├── sessions.ts       Splitting into sessions at idle gaps
│   │   ├── svgChart.ts       SVG chart rendering (single and per-session)
│   │   ├── flow.ts           flow: rebuilding the document at session boundaries, schematic pages, classifying changes
│   │   ├── flowSvg.ts        flow: SVG rendering
│   │   ├── suspicionScore.ts AI-misuse suspicion scoring
│   │   ├── historySettings.ts Checking and rewriting the Track Changes / personal-information settings (--preserve-history)
│   │   ├── integrity.ts      Integrity notes (integrity in --json)
│   │   ├── tamperEvidence.ts Traces of tampering and the warning in figures (-tampered.svg)
│   │   ├── trackLock.ts      Locking Track Changes (--lock; Word-compatible password hash)
│   │   ├── snapshots.ts      Analysis across snapshots (new changes per submission, without duplicates)
│   │   ├── versions.ts       Rebuilding the document (.docx) at a given time from its tracked changes
│   │   ├── filenames.ts      Output file names and dates in titles
│   │   ├── input.ts          The input type (.docx bytes)
│   │   └── i18n.ts           Display language detection and messages (English / Japanese)
│   └── node/              Node.js only (uses files, crypto and OS features)
│       ├── files.ts          Reading files; --drop output names
│       ├── historyFile.ts    Applying settings changes to files (backups, checking the file isn't open)
│       ├── lockFile.ts       Applying the template lock to files
│       ├── snapshotArchive.ts Archiving snapshots under serial numbers
│       └── locale.ts         Reading the macOS language setting
├── scripts/
│   ├── build-all.mjs         Builds everything this environment can (npm run build and build:*)
│   ├── make-mac-droplet.sh   Builds the macOS Finder droplets (.app)
│   ├── makeFixtures.ts       Generates the test .docx files
│   ├── make-revision-states.mjs Generates the state diagrams
│   └── test-*.mjs            Tests (npm test)
├── desktop/               Desktop app (Electron)
├── examples/              Example highlight rules
├── fixtures/              Generated test .docx files, sample outputs and state diagrams
├── dist/                  Output of `npm run build:lib` (gitignored)
└── release/               Output of `npm run build` for distribution (gitignored)
```

## Contributing

Issues and pull requests are welcome — please use
[GitHub Issues](https://github.com/kubohiroya/docx-revision-analyzer/issues)
for bug reports and feature requests.
