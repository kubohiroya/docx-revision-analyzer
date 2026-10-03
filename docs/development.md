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
[Usage](usage.md)). Running `npm link` exposes them globally as `docx-revision-chart` /
`docx-revision-flow` / `docx-ai-suspicion-score`.

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
  paste (recorded as a move) and by copy + paste + delete (for `docx-revision-flow`; sample output in
  `fixtures/flow-demo.svg`)
- `fixtures/snapshot-demo/sprint{1,2,3}/thesis.docx`: a thesis handed in at three sprints. Before each one, the
  teacher accepted all changes and handed it back; sprint 3 includes text written elsewhere and pasted in, and a
  paragraph typed with Track Changes off (for checking `docx-revision-snapshot`). The sample outputs
  `fixtures/snapshot-through.svg` (the chart across all submissions) and `fixtures/tamper-warning.svg` (sprint 3's
  chart, with the tampering warning) are made like this:

  ```bash
  for n in 1 2 3; do node dist/cli/snapshot.js /tmp/archive fixtures/snapshot-demo/sprint$n --lang en; done
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
the published package via `files`); `dist-bin/` is where Bun writes the
single-binary build.

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
│   │   ├── chart.ts       docx-revision-chart CLI
│   │   ├── flow.ts        docx-revision-flow CLI
│   │   ├── score.ts       docx-ai-suspicion-score CLI
│   │   ├── snapshot.ts    docx-revision-snapshot CLI (archiving and charting submissions over time)
│   │   ├── versions.ts    docx-revision-versions CLI (writing out the document at earlier points in time)
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
│   ├── build-all.mjs         Builds everything this environment can (npm run build)
│   ├── build-binary.sh       Builds single executables with Bun
│   ├── make-mac-droplet.sh   Builds the macOS Finder droplets (.app)
│   ├── makeFixtures.ts       Generates the test .docx files
│   ├── make-revision-states.mjs Generates the state diagrams
│   └── test-*.mjs            Tests (npm test)
├── desktop/               Desktop app (Electron)
├── examples/              Example highlight rules
├── fixtures/              Generated test .docx files, sample outputs and state diagrams
├── dist/                  Output of `npm run build:lib` (gitignored)
└── dist-bin/              Output of `npm run build:binary` (gitignored)
```

## Contributing

Issues and pull requests are welcome — please use
[GitHub Issues](https://github.com/kubohiroya/docx-revision-analyzer/issues)
for bug reports and feature requests.
