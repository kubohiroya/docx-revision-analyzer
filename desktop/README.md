# Docx Revision Analyzer (desktop app)

A desktop app for looking back on how a Word document (.docx) was written: open or drop a file and see the
revision chart, the edit flow, and the highlighted insertions. Everything is analyzed locally — the document and
its text are never sent anywhere (the app blocks all network access).

デスクトップ版です。.docx を開く (またはドロップする) だけで、編集履歴のチャート・編集フロー・ハイライトの一覧を表示します。
解析はすべてこのコンピュータの中で行い、文書や本文をどこにも送りません (アプリは外部への通信をすべて遮断します)。

## Features / 機能 (MVP)

- Open from a OneDrive / SharePoint URL ("Open from URL…"): paste the link from Share → Copy link; recently opened
  documents are listed for one-click reopening. Signs in to Microsoft in your browser the first time
  (setup: [DISTRIBUTION.md](DISTRIBUTION.md)). / OneDrive・SharePoint の URL から開く (最近開いた文書の一覧付き)
- Batch: drop a folder (SVGs next to every .docx), or paste a SharePoint folder link such as the Teams assignment
  submissions (choose an output folder; the folder hierarchy is recreated) — with `summary.csv` and `index.html`. /
  フォルダの一括処理 (ローカルのフォルダ、SharePoint のフォルダ)
- Open / drop a .docx. If Word removes tracked-change dates on save, the app offers to fix the setting (keeping a backup) —
  same as `--preserve-history`. / ファイルを開く・ドロップ。日時が削除される設定なら修正を案内 (バックアップを残す)
- Chart and Flow tabs; change the idle-gap threshold, period, bulk-insertion size, and time step. /
  チャートとフローのタブ。区間のしきい値・期間・一括挿入の文字数・時間の刻みを変更できる
- Highlights tab: the classifiers' highlights with reasons and the inserted text; click to show where they are in
  the figures. / ハイライトの一覧 (理由・本文)。クリックで図の該当箇所を示す
- Settings: language, rules file (`--rules`), extensions ([EXTENSIONS.md](EXTENSIONS.md)). / 設定: 表示言語・ルールファイル・拡張機能
- Save the figure as SVG or PNG. / 図を SVG / PNG で保存

## Development / 開発

```bash
npm install            # in the repository root (library dependencies, esbuild)
cd desktop
npm install            # Electron
npm start              # build and launch
npm run typecheck
npm run smoke          # launch hidden, open fixtures/flow-demo.docx, save screenshots of each tab to smoke-out/
```

If Electron's binary wasn't downloaded (install scripts disabled), run `node node_modules/electron/install.js`.

### Structure / 構成

- `src/main.ts` — main process: file I/O, dialogs, settings (`userData/settings.json`), rewriting the docx settings,
  blocking network access and navigation.
- `src/preload.ts` — exposes `window.app` (`src/shared.ts`) to the sandboxed renderer.
- `src/renderer/` — UI. Analysis and drawing use the library core (`../src/core.ts`, no file system).

Packaging, signing, notarization, releases and auto-update: see [DISTRIBUTION.md](DISTRIBUTION.md).
