# 開発者向けガイド

[English](./development.md) | 日本語 · [README に戻る](../README.ja.md)

ソースからのビルド、単体実行ファイルとドラッグ&ドロップ版の作り方、公開の手順、テスト用フィクスチャ、ディレクトリ構成です。

**目次**

- [セットアップ](#セットアップ)
- [Bunでシングルバイナリ化する](#bunでシングルバイナリ化する)
- [npm への公開 (メンテナー向け)](#npm-への公開-メンテナー向け)
- [GitHub Releases への単体バイナリの添付 (メンテナー向け)](#github-releases-への単体バイナリの添付-メンテナー向け)
- [テスト用フィクスチャ](#テスト用フィクスチャ)
- [ディレクトリ構成](#ディレクトリ構成)
- [コントリビュート](#コントリビュート)

## セットアップ

```bash
git clone https://github.com/kubohiroya/docx-revision-analyzer.git
cd docx-revision-analyzer
npm install
npm run build
```

パッケージ管理は **npm** に統一しています (`package-lock.json`、`"packageManager": "npm@…"`。CI は `npm ci`)。手元のビルドには
pnpm も使えます (`pnpm install`、`pnpm run build`) が、`pnpm-lock.yaml` はコミットしないでください (gitignore 済み)。
`pnpm-workspace.yaml` は、pnpm で esbuild のインストールスクリプトを許可するためだけのものです。

`npm run build` (`pnpm run build`) は、**この環境で作れるものを一括で**作り、最後に一覧を表示します。

| 工程 | 作るもの | 条件 |
|---|---|---|
| `lib` | ライブラリと CLI → `dist/` | 常に |
| `binaries` | 単体実行ファイル `docx-revision-chart` / `docx-revision-flow` / `docx-ai-suspicion-score` → `dist-bin/` | [Bun](https://bun.sh) があれば |
| `droplets` | macOS のドロップレット `docx-revision-chart.app` / `docx-revision-flow.app` → `dist-bin/` | macOS で Bun があれば |
| `desktop` | この OS / CPU 向けのデスクトップアプリ (署名なし) → `desktop/release/` (初回は `desktop/` の依存も入れる) | 常に |

必要なツールが無い工程は、理由を表示して飛ばします。オプション (`--` の後に指定):

```bash
npm run build -- --skip=desktop          # デスクトップアプリ以外
npm run build -- --only=lib,binaries     # 指定した工程だけ
npm run build -- --all-platforms         # 単体実行ファイルを Linux/macOS/Windows × x64/arm64 向けにクロスビルド
npm run build -- --installers            # デスクトップアプリのインストーラ (dmg/zip/NSIS) も作る (時間がかかる)
```

ライブラリのコンパイルだけなら `npm run build:lib` です (`prepublishOnly` と CI はこちらを使います)。以下の個別のコマンド
(`build:binary:flow`、`build:mac-app:flow` など) も引き続き使えます。

Node.js 18 以降を想定しています (開発・動作確認は Node.js 22 で実施)。
ビルド後は `dist/cli/chart.js` / `dist/cli/flow.js` / `dist/cli/score.js` を `node` で直接実行できます
(使い方は[使い方](usage.ja.md)を参照)。`npm link` するとコマンド名 `docx-revision-chart` /
`docx-revision-flow` / `docx-ai-suspicion-score` としてどこからでも実行できるようになります。

## Bunでシングルバイナリ化する

[Bun](https://bun.sh) がインストールされていれば、`docx-revision-chart` を
Node.js 本体すら不要な単一の実行ファイルにまとめられます (JSZip / fast-xml-parser /
commander などの依存パッケージもすべてバイナリに埋め込まれます)。TypeScript の
ソース (`src/cli/chart.ts`) を直接コンパイルするため、事前に `npm run build:lib` する
必要はありません。

```bash
# Bun のインストール (未導入の場合)
curl -fsSL https://bun.sh/install | bash

# 実行中のOS/CPU向けに1つビルド → dist-bin/docx-revision-chart
npm run build:binary

# 主要なOS/CPU向けに一括クロスビルド (Linux/macOS/Windows × x64/arm64)
npm run build:binary:all
```

生成された実行ファイルはそのまま配布・実行できます (Node.jsのインストール不要):

```bash
./dist-bin/docx-revision-chart fixtures/multi-session.docx -p 12 -o out.svg
```

内部的には `scripts/build-binary.sh` を呼んでいます。`docx-ai-suspicion-score` の
方をバイナリ化したい場合は `npm run build:binary:score`、または
`bash scripts/build-binary.sh score` を実行してください
(`--all` を追加すると同様に一括クロスビルドできます)。`docx-revision-flow` は
`npm run build:binary:flow` (`bash scripts/build-binary.sh flow`) でバイナリ化できます。

> クロスコンパイル (`--all`) で生成した他OS向けの実行ファイルは、ビルドを実行した
> マシン上では動作確認できません。配布前に対象OSでの動作確認を推奨します。

## npm への公開 (メンテナー向け)

```bash
npm version <patch|minor|major>   # package.json のバージョンを上げてタグを作る
npm publish --dry-run              # 公開されるファイルを確認 (prepublishOnly で dist/ を作り直す)
npm publish
```

公開されるのは `dist/`、`README.md` / `README.ja.md`、`LICENSE` だけです (`package.json` の `files`)。
`prepublishOnly` で `dist/` を消してからビルドし直すため、古いビルド結果が混ざることはありません。

## GitHub Releases への単体バイナリの添付 (メンテナー向け)

```bash
npm run build:binary:all            # docx-revision-chart
bash scripts/build-binary.sh flow --all
bash scripts/build-binary.sh score --all
npm run build:mac-app && npm run build:mac-app:flow   # macOS 上でのみ
```

`dist-bin/` にできたファイルを zip にして、リリースに添付します。

## テスト用フィクスチャ

実際の変更履歴付きWordファイルが手元になくても動作確認できるよう、
`scripts/makeFixtures.ts` で人工的な `.docx` を生成できます。

```bash
npm run fixtures
# 内部で: ts-node scripts/makeFixtures.ts (ビルド不要でその場でTSを実行)
```

- `fixtures/chart-demo.docx`: 2日にわたる3回の執筆 (間に4.5時間・17時間の無編集期間) で、1回まとめて
  貼り付け、下書きを散発的に削除したことを想定 ([使い方](usage.ja.md)の `docx-revision-chart` の図の例)
- `fixtures/natural-writing.docx`: 約37分かけて少しずつタイプしたことを想定
  (スコア: 0 / low)
- `fixtures/suspicious-paste.docx`: 最初に少しタイプした直後、295文字を1秒で
  一括挿入 (外部での作文の貼り付けを模擬) したことを想定 (スコア: 93 / very_high)
- `fixtures/multi-session.docx`: 3日間に分けて執筆し、間に29時間・20.5時間の
  無編集期間があったことを想定 (`-p` オプションの動作確認用)
- `fixtures/flow-demo.docx`: 見出し・図を含む複数段落の文書を、3つの時間区間で
  「手で入力」「3段落の一括貼り付けと手直し」「段落の置き換え・削除と細かい修正・図の追加」と編集し、
  4つ目の区間で「カット＋貼り付け (移動として記録)」「コピー＋貼り付け＋削除」による並べ替えを行ったことを想定
  (`docx-revision-flow` の動作確認用。出力例は `fixtures/flow-demo.svg`)
- `fixtures/snapshot-demo/sprint{1,2,3}/thesis.docx`: 卒論を3回の区切りで提出したことを想定。各回の前に教員がすべての変更を
  承諾して返し、第3回には外部で作った文章の貼り付けと、記録をオフにして入力した段落がある (`docx-revision-snapshot` の
  動作確認用)。出力例の `fixtures/snapshot-through.svg` (すべての回を通したチャート) と `fixtures/tamper-warning.svg`
  (第3回のチャート。改ざんの痕跡の警告付き) は、次のように作ります

  ```bash
  for n in 1 2 3; do node dist/cli/snapshot.js /tmp/archive fixtures/snapshot-demo/sprint$n --lang en; done
  # /tmp/archive/thesis/thesis-through-tampered.svg → fixtures/snapshot-through.svg
  # /tmp/archive/thesis/thesis-s03-tampered.svg     → fixtures/tamper-warning.svg  (--lang ja で .ja.svg)
  ```

これらに対する出力例が `fixtures/*.svg` (`*.png` は確認用にラスタライズしたもの) として同梱されています。
既定の名前のものは英語 (`--lang en`)、`*.ja.svg` / `*.ja.png` は日本語 (`--lang ja`) で出力したもので、
英語版と日本語版の README・ドキュメントでそれぞれ使っています。
なお `fixtures/revision-states.svg`・`revision-states-onedrive.svg` (と各 `.ja.svg` / `.png`) は README と運用ガイドの状態遷移図として手で作成したもので、
ツールの出力ではありません。`npm run fixtures` ではなく `npm run diagrams` (`scripts/make-revision-states.mjs`。PNG の生成には Inkscape が必要) で再生成します。

## ディレクトリ構成

標準的な npm パッケージのレイアウトに沿っています。`src/` が TypeScript ソース、
`dist/` がビルド成果物 (JS + 型定義、`.gitignore` 済・npm publish 時のみ `files` で
同梱)、`dist-bin/` が Bun によるシングルバイナリの出力先です。

```
docx-revision-analyzer/
├── package.json          パッケージ定義 (bin, files, license: MIT 等)
├── tsconfig.json
├── LICENSE                MIT ライセンス全文
├── README.md / README.ja.md   概要 (英語 / 日本語)
├── docs/                  ドキュメント (installation, usage, operations-guide, advanced, library, development。各 .ja.md が日本語版)
├── src/
│   ├── index.ts           ライブラリとして import する場合のエントリポイント (Node.js 用)
│   ├── core.ts            ファイルシステムを使わないコア (ブラウザ用。docx-revision-analyzer/core)
│   ├── cli/
│   │   ├── chart.ts       docx-revision-chart CLI本体
│   │   ├── flow.ts        docx-revision-flow CLI本体
│   │   ├── score.ts       docx-ai-suspicion-score CLI本体
│   │   ├── snapshot.ts    docx-revision-snapshot CLI本体 (区切りごとの提出の保管と通し解析)
│   │   ├── versions.ts    docx-revision-versions CLI本体 (過去の時点の文書の書き出し)
│   │   ├── config.ts      設定ファイル (<ツール名>.yml) の読み込み
│   │   └── common.ts      CLI共通処理 (--preserve-history・--lock・--template の処理、複数ファイルの処理と結果表示)
│   ├── lib/
│   │   ├── docxRevisions.ts  docxから変更履歴イベントを抽出する共通ライブラリ
│   │   ├── docxLayout.ts     本文を段落と変更履歴付きの断片に分解し、用紙サイズ等を読み取る
│   │   ├── revisionPositions.ts 各編集の位置 (最終文書と編集時点)
│   │   ├── insertionKinds.ts 挿入の分類 (一括挿入・細かい編集・移動/並べ替え)
│   │   ├── insertionWindows.ts 挿入の窓と特徴量 (--json)
│   │   ├── insertionRules.ts ハイライトの判定ルール (--rules)
│   │   ├── classifiers.ts    分類器のパイプライン
│   │   ├── categories.ts     図の色のカテゴリ
│   │   ├── figureTargets.ts  図の部分のキーと注釈 (--annotations)
│   │   ├── timeBuckets.ts    時間バケットへの集計 (チャート用)
│   │   ├── sessions.ts       無編集期間で区切ったセッション分割
│   │   ├── svgChart.ts       SVGチャートのレンダリング (全体版・セッション分割版)
│   │   ├── flow.ts           flow 用: 区間の開始・終了時点の文書の復元・模式的なページ割り・変化の分類
│   │   ├── flowSvg.ts        flow 用: SVG のレンダリング
│   │   ├── suspicionScore.ts AI不正利用疑いスコアの算出ロジック
│   │   ├── historySettings.ts 変更履歴の記録・個人情報削除の設定の確認と書き換え (--preserve-history)
│   │   ├── integrity.ts      整合性についての情報 (--json の integrity)
│   │   ├── tamperEvidence.ts 改ざんの痕跡の検出と図の警告 (-tampered.svg)
│   │   ├── trackLock.ts      変更履歴のロック (--lock。Word と同じ形式のパスワードのハッシュ値)
│   │   ├── snapshots.ts      スナップショットの通し解析 (重複を除いて各回の新しい変更を取り出す)
│   │   ├── versions.ts       変更履歴から指定時刻の文書 (.docx) を復元する
│   │   ├── filenames.ts      出力ファイル名・見出しの日時
│   │   ├── input.ts          入力 (docx のバイト列) の型
│   │   └── i18n.ts           表示言語の判定とメッセージ (英語・日本語)
│   └── node/              Node.js 専用 (ファイル・暗号・OS の機能を使う部分)
│       ├── files.ts          ファイルからの読み込み、--drop の出力名
│       ├── historyFile.ts    設定の書き換えのファイルへの適用 (バックアップ、開かれていないかの確認)
│       ├── lockFile.ts       テンプレートへのロックのファイルへの適用
│       ├── snapshotArchive.ts スナップショットの通し番号付きの保管
│       └── locale.ts         macOS の言語設定の読み取り
├── scripts/
│   ├── build-all.mjs         作れるものを一括でビルド (npm run build)
│   ├── build-binary.sh       Bunでシングルバイナリ化するビルドスクリプト
│   ├── make-mac-droplet.sh   macOS用Finderドロップレット(.app)を作るビルドスクリプト
│   ├── makeFixtures.ts       テスト用docx生成スクリプト
│   ├── make-revision-states.mjs 状態遷移図の生成
│   └── test-*.mjs            テスト (npm test)
├── desktop/               デスクトップアプリ (Electron)
├── examples/              判定ルールの例
├── fixtures/              生成済みのテスト用docx・出力例・状態遷移図
├── dist/                  `npm run build:lib` の出力 (gitignore対象)
└── dist-bin/              `npm run build:binary` の出力 (gitignore対象)
```

## コントリビュート

Issue / Pull Request 歓迎です。バグ報告や改善提案は
[GitHub Issues](https://github.com/kubohiroya/docx-revision-analyzer/issues) へ。
