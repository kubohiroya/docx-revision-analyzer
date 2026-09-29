# docx-revision-analyzer のインストールとビルド

[English](./INSTALL.md) | 日本語

使い方は [README.ja.md](./README.ja.md) を参照してください。配布用のファイル (npm パッケージ・単体バイナリ) を
使うだけなら、README の「インストール」の手順で足ります。ここでは、ソースからのビルドやパッケージの作り方を説明します。

---

## セットアップ

```bash
git clone https://github.com/kubohiroya/docx-revision-analyzer.git
cd docx-revision-analyzer
npm install
npm run build
```

Node.js 18 以降を想定しています (開発・動作確認は Node.js 22 で実施)。
ビルド後は `dist/cli/chart.js` / `dist/cli/flow.js` / `dist/cli/score.js` を `node` で直接実行できます
(使い方は [README.ja.md](./README.ja.md) を参照)。`npm link` するとコマンド名 `docx-revision-chart` /
`docx-revision-flow` / `docx-ai-suspicion-score` としてどこからでも実行できるようになります。

---

## Bunでシングルバイナリ化する

[Bun](https://bun.sh) がインストールされていれば、`docx-revision-chart` を
Node.js 本体すら不要な単一の実行ファイルにまとめられます (JSZip / fast-xml-parser /
commander などの依存パッケージもすべてバイナリに埋め込まれます)。TypeScript の
ソース (`src/cli/chart.ts`) を直接コンパイルするため、事前に `npm run build` する
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

---

## Finderドラッグ&ドロップ対応アプリ (macOS専用)

コマンドラインだけでなく、`docx-revision-chart` をmacOSの小さなアプリ
(「ドロップレット」) としてパッケージすることもできます。Finder上で1つ、
または複数の `.docx` ファイルをこのアプリのアイコンに直接ドラッグ&ドロップすると、
ターミナルを開かずにグラフが生成されます。ドロップした各ファイルについて、
その隣に `<ファイル名>-<そのファイルの最終更新日時>.svg`
(例: `report-20260919-143000.svg`) という名前でグラフを出力します。
成功/失敗はコンソール出力ではなく、macOSのネイティブな通知/ダイアログで
知らされます。

```bash
npm run build:mac-app
```

このコマンドはBunバイナリをビルドしたうえで、`osacompile`
(AppleScriptを.appにコンパイルする、Macに標準搭載されているツール。Bun以外の
追加インストール不要) を使って `dist-bin/docx-revision-chart.app` として
パッケージします。生成された `.app` は好きな場所 (Applicationsフォルダや
Dockなど) に置いて、`.docx` ファイルをそのアイコンにドロップしてください。

`npm run build:mac-app:flow` を実行すると、同じ仕組みで `docx-revision-flow` の
ドロップレット `dist-bin/docx-revision-flow.app` を作れます。こちらは全期間・既定の設定で
`<ファイル名>-flow-<そのファイルの最終更新日時>.svg` を出力します。

なぜラッパーが必要か: macOSのFinderは、ドロップされたファイルを
(Apple Eventという仕組み経由で) 正式な「アプリケーションバンドル」にしか
配送せず、バンドル化されていない素の実行ファイルには直接コマンドライン引数
として渡すことができません。そのため、Bunでビルドしたバイナリ単体では
ドロップ対象になれません。`scripts/make-mac-droplet.sh` は、`on open`
ハンドラでバンドルされたバイナリを `--drop` 付きで呼び出す、極小の
AppleScriptアプリを生成することでこれを解決しています。

知っておくべき点:

- **macOS専用**であり、ビルドスクリプトはお使いのMacの実際のターミナルで
  実行する必要があります (サンドボックス化されたLinux環境からは実行できません)。
- 生成されるアプリはビルドスクリプトによって **ad-hoc署名** (`codesign -s -`)
  されます。Apple Silicon上でローカルに実行するにはこれで十分ですが、
  ノータライズ (notarization) はしていないため、初回起動時にGatekeeperが
  「開発元を確認できません」と警告することがあります。その場合はアプリを
  右クリックして「開く」を選ぶと、確認ダイアログ経由で起動できます。
- 一度ビルドすれば、その `.app` はBun・Node.js・このリポジトリのいずれにも
  依存せず単体で動作し続けます。
- この機能は現時点の公開情報を調査したうえで実装していますが、実際のMac上で
  作者自身が動作確認を完了できていません。ドラッグ&ドロップの挙動が
  説明と異なる場合は、Issueで報告してください。

---

## Windowsでのドラッグ&ドロップ (ラッパーアプリ不要)

Windowsでは、macOSのような専用ラッパーアプリは不要です。エクスプローラーは
`.exe` にドロップされたファイルをそのままコマンドライン引数として渡して
直接起動するため、`npm run build:binary:all` で生成されるWindows向けバイナリ
(`bun-windows-x64` ターゲット) には `osacompile` のようなパッケージング工程
なしで直接ファイルをドロップできます。

- `docx-revision-chart.exe` に **1つまたは複数** の `.docx` ファイルを
  まとめてドロップできます。エクスプローラーはドロップされた全ファイルを
  1回のプロセス起動に対する複数の引数として渡すため、各ファイルは独立に
  処理されます (1件が失敗しても残りの処理は続行されます)。これは今回、
  この用途のために `docx-revision-chart` に追加した複数ファイル対応の
  機能そのものです ([README.ja.md](./README.ja.md) のCLI仕様を参照)。
- `--drop` はmacOSと同様に使えます。指定すると、各出力ファイル名が通常の
  `<ファイル名>.svg` ではなく `<ファイル名>-<そのファイルの最終更新日時>.svg`
  になります。
- Windowsでは、ダブルクリックやドラッグ&ドロップで起動したコンソール
  サブシステムの実行ファイルは、プロセス終了と同時にコンソールウィンドウが
  開いてすぐ閉じてしまうため、そこに表示された内容を読むことができません。
  ターミナルなしでも結果を確認できるよう、Windows上で `--drop` を使うと、
  同梱のPowerShell呼び出し経由でネイティブなメッセージボックスが表示され、
  各ファイルの成功/失敗が要約表示されます。
- macOS用ドロップレットと同様、**この機能も実装・検討は済んでいますが、
  実際のWindows環境での動作確認は完了していません**。挙動が説明と異なる
  場合は、Issueで報告してください。

```powershell
# macOS/Linux上でクロスコンパイル (Windows上でBunをインストールして直接ビルドも可)
npm run build:binary:all
# -> dist-bin/docx-revision-chart-windows-x64.exe (正確なファイル名は build-binary.sh に依存)
```

---

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

---

## テスト用フィクスチャ

実際の変更履歴付きWordファイルが手元になくても動作確認できるよう、
`scripts/makeFixtures.ts` で人工的な `.docx` を生成できます。

```bash
npm run fixtures
# 内部で: ts-node scripts/makeFixtures.ts (ビルド不要でその場でTSを実行)
```

- `fixtures/chart-demo.docx`: 約3時間の執筆で、細かい入力の途中に一度まとめて貼り付け、下書きを散発的に
  削除したことを想定 (README の `docx-revision-chart` の図の例)
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

これらに対する `docx-revision-chart` の出力例が `fixtures/*.svg` (`*.png` は
確認用にラスタライズしたもの) として同梱されています。

---

## ディレクトリ構成

標準的な npm パッケージのレイアウトに沿っています。`src/` が TypeScript ソース、
`dist/` がビルド成果物 (JS + 型定義、`.gitignore` 済・npm publish 時のみ `files` で
同梱)、`dist-bin/` が Bun によるシングルバイナリの出力先です。

```
docx-revision-analyzer/
├── package.json          パッケージ定義 (bin, files, license: MIT 等)
├── tsconfig.json
├── LICENSE                MIT ライセンス全文
├── README.md              英語版マニュアル
├── README.ja.md           日本語版マニュアル (このファイル)
├── .gitignore
├── src/
│   ├── index.ts           ライブラリとしてimportする場合のエントリポイント
│   ├── cli/
│   │   ├── chart.ts       docx-revision-chart CLI本体
│   │   ├── flow.ts        docx-revision-flow CLI本体
│   │   ├── common.ts      CLI共通処理 (--preserve-history の確認、複数ファイルの処理と結果表示)
│   │   └── score.ts       docx-ai-suspicion-score CLI本体
│   └── lib/
│       ├── docxRevisions.ts  docxから変更履歴イベントを抽出する共通ライブラリ
│       ├── insertionKinds.ts 挿入の分類 (一括挿入・細かい編集・移動/並べ替え)。chart と flow で共通
│       ├── historySettings.ts 変更履歴の記録・個人情報削除の設定の確認と書き換え (--preserve-history)
│       ├── timeBuckets.ts    時間バケットへの集計 (チャート用)
│       ├── sessions.ts       -p オプション用: 無編集期間で区切ったセッション分割
│       ├── svgChart.ts       SVGチャートのレンダリング (全体版・セッション分割版)
│       ├── docxLayout.ts     flow 用: 本文を段落と変更履歴付きの断片に分解し、用紙サイズ等を読み取る
│       ├── flow.ts           flow 用: 区間の開始・終了時点の文書の復元・模式的なページ割り・段落ごとの度合いと変化の分類
│       ├── flowSvg.ts        flow 用: SVG のレンダリング
│       ├── suspicionScore.ts AI不正利用疑いスコアの算出ロジック
│       └── filenames.ts      --drop 用: 最終更新日時を使った出力ファイル名の生成
├── scripts/
│   ├── build-binary.sh       Bunでシングルバイナリ化するビルドスクリプト
│   ├── make-mac-droplet.sh   macOS用Finderドロップレット(.app)を作るビルドスクリプト
│   └── makeFixtures.ts       テスト用docx生成スクリプト
├── fixtures/              生成済みのテスト用docx・出力例
├── dist/                  `npm run build` の出力 (gitignore対象)
└── dist-bin/              `npm run build:binary` の出力 (gitignore対象)
```
