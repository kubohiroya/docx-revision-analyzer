# docx-revision-analyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[English](./README.md) | 日本語

変更履歴 (Track Changes / 校閲) を有効にして編集された Word ファイル (`.docx`) から、

1. **`docx-revision-chart`** — 編集の時系列を可視化する SVG チャートを生成
2. **`docx-ai-suspicion-score`** — 「外部で作文した文章を貼り付けたのでは」と疑われる
   不自然な文字数増加を検出し、0〜100の「AI不正利用疑いスコア」を算出

する2つの CLI ツールです。npm パッケージとして構成されており、`docx-revision-chart`
は [Bun](https://bun.sh) で依存パッケージ込みの単一実行ファイル (シングルバイナリ) に
することもできます (Node.js 不要で配布可能)。ライセンスは MIT です。

> **重要な前提**: どちらのツールも、Word の「変更履歴の記録」(校閲タブ → 変更履歴の記録) を
> 有効にした状態で作成・編集された `.docx` ファイルが必要です。変更履歴の記録がオフの状態で
> 編集されたファイルには挿入(`w:ins`)/削除(`w:del`)の情報が残らないため、解析できません
> (エラーにはなりませんが、イベント数0として扱われます)。

---

## セットアップ

```bash
git clone https://github.com/kubohiroya/docx-revision-analyzer.git
cd docx-revision-analyzer
npm install
npm run build
```

Node.js 18 以降を想定しています (開発・動作確認は Node.js 22 で実施)。
ビルド後は `dist/cli/chart.js` / `dist/cli/score.js` を `node` で直接実行できます
(下記の使用例を参照)。`npm link` するとコマンド名 `docx-revision-chart` /
`docx-ai-suspicion-score` としてどこからでも実行できるようになります。

---

## Bunでシングルバイナリ化する (`docx-revision-chart`)

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
(`--all` を追加すると同様に一括クロスビルドできます)。

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
  機能そのものです (下記のCLI仕様を参照)。
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

## 1. `docx-revision-chart` — 編集履歴のSVGチャート化

```bash
node dist/cli/chart.js <input.docx> [input2.docx ...] [options]
```

複数の入力ファイルをまとめて指定できます (各ファイルは独立に処理され、
1件が失敗しても残りの処理は続行されます)。主にWindowsで、エクスプローラーから
複数ファイルを実行ファイルにドロップした際に、それらが1回の起動に対する
複数の引数として渡されるケースに対応するためのものです。`-o/--output` は
入力ファイルが1つの場合のみ使用できます。

| オプション | 説明 | 既定値 |
|---|---|---|
| `-o, --output <file.svg>` | 出力SVGのパス (単一ファイル指定時のみ使用可) | `<入力ファイル名>.svg` |
| `-b, --bucket <spec>` | 時間バケットの粒度: `auto`\|`second`\|`minute`\|`hour`\|`day`、または秒数 | `auto` |
| `-p, --gap-threshold <hours>` | 「連続的に更新が行われた期間」と「更新が無かった期間」を区別する閾値(時間)。指定すると期間ごとに分割した詳細なグラフを水平に並べて表示する (下記参照) | 指定なし (全体を1枚のグラフにする) |
| `-w, --width <px>` | 画像の幅 | `1100` (`-p` 指定時は内容に応じて自動計算) |
| `-H, --height <px>` | 画像の高さ | `550` |
| `-t, --title <text>` | チャートタイトル | `編集履歴: <ファイル名>` |
| `--drop` | デスクトップからのドラッグ&ドロップ起動モード。`-o` 未指定時、各ファイルの出力先を通常の `<ファイル名>.svg` ではなく `<入力と同じディレクトリ>/<そのファイル名>-<そのファイルの最終更新日時>.svg` にする。上記のmacOS用Finderドロップレットや、Windowsエクスプローラーからの直接ドロップ(またはターミナルを介さない他の起動方法)向け。Windows上ではさらにネイティブなメッセージボックスで結果を表示する | off |

### 出力される図の見方

- **横軸**: 時間 (バケットごとに集計)
- **縦軸・上向きの棒 (緑)**: そのバケット内で追加された文字数
- **縦軸・下向きの棒 (赤)**: そのバケット内で削除された文字数 (絶対値で表示)
- **折れ線 (青、右軸)**: バケット終了時点での文書の総文字数 (追跡開始前の
  文字数 + それまでの純増減の累計)

```bash
node dist/cli/chart.js fixtures/natural-writing.docx -o out.svg
```

同梱の `fixtures/*.svg` (および見た目確認用の `fixtures/*.png`) がサンプル出力です。

### `-p` オプション: 更新期間ごとの詳細表示

`-p <時間>` を指定すると、直前の変更履歴イベントからの経過時間がこの閾値を
超えるたびに「無編集期間」とみなして区切り、**連続的に更新が行われた期間 (セッション)
ごとに独立したグラフを作成して水平に並べます**。各セッションは独立に時間バケットへ
集計されるため、たとえば数日にわたる編集履歴でも、1本の全体グラフでは埋もれてしまう
短時間の編集の起伏を、セッションごとに詳細に見ることができます。

セッションとセッションの間には、ラベル文字 (`"12 h"` のように、およそ4文字分) が
収まる程度の狭い間隔を破線区切りで挿入し、その間の無編集期間の長さを表示します。
追加/削除文字数の縦軸、総文字数の縦軸 (右軸) はすべてのセッションで共通のスケールを
使うため、セッション間で活動量を比較できます。総文字数の折れ線は、無編集期間をまたいで
実際には経過していない時間を線でつながないよう、セッションごとに独立させています。

```bash
node dist/cli/chart.js fixtures/multi-session.docx -p 12 -o out.svg
```

同梱の `fixtures/multi-session.docx` (3日間、間に29時間・20.5時間の空白を挟んだ想定)
に対する `-p 12` の出力例が `fixtures/multi-session-p12.svg` / `.png` です。
`-p` を指定しない場合の全体表示 (`fixtures/multi-session-single.svg` / `.png`) と
見比べると、セッションに分割することで編集の起伏がどれだけ詳細に見えるようになるかが
わかります。

---

## 2. `docx-ai-suspicion-score` — AI不正利用疑いスコア

```bash
node dist/cli/score.js <input.docx> [options]
```

| オプション | 説明 | 既定値 |
|---|---|---|
| `-o, --output <file.json>` | 結果の出力先 (省略時は標準出力) | - |
| `--min-chars <n>` | この文字数未満の挿入は誤検知防止のため無視 | `20` |
| `--burst-low <n>` | バーストスコアが0になる文字数境界 | `20` |
| `--burst-high <n>` | バーストスコアが100になる文字数境界 | `300` |
| `--rate-low <cps>` | 速度スコアが0になる挿入速度 (文字/秒) | `8` |
| `--rate-high <cps>` | 速度スコアが100になる挿入速度 (文字/秒) | `40` |
| `--max-weight <0-1>` | 全体スコアにおける「最も疑わしい1件」の重み (残りは文字数加重平均) | `0.6` |
| `--pretty` | JSONを整形して出力 | off |

```bash
node dist/cli/score.js fixtures/suspicious-paste.docx --pretty
```

### スコアの算出方法 (アルゴリズムの説明)

Word は挿入された文字列を、挿入した日時 (秒単位) と一緒に `w:ins` 要素として
記録します。**WordのGUIで自分の手でタイプ・校閲した場合、1回の挿入イベントで
一度に増える文字数はせいぜい数十文字程度に収まる**のに対し、**外部アプリ
(ChatGPT等) で作文した完成文をコピー&ペーストした場合、1回の挿入イベントで
数百文字が一瞬 (1〜数秒) で増加**します。この違いを検出しています。

各挿入イベントについて、次の2つのスコア (0〜100) を計算し、大きい方を
そのイベントの疑いスコアとします。

- **バーストスコア**: そのイベント単体の文字数の大きさ
  (`--burst-low` 文字以下なら0、`--burst-high` 文字以上なら100、線形補間)
- **速度スコア**: 直前の挿入イベントからの経過時間に対する挿入速度 (文字/秒)
  (`--rate-low` cps以下なら0、`--rate-high` cps以上なら100、線形補間)

ただし `--min-chars` 未満の小さな挿入は常にスコア0とし、誤検知を防ぎます。

文書全体のスコアは、

```
score = max_weight × (最も疑わしいイベントのスコア)
      + (1 − max_weight) × (文字数で重み付けした全イベントの平均スコア)
```

の加重平均です (既定では `max_weight = 0.6`)。最大値だけで決めると偶然の
誤検知に引っ張られやすく、平均だけで決めると大きな1回の貼り付けが他の
多数の正常な編集で薄まってしまうため、両者をブレンドしています。

`riskLevel` は参考の目安です: `low` (0〜19) / `medium` (20〜44) /
`high` (45〜74) / `very_high` (75〜100)。

### ⚠️ 限界・注意点 (必ずお読みください)

このスコアは **統計的なヒューリスティックであり、不正の「証拠」にはなりません**。
以下のような正当なケースでもスコアが高くなり得ます (誤検知の可能性):

- 非常にタイピングが速い人、音声入力を使っている人
- 別のメモアプリやエディタで下書きしてから Word に貼り付けた人
  (本人の文章でも、貼り付けた時点では「一瞬での大量挿入」に見える)
- 表や参考文献リストなど、機械的にまとまった量を貼り付けた場合

逆に、以下のようなケースは検出をすり抜けます (見逃しの可能性):

- AIが生成した文章を、時間をかけて少しずつ手で書き写した場合
- 貼り付け後に文章全体を大きく手直しして、挿入イベントを細かく分割した場合

そのため、**このスコアは「疑いの度合いの参考情報」として使い、最終的な判断は
必ず人間 (教員等) が文脈と合わせて行ってください**。学生に不利益な判定を
自動的に下す用途には使わないことを強く推奨します。

その他の技術的な制約:

- `w:date` のタイムスタンプは秒単位精度のため、1秒未満の間隔は区別できません。
- `w:moveFrom` / `w:moveTo` (文書内でのドラッグ移動等) は本バージョンでは
  特別扱いしておらず、Wordのバージョン・設定によっては通常の挿入と同様に
  カウントされる場合があります。
- 既定では `word/document.xml` (本文) のみを解析します。ヘッダー/フッター/
  コメント/脚注中の変更履歴は対象外です。

---

## テスト用フィクスチャ

実際の変更履歴付きWordファイルが手元になくても動作確認できるよう、
`scripts/makeFixtures.ts` で人工的な `.docx` を生成できます。

```bash
npm run fixtures
# 内部で: ts-node scripts/makeFixtures.ts (ビルド不要でその場でTSを実行)
```

- `fixtures/natural-writing.docx`: 約37分かけて少しずつタイプしたことを想定
  (スコア: 0 / low)
- `fixtures/suspicious-paste.docx`: 最初に少しタイプした直後、295文字を1秒で
  一括挿入 (外部での作文の貼り付けを模擬) したことを想定 (スコア: 93 / very_high)
- `fixtures/multi-session.docx`: 3日間に分けて執筆し、間に29時間・20.5時間の
  無編集期間があったことを想定 (`-p` オプションの動作確認用)

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
│   │   └── score.ts       docx-ai-suspicion-score CLI本体
│   └── lib/
│       ├── docxRevisions.ts  docxから変更履歴イベントを抽出する共通ライブラリ
│       ├── timeBuckets.ts    時間バケットへの集計 (チャート用)
│       ├── sessions.ts       -p オプション用: 無編集期間で区切ったセッション分割
│       ├── svgChart.ts       SVGチャートのレンダリング (全体版・セッション分割版)
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

`src/index.ts` から主要な関数・型 (`extractRevisionsFromFile`, `computeSuspicionScore`,
`renderRevisionChart` など) を再エクスポートしているため、CLIとしてだけでなく
他の Node.js / Bun プロジェクトから `import { ... } from "docx-revision-analyzer"`
としてライブラリ的に使うこともできます。

## コントリビュート

Issue / Pull Request 歓迎です。バグ報告や改善提案は
[GitHub Issues](https://github.com/kubohiroya/docx-revision-analyzer/issues) へ。

## ライセンス

[MIT](./LICENSE) © 2026 Hiroya Kubo
