# インストール

[English](./installation.md) | 日本語 · [README に戻る](../README.ja.md)

npm パッケージ・単体実行ファイル・ドラッグ&ドロップ版の入手と準備の方法です。ソースからのビルドは[開発者向けガイド](development.ja.md)を参照してください。

このページは次の CLI リリースに向けた統合版の説明です。公開済み **v1.2.0** の npm パッケージと配布物は個別コマンド方式です。統合版を今試す場合は、[開発ガイド](development.ja.md)の手順でソースからビルドしてください。

**目次**

- [npm から (Node.js 18 以降)](#npm-から-nodejs-18-以降)
- [単体バイナリ (Node.js 不要)](#単体バイナリ-nodejs-不要)
- [Finderドラッグ&ドロップ対応アプリ (macOS専用)](#finderドラッグドロップ対応アプリ-macos専用)
- [Windows での実行](#windows-での実行)
- [旧コマンドからの移行](#旧コマンドからの移行)
- [デスクトップアプリ](#デスクトップアプリ)

## npm から (Node.js 18 以降)

```bash
npm install -g docx-revision-analyzer
```

`docx-revision-analyzer` という1つのコマンドに、`chart` / `flow` / `score` / `snapshot` / `versions` の5サブコマンドが入ります。別名コマンドや別名リンクは作成しません。

```bash
docx-revision-analyzer --help
docx-revision-analyzer flow 報告書.docx
docx-revision-analyzer help versions
```

## 単体バイナリ (Node.js 不要)

次のリリースからは、[GitHub Releases](https://github.com/kubohiroya/docx-revision-analyzer/releases) で OS / CPU に合う `docx-revision-analyzer-<os>-<cpu>.zip` を **1つ**選び、展開します。

| 環境 | 選ぶファイル | 中の実行ファイル |
|---|---|---|
| Linux x64 / ARM64 | `docx-revision-analyzer-linux-x64.zip` / `-linux-arm64.zip` | `docx-revision-analyzer` |
| Mac Intel / Apple Silicon | `docx-revision-analyzer-macos-x64.zip` / `-macos-arm64.zip` | `docx-revision-analyzer` |
| Windows x64 | `docx-revision-analyzer-windows-x64.zip` | `docx-revision-analyzer.exe` |

どの実行ファイルにも5つのサブコマンドが含まれます。macOS / Linux で必要な場合は `chmod +x docx-revision-analyzer` を実行し、`./docx-revision-analyzer flow 報告書.docx` のように使います。Windows の例は後述します。

macOS のドロップレットは `docx-revision-chart.app` / `docx-revision-flow.app` の名前を維持し、それぞれ共通 CLI に `chart` / `flow` を渡します。各 `.app` だけで動作します。ad-hoc 署名のみのため、初回は Finder で右クリック →「開く」を選んでください。

ソースからのビルドと出力先は[開発ガイド](development.ja.md)を参照してください。

## Finderドラッグ&ドロップ対応アプリ (macOS専用)

コマンドラインだけでなく、`docx-revision-analyzer chart` をmacOSの小さなアプリ
(「ドロップレット」) としてパッケージすることもできます。Finder上で1つ、
または複数の `.docx` ファイルをこのアプリのアイコンに直接ドラッグ&ドロップすると、
ターミナルを開かずにグラフが生成されます。ドロップした各ファイルについて、
その隣に `<ファイル名>-<そのファイルの最終更新日時>.svg`
(例: `report-20260919-143000.svg`) という名前でグラフを出力します。
成功/失敗はコンソール出力ではなく、macOSのネイティブな通知/ダイアログで
知らされます。

```bash
npm run build:droplet
```

このコマンドは 共通のBunバイナリを1つビルドしたうえで、`osacompile`
(AppleScriptを.appにコンパイルする、Macに標準搭載されているツール。Bun以外の
追加インストール不要) を使って `release/droplet/mac-<arch>/docx-revision-chart.app`
(`<arch>` は `arm64` または `x64`) としてパッケージします。生成された `.app` は好きな場所 (Applicationsフォルダや
Dockなど) に置いて、`.docx` ファイルをそのアイコンにドロップしてください。

同じ仕組みで `docx-revision-analyzer flow` の
ドロップレット `release/droplet/mac-<arch>/docx-revision-flow.app` も作られます。こちらは全期間・既定の設定で
`<ファイル名>-flow-<そのファイルの最終更新日時>.svg` を出力します。

なぜラッパーが必要か: macOSのFinderは、ドロップされたファイルを
(Apple Eventという仕組み経由で) 正式な「アプリケーションバンドル」にしか
配送せず、バンドル化されていない素の実行ファイルには直接コマンドライン引数
として渡すことができません。そのため、Bunでビルドしたバイナリ単体では
ドロップ対象になれません。`scripts/make-mac-droplet.sh` は、`on open`
ハンドラでバンドルされたバイナリを `chart --drop` / `flow --drop` 付きで呼び出す、極小の
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

## Windows での実行

Windows 上では `npm run build:cli`、クロスビルドでは `npm run build:cli:all-platforms` を使います。Windows x64 向けの出力は `release/cli/win-x64/docx-revision-analyzer.exe` です。

```powershell
.\docx-revision-analyzer.exe --help
.\docx-revision-analyzer.exe chart .\report.docx
.\docx-revision-analyzer.exe flow .\report.docx .\report2.docx
```

CLI はサブコマンドが必要なため、`.exe` に文書を直接ドロップしても解析方法を選べません。画面でドラッグ＆ドロップして使う場合は[デスクトップアプリ](#デスクトップアプリ)を使ってください。ランチャーなどから `chart` / `flow` に `--drop` を付けて実行した場合の、日時付き出力名と Windows の結果ダイアログは維持しています。

## 旧コマンドからの移行

npm と単体バイナリの公開コマンドを1つに統合しました。呼び出しを次のように置き換えてください。各サブコマンドのオプション・出力形式は維持しています。旧名の実行ファイルや別名リンクは生成しません。

| 旧コマンド | 新しい呼び出し |
|---|---|
| `docx-revision-chart` | `docx-revision-analyzer chart` |
| `docx-revision-flow` | `docx-revision-analyzer flow` |
| `docx-ai-suspicion-score` | `docx-revision-analyzer score` |
| `docx-revision-snapshot` | `docx-revision-analyzer snapshot` |
| `docx-revision-versions` | `docx-revision-analyzer versions` |

既存の設定ファイル名 (`docx-revision-chart.yml` など) と解析 JSON の `tool` の識別子は変わりません。
ビルド時の `--cli=...` は廃止し、常に全5サブコマンドを含めます。

## デスクトップアプリ

デスクトップアプリ (Electron、プレビュー) のインストーラは、[desktop-v0.2.0](https://github.com/kubohiroya/docx-revision-analyzer/releases/tag/desktop-v0.2.0) で配布しています。Mac は `.dmg`、Windows は `Setup-0.2.0.exe` を選んでください。CLI とは別の配布物です。
ソースからの起動は [desktop/README.md](../desktop/README.md)、パッケージの作成は [desktop/DISTRIBUTION.md](../desktop/DISTRIBUTION.md)、機能は[使い方](usage.ja.md#デスクトップアプリ-プレビュー)を参照してください。
