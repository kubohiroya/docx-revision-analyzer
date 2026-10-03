# インストール

[English](./installation.md) | 日本語 · [README に戻る](../README.ja.md)

npm パッケージ・単体実行ファイル・ドラッグ&ドロップ版の入手と準備の方法です。ソースからのビルドは[開発者向けガイド](development.ja.md)を参照してください。

**目次**

- [npm から (Node.js 18 以降)](#npm-から-nodejs-18-以降)
- [単体バイナリ (Node.js 不要)](#単体バイナリ-nodejs-不要)
- [Finderドラッグ&ドロップ対応アプリ (macOS専用)](#finderドラッグドロップ対応アプリ-macos専用)
- [Windowsでのドラッグ&ドロップ (ラッパーアプリ不要)](#windowsでのドラッグドロップ-ラッパーアプリ不要)
- [デスクトップアプリ](#デスクトップアプリ)

## npm から (Node.js 18 以降)

```bash
npm install -g docx-revision-analyzer
```

`docx-revision-chart` / `docx-revision-flow` / `docx-ai-suspicion-score` / `docx-revision-snapshot` / `docx-revision-versions` の5つのコマンドが使えるようになります。

## 単体バイナリ (Node.js 不要)

[GitHub Releases](https://github.com/kubohiroya/docx-revision-analyzer/releases) から、お使いの OS / CPU に合ったファイルをダウンロードして展開してください。

| ファイル | 内容 |
|---|---|
| `docx-revision-<chart\|flow>-<OS>-<CPU>.zip`、`docx-ai-suspicion-score-<OS>-<CPU>.zip` | 各ツールの単体実行ファイル (`linux-x64` / `linux-arm64` / `macos-x64` / `macos-arm64` / `windows-x64`) |
| `docx-revision-<chart\|flow>-macos-arm64.app.zip` | macOS (Apple Silicon) 用のドラッグ&ドロップ版アプリ。Finder で `.docx` をアイコンにドロップすると、同じフォルダに SVG を出力します |

ドラッグ&ドロップ版は署名が ad-hoc のため、初回は Finder でアプリを右クリックして「開く」を選んでください。
Windows では `.exe` に `.docx` を直接ドロップできます。

ソースからのビルド、単体バイナリやドラッグ&ドロップ版の作り方、テスト用フィクスチャ、ディレクトリ構成は
[開発者向けガイド](development.ja.md)を参照してください。

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
  機能そのものです ([使い方](usage.ja.md)を参照)。
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

## デスクトップアプリ

コマンドラインに慣れていない人向けのデスクトップアプリ (Electron) は、インストーラをまだ配布していません。
ソースから起動する方法は [desktop/README.md](../desktop/README.md)、配布用のパッケージの作り方は
[desktop/DISTRIBUTION.md](../desktop/DISTRIBUTION.md) を参照してください。機能は[使い方](usage.ja.md#デスクトップアプリ-プレビュー)にあります。
