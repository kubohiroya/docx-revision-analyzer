# AGENTS.md

このリポジトリで作業する AI エージェント向けの手順です。人向けの詳しい説明は [desktop/DISTRIBUTION.md](desktop/DISTRIBUTION.md)
と [docs/development.md](docs/development.md) にあります。

## デスクトップアプリの自動更新

### 仕組み

- 自動更新は既定では無効で、利用者が設定 (「起動時に GitHub Releases で新しいバージョンを確認する」) で有効にしたときだけ
  動く。メインプロセスの `electron-updater` が確認し、ダウンロードして、アプリの終了時に入れ替える (`desktop/src/main.ts` の
  `checkForUpdates`)。文書や解析結果は送らない。
- 参照先は **`generic` プロバイダ**で、デスクトップ版専用のプレリリース **`desktop-latest`** に置いた更新情報のファイルを読む
  (`desktop/electron-builder.config.cjs` の `publish`。アプリの `Contents/Resources/app-update.yml` に組み込まれる)。
  - `https://github.com/kubohiroya/docx-revision-analyzer/releases/download/desktop-latest/latest-mac.yml` (macOS)
  - `…/desktop-latest/latest.yml` (Windows)
- 更新情報のファイルの配布物の場所 (`url` / `path`) は、各版のリリース `desktop-vX.Y.Z` のダウンロード用の**絶対 URL**
  (`desktop/scripts/update-feed.mjs` が書き換える)。配布物そのものは各版のリリースから落とす。
- **`github` プロバイダを使わない理由**: リポジトリの「Latest」のリリースを読むため、ライブラリ・CLI のリリース `vX.Y.Z`
  (更新情報のファイルが無い) と取り違え、更新が見つからない (electron-updater の `GitHubProvider.getLatestTagName`)。
  プレリリースは「Latest」にならないので、`desktop-latest` ならライブラリのリリースと混ざらない。
- 署名: macOS は署名した配布物でないと自動更新が働かない (Squirrel.Mac が、起動中のアプリと同じ Developer ID の署名を求める)。
  Windows はリポジトリの変数 `WIN_PUBLISHER_NAME` があれば、インストーラの発行者名を検証する。
- 自動更新は新しい版にしか更新しない (`allowDowngrade` は使っていない)。`desktop-latest` を古い版に向け直しても、
  すでに新しい版を入れた利用者は戻らない。

### 関係するファイル

| ファイル | 役割 |
|---|---|
| `desktop/electron-builder.config.cjs` | `publish` (generic、`desktop-latest`)、空白の無い `artifactName` |
| `.github/workflows/desktop-release.yml` | タグ `desktop-v*` で配布物を作り、**下書きのリリースを1つ**作って添付する (各 OS のジョブは `--publish never`) |
| `.github/workflows/desktop-update-feed.yml` | `desktop-v*` のリリースを公開したとき (または手動実行で指定したタグ) に、`latest*.yml` を `desktop-latest` に置く |
| `desktop/scripts/update-feed.mjs` | `latest*.yml` の `url` / `path` を絶対 URL に書き換える (何度実行しても同じ結果) |

### デスクトップ版をリリースする手順

1. `desktop/` で `npm version X.Y.Z --no-git-tag-version` を実行し (`package.json` と `package-lock.json`)、PR で `main` に入れる。
2. `main` のマージコミットにタグ `desktop-vX.Y.Z` を付けて push する。タグと `desktop/package.json` のバージョンが
   違うとワークフローが止まる。
3. `desktop-release.yml` の完了を待ち、下書きを確かめる。
   - 下書きが**1つだけ**で、タグが `desktop-vX.Y.Z` であること (`v` だけのタグ `vX.Y.Z` が作られていないこと)
   - 配布物の名前に空白や「.」への置き換えが無いこと (`Docx-Revision-Analyzer-X.Y.Z-…`)。dmg / mac の zip / Windows の
     インストーラと zip (x64・arm64) / それぞれの `.blockmap` / `latest-mac.yml` / `latest.yml`
   - `latest*.yml` の `url` / `path` が `…/releases/download/desktop-vX.Y.Z/…` の絶対 URL であること
   - 可能なら、macOS の zip を展開してスモークテストを行い、`app-update.yml` が `provider: generic` と `desktop-latest` の
     URL であることを確かめる:
     `DRA_SMOKE_FILE=fixtures/flow-demo.docx DRA_SMOKE_OUT=<出力先> DRA_USER_DATA=<一時フォルダ> "<展開先>/Docx Revision Analyzer.app/Contents/MacOS/Docx Revision Analyzer"`
     (`smoke: ok` と `smoke: interaction ok` が出ればよい)
4. 下書きにリリースノート (日本語と英語) を書く。
5. **公開は利用者に確認してから**行う (外部に公開する操作のため)。公開するときは「Latest」にしない:
   `gh release edit desktop-vX.Y.Z --draft=false --latest=false`
   (GitHub の画面で公開する場合は「Set as the latest release」のチェックを外す)
6. 公開すると `desktop-update-feed.yml` が動く。完了後に確かめる。
   - `gh api repos/kubohiroya/docx-revision-analyzer/releases/latest --jq .tag_name` が、ライブラリの `vX.Y.Z` のままであること
   - `desktop-latest` がプレリリースで、`latest-mac.yml` / `latest.yml` の `version` が X.Y.Z であること
     (`curl -sL https://github.com/kubohiroya/docx-revision-analyzer/releases/download/desktop-latest/latest-mac.yml`)
   - その中の配布物の URL がダウンロードできること (`curl -sIL <url>` で 200 と正しい大きさ)

### してはいけないこと

- `electron-builder` に GitHub へ公開させない (`--publish always` / `onTag` など)。CI では各 OS のジョブが並行して別々の下書きを
  作り、タグも electron-builder の既定の `vX.Y.Z` (ライブラリのリリースと同じ形) になる。
- デスクトップ版に `v` だけのタグを付けない。デスクトップ版のリリースを「Latest」にしない。
- `publish` を `github` プロバイダに戻さない。
- `artifactName` に空白を入れない (GitHub は添付するときに空白を「.」に置き換え、`latest*.yml` の名前とずれる)。
- `desktop-latest` のファイルを手で編集・削除しない。向け直すときは、Actions の「Desktop update feed」を、公開済みの
  `desktop-vX.Y.Z` を指定して手動実行する。
- 公開済みのリリースの配布物を差し替えない (チェックサムが `latest*.yml` と合わなくなる)。直すときは新しい版を出す。

### やり直し

- 下書きの段階なら作り直せる: `gh release delete desktop-vX.Y.Z --yes --cleanup-tag` で下書きとタグを消し、`main` で
  タグを付け直して push する。公開済みのリリースは消さず、新しい版で直す。
- `desktop-latest` を作り直したいときは、「Desktop update feed」を手動実行する (無ければ作る)。

## ライブラリ・CLI のリリース (参考)

- タグは `vX.Y.Z`。`package.json` のバージョンを PR で上げてから、`main` のマージコミットで GitHub Release を作る
  (このリリースが「Latest」になる)。
- 単体実行ファイルとドロップレットは `npm run build:cli:all-platforms` と `npm run build:droplet` で作り、
  `<コマンド>-<os>-<cpu>.zip` (os は `linux` / `macos` / `windows`)、`<コマンド>-macos-arm64.app.zip` にして添付する。
- `npm publish` は利用者の npm のアカウントの認証 (2段階認証) が必要なので、エージェントは実行せず、利用者に頼む
  (`npm publish --dry-run` で中身を確かめておく)。
