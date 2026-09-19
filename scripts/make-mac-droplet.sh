#!/usr/bin/env bash
# ----------------------------------------------------------------------------
# macOS 専用: docx-revision-chart を Finder にドラッグ&ドロップできる
# ".app" (ドロップレット) にパッケージする。
#
# 仕組み:
#   - 素の (バンドル化されていない) Unix実行ファイルは Finder の
#     ドラッグ&ドロップ対象にはなれない (macOSはドロップされたファイルを
#     Apple Event として、バンドル化されたアプリにしか配送しない)。
#   - そこで、macOS標準搭載の `osacompile` (AppleScript を .app にコンパイルする
#     ツール。追加インストール不要) を使い、"on open" ハンドラを持つ極小の
#     AppleScriptアプリ (伝統的な「ドロップレット」) を作る。
#   - Bunでコンパイルした本体バイナリは、このアプリの Contents/Resources/ に
#     同梱し、AppleScript側から `do shell script` で呼び出す。
#
# 出力: dist-bin/docx-revision-chart.app
#   ここに .docx ファイルをドラッグ&ドロップすると、
#   同じディレクトリに "<ファイル名>-<最終更新日時 YYYYMMDD-HHMMSS>.svg"
#   が作成される (docx-revision-chart の --drop オプションを使用)。
#   完了時はmacOSの通知、エラー時はダイアログで結果を知らせる
#   (ターミナルを開かないため)。
#
# 必要なもの:
#   - macOS (このスクリプト自体はmacOS上のターミナルで実行すること。
#     Claude Desktopの「接続されたコンピュータ」機能経由のサンドボックスは
#     Linuxなので、そこからは実行できない)
#   - Bun (ビルド時のみ必要。生成された .app の実行にBun/Node.jsは不要)
#   - osacompile, codesign (どちらもmacOS標準搭載)
# ----------------------------------------------------------------------------
set -euo pipefail

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "エラー: このスクリプトは macOS 専用です ($(uname -s) 上では実行できません)。" >&2
  exit 1
fi

if ! command -v osacompile >/dev/null 2>&1; then
  echo "エラー: osacompile が見つかりません (通常macOSに標準搭載されています)。" >&2
  exit 1
fi

if ! command -v bun >/dev/null 2>&1; then
  echo "エラー: bun が見つかりません。https://bun.sh からインストールしてください。" >&2
  echo "  curl -fsSL https://bun.sh/install | bash" >&2
  exit 1
fi

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="$ROOT_DIR/dist-bin"
APP_NAME="docx-revision-chart.app"
APP_PATH="$OUT_DIR/$APP_NAME"
BIN_NAME="docx-revision-chart"

mkdir -p "$OUT_DIR"

echo "==> [1/4] Bunバイナリをビルド中 (scripts/build-binary.sh chart)..."
bash "$ROOT_DIR/scripts/build-binary.sh" chart

echo "==> [2/4] AppleScriptドロップレットを作成中 (osacompile)..."
SCRIPT_SRC="$(mktemp -t docx-revision-chart-droplet).applescript"
cat > "$SCRIPT_SRC" <<'APPLESCRIPT'
on run
	display dialog "Wordファイル (.docx) をこのアイコンにドラッグ&ドロップしてください。" & return & return & "変更履歴の記録が有効な状態で編集された .docx が対象です。" ¬
		with title "docx-revision-chart" buttons {"OK"} default button 1
end run

on open theFiles
	repeat with aFile in theFiles
		set posixPath to POSIX path of aFile
		set appPosix to POSIX path of (path to me)
		set toolPath to appPosix & "Contents/Resources/docx-revision-chart"
		try
			set outPath to do shell script quoted form of toolPath & " " & quoted form of posixPath & " --drop"
			display notification outPath with title "docx-revision-chart" subtitle "SVGチャートを作成しました"
		on error errMsg
			display dialog errMsg with title "docx-revision-chart - エラー" buttons {"OK"} default button 1 with icon caution
		end try
	end repeat
end open
APPLESCRIPT

# 既存の同名 .app があると osacompile が失敗するため、事前に退避 (削除ではなくrename)。
# 通常のターミナルなので rm -rf でも問題ないが、上書き対象が壊れていた場合に
# 備えて安全側に倒している。
if [[ -e "$APP_PATH" ]]; then
  rm -rf "$APP_PATH"
fi

osacompile -o "$APP_PATH" "$SCRIPT_SRC"
rm -f "$SCRIPT_SRC"

echo "==> [3/4] 本体バイナリを同梱中..."
mkdir -p "$APP_PATH/Contents/Resources"
cp "$OUT_DIR/$BIN_NAME" "$APP_PATH/Contents/Resources/$BIN_NAME"
chmod +x "$APP_PATH/Contents/Resources/$BIN_NAME"

echo "==> [4/4] 署名中 (ad-hoc)..."
# Apple SiliconではGatekeeperの警告とは別に、有効な署名 (ad-hocで十分) が
# 無いと実行自体ができない。配布物ではなく自分用のツールなので ad-hoc署名で十分。
codesign -s - --force --deep "$APP_PATH"

echo ""
echo "完了: $APP_PATH"
echo "Finderで .docx ファイルをこのアプリのアイコンにドラッグ&ドロップしてください。"
echo "(初回起動時に「開発元を確認できないため開けません」と出た場合は、"
echo " Finderでアプリを右クリック→「開く」を選ぶと確認ダイアログ経由で起動できます。)"
