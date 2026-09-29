#!/usr/bin/env bash
# ----------------------------------------------------------------------------
# macOS 専用: docx-revision-chart (または docx-revision-flow) を
# Finder にドラッグ&ドロップできる ".app" (ドロップレット) にパッケージする。
#
# 使い方:
#   ./scripts/make-mac-droplet.sh            # docx-revision-chart.app
#   ./scripts/make-mac-droplet.sh flow    # docx-revision-flow.app
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
# 出力: dist-bin/docx-revision-chart.app (flow 指定時は docx-revision-flow.app)
#   ここに .docx ファイルをドラッグ&ドロップすると、同じディレクトリに
#   "<ファイル名>-<最終更新日時 YYYYMMDD-HHMMSS>.svg"
#   (flow は "<ファイル名>-flow-<最終更新日時>.svg") が作成される
#   (各ツールの --drop オプションを使用。flow は全期間が対象)。
#   完了時はmacOSの通知、エラー時はダイアログで結果を知らせる
#   (ターミナルを開かないため)。
#   文書が「変更履歴の作成者・日時を削除する」設定の
#   場合は、解析前に OK/キャンセルのダイアログで設定を有効化するか尋ねる
#   (OK なら --preserve-history を付けて実行する)。
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

TARGET_CLI="${1:-chart}"
case "$TARGET_CLI" in
  chart)
    BIN_NAME="docx-revision-chart"
    RESULT_LABEL="SVGチャート"
    ;;
  flow)
    BIN_NAME="docx-revision-flow"
    RESULT_LABEL="編集フロー図"
    ;;
  *)
    echo "エラー: 第1引数には 'chart' または 'flow' を指定してください (指定値: '$TARGET_CLI')" >&2
    exit 1
    ;;
esac

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="$ROOT_DIR/dist-bin"
APP_NAME="$BIN_NAME.app"
APP_PATH="$OUT_DIR/$APP_NAME"

mkdir -p "$OUT_DIR"

echo "==> [1/4] Bunバイナリをビルド中 (scripts/build-binary.sh $TARGET_CLI)..."
bash "$ROOT_DIR/scripts/build-binary.sh" "$TARGET_CLI"

echo "==> [2/4] AppleScriptドロップレットを作成中 (osacompile)..."
SCRIPT_SRC="$(mktemp -t "$BIN_NAME-droplet").applescript"
cat > "$SCRIPT_SRC" <<'APPLESCRIPT'
on run
	display dialog "Wordファイル (.docx) をこのアイコンにドラッグ&ドロップしてください。" & return & return & "変更履歴の記録が有効な状態で編集された .docx が対象です。" ¬
		with title "__TOOL__" buttons {"OK"} default button 1
end run

on open theFiles
	repeat with aFile in theFiles
		set posixPath to POSIX path of aFile
		set appPosix to POSIX path of (path to me)
		set toolPath to appPosix & "Contents/Resources/__TOOL__"
		set extraArgs to ""
		set subtitleText to "__RESULT_LABEL__を作成しました"
		-- 変更履歴の作成者・日時を保存する設定が無効な文書なら、有効化するか尋ねる
		try
			set checkOut to do shell script quoted form of toolPath & " --check-history-settings " & quoted form of posixPath
			if paragraph 1 of checkOut is "needs-fix" then
				set AppleScript's text item delimiters to return
				set promptText to (paragraphs 2 thru -1 of checkOut) as text
				set AppleScript's text item delimiters to ""
				try
					display dialog promptText with title "__TOOL__" buttons {"キャンセル", "OK"} default button "OK" cancel button "キャンセル" with icon caution
					set extraArgs to " --preserve-history"
					set subtitleText to "設定を有効化し、__RESULT_LABEL__を作成しました"
				on error number -128
					-- キャンセル: 設定は変えずに解析だけ行う
				end try
			end if
		end try
		try
			set outPath to do shell script quoted form of toolPath & " " & quoted form of posixPath & " --drop" & extraArgs
			display notification outPath with title "__TOOL__" subtitle subtitleText
		on error errMsg
			display dialog errMsg with title "__TOOL__ - エラー" buttons {"OK"} default button 1 with icon caution
		end try
	end repeat
end open
APPLESCRIPT
# AppleScript 内のツール名・結果の呼び名を差し込む (ヒアドキュメントは展開しない形で書いているため)
sed -i '' -e "s/__TOOL__/$BIN_NAME/g" -e "s/__RESULT_LABEL__/$RESULT_LABEL/g" "$SCRIPT_SRC"

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
