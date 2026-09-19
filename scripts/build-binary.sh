#!/usr/bin/env bash
# ----------------------------------------------------------------------------
# Bun (https://bun.sh) を使って CLI を単一実行ファイル (シングルバイナリ) 化する
# ビルドスクリプト。生成された実行ファイルは Node.js が入っていない環境でも
# そのまま動作する (JSZip / fast-xml-parser / commander も含めてバイナリに
# 埋め込まれる)。
#
# 使い方:
#   ./scripts/build-binary.sh                    # chart CLI を実行中のOS/CPU向けにビルド
#   ./scripts/build-binary.sh chart               # 同上 (明示指定)
#   ./scripts/build-binary.sh score                # score (AI不正利用疑いスコア) CLI をビルド
#   ./scripts/build-binary.sh chart --all          # 主要OS/CPU向けに一括クロスビルド
#
# 出力先: dist-bin/
#
# 事前に Bun のインストールが必要です: https://bun.sh
#   curl -fsSL https://bun.sh/install | bash
# ----------------------------------------------------------------------------
set -euo pipefail

TARGET_CLI="${1:-chart}"
ALL_FLAG="${2:-}"

if [[ "$TARGET_CLI" != "chart" && "$TARGET_CLI" != "score" ]]; then
  echo "エラー: 第1引数には 'chart' または 'score' を指定してください (指定値: '$TARGET_CLI')" >&2
  exit 1
fi

if ! command -v bun >/dev/null 2>&1; then
  echo "エラー: bun が見つかりません。https://bun.sh からインストールしてください。" >&2
  echo "  curl -fsSL https://bun.sh/install | bash" >&2
  exit 1
fi

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="$ROOT_DIR/dist-bin"
mkdir -p "$OUT_DIR"

ENTRY="$ROOT_DIR/src/cli/${TARGET_CLI}.ts"
if [[ "$TARGET_CLI" == "chart" ]]; then
  BASENAME="docx-revision-chart"
else
  BASENAME="docx-ai-suspicion-score"
fi

build_one() {
  local bun_target="$1"   # 空文字列なら実行中のホスト向け
  local suffix="$2"
  local outfile="$OUT_DIR/${BASENAME}${suffix}"

  echo "==> ビルド中: ${outfile} (target: ${bun_target:-host})"
  if [[ -n "$bun_target" ]]; then
    bun build "$ENTRY" --compile --target="$bun_target" --outfile "$outfile"
  else
    bun build "$ENTRY" --compile --outfile "$outfile"
  fi
}

if [[ "$ALL_FLAG" == "--all" ]]; then
  # Bun がサポートする主要なクロスコンパイルターゲット
  # (参考: https://bun.sh/docs/bundler/executables#cross-compile-to-other-platforms)
  build_one "bun-linux-x64"    "-linux-x64"
  build_one "bun-linux-arm64"  "-linux-arm64"
  build_one "bun-darwin-x64"   "-macos-x64"
  build_one "bun-darwin-arm64" "-macos-arm64"
  build_one "bun-windows-x64"  "-windows-x64.exe"
else
  # 実行中のOS/CPU向けに1つだけビルド
  EXT=""
  if [[ "$(uname -s)" == MINGW* || "$(uname -s)" == MSYS* ]]; then
    EXT=".exe"
  fi
  build_one "" "$EXT"
fi

echo "完了: ${OUT_DIR} に出力しました。"
ls -lh "$OUT_DIR"
