// 自動更新の情報 (latest-mac.yml / latest.yml / latest-linux.yml) の、配布物の名前 (url / path) を
// リリースのダウンロード用の絶対 URL に書き換える。
//
// 自動更新は、デスクトップ版専用の固定の場所 (プレリリース desktop-latest) に置いたこれらのファイルを読む
// (electron-builder.config.cjs の publish)。リポジトリの「Latest」のリリース (ライブラリの vX.Y.Z) とは取り違えない。
// ファイルは desktop-latest に置き、配布物そのものは各版のリリース (desktop-vX.Y.Z) から落とすので、絶対 URL にする。
//
//   node scripts/update-feed.mjs <タグ (desktop-vX.Y.Z)> <ファイル...> [--repo owner/repo]
import * as fs from "node:fs";

const args = process.argv.slice(2);
const repoAt = args.indexOf("--repo");
const repo = repoAt >= 0 ? args.splice(repoAt, 2)[1] : "kubohiroya/docx-revision-analyzer";
const [tag, ...files] = args;
if (!/^desktop-v\d+\.\d+\.\d+/.test(tag ?? "") || files.length === 0) {
  console.error("usage: node scripts/update-feed.mjs <desktop-vX.Y.Z> <latest*.yml...> [--repo owner/repo]");
  process.exit(2);
}
const base = `https://github.com/${repo}/releases/download/${tag}/`;
for (const file of files) {
  const before = fs.readFileSync(file, "utf-8");
  // "  - url: <名前>" と "path: <名前>" (すでに URL のものはそのまま)
  const after = before.replace(/^(\s*(?:-\s+)?(?:url|path):\s*)(['"]?)([^'"\s]+)\2\s*$/gm, (m, head, _q, name) =>
    /^https?:\/\//.test(name) ? m : `${head}${base}${encodeURI(name)}`
  );
  if (!after.includes(base)) {
    console.error(`${file}: no url / path to rewrite`);
    process.exit(1);
  }
  fs.writeFileSync(file, after);
  console.log(`${file}: -> ${base}`);
}
