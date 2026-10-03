# ライブラリとして使う

[English](./library.md) | 日本語 · [README に戻る](../README.ja.md)

`src/index.ts` から主要な関数・型 (`extractRevisionsFromFile`, `computeSuspicionScore`,
`renderRevisionChart` など) を再エクスポートしているため、CLIとしてだけでなく
他の Node.js / Bun プロジェクトから `import { ... } from "docx-revision-analyzer"`
としてライブラリ的に使うこともできます。

**目次**

- [ブラウザで使う (`docx-revision-analyzer/core`)](#ブラウザで使う-docx-revision-analyzercore)
- [分類器](#分類器)
- [色のカテゴリ](#色のカテゴリ)
- [各編集の位置](#各編集の位置)

## ブラウザで使う (`docx-revision-analyzer/core`)

解析と描画のコードは .docx をバイト列 (`ArrayBuffer` / `Uint8Array`) で受け取り、ファイルシステムや子プロセスを
使わないため、ブラウザ・WebView・隔離されたワーカーでも動きます。`docx-revision-analyzer/core` から import してください。
既定のエントリは、これに Node.js 専用の機能 (`extractRevisionsFromFile`、ファイルを書き換える `enableHistoryPreservation`、
`buildDropOutputPath`、macOS の言語設定の読み取り) を加えたものです。

```ts
import { extractRevisions, parseDocxLayout, buildFlow, renderFlowSvg, preserveHistoryInDocx } from "docx-revision-analyzer/core";

const bytes = new Uint8Array(await file.arrayBuffer());
const data = await extractRevisions(bytes);
const svg = renderFlowSvg(buildFlow(await parseDocxLayout(bytes), { gapThresholdHours: 1, bulkChars: 150 }));
const { output } = await preserveHistoryInDocx(bytes); // 書き換えが必要なら、書き換えた .docx のバイト列
```

`npm test` は、コアを esbuild でブラウザ向けにバンドルし (Node.js のモジュールを import していれば失敗)、`process` /
`require` / `Buffer` の無い隔離環境で fixtures を処理して、Node.js での結果と一致することを確かめます。

## 分類器

ハイライトは「分類器」のパイプラインで決めます。分類器 (`{ id, version, categories?, classify(ctx) }`) は解析結果
(`ctx.positioned` の位置付きイベント、`ctx.windowsFor(options)` の挿入の窓と特徴量、`ctx.sessions` の時間区間) を受け取り、
ハイライト (カテゴリ id・対象の挿入の `w:id`・時間範囲・文書範囲・特徴量・理由の文章) を返します。`runClassifiers` は
分類器を順に実行し、挿入ごとにカテゴリの priority が最も高いハイライト (同じなら先の分類器のもの) を採用します。
既定のパイプライン (`defaultClassifiers(rules)`) は `builtin.relocation` (並べ替え・複製 → 移動) と `builtin.rules`
(判定ルールのレベル) です。JSON 出力には分類器の id とバージョン (`classifiers`) と、すべてのハイライトを理由付きで
(`highlights`) 記録します。

## 色のカテゴリ

図の色は `CategoryRegistry` (`{ id, role, color, label, priority, pattern }`) から取ります。既定のカテゴリは
細かい編集・一括挿入・移動・削除・変化なしです。`highlight` のカテゴリを登録し、レジストリを `categories` として
`renderRevisionChart` / `renderSessionedRevisionChart` / `renderFlowSvg` に渡すと、その色で描きます。挿入とカテゴリは
`RevisionEvent.category` (chart) や `FlowOptions.highlightOf` (flow) で対応づけます。

## 各編集の位置

`extractRevisionPositions` は、挿入・削除 (`w:id`) ごとに、時刻・作成者・文字数・本文と、位置 (段落インデックス・
段落内の文字オフセット・文書先頭からの文字オフセット) を持つイベント列を返します。位置は、最終文書 (すべての変更を
反映した状態) と、編集の時点 (挿入は挿入直後、削除は削除直前) の両方で求めます。オフセットは
`finalDocumentText(model)` / `documentTextAt(model, date)` (段落を `"\n"` で連結した文字列、UTF-16 単位) の中の位置です。
移動 (`w:moveFrom` / `w:moveTo`) は `move: true` として含め、`moveName` で移動元と移動先を対応づけます。
`attachRevisionPositions` で、`w:id` の対応する `RevisionEvent` に `position` を設定できます。

```ts
import { readFileSync } from "fs";
import { parseDocxLayout, extractRevisionPositions } from "docx-revision-analyzer";

const model = await parseDocxLayout(readFileSync("report.docx"));
for (const e of extractRevisionPositions(model)) {
  console.log(e.type, e.date, e.chars, e.final.start.paraIndex, e.final.start.offsetInPara);
}
```
