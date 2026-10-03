# Using it as a library

English | [日本語](./library.ja.md) · [Back to the README](../README.md)

`src/index.ts` re-exports the main functions and types (`extractRevisionsFromFile`,
`computeSuspicionScore`, `renderRevisionChart`, etc.), so beyond the CLIs, the
package can also be used as a library from another Node.js / Bun project via
`import { ... } from "docx-revision-analyzer"`.

**Contents**

- [In the browser (`docx-revision-analyzer/core`)](#in-the-browser-docx-revision-analyzercore)
- [Classifiers](#classifiers)
- [Color categories](#color-categories)
- [Where each edit happened](#where-each-edit-happened)

## In the browser (`docx-revision-analyzer/core`)

The analysis and drawing code takes the .docx as bytes (`ArrayBuffer` / `Uint8Array`) and doesn't use the file
system or child processes, so it also runs in a browser, a WebView or a sandboxed worker. Import it from
`docx-revision-analyzer/core`; the default entry adds the Node.js-only helpers (`extractRevisionsFromFile`,
`enableHistoryPreservation` that rewrites a file in place, `buildDropOutputPath`, and reading the macOS language
setting).

```ts
import { extractRevisions, parseDocxLayout, buildFlow, renderFlowSvg, preserveHistoryInDocx } from "docx-revision-analyzer/core";

const bytes = new Uint8Array(await file.arrayBuffer());
const data = await extractRevisions(bytes);
const svg = renderFlowSvg(buildFlow(await parseDocxLayout(bytes), { gapThresholdHours: 1, bulkChars: 150 }));
const { output } = await preserveHistoryInDocx(bytes); // the fixed .docx as bytes, if it needed fixing
```

`npm test` bundles the core for the browser with esbuild (failing if it imports a Node.js module) and runs it on
the fixtures in a sandbox without `process` / `require` / `Buffer`, checking that the results match Node.js.

## Classifiers

Highlights are produced by a pipeline of *classifiers*. A classifier (`{ id, version, categories?, classify(ctx) }`)
receives the analysis (`ctx.positioned` events, `ctx.windowsFor(options)` insertion windows with features,
`ctx.sessions` time ranges) and returns highlights: category id, target insertions (`w:id`), time range, document
range, features and a reason. `runClassifiers` runs them in order and, for each insertion, keeps the highlight
whose category has the highest priority (ties: the earlier classifier). The built-in pipeline
(`defaultClassifiers(rules)`) is `builtin.relocation` (reordering/duplication → moves) followed by `builtin.rules`
(the rule levels). The JSON output records the classifiers' ids and versions (`classifiers`) and every highlight
with its reason (`highlights`).

## Color categories

The colors in the figures come from a `CategoryRegistry` (`{ id, role, color, label, priority, pattern }`). The
built-in categories are fine-grained edits, bulk insertion, moves, deletions and unchanged; register more
`highlight` categories and pass the registry as `categories` to `renderRevisionChart` /
`renderSessionedRevisionChart` / `renderFlowSvg`. Insertions are tied to a category through
`RevisionEvent.category` (chart) or `FlowOptions.highlightOf` (flow).

## Where each edit happened

`extractRevisionPositions` returns one event per insertion/deletion (`w:id`) with its time, author, character
count, text, and position: paragraph index, offset within the paragraph, and offset from the start of the
document, both in the final document (all changes accepted) and at the moment of the edit (right after an
insertion / right before a deletion). Offsets index into `finalDocumentText(model)` / `documentTextAt(model, date)`
(paragraphs joined with `"\n"`, counted in UTF-16 code units). Moves (`w:moveFrom` / `w:moveTo`) are included with
`move: true` and paired by `moveName`. `attachRevisionPositions` sets `position` on matching `RevisionEvent`s.

```ts
import { readFileSync } from "fs";
import { parseDocxLayout, extractRevisionPositions } from "docx-revision-analyzer";

const model = await parseDocxLayout(readFileSync("report.docx"));
for (const e of extractRevisionPositions(model)) {
  console.log(e.type, e.date, e.chars, e.final.start.paraIndex, e.final.start.offsetInPara);
}
```
