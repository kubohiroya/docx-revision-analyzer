/**
 * docx-revision-analyzer のライブラリエントリポイント。
 * CLI (`docx-revision-chart` / `docx-revision-flow` / `docx-ai-suspicion-score`) からだけでなく、
 * 他の Node.js / Bun プロジェクトから直接 import して使うこともできる。
 *
 *   import { extractRevisionsFromFile, computeSuspicionScore } from "docx-revision-analyzer";
 */
export * from "./lib/docxRevisions";
export * from "./lib/insertionKinds";
export * from "./lib/historySettings";
export * from "./lib/timeBuckets";
export * from "./lib/sessions";
export * from "./lib/svgChart";
export * from "./lib/docxLayout";
export * from "./lib/flow";
export * from "./lib/flowSvg";
export * from "./lib/suspicionScore";
export * from "./lib/filenames";
