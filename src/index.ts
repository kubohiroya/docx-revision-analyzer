/**
 * docx-revision-tools のライブラリエントリポイント。
 * CLI (`docx-revision-chart` / `docx-ai-suspicion-score`) からだけでなく、
 * 他の Node.js / Bun プロジェクトから直接 import して使うこともできる。
 *
 *   import { extractRevisionsFromFile, computeSuspicionScore } from "docx-revision-tools";
 */
export * from "./lib/docxRevisions";
export * from "./lib/historySettings";
export * from "./lib/timeBuckets";
export * from "./lib/sessions";
export * from "./lib/svgChart";
export * from "./lib/docxLayout";
export * from "./lib/heatmap";
export * from "./lib/heatmapSvg";
export * from "./lib/suspicionScore";
export * from "./lib/filenames";
