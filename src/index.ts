/**
 * docx-revision-analyzer のライブラリエントリポイント (Node.js 用)。
 * CLI (`docx-revision-chart` / `docx-revision-flow` / `docx-ai-suspicion-score`) からだけでなく、
 * 他の Node.js / Bun プロジェクトから直接 import して使うこともできる。
 *
 *   import { extractRevisionsFromFile, computeSuspicionScore } from "docx-revision-analyzer";
 *
 * ファイルシステムを使わないコアだけが必要な場合 (ブラウザ等) は "docx-revision-analyzer/core" を使う。
 */
import "./node/locale";

export * from "./core";
export * from "./node/files";
export * from "./node/historyFile";
