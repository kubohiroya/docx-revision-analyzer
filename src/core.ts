/**
 * docx-revision-analyzer のコア (ファイルシステム・子プロセスを使わない部分)。
 * ブラウザ・WebView・拡張の隔離環境でも動く。docx は ArrayBuffer / Uint8Array で渡す。
 *
 *   import { extractRevisions, parseDocxLayout, renderFlowSvg } from "docx-revision-analyzer/core";
 */
export * from "./lib/input";
export * from "./lib/docxRevisions";
export * from "./lib/insertionKinds";
export * from "./lib/historySettings";
export * from "./lib/timeBuckets";
export * from "./lib/sessions";
export * from "./lib/svgChart";
export * from "./lib/docxLayout";
export * from "./lib/revisionPositions";
export * from "./lib/insertionWindows";
export * from "./lib/insertionRules";
export * from "./lib/categories";
export * from "./lib/classifiers";
export * from "./lib/integrity";
export * from "./lib/tamperEvidence";
export * from "./lib/trackLock";
export * from "./lib/snapshots";
export * from "./lib/figureTargets";
export * from "./lib/flow";
export * from "./lib/flowSvg";
export * from "./lib/suspicionScore";
export * from "./lib/filenames";
export { t, getLang, setLang, detectLang, langFromLocale, setSystemLocaleProvider, SUPPORTED_LANGS } from "./lib/i18n";
export type { Lang } from "./lib/i18n";
