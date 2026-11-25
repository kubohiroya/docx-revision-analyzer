import type { SavedVersionReport } from "../reportData";
import { REPORT_LABELS } from "./savedReportView";

declare const __DRA_SAVED_REPORT_SCRIPT__: string;
declare const __DRA_SAVED_REPORT_CSS__: string;

export function savedReportHtml(report: SavedVersionReport): string {
  const nonce = Array.from(crypto.getRandomValues(new Uint8Array(16)), (byte) => byte.toString(16).padStart(2, "0")).join("");
  // An inert script still uses the HTML raw-text parser. Escape angle brackets before placing JSON inside it.
  const json = JSON.stringify(report).replace(/[<>&\u2028\u2029]/gu, (char) => `\\u${char.charCodeAt(0).toString(16).padStart(4, "0")}`);
  return `<!doctype html>
<html lang="${report.lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="referrer" content="no-referrer">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src 'nonce-${nonce}'; base-uri 'none'; form-action 'none'">
<title>${REPORT_LABELS[report.lang].title}</title>
<style nonce="${nonce}">${__DRA_SAVED_REPORT_CSS__}</style>
</head>
<body class="report-document">
<main id="report"></main>
<script id="report-data" type="application/json">${json}</script>
<script nonce="${nonce}">${__DRA_SAVED_REPORT_SCRIPT__}</script>
</body>
</html>`;
}
