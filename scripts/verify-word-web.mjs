// Read-only preflight of deliberately edited Word for the web test files (issue #39).
// Use test documents, not student submissions. Emit metadata and flags, never document text.
import { createRequire } from "node:module";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const paths = process.argv.slice(2);
if (paths.length !== 4) {
  console.error("Usage: node scripts/verify-word-web.mjs locked-template.docx web-tracked.docx unlocked-template.docx web-tracking-off.docx");
  process.exit(2);
}
const require = createRequire(import.meta.url);
const { checkTamperEvidence, isTrackChangesLocked } = require("../dist/core.js");
const parser = new XMLParser({ preserveOrder: true, ignoreAttributes: false, removeNSPrefix: true, parseAttributeValue: false });

async function metadata(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const settings = await zip.file("word/settings.xml")?.async("string");
  const document = await zip.file("word/document.xml")?.async("string");
  const data = { settings: !!settings, root: "", rsidCount: 0, bodyRsidCount: 0, tracked: 0, labeled: 0 };
  const ids = new Set();
  function walk(nodes, isSettings) {
    if (!Array.isArray(nodes)) return;
    for (const node of nodes) {
      const attrs = node[":@"] ?? {};
      if (!isSettings && attrs["@_rsidR"]) ids.add(attrs["@_rsidR"]);
      for (const [tag, children] of Object.entries(node)) {
        if (tag === ":@" || tag === "#text") continue;
        if (isSettings && tag === "rsidRoot") data.root = attrs["@_val"] ?? "";
        if (isSettings && tag === "rsid" && attrs["@_val"]) data.rsidCount++;
        if (!isSettings && ["ins", "del"].includes(tag)) {
          data.tracked++;
          if (attrs["@_author"] && attrs["@_date"]) data.labeled++;
        }
        walk(children, isSettings);
      }
    }
  }
  if (settings) walk(parser.parse(settings), true);
  if (document) walk(parser.parse(document), false);
  data.bodyRsidCount = ids.size;
  return data;
}
const files = await Promise.all(paths.map(path => readFile(path)));
const [template, faithful, unlocked, bypass] = files;
const metas = await Promise.all(files.map(metadata));
const reports = await Promise.all([
  checkTamperEvidence(faithful, { template }),
  checkTamperEvidence(bypass, { template: unlocked }),
]);
const codes = reports.map(report => report.evidence.map(item => item.id));
const checks = {
  lockedTemplate: await isTrackChangesLocked(template),
  lockedAfterSave: await isTrackChangesLocked(faithful),
  noFalsePositive: codes[0].length === 0,
  trackedEditWithAuthorAndDate: metas[1].tracked > 0 && metas[1].tracked === metas[1].labeled,
  sameRsidRoot: !!metas[0].root && metas[0].root === metas[1].root,
  rsidsInEditedDocument: metas[1].rsidCount > 0 && metas[1].bodyRsidCount > 0,
  trackingOffWasDetected: codes[1].includes("trackingOff"),
  untrackedEditWasDetected: codes[1].includes("untrackedText"),
};
console.log(JSON.stringify({
  checks, trackedCaseFlags: codes[0], trackingOffCaseFlags: codes[1],
  metadata: metas.map(({ root, ...counts }) => ({ hasRsidRoot: !!root, ...counts })),
  scope: "File inspection only; not proof of tampering. Web UI and the next submission cycle are separate checks.",
}, null, 2));
if (Object.values(checks).some(ok => !ok)) process.exitCode = 1;
