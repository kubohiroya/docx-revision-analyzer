// 改ざんの痕跡の検出 (src/lib/tamperEvidence.ts) のテスト。
//
// 変更履歴をロックしたテンプレートと、それをもとにした提出物のいろいろな形 (そのまま書いたもの・ロックを外した
// もの・記録をオフにして書いたもの・別の文書など) を JSZip で作り、見つかる痕跡が期待どおりかを確かめる。
// あわせて、警告を加えた SVG が整形式の XML であることと、ファイル名の付け方を確かめる。
//
//   node scripts/test-tamper-evidence.mjs
import * as esbuild from "esbuild";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { XMLValidator } = require("fast-xml-parser");

const built = await esbuild.build({
  entryPoints: [path.join(root, "src/core.ts")],
  bundle: true,
  platform: "node",
  format: "cjs",
  write: false,
  logLevel: "silent",
});
const mod = { exports: {} };
new Function("module", "exports", "require", built.outputFiles[0].text)(mod, mod.exports, require);
const DRA = mod.exports;
DRA.setLang("en");

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.error(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const LOCK = '<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="HASH1" w:salt="SALT1"/>';
const TEMPLATE_RSIDS = ["00A10001", "00A10002"];

function settingsXml({ track = true, lock = LOCK, rsidRoot = "00A10001", rsids = TEMPLATE_RSIDS } = {}) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}>` +
    `<w:zoom w:percent="100"/>${track ? "<w:trackRevisions/>" : ""}${lock ?? ""}<w:defaultTabStop w:val="840"/>` +
    `<w:rsids><w:rsidRoot w:val="${rsidRoot}"/>${rsids.map((r) => `<w:rsid w:val="${r}"/>`).join("")}</w:rsids>` +
    `</w:settings>`
  );
}

/** 本文: テンプレートの見出し (rsid 00A10002) と、続く段落 */
function documentXml(body) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
    `<w:p w:rsidR="00A10002"><w:r w:rsidR="00A10002"><w:t>Title:</w:t></w:r></w:p>` +
    body +
    `</w:body></w:document>`
  );
}

const ins = (id, rsid, text, date = `2026-06-01T10:0${id}:00Z`, author = ' w:author="Student"') =>
  `<w:ins w:id="${id}"${author}${date ? ` w:date="${date}"` : ""}><w:r w:rsidR="${rsid}"><w:t>${text}</w:t></w:r></w:ins>`;

async function docx(settings, body) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file("word/document.xml", documentXml(body));
  if (settings !== null) zip.file("word/settings.xml", settings);
  return zip.generateAsync({ type: "uint8array" });
}

const template = await docx(settingsXml(), "");
const studentRsids = [...TEMPLATE_RSIDS, "00B20001", "00B20002"];
const tracked = `<w:p w:rsidR="00B20001">${ins(1, "00B20001", "We study")}${ins(2, "00B20002", " tracked changes.")}</w:p>`;

const cases = [
  {
    name: "written as asked",
    file: await docx(settingsXml({ rsids: studentRsids }), tracked),
    expect: [],
    expectNoTemplate: [],
  },
  {
    name: "lock removed",
    file: await docx(settingsXml({ lock: "", rsids: studentRsids }), tracked),
    expect: ["lockReleased"],
    expectNoTemplate: [],
  },
  {
    name: "lock released (enforcement off)",
    file: await docx(settingsXml({ lock: LOCK.replace('w:enforcement="1"', 'w:enforcement="0"'), rsids: studentRsids }), tracked),
    expect: ["lockReleased"],
    expectNoTemplate: ["lockReleased"],
  },
  {
    name: "password replaced",
    file: await docx(settingsXml({ lock: LOCK.replace("HASH1", "HASH2"), rsids: studentRsids }), tracked),
    expect: ["lockChanged"],
    expectNoTemplate: [],
  },
  {
    name: "tracking turned off and text typed untracked",
    file: await docx(
      settingsXml({ track: false, lock: "", rsids: [...studentRsids, "00C30001"] }),
      tracked + `<w:p w:rsidR="00C30001"><w:r w:rsidR="00C30001"><w:t>Pasted without tracking.</w:t></w:r></w:p>`
    ),
    expect: ["trackingOff", "lockReleased", "untrackedText"],
    expectNoTemplate: ["trackingOff"],
  },
  {
    name: "untracked whitespace only is ignored",
    file: await docx(
      settingsXml({ rsids: [...studentRsids, "00C30001"] }),
      tracked + `<w:p w:rsidR="00C30001"><w:r w:rsidR="00C30001"><w:t xml:space="preserve">   </w:t></w:r></w:p>`
    ),
    expect: [],
    expectNoTemplate: [],
  },
  {
    name: "text deleted after being typed untracked",
    file: await docx(
      settingsXml({ rsids: [...studentRsids, "00C30001"] }),
      tracked +
        `<w:p><w:del w:id="9" w:author="Student" w:date="2026-06-01T11:00:00Z"><w:r w:rsidR="00C30001" w:rsidDel="00B20002"><w:delText>gone</w:delText></w:r></w:del></w:p>`
    ),
    expect: ["untrackedText"],
    expectNoTemplate: [],
  },
  {
    name: "a different document",
    file: await docx(settingsXml({ lock: "", rsidRoot: "00D40001", rsids: ["00D40001"] }), `<w:p>${ins(1, "00D40001", "Other file")}</w:p>`),
    expect: ["lockReleased", "notFromTemplate"],
    expectNoTemplate: [],
  },
  {
    name: "authors and dates removed",
    file: await docx(settingsXml({ rsids: studentRsids }), `<w:p>${ins(1, "00B20001", "Dated")}${ins(2, "00B20002", " undated", "", "")}</w:p>`),
    expect: ["authorDateRemoved"],
    expectNoTemplate: ["authorDateRemoved"],
  },
  {
    name: "no settings part (not judged)",
    file: await docx(null, tracked),
    expect: ["lockReleased"],
    expectNoTemplate: [],
  },
];

for (const c of cases) {
  const withTpl = (await DRA.checkTamperEvidence(c.file, { template })).evidence.map((e) => e.id).sort();
  const without = (await DRA.checkTamperEvidence(c.file)).evidence.map((e) => e.id).sort();
  check(`${c.name} (with template)`, JSON.stringify(withTpl) === JSON.stringify([...c.expect].sort()), `got ${withTpl.join(",") || "none"}`);
  check(`${c.name} (without template)`, JSON.stringify(without) === JSON.stringify([...c.expectNoTemplate].sort()), `got ${without.join(",") || "none"}`);
}

// 警告を加えた SVG
const report = await DRA.checkTamperEvidence(cases[4].file, { template });
const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="600" height="300" viewBox="0 0 600 300" font-family="Arial"><rect width="100%" height="100%" fill="#fff"/><text x="1" y="20">chart</text></svg>';
for (const lang of ["en", "ja"]) {
  const warned = DRA.addTamperWarning(svg, report, lang);
  check(`warning SVG is well-formed (${lang})`, XMLValidator.validate(warned) === true, JSON.stringify(XMLValidator.validate(warned)));
  check(`warning SVG keeps the figure (${lang})`, warned.includes('<text x="1" y="20">chart</text>') && warned.includes('id="tamper-warning"'));
  const h = Number(warned.match(/^<svg[^>]*\bheight="(\d+)"/)[1]);
  check(`warning SVG is taller than the figure (${lang})`, h > 300, String(h));
}
check("no warning without evidence", DRA.addTamperWarning(svg, { checkedAgainstTemplate: false, evidence: [], note: "" }) === svg);
check("file name for a chart", DRA.tamperedFileName("/a/report.svg") === "/a/report-tampered.svg");
check("file name for a flow", DRA.tamperedFileName("/a/report-flow.svg") === "/a/report-flow-tampered.svg");

if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("tamper evidence: all checks passed");
