// テンプレートへの「変更履歴のロック」(src/lib/trackLock.ts, src/node/lockFile.ts) のテスト。
//
// パスワードの鍵が ECMA-376 の例と一致すること、settings.xml の書き換え (要素の順序・記録のオン・個人情報削除の
// 設定の除去)、パスワードの照合、ロックしたテンプレートをそのまま提出した場合に改ざんの痕跡が出ないことを確かめる。
//
//   node scripts/test-track-lock.mjs
import * as esbuild from "esbuild";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { XMLValidator } = require("fast-xml-parser");

const built = await esbuild.build({
  entryPoints: [path.join(root, "src/index.ts")],
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
async function docx(settings) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<w:document ${W}><w:body><w:p w:rsidR="00A10001"><w:r w:rsidR="00A10001"><w:t>Title:</w:t></w:r></w:p></w:body></w:document>`);
  if (settings !== null) zip.file("word/settings.xml", settings);
  return zip.generateAsync({ type: "uint8array" });
}
const settingsOf = async (bytes) => (await JSZip.loadAsync(bytes)).file("word/settings.xml").async("string");

// ECMA-376 Part 4 の例: パスワード "Example" の鍵 0x64CEED7E は、バイト順を逆にして 7EEDCE64
check("legacy key matches the ECMA-376 example", DRA.legacyPasswordKey("Example") === "7EEDCE64", DRA.legacyPasswordKey("Example"));
check("legacy key of an empty password", DRA.legacyPasswordKey("") === "00000000");
check("password is truncated to 15 characters", DRA.legacyPasswordKey("123456789012345") === DRA.legacyPasswordKey("1234567890123456789"));

// Web Crypto と node:crypto で同じハッシュ値になる
const salt = new Uint8Array(16).map((_, i) => i);
const web = await DRA.hashLockPassword("Teacher2026", salt, 1000);
const node = await DRA.hashLockPassword("Teacher2026", salt, 1000, DRA.nodeSha512);
check("Web Crypto and node:crypto agree", Buffer.from(web).equals(Buffer.from(node)) && web.length === 64);

// Word が書き出す settings.xml に近い形 (個人情報削除の設定あり、記録はオフ)
const plain = await docx(
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
    `<w:zoom w:percent="100"/><w:removePersonalInformation/><w:removeDateAndTime/><w:proofState w:spelling="clean"/>` +
    `<w:defaultTabStop w:val="840"/><w:characterSpacingControl w:val="doNotCompress"/>` +
    `<w:rsids><w:rsidRoot w:val="00A10001"/><w:rsid w:val="00A10001"/></w:rsids><w14:docId w14:val="1"/></w:settings>`
);
const r = await DRA.lockDocxBytes(plain, "Teacher2026");
const xml = await settingsOf(r.output);
check("settings.xml stays well-formed", XMLValidator.validate(xml) === true);
check("Track Changes is turned on", /<w:trackRevisions\/>/.test(xml) && !r.wasTracking);
check("personal-information settings are removed", !/removePersonalInformation|removeDateAndTime/.test(xml) && r.removedPersonalInfoSetting);
const order = ["proofState", "trackRevisions", "documentProtection", "defaultTabStop"].map((n) => xml.indexOf(`<w:${n}`));
check("elements are in schema order", order.every((x, i) => x >= 0 && (i === 0 || x > order[i - 1])), JSON.stringify(order));
check("lock attributes", /<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="[A-Za-z0-9+/=]{88}" w:salt="[A-Za-z0-9+/=]{24}"\/>/.test(xml), xml.match(/<w:documentProtection[^>]*>/)?.[0]);
check("the file is reported as locked", await DRA.isTrackChangesLocked(r.output));
check("the right password unlocks", await DRA.verifyLockPasswordNode(r.output, "Teacher2026"));
check("a wrong password doesn't", !(await DRA.verifyLockPasswordNode(r.output, "teacher2026")));

// もう一度ロックすると、ロックとパスワードが置き換わる (documentProtection は1つだけ)
const again = await DRA.lockDocxBytes(r.output, "other");
const xml2 = await settingsOf(again.output);
check("re-locking replaces the lock", again.wasLocked && (xml2.match(/documentProtection/g) ?? []).length === 1);
check("re-locking replaces the password", (await DRA.verifyLockPasswordNode(again.output, "other")) && !(await DRA.verifyLockPasswordNode(again.output, "Teacher2026")));

// パスワード無し
const nopw = await DRA.lockDocxBytes(plain, undefined);
const xml3 = await settingsOf(nopw.output);
check("lock without a password", /<w:documentProtection w:edit="trackedChanges" w:enforcement="1"\/>/.test(xml3) && (await DRA.verifyLockPasswordNode(nopw.output, "")));

// settings.xml が無い文書はロックできない
let threw = false;
try {
  await DRA.lockDocxBytes(await docx(null), "x");
} catch {
  threw = true;
}
check("a document without settings.xml can't be locked", threw);

// ロックしたテンプレートを、そのまま (書き足して) 提出した場合は痕跡が出ない
const tamper = await DRA.checkTamperEvidence(r.output, { template: r.output });
check("the locked template itself shows no trace", tamper.evidence.length === 0, tamper.evidence.map((e) => e.id).join(","));

if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("track lock: all checks passed");
