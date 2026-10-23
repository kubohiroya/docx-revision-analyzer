// 保存時のページ (issue #59) のテスト: 埋め込みフォントの取り除き (src/pages/pdfFonts.ts) と LibreOffice の呼び出し方
// (src/pages/libreoffice.ts)。
//
// pdf-lib で作った PDF (フォントの本体は中身の無いダミー) に対して、次のことを確かめる。
//   - CID フォント: 本体を消し、文字の幅 (W) と ToUnicode を残す。接頭辞 (ABCDEF+) を外し、明朝系なら serif の印を立てる
//   - 1バイトのフォント: ToUnicode から /Differences (uniXXXX) を作り直し、非記号にしてから本体を消す
//   - ToUnicode の無い 1バイトのフォントと /Type3 は触らない。何も取り除かなければ元のバイト列を返す
//   - 取り除いた後も pdf.js で開け、本文の文字が変わらない
// LibreOffice に渡す .docx は、変更履歴の印を表示しない設定 (w:revisionView) にする (settings.xml の無い文書にも)。
// LibreOffice が入っていれば (または DRA_TEST_SOFFICE で場所を指定すれば)、fixtures の .docx を実際に変換し、
// PDF の本文が変更を反映した後の本文と同じこと (削除した文字が描かれていないこと)、フォントを取り除く前と後で
// pdf.js が取り出す本文が同じことも確かめる (無ければ飛ばす)。
//
//   node test-pages.mjs
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const { PDFDocument, PDFName, PDFNumber, PDFArray, PDFDict, PDFRawStream } = require("pdf-lib");
const pdfjs = await import(pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.mjs")).href);

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-pages-"));
const bundle = async (entry, name) => {
  const out = path.join(tmp, name);
  await esbuild.build({ entryPoints: [path.join(here, entry)], bundle: true, platform: "node", format: "cjs", outfile: out, logLevel: "silent" });
  return require(out);
};
const { stripEmbeddedFonts, parseToUnicode, glyphName } = await bundle("src/pages/pdfFonts.ts", "pdfFonts.cjs");
const { exportFilter, macFontconfig, findSoffice, convertDocxToPdf, withoutRevisionMarks } = await bundle("src/pages/libreoffice.ts", "libreoffice.cjs");
const { parseDocxLayout } = await bundle("../src/lib/docxLayout.ts", "layout.cjs");
const { finalDocumentText } = await bundle("../src/lib/revisionPositions.ts", "positions.cjs");
const JSZip = require("jszip");

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.log(`NG ${name}${detail ? `: ${detail}` : ""}`);
  }
}

// ---------------------------------------------------------------------------
// ToUnicode と字の名前
// ---------------------------------------------------------------------------

const cmap = (body) =>
  `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n${body}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`;
{
  const m = parseToUnicode(
    cmap(
      "2 beginbfchar\n<01> <3042>\n<06> <D842DF9F>\nendbfchar\n" +
        "2 beginbfrange\n<03> <05> <3046>\n<10> <11> [<4E00> <0041>]\nendbfrange"
    )
  );
  check("ToUnicode: bfchar", m.get(1) === "あ");
  check("ToUnicode: bfrange (base)", m.get(3) === "う" && m.get(4) === "ぇ" && m.get(5) === "え");
  check("ToUnicode: bfrange (array)", m.get(0x10) === "一" && m.get(0x11) === "A");
  check("ToUnicode: surrogate pair", m.get(6) === "𠮟");
  check("glyph name: BMP", glyphName("あ") === "uni3042");
  check("glyph name: outside the BMP", glyphName("𠮟") === "u20B9F");
  check("glyph name: ligature", glyphName("ff") === "uni00660066");
}

// ---------------------------------------------------------------------------
// フォントを取り除く
// ---------------------------------------------------------------------------

/** 中身の無いフォントの本体 (取り除くと小さくなることを確かめるため、大きめにする) */
const dummyFont = (ctx) => ctx.register(ctx.stream(new Uint8Array(40_000).map((_, i) => (i * 7919) % 251), {}));

async function makePdf({ withStripTargets = true } = {}) {
  const doc = await PDFDocument.create();
  const ctx = doc.context;
  const fonts = {};
  const descriptor = (name, flags, fileKey = "FontFile2") =>
    ctx.register(ctx.obj({ Type: "FontDescriptor", FontName: name, Flags: flags, FontBBox: [0, -120, 1000, 880], ItalicAngle: 0, Ascent: 880, Descent: -120, CapHeight: 700, StemV: 80, [fileKey]: dummyFont(ctx) }));
  if (withStripTargets) {
    // 1バイトのフォント (LibreOffice の日本語の小分けのフォントの形)
    fonts.F1 = ctx.register(
      ctx.obj({
        Type: "Font",
        Subtype: "TrueType",
        BaseFont: "ABCDEF+HiraMinProN-W3",
        FirstChar: 1,
        LastChar: 6,
        Widths: [1000, 1000, 1000, 1000, 1000, 1000],
        FontDescriptor: descriptor("ABCDEF+HiraMinProN-W3", 4),
        ToUnicode: ctx.register(ctx.flateStream(cmap("2 beginbfchar\n<01> <3042>\n<06> <D842DF9F>\nendbfchar\n1 beginbfrange\n<02> <05> <3044>\nendbfrange"))),
      })
    );
    // CID フォント (明朝系)
    const cid = (prefixed, flags) =>
      ctx.register(
        ctx.obj({
          Type: "Font",
          Subtype: "Type0",
          BaseFont: prefixed,
          Encoding: "Identity-H",
          DescendantFonts: [
            ctx.register(
              ctx.obj({
                Type: "Font",
                Subtype: "CIDFontType2",
                BaseFont: prefixed,
                CIDSystemInfo: { Registry: ctx.obj("Adobe"), Ordering: ctx.obj("Identity"), Supplement: 0 },
                FontDescriptor: descriptor(prefixed, flags),
                W: [1, [1000, 1000, 500]],
                CIDToGIDMap: "Identity",
              })
            ),
          ],
          ToUnicode: ctx.register(ctx.flateStream(cmap("3 beginbfchar\n<0001> <6587>\n<0002> <66F8>\n<0003> <0041>\nendbfchar"))),
        })
      );
    fonts.F2 = cid("GHIJKL+NotoSerifJP-Regular", 4);
    fonts.F3 = cid("MNOPQR+NotoSansJP-Regular", 4);
  }
  // ToUnicode の無い 1バイトのフォント (取り除かない)
  fonts.F4 = ctx.register(
    ctx.obj({ Type: "Font", Subtype: "TrueType", BaseFont: "STUVWX+Foo", FirstChar: 65, LastChar: 66, Widths: [600, 600], FontDescriptor: descriptor("STUVWX+Foo", 32) })
  );
  // Type3 (触らない)
  fonts.F5 = ctx.register(
    ctx.obj({
      Type: "Font",
      Subtype: "Type3",
      FontBBox: [0, 0, 1000, 1000],
      FontMatrix: [0.001, 0, 0, 0.001, 0, 0],
      CharProcs: { g1: ctx.register(ctx.stream("1000 0 0 0 1000 1000 d1 0 0 1000 1000 re f", {})) },
      Encoding: { Type: "Encoding", Differences: [1, PDFName.of("g1")] },
      FirstChar: 1,
      LastChar: 1,
      Widths: [1000],
    })
  );
  const page = doc.addPage([300, 200]);
  const ops = [
    withStripTargets ? "BT /F1 12 Tf 10 170 Td <010203040506> Tj ET" : "",
    withStripTargets ? "BT /F2 12 Tf 10 140 Td <000100020003> Tj ET" : "",
    withStripTargets ? "BT /F3 12 Tf 10 110 Td <00010002> Tj ET" : "",
    "BT /F4 12 Tf 10 80 Td (AB) Tj ET",
    "BT /F5 12 Tf 10 50 Td <01> Tj ET",
  ].filter(Boolean);
  page.node.set(PDFName.of("Resources"), ctx.obj({ Font: fonts }));
  page.node.set(PDFName.of("Contents"), ctx.register(ctx.flateStream(ops.join("\n"))));
  return { bytes: await doc.save({ useObjectStreams: false }), fontNames: Object.keys(fonts) };
}

/** PDF のフォントの辞書 (リソースの名前ごと) */
async function fontsOf(bytes) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  const ctx = doc.context;
  const res = doc.getPage(0).node.Resources();
  const fontRes = ctx.lookup(res.get(PDFName.of("Font")));
  const out = {};
  for (const [k, v] of fontRes.entries()) {
    const font = ctx.lookup(v);
    const desc = font.get(PDFName.of("DescendantFonts"));
    const cidFont = desc ? ctx.lookup(ctx.lookup(desc).get(0)) : undefined;
    const fd = ctx.lookup((cidFont ?? font).get(PDFName.of("FontDescriptor")));
    out[k.asString().slice(1)] = { font, cidFont, fd, ctx };
  }
  return out;
}

const hasFontFile = (fd) => ["FontFile", "FontFile2", "FontFile3"].some((k) => fd?.has(PDFName.of(k)));
const name = (d, key) => d?.get(PDFName.of(key))?.toString();

/** pdf.js で取り出した本文 (行ごと) */
async function textOf(bytes) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false, isEvalSupported: false, disableFontFace: true, verbosity: 0 }).promise;
  const lines = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const tc = await page.getTextContent();
    lines.push(tc.items.map((it) => it.str).join(""));
  }
  await doc.destroy();
  return lines.join("\n").replace(/\s+/g, "");
}

{
  const { bytes } = await makePdf();
  const r = await stripEmbeddedFonts(bytes);
  check("strip: counts", r.stripped === 3 && r.reencoded === 1 && r.kept === 1, JSON.stringify({ ...r, bytes: undefined }));
  check("strip: smaller", r.bytes.byteLength < bytes.byteLength / 3, `${bytes.byteLength} → ${r.bytes.byteLength}`);
  const f = await fontsOf(r.bytes);

  // 1バイトのフォント
  check("simple font: font program removed", !hasFontFile(f.F1.fd));
  check("simple font: prefix removed", name(f.F1.font, "BaseFont") === "/HiraMinProN-W3" && name(f.F1.fd, "FontName") === "/HiraMinProN-W3");
  const enc = f.F1.ctx.lookup(f.F1.font.get(PDFName.of("Encoding")));
  const diffs = enc instanceof PDFDict ? f.F1.ctx.lookup(enc.get(PDFName.of("Differences"))) : undefined;
  const diffText = diffs instanceof PDFArray ? diffs.asArray().map((x) => x.toString()).join(" ") : "";
  check(
    "simple font: /Differences from ToUnicode",
    diffText === "1 /uni3042 2 /uni3044 3 /uni3045 4 /uni3046 5 /uni3047 6 /u20B9F",
    diffText
  );
  check("simple font: nonsymbolic and serif", name(f.F1.fd, "Flags") === String(32 | 2), name(f.F1.fd, "Flags"));
  check("simple font: ToUnicode kept", f.F1.font.has(PDFName.of("ToUnicode")));

  // CID フォント
  check("CID font: font program removed", !hasFontFile(f.F2.fd) && !hasFontFile(f.F3.fd));
  check(
    "CID font: prefix removed (Type0 and CIDFont)",
    name(f.F2.font, "BaseFont") === "/NotoSerifJP-Regular" && name(f.F2.cidFont, "BaseFont") === "/NotoSerifJP-Regular"
  );
  check("CID font: widths and ToUnicode kept", f.F2.cidFont.has(PDFName.of("W")) && f.F2.font.has(PDFName.of("ToUnicode")));
  check("CID font: serif flag for mincho / serif names", name(f.F2.fd, "Flags") === String(4 | 2), name(f.F2.fd, "Flags"));
  check("CID font: no serif flag for sans names", name(f.F3.fd, "Flags") === "4", name(f.F3.fd, "Flags"));

  // 触らないもの
  check("simple font without ToUnicode: kept", hasFontFile(f.F4.fd) && name(f.F4.font, "BaseFont") === "/STUVWX+Foo");
  check("Type3: untouched", f.F5.font.has(PDFName.of("CharProcs")) && name(f.F5.font, "Subtype") === "/Type3");

  // pdf.js で開けて、本文が変わらない
  const text = await textOf(r.bytes);
  check("pdf.js: text after stripping", text.startsWith("あいぅうぇ𠮟文書A文書AB"), text);
}

{
  const { bytes } = await makePdf({ withStripTargets: false });
  const r = await stripEmbeddedFonts(bytes);
  check("nothing to strip: input returned as is", r.stripped === 0 && r.kept === 1 && r.bytes === bytes);
}

// ---------------------------------------------------------------------------
// LibreOffice の呼び出し方
// ---------------------------------------------------------------------------

check("export filter: high keeps images", exportFilter("high") === "pdf:writer_pdf_Export");
for (const [q, dpi] of [
  ["medium", "150"],
  ["low", "96"],
]) {
  const f = exportFilter(q);
  const opts = JSON.parse(f.slice("pdf:writer_pdf_Export:".length));
  check(
    `export filter: ${q} reduces images to ${dpi} dpi`,
    f.startsWith("pdf:writer_pdf_Export:{") && opts.ReduceImageResolution.value === "true" && opts.MaxImageResolution.value === dpi
  );
}
{
  const conf = macFontconfig("/Applications/LibreOffice.app/Contents/Resources/fonts", "/tmp/a&b");
  check("fontconfig: macOS font folders", conf.includes("<dir>/System/Library/Fonts</dir>") && conf.includes("<dir>~/Library/Fonts</dir>"));
  check("fontconfig: LibreOffice's fonts", conf.includes("<dir>/Applications/LibreOffice.app/Contents/Resources/fonts</dir>"));
  check("fontconfig: escapes", conf.includes("<cachedir>/tmp/a&amp;b</cachedir>"));
  check(
    "fontconfig: Windows fonts → Hiragino",
    conf.includes("<family>ＭＳ 明朝</family><prefer><family>Hiragino Mincho ProN</family>") &&
      conf.includes("<family>Yu Gothic</family><prefer><family>Hiragino Sans</family>")
  );
}
{
  const fake = path.join(tmp, "soffice");
  fs.writeFileSync(fake, "");
  check("find soffice: the configured path first", findSoffice(fake) === fake);
  check("find soffice: a missing configured path is skipped", findSoffice(path.join(tmp, "missing")) !== path.join(tmp, "missing"));
}

// ---------------------------------------------------------------------------
// 変更履歴の印を表示しない設定
// ---------------------------------------------------------------------------

const fixtureDocx = new Uint8Array(fs.readFileSync(path.join(here, "..", "fixtures", "flow-demo.docx")));
{
  // settings.xml の無い文書 (fixture): settings.xml と、その関係・種類を加える
  const zip = await JSZip.loadAsync(await withoutRevisionMarks(fixtureDocx));
  const settings = await zip.file("word/settings.xml")?.async("string");
  const rels = await zip.file("word/_rels/document.xml.rels").async("string");
  const types = await zip.file("[Content_Types].xml").async("string");
  check("no settings.xml: added with revisionView", !!settings && settings.includes(`<w:revisionView w:markup="0" w:insDel="0" w:formatting="0"/>`));
  check("no settings.xml: relationship and content type", rels.includes(`relationships/settings" Target="settings.xml"`) && types.includes(`PartName="/word/settings.xml"`));
  // settings.xml のある文書: 既存の revisionView を置き換え、trackRevisions の前に置く
  const z2 = await JSZip.loadAsync(fixtureDocx);
  z2.file(
    "word/settings.xml",
    `<?xml version="1.0"?><x:settings xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:zoom x:percent="100"/>` +
      `<x:revisionView x:insDel="1"/><x:trackRevisions/><x:defaultTabStop x:val="720"/></x:settings>`
  );
  const z3 = await JSZip.loadAsync(await withoutRevisionMarks(await z2.generateAsync({ type: "uint8array" })));
  const s3 = await z3.file("word/settings.xml").async("string");
  check(
    "settings.xml: revisionView replaced before trackRevisions",
    s3.includes(`<x:zoom x:percent="100"/><x:revisionView x:markup="0" x:insDel="0" x:formatting="0"/><x:trackRevisions/>`) && s3.split("revisionView").length === 2,
    s3
  );
}

// ---------------------------------------------------------------------------
// 実際の変換 (LibreOffice があるときだけ)
// ---------------------------------------------------------------------------

const soffice = findSoffice(process.env.DRA_TEST_SOFFICE);
if (!soffice) {
  console.log("skip LibreOffice conversion (LibreOffice was not found; set DRA_TEST_SOFFICE to its soffice)");
} else {
  const docx = fixtureDocx;
  const workDir = path.join(tmp, "work");
  fs.mkdirSync(workDir);
  const t0 = Date.now();
  const pdf = await convertDocxToPdf(docx, { soffice, workDir, quality: "medium" });
  check("LibreOffice: converted", pdf.byteLength > 1000 && Buffer.from(pdf.slice(0, 5)).toString() === "%PDF-", `${pdf.byteLength} bytes`);
  const r = await stripEmbeddedFonts(pdf);
  check("LibreOffice: fonts stripped", r.stripped > 0, JSON.stringify({ ...r, bytes: undefined }));
  check("LibreOffice: smaller after stripping", r.bytes.byteLength < pdf.byteLength, `${pdf.byteLength} → ${r.bytes.byteLength}`);
  const [before, after] = [await textOf(pdf), await textOf(r.bytes)];
  check("LibreOffice: text unchanged by stripping", before.length > 20 && before === after, `${before.slice(0, 40)} / ${after.slice(0, 40)}`);
  check("LibreOffice: Japanese text present", /[぀-ヿ一-鿿]/.test(after), after.slice(0, 40));
  const final = finalDocumentText(await parseDocxLayout(docx)).replace(/\s+/g, "");
  check("LibreOffice: the text after the changes (no deleted text drawn)", after === final, `${after.length} chars in the PDF, ${final.length} in the document`);
  console.log(`  (LibreOffice ${soffice}: ${Date.now() - t0} ms, ${pdf.byteLength} → ${r.bytes.byteLength} bytes)`);
}

fs.rmSync(tmp, { recursive: true, force: true });
if (failed) {
  console.log(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("all checks passed");
