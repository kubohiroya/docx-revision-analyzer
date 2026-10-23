// 模式的なページの割り付け (src/lib/pageLayout.ts, src/lib/docxLayout.ts) のテスト。
//
// 機能ごとに小さな文書を JSZip で作り、layoutDocument の結果 (ページ・行・図・セルの位置) を確かめる。
//   - 段組み (2段組み、連続のセクション区切りでの段数の切り替えと段の高さの均等化、段区切り)
//   - セクションごとの用紙の向き、文字数と行数 (docGrid)、設定の変更の履歴 (w:sectPrChange)
//   - 段落の書式 (間隔・インデント・スタイルの継承)、書式の変更の履歴 (w:pPrChange)
//   - ページ送りの規則 (次の段落と分離しない・1行残して区切らない)、脚注
//   - 表 (列の幅・行の高さ・行の挿入)、図 (幅・浮動配置の回り込み)
//   - Word が保存した改ページ位置 (w:lastRenderedPageBreak)、縦書き
//
//   node scripts/test-page-layout.mjs
import * as esbuild from "esbuild";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dra-layout-"));
const lib = path.join(tmp, "core.cjs");
await esbuild.build({
  entryPoints: [path.join(root, "src/core.ts")],
  bundle: true,
  platform: "node",
  format: "cjs",
  outfile: lib,
  logLevel: "silent",
});
const { parseDocxLayout, layoutDocument, buildFlow, renderFlowSvg } = require(lib);

let failed = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`ok ${name}`);
  else {
    failed++;
    console.error(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const T1 = "2026-06-01T10:00:00Z";
const T2 = "2026-06-02T10:00:00Z";
let id = 0;
const rev = (date) => `w:id="${++id}" w:author="Writer" w:date="${date}"`;
const run = (text, rPr = "") => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (text, pPr = "", extra = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${text ? run(text) : ""}${extra}</w:p>`;
/** 全角 n 文字 */
const jp = (n) => "あいうえおかきくけこさしすせそたちつてと".repeat(Math.ceil(n / 20)).slice(0, n);
/** A4、余白 1701 (本文の幅 8504)、行送り 360 */
const sect = (inner = "", extra = "") =>
  `<w:sectPr>${extra}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>${inner}</w:sectPr>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Indented"><w:name w:val="Indented"/><w:basedOn w:val="a"/><w:pPr><w:ind w:left="840"/></w:pPr></w:style>
</w:styles>`;

async function model(body, { footnotes, app } = {}) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  zip.file("word/styles.xml", STYLES);
  if (footnotes) {
    zip.file(
      "word/footnotes.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NS}>` +
        `<w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>` +
        Object.entries(footnotes)
          .map(([fid, text]) => `<w:footnote w:id="${fid}">${p(text)}</w:footnote>`)
          .join("") +
        `</w:footnotes>`
    );
  }
  if (app) zip.file("docProps/app.xml", `<Properties><Pages>${app}</Pages></Properties>`);
  return parseDocxLayout(await zip.generateAsync({ type: "uint8array" }));
}

const lines = (page) => page.items.filter((it) => it.type === "line");
const ofPara = (pages, i) =>
  pages.flatMap((pg, n) => pg.items.filter((it) => it.paraIndex === i).map((it) => ({ ...it, page: n })));
const near = (a, b, tol = 2) => Math.abs(a - b) <= tol;

// ---------------------------------------------------------------------------
// 2段組み
{
  const body = Array.from({ length: 60 }, () => p(jp(40))).join("") + sect(`<w:cols w:num="2" w:space="425" w:sep="1"/><w:docGrid w:linePitch="360"/>`);
  const m = await model(body);
  const { pages } = layoutDocument(m, Infinity);
  const xs = [...new Set(lines(pages[0]).map((it) => Math.round(it.x)))].sort((a, b) => a - b);
  const colW = (8504 - 425) / 2;
  check("two columns: two column positions on page 1", xs.length === 2 && near(xs[0], 1701) && near(xs[1], 1701 + colW + 425), JSON.stringify(xs));
  check("two columns: line width is the column width", lines(pages[0]).every((it) => near(it.w, colW, 1)));
  const right = lines(pages[0]).filter((it) => it.x > 5000);
  check("two columns: right column starts at the top", near(Math.min(...right.map((it) => it.y)), 1985));
  check("two columns: separator line", (pages[0].rules ?? []).length === 1);
  // 1段組みより約半分のページ数 (1段: 1行約40字。2段: 1行約19字で行数は倍、段は2つ)
  const single = layoutDocument(await model(Array.from({ length: 60 }, () => p(jp(40))).join("") + sect(`<w:docGrid w:linePitch="360"/>`)), Infinity).pages.length;
  check("two columns: page count is close to single column", Math.abs(pages.length - single) <= 1, `${pages.length} vs ${single}`);

  // 段区切り
  const brk = await model(p("first") + `<w:p><w:r><w:br w:type="column"/></w:r>${run("second")}</w:p>` + sect(`<w:cols w:num="2" w:space="425"/>`));
  const items = ofPara(layoutDocument(brk, Infinity).pages, 1);
  check("column break: text after the break is in the second column", items.some((it) => it.x > 5000 && near(it.y, 1985)));
}

// 表題は1段、本文は2段 (連続のセクション区切り) と、段の高さの均等化
{
  const body =
    p("Title", `<w:jc w:val="center"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>`) +
    Array.from({ length: 6 }, () => p(jp(60))).join("") +
    p("After", "", "") +
    sect(`<w:cols w:num="2" w:space="425"/>`, `<w:type w:val="continuous"/>`).replace("<w:sectPr>", `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:num="2" w:space="425"/></w:sectPr></w:pPr></w:p><w:sectPr>`);
  const m = await model(body);
  const { pages } = layoutDocument(m, Infinity);
  check("continuous: one page", pages.length === 1, String(pages.length));
  const title = ofPara(pages, 0)[0];
  const bodyLines = pages[0].items.filter((it) => it.paraIndex >= 1);
  const left = bodyLines.filter((it) => it.x < 5000);
  const right = bodyLines.filter((it) => it.x > 5000);
  check("continuous: title spans the full width", near(title.w, 8504, 1));
  check("continuous: body below the title in two columns", left.length > 0 && right.length > 0 && Math.min(...right.map((it) => it.y)) >= title.y + title.h - 1);
  const bottom = (a) => Math.max(...a.map((it) => it.y + it.h));
  check("continuous: columns balanced", Math.abs(bottom(left) - bottom(right)) <= 400, `${bottom(left)} vs ${bottom(right)}`);
  const after = ofPara(pages, 7)[0];
  check("continuous: following paragraph", after !== undefined);
}

// セクションごとの用紙の向き、設定の変更の履歴
{
  const landscape = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const body = p("Portrait", `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`) + p("Landscape") + landscape;
  const { pages } = layoutDocument(await model(body), Infinity);
  check("sections: portrait then landscape page", pages.length === 2 && pages[0].width === 11906 && pages[1].width === 16838, JSON.stringify(pages.map((pg) => pg.width)));
  const wide = ofPara(pages, 1)[0];
  check("sections: landscape text width", near(wide.w, 16838 - 2880, 1));

  const changed =
    p("Text") +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="2880" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:sectPrChange ${rev(T1)}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:sectPrChange></w:sectPr>`;
  const mc = await model(changed);
  const before = ofPara(layoutDocument(mc, Date.parse(T1) - 1).pages, 0)[0];
  const after = ofPara(layoutDocument(mc, Infinity).pages, 0)[0];
  check("sectPrChange: old margin before the change", near(before.x, 1440) && near(after.x, 2880), `${before.x} / ${after.x}`);
}

// 文字数と行数 (docGrid)
{
  // 本文の幅 8504、10.5pt。40字: 文字送り 212.6 → charSpace = (212.6 - 210) * 4096 / 20
  const grid = (chars) => {
    const cs = Math.round(((8504 / chars - 210) * 4096) / 20);
    return sect(`<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="${cs}"/>`);
  };
  const count = async (chars) => lines(layoutDocument(await model(p(jp(80)) + grid(chars)), Infinity).pages[0]).length;
  check("docGrid: 40 chars per line → 80 chars in 2 lines", (await count(40)) === 2, String(await count(40)));
  check("docGrid: 30 chars per line → 80 chars in 3 lines", (await count(30)) === 3, String(await count(30)));
  const { pages } = layoutDocument(await model(p("a") + p("b") + grid(40)), Infinity);
  check("docGrid: line pitch", near(lines(pages[0])[1].y - lines(pages[0])[0].y, 360));
}

// 段落の書式 (間隔・インデント・スタイルの継承)、書式の変更の履歴
{
  const body =
    p("one") +
    p("two", `<w:spacing w:before="480" w:after="240"/>`) +
    p("three") +
    p("styled", `<w:pStyle w:val="Indented"/>`) +
    p("direct", `<w:ind w:left="1440" w:firstLine="420"/>`) +
    `<w:p><w:pPr><w:pStyle w:val="1"/><w:pPrChange ${rev(T1)}><w:pPr/></w:pPrChange></w:pPr>${run("Heading later")}</w:p>` +
    sect(`<w:docGrid w:linePitch="360"/>`);
  const m = await model(body);
  const { pages } = layoutDocument(m, Infinity);
  const [a, b, c, d, e, f] = [0, 1, 2, 3, 4, 5].map((i) => ofPara(pages, i)[0]);
  check("spacing: before", near(b.y - (a.y + a.h), 480), String(b.y - (a.y + a.h)));
  check("spacing: after", near(c.y - (b.y + b.h), 240), String(c.y - (b.y + b.h)));
  check("indent: from the style (basedOn)", near(d.x, 1701 + 840));
  check("indent: left + first line", near(e.x, 1701 + 1440 + 420));
  check("heading style: heading after the change", f.style === "heading");
  const old = ofPara(layoutDocument(m, Date.parse(T1) - 1).pages, 5)[0];
  check("pPrChange: body text before the change", old.style === "body");
}

// ページ送りの規則
{
  // 1ページ 36 行 (本文の高さ 16838-1985-1701 = 13152、行送り 360)
  const filler = (n) => Array.from({ length: n }, () => p("x")).join("");
  const keep = await model(filler(34) + p("Heading", `<w:pStyle w:val="1"/>`) + p(jp(40)) + sect(`<w:docGrid w:type="lines" w:linePitch="360"/>`));
  const kp = layoutDocument(keep, Infinity).pages;
  check("keepNext: heading moves with the next paragraph", ofPara(kp, 34)[0].page === ofPara(kp, 35)[0].page && ofPara(kp, 34)[0].page === 1);

  const widow = await model(filler(35) + p(jp(40 * 3)) + sect(`<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="${Math.round(((8504 / 40 - 210) * 4096) / 20)}"/>`));
  const wp = ofPara(layoutDocument(widow, Infinity).pages, 35);
  check("widowControl: a single first line is not left at the bottom", wp.every((it) => it.page === 1), JSON.stringify(wp.map((it) => it.page)));

  const lines3 = await model(filler(34) + p(jp(40 * 3)) + sect(`<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="${Math.round(((8504 / 40 - 210) * 4096) / 20)}"/>`));
  const lp = ofPara(layoutDocument(lines3, Infinity).pages, 34).map((it) => it.page);
  check("widowControl: a single last line is not carried over", lp.filter((n) => n === 0).length === 1 && lp.filter((n) => n === 1).length === 2, JSON.stringify(lp));
}

// 脚注
{
  const body = `<w:p>${run("text")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` + p("more") + sect(`<w:docGrid w:linePitch="360"/>`);
  const m = await model(body, { footnotes: { 1: jp(200) } });
  const { pages } = layoutDocument(m, Infinity);
  const notes = pages[0].items.filter((it) => it.type === "note");
  check("footnote: note lines at the bottom of the page", notes.length >= 4 && near(Math.max(...notes.map((it) => it.y + it.h)), 16838 - 1701, 2), String(notes.length));
}

// 表
{
  const cell = (text, w) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
  const body =
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${cell("A", 2000)}${cell("B", 2000)}${cell(jp(30), 2000)}</w:tr>` +
    `<w:tr><w:trPr><w:ins ${rev(T1)}/></w:trPr>${cell("D", 2000)}${cell("E", 2000)}${cell("F", 2000)}</w:tr>` +
    `</w:tbl>` +
    p("after") +
    sect(`<w:docGrid w:linePitch="360"/>`);
  const m = await model(body);
  const { pages } = layoutDocument(m, Infinity);
  const [A, B, C] = [0, 1, 2].map((i) => ofPara(pages, i));
  check("table: cells side by side", near(A[0].y, B[0].y) && near(B[0].x - A[0].x, 2000, 1) && near(C[0].x - B[0].x, 2000, 1));
  check("table: tall cell wraps in its width", C.length >= 3, String(C.length));
  const cells = pages[0].items.filter((it) => it.type === "cell");
  check("table: cell boxes (3 per row)", cells.length === 6, String(cells.length));
  const row1 = cells.filter((it) => near(it.y, cells[0].y));
  check("table: row height is the tallest cell", row1.every((it) => near(it.h, row1[0].h)) && row1[0].h >= C.reduce((n, it) => n + it.h, 0));
  const after = ofPara(pages, 6)[0];
  check("table: paragraph after the table below both rows", after.y >= Math.max(...cells.map((it) => it.y + it.h)) - 1);
  const early = layoutDocument(m, Date.parse(T1) - 1).pages[0].items.filter((it) => it.type === "cell");
  check("table: inserted row absent before its insertion", early.length === 3, String(early.length));
}

// 図 (幅・浮動配置)
{
  const emu = (tw) => Math.round(tw * 635);
  const inline = `<w:r><w:drawing><wp:inline><wp:extent cx="${emu(2880)}" cy="${emu(1440)}"/></wp:inline></w:drawing></w:r>`;
  const anchor =
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" behindDoc="0"><wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${emu(3000)}" cy="${emu(2000)}"/><wp:wrapSquare wrapText="bothSides"/></wp:anchor></w:drawing></w:r>`;
  const body = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${inline}</w:p>` + `<w:p>${anchor}${run(jp(600))}</w:p>` + sect(`<w:docGrid w:linePitch="360"/>`);
  const { pages } = layoutDocument(await model(body), Infinity);
  const fig = ofPara(pages, 0).find((it) => it.type === "figure");
  check("figure: inline width and centering", near(fig.w, 2880) && near(fig.x, 1701 + (8504 - 2880) / 2, 1));
  const float = ofPara(pages, 1).find((it) => it.type === "figure");
  check("figure: floating figure at the right margin", float?.float && near(float.x + float.w, 1701 + 8504, 1));
  const beside = ofPara(pages, 1).filter((it) => it.type === "line" && it.y < float.y + float.h);
  const below = ofPara(pages, 1).filter((it) => it.type === "line" && it.y >= float.y + float.h);
  check("figure: lines beside the figure are shorter", beside.length > 0 && beside.every((it) => it.x + it.w <= float.x + 1));
  check("figure: lines below the figure use the full width", below.length > 0 && below.every((it) => near(it.w, 8504, 1)));
}

// Word が保存した改ページ位置
{
  // 30 行目の段落で Word は改ページしていた (近似の割り付けでは 36 行目まで1ページ目)
  const body =
    Array.from({ length: 30 }, () => p("x")).join("") +
    `<w:p><w:r><w:lastRenderedPageBreak/><w:t>hint</w:t></w:r></w:p>` +
    Array.from({ length: 20 }, () => p("y")).join("") +
    sect(`<w:docGrid w:linePitch="360"/>`);
  const m = await model(body);
  const fin = layoutDocument(m, Infinity, { final: true });
  check("lastRenderedPageBreak: page starts at the hint", ofPara(fin.pages, 30)[0].page === 1 && fin.hintBreaks === 1);
  const approx = layoutDocument(m, Infinity);
  check("lastRenderedPageBreak: not used for past versions", ofPara(approx.pages, 30)[0].page === 0);
  // 段の前半にある手がかりは使わない (二重の改ページを避ける)
  const early = await model(p("x") + `<w:p><w:r><w:lastRenderedPageBreak/><w:t>hint</w:t></w:r></w:p>` + sect());
  check("lastRenderedPageBreak: ignored near the top", layoutDocument(early, Infinity, { final: true }).pages.length === 1);
}

// 縦書き
{
  const body = Array.from({ length: 5 }, () => p(jp(60))).join("") + sect(`<w:textDirection w:val="tbRl"/><w:docGrid w:linePitch="360"/>`);
  const { pages } = layoutDocument(await model(body), Infinity);
  const ls = lines(pages[0]);
  check("vertical: page marked vertical", pages[0].vertical === true);
  check("vertical: lines are vertical", ls.every((it) => it.h > it.w));
  check("vertical: first line at the right", ls[0].x > ls[ls.length - 1].x && near(ls[0].x + ls[0].w, 11906 - 1701, 2), `${ls[0].x}`);
}

// 総ページ数の補正と、flow の SVG (段組みの文書)
{
  const body = Array.from({ length: 50 }, () => p(jp(80))).join("") + sect(`<w:cols w:num="2" w:space="425"/><w:docGrid w:linePitch="360"/>`);
  const m = await model(body, { app: 2 });
  const flow = buildFlow(m, { gapThresholdHours: 1, bulkChars: 150 });
  check("calibration: line scale within range", flow.lineScale >= 0.8 && flow.lineScale <= 1.25, String(flow.lineScale));
  const insBody = body.replace(`${run(jp(80))}</w:p>`, `${run(jp(80))}<w:ins ${rev(T1)}>${run("added")}</w:ins></w:p>`);
  const flow2 = buildFlow(await model(insBody), { gapThresholdHours: 1, bulkChars: 150 });
  const svg = renderFlowSvg(flow2, { title: "t" });
  check("flow svg: renders a two-column document", svg.startsWith("<svg") && !svg.includes("NaN"));
}

if (failed) {
  console.error(`${failed} check(s) failed`);
  process.exit(1);
}
console.log("page layout: all checks passed");
fs.rmSync(tmp, { recursive: true, force: true });
