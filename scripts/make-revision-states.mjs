#!/usr/bin/env node
// Draws the state diagrams in the README's "The .docx files these tools expect" section:
//   fixtures/revision-states{,.ja}.svg           a .docx saved on your computer
//   fixtures/revision-states-onedrive{,.ja}.svg  a .docx saved on OneDrive (adds restoring from Version History)
// With Inkscape on the PATH it also renders each one to .png at 2x; otherwise the PNGs are left as they are.
//
//   npm run diagrams

import { execFileSync } from "node:child_process";
import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");

const L = {
  en: {
    font: "'Helvetica Neue',Arial,sans-serif",
    title: "Four states of a .docx's tracked-change history",
    titleOneDrive: "Four states of a .docx's tracked-change history — file saved on OneDrive",
    sub: "These tools analyze documents in state D. Dashed red = tracked changes reduced or erased, which can't be undone",
    subOneDrive: "These tools analyze documents in state D. Dashed red = tracked changes reduced or erased",
    start: "Start",
    forward: ["Turn on", "Edit", "More edits"],
    off: "Turn off",
    some: "Accept / reject some changes",
    toB: "Accept All / Reject All",
    toA: "Accept All Changes and Stop Tracking",
    restore: "Restore a version from when it was in D, if Version History has one (not guaranteed)",
    boxes: [
      ["A: No tracking", "Track Changes OFF"],
      ["B: Tracking on", "0 changes yet"],
      ["C: Too sparse", "Too few / too short"],
      ["D: Sufficient", "Analysis target"],
    ],
  },
  ja: {
    font: "'Noto Sans CJK JP','Noto Sans JP','Yu Gothic','Hiragino Sans',Meiryo,'MS PGothic','Helvetica Neue',Arial,sans-serif",
    title: "変更履歴から見た .docx ファイルの4つの状態",
    titleOneDrive: "変更履歴から見た .docx ファイルの4つの状態 — OneDrive に保存したファイルの場合",
    sub: "本ツールが解析の対象とするのは「D: 履歴十分」の文書。赤の破線は変更履歴が減る・消える経路で、元には戻せない",
    subOneDrive: "本ツールが解析の対象とするのは「D: 履歴十分」の文書。赤の破線は変更履歴が減る・消える経路",
    start: "開始",
    forward: ["記録ON", "編集", "編集が蓄積"],
    off: "記録OFF",
    some: "一部の変更を承諾・元に戻す",
    toB: "すべて反映・すべて元に戻す",
    toA: "「すべての変更を反映し、変更の記録を停止」",
    restore: "バージョン履歴にDだった時点の版があれば、それを復元 (保証なし)",
    boxes: [
      ["A: 未記録", "Track Changes OFF"],
      ["B: 記録ON", "履歴まだ0件"],
      ["C: 履歴不十分", "件数・期間が少ない"],
      ["D: 履歴十分", "本ツールの解析対象"],
    ],
  },
};

const W = 920;
const BX = [80, 300, 520, 740];
const BW = 150;
const GRAY = "#9CA3AF", DARK = "#111827", MUTED = "#6B7280", RED = "#DC2626", GREEN = "#16A34A";

function text(x, y, s, { size = 11.5, fill = MUTED, anchor = "middle", weight } = {}) {
  const w = weight ? ` font-weight="${weight}"` : "";
  return `<text x="${x}" y="${y}" text-anchor="${anchor}" font-size="${size}"${w} fill="${fill}">${s}</text>`;
}

function svg(lang, oneDrive) {
  const d = L[lang];
  const o = [];
  const TOP = oneDrive ? 170 : 134;
  const BOT = TOP + 56, MID = TOP + 28;
  const H = BOT + 92;
  o.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" font-family="${d.font}" font-size="13">`);
  o.push("<defs>");
  for (const [id, c] of [["fwd", GRAY], ["start", DARK], ["danger", RED], ["restore", GREEN]]) {
    o.push(`<marker id="${id}-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${c}"/></marker>`);
  }
  o.push("</defs>");
  o.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#FFFFFF"/>`);
  o.push(text(24, 28, oneDrive ? d.titleOneDrive : d.title, { size: 15, fill: DARK, anchor: "start", weight: 600 }));
  o.push(text(24, 46, oneDrive ? d.subOneDrive : d.sub, { anchor: "start" }));

  // Initial state
  o.push(`<circle cx="30" cy="${MID}" r="7" fill="${DARK}"/>`);
  o.push(`<path d="M38 ${MID}H${BX[0]}" fill="none" stroke="${DARK}" stroke-width="1.25" marker-end="url(#start-arrow)"/>`);
  o.push(text(30, MID + 24, d.start));

  // Forward transitions
  d.forward.forEach((label, i) => {
    const x1 = BX[i] + BW, x2 = BX[i + 1];
    const y = i === 0 ? MID - 8 : MID;
    o.push(`<path d="M${x1} ${y}H${x2}" fill="none" stroke="${GRAY}" stroke-width="1.25" marker-end="url(#fwd-arrow)"/>`);
    o.push(text((x1 + x2) / 2, y - 7, label));
  });

  // B -> A: turning Track Changes off loses nothing
  {
    const x1 = BX[0] + BW, x2 = BX[1];
    o.push(`<path d="M${x2} ${MID + 8}H${x1}" fill="none" stroke="${GRAY}" stroke-width="1.25" stroke-dasharray="4 4" marker-end="url(#fwd-arrow)"/>`);
    o.push(text((x1 + x2) / 2, MID + 23, d.off));
  }

  const a = BX[0] + BW / 2, b = BX[1] + BW / 2, c = BX[2] + BW / 2, dd = BX[3] + BW / 2;

  // D -> C: accept / reject some changes
  o.push(`<path d="M${dd} ${TOP}C${dd} ${TOP - 46} ${c} ${TOP - 46} ${c} ${TOP}" fill="none" stroke="${RED}" stroke-width="1.25" stroke-dasharray="4 4" marker-end="url(#danger-arrow)"/>`);
  o.push(text((c + dd) / 2, TOP - 44, d.some, { fill: DARK }));

  // C, D -> B (Accept All / Reject All) and C, D -> A (Accept All Changes and Stop Tracking)
  const y1 = BOT + 34, y2 = BOT + 74, j = (c + dd) / 2;
  o.push(`<g fill="none" stroke="${RED}" stroke-width="1.25" stroke-dasharray="4 4">`);
  o.push(`<path d="M${c} ${BOT}V${y1}"/><path d="M${dd} ${BOT}V${y1}"/><path d="M${dd} ${y1}H${c}"/>`);
  o.push(`<path d="M${c} ${y1}H${b}V${BOT}" marker-end="url(#danger-arrow)"/>`);
  o.push(`<path d="M${j} ${y1}V${y2}H${a}V${BOT}" marker-end="url(#danger-arrow)"/>`);
  o.push("</g>");
  o.push(`<circle cx="${c}" cy="${y1}" r="2.5" fill="${RED}"/><circle cx="${j}" cy="${y1}" r="2.5" fill="${RED}"/>`);
  o.push(text((b + c) / 2, y1 - 6, d.toB, { fill: DARK }));
  o.push(text((a + j) / 2, y2 - 6, d.toA, { fill: DARK }));

  if (oneDrive) {
    // A, B, C -> D: restore a version saved while the document was in D
    const yg = 92, cs = BX[2] + 30, dt = BX[3] + BW - 30;
    o.push(`<g fill="none" stroke="${GREEN}" stroke-width="1.5" stroke-dasharray="6 3">`);
    o.push(`<path d="M${a} ${TOP}V${yg}"/><path d="M${b} ${TOP}V${yg}"/><path d="M${cs} ${TOP}V${yg}"/>`);
    o.push(`<path d="M${a} ${yg}H${dt}V${TOP}" marker-end="url(#restore-arrow)"/>`);
    o.push("</g>");
    for (const x of [b, cs]) o.push(`<circle cx="${x}" cy="${yg}" r="2.5" fill="${GREEN}"/>`);
    o.push(text((a + dt) / 2, yg - 8, d.restore, { size: 12, fill: GREEN, weight: 600 }));
  }

  // States; D, the one these tools analyze, is highlighted
  d.boxes.forEach(([heading, detail], i) => {
    const x = BX[i], cx = x + BW / 2;
    if (i === 3) {
      o.push(`<rect x="${x}" y="${TOP}" width="${BW}" height="56" rx="8" fill="#DBEAFE" stroke="#2563EB" stroke-width="2.5"/>`);
      o.push(text(cx, TOP + 24, heading, { size: 14, fill: "#1D4ED8", weight: 700 }));
      o.push(text(cx, TOP + 42, detail, { fill: "#1D4ED8", weight: 600 }));
    } else {
      o.push(`<rect x="${x}" y="${TOP}" width="${BW}" height="56" rx="8" fill="#FFFFFF" stroke="${GRAY}" stroke-width="1.25"/>`);
      o.push(text(cx, TOP + 24, heading, { size: 14, fill: DARK, weight: 600 }));
      o.push(text(cx, TOP + 42, detail));
    }
  });
  o.push("</svg>");
  return o.join("\n") + "\n";
}

function hasInkscape() {
  try {
    execFileSync("inkscape", ["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

const png = hasInkscape();
for (const oneDrive of [false, true]) {
  for (const lang of ["en", "ja"]) {
    const name = "revision-states" + (oneDrive ? "-onedrive" : "") + (lang === "ja" ? ".ja" : "");
    const svgPath = join(FIXTURES, `${name}.svg`);
    writeFileSync(svgPath, svg(lang, oneDrive));
    console.log(`wrote ${svgPath}`);
    if (png) {
      const pngPath = join(FIXTURES, `${name}.png`);
      execFileSync("inkscape", [svgPath, "--export-type=png", "--export-dpi=192", `--export-filename=${pngPath}`], { stdio: "ignore" });
      console.log(`wrote ${pngPath}`);
    }
  }
}
if (!png) console.log("Inkscape not found; PNGs were not updated.");
