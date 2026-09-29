/**
 * テスト用の .docx フィクスチャを生成するスクリプト。
 * 実際のWordファイルが手元になくても、両CLIの動作確認ができるようにする。
 *
 * 生成するファイル:
 *  - fixtures/natural-writing.docx   : 人が時間をかけて少しずつタイプしたことを想定
 *  - fixtures/suspicious-paste.docx  : 最初は少し自分でタイプした後、大きな塊を
 *                                       一瞬で貼り付けたことを想定 (AI生成文の貼付を模擬)
 *  - fixtures/flow-demo.docx      : 見出し・図を含む複数段落の文書を3つの時間区間で編集
 *                                       (docx-revision-flow 用)
 */
import JSZip from "jszip";
import * as fs from "fs";
import * as path from "path";

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

type Segment =
  | { kind: "text"; text: string }
  | { kind: "ins"; text: string; author: string; date: Date; id: number }
  | { kind: "del"; text: string; author: string; date: Date; id: number };

function buildDocumentXml(segments: Segment[]): string {
  const runs = segments
    .map((seg) => {
      if (seg.kind === "text") {
        return `<w:r><w:t xml:space="preserve">${escapeXml(seg.text)}</w:t></w:r>`;
      } else if (seg.kind === "ins") {
        return `<w:ins w:id="${seg.id}" w:author="${escapeXml(seg.author)}" w:date="${seg.date.toISOString()}"><w:r><w:t xml:space="preserve">${escapeXml(
          seg.text
        )}</w:t></w:r></w:ins>`;
      } else {
        return `<w:del w:id="${seg.id}" w:author="${escapeXml(seg.author)}" w:date="${seg.date.toISOString()}"><w:r><w:delText xml:space="preserve">${escapeXml(
          seg.text
        )}</w:delText></w:r></w:del>`;
      }
    })
    .join("\n      ");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      ${runs}
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults/>
</w:styles>`;

async function writeDocx(outPath: string, segments: Segment[]): Promise<void> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/document.xml", buildDocumentXml(segments));
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/styles.xml", STYLES);
  const buf = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outPath, buf);
  console.log(`生成: ${outPath} (${segments.length} セグメント)`);
}

// 再現性のための簡易シード付き乱数 (mulberry32)
function mulberry32(seed: number) {
  return function () {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const WORDS = [
  "本研究では",
  "従来手法の",
  "課題を整理し",
  "新しい評価指標を",
  "提案する。",
  "実験の結果",
  "提案手法は",
  "既存手法と比較して",
  "高い精度を示した。",
  "今後の課題として",
  "データ数の拡大や",
  "他分野への応用が",
  "考えられる。",
  "また、",
  "評価者による",
  "主観評価との",
  "整合性についても",
  "検討が必要である。",
  "特に、",
  "面接評価の場面では",
  "口頭での説明力を",
  "定量的に測ることが",
  "重要だと考えられる。",
];

async function makeNatural() {
  const rnd = mulberry32(42);
  const author = "学生A";
  let t = new Date("2026-05-10T09:00:00Z").getTime();
  const segments: Segment[] = [];
  let id = 1;

  for (let i = 0; i < 40; i++) {
    // 20〜100秒ランダムな間隔 (人が考えながら打っている想定)
    t += 20_000 + Math.floor(rnd() * 80_000);
    const word = WORDS[i % WORDS.length];
    segments.push({ kind: "ins", text: word, author, date: new Date(t), id: id++ });

    // たまに typo 修正の削除を挟む
    if (i > 0 && i % 9 === 0) {
      t += 3_000 + Math.floor(rnd() * 5_000);
      segments.push({
        kind: "del",
        text: "、えー",
        author,
        date: new Date(t),
        id: id++,
      });
    }
  }

  await writeDocx(
    path.join(__dirname, "..", "fixtures", "natural-writing.docx"),
    segments
  );
}

async function makeSuspicious() {
  const rnd = mulberry32(7);
  const author = "学生B";
  let t = new Date("2026-05-11T14:00:00Z").getTime();
  const segments: Segment[] = [];
  let id = 1;

  // 最初は少しだけ自分でタイプしたように見せる (6イベント、約3分)
  for (let i = 0; i < 6; i++) {
    t += 15_000 + Math.floor(rnd() * 25_000);
    segments.push({
      kind: "ins",
      text: WORDS[i % WORDS.length],
      author,
      date: new Date(t),
      id: id++,
    });
  }

  // ここで外部AIアプリで作文した文章を一瞬で貼り付けたことを模擬
  // (直前のイベントから1秒後に、大量の文字が一括挿入される)
  t += 1_000;
  const pastedText =
    "本研究の目的は、AIを活用した面接評価システムにおいて、受験者の口頭説明の" +
    "論理性と一貫性を定量的に評価する新しい指標を提案することである。" +
    "既存研究では主に音声認識精度やキーワード一致度に基づく評価が中心であったが、" +
    "本研究ではそれに加えて、発話の時系列構造を考慮した意味的つながりの評価を導入する。" +
    "具体的には、発話をセグメントに分割し、各セグメント間の意味的類似度を" +
    "埋め込みベクトルにより算出したうえで、論理展開の自然さをスコア化する。" +
    "実験では模擬面接データを用い、提案手法が人間評価者のスコアと高い相関を" +
    "示すことを確認した。今後は評価対象分野の拡大や、リアルタイム評価への" +
    "応用が期待される。";
  segments.push({
    kind: "ins",
    text: pastedText,
    author,
    date: new Date(t),
    id: id++,
  });

  // 貼り付け後、少しだけ自分で手直ししたように見せる (5イベント)
  for (let i = 0; i < 5; i++) {
    t += 20_000 + Math.floor(rnd() * 40_000);
    segments.push({
      kind: "ins",
      text: WORDS[(i + 10) % WORDS.length],
      author,
      date: new Date(t),
      id: id++,
    });
  }

  await writeDocx(
    path.join(__dirname, "..", "fixtures", "suspicious-paste.docx"),
    segments
  );
}

async function makeMultiSession() {
  const rnd = mulberry32(99);
  const author = "学生C";
  const segments: Segment[] = [];
  let id = 1;

  // 3日間にわたり、間に大きな空白期間 (無編集期間) を挟みながら執筆した想定。
  // 1日目 09:00〜, 2日目 (約29時間後), 3日目 (約20.5時間後)
  let t = new Date("2026-06-01T09:00:00Z").getTime();
  const sessionGapsHours = [0, 29, 20.5]; // 各セッション開始前の空白 (先頭は無視)
  const eventsPerSession = [10, 14, 8];

  for (let s = 0; s < sessionGapsHours.length; s++) {
    if (s > 0) {
      t += sessionGapsHours[s] * 3600_000;
    }
    for (let i = 0; i < eventsPerSession[s]; i++) {
      t += 15_000 + Math.floor(rnd() * 60_000);
      segments.push({
        kind: "ins",
        text: WORDS[(i + s * 5) % WORDS.length],
        author,
        date: new Date(t),
        id: id++,
      });
    }
  }

  await writeDocx(
    path.join(__dirname, "..", "fixtures", "multi-session.docx"),
    segments
  );
}

// ---------------------------------------------------------------------------
// docx-revision-flow 用: 見出し・図・複数段落を含み、3つの時間区間で
// 「手での入力」「複数段落の一括貼り付け」「段落の置き換え」を行った文書
// ---------------------------------------------------------------------------

type Rev = { author: string; date: Date; id: number };

interface ParaSpec {
  style?: string;
  /** 段落記号の挿入 (新しく追加された段落) / 削除 (段落ごと削除された) */
  markIns?: Rev;
  markDel?: Rev;
  /** 段落記号の移動 (Word が移動として記録した段落) */
  markMoveFrom?: Rev;
  markMoveTo?: Rev;
  segments: (
    | Segment
    | { kind: "figure"; heightEmu: number; ins?: Rev }
    | ({ kind: "moveFrom" | "moveTo"; text: string; name: string; rangeId: number } & Rev)
  )[];
}

function revAttrs(r: { author: string; date: Date; id: number }): string {
  return `w:id="${r.id}" w:author="${escapeXml(r.author)}" w:date="${r.date.toISOString()}"`;
}

function buildParagraphsXml(paras: ParaSpec[]): string {
  const body = paras
    .map((p) => {
      const pPr: string[] = [];
      if (p.style) pPr.push(`<w:pStyle w:val="${p.style}"/>`);
      const marks = [
        p.markIns && `<w:ins ${revAttrs(p.markIns)}/>`,
        p.markDel && `<w:del ${revAttrs(p.markDel)}/>`,
        p.markMoveFrom && `<w:moveFrom ${revAttrs(p.markMoveFrom)}/>`,
        p.markMoveTo && `<w:moveTo ${revAttrs(p.markMoveTo)}/>`,
      ].filter(Boolean);
      if (marks.length > 0) pPr.push(`<w:rPr>${marks.join("")}</w:rPr>`);
      const runs = p.segments
        .map((seg) => {
          if (seg.kind === "moveFrom" || seg.kind === "moveTo") {
            // Word は移動元・移動先を同じ範囲名 (w:name) の範囲で囲んで記録する
            const range = seg.kind === "moveFrom" ? "moveFromRange" : "moveToRange";
            return (
              `<w:${range}Start w:id="${seg.rangeId}" ${revAttrs(seg).replace(/w:id="\d+" /, "")} w:name="${seg.name}"/>` +
              `<w:${seg.kind} ${revAttrs(seg)}><w:r><w:t xml:space="preserve">${escapeXml(seg.text)}</w:t></w:r></w:${seg.kind}>` +
              `<w:${range}End w:id="${seg.rangeId}"/>`
            );
          }
          if (seg.kind === "figure") {
            const drawing =
              `<w:r><w:drawing><wp:inline><wp:extent cx="3600000" cy="${seg.heightEmu}"/>` +
              `<wp:docPr id="1" name="図"/></wp:inline></w:drawing></w:r>`;
            return seg.ins ? `<w:ins ${revAttrs(seg.ins)}>${drawing}</w:ins>` : drawing;
          }
          if (seg.kind === "text") {
            return `<w:r><w:t xml:space="preserve">${escapeXml(seg.text)}</w:t></w:r>`;
          }
          if (seg.kind === "ins") {
            return `<w:ins ${revAttrs(seg)}><w:r><w:t xml:space="preserve">${escapeXml(seg.text)}</w:t></w:r></w:ins>`;
          }
          return `<w:del ${revAttrs(seg)}><w:r><w:delText xml:space="preserve">${escapeXml(seg.text)}</w:delText></w:r></w:del>`;
        })
        .join("");
      return `<w:p>${pPr.length ? `<w:pPr>${pPr.join("")}</w:pPr>` : ""}${runs}</w:p>`;
    })
    .join("\n    ");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    ${body}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>
  </w:body>
</w:document>`;
}

const FLOW_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
</w:styles>`;

async function makeFlowDemo() {
  const rnd = mulberry32(2024);
  const author = "学生D";
  let id = 1;
  const rev = (date: number) => ({ author, date: new Date(date), id: id++ });
  // 段落ごとに単語の並びを変え、別々の段落の文章が偶然一致しないようにする
  // (一致すると、文書内の並べ替えと判定されてしまうため)
  const filler = (n: number, seed: number) => {
    const r = mulberry32(1000 + seed);
    return Array.from({ length: n }, () => WORDS[Math.floor(r() * WORDS.length)]).join("");
  };

  const paras: ParaSpec[] = [];
  // 変更履歴の記録を始める前からあった部分 (様式の説明文など)
  paras.push({ style: "1", segments: [{ kind: "text", text: "研究計画調書" }] });
  for (let i = 0; i < 6; i++) {
    paras.push({ segments: [{ kind: "text", text: filler(14, i * 3) }] });
  }

  // 区間1: 見出しと2段落を少しずつ手で入力 (約70分)
  let t = new Date("2026-06-01T00:00:00Z").getTime();
  const typeInto = (p: ParaSpec, count: number, offset: number) => {
    for (let i = 0; i < count; i++) {
      t += 20_000 + Math.floor(rnd() * 70_000);
      p.segments.push({ kind: "ins", text: WORDS[(i + offset) % WORDS.length], ...rev(t) });
      if (i % 7 === 6) {
        t += 5_000;
        // Word は自分で入力した文字を自分で消しても記録を残さないため、
        // 削除として記録されるのは記録開始前からあった文章 (様式の説明文) の削除
        paras[1 + (i % 6)].segments.push({ kind: "del", text: "えーと", ...rev(t) });
      }
    }
  };
  const h1: ParaSpec = { style: "1", markIns: rev(t), segments: [] };
  paras.push(h1);
  typeInto(h1, 2, 0);
  const p1: ParaSpec = { markIns: rev(t), segments: [] };
  paras.push(p1);
  typeInto(p1, 22, 3);
  const p2: ParaSpec = { markIns: rev(t), segments: [] };
  paras.push(p2);
  typeInto(p2, 18, 8);

  // 区間2 (約26時間後): 見出しを入力し、3段落をまとめて貼り付けてから小さく手直し
  t += 26 * 3600_000;
  const h2: ParaSpec = { style: "1", markIns: rev(t), segments: [] };
  paras.push(h2);
  typeInto(h2, 2, 5);
  t += 30_000;
  const pasteAt = t;
  const pasted: ParaSpec[] = [0, 1, 2].map((k) => ({
    markIns: rev(pasteAt),
    segments: [{ kind: "ins" as const, text: filler(9, k * 5 + 2), ...rev(pasteAt) }],
  }));
  paras.push(...pasted);
  for (let i = 0; i < 6; i++) {
    t += 30_000 + Math.floor(rnd() * 60_000);
    pasted[i % 3].segments.push({ kind: "ins", text: WORDS[(i + 4) % WORDS.length], ...rev(t) });
  }

  // 区間3 (約5時間後): 記録前からあった段落を丸ごと置き換え・1段落を削除し、区間1の段落を細かく修正し、図を追加
  t += 5 * 3600_000;
  const replaced = paras[3];
  const oldText = (replaced.segments[0] as Segment & { kind: "text" }).text;
  replaced.segments = [
    { kind: "del", text: oldText, ...rev(t) },
    { kind: "ins", text: filler(24, 11), ...rev(t) },
  ];
  t += 40_000;
  const removed = paras[6];
  const removedText = (removed.segments[0] as Segment & { kind: "text" }).text;
  removed.segments = [{ kind: "del", text: removedText, ...rev(t) }];
  removed.markDel = rev(t);
  for (let i = 0; i < 9; i++) {
    t += 25_000 + Math.floor(rnd() * 50_000);
    const text = WORDS[(i + 15) % WORDS.length].slice(0, 4);
    if (i % 3 === 2) paras[5].segments.push({ kind: "del", text, ...rev(t) });
    else p1.segments.push({ kind: "ins", text, ...rev(t) });
  }
  t += 60_000;
  paras.push({ markIns: rev(t), segments: [{ kind: "figure", heightEmu: 1_800_000, ins: rev(t) }] });

  // 区間4 (約20時間後): 段落の並べ替え
  //  - 冒頭の段落をカット＋貼り付けで後ろへ移動 (Word が移動として記録)
  //  - 2つ目の段落をコピー＋貼り付けで後ろへ複製してから元を削除 (挿入＋削除として記録)
  t += 20 * 3600_000;
  const plainText = (p: ParaSpec) =>
    p.segments.filter((sg): sg is Segment & { kind: "text" } => sg.kind === "text").map((sg) => sg.text).join("");
  const moved = paras[1];
  const movedText = plainText(moved);
  const moveFromRev = rev(t);
  const moveToRev = rev(t);
  moved.segments = [
    ...moved.segments.filter((sg) => sg.kind !== "text"),
    { kind: "moveFrom", text: movedText, name: "move1", rangeId: 901, ...moveFromRev },
  ];
  moved.markMoveFrom = moveFromRev;
  const movedDest: ParaSpec = {
    markMoveTo: moveToRev,
    segments: [{ kind: "moveTo", text: movedText, name: "move1", rangeId: 902, ...moveToRev }],
  };
  paras.splice(paras.indexOf(pasted[2]) + 1, 0, movedDest);

  t += 60_000;
  const copied = paras[2];
  const copiedText = plainText(copied);
  const copyDest: ParaSpec = { markIns: rev(t), segments: [{ kind: "ins", text: copiedText, ...rev(t) }] };
  paras.splice(paras.indexOf(h2), 0, copyDest);
  t += 30_000;
  copied.segments = [...copied.segments.filter((sg) => sg.kind !== "text"), { kind: "del", text: copiedText, ...rev(t) }];
  copied.markDel = rev(t);

  const outPath = path.join(__dirname, "..", "fixtures", "flow-demo.docx");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/document.xml", buildParagraphsXml(paras));
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/styles.xml", FLOW_STYLES);
  fs.writeFileSync(outPath, await zip.generateAsync({ type: "nodebuffer" }));
  console.log(`生成: ${outPath} (${paras.length} 段落)`);
}

async function main() {
  fs.mkdirSync(path.join(__dirname, "..", "fixtures"), { recursive: true });
  await makeNatural();
  await makeSuspicious();
  await makeMultiSession();
  await makeFlowDemo();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
