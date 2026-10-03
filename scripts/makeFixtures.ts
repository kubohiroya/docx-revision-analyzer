/**
 * テスト用の .docx フィクスチャを生成するスクリプト。
 * 実際のWordファイルが手元になくても、両CLIの動作確認ができるようにする。
 *
 * 生成するファイル:
 *  - fixtures/natural-writing.docx   : 人が時間をかけて少しずつタイプしたことを想定
 *  - fixtures/suspicious-paste.docx  : 最初は少し自分でタイプした後、大きな塊を
 *                                       一瞬で貼り付けたことを想定 (AI生成文の貼付を模擬)
 *  - fixtures/chart-demo.docx       : 2日にわたる3回の執筆。1回まとめて貼り付け、下書きを散発的に削除
 *                                       (README と docs/usage の docx-revision-chart の図の例)
 *  - fixtures/flow-demo.docx      : 見出し・図を含む複数段落の文書を3つの時間区間で編集
 *                                       (docx-revision-flow 用)
 *  - fixtures/snapshot-demo/sprint{1,2,3}/thesis.docx
 *                                     : 卒論を3回の区切り (スプリント) で提出したもの。各回の前に教員がすべて承諾して返し、
 *                                       第3回には記録をオフにして入力した段落がある (docx-revision-snapshot 用)
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

/** rsid: その文字列を入力した編集セッションの識別子 (w:rsidR)。省略すると付けない */
type Segment =
  | { kind: "text"; text: string; rsid?: string }
  | { kind: "ins"; text: string; author: string; date: Date; id: number; rsid?: string }
  | { kind: "del"; text: string; author: string; date: Date; id: number; rsid?: string };

const rsidAttr = (seg: Segment) => (seg.rsid ? ` w:rsidR="${seg.rsid}"` : "");

function buildDocumentXml(segments: Segment[]): string {
  const runs = segments
    .map((seg) => {
      if (seg.kind === "text") {
        return `<w:r${rsidAttr(seg)}><w:t xml:space="preserve">${escapeXml(seg.text)}</w:t></w:r>`;
      } else if (seg.kind === "ins") {
        return `<w:ins w:id="${seg.id}" w:author="${escapeXml(seg.author)}" w:date="${seg.date.toISOString()}"><w:r${rsidAttr(seg)}><w:t xml:space="preserve">${escapeXml(
          seg.text
        )}</w:t></w:r></w:ins>`;
      } else {
        return `<w:del w:id="${seg.id}" w:author="${escapeXml(seg.author)}" w:date="${seg.date.toISOString()}"><w:r${rsidAttr(seg)}><w:delText xml:space="preserve">${escapeXml(
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

const SETTINGS_TYPE =
  '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
const SETTINGS_REL =
  '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';

/** 変更履歴の記録をオンにし、編集セッションの一覧 (rsid) を持つ settings.xml */
function settingsWithRsids(rsidRoot: string, rsids: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions/>
  <w:defaultTabStop w:val="840"/>
  <w:rsids><w:rsidRoot w:val="${rsidRoot}"/>${[...new Set([rsidRoot, ...rsids])].map((r) => `<w:rsid w:val="${r}"/>`).join("")}</w:rsids>
</w:settings>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults/>
</w:styles>`;

/** settingsXml を渡すと word/settings.xml も入れる (変更履歴の記録・編集セッションの一覧など) */
async function writeDocx(outPath: string, segments: Segment[], settingsXml?: string): Promise<void> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", settingsXml ? CONTENT_TYPES.replace("</Types>", `${SETTINGS_TYPE}\n</Types>`) : CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/document.xml", buildDocumentXml(segments));
  zip.file(
    "word/_rels/document.xml.rels",
    settingsXml ? DOCUMENT_RELS.replace("</Relationships>", `${SETTINGS_REL}\n</Relationships>`) : DOCUMENT_RELS
  );
  if (settingsXml) zip.file("word/settings.xml", settingsXml);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
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

// ---------------------------------------------------------------------------
// docx-revision-chart の README 用: 2日にわたる3回の執筆 (間に無編集期間)。細かく入力する合間に、
// 下書き (記録開始前からあった文章) を大小さまざまな単位で散発的に削除し、2回目の最初に文章をまとめて貼り付ける
// ---------------------------------------------------------------------------

async function makeChartDemo() {
  const rnd = mulberry32(314);
  const author = "学生E";
  let t = new Date("2026-06-10T00:00:00Z").getTime();
  let id = 1;
  const draft = "（下書き）" + WORDS.join("").repeat(6);
  const segments: Segment[] = [{ kind: "text", text: draft }];
  let draftPos = 5;
  const deleteDraft = (len: number) => {
    const text = draft.slice(draftPos, draftPos + len);
    draftPos += len;
    segments.push({ kind: "del", text, author, date: new Date(t), id: id++ });
  };
  const deletionSizes = [3, 6, 10, 18, 35, 60, 110];

  // 1日目 09:00〜10:20 手で入力しながら、ときどき下書きを削る
  const typeFor = (minutes: number) => {
    const until = t + minutes * 60_000;
    while (t < until) {
      t += 15_000 + Math.floor(rnd() * 60_000);
      segments.push({ kind: "ins", text: WORDS[Math.floor(rnd() * WORDS.length)], author, date: new Date(t), id: id++ });
      if (rnd() < 0.09) {
        t += 5_000;
        deleteDraft(deletionSizes[Math.floor(rnd() * deletionSizes.length)]);
      }
    }
  };
  typeFor(80);

  // 約4.5時間の無編集の後、1日目 14:50 外部で作った文章をまとめて貼り付け (同じ時刻に2段落分)
  t += 4.5 * 3600_000;
  const pasted =
    "本研究の目的は、AIを活用した面接評価システムにおいて、受験者の口頭説明の論理性と一貫性を定量的に" +
    "評価する新しい指標を提案することである。既存研究では主に音声認識精度やキーワード一致度に基づく評価が" +
    "中心であったが、本研究ではそれに加えて、発話の時系列構造を考慮した意味的つながりの評価を導入する。" +
    "具体的には、発話をセグメントに分割し、各セグメント間の意味的類似度を埋め込みベクトルにより算出したうえで、" +
    "論理展開の自然さをスコア化する。";
  const half = Math.floor(pasted.length / 2);
  segments.push({ kind: "ins", text: pasted.slice(0, half), author, date: new Date(t), id: id++ });
  segments.push({ kind: "ins", text: pasted.slice(half), author, date: new Date(t), id: id++ });
  t += 20_000;
  deleteDraft(90);

  // 〜16:20 貼り付けた文章の手直しと続きの入力
  typeFor(90);

  // 夜をはさんで約17時間の無編集の後、2日目 09:30〜10:30 に見直し
  t += 17 * 3600_000;
  typeFor(60);

  await writeDocx(path.join(__dirname, "..", "fixtures", "chart-demo.docx"), segments);
}

/**
 * 卒論を3回の区切り (スプリント) で提出したもの。各回の提出の後、教員がすべての変更を承諾して返す
 * (承諾した文章は変更履歴の外の本文になり、次の回にそれを消すと削除として記録される)。
 * 第3回には、別のアプリで書いた段落を記録をオフにして入力したもの (前回までに無い編集セッションの本文) がある。
 */
async function makeSnapshotDemo() {
  const rnd = mulberry32(2026);
  const author = "学生F";
  const root = "00A00000";
  let id = 1;
  let sessionNo = 0;
  const rsids: string[] = [];
  const newSession = () => {
    const r = `00B${String(++sessionNo).padStart(5, "0")}`;
    rsids.push(r);
    return r;
  };
  /** 承諾した状態: 削除を除き、挿入を本文にする */
  const accept = (segs: Segment[]): Segment[] =>
    segs.filter((s) => s.kind !== "del").map((s) => ({ kind: "text", text: s.text, rsid: s.rsid }));
  /** start から minutes 分、20〜80秒ごとに語句を入力する */
  const type = (segs: Segment[], start: string, minutes: number) => {
    const rsid = newSession();
    let t = new Date(start).getTime();
    const until = t + minutes * 60_000;
    while (t < until) {
      t += 20_000 + Math.floor(rnd() * 60_000);
      segs.push({ kind: "ins", text: WORDS[Math.floor(rnd() * WORDS.length)], author, date: new Date(t), id: id++, rsid });
    }
  };
  /** 承諾済みの本文の語句を、いくつか削除する */
  const trim = (segs: Segment[], start: string, count: number) => {
    let t = new Date(start).getTime();
    const texts = segs.map((s, i) => (s.kind === "text" ? i : -1)).filter((i) => i >= 0);
    for (let k = 0; k < count && texts.length; k++) {
      const i = texts.splice(Math.floor(rnd() * texts.length), 1)[0];
      const s = segs[i];
      t += 30_000 + Math.floor(rnd() * 90_000);
      segs[i] = { kind: "del", text: s.text, author, date: new Date(t), id: id++, rsid: s.rsid };
    }
  };
  const dir = path.join(__dirname, "..", "fixtures", "snapshot-demo");
  const write = (n: number, segs: Segment[]) =>
    writeDocx(path.join(dir, `sprint${n}`, "thesis.docx"), segs, settingsWithRsids(root, rsids));

  // 第1回 (2週間): 3回に分けて書き始める
  const s1: Segment[] = [];
  type(s1, "2026-10-06T01:00:00Z", 50);
  type(s1, "2026-10-09T06:00:00Z", 40);
  type(s1, "2026-10-14T02:30:00Z", 60);
  await write(1, s1);

  // 第2回: 承諾して返したものを推敲 (確定した語句の削除) し、書き進める
  const s2 = accept(s1);
  trim(s2, "2026-10-20T01:00:00Z", 6);
  type(s2, "2026-10-20T01:20:00Z", 45);
  type(s2, "2026-10-27T05:00:00Z", 70);
  await write(2, s2);

  // 第3回: 推敲・入力に加えて、外部で作った文章の貼り付けと、記録をオフにして入力した段落
  const s3 = accept(s2);
  trim(s3, "2026-11-03T02:00:00Z", 8);
  type(s3, "2026-11-03T02:30:00Z", 40);
  const pasteRsid = newSession();
  const pasted =
    "先行研究では、面接評価の自動化において主に音声認識の精度や語彙の多様性が指標として用いられてきた。" +
    "しかし、これらの指標は発話の論理的なつながりを十分に捉えられないという課題が指摘されている。" +
    "そこで本研究では、発話の時系列構造に着目した新たな評価指標を提案し、その妥当性を検証する。" +
    "具体的には、発話をいくつかの区間に分け、区間どうしの意味的な近さを埋め込みベクトルから求めたうえで、" +
    "論理展開の自然さを点数化する。";
  s3.push({ kind: "ins", text: pasted, author, date: new Date("2026-11-08T07:10:00Z"), id: id++, rsid: pasteRsid });
  type(s3, "2026-11-08T07:12:00Z", 35);
  const untracked = newSession();
  s3.push({ kind: "text", text: "（記録をオフにして入力した段落）今後の課題として、評価データの拡充が挙げられる。", rsid: untracked });
  await write(3, s3);
}

async function main() {
  fs.mkdirSync(path.join(__dirname, "..", "fixtures"), { recursive: true });
  await makeNatural();
  await makeSuspicious();
  await makeMultiSession();
  await makeFlowDemo();
  await makeChartDemo();
  await makeSnapshotDemo();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
