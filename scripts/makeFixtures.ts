/**
 * テスト用の .docx フィクスチャを生成するスクリプト。
 * 実際のWordファイルが手元になくても、両CLIの動作確認ができるようにする。
 *
 * 生成するファイル:
 *  - fixtures/natural-writing.docx   : 人が時間をかけて少しずつタイプしたことを想定
 *  - fixtures/suspicious-paste.docx  : 最初は少し自分でタイプした後、大きな塊を
 *                                       一瞬で貼り付けたことを想定 (AI生成文の貼付を模擬)
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

async function main() {
  fs.mkdirSync(path.join(__dirname, "..", "fixtures"), { recursive: true });
  await makeNatural();
  await makeSuspicious();
  await makeMultiSession();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
