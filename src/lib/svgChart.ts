import { Bucket, buildBuckets, BucketSpec } from "./timeBuckets";
import { Session } from "./sessions";

export interface ChartOptions {
  width?: number;
  height?: number;
  title?: string;
  /** 文字コード表記用フォント */
  fontFamily?: string;
}

function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmtTime(d: Date, spanMs: number): string {
  const useDate = spanMs > 1000 * 60 * 60 * 36; // 36時間超なら日付表示
  const pad = (n: number) => String(n).padStart(2, "0");
  if (useDate) {
    return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())}`;
  }
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/**
 * 変更履歴のバケット集計から SVG チャートを生成する。
 * - 上向き棒: バケット内の追加文字数
 * - 下向き棒: バケット内の削除文字数
 * - 折れ線 (右軸): バケット終了時点での文書総文字数
 */
export function renderRevisionChart(buckets: Bucket[], opts: ChartOptions = {}): string {
  const width = opts.width ?? 1100;
  const height = opts.height ?? 550;
  const fontFamily =
    opts.fontFamily ??
    "'Noto Sans CJK JP', 'Noto Sans JP', 'Yu Gothic', 'Hiragino Sans', Meiryo, 'MS PGothic', 'Helvetica Neue', Arial, sans-serif";
  const title = opts.title ?? "編集履歴 (追加/削除文字数・総文字数)";

  const marginLeft = 70;
  const marginRight = 70;
  const marginTop = 60;
  const marginBottom = 70;
  const plotW = width - marginLeft - marginRight;
  const plotH = height - marginTop - marginBottom;

  if (buckets.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
      <rect width="100%" height="100%" fill="#ffffff"/>
      <text x="50%" y="50%" text-anchor="middle" font-family="${fontFamily}" font-size="16" fill="#555">
        変更履歴 (挿入/削除) が見つかりませんでした
      </text>
    </svg>`;
  }

  const maxAdded = Math.max(1, ...buckets.map((b) => b.added));
  const maxDeleted = Math.max(1, ...buckets.map((b) => b.deleted));
  const barScaleMax = Math.max(maxAdded, maxDeleted); // 追加/削除を同一スケールで比較できるようにする

  const totals = buckets.map((b) => b.totalAtEnd);
  const minTotal = Math.min(...totals, 0);
  const maxTotal = Math.max(...totals, 1);
  const totalRange = Math.max(1, maxTotal - minTotal);

  // ゼロ軸 (追加/削除の境界) の垂直位置。追加側・削除側それぞれに半分ずつ割り当てる。
  const zeroY = marginTop + plotH / 2;
  const halfH = plotH / 2 - 10; // 上下に少し余白

  const n = buckets.length;
  const bandW = plotW / n;
  const barW = Math.max(1, bandW * 0.7);

  const xCenter = (i: number) => marginLeft + bandW * i + bandW / 2;
  const yAdded = (v: number) => zeroY - (v / barScaleMax) * halfH;
  const yDeleted = (v: number) => zeroY + (v / barScaleMax) * halfH;
  const yTotal = (v: number) =>
    marginTop + plotH - ((v - minTotal) / totalRange) * plotH;

  const spanMs = buckets[buckets.length - 1].end.getTime() - buckets[0].start.getTime();

  // --- 棒グラフ ---
  const bars: string[] = [];
  for (let i = 0; i < n; i++) {
    const b = buckets[i];
    const x = marginLeft + bandW * i + (bandW - barW) / 2;
    if (b.added > 0) {
      const y = yAdded(b.added);
      bars.push(
        `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${barW.toFixed(2)}" height="${(
          zeroY - y
        ).toFixed(2)}" fill="#2e7d32" fill-opacity="0.85"><title>追加 ${b.added}文字 (${b.start.toISOString()})</title></rect>`
      );
    }
    if (b.deleted > 0) {
      const y = yDeleted(b.deleted);
      bars.push(
        `<rect x="${x.toFixed(2)}" y="${zeroY.toFixed(2)}" width="${barW.toFixed(2)}" height="${(
          y - zeroY
        ).toFixed(2)}" fill="#c62828" fill-opacity="0.85"><title>削除 ${b.deleted}文字 (${b.start.toISOString()})</title></rect>`
      );
    }
  }

  // --- 総文字数の折れ線 ---
  const linePoints = buckets
    .map((b, i) => `${xCenter(i).toFixed(2)},${yTotal(b.totalAtEnd).toFixed(2)}`)
    .join(" ");

  // --- 軸ラベル (X軸: 間引いて表示) ---
  const maxLabels = 10;
  const labelStep = Math.max(1, Math.ceil(n / maxLabels));
  const xLabels: string[] = [];
  for (let i = 0; i < n; i += labelStep) {
    const x = xCenter(i);
    xLabels.push(
      `<text x="${x.toFixed(2)}" y="${(marginTop + plotH + 20).toFixed(2)}" font-size="11" fill="#333" text-anchor="middle">${esc(
        fmtTime(buckets[i].start, spanMs)
      )}</text>`
    );
  }

  // --- 左軸 (追加/削除 文字数) 目盛り ---
  const leftTicks: string[] = [];
  const leftTickCount = 4;
  for (let t = -leftTickCount; t <= leftTickCount; t++) {
    const val = (t / leftTickCount) * barScaleMax;
    const y = zeroY - (val / barScaleMax) * halfH;
    leftTicks.push(
      `<line x1="${marginLeft - 5}" y1="${y.toFixed(2)}" x2="${marginLeft}" y2="${y.toFixed(
        2
      )}" stroke="#999"/>`
    );
    leftTicks.push(
      `<text x="${marginLeft - 10}" y="${(y + 4).toFixed(2)}" font-size="10" fill="#555" text-anchor="end">${Math.round(
        Math.abs(val)
      )}</text>`
    );
  }

  // --- 右軸 (総文字数) 目盛り ---
  const rightTicks: string[] = [];
  const rightTickCount = 4;
  for (let t = 0; t <= rightTickCount; t++) {
    const val = minTotal + (t / rightTickCount) * totalRange;
    const y = yTotal(val);
    rightTicks.push(
      `<line x1="${(marginLeft + plotW).toFixed(2)}" y1="${y.toFixed(2)}" x2="${(
        marginLeft + plotW + 5
      ).toFixed(2)}" y2="${y.toFixed(2)}" stroke="#1565c0"/>`
    );
    rightTicks.push(
      `<text x="${(marginLeft + plotW + 10).toFixed(2)}" y="${(y + 4).toFixed(
        2
      )}" font-size="10" fill="#1565c0" text-anchor="start">${Math.round(val)}</text>`
    );
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="${fontFamily}">
  <rect width="100%" height="100%" fill="#ffffff"/>
  <text x="${width / 2}" y="28" text-anchor="middle" font-size="18" font-weight="bold" fill="#1a1a1a">${esc(
    title
  )}</text>

  <!-- ゼロ軸 -->
  <line x1="${marginLeft}" y1="${zeroY}" x2="${marginLeft + plotW}" y2="${zeroY}" stroke="#333" stroke-width="1"/>
  <!-- プロット枠 -->
  <line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${marginTop + plotH}" stroke="#333" stroke-width="1"/>
  <line x1="${marginLeft + plotW}" y1="${marginTop}" x2="${marginLeft + plotW}" y2="${
    marginTop + plotH
  }" stroke="#ccd7e6" stroke-width="1"/>

  ${leftTicks.join("\n  ")}
  ${rightTicks.join("\n  ")}
  ${xLabels.join("\n  ")}

  <g>${bars.join("\n  ")}</g>

  <polyline points="${linePoints}" fill="none" stroke="#1565c0" stroke-width="2.5"/>
  <g fill="#1565c0">${buckets
    .map((b, i) => `<circle cx="${xCenter(i).toFixed(2)}" cy="${yTotal(b.totalAtEnd).toFixed(2)}" r="2.5"><title>総文字数 ${b.totalAtEnd} (${b.end.toISOString()})</title></circle>`)
    .join("")}</g>

  <text x="${marginLeft - 50}" y="${marginTop - 10}" font-size="11" fill="#555">文字数(追加/削除)</text>
  <text x="${marginLeft + plotW - 20}" y="${marginTop - 10}" font-size="11" fill="#1565c0">総文字数</text>

  <!-- 凡例 -->
  <g transform="translate(${marginLeft}, ${height - 22})">
    <rect x="0" y="-10" width="14" height="14" fill="#2e7d32" fill-opacity="0.85"/>
    <text x="20" y="1" font-size="12" fill="#333">追加</text>
    <rect x="70" y="-10" width="14" height="14" fill="#c62828" fill-opacity="0.85"/>
    <text x="90" y="1" font-size="12" fill="#333">削除</text>
    <line x1="140" y1="-3" x2="160" y2="-3" stroke="#1565c0" stroke-width="2.5"/>
    <text x="166" y="1" font-size="12" fill="#333">総文字数</text>
  </g>
</svg>`;
}

// ============================================================================
// セッション分割チャート (-p / --gap-threshold オプション用)
// ============================================================================

/** 文字幅の粗い近似 (プロポーショナルフォント、半角英数字・記号を想定) */
function estimateTextWidth(text: string, fontSize: number): number {
  return text.length * fontSize * 0.62;
}

/** 無編集期間 (時間) を "12 h" のような表記に整形する */
function formatGapHours(hours: number): string {
  const rounded = Math.round(hours * 10) / 10;
  const text = Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
  return `${text} h`;
}

interface PanelGeom {
  session: Session;
  buckets: Bucket[];
  x: number; // このパネルの左端 (プロット座標系)
  width: number;
  gapBeforeLabel: string | null; // このパネルの直前に描画するギャップ表記 (先頭パネルはnull)
  gapBeforeWidth: number; // 直前のギャップの幅 (先頭パネルは0)
}

export interface SessionedChartOptions extends ChartOptions {
  /** -p で指定した閾値 (時間)。タイトル直下の説明表記に使う */
  gapThresholdHours?: number;
  /** 1バケットあたりのおおよその横幅(px) */
  pixelsPerBucket?: number;
}

/**
 * 「連続的に更新が行われた期間 (セッション)」ごとに個別のバー+折れ線チャートを作成し、
 * 水平方向に並べて1枚のSVGにまとめる。セッションとセッションの間には、無編集期間の
 * 長さ (例: "12 h") を表示する程度の狭い間隔を空ける。
 *
 * 各セッションは独立に時間バケットへ集計する (buildBuckets の "auto" は目標バケット数を
 * 一定に保とうとするため、セッション単位で区切ることで結果的に元の1本の全体チャートより
 * 細かい時間粒度で描画される = 「詳細に作成する」)。
 *
 * 追加/削除文字数の縦軸スケール、および総文字数の縦軸スケールは、すべてのパネルで
 * 共通 (文書全体を通した同一スケール) にし、パネル間で高さを比較できるようにする。
 * 総文字数の折れ線は、無編集期間をまたいで直線でつながないよう、パネルごとに
 * 独立したポリラインとして描画する (時間が実際に経過していない空白を線で埋めない)。
 */
export function renderSessionedRevisionChart(
  sessions: Session[],
  baselineCharCount: number,
  bucketSpec: BucketSpec,
  opts: SessionedChartOptions = {}
): string {
  const height = opts.height ?? 550;
  const fontFamily =
    opts.fontFamily ??
    "'Noto Sans CJK JP', 'Noto Sans JP', 'Yu Gothic', 'Hiragino Sans', Meiryo, 'MS PGothic', 'Helvetica Neue', Arial, sans-serif";
  const title = opts.title ?? "編集履歴 (期間ごと・追加/削除文字数・総文字数)";
  const pixelsPerBucket = opts.pixelsPerBucket ?? 16;
  const axisFontSize = 11;

  const marginLeft = 70;
  const marginRight = 70;
  const marginTop = opts.gapThresholdHours !== undefined ? 78 : 60;
  const marginBottom = 70;
  const plotH = height - marginTop - marginBottom;

  if (sessions.length === 0) {
    const width = opts.width ?? 1100;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
      <rect width="100%" height="100%" fill="#ffffff"/>
      <text x="50%" y="50%" text-anchor="middle" font-family="${fontFamily}" font-size="16" fill="#555">
        変更履歴 (挿入/削除) が見つかりませんでした
      </text>
    </svg>`;
  }

  // --- 各セッションを独立にバケット集計しつつ、累計総文字数はセッションをまたいで引き継ぐ ---
  let running = baselineCharCount;
  const panels: PanelGeom[] = [];
  let cursorX = marginLeft;

  for (const session of sessions) {
    const buckets = buildBuckets(session.events, running, bucketSpec);
    for (const ev of session.events) {
      running += ev.type === "ins" ? ev.chars : -ev.chars;
    }

    let gapBeforeLabel: string | null = null;
    let gapBeforeWidth = 0;
    if (session.gapBeforeHours !== null) {
      gapBeforeLabel = formatGapHours(session.gapBeforeHours);
      const minGapWidth = 4 * axisFontSize * 0.62; // 「4文字程度」の最小幅
      gapBeforeWidth = Math.max(minGapWidth, estimateTextWidth(gapBeforeLabel, axisFontSize) + 14);
      cursorX += gapBeforeWidth;
    }

    const panelWidth = Math.max(50, buckets.length * pixelsPerBucket);
    panels.push({
      session,
      buckets,
      x: cursorX,
      width: panelWidth,
      gapBeforeLabel,
      gapBeforeWidth,
    });
    cursorX += panelWidth;
  }

  const totalWidth = cursorX + marginRight;
  const width = opts.width ?? Math.max(600, totalWidth);
  const plotRight = cursorX; // 最後のパネルの右端

  // --- 共通スケール (全パネル共通) ---
  const allBuckets = panels.flatMap((p) => p.buckets);
  const maxAdded = Math.max(1, ...allBuckets.map((b) => b.added));
  const maxDeleted = Math.max(1, ...allBuckets.map((b) => b.deleted));
  const barScaleMax = Math.max(maxAdded, maxDeleted);

  const totals = allBuckets.map((b) => b.totalAtEnd);
  const minTotal = Math.min(0, baselineCharCount, ...totals);
  const maxTotal = Math.max(1, ...totals);
  const totalRange = Math.max(1, maxTotal - minTotal);

  const zeroY = marginTop + plotH / 2;
  const halfH = plotH / 2 - 10;
  const yAdded = (v: number) => zeroY - (v / barScaleMax) * halfH;
  const yDeleted = (v: number) => zeroY + (v / barScaleMax) * halfH;
  const yTotal = (v: number) => marginTop + plotH - ((v - minTotal) / totalRange) * plotH;

  const svgParts: string[] = [];

  // --- 左軸 (追加/削除) 目盛り: 先頭パネルの左端にのみ描画 ---
  const leftTicks: string[] = [];
  const leftTickCount = 4;
  for (let t = -leftTickCount; t <= leftTickCount; t++) {
    const val = (t / leftTickCount) * barScaleMax;
    const y = zeroY - (val / barScaleMax) * halfH;
    leftTicks.push(
      `<line x1="${marginLeft - 5}" y1="${y.toFixed(2)}" x2="${marginLeft}" y2="${y.toFixed(2)}" stroke="#999"/>`
    );
    leftTicks.push(
      `<text x="${marginLeft - 10}" y="${(y + 4).toFixed(2)}" font-size="10" fill="#555" text-anchor="end">${Math.round(
        Math.abs(val)
      )}</text>`
    );
  }

  // --- 右軸 (総文字数) 目盛り: 末尾パネルの右端にのみ描画 ---
  const rightTicks: string[] = [];
  const rightTickCount = 4;
  for (let t = 0; t <= rightTickCount; t++) {
    const val = minTotal + (t / rightTickCount) * totalRange;
    const y = yTotal(val);
    rightTicks.push(
      `<line x1="${plotRight.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(plotRight + 5).toFixed(
        2
      )}" y2="${y.toFixed(2)}" stroke="#1565c0"/>`
    );
    rightTicks.push(
      `<text x="${(plotRight + 10).toFixed(2)}" y="${(y + 4).toFixed(
        2
      )}" font-size="10" fill="#1565c0" text-anchor="start">${Math.round(val)}</text>`
    );
  }

  // --- ゼロ軸・プロット枠 (パネルごとに区切って描画) ---
  const frames: string[] = [];
  for (const p of panels) {
    frames.push(
      `<line x1="${p.x.toFixed(2)}" y1="${zeroY}" x2="${(p.x + p.width).toFixed(2)}" y2="${zeroY}" stroke="#333" stroke-width="1"/>`
    );
  }

  // --- ギャップ (無編集期間) の区切り表示 ---
  const gapMarks: string[] = [];
  for (const p of panels) {
    if (p.gapBeforeLabel === null) continue;
    const gapStart = p.x - p.gapBeforeWidth;
    const gapEnd = p.x;
    const gapMid = (gapStart + gapEnd) / 2;
    gapMarks.push(
      `<line x1="${gapStart.toFixed(2)}" y1="${marginTop}" x2="${gapStart.toFixed(2)}" y2="${(
        marginTop + plotH
      ).toFixed(2)}" stroke="#bbb" stroke-dasharray="2,3"/>`,
      `<line x1="${gapEnd.toFixed(2)}" y1="${marginTop}" x2="${gapEnd.toFixed(2)}" y2="${(
        marginTop + plotH
      ).toFixed(2)}" stroke="#bbb" stroke-dasharray="2,3"/>`,
      `<rect x="${(gapMid - estimateTextWidth(p.gapBeforeLabel, axisFontSize) / 2 - 3).toFixed(
        2
      )}" y="${(zeroY - 8).toFixed(2)}" width="${(
        estimateTextWidth(p.gapBeforeLabel, axisFontSize) + 6
      ).toFixed(2)}" height="16" fill="#ffffff" fill-opacity="0.85"/>`,
      `<text x="${gapMid.toFixed(2)}" y="${(zeroY + 4).toFixed(
        2
      )}" font-size="${axisFontSize}" fill="#777" text-anchor="middle">${esc(p.gapBeforeLabel)}</text>`
    );
  }

  // --- パネルごとの棒・折れ線・X軸ラベル ---
  const panelContents: string[] = [];
  for (const p of panels) {
    const n = p.buckets.length;
    if (n === 0) continue;
    const bandW = p.width / n;
    const barW = Math.max(1, bandW * 0.7);
    const xCenter = (i: number) => p.x + bandW * i + bandW / 2;
    const spanMs = p.buckets[n - 1].end.getTime() - p.buckets[0].start.getTime();

    const bars: string[] = [];
    for (let i = 0; i < n; i++) {
      const b = p.buckets[i];
      const x = p.x + bandW * i + (bandW - barW) / 2;
      if (b.added > 0) {
        const y = yAdded(b.added);
        bars.push(
          `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${barW.toFixed(2)}" height="${(
            zeroY - y
          ).toFixed(2)}" fill="#2e7d32" fill-opacity="0.85"><title>追加 ${b.added}文字 (${b.start.toISOString()})</title></rect>`
        );
      }
      if (b.deleted > 0) {
        const y = yDeleted(b.deleted);
        bars.push(
          `<rect x="${x.toFixed(2)}" y="${zeroY.toFixed(2)}" width="${barW.toFixed(2)}" height="${(
            y - zeroY
          ).toFixed(2)}" fill="#c62828" fill-opacity="0.85"><title>削除 ${b.deleted}文字 (${b.start.toISOString()})</title></rect>`
        );
      }
    }

    const linePoints = p.buckets
      .map((b, i) => `${xCenter(i).toFixed(2)},${yTotal(b.totalAtEnd).toFixed(2)}`)
      .join(" ");
    const dots = p.buckets
      .map(
        (b, i) =>
          `<circle cx="${xCenter(i).toFixed(2)}" cy="${yTotal(b.totalAtEnd).toFixed(
            2
          )}" r="2.5"><title>総文字数 ${b.totalAtEnd} (${b.end.toISOString()})</title></circle>`
      )
      .join("");

    const maxLabels = Math.max(2, Math.floor(p.width / 70));
    const labelStep = Math.max(1, Math.ceil(n / maxLabels));
    const xLabels: string[] = [];
    for (let i = 0; i < n; i += labelStep) {
      xLabels.push(
        `<text x="${xCenter(i).toFixed(2)}" y="${(marginTop + plotH + 20).toFixed(
          2
        )}" font-size="10" fill="#333" text-anchor="middle">${esc(
          fmtTime(p.buckets[i].start, spanMs)
        )}</text>`
      );
    }

    panelContents.push(
      `<g>${bars.join("")}</g>` +
        `<polyline points="${linePoints}" fill="none" stroke="#1565c0" stroke-width="2.5"/>` +
        `<g fill="#1565c0">${dots}</g>` +
        `${xLabels.join("")}`
    );
  }

  const thresholdNote =
    opts.gapThresholdHours !== undefined
      ? `<text x="${width / 2}" y="48" text-anchor="middle" font-size="12" fill="#777">(${opts.gapThresholdHours}時間以上更新が無い期間で区切って表示)</text>`
      : "";

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="${fontFamily}">
  <rect width="100%" height="100%" fill="#ffffff"/>
  <text x="${width / 2}" y="28" text-anchor="middle" font-size="18" font-weight="bold" fill="#1a1a1a">${esc(
    title
  )}</text>
  ${thresholdNote}

  ${frames.join("\n  ")}

  ${leftTicks.join("\n  ")}
  ${rightTicks.join("\n  ")}
  ${gapMarks.join("\n  ")}

  ${panelContents.join("\n  ")}

  <text x="${marginLeft - 50}" y="${marginTop - 10}" font-size="11" fill="#555">文字数(追加/削除)</text>
  <text x="${plotRight - 20}" y="${marginTop - 10}" font-size="11" fill="#1565c0">総文字数</text>

  <!-- 凡例 -->
  <g transform="translate(${marginLeft}, ${height - 22})">
    <rect x="0" y="-10" width="14" height="14" fill="#2e7d32" fill-opacity="0.85"/>
    <text x="20" y="1" font-size="12" fill="#333">追加</text>
    <rect x="70" y="-10" width="14" height="14" fill="#c62828" fill-opacity="0.85"/>
    <text x="90" y="1" font-size="12" fill="#333">削除</text>
    <line x1="140" y1="-3" x2="160" y2="-3" stroke="#1565c0" stroke-width="2.5"/>
    <text x="166" y="1" font-size="12" fill="#333">総文字数</text>
  </g>
</svg>`;
}
