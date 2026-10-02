// @ts-check
/**
 * 拡張機能のサンプル (API v1)。
 *  - カテゴリ「長い単独の挿入」を登録し、1回の挿入 (w:ins) で 200 文字以上入った箇所をハイライトする分類器を登録する
 *  - 図の注釈: 長い単独の挿入があった段落 (flow) と棒 (chart) に、マウスオーバーの説明とポップアップ、リンクを付ける。
 *    リンク先は、OneDrive / SharePoint の文書なら元の文書の該当箇所 (target.docLink。段落を含むセクション
 *    (ブックマーク) があればそこ、無ければ冒頭)、それ以外は README。flow の区間のキャプションには、区間のまとめのポップアップを付ける。
 *    段落の冒頭 (約 40 文字) は、アプリがポップアップの先頭に添える
 *  - 解析が終わるたびに、区間・文字数・ハイライトの件数と、この拡張で解析した回数 (storage) をパネルに表示する
 * 本文は使わない (documentText の権限を求めない) ため、解析結果の本文は空文字列で届く。
 *
 * @typedef {import("../../src/extension-api").HostApi} HostApi
 */

const ID = "io.github.kubohiroya.sample-summary";
const CATEGORY = `${ID}.long-single`;
const MIN_CHARS = 200;

/** @param {HostApi} host */
export function activate(host) {
  const ja = host.app.locale === "ja";

  host.registerCategory({
    id: CATEGORY,
    color: "#7C3AED",
    label: { en: "Long single insertion", ja: "長い単独の挿入" },
    priority: 5,
    pattern: "dots",
  });

  host.registerClassifier({
    id: `${ID}.long-single`,
    version: "1.0.0",
    classify(ctx) {
      return ctx.positioned
        .filter((e) => e.type === "ins" && !e.move && e.date && e.chars >= MIN_CHARS)
        .map((e) => ({
          categoryId: CATEGORY,
          eventIds: [e.id],
          start: /** @type {string} */ (e.date),
          end: /** @type {string} */ (e.date),
          features: { chars: e.chars },
          reason: {
            en: `${e.chars} chars were inserted in a single insertion`,
            ja: `1回の挿入で ${e.chars} 文字が入った`,
          },
        }));
    },
  });

  const README = "https://github.com/kubohiroya/docx-revision-analyzer#readme";

  host.registerFigureAnnotator({
    id: `${ID}.figure`,
    version: "1.0.0",
    annotate({ figure, targets, analysis }) {
      const mine = analysis.highlights.filter((h) => h.categoryId === CATEGORY);
      const byId = new Map(analysis.positioned.map((e) => [e.id, e]));
      /** @type {import("../../src/extension-api").FigureAnnotationSpec[]} */
      const out = [];
      for (const h of mine) {
        const t = Date.parse(h.start);
        const chars = Number(h.features?.chars ?? 0);
        const tooltip = {
          en: `Long single insertion: ${chars} chars at ${new Date(t).toLocaleString("en-US")}`,
          ja: `長い単独の挿入: ${new Date(t).toLocaleString("ja-JP")} に ${chars} 文字`,
        };
        if (figure === "flow") {
          // ハイライトの時刻を含む区間の、終了時点の列 (列 = 区間 + 1) の段落
          const session = analysis.sessions.findIndex((s) => Date.parse(s.start) <= t && t <= Date.parse(s.end));
          const paras = new Set(h.eventIds.flatMap((id) => byId.get(id)?.paraModelIndices ?? []));
          for (const tg of targets) {
            if (tg.kind === "paragraph" && tg.column === session + 1 && paras.has(tg.paraIndex)) {
              out.push({
                target: tg.key,
                tooltip,
                popup: {
                  title: { en: "Long single insertion", ja: "長い単独の挿入" },
                  blocks: [
                    { type: "keyValue", rows: [{ key: { en: "Characters", ja: "文字数" }, value: chars }, { key: { en: "Session", ja: "区間" }, value: session + 1 }] },
                    { type: "text", text: { en: "Was this pasted, or typed in one go?", ja: "貼り付けたもの、それとも一気に入力したもの?" } },
                  ],
                },
                href: tg.docLink ?? README,
              });
            }
          }
        } else {
          for (const tg of targets) {
            if (tg.kind === "bar" && Date.parse(tg.start) <= t && t < Date.parse(tg.end)) {
              out.push({ target: tg.key, tooltip, href: tg.docLink ?? README });
            }
          }
        }
      }
      if (figure === "flow") {
        for (const tg of targets) {
          if (tg.kind !== "caption") continue;
          const s = analysis.sessions[tg.session];
          out.push({
            target: tg.key,
            popup: {
              title: { en: `Session ${tg.session + 1}`, ja: `区間 ${tg.session + 1}` },
              blocks: [
                {
                  type: "keyValue",
                  rows: [
                    { key: { en: "Inserted", ja: "挿入" }, value: s.insChars },
                    { key: { en: "Deleted", ja: "削除" }, value: s.delChars },
                    { key: { en: "Bulk-inserted", ja: "一括挿入" }, value: s.bulkInsChars },
                  ],
                },
              ],
            },
          });
        }
      }
      return out;
    },
  });

  host.onAnalysisComplete(async (r) => {
    const runs = /** @type {number} */ ((await host.storage.get("runs")) ?? 0) + 1;
    await host.storage.set("runs", runs);
    /** @type {Record<string, number>} */
    const byCategory = {};
    for (const h of r.highlights) byCategory[h.categoryId] = (byCategory[h.categoryId] ?? 0) + 1;
    host.ui.showPanel({
      id: "summary",
      title: { en: "Writing summary", ja: "執筆のまとめ" },
      blocks: [
        {
          type: "keyValue",
          rows: [
            { key: { en: "Editing sessions", ja: "編集の区間" }, value: r.sessions.length },
            { key: { en: "Inserted characters", ja: "挿入した文字数" }, value: r.revisions.totalInserted },
            { key: { en: "Deleted characters", ja: "削除した文字数" }, value: r.revisions.totalDeleted },
            { key: { en: "Insertion windows", ja: "挿入の窓" }, value: r.windows.length },
          ],
        },
        {
          type: "table",
          columns: [{ en: "Category", ja: "分類" }, { en: "Highlights", ja: "件数" }],
          rows: Object.entries(byCategory).map(([k, v]) => [k, v]),
        },
        {
          type: "text",
          text: ja
            ? `この拡張機能での解析は ${runs} 回目です (このコンピュータに保存)。`
            : `This is analysis #${runs} with this extension (stored on this computer).`,
        },
      ],
    });
  });
}
