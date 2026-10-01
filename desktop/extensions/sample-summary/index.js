// @ts-check
/**
 * 拡張機能のサンプル (API v1)。
 *  - カテゴリ「長い単独の挿入」を登録し、1回の挿入 (w:ins) で 200 文字以上入った箇所をハイライトする分類器を登録する
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
