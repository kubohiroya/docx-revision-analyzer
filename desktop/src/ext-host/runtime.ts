/**
 * 拡張の実行環境 (拡張ごとの、表示しない sandbox のページ)。
 * 拡張のエントリ (ES モジュール) を読み込んで activate(host) を呼び、HostApi の呼び出しをメインプロセスへ中継する。
 * 分類器に渡す windowsFor は、ライブラリのコアでこのページの中で計算する。
 */
import { detectInsertionWindows, PositionedRevisionEvent } from "../../../src/core";
import type {
  AnalysisResult,
  CategorySpec,
  ClassifierContext,
  ClassifierSpec,
  DialogSpec,
  FigureAnnotatorSpec,
  FigureContext,
  FormSpec,
  HighlightSpec,
  HostApi,
  PanelSpec,
  PositionedEvent,
  WindowOptions,
} from "../extension-api";

interface Bridge {
  call(method: string, args?: unknown): Promise<unknown>;
  onInvoke(handler: (method: string, args: unknown) => Promise<unknown>): void;
  ready(): void;
}

const rawBridge = (window as unknown as { extBridge: Bridge }).extBridge;
/** メインプロセスのエラーを、Electron の接頭辞を除いて拡張に渡す */
const bridge: Bridge = {
  ...rawBridge,
  call: (method, args) =>
    rawBridge.call(method, args).catch((err: unknown) => {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(msg.replace(/^Error invoking remote method '[^']+': (Error: )?/, ""));
    }),
};
const classifiers = new Map<string, ClassifierSpec>();
const annotators = new Map<string, FigureAnnotatorSpec>();
const analysisCallbacks: ((r: AnalysisResult) => void | Promise<void>)[] = [];

/** JSON の位置付きイベントを、ライブラリの型 (日時は Date) に戻す */
function toLib(e: PositionedEvent): PositionedRevisionEvent {
  return {
    ...e,
    date: e.date ? new Date(e.date) : undefined,
    author: e.author ?? undefined,
    moveName: e.moveName ?? undefined,
    atEdit: e.atEdit ?? undefined,
  };
}

function makeHost(app: HostApi["app"]): HostApi {
  return {
    registerCategory: (c: CategorySpec) => void bridge.call("registerCategory", c),
    registerClassifier: (c: ClassifierSpec) => {
      classifiers.set(c.id, c);
      void bridge.call("registerClassifier", { id: c.id, version: c.version });
    },
    registerFigureAnnotator: (a: FigureAnnotatorSpec) => {
      annotators.set(a.id, a);
      void bridge.call("registerFigureAnnotator", { id: a.id, version: a.version });
    },
    onAnalysisComplete: (cb) => void analysisCallbacks.push(cb),
    ui: {
      showDialog: (d: DialogSpec) => bridge.call("ui.showDialog", d) as Promise<{ buttonId: string | null }>,
      showPanel: (p: PanelSpec) => void bridge.call("ui.showPanel", p),
      openForm: (f: FormSpec) => bridge.call("ui.openForm", f) as Promise<Record<string, string | number | boolean> | null>,
    },
    storage: {
      get: (key) => bridge.call("storage.get", { key }) as Promise<never>,
      set: async (key, value) => void (await bridge.call("storage.set", { key, value })),
      delete: async (key) => void (await bridge.call("storage.delete", { key })),
      keys: () => bridge.call("storage.keys", {}) as Promise<string[]>,
    },
    net: {
      post: (url, body, options) =>
        bridge.call("net.post", { url, body, confirm: options?.confirm === true }) as ReturnType<HostApi["net"]["post"]>,
    },
    app: Object.freeze({ ...app }),
  };
}

bridge.onInvoke(async (method, args) => {
  const a = args as Record<string, unknown>;
  if (method === "activate") {
    const mod = (await import(String(a.entryUrl))) as { activate?: unknown; default?: { activate?: unknown } };
    const activate = (mod.activate ?? mod.default?.activate) as ((h: HostApi) => unknown) | undefined;
    if (typeof activate !== "function") throw new Error("the extension entry must export activate(host)");
    await activate(makeHost({ version: String(a.version), locale: a.locale as "ja" | "en", apiVersion: Number(a.apiVersion) }));
    return true;
  }
  if (method === "classify") {
    const c = classifiers.get(String(a.classifierId));
    if (!c) throw new Error(`unknown classifier ${String(a.classifierId)}`);
    const raw = a.ctx as Omit<ClassifierContext, "windowsFor">;
    const lib = raw.positioned.map(toLib);
    const ctx: ClassifierContext = {
      ...raw,
      windowsFor(options: WindowOptions) {
        const r = detectInsertionWindows(lib, options);
        return {
          timeResolutionSec: r.timeResolutionSec,
          windows: r.windows.map((w) => ({
            index: w.index,
            start: w.start.toISOString(),
            end: w.end.toISOString(),
            authors: w.authors,
            eventIds: w.eventIds,
            range: w.range,
            features: { ...w.features },
          })),
        };
      },
    };
    const out = await c.classify(ctx);
    if (!Array.isArray(out)) throw new Error("classify must return an array");
    // JSON にできる形だけを返す
    return JSON.parse(JSON.stringify(out)) as HighlightSpec[];
  }
  if (method === "annotateFigure") {
    const an = annotators.get(String(a.annotatorId));
    if (!an) throw new Error(`unknown figure annotator ${String(a.annotatorId)}`);
    const out = await an.annotate(a.ctx as FigureContext);
    if (!Array.isArray(out)) throw new Error("annotate must return an array");
    return JSON.parse(JSON.stringify(out));
  }
  if (method === "analysisComplete") {
    for (const cb of analysisCallbacks) await cb(a as unknown as AnalysisResult);
    return true;
  }
  throw new Error(`unknown method ${method}`);
});

bridge.ready();
