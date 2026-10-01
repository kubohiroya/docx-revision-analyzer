/**
 * 拡張モジュールの管理 (メインプロセス)。
 *
 *  - アプリに同梱した拡張 (extensions/<dir>/manifest.json) だけを読み込む。既定は無効で、設定で有効にする。
 *  - 拡張ごとに、表示しない sandbox のウィンドウ (専用のセッション・オリジン dra-ext://<id>) を作り、
 *    その中で拡張のコードを動かす。拡張は Node.js の権限を持たず、ネットワークにも出られない
 *    (CSP と webRequest で遮断)。アプリとは ext-preload.ts の RPC だけでやり取りする。
 *  - 拡張からの呼び出しは manifest の permissions で確かめる。UI はアプリのウィンドウが描画する。
 *  - 拡張の呼び出しには時間制限を設け、例外・時間切れ・クラッシュはその拡張だけの失敗として扱う。
 */

import { BrowserWindow, ipcMain, session, WebContents } from "electron";
import * as fs from "fs";
import * as path from "path";
import type {
  AnalysisResult,
  CategorySpec,
  ClassifierContext,
  DialogSpec,
  ExtensionManifest,
  FormSpec,
  HighlightSpec,
  NetResponse,
  PanelSpec,
  PositionedEvent,
} from "./extension-api";
import { SUPPORTED_API_VERSIONS } from "./extension-api";

export const EXT_SCHEME = "dra-ext";

const ACTIVATE_TIMEOUT_MS = 10_000;
const CLASSIFY_TIMEOUT_MS = 10_000;
const ANALYSIS_TIMEOUT_MS = 30_000;
const MAX_STORAGE_BYTES = 1_000_000;

export type ExtensionState = "disabled" | "starting" | "active" | "error";

export interface ExtensionInfo {
  id: string;
  dir: string;
  manifest?: ExtensionManifest;
  state: ExtensionState;
  error?: string;
  categories: CategorySpec[];
  classifiers: { id: string; version: string }[];
}

/** アプリのウィンドウへ、拡張の UI の表示を頼む関数 */
export type UiRequester = (
  kind: "dialog" | "panel" | "form" | "confirmSend",
  ext: { id: string; name: ExtensionManifest["name"] },
  spec: DialogSpec | PanelSpec | FormSpec | { url: string; body: string }
) => Promise<unknown>;

/** net.post を仲介する関数 (sendLog.ts。allowed は宛先が manifest で宣言されたオリジンか) */
export type NetPoster = (
  ext: ExtensionManifest,
  url: string,
  body: unknown,
  opts: { confirm?: boolean; allowed: boolean }
) => Promise<NetResponse>;

interface Host {
  info: ExtensionInfo;
  win?: BrowserWindow;
  ready?: Promise<void>;
  pending: Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }>;
}

const ID_RE = /^[a-z0-9]+([.\-_][a-z0-9]+)+$/;

/** manifest を検証する。問題があれば理由を返す */
export function validateManifest(m: unknown): string | undefined {
  if (typeof m !== "object" || m === null) return "manifest.json must be an object";
  const v = m as Record<string, unknown>;
  if (typeof v.id !== "string" || !ID_RE.test(v.id)) return "id must be like com.example.sample";
  if (typeof v.version !== "string" || v.version === "") return "version is required";
  if (typeof v.main !== "string" || v.main.includes("..") || path.isAbsolute(v.main)) return "main must be a relative path";
  if (!SUPPORTED_API_VERSIONS.includes(v.apiVersion as 1)) {
    return `apiVersion ${String(v.apiVersion)} is not supported (supported: ${SUPPORTED_API_VERSIONS.join(", ")})`;
  }
  const p = (v.permissions ?? {}) as Record<string, unknown>;
  for (const origin of (p.network as unknown[]) ?? []) {
    try {
      const u = new URL(String(origin));
      if (u.protocol !== "https:" || u.origin !== String(origin).replace(/\/$/, "")) return `network: "${origin}" must be an https origin`;
    } catch {
      return `network: "${origin}" is not a URL`;
    }
  }
  return undefined;
}

export class ExtensionManager {
  private hosts = new Map<string, Host>();
  private byWebContents = new Map<number, Host>();
  private nextId = 1;

  constructor(
    private readonly extensionsDir: string,
    private readonly hostPageDir: string,
    private readonly preloadPath: string,
    private readonly storageDir: string,
    private readonly appInfo: () => { version: string; locale: "ja" | "en" },
    private readonly requestUi: UiRequester,
    private readonly post: NetPoster
  ) {
    ipcMain.handle("ext:call", (e, method: string, args: unknown) => this.handleCall(e.sender, method, args));
    ipcMain.on("ext:reply", (e, msg: { id: number; ok: boolean; value?: unknown; error?: string }) => {
      const host = this.byWebContents.get(e.sender.id);
      const p = host?.pending.get(msg.id);
      if (!host || !p) return;
      host.pending.delete(msg.id);
      clearTimeout(p.timer);
      if (msg.ok) p.resolve(msg.value);
      else p.reject(new Error(msg.error ?? "extension error"));
    });
  }

  /** 同梱の拡張を探す */
  discover(): ExtensionInfo[] {
    let dirs: string[] = [];
    try {
      dirs = fs.readdirSync(this.extensionsDir).filter((d) => fs.existsSync(path.join(this.extensionsDir, d, "manifest.json")));
    } catch {
      dirs = [];
    }
    for (const d of dirs) {
      const dir = path.join(this.extensionsDir, d);
      let manifest: ExtensionManifest | undefined;
      let error: string | undefined;
      try {
        const raw = JSON.parse(fs.readFileSync(path.join(dir, "manifest.json"), "utf-8"));
        error = validateManifest(raw);
        if (!error) manifest = raw as ExtensionManifest;
      } catch (err) {
        error = `manifest.json: ${err instanceof Error ? err.message : String(err)}`;
      }
      const id = manifest?.id ?? d;
      if (this.hosts.has(id)) continue;
      this.hosts.set(id, {
        info: { id, dir, manifest, state: error ? "error" : "disabled", error, categories: [], classifiers: [] },
        pending: new Map(),
      });
    }
    return this.list();
  }

  list(): ExtensionInfo[] {
    return [...this.hosts.values()].map((h) => ({ ...h.info }));
  }

  /** 有効/無効の設定に合わせて拡張を起動・停止する */
  async sync(enabled: Record<string, boolean>): Promise<void> {
    for (const [id, host] of this.hosts) {
      if (!host.info.manifest) continue;
      const want = !!enabled[id];
      const running = host.info.state === "active" || host.info.state === "starting";
      if (want && !running) await this.start(host);
      else if (!want && running) this.stop(host);
      else if (!want && host.info.state === "error") host.info.state = "disabled";
    }
  }

  private extSession(id: string): Electron.Session {
    const ses = session.fromPartition(`ext-${id}`); // persist: で始まらないのでメモリ上だけ
    if (!(ses as unknown as { __draLocked?: boolean }).__draLocked) {
      (ses as unknown as { __draLocked?: boolean }).__draLocked = true;
      ses.webRequest.onBeforeRequest((d, cb) => cb({ cancel: !d.url.startsWith(`${EXT_SCHEME}://`) && !d.url.startsWith("devtools:") }));
      ses.setPermissionRequestHandler((_wc, _p, cb) => cb(false));
      ses.protocol.handle(EXT_SCHEME, (req) => this.serve(id, req.url));
    }
    return ses;
  }

  /** dra-ext://<id>/... を返す。__host__/ はホストのページ、それ以外は拡張のディレクトリ */
  private async serve(id: string, url: string): Promise<Response> {
    const u = new URL(url);
    const host = this.hosts.get(id);
    if (!host || u.hostname !== id) return new Response("not found", { status: 404 });
    const rel = decodeURIComponent(u.pathname).replace(/^\/+/, "");
    const base = rel.startsWith("__host__/") ? this.hostPageDir : host.info.dir;
    const file = path.resolve(base, rel.startsWith("__host__/") ? rel.slice("__host__/".length) : rel);
    if (!file.startsWith(path.resolve(base) + path.sep)) return new Response("forbidden", { status: 403 });
    try {
      const body = await fs.promises.readFile(file);
      const type =
        { ".js": "text/javascript", ".mjs": "text/javascript", ".html": "text/html", ".json": "application/json", ".css": "text/css" }[
          path.extname(file)
        ] ?? "application/octet-stream";
      return new Response(body, {
        headers: {
          "content-type": `${type}; charset=utf-8`,
          "content-security-policy":
            "default-src 'none'; script-src 'self'; connect-src 'none'; img-src 'none'; style-src 'none'; object-src 'none'; base-uri 'none'",
        },
      });
    } catch {
      return new Response("not found", { status: 404 });
    }
  }

  private async start(host: Host): Promise<void> {
    const m = host.info.manifest!;
    host.info.state = "starting";
    host.info.error = undefined;
    host.info.categories = [];
    host.info.classifiers = [];
    const win = new BrowserWindow({
      show: false,
      webPreferences: {
        session: this.extSession(m.id),
        preload: this.preloadPath,
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
        webSecurity: true,
        spellcheck: false,
      },
    });
    host.win = win;
    this.byWebContents.set(win.webContents.id, host);
    win.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    win.webContents.on("will-navigate", (e) => e.preventDefault());
    win.webContents.on("render-process-gone", (_e, d) => this.fail(host, `the extension process stopped (${d.reason})`));
    host.ready = new Promise<void>((resolve) => {
      const onReady = (e: Electron.IpcMainEvent) => {
        if (e.sender.id !== win.webContents.id) return;
        ipcMain.removeListener("ext:ready", onReady);
        resolve();
      };
      ipcMain.on("ext:ready", onReady);
    });
    try {
      await win.loadURL(`${EXT_SCHEME}://${m.id}/__host__/index.html`);
      await withTimeout(host.ready, ACTIVATE_TIMEOUT_MS, "the extension host did not start");
      const app = this.appInfo();
      await this.invoke(host, "activate", { entryUrl: `${EXT_SCHEME}://${m.id}/${m.main}`, ...app, apiVersion: 1 }, ACTIVATE_TIMEOUT_MS);
      host.info.state = "active";
      this.onActivated?.(m.id);
    } catch (err) {
      this.fail(host, err instanceof Error ? err.message : String(err));
    }
  }

  private stop(host: Host): void {
    if (host.win && !host.win.isDestroyed()) {
      this.byWebContents.delete(host.win.webContents.id);
      host.win.destroy();
    }
    host.win = undefined;
    for (const p of host.pending.values()) {
      clearTimeout(p.timer);
      p.reject(new Error("the extension was stopped"));
    }
    host.pending.clear();
    host.info.state = "disabled";
    host.info.categories = [];
    host.info.classifiers = [];
  }

  private fail(host: Host, message: string): void {
    this.stop(host);
    host.info.state = "error";
    host.info.error = message;
  }

  private invoke(host: Host, method: string, args: unknown, timeoutMs: number): Promise<unknown> {
    const win = host.win;
    if (!win || win.isDestroyed()) return Promise.reject(new Error("the extension is not running"));
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        host.pending.delete(id);
        reject(new Error(`the extension did not answer "${method}" within ${timeoutMs / 1000} s`));
      }, timeoutMs);
      host.pending.set(id, { resolve, reject, timer });
      win.webContents.send("ext:invoke", { id, method, args });
    });
  }

  /** 拡張からの呼び出し (権限を確かめてから処理する) */
  private async handleCall(sender: WebContents, method: string, args: unknown): Promise<unknown> {
    const host = this.byWebContents.get(sender.id);
    if (!host?.info.manifest) throw new Error("unknown extension");
    const m = host.info.manifest;
    const perms = m.permissions ?? {};
    const a = args as Record<string, unknown>;
    switch (method) {
      case "registerCategory": {
        const c = args as CategorySpec;
        if (typeof c?.id !== "string" || !c.id.startsWith(`${m.id}.`)) throw new Error(`category ids must start with "${m.id}."`);
        if (typeof c.color !== "string" || !/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(c.color)) throw new Error("color must be #rrggbb");
        host.info.categories = [...host.info.categories.filter((x) => x.id !== c.id), { ...c, priority: Number(c.priority) || 0 }];
        return null;
      }
      case "registerClassifier": {
        const c = args as { id: string; version: string };
        if (typeof c?.id !== "string" || !c.id.startsWith(`${m.id}.`)) throw new Error(`classifier ids must start with "${m.id}."`);
        host.info.classifiers = [...host.info.classifiers.filter((x) => x.id !== c.id), { id: c.id, version: String(c.version) }];
        return null;
      }
      case "ui.showDialog":
      case "ui.showPanel":
      case "ui.openForm": {
        const kind = method === "ui.showDialog" ? "dialog" : method === "ui.showPanel" ? "panel" : "form";
        if (!perms.ui?.includes(kind)) throw new Error(`the "${kind}" UI permission is not declared in manifest.json`);
        return this.requestUi(kind, { id: m.id, name: m.name }, args as DialogSpec);
      }
      case "storage.get":
      case "storage.set":
      case "storage.delete":
      case "storage.keys":
        if (!perms.storage) throw new Error("the storage permission is not declared in manifest.json");
        return this.storage(m.id, method.slice("storage.".length), a);
      case "net.post": {
        const url = String(a.url);
        let origin = "";
        try {
          origin = new URL(url).origin;
        } catch {
          throw new Error(`invalid URL: ${url}`);
        }
        const declared = (perms.network ?? []).map((o) => o.replace(/\/$/, ""));
        const allowed = url.startsWith("https://") && declared.includes(origin);
        return this.post(m, url, a.body, { confirm: a.confirm === true, allowed });
      }
    }
    throw new Error(`unknown host API: ${method}`);
  }

  private storage(id: string, op: string, a: Record<string, unknown>): unknown {
    const file = path.join(this.storageDir, id, "storage.json");
    let data: Record<string, unknown> = {};
    try {
      data = JSON.parse(fs.readFileSync(file, "utf-8"));
    } catch {
      data = {};
    }
    if (op === "get") return data[String(a.key)];
    if (op === "keys") return Object.keys(data);
    if (op === "set") data[String(a.key)] = a.value;
    if (op === "delete") delete data[String(a.key)];
    const json = JSON.stringify(data);
    if (json.length > MAX_STORAGE_BYTES) throw new Error(`storage is limited to ${MAX_STORAGE_BYTES} bytes`);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, json, "utf-8");
    return null;
  }

  private active(): Host[] {
    return [...this.hosts.values()].filter((h) => h.info.state === "active");
  }

  /** 有効な拡張のカテゴリと分類器 */
  registrations(): { extId: string; categories: CategorySpec[]; classifiers: { id: string; version: string }[] }[] {
    return this.active().map((h) => ({ extId: h.info.id, categories: h.info.categories, classifiers: h.info.classifiers }));
  }

  /** 拡張の分類器を実行する。失敗したら空の結果とエラーを返す (アプリの解析は続ける) */
  async classify(
    extId: string,
    classifierId: string,
    ctx: Omit<ClassifierContext, "windowsFor">
  ): Promise<{ highlights: HighlightSpec[]; error?: string }> {
    const host = this.hosts.get(extId);
    if (!host || host.info.state !== "active") return { highlights: [], error: "the extension is not running" };
    try {
      const r = await this.invoke(host, "classify", { classifierId, ctx: this.redact(host, ctx) }, CLASSIFY_TIMEOUT_MS);
      return { highlights: Array.isArray(r) ? (r as HighlightSpec[]) : [] };
    } catch (err) {
      return { highlights: [], error: err instanceof Error ? err.message : String(err) };
    }
  }

  /** 解析の完了を、有効なすべての拡張に知らせる (待たない。失敗はその拡張のエラーとして記録する) */
  analysisComplete(result: AnalysisResult): void {
    for (const host of this.active()) {
      this.invoke(host, "analysisComplete", this.redact(host, result), ANALYSIS_TIMEOUT_MS).catch((err) => {
        host.info.error = err instanceof Error ? err.message : String(err);
      });
    }
  }

  /** documentText の権限が無い拡張には、本文を渡さない */
  private redact<T extends { positioned: PositionedEvent[]; finalText?: string }>(host: Host, v: T): T {
    if (host.info.manifest?.permissions?.documentText) return v;
    return {
      ...v,
      positioned: v.positioned.map((e) => ({ ...e, text: "" })),
      ...(v.finalText !== undefined ? { finalText: "" } : {}),
    };
  }

  /** 拡張が動き始めたら呼ぶ関数 (送信のキューを送るため) */
  onActivated?: (id: string) => void;

  stopAll(): void {
    for (const h of this.hosts.values()) if (h.win) this.stop(h);
  }
}

function withTimeout<T>(p: Promise<T>, ms: number, message: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), ms);
    p.then(
      (v) => {
        clearTimeout(timer);
        resolve(v);
      },
      (e) => {
        clearTimeout(timer);
        reject(e);
      }
    );
  });
}
