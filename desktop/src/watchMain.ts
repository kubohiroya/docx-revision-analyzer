/**
 * 文書の監視 (メインプロセス)。
 *
 *  - 開いている文書を監視し、保存されるたびに読み直してレンダラへ送る (図を描き直すため)
 *    - ローカルの文書 (OneDrive の同期フォルダの中を含む): フォルダの変更の通知と、更新日時の確認
 *    - URL から開いた OneDrive / SharePoint の文書: Graph に一定間隔で問い合わせ、中身の版 (cTag) が変わったら読み直す。
 *      問い合わせではサインインのためにブラウザを開かない。アプリが前面にないときは間隔を延ばす
 *  - 保存ごとの版とコメントの出来事を、記録のフォルダ (既定はアプリのデータフォルダの watch/) に積む
 *  - ローカルの文書が閉じられたことを検知したら、新しく見たコメントに「保全済み」の返信を書き込む
 *    (URL から開いた文書には書き込まない。読み取りだけの権限で開いているため)
 * 処理の本体は src/node/docWatcher.ts。監視している文書は設定に覚えておき、次に起動したら監視を再開する。
 */
import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import * as fs from "fs";
import * as path from "path";
import { DocWatcher, WatchSource, WatchUpdate } from "../../src/node/docWatcher";
import { setLang, t } from "../../src/lib/i18n";
import type { MicrosoftClient } from "./microsoft";
import { rememberSource } from "./batchMain";
import type { AppLang, AppSettings, OpenedFile, WatchStatus, WatchTarget, WatchUpdateMessage } from "./shared";

let watcher: DocWatcher | undefined;
let target: WatchTarget | undefined;
/** 新しい版を保管したとき・監視を始めたときに呼ぶ (保存時のページの作成の予定を立てるため) */
let versionListener: (() => void) | undefined;

/** 監視している文書と、その読み取り元 */
export function currentWatch(): { watcher: DocWatcher; target: WatchTarget } | undefined {
  return watcher?.archive && target ? { watcher, target } : undefined;
}

export function onWatchVersion(cb: () => void): void {
  versionListener = cb;
}

interface WatchDeps {
  window: () => BrowserWindow | undefined;
  microsoft: () => MicrosoftClient | undefined;
  loadSettings: () => AppSettings;
  saveSettings: (s: AppSettings) => void;
  lang: () => AppLang;
}

/** OneDrive / SharePoint の文書を問い合わせる間隔 (ミリ秒)。アプリが前面にないときはこの4倍 */
const CLOUD_POLL_MS = Number(process.env.DRA_WATCH_CLOUD_POLL_MS ?? 90_000);

function recordsRoot(s: AppSettings): string {
  return s.watchRecordsDir ?? path.join(app.getPath("userData"), "watch");
}

function fmtWhen(d: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function statusOf(s: AppSettings): WatchStatus {
  if (!watcher?.archive) {
    return { watching: false, canWrite: false, recordsRoot: recordsRoot(s), versions: 0, comments: [], pendingReceipts: 0, writtenReceipts: 0, gaps: 0 };
  }
  const a = watcher.archive;
  const q = a.receipts();
  return {
    watching: true,
    canWrite: watcher.source.canWrite,
    path: watcher.path,
    recordsRoot: a.root,
    dir: a.dir,
    versions: a.versions().length,
    comments: watcher.comments(),
    pendingReceipts: q.pending.length,
    writtenReceipts: q.written.length,
    gaps: a.info.gaps.length,
  };
}

function openedFile(t: WatchTarget, bytes: Uint8Array, mtime: Date): OpenedFile {
  if (t.kind === "local") return { path: t.path, name: path.basename(t.path), mtime: mtime.toISOString(), bytes };
  return {
    path: t.url,
    name: t.name,
    mtime: mtime.toISOString(),
    bytes,
    source: { kind: "url", url: t.url, webUrl: t.webUrl, driveId: t.driveId, itemId: t.itemId },
  };
}

/** OneDrive / SharePoint の文書の読み取り元: 中身の版 (cTag) を一定間隔で問い合わせる */
function cloudSource(ms: MicrosoftClient, t: Extract<WatchTarget, { kind: "cloud" }>, focused: () => boolean): WatchSource {
  const ref = { driveId: t.driveId, itemId: t.itemId };
  let lastTag: string | undefined;
  let first = true;
  return {
    location: t.url,
    name: t.name,
    canWrite: false,
    async read() {
      lastTag = await ms.contentTag(ref);
      // 監視を始めるとき (利用者の操作) だけ、必要ならサインインする
      const f = await ms.open(ref, first);
      first = false;
      return { bytes: f.bytes, mtime: new Date(f.mtime) };
    },
    subscribe(onChange, onError) {
      let last = Date.now();
      let lastError = "";
      const timer = setInterval(async () => {
        if (!focused() && Date.now() - last < CLOUD_POLL_MS * 4) return;
        last = Date.now();
        try {
          const tag = await ms.contentTag(ref);
          lastError = "";
          if (lastTag !== undefined && tag !== lastTag) onChange();
          lastTag = tag;
        } catch (err) {
          // 同じエラーは繰り返し知らせない (ネットワークが切れている間など)
          const msg = err instanceof Error ? err.message : String(err);
          if (msg !== lastError) onError(err);
          lastError = msg;
        }
      }, CLOUD_POLL_MS);
      return () => clearInterval(timer);
    },
  };
}

/** 監視をやめる (設定に覚えた監視している文書は消さない。アプリの終了時など) */
export function stopWatching(): void {
  watcher?.stop();
  watcher = undefined;
  target = undefined;
}

export function registerWatchIpc(d: WatchDeps): void {
  const send = (channel: string, payload: unknown) => d.window()?.webContents.send(channel, payload);
  const sendError = (err: unknown) => send("watch:error", err instanceof Error ? err.message : String(err));
  const message = (t: WatchTarget, u: WatchUpdate): WatchUpdateMessage => ({
    kind: u.kind,
    file: openedFile(t, u.bytes, u.mtime),
    status: statusOf(d.loadSettings()),
    events: u.events,
    version: u.version?.n,
    gap: u.gap,
  });

  /** 文書の監視を始める (ほかの文書の監視はやめる) */
  async function start(wt: WatchTarget): Promise<WatchUpdateMessage> {
    stopWatching();
    const s = d.loadSettings();
    let source: WatchSource | undefined;
    if (wt.kind === "cloud") {
      const ms = d.microsoft();
      if (!ms?.configured()) throw new Error("OneDrive / SharePoint is not configured");
      // 元の文書へのリンク (図の見出し・コメントの印) を、確認なしで開いてよいものとして覚える (起動時に監視を再開したときも)
      rememberSource(wt.webUrl ?? wt.url);
      source = cloudSource(ms, wt, () => !!d.window()?.isFocused());
    }
    const w = new DocWatcher({
      path: wt.kind === "local" ? path.resolve(wt.path) : undefined,
      source,
      recordsRoot: recordsRoot(s),
      receiptText: (seq, at) => {
        setLang(d.lang());
        return t("commentReceipt", seq, fmtWhen(at));
      },
      onUpdate: (u) => {
        send("watch:update", message(wt, u));
        if (u.version) versionListener?.();
      },
      onReceipts: async (r) => {
        try {
          const bytes = new Uint8Array(await fs.promises.readFile(r.path));
          send("watch:receipts", {
            file: openedFile(wt, bytes, (await fs.promises.stat(r.path)).mtime),
            status: statusOf(d.loadSettings()),
            written: r.written.length,
            skipped: r.skipped.length,
          });
        } catch (err) {
          sendError(err);
        }
      },
      onError: sendError,
    });
    const first = await w.start();
    watcher = w;
    target = wt.kind === "local" ? { kind: "local", path: path.resolve(wt.path) } : wt;
    s.watchTarget = target;
    d.saveSettings(s);
    versionListener?.();
    return message(target, first);
  }

  ipcMain.handle("watch:start", (_e, wt: WatchTarget) => start(wt));

  ipcMain.handle("watch:stop", () => {
    stopWatching();
    const s = d.loadSettings();
    s.watchTarget = undefined;
    d.saveSettings(s);
    return statusOf(s);
  });

  ipcMain.handle("watch:status", () => statusOf(d.loadSettings()));

  ipcMain.handle("watch:choose-records-dir", async (_e, reset?: boolean) => {
    const s = d.loadSettings();
    if (reset) {
      s.watchRecordsDir = undefined;
    } else {
      const r = await dialog.showOpenDialog(d.window()!, {
        properties: ["openDirectory", "createDirectory"],
        defaultPath: recordsRoot(s),
      });
      if (r.canceled || r.filePaths.length === 0) return statusOf(s);
      s.watchRecordsDir = r.filePaths[0];
    }
    d.saveSettings(s);
    // 監視中なら、新しい保管先で監視し直す
    if (watcher && target) await start(target);
    return statusOf(d.loadSettings());
  });

  ipcMain.handle("watch:open-records", async () => {
    const dir = watcher?.archive?.dir ?? recordsRoot(d.loadSettings());
    fs.mkdirSync(dir, { recursive: true });
    await shell.openPath(dir);
  });
}
