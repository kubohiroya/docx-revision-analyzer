/**
 * デスクトップアプリのメインプロセス。
 *
 *  - ファイルの読み書き・ダイアログ・設定の保存・docx の設定の書き換えを受け持つ (IPC)。
 *  - 解析はすべてローカル (レンダラ) で行う。文書や本文をどこにも送らないよう、外部への通信を
 *    すべて遮断し、別ウィンドウの表示やページの移動も禁止する。
 *  - スモークテスト: 環境変数 DRA_SMOKE_FILE を指定すると、ウィンドウを表示せずにそのファイルを開き、
 *    各タブの画面を DRA_SMOKE_OUT (フォルダ) に PNG で保存して終了する。
 */

import { app, BrowserWindow, dialog, ipcMain, protocol, session, shell } from "electron";
import { autoUpdater } from "electron-updater";
import * as fs from "fs";
import * as path from "path";
import { parse as parseYaml } from "yaml";
import { enableHistoryPreservation } from "../../src/node/historyFile";
import { macSystemLocale } from "../../src/node/locale";
import { langFromLocale } from "../../src/lib/i18n";
import { AppLang, AppSettings, DEFAULT_SETTINGS, ExtensionListItem, LoadedRules, MicrosoftStatus, OpenedFile, RecentUrl } from "./shared";
import { MicrosoftClient, MicrosoftConfig, MicrosoftError } from "./microsoft";
import { EXT_SCHEME, ExtensionManager } from "./extensions";
import { NetMediator } from "./sendLog";
import { registerBatchIpc } from "./batchMain";
import type { AnalysisResult, ClassifierContext } from "./extension-api";

// 拡張のページ (dra-ext://<id>/) を、ES モジュールを読み込める安全なオリジンとして扱う
protocol.registerSchemesAsPrivileged([{ scheme: EXT_SCHEME, privileges: { standard: true, secure: true, supportFetchAPI: true } }]);

const smokeFile = process.env.DRA_SMOKE_FILE ?? process.env.DRA_SMOKE_URL ?? process.env.DRA_SMOKE_START ?? process.env.DRA_SMOKE_BATCH;
// テスト用: 設定・拡張のデータの置き場所を変える (利用者のデータに触れないため)
if (process.env.DRA_USER_DATA) app.setPath("userData", path.resolve(process.env.DRA_USER_DATA));
const smokeOut = process.env.DRA_SMOKE_OUT;

let mainWindow: BrowserWindow | undefined;
let extensions: ExtensionManager | undefined;
let mediator: NetMediator | undefined;
let microsoft: MicrosoftClient | undefined;
const MAX_RECENT_URLS = 20;

/**
 * Microsoft のアプリ (クライアント) ID: 環境変数 DRA_MS_CLIENT_ID > 設定 > 配布物に同梱した microsoft.json の順。
 * 登録の手順は DISTRIBUTION.md
 */
function microsoftConfig(): MicrosoftConfig | undefined {
  const s = loadSettings().microsoft;
  let bundled: MicrosoftConfig | undefined;
  try {
    bundled = JSON.parse(fs.readFileSync(path.join(app.getAppPath(), "microsoft.json"), "utf-8"));
  } catch {
    bundled = undefined;
  }
  const clientId = process.env.DRA_MS_CLIENT_ID || s?.clientId || bundled?.clientId;
  if (!clientId) return undefined;
  return { clientId, authority: process.env.DRA_MS_AUTHORITY || s?.authority || bundled?.authority };
}

async function microsoftStatus(): Promise<MicrosoftStatus> {
  if (!microsoft?.configured()) return { configured: false };
  try {
    const a = await microsoft.account();
    return { configured: true, account: a ? a.name || a.username : undefined };
  } catch {
    return { configured: true };
  }
}

function rememberUrl(entry: RecentUrl): RecentUrl[] {
  const s = loadSettings();
  const list = [entry, ...(s.recentUrls ?? []).filter((r) => r.url !== entry.url)].slice(0, MAX_RECENT_URLS);
  s.recentUrls = list;
  saveSettings(s);
  return list;
}
/** スモークテストで有効にする拡張 (設定には保存しない) */
const smokeExtension = process.env.DRA_SMOKE_EXTENSION;
let nextUiRequest = 1;
/** 起動時の拡張の起動が終わったら解決する (ファイルを開くのはその後) */
let extensionsReady: Promise<void> = Promise.resolve();
const uiRequests = new Map<number, (value: unknown) => void>();
/** 起動前に Finder から渡されたファイル (macOS の open-file) */
let pendingOpenPath: string | undefined;

function settingsPath(): string {
  return path.join(app.getPath("userData"), "settings.json");
}

function loadSettings(): AppSettings {
  try {
    const raw = JSON.parse(fs.readFileSync(settingsPath(), "utf-8"));
    return {
      ...DEFAULT_SETTINGS,
      ...raw,
      extensions: { ...DEFAULT_SETTINGS.extensions, ...(raw.extensions ?? {}) },
      analysis: { ...DEFAULT_SETTINGS.analysis, ...(raw.analysis ?? {}) },
    };
  } catch {
    return structuredClone(DEFAULT_SETTINGS);
  }
}

function saveSettings(s: AppSettings): void {
  fs.mkdirSync(path.dirname(settingsPath()), { recursive: true });
  fs.writeFileSync(settingsPath(), JSON.stringify(s, null, 2), "utf-8");
}

async function readDocx(p: string): Promise<OpenedFile> {
  const [bytes, stat] = await Promise.all([fs.promises.readFile(p), fs.promises.stat(p)]);
  return { path: p, name: path.basename(p), mtime: stat.mtime.toISOString(), bytes: new Uint8Array(bytes) };
}

async function readRules(p: string): Promise<LoadedRules> {
  return { path: p, value: parseYaml(await fs.promises.readFile(p, "utf-8")) };
}

function systemLang(): AppLang {
  return langFromLocale(macSystemLocale() ?? app.getLocale()) ?? "en";
}

/** 外部への通信・ページの移動・新しいウィンドウを禁止する */
function lockDown(): void {
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    const ok = details.url.startsWith("file:") || details.url.startsWith("devtools:") || details.url.startsWith("blob:") || details.url.startsWith("data:");
    callback({ cancel: !ok });
  });
  session.defaultSession.setPermissionRequestHandler((_wc, _perm, callback) => callback(false));
  app.on("web-contents-created", (_e, contents) => {
    contents.setWindowOpenHandler(({ url }) => {
      // README などのリンクは既定のブラウザで開く (アプリ内には読み込まない)
      if (/^https:\/\/github\.com\//.test(url)) void shell.openExternal(url);
      return { action: "deny" };
    });
    contents.on("will-navigate", (e) => e.preventDefault());
  });
}

function registerIpc(): void {
  ipcMain.handle("open-docx-dialog", async () => {
    const r = await dialog.showOpenDialog(mainWindow!, {
      properties: ["openFile"],
      filters: [{ name: "Word", extensions: ["docx"] }],
    });
    return r.canceled || r.filePaths.length === 0 ? null : readDocx(r.filePaths[0]);
  });
  ipcMain.handle("read-docx", (_e, p: string) => readDocx(p));
  ipcMain.handle("open-url", async (_e, target: { url: string } | { recent: string }) => {
    if (!microsoft) return { error: { code: "notConfigured", message: "" } };
    const recent = "recent" in target ? (loadSettings().recentUrls ?? []).find((r) => r.url === target.recent) : undefined;
    const url = "url" in target ? target.url.trim() : target.recent;
    // 最近開いたフォルダは、一括処理として開き直す
    if (recent?.kind === "folder") return { error: { code: "folder", message: recent.name } };
    try {
      // 最近開いた文書は、前回の driveId / itemId で開く (共有リンクの解決を省く)。失敗したら URL から開き直す
      let f;
      try {
        f = recent?.driveId && recent.itemId ? await microsoft.open({ driveId: recent.driveId, itemId: recent.itemId }) : undefined;
      } catch {
        f = undefined;
      }
      f ??= await microsoft.open({ url });
      rememberUrl({ url, name: f.name, driveId: f.driveId, itemId: f.itemId, openedAt: new Date().toISOString() });
      const file: OpenedFile = { path: url, name: f.name, mtime: f.mtime, bytes: f.bytes, source: { kind: "url", url } };
      return { file };
    } catch (err) {
      const code = err instanceof MicrosoftError ? err.code : "network";
      return { error: { code, message: err instanceof Error ? err.message : String(err) } };
    }
  });
  ipcMain.handle("recent-urls", () => loadSettings().recentUrls ?? []);
  ipcMain.handle("remove-recent-url", (_e, url: string) => {
    const s = loadSettings();
    s.recentUrls = (s.recentUrls ?? []).filter((r) => r.url !== url);
    saveSettings(s);
    return s.recentUrls;
  });
  ipcMain.handle("ms-status", () => microsoftStatus());
  ipcMain.handle("ms-sign-out", async () => {
    await microsoft?.signOut();
    return microsoftStatus();
  });
  ipcMain.handle("preserve-history", async (_e, p: string) => {
    const r = await enableHistoryPreservation(p);
    return { backupPath: r.backupPath };
  });
  ipcMain.handle("choose-rules-dialog", async () => {
    const r = await dialog.showOpenDialog(mainWindow!, {
      properties: ["openFile"],
      filters: [{ name: "Rules (YAML / JSON)", extensions: ["yml", "yaml", "json"] }],
    });
    return r.canceled || r.filePaths.length === 0 ? null : readRules(r.filePaths[0]);
  });
  ipcMain.handle("read-rules", (_e, p: string) => readRules(p));
  ipcMain.handle("save-file", async (_e, defaultName: string, data: Uint8Array | string, kind: "svg" | "png") => {
    const r = await dialog.showSaveDialog(mainWindow!, {
      defaultPath: defaultName,
      filters: [kind === "svg" ? { name: "SVG", extensions: ["svg"] } : { name: "PNG", extensions: ["png"] }],
    });
    if (r.canceled || !r.filePath) return null;
    await fs.promises.writeFile(r.filePath, typeof data === "string" ? data : Buffer.from(data));
    return r.filePath;
  });
  ipcMain.handle("get-settings", () => loadSettings());
  ipcMain.handle("set-settings", (_e, s: AppSettings) => {
    const turnedOn = s.checkForUpdates && !loadSettings().checkForUpdates;
    saveSettings(s);
    if (turnedOn) checkForUpdates();
  });
  ipcMain.handle("updates-available", () => app.isPackaged);

  // 拡張
  ipcMain.handle("ext:list", () => extensionList());
  ipcMain.handle("ext:set-enabled", async (_e, id: string, enabled: boolean) => {
    const s = loadSettings();
    s.extensions[id] = enabled;
    saveSettings(s);
    // 無効にした拡張の、まだ送っていない送信は破棄する
    if (!enabled) mediator?.discard(id);
    await syncExtensions();
    return extensionList();
  });
  ipcMain.handle("ext:send-log", (_e, id: string) => mediator?.log(id) ?? []);
  ipcMain.handle("ext:registrations", () => extensions?.registrations() ?? []);
  ipcMain.handle("ext:classify", (_e, extId: string, classifierId: string, ctx: Omit<ClassifierContext, "windowsFor">) =>
    extensions ? extensions.classify(extId, classifierId, ctx) : { highlights: [], error: "extensions are not available" }
  );
  ipcMain.on("ext:analysis-complete", (_e, result: AnalysisResult) => extensions?.analysisComplete(result));
  ipcMain.on("ext-ui-reply", (_e, msg: { reqId: number; value: unknown }) => {
    uiRequests.get(msg.reqId)?.(msg.value);
    uiRequests.delete(msg.reqId);
  });
  ipcMain.handle("system-lang", () => systemLang());
  ipcMain.handle("version", () => app.getVersion());
  ipcMain.on("renderer-ready", () => {
    // スモークテスト: フォルダの一括処理 (DRA_SMOKE_BATCH にローカルのフォルダ、または SharePoint の URL)
    if (process.env.DRA_SMOKE_BATCH) {
      const b = process.env.DRA_SMOKE_BATCH;
      const target = /^https:/.test(b) ? { kind: "cloud", url: b } : { kind: "local", path: path.resolve(b) };
      void extensionsReady.then(() => mainWindow?.webContents.send("smoke-batch", target));
      return;
    }
    // スモークテスト: 起動画面 (ファイルを開く前) だけを撮る
    if (process.env.DRA_SMOKE_START) {
      setTimeout(() => void captureStart(), 800);
      return;
    }
    const p = process.env.DRA_SMOKE_FILE ?? pendingOpenPath;
    pendingOpenPath = undefined;
    if (p) void extensionsReady.then(() => mainWindow?.webContents.send("open-path", path.resolve(p)));
    else if (process.env.DRA_SMOKE_URL) void extensionsReady.then(() => mainWindow?.webContents.send("open-url", process.env.DRA_SMOKE_URL));
  });
  ipcMain.on("smoke-batch-done", () => void captureStart("smoke-batch.png"));
  ipcMain.on("smoke-analyzed", (_e, summary: { ok: boolean; message: string }) => void runSmokeCapture(summary));
}

async function captureStart(name = "smoke-start.png"): Promise<void> {
  if (!mainWindow) return;
  const out = smokeOut ?? process.cwd();
  fs.mkdirSync(out, { recursive: true });
  const file = path.join(out, name);
  fs.writeFileSync(file, (await mainWindow.webContents.capturePage()).toPNG());
  console.log(`smoke: wrote ${file}`);
  app.exit(0);
}

/** スモークテスト: 各タブを表示して PNG に保存し、終了する */
async function runSmokeCapture(summary: { ok: boolean; message: string }): Promise<void> {
  if (!smokeFile || !mainWindow) return;
  const out = smokeOut ?? process.cwd();
  fs.mkdirSync(out, { recursive: true });
  console.log(`smoke: ${summary.ok ? "ok" : "error"}: ${summary.message}`);
  // 拡張のパネルなど、解析の後から届く表示を待つ (DRA_SMOKE_WAIT_MS)
  await new Promise((r) => setTimeout(r, Number(process.env.DRA_SMOKE_WAIT_MS ?? 0)));
  for (const tab of ["chart", "flow", "highlights", "locate", "settings", "url"]) {
    // locate: ハイライトの一覧の最初の行の「フロー」を押し、図の中の位置が示されることを確かめる
    const js =
      tab === "locate"
        ? `window.__showTab("highlights"); document.querySelector('tr.item button[data-goto="flow"]')?.click()`
        : tab === "url"
          ? `window.__showTab("chart"); window.__showUrlDialog()`
          : `window.__showTab(${JSON.stringify(tab)})`;
    await mainWindow.webContents.executeJavaScript(js);
    await new Promise((r) => setTimeout(r, 400));
    const img = await mainWindow.webContents.capturePage();
    const file = path.join(out, `smoke-${tab}.png`);
    fs.writeFileSync(file, img.toPNG());
    console.log(`smoke: wrote ${file}`);
  }
  app.exit(summary.ok ? 0 : 1);
}

function extensionList(): ExtensionListItem[] {
  const enabled = loadSettings().extensions;
  return (extensions?.list() ?? []).map((x) => ({
    id: x.id,
    manifest: x.manifest,
    state: x.state,
    error: x.error,
    enabled: !!enabled[x.id] || x.id === smokeExtension,
  }));
}

async function syncExtensions(): Promise<void> {
  const enabled = { ...loadSettings().extensions };
  if (smokeExtension) enabled[smokeExtension] = true;
  await extensions?.sync(enabled);
}

/** 拡張の UI をアプリのウィンドウで表示し、結果を待つ */
function requestUi(kind: string, ext: { id: string; name: unknown }, spec: unknown): Promise<unknown> {
  if (kind === "confirmSend" && process.env.DRA_SMOKE_CONFIRM) {
    // スモークテスト: 確認に自動で答える (accept / reject)
    return Promise.resolve(process.env.DRA_SMOKE_CONFIRM === "accept");
  }
  if (!mainWindow) return Promise.resolve(kind === "form" ? null : kind === "confirmSend" ? false : { buttonId: null });
  const reqId = nextUiRequest++;
  return new Promise((resolve) => {
    uiRequests.set(reqId, resolve);
    mainWindow!.webContents.send("ext-ui", { reqId, kind, ext, spec });
    if (kind === "panel") {
      uiRequests.delete(reqId);
      resolve(null);
    }
  });
}

/**
 * GitHub Releases で新しいバージョンを確認し、あればダウンロードして終了時に入れ替える (利用者が設定で
 * 有効にした場合だけ)。electron-updater はメインプロセスの専用セッションで通信し、更新情報 (latest.yml) の
 * SHA-512 と、macOS ではアプリの署名 (Squirrel.Mac)、Windows では発行者名 (publisherName) を検証する。
 * 文書や解析結果は送らない。開発中 (パッケージしていない起動) は何もしない。
 */
function checkForUpdates(): void {
  if (!app.isPackaged || smokeFile) return;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.checkForUpdatesAndNotify().catch((err) => console.error(`update check failed: ${err}`));
}

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    show: !smokeFile,
    title: "Docx Revision Analyzer",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  void mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"));
  mainWindow.on("closed", () => (mainWindow = undefined));
}

// macOS: Finder でアプリのアイコンへドロップされたファイル
app.on("open-file", (e, p) => {
  e.preventDefault();
  if (mainWindow) mainWindow.webContents.send("open-path", p);
  else pendingOpenPath = p;
});

/**
 * 起動時の引数で渡されたファイル・フォルダ (Windows でアイコンへドロップしたとき・「送る」など)。
 * 開発中の起動 (electron .) の引数は対象にしない
 */
function pathFromArgv(argv: string[]): string | undefined {
  if (!app.isPackaged) return undefined;
  return argv
    .slice(1)
    .filter((a) => !a.startsWith("-"))
    .find((a) => {
      try {
        const st = fs.statSync(a);
        return st.isDirectory() || /\.docx$/i.test(a);
      } catch {
        return false;
      }
    });
}

// 2つ目の起動 (すでに開いているアプリのアイコンへドロップしたなど) は、開いているウィンドウで受け取る
if (!process.env.DRA_SMOKE_FILE && !process.env.DRA_SMOKE_URL && !process.env.DRA_SMOKE_START && !process.env.DRA_SMOKE_BATCH) {
  if (!app.requestSingleInstanceLock()) app.exit(0);
  app.on("second-instance", (_e, argv) => {
    const p = pathFromArgv(argv);
    if (p && mainWindow) mainWindow.webContents.send("open-path", path.resolve(p));
    if (mainWindow?.isMinimized()) mainWindow.restore();
    mainWindow?.focus();
  });
}

app.whenReady().then(() => {
  lockDown();
  registerIpc();
  registerBatchIpc(
    () => mainWindow,
    () => microsoft,
    (e) => rememberUrl({ ...e, openedAt: new Date().toISOString() })
  );
  pendingOpenPath ??= pathFromArgv(process.argv);
  mediator = new NetMediator({
    dir: path.join(app.getPath("userData"), "extensions"),
    confirmByDefault: (id) => loadSettings().extensionConfirmSends?.[id] !== false,
    askUser: async (ext, url, body) => (await requestUi("confirmSend", { id: ext.id, name: ext.name }, { url, body })) === true,
    onChange: (id) => mainWindow?.webContents.send("ext-send-log-changed", id),
  });
  microsoft = new MicrosoftClient(microsoftConfig, path.join(app.getPath("userData"), "msal-cache.bin"));
  extensions = new ExtensionManager(
    path.join(app.getAppPath(), "extensions"),
    path.join(__dirname, "ext-host"),
    path.join(__dirname, "ext-preload.js"),
    path.join(app.getPath("userData"), "extensions"),
    () => ({ version: app.getVersion(), locale: loadSettings().lang ?? systemLang() }),
    requestUi,
    (ext, url, body, opts) => mediator!.post(ext, url, body, opts)
  );
  extensions.onActivated = (id) => void mediator?.flush(id);
  extensions.discover();
  createWindow();
  extensionsReady = syncExtensions().catch((err) => console.error(`extensions: ${err}`));
  if (loadSettings().checkForUpdates) checkForUpdates();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("before-quit", () => extensions?.stopAll());

app.on("window-all-closed", () => {
  if (process.platform !== "darwin" || smokeFile) app.quit();
});
