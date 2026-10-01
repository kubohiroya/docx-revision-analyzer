/**
 * デスクトップアプリのメインプロセス。
 *
 *  - ファイルの読み書き・ダイアログ・設定の保存・docx の設定の書き換えを受け持つ (IPC)。
 *  - 解析はすべてローカル (レンダラ) で行う。文書や本文をどこにも送らないよう、外部への通信を
 *    すべて遮断し、別ウィンドウの表示やページの移動も禁止する。
 *  - スモークテスト: 環境変数 DRA_SMOKE_FILE を指定すると、ウィンドウを表示せずにそのファイルを開き、
 *    各タブの画面を DRA_SMOKE_OUT (フォルダ) に PNG で保存して終了する。
 */

import { app, BrowserWindow, dialog, ipcMain, session, shell } from "electron";
import { autoUpdater } from "electron-updater";
import * as fs from "fs";
import * as path from "path";
import { parse as parseYaml } from "yaml";
import { enableHistoryPreservation } from "../../src/node/historyFile";
import { macSystemLocale } from "../../src/node/locale";
import { langFromLocale } from "../../src/lib/i18n";
import { AppLang, AppSettings, DEFAULT_SETTINGS, LoadedRules, OpenedFile } from "./shared";

const smokeFile = process.env.DRA_SMOKE_FILE;
const smokeOut = process.env.DRA_SMOKE_OUT;

let mainWindow: BrowserWindow | undefined;
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
  ipcMain.handle("system-lang", () => systemLang());
  ipcMain.handle("version", () => app.getVersion());
  ipcMain.on("renderer-ready", () => {
    const p = smokeFile ?? pendingOpenPath;
    pendingOpenPath = undefined;
    if (p) mainWindow?.webContents.send("open-path", path.resolve(p));
  });
  ipcMain.on("smoke-analyzed", (_e, summary: { ok: boolean; message: string }) => void runSmokeCapture(summary));
}

/** スモークテスト: 各タブを表示して PNG に保存し、終了する */
async function runSmokeCapture(summary: { ok: boolean; message: string }): Promise<void> {
  if (!smokeFile || !mainWindow) return;
  const out = smokeOut ?? process.cwd();
  fs.mkdirSync(out, { recursive: true });
  console.log(`smoke: ${summary.ok ? "ok" : "error"}: ${summary.message}`);
  for (const tab of ["chart", "flow", "highlights", "locate", "settings"]) {
    // locate: ハイライトの一覧の最初の行の「フロー」を押し、図の中の位置が示されることを確かめる
    const js =
      tab === "locate"
        ? `window.__showTab("highlights"); document.querySelector('tr.item button[data-goto="flow"]')?.click()`
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

app.whenReady().then(() => {
  lockDown();
  registerIpc();
  createWindow();
  if (loadSettings().checkForUpdates) checkForUpdates();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin" || smokeFile) app.quit();
});
