/**
 * 保存時のページ (PDF) の作成・保管・表示 (メインプロセス、issue #59)。
 *
 * 作り方:
 *  - OneDrive / SharePoint の文書の最新の版: Graph の format=pdf (Word for the web と同じ割り付け)
 *  - それ以外 (ローカルの文書、OneDrive の文書の過去の版): インストールされている LibreOffice で、保管した .docx を変換する
 *  - 横書きだけの文書では、埋め込みフォントを取り除く (表示するときは OS のフォントで代用する。設定で切り替え)
 * 保管する場所:
 *  - 監視している文書: 記録のフォルダ (pages.json と <名前>-sNN.pdf)。版の番号は保存ごとの版と同じ
 *  - 監視していない文書: 今の文書だけを、アプリのデータフォルダの pages/cache/ に (内容の SHA-256 ごと、新しいものから残す)
 * 変換する版 (自動保存は数秒ごとなので、保存のたびには変換しない):
 *  - 最新の版で、文書が閉じられた (ローカルの文書) か、最後の保存から無編集期間 (区間を分ける閾値) が過ぎたとき
 *  - 右の画面で、まだ PDF の無い版を開いたとき (その場で変換する)
 * 変換はいつも1つずつ、別のプロセス (LibreOffice) を優先度を下げて動かす。pdf.js は非表示のウィンドウで動かす (pdfHost.ts)。
 */
import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { createHash } from "crypto";
import * as fs from "fs";
import * as path from "path";
import { PDFDocument } from "pdf-lib";
import { parseDocxLayout } from "../../../src/lib/docxLayout";
import { isOpenInWord } from "../../../src/node/docWatcher";
import { snapshotFileName } from "../../../src/node/snapshotArchive";
import type { MicrosoftClient } from "../microsoft";
import { DOC_CHANGED, NO_LIBREOFFICE, type AppSettings, type PagesStatus, type PagesTarget, type SavedVersion } from "../shared";
import { currentWatch, onWatchVersion } from "../watchMain";
import { convertDocxToPdf, findSoffice } from "./libreoffice";
import { stripEmbeddedFonts } from "./pdfFonts";
import { PdfHost } from "./pdfHost";

interface PagesDeps {
  window: () => BrowserWindow | undefined;
  loadSettings: () => AppSettings;
  saveSettings: (s: AppSettings) => void;
  microsoft: () => MicrosoftClient | undefined;
}

interface IndexItem {
  file?: string;
  state: "done" | "failed";
  source?: "libreoffice" | "graph";
  createdAt: string;
  fontsStripped?: boolean;
  bytes?: number;
  pageCount?: number;
  error?: string;
}

interface PagesIndex {
  version: 1;
  items: Record<string, IndexItem>;
}

const INDEX = "pages.json";
const CACHE_KEEP = 10;
/** 自動で変換するかを確かめる間隔 */
const SCHEDULE_MS = Number(process.env.DRA_PAGES_SCHEDULE_MS ?? 30_000);

let host: PdfHost | undefined;

export function disposePages(): void {
  host?.dispose();
}

function readIndex(dir: string): PagesIndex {
  try {
    const v = JSON.parse(fs.readFileSync(path.join(dir, INDEX), "utf-8")) as PagesIndex;
    return v.items ? v : { version: 1, items: {} };
  } catch {
    return { version: 1, items: {} };
  }
}

function writeIndex(dir: string, idx: PagesIndex): void {
  const file = path.join(dir, INDEX);
  fs.writeFileSync(`${file}.tmp`, JSON.stringify(idx, null, 2) + "\n");
  fs.renameSync(`${file}.tmp`, file);
}

/** 縦書きの部分がある文書か (縦書きの字形は OS のフォントでは取り戻せないので、フォントを取り除かない) */
async function hasVerticalText(docx: Uint8Array): Promise<boolean> {
  try {
    const model = await parseDocxLayout(docx);
    return model.sections.some((s) => s.props.vertical || !!s.before?.vertical);
  } catch {
    return true;
  }
}

export function registerPagesIpc(d: PagesDeps): void {
  host = new PdfHost(path.join(__dirname, "pdf-host"));
  const workDir = path.join(app.getPath("userData"), "pages");
  const cacheDir = path.join(workDir, "cache");
  const pending = new Set<string>();
  const failures = new Map<string, string>();
  const send = (p: string) => d.window()?.webContents.send("pages:changed", p);

  const settings = () => d.loadSettings().pages ?? {};
  const status = (): PagesStatus => {
    const s = settings();
    return {
      enabled: s.enabled !== false,
      libreOffice: findSoffice(s.libreOfficePath),
      imageQuality: s.imageQuality ?? "medium",
      fonts: s.fonts ?? "substitute",
    };
  };

  /** 監視している文書か */
  const watched = (t: PagesTarget) => {
    const w = currentWatch();
    return w && w.watcher.path === (/^https?:/i.test(t.path) ? t.path : path.resolve(t.path)) ? w : undefined;
  };

  /** 版の .docx と、PDF を置くフォルダ・ファイル名 */
  const locate = async (t: PagesTarget, n: number) => {
    const w = watched(t);
    if (w) {
      const v = w.watcher.archive.versions().find((x) => x.n === n);
      if (!v) throw new Error(`version ${n} not found`);
      const versions = w.watcher.archive.versions();
      return {
        key: `${w.watcher.archive.dir}#${n}`,
        dir: w.watcher.archive.dir,
        pdfName: snapshotFileName(w.watcher.archive.key, n).replace(/\.docx$/i, ".pdf"),
        docx: () => fs.promises.readFile(path.join(w.watcher.archive.dir, v.file)).then((b) => new Uint8Array(b)),
        savedAt: v.sourceModified,
        cloud: w.target.kind === "cloud" && n === versions[versions.length - 1]?.n ? { driveId: w.target.driveId, itemId: w.target.itemId } : undefined,
      };
    }
    // 監視していない文書: 開いている文書だけ (n = 0)。内容の SHA-256 ごとに置く。文書を読むのは変換するときだけ
    // (OneDrive の文書をそのたびに落とさない)。開いた後に変わっていたら、作らない (別の内容の PDF を置かない)
    if (n !== 0) throw new Error("only the current document can be converted when it is not watched");
    const read = async () => {
      if (!t.cloud) return new Uint8Array(await fs.promises.readFile(t.path));
      const ms = d.microsoft();
      if (!ms) throw new Error("OneDrive / SharePoint is not configured");
      return (await ms.open(t.cloud, false)).bytes;
    };
    const sha = (b: Uint8Array) => createHash("sha256").update(b).digest("hex");
    let known = t.sha256 && t.mtime ? { sha256: t.sha256, mtime: t.mtime } : undefined;
    if (!known) {
      if (t.cloud) throw new Error("the document's content is not known");
      const bytes = await read();
      known = { sha256: sha(bytes), mtime: (await fs.promises.stat(t.path)).mtime.toISOString() };
    }
    const expected = known.sha256;
    return {
      key: `cache#${expected}`,
      dir: cacheDir,
      pdfName: `${expected}.pdf`,
      docx: async () => {
        const bytes = await read();
        if (sha(bytes) !== expected) throw new Error(DOC_CHANGED);
        return bytes;
      },
      savedAt: known.mtime,
      cloud: t.cloud,
    };
  };

  const versionOf = (dir: string, n: number, name: string, savedAt: string, key: string): SavedVersion => {
    const item = readIndex(dir).items[name];
    if (pending.has(key)) return { n, savedAt, state: "pending" };
    if (item?.state === "done" && item.file && fs.existsSync(path.join(dir, item.file))) {
      return { n, savedAt, state: "done", pageCount: item.pageCount, source: item.source, fontsStripped: item.fontsStripped, bytes: item.bytes };
    }
    if (item?.state === "failed" || failures.has(key)) return { n, savedAt, state: "failed", error: item?.error ?? failures.get(key) };
    return { n, savedAt, state: "none" };
  };

  const list = async (t: PagesTarget): Promise<SavedVersion[]> => {
    const w = watched(t);
    if (w) {
      const dir = w.watcher.archive.dir;
      return w.watcher.archive
        .versions()
        .map((v) => versionOf(dir, v.n, snapshotFileName(w.watcher.archive.key, v.n).replace(/\.docx$/i, ".pdf"), v.sourceModified, `${dir}#${v.n}`));
    }
    try {
      const loc = await locate(t, 0);
      return [versionOf(loc.dir, 0, loc.pdfName, loc.savedAt, loc.key)];
    } catch {
      return [];
    }
  };

  /** PDF を作って保管する */
  const convert = async (t: PagesTarget, n: number): Promise<void> => {
    const loc = await locate(t, n);
    if (pending.has(loc.key)) return;
    const st = status();
    pending.add(loc.key);
    failures.delete(loc.key);
    send(t.path);
    const idx = () => readIndex(loc.dir);
    try {
      await fs.promises.mkdir(loc.dir, { recursive: true });
      await fs.promises.mkdir(workDir, { recursive: true });
      const docx = await loc.docx();
      let pdf: Uint8Array | undefined;
      let source: "libreoffice" | "graph" = "libreoffice";
      const ms = d.microsoft();
      if (loc.cloud && ms?.configured()) {
        try {
          pdf = await ms.contentPdf(loc.cloud);
          source = "graph";
        } catch {
          pdf = undefined;
        }
      }
      if (!pdf) {
        if (!st.libreOffice) throw new Error(NO_LIBREOFFICE);
        pdf = await convertDocxToPdf(docx, { soffice: st.libreOffice, workDir, quality: st.imageQuality });
      }
      let fontsStripped = false;
      if (st.fonts === "substitute" && !(await hasVerticalText(docx))) {
        const r = await stripEmbeddedFonts(pdf);
        pdf = r.bytes;
        fontsStripped = r.stripped > 0;
      }
      const pageCount = (await PDFDocument.load(pdf, { updateMetadata: false })).getPageCount();
      const file = path.join(loc.dir, loc.pdfName);
      await host?.forget(file);
      await fs.promises.writeFile(`${file}.tmp`, pdf);
      await fs.promises.rename(`${file}.tmp`, file);
      const i = idx();
      i.items[loc.pdfName] = { file: loc.pdfName, state: "done", source, createdAt: new Date().toISOString(), fontsStripped, bytes: pdf.byteLength, pageCount };
      writeIndex(loc.dir, i);
      if (loc.dir === cacheDir) pruneCache();
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      failures.set(loc.key, message);
      // LibreOffice が無いだけなら記録しない (入れた後に、自動で作る対象に戻す)
      if (message !== NO_LIBREOFFICE) {
        try {
          const i = idx();
          i.items[loc.pdfName] = { state: "failed", createdAt: new Date().toISOString(), error: message };
          writeIndex(loc.dir, i);
        } catch {
          // 記録できなくても、状態は failures で返す
        }
      }
    } finally {
      pending.delete(loc.key);
      send(t.path);
    }
  };

  /** 監視していない文書の PDF は、新しいものから CACHE_KEEP 個だけ残す */
  const pruneCache = () => {
    const i = readIndex(cacheDir);
    const done = Object.entries(i.items)
      .filter(([, v]) => v.state === "done")
      .sort((a, b) => b[1].createdAt.localeCompare(a[1].createdAt));
    for (const [name] of done.slice(CACHE_KEEP)) {
      fs.rmSync(path.join(cacheDir, name), { force: true });
      delete i.items[name];
    }
    writeIndex(cacheDir, i);
  };

  // 監視している文書: 最新の版を、閉じられたか無編集期間が過ぎたら変換する
  let timer: NodeJS.Timeout | undefined;
  const check = async () => {
    const w = currentWatch();
    const st = status();
    if (!w || !st.enabled) return;
    const versions = w.watcher.archive.versions();
    const latest = versions[versions.length - 1];
    if (!latest) return;
    const t: PagesTarget = {
      path: w.watcher.path,
      cloud: w.target.kind === "cloud" ? { driveId: w.target.driveId, itemId: w.target.itemId } : undefined,
    };
    const v = (await list(t)).find((x) => x.n === latest.n);
    if (!v || !(v.state === "none" || (v.state === "failed" && v.error === NO_LIBREOFFICE))) return;
    if (!st.libreOffice && !(t.cloud && d.microsoft()?.configured())) return;
    const idleMs = Date.now() - Date.parse(latest.sourceModified);
    const gapMs = (d.loadSettings().analysis.gapThresholdHours ?? 1) * 3600_000;
    const closed = w.target.kind === "local" && !(await isOpenInWord(w.watcher.path));
    if (closed || idleMs >= gapMs) await convert(t, latest.n);
  };
  onWatchVersion(() => {
    if (!timer) timer = setInterval(() => void check().catch((err) => console.error(`pages: ${err}`)), SCHEDULE_MS);
  });

  ipcMain.handle("pages:status", () => status());
  ipcMain.handle("pages:list", (_e, t: PagesTarget) => list(t));
  ipcMain.handle("pages:request", async (_e, t: PagesTarget, n: number) => {
    const before = (await list(t)).find((v) => v.n === n);
    if (status().enabled && before && (before.state === "none" || before.state === "failed")) {
      void convert(t, n).catch((err) => console.error(`pages: ${err}`));
    }
    return (await list(t)).find((v) => v.n === n);
  });
  ipcMain.handle("pages:render", async (_e, t: PagesTarget, n: number, page: number, w: number, h: number) => {
    const loc = await locate(t, n);
    const file = path.join(loc.dir, loc.pdfName);
    if (!fs.existsSync(file)) throw new Error("the page is not ready");
    const r = await host!.render(file, page, w, h);
    return { dataUrl: r.dataUrl, pageCount: r.pageCount };
  });
  ipcMain.handle("pages:choose-libreoffice", async (_e, reset?: boolean) => {
    const s = d.loadSettings();
    if (reset) {
      s.pages = { ...(s.pages ?? {}), libreOfficePath: undefined };
    } else {
      const r = await dialog.showOpenDialog(d.window()!, {
        properties: ["openFile"],
        defaultPath: process.platform === "darwin" ? "/Applications" : undefined,
      });
      if (r.canceled || r.filePaths.length === 0) return status();
      let p = r.filePaths[0];
      // macOS: LibreOffice.app を選んだら、その中の soffice を使う
      if (/\.app$/i.test(p)) p = path.join(p, "Contents", "MacOS", "soffice");
      s.pages = { ...(s.pages ?? {}), libreOfficePath: p };
    }
    d.saveSettings(s);
    return status();
  });
}
