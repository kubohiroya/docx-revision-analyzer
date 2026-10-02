/**
 * フォルダの一括処理 (メインプロセス側)。
 *
 *  - ローカルのフォルダ (ドラッグ&ドロップ・アイコンへのドロップ・起動時の引数) の下の .docx を一覧する
 *  - SharePoint / OneDrive のフォルダの URL の下の .docx を一覧し、1つずつ読み込む (microsoft.ts)
 *  - 出力先のフォルダを選ぶ (新しいフォルダも作れるダイアログ)
 *  - 結果のファイルを書き込む。書き込めるのは、利用者が選んだ・ドロップしたフォルダ (出力先として登録したもの) の中の
 *    .svg / .csv / .html だけ (レンダラから任意のファイルを書き換えられないようにする)
 *
 * 解析と描画はレンダラ (ライブラリのコア) で行う。
 */

import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import * as fs from "fs";
import * as path from "path";
import { isTargetDocx, MAX_FOLDER_FILES, MicrosoftClient, MicrosoftError } from "./microsoft";
import type { BatchFolder } from "./shared";

/**
 * 開いた (一括処理した) OneDrive / SharePoint の文書の URL。図の見出しや、拡張の注釈の「元の文書へのリンク」として、
 * 確認なしで開いてよいもの (アプリが Graph から得た URL で、拡張が作ったものではない)
 */
const sourceUrls = new Set<string>();

/** 文書の URL から、該当箇所の指定 (#ブックマーク、Word for the web の見出しリンクの nav=) を除いたもの */
function baseOf(url: string): string {
  return url
    .replace(/#.*$/, "")
    .replace(/([?&])nav=[^&]*&?/, "$1")
    .replace(/[?&]$/, "");
}

export function rememberSource(webUrl: string): void {
  sourceUrls.add(baseOf(webUrl));
}

/** 開いた文書の URL (と、その #ブックマーク・見出しリンク) か */
export function isSourceLink(url: string): boolean {
  return /^https:\/\//i.test(url) && sourceUrls.has(baseOf(url));
}

/** 書き込みを許した出力先のフォルダ */
const outputRoots = new Set<string>();
const ALLOWED_EXT = new Set([".svg", ".csv", ".html"]);

/** ローカルのフォルダの下の .docx を一覧する (隠しフォルダ・一時ファイルは除く) */
export async function listLocalFolder(dir: string): Promise<BatchFolder> {
  const files: BatchFolder["files"] = [];
  let truncated = false;
  const walk = async (rel: string): Promise<void> => {
    const entries = await fs.promises.readdir(path.join(dir, rel), { withFileTypes: true });
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const e of entries) {
      if (truncated) return;
      if (e.name.startsWith(".")) continue;
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) await walk(r);
      else if (e.isFile() && isTargetDocx(e.name)) {
        if (files.length >= MAX_FOLDER_FILES) {
          truncated = true;
          return;
        }
        const st = await fs.promises.stat(path.join(dir, r));
        files.push({ path: r, name: e.name, size: st.size, mtime: st.mtime.toISOString(), ref: { kind: "local", path: path.join(dir, r) } });
      }
    }
  };
  await walk("");
  return { name: path.basename(dir), location: dir, files, truncated };
}

function isInside(root: string, p: string): boolean {
  const rel = path.relative(root, p);
  return rel !== "" && !rel.startsWith("..") && !path.isAbsolute(rel);
}

export function registerBatchIpc(
  window: () => BrowserWindow | undefined,
  microsoft: () => MicrosoftClient | undefined,
  remember: (entry: { url: string; name: string; kind: "folder" }) => void
): void {
  ipcMain.handle("inspect-path", async (_e, p: string) => {
    try {
      const st = await fs.promises.stat(p);
      return st.isDirectory() ? "directory" : st.isFile() ? "file" : "other";
    } catch {
      return "missing";
    }
  });

  ipcMain.handle("list-local-folder", async (_e, dir: string) => {
    const folder = await listLocalFolder(dir);
    // ドロップしたフォルダには、その中に結果を書き込む
    outputRoots.add(path.resolve(dir));
    return folder;
  });

  ipcMain.handle("list-cloud-folder", async (_e, url: string) => {
    const ms = microsoft();
    if (!ms) return { error: { code: "notConfigured", message: "" } };
    try {
      const f = await ms.listFolder(url);
      remember({ url, name: f.name, kind: "folder" });
      for (const x of f.files) if (x.webUrl) rememberSource(x.webUrl);
      const folder: BatchFolder = {
        name: f.name,
        location: f.webUrl ?? url,
        truncated: f.truncated,
        files: f.files.map((x) => ({
          path: x.path,
          name: x.name,
          size: x.size,
          mtime: x.mtime,
          ref: { kind: "cloud", driveId: x.driveId, itemId: x.itemId },
          webUrl: x.webUrl,
        })),
      };
      return { folder };
    } catch (err) {
      return {
        error: { code: err instanceof MicrosoftError ? err.code : "network", message: err instanceof Error ? err.message : String(err) },
      };
    }
  });

  ipcMain.handle("read-batch-file", async (_e, ref: BatchFolder["files"][number]["ref"]) => {
    if (ref.kind === "local") {
      // 一覧したフォルダ (出力先として登録したもの) の中の .docx だけを読む
      const p = path.resolve(ref.path);
      if (![...outputRoots].some((r) => isInside(r, p)) || !isTargetDocx(path.basename(p))) throw new Error("not allowed");
      return new Uint8Array(await fs.promises.readFile(p));
    }
    const ms = microsoft();
    if (!ms) throw new Error("not configured");
    return (await ms.open({ driveId: ref.driveId, itemId: ref.itemId })).bytes;
  });

  ipcMain.handle("choose-output-dir", async (_e, suggestedName: string) => {
    if (process.env.DRA_SMOKE_OUTPUT_DIR) {
      const d = path.resolve(process.env.DRA_SMOKE_OUTPUT_DIR);
      fs.mkdirSync(d, { recursive: true });
      outputRoots.add(d);
      return d;
    }
    const win = window();
    const opts: Electron.OpenDialogOptions = {
      title: suggestedName,
      buttonLabel: app.getLocale().startsWith("ja") ? "ここに保存" : "Save here",
      defaultPath: path.join(app.getPath("documents"), suggestedName.replace(/[\\/:*?"<>|]/g, "_")),
      properties: ["openDirectory", "createDirectory", "promptToCreate"],
    };
    const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
    if (r.canceled || !r.filePaths[0]) return null;
    const d = path.resolve(r.filePaths[0]);
    // 「作成する」を選んだ場合 (promptToCreate) はまだ無いので作る
    fs.mkdirSync(d, { recursive: true });
    outputRoots.add(d);
    return d;
  });

  ipcMain.handle("write-output", async (_e, root: string, rel: string, content: string) => {
    const r = path.resolve(root);
    const p = path.resolve(r, rel);
    if (!outputRoots.has(r) || !isInside(r, p) || !ALLOWED_EXT.has(path.extname(p).toLowerCase())) {
      throw new Error(`writing ${rel} is not allowed`);
    }
    await fs.promises.mkdir(path.dirname(p), { recursive: true });
    await fs.promises.writeFile(p, content, "utf-8");
    return p;
  });

  ipcMain.handle("open-source-link", async (_e, url: string) => {
    if (!isSourceLink(url)) return false;
    if (process.env.DRA_SMOKE_FILE || process.env.DRA_SMOKE_URL || process.env.DRA_SMOKE_BATCH) console.log(`smoke: would open ${url}`);
    else await shell.openExternal(url);
    return true;
  });

  ipcMain.handle("show-folder", (_e, dir: string) => {
    if (outputRoots.has(path.resolve(dir))) void shell.openPath(dir);
  });
}
