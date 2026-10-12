/** 書き出したファイル (SVG・一括処理の一覧) を既定のブラウザで開く (メインプロセス) */
import { app, shell } from "electron";
import { execFile } from "child_process";
import { pathToFileURL } from "url";

/**
 * ローカルのファイルを既定のブラウザで開く。ファイルの種類の既定のアプリ (SVG ならプレビューや画像編集ソフト) ではなく、
 * https を開くアプリで開く。それが分からない環境 (Linux など) では、ファイルの既定のアプリで開く
 */
export async function openFileInBrowser(file: string): Promise<void> {
  if (process.env.DRA_SMOKE_FILE || process.env.DRA_SMOKE_URL || process.env.DRA_SMOKE_START || process.env.DRA_SMOKE_BATCH) {
    console.log(`smoke: would open ${file} in the browser`);
    return;
  }
  try {
    if (process.platform === "darwin" || process.platform === "win32") {
      const browser = await app.getApplicationInfoForProtocol("https://");
      if (browser.path) {
        const [cmd, args] =
          process.platform === "darwin" ? ["open", ["-a", browser.path, file]] : [browser.path, [pathToFileURL(file).href]];
        await new Promise<void>((resolve, reject) => execFile(cmd, args, (err) => (err ? reject(err) : resolve())));
        return;
      }
    }
  } catch {
    // 下の方法で開く
  }
  const err = await shell.openPath(file);
  if (err) throw new Error(err);
}
