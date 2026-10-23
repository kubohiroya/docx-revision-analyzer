/**
 * PDF のページを画像にするための、非表示のウィンドウ (メインプロセス)。
 *
 * pdf.js は、アプリの画面とは別の非表示のウィンドウで動かす (拡張機能の実行環境と同じ作り)。
 *  - 独自のプロトコル dra-pdf:// だけを読める専用のセッション (メモリ上だけ)。ネットワークには出られない
 *  - dra-pdf://app/ は pdf.js とホストのページ (dist/pdf-host/)、dra-pdf://app/doc/<番号> はメインプロセスが許可した PDF
 *    (pdf.js は PDF を XHR で読むので、ホストのページと同じオリジンに置く)
 *  - 初めて使うときに作る (起動の処理は増やさない)
 */
import { BrowserWindow, session } from "electron";
import * as fs from "fs";
import * as path from "path";

export const PDF_SCHEME = "dra-pdf";

export interface RenderedPage {
  dataUrl: string;
  pageCount: number;
  /** ページの大きさ (pt) */
  width: number;
  height: number;
}

const TYPES: Record<string, string> = {
  ".html": "text/html",
  ".mjs": "text/javascript",
  ".js": "text/javascript",
  ".bcmap": "application/octet-stream",
  ".pdf": "application/pdf",
};

export class PdfHost {
  private win?: BrowserWindow;
  private starting?: Promise<BrowserWindow>;
  private readonly docs = new Map<string, string>();
  private nextDoc = 1;
  private sessionReady = false;

  constructor(private readonly hostDir: string) {}

  private session(): Electron.Session {
    const ses = session.fromPartition("pdf-host");
    if (!this.sessionReady) {
      this.sessionReady = true;
      ses.webRequest.onBeforeRequest((d, cb) => cb({ cancel: !d.url.startsWith(`${PDF_SCHEME}://`) && !d.url.startsWith("devtools:") }));
      ses.setPermissionRequestHandler((_wc, _p, cb) => cb(false));
      ses.protocol.handle(PDF_SCHEME, (req) => this.serve(req.url));
    }
    return ses;
  }

  private async serve(url: string): Promise<Response> {
    const u = new URL(url);
    let file: string | undefined;
    const doc = /^\/doc\/(\d+)$/.exec(u.pathname);
    if (u.hostname === "app" && doc) {
      file = this.docs.get(doc[1]);
    } else if (u.hostname === "app") {
      const rel = decodeURIComponent(u.pathname).replace(/^\/+/, "");
      const resolved = path.resolve(this.hostDir, rel);
      if (resolved.startsWith(path.resolve(this.hostDir) + path.sep)) file = resolved;
    }
    if (!file) return new Response("not found", { status: 404 });
    try {
      const body = await fs.promises.readFile(file);
      return new Response(body, {
        headers: {
          "content-type": TYPES[path.extname(file)] ?? "application/octet-stream",
          "content-security-policy":
            `default-src 'none'; script-src 'self'; worker-src 'self'; connect-src ${PDF_SCHEME}:; img-src data: blob:; ` +
            `font-src data: blob:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'`,
        },
      });
    } catch {
      return new Response("not found", { status: 404 });
    }
  }

  private window(): Promise<BrowserWindow> {
    if (this.win && !this.win.isDestroyed()) return Promise.resolve(this.win);
    this.starting ??= (async () => {
      const win = new BrowserWindow({
        show: false,
        webPreferences: { session: this.session(), sandbox: true, contextIsolation: true, nodeIntegration: false, webSecurity: true, spellcheck: false },
      });
      win.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
      win.webContents.on("will-navigate", (e) => e.preventDefault());
      await win.loadURL(`${PDF_SCHEME}://app/index.html`);
      for (let i = 0; i < 100 && !(await win.webContents.executeJavaScript("window.ready === true")); i++) {
        await new Promise((r) => setTimeout(r, 50));
      }
      this.win = win;
      this.starting = undefined;
      return win;
    })();
    return this.starting;
  }

  /** PDF のファイルを読めるようにして、その URL を返す */
  private docUrl(file: string): string {
    for (const [id, f] of this.docs) if (f === file) return `${PDF_SCHEME}://app/doc/${id}`;
    const id = String(this.nextDoc++);
    this.docs.set(id, file);
    return `${PDF_SCHEME}://app/doc/${id}`;
  }

  /** PDF のページ (0 から) を、maxWidth × maxHeight (画素) に収まる PNG にする */
  async render(file: string, pageIndex: number, maxWidth: number, maxHeight: number): Promise<RenderedPage> {
    const win = await this.window();
    const url = this.docUrl(file);
    try {
      return await win.webContents.executeJavaScript(
        `window.renderPage(${JSON.stringify(url)}, ${Math.max(0, Math.floor(pageIndex))}, ${Math.max(16, Math.floor(maxWidth))}, ${Math.max(16, Math.floor(maxHeight))})`
      );
    } catch (err) {
      // pdf.js の例外はただのオブジェクトで届くので、メッセージだけを取り出す
      const message = (err as { message?: unknown })?.message;
      throw new Error(typeof message === "string" ? message : String(err));
    }
  }

  /** ファイルが書き換わったときなどに、開いた PDF を手放す */
  async forget(file: string): Promise<void> {
    for (const [id, f] of this.docs) {
      if (f !== file) continue;
      this.docs.delete(id);
      if (this.win && !this.win.isDestroyed()) await this.win.webContents.executeJavaScript(`window.closeDoc(${JSON.stringify(`${PDF_SCHEME}://app/doc/${id}`)})`);
    }
  }

  dispose(): void {
    if (this.win && !this.win.isDestroyed()) this.win.destroy();
    this.win = undefined;
  }
}
