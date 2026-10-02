/**
 * OneDrive / SharePoint 上の docx を URL から開く (メインプロセス)。
 *
 *  - Microsoft へのサインインは MSAL (認可コード + PKCE)。システムのブラウザでサインインし、
 *    localhost のループバックでコードを受け取る。組織のアカウント (Microsoft 365) と個人のアカウントの両方に対応する。
 *  - トークンは OS の暗号化 (Electron の safeStorage) をかけて userData に保存し、次回からはサインインを省く。
 *  - URL は Microsoft Graph の /shares で解決する (Word / OneDrive の「リンクのコピー」の URL、ファイルの URL)。
 *    docx はメモリに読み込むだけで、どこにも保存・送信しない。
 *  - 通信はメインプロセスからだけ行う (アプリの画面は引き続き外部と通信できない)。
 *    通信先は login.microsoftonline.com / graph.microsoft.com と、Graph が返すダウンロード用の URL だけ。
 *
 * フォルダの URL (Teams の課題の提出物が集まる SharePoint のフォルダなど) からは、その下の .docx を一覧して
 * 1つずつ読み込める (listFolder / open({ driveId, itemId }))。
 *
 * テスト用: DRA_GRAPH_MOCK に docx のパスを指定すると、サインインも通信もせずにそのファイルを返す。
 * DRA_GRAPH_MOCK_DIR にフォルダを指定すると、フォルダの URL をそのフォルダとして扱う。
 */

import { safeStorage, shell } from "electron";
import * as fs from "fs";
import * as path from "path";
import type { AccountInfo, ICachePlugin, PublicClientApplication as PCA, TokenCacheContext } from "@azure/msal-node";

/** Graph で読むために求める権限 (自分がアクセスできるファイル。組織の多くでは利用者の同意だけで使える) */
export const GRAPH_SCOPES = ["Files.Read.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_BYTES = 200 * 1024 * 1024;

export interface MicrosoftConfig {
  /** Azure に登録したアプリの (クライアント) ID */
  clientId: string;
  /** 既定は common (組織と個人の両方)。組織だけなら organizations やテナント ID */
  authority?: string;
}

export interface CloudFile {
  name: string;
  /** 最終更新日時 (ISO 8601) */
  mtime: string;
  bytes: Uint8Array;
  /** 次に開くときに使う */
  driveId: string;
  itemId: string;
  webUrl?: string;
}

export class MicrosoftError extends Error {
  constructor(
    message: string,
    /** 画面で案内を出し分けるための種類 */
    readonly code:
      | "notConfigured"
      | "badUrl"
      | "notDocx"
      | "notFolder"
      | "folder"
      | "notFound"
      | "forbidden"
      | "tooLarge"
      | "signInCancelled"
      | "network"
  ) {
    super(message);
  }
}

/** OneDrive / SharePoint の URL か (それ以外の URL は Graph に渡さない) */
export function isMicrosoftFileUrl(url: string): boolean {
  try {
    const u = new URL(url.trim());
    if (u.protocol !== "https:") return false;
    const h = u.hostname.toLowerCase();
    return (
      h.endsWith(".sharepoint.com") ||
      h === "onedrive.live.com" ||
      h === "1drv.ms" ||
      h === "d.docs.live.net" ||
      h.endsWith(".sharepoint.us") ||
      h.endsWith(".sharepoint.cn")
    );
  } catch {
    return false;
  }
}

/** Graph の /shares で使う ID ("u!" + URL の base64url) */
export function shareIdOf(url: string): string {
  return "u!" + Buffer.from(url.trim(), "utf-8").toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

/** トークンのキャッシュを、OS の暗号化をかけてファイルに保存する */
function cachePlugin(file: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      try {
        const raw = await fs.promises.readFile(file);
        const text = safeStorage.isEncryptionAvailable() ? safeStorage.decryptString(raw) : raw.toString("utf-8");
        ctx.tokenCache.deserialize(text);
      } catch {
        // まだ無い・読めない場合はサインインし直す
      }
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (!ctx.cacheHasChanged) return;
      const text = ctx.tokenCache.serialize();
      const data = safeStorage.isEncryptionAvailable() ? safeStorage.encryptString(text) : Buffer.from(text, "utf-8");
      await fs.promises.mkdir(path.dirname(file), { recursive: true });
      await fs.promises.writeFile(file, data, { mode: 0o600 });
    },
  };
}

const SIGNED_IN_PAGE = (title: string, body: string) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title>` +
  `<body style="font-family:-apple-system,'Segoe UI',sans-serif;padding:48px;text-align:center">` +
  `<h2>${title}</h2><p>${body}</p></body>`;

/** フォルダの中の .docx */
export interface CloudFolderFile {
  driveId: string;
  itemId: string;
  name: string;
  /** フォルダからの相対パス ("/" 区切り。ファイル名を含む) */
  path: string;
  size: number;
  mtime: string;
  /** Web で開く URL (図の見出しのリンクに使う) */
  webUrl?: string;
}

export interface CloudFolder {
  name: string;
  webUrl?: string;
  files: CloudFolderFile[];
  /** 上限に達して一覧を打ち切ったか */
  truncated: boolean;
}

/** 一覧にする .docx の上限 */
export const MAX_FOLDER_FILES = 2000;

/** 一括処理の対象にする .docx か (Word の一時ファイル ~$xxx.docx や隠しファイルは除く) */
export function isTargetDocx(name: string): boolean {
  return /\.docx$/i.test(name) && !name.startsWith("~$") && !name.startsWith(".");
}

export class MicrosoftClient {
  private pca?: PCA;

  constructor(
    private readonly config: () => MicrosoftConfig | undefined,
    private readonly cacheFile: string,
    private readonly mockFile = process.env.DRA_GRAPH_MOCK
  ) {}

  configured(): boolean {
    return !!this.mockFile || !!this.config()?.clientId;
  }

  private async app(): Promise<PCA> {
    const cfg = this.config();
    if (!cfg?.clientId) {
      throw new MicrosoftError("The Microsoft application (client) ID is not set", "notConfigured");
    }
    if (!this.pca) {
      const { PublicClientApplication } = await import("@azure/msal-node");
      this.pca = new PublicClientApplication({
        auth: { clientId: cfg.clientId, authority: cfg.authority ?? "https://login.microsoftonline.com/common" },
        cache: { cachePlugin: cachePlugin(this.cacheFile) },
      });
    }
    return this.pca;
  }

  /** サインインしているアカウント (無ければ undefined) */
  async account(): Promise<AccountInfo | undefined> {
    if (this.mockFile || !this.config()?.clientId) return undefined;
    const accounts = await (await this.app()).getTokenCache().getAllAccounts();
    return accounts[0];
  }

  async signOut(): Promise<void> {
    if (!this.config()?.clientId) return;
    const cache = (await this.app()).getTokenCache();
    for (const a of await cache.getAllAccounts()) await cache.removeAccount(a);
    await fs.promises.rm(this.cacheFile, { force: true });
  }

  /** アクセストークン。保存したトークンで足りなければ、ブラウザでサインインする */
  private async token(): Promise<string> {
    const pca = await this.app();
    const account = (await pca.getTokenCache().getAllAccounts())[0];
    if (account) {
      try {
        return (await pca.acquireTokenSilent({ account, scopes: GRAPH_SCOPES })).accessToken;
      } catch {
        // 期限切れ・権限の追加などは、サインインし直す
      }
    }
    try {
      const r = await pca.acquireTokenInteractive({
        scopes: GRAPH_SCOPES,
        openBrowser: async (url) => {
          await shell.openExternal(url);
        },
        successTemplate: SIGNED_IN_PAGE("Signed in / サインインしました", "You can close this tab and return to the app. / このタブを閉じてアプリに戻ってください。"),
        errorTemplate: SIGNED_IN_PAGE("Sign-in failed / サインインできませんでした", "Please return to the app and try again. / アプリに戻ってやり直してください。"),
      });
      return r.accessToken;
    } catch (err) {
      throw new MicrosoftError(err instanceof Error ? err.message : String(err), "signInCancelled");
    }
  }

  private async graph(pathAndQuery: string, token: string, init: RequestInit = {}): Promise<Response> {
    let res: Response;
    try {
      res = await fetch(`${GRAPH}${pathAndQuery}`, { ...init, headers: { ...(init.headers ?? {}), authorization: `Bearer ${token}` } });
    } catch (err) {
      throw new MicrosoftError(err instanceof Error ? err.message : String(err), "network");
    }
    if (res.status === 404 || res.status === 400) throw new MicrosoftError(await errorText(res), "notFound");
    if (res.status === 401 || res.status === 403) throw new MicrosoftError(await errorText(res), "forbidden");
    if (!res.ok) throw new MicrosoftError(await errorText(res), "network");
    return res;
  }

  /** URL (または前回開いたときの driveId / itemId) の docx を読み込む */
  async open(target: { url: string } | { driveId: string; itemId: string }): Promise<CloudFile> {
    if (process.env.DRA_GRAPH_MOCK_DIR && "driveId" in target && target.driveId === "mock-dir") {
      const p = path.join(process.env.DRA_GRAPH_MOCK_DIR, target.itemId);
      const stat = await fs.promises.stat(p);
      return { name: path.basename(p), mtime: stat.mtime.toISOString(), bytes: new Uint8Array(await fs.promises.readFile(p)), driveId: "mock-dir", itemId: target.itemId };
    }
    if (this.mockFile) {
      const bytes = await fs.promises.readFile(this.mockFile);
      const stat = await fs.promises.stat(this.mockFile);
      return { name: path.basename(this.mockFile), mtime: stat.mtime.toISOString(), bytes: new Uint8Array(bytes), driveId: "mock", itemId: "mock" };
    }
    if ("url" in target && !isMicrosoftFileUrl(target.url)) throw new MicrosoftError(target.url, "badUrl");
    const token = await this.token();
    const select = "$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,file,folder";
    const meta = (await (
      "url" in target
        ? await this.graph(`/shares/${shareIdOf(target.url)}/driveItem?${select}`, token, { headers: { prefer: "redeemSharingLinkIfNecessary" } })
        : await this.graph(`/drives/${encodeURIComponent(target.driveId)}/items/${encodeURIComponent(target.itemId)}?${select}`, token)
    ).json()) as {
      id: string;
      name: string;
      size?: number;
      lastModifiedDateTime?: string;
      webUrl?: string;
      parentReference?: { driveId?: string };
      folder?: unknown;
    };
    if (meta.folder) throw new MicrosoftError(meta.name, "folder");
    if (!/\.docx$/i.test(meta.name)) throw new MicrosoftError(meta.name, "notDocx");
    if ((meta.size ?? 0) > MAX_BYTES) throw new MicrosoftError(meta.name, "tooLarge");
    const driveId = meta.parentReference?.driveId ?? ("driveId" in target ? target.driveId : "");
    // 中身は Graph が返す、期限付きのダウンロード用の URL へ転送される (fetch が転送をたどる)
    const res = await this.graph(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(meta.id)}/content`, token);
    const bytes = new Uint8Array(await res.arrayBuffer());
    if (bytes.byteLength > MAX_BYTES) throw new MicrosoftError(meta.name, "tooLarge");
    return {
      name: meta.name,
      mtime: meta.lastModifiedDateTime ?? new Date().toISOString(),
      bytes,
      driveId,
      itemId: meta.id,
      webUrl: meta.webUrl,
    };
  }

  /**
   * フォルダの URL の下にある .docx を、サブフォルダまでたどって一覧する (MAX_FOLDER_FILES 件まで)。
   * Teams の課題の提出物 (SharePoint の「Student Work」などのフォルダ) を想定している
   */
  async listFolder(url: string): Promise<CloudFolder> {
    const mockDir = process.env.DRA_GRAPH_MOCK_DIR;
    if (mockDir) return listLocalAsCloud(mockDir);
    if (!isMicrosoftFileUrl(url)) throw new MicrosoftError(url, "badUrl");
    const token = await this.token();
    const root = (await (
      await this.graph(`/shares/${shareIdOf(url)}/driveItem?$select=id,name,webUrl,folder,parentReference`, token, {
        headers: { prefer: "redeemSharingLinkIfNecessary" },
      })
    ).json()) as { id: string; name: string; webUrl?: string; folder?: unknown; parentReference?: { driveId?: string } };
    if (!root.folder) throw new MicrosoftError(root.name, "notFolder");
    const driveId = root.parentReference?.driveId ?? "";
    const files: CloudFolderFile[] = [];
    let truncated = false;
    // 幅優先でたどる (フォルダ ID と、そこまでの相対パス)
    const queue: { id: string; prefix: string }[] = [{ id: root.id, prefix: "" }];
    while (queue.length && !truncated) {
      const { id, prefix } = queue.shift()!;
      let next: string | undefined = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(id)}/children?$select=id,name,size,lastModifiedDateTime,webUrl,folder,file&$top=200`;
      while (next) {
        const page = (await (await this.graph(next, token)).json()) as {
          value: { id: string; name: string; size?: number; lastModifiedDateTime?: string; webUrl?: string; folder?: unknown }[];
          "@odata.nextLink"?: string;
        };
        for (const it of page.value) {
          if (it.folder) queue.push({ id: it.id, prefix: `${prefix}${it.name}/` });
          else if (isTargetDocx(it.name)) {
            if (files.length >= MAX_FOLDER_FILES) {
              truncated = true;
              break;
            }
            files.push({
              driveId,
              itemId: it.id,
              name: it.name,
              path: `${prefix}${it.name}`,
              size: it.size ?? 0,
              mtime: it.lastModifiedDateTime ?? "",
              webUrl: it.webUrl,
            });
          }
        }
        // 次のページは完全な URL で返る
        next = truncated ? undefined : page["@odata.nextLink"]?.replace(GRAPH, "");
      }
    }
    return { name: root.name, webUrl: root.webUrl, files, truncated };
  }
}

/** テスト用: ローカルのフォルダを、クラウドのフォルダの一覧として返す */
async function listLocalAsCloud(dir: string): Promise<CloudFolder> {
  const files: CloudFolderFile[] = [];
  const walk = async (rel: string) => {
    for (const e of await fs.promises.readdir(path.join(dir, rel), { withFileTypes: true })) {
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) await walk(r);
      else if (isTargetDocx(e.name)) {
        const st = await fs.promises.stat(path.join(dir, r));
        files.push({
          driveId: "mock-dir",
          itemId: r,
          name: e.name,
          path: r,
          size: st.size,
          mtime: st.mtime.toISOString(),
          webUrl: `https://contoso.sharepoint.com/mock/${encodeURI(r)}`,
        });
      }
    }
  };
  await walk("");
  return { name: path.basename(dir), files, truncated: false };
}

async function errorText(res: Response): Promise<string> {
  try {
    const j = (await res.json()) as { error?: { message?: string } };
    return j.error?.message ?? `HTTP ${res.status}`;
  } catch {
    return `HTTP ${res.status}`;
  }
}
