/**
 * 拡張のホストのページ用の preload。拡張の実行環境 (ext-host/runtime.ts) に、メインプロセスとの RPC だけを公開する。
 */
import { contextBridge, ipcRenderer } from "electron";

type Handler = (method: string, args: unknown) => Promise<unknown>;

contextBridge.exposeInMainWorld("extBridge", {
  call: (method: string, args: unknown) => ipcRenderer.invoke("ext:call", method, args),
  onInvoke: (handler: Handler) => {
    ipcRenderer.on("ext:invoke", async (_e, msg: { id: number; method: string; args: unknown }) => {
      try {
        const value = await handler(msg.method, msg.args);
        ipcRenderer.send("ext:reply", { id: msg.id, ok: true, value });
      } catch (err) {
        ipcRenderer.send("ext:reply", { id: msg.id, ok: false, error: err instanceof Error ? err.message : String(err) });
      }
    });
  },
  ready: () => ipcRenderer.send("ext:ready"),
});
