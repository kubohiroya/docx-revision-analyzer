/**
 * preload: レンダラに window.app (shared.ts の AppApi) を公開する。
 * レンダラは Node.js の機能を持たず (sandbox)、この API を通してだけメインプロセスとやり取りする。
 */
import { contextBridge, ipcRenderer, webUtils } from "electron";
import type { AppApi, ExtensionUiRequest } from "./shared";

const api: AppApi = {
  extensions: {
    list: () => ipcRenderer.invoke("ext:list"),
    setEnabled: (id, enabled) => ipcRenderer.invoke("ext:set-enabled", id, enabled),
    registrations: () => ipcRenderer.invoke("ext:registrations"),
    classify: (extId, classifierId, ctx) => ipcRenderer.invoke("ext:classify", extId, classifierId, ctx),
    analysisComplete: (result) => ipcRenderer.send("ext:analysis-complete", result),
    sendLog: (id) => ipcRenderer.invoke("ext:send-log", id),
    onSendLogChanged: (cb) => {
      ipcRenderer.on("ext-send-log-changed", (_e, id: string) => cb(id));
    },
    onUiRequest: (cb) => {
      ipcRenderer.on("ext-ui", async (_e, req: ExtensionUiRequest) => {
        let value: unknown = null;
        try {
          value = await cb(req);
        } finally {
          ipcRenderer.send("ext-ui-reply", { reqId: req.reqId, value });
        }
      });
    },
  },
  openDocxDialog: () => ipcRenderer.invoke("open-docx-dialog"),
  readDocx: (p) => ipcRenderer.invoke("read-docx", p),
  pathForFile: (file) => webUtils.getPathForFile(file),
  preserveHistory: (p) => ipcRenderer.invoke("preserve-history", p),
  chooseRulesDialog: () => ipcRenderer.invoke("choose-rules-dialog"),
  readRules: (p) => ipcRenderer.invoke("read-rules", p),
  saveFile: (name, data, kind) => ipcRenderer.invoke("save-file", name, data, kind),
  getSettings: () => ipcRenderer.invoke("get-settings"),
  setSettings: (s) => ipcRenderer.invoke("set-settings", s),
  systemLang: () => ipcRenderer.invoke("system-lang"),
  onOpenPath: (cb) => {
    ipcRenderer.on("open-path", (_e, p: string) => cb(p));
  },
  version: () => ipcRenderer.invoke("version"),
  updatesAvailable: () => ipcRenderer.invoke("updates-available"),
};

contextBridge.exposeInMainWorld("app", api);
contextBridge.exposeInMainWorld("appInternal", {
  ready: () => ipcRenderer.send("renderer-ready"),
  smokeAnalyzed: (summary: { ok: boolean; message: string }) => ipcRenderer.send("smoke-analyzed", summary),
});
