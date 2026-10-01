# Extension API v1 / 拡張モジュール API v1

Extensions add their own highlights, colors, panels, dialogs and forms to the desktop app — and, with permission,
send data to a declared web service — without changing the app itself.

拡張モジュールは、本体を変えずに、独自のハイライト・色・パネル・ダイアログ・フォーム (権限があれば外部サービスへの送信も)
を追加します。

Types: [`src/extension-api.ts`](src/extension-api.ts). Example: [`extensions/sample-summary/`](extensions/sample-summary/).

## Form / 形態

An extension is a directory with `manifest.json` and an ES-module entry:

```
extensions/my-extension/
  manifest.json
  index.js        # export function activate(host) { ... }
```

```json
{
  "id": "com.example.sample",
  "name": { "en": "Sample", "ja": "サンプル" },
  "version": "0.1.0",
  "apiVersion": 1,
  "main": "index.js",
  "description": { "en": "…", "ja": "…" },
  "permissions": {
    "ui": ["dialog", "panel", "form"],
    "network": ["https://api.example.com"],
    "storage": true,
    "documentText": false
  }
}
```

- `id`: reverse-domain, lowercase (`[a-z0-9]` separated by `.`, `-`, `_`). Category and classifier ids must start with
  `<id>.`.
- `main`: a relative path; the file is loaded as an ES module (it may `import` other files in the same directory).
- `apiVersion`: must be one this app supports (currently `1`); otherwise the extension isn't loaded.

### Distribution / 配布

For now only extensions **bundled with the app** (`desktop/extensions/`) are loaded, and each is **disabled by
default**. When the user enables one in Settings, the app shows the permissions it asks for and asks for
confirmation. Installing third-party extensions is not supported yet; it will require an explicit install step and
the same permission confirmation.

当面はアプリに同梱した拡張だけを読み込み、既定は無効です。設定で有効にするときに、求める権限を示して確認します。
第三者の拡張のインストールは、明示的なインストール操作と権限の確認を用意してから対応します。

## Isolation / 隔離

Each enabled extension runs in its own hidden, sandboxed renderer process:

- no Node.js (`require`, `process`, file system), no access to the app's window or other extensions;
- its own origin (`dra-ext://<id>/`) and in-memory session; it can only load files from its own directory;
- no network: Content-Security-Policy `connect-src 'none'` and a request filter that blocks every non-`dra-ext:` URL.
  The only way out is `host.net.post`, which the app performs after checking the manifest (#16 adds a send log and
  confirmation);
- it talks to the app only through the host API (an RPC over IPC).

拡張ごとに、表示しない sandbox のレンダラで動かします。Node.js の機能・アプリのウィンドウ・ほかの拡張には触れられず、
ネットワークにも出られません (外へ出る手段は、アプリが manifest を確かめてから送る `host.net.post` だけ)。

### Errors and timeouts / 例外と時間切れ

- `activate` must finish within 10 s, `classify` within 10 s, `onAnalysisComplete` callbacks within 30 s.
- A throwing or timed-out classifier contributes no highlights; the app shows a notice and continues.
- If the extension's process crashes, it's marked as failed (shown in Settings); the app keeps working.

## Host API

```ts
export function activate(host: HostApi): void | Promise<void>;

interface HostApi {
  registerCategory(c: { id; color; label; priority; pattern? }): void;
  registerClassifier(c: { id; version; classify(ctx): HighlightSpec[] | Promise<HighlightSpec[]> }): void;
  onAnalysisComplete(cb: (r: AnalysisResult) => void | Promise<void>): void;
  ui: {
    showDialog(d: DialogSpec): Promise<{ buttonId: string | null }>;   // permission ui: "dialog"
    showPanel(p: PanelSpec): void;                                      // permission ui: "panel"
    openForm(f: FormSpec): Promise<FormResult | null>;                  // permission ui: "form"
  };
  storage: { get; set; delete; keys };     // permission storage (per extension, ≤ 1 MB, on this computer)
  net: { post(url, body): Promise<{ ok; status; body }> };  // permission network (declared https origins only)
  app: { version: string; locale: "ja" | "en"; apiVersion: number };
}
```

### Categories and classifiers / カテゴリと分類器

Categories are drawn like the built-in ones (see the main README, "Color categories"): `priority` decides which
category fills a paragraph when several apply (built-ins are 0, rule levels 10+). Give a `pattern` so the category
isn't distinguished by color alone, and use colors with at least 3:1 contrast against white.

A classifier receives `ctx` — `positioned` (every insertion/deletion with its position), `sessions`, and
`windowsFor(options)` to compute insertion windows and their features (#8) — and returns highlights:

```js
{ categoryId: "com.example.sample.long", eventIds: ["12"], start: "2026-05-11T14:02:00Z", end: "…",
  features: { chars: 295 }, reason: { en: "…", ja: "…" } }
```

Highlights whose category isn't registered are ignored. For each insertion the app keeps the highlight with the
highest category priority across all classifiers (ties: the built-in classifiers first, then extensions in order).

### Analysis result / 解析結果

`onAnalysisComplete` receives `AnalysisResult`: file name and time, revision totals, positioned events, sessions,
the rule set and insertion windows with their levels, all highlights, and the final text.

**Document text** (`positioned[].text`, `finalText`) is included only when the manifest declares
`"documentText": true`; otherwise those fields are empty strings. Declare it only if the extension really needs the
text — users see it as a separate permission.

本文 (`positioned[].text`, `finalText`) は、manifest で `"documentText": true` を宣言した拡張にだけ渡します
(無ければ空文字列)。利用者には別の権限として表示されます。

### UI (declarative) / UI (宣言的)

Extensions describe UI as JSON; the app renders it (as plain text — no HTML), marked "From the extension …".

- `DialogSpec`: `{ title, message, buttons?: [{ id, label, primary? }] }` → `{ buttonId }` (`null` if dismissed)
- `FormSpec`: `{ title, description?, fields: [...], submitLabel? }` with fields of type `text`, `textarea`, `number`,
  `checkbox`, `select` → values keyed by field id, or `null` if cancelled
- `PanelSpec`: `{ id, title, blocks: [...] }` with blocks `heading`, `text`, `list`, `keyValue`, `table`; shown under
  the Highlights tab; a panel with the same id replaces the previous one. Panels are cleared on each analysis.

All texts are `LocalizedText`: a string or `{ ja, en }`.

## Versioning / バージョンの方針

- Within API v1, changes are additive only (new optional fields, new methods). Existing behavior isn't changed.
- A breaking change becomes API v2; the app may support several versions at once (`SUPPORTED_API_VERSIONS`), and an
  extension declaring an unsupported `apiVersion` is not loaded (shown with the reason in Settings).
- `host.app.apiVersion` tells the extension which version it runs under.

## Writing an extension / 作り方

1. Copy `extensions/sample-summary/` to `extensions/<your-name>/` and change `id`, `name`, permissions.
2. Write `index.js` (plain ES module; add `// @ts-check` and the JSDoc type import from the sample for type checking).
3. Run the app (`npm start`), enable it in Settings, and open a document.
   To try it without the UI: `DRA_SMOKE_EXTENSION=<id> DRA_SMOKE_WAIT_MS=3000 npm run smoke`.
