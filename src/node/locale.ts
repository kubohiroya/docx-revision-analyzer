/**
 * locale.ts (Node.js 専用)
 *
 * macOS のシステムの言語設定 (defaults read -g AppleLocale) を読む関数を i18n に登録する。
 * import するだけで登録される。
 */

import { spawnSync } from "child_process";
import { setSystemLocaleProvider } from "../lib/i18n";

export function macSystemLocale(): string | undefined {
  if (process.platform !== "darwin") return undefined;
  try {
    const r = spawnSync("defaults", ["read", "-g", "AppleLocale"], { encoding: "utf-8", timeout: 2000 });
    if (r.status === 0 && r.stdout.trim()) return r.stdout.trim();
  } catch {
    // 取得できなければ Intl に任せる
  }
  return undefined;
}

setSystemLocaleProvider(macSystemLocale);
