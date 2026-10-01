/**
 * electron-builder の設定 (配布物の作成・署名・公証・GitHub Releases への公開)。
 *
 * 署名と公証は、環境変数 (CI では GitHub の Secrets) があるときだけ行う。無ければ署名なしで作る。
 *   macOS の署名:   CSC_LINK (Developer ID Application 証明書の .p12 を base64 にしたもの), CSC_KEY_PASSWORD
 *   macOS の公証:   APPLE_API_KEY (.p8 のパス), APPLE_API_KEY_ID, APPLE_API_ISSUER
 *                   または APPLE_ID, APPLE_APP_SPECIFIC_PASSWORD, APPLE_TEAM_ID
 *   Windows の署名: WIN_CSC_LINK, WIN_CSC_KEY_PASSWORD
 *                   (Windows 用の証明書の発行者名を WIN_PUBLISHER_NAME に入れると、自動更新で署名を検証する)
 * 詳しくは DISTRIBUTION.md。
 */
const env = process.env;
const notarize =
  !!(env.APPLE_API_KEY && env.APPLE_API_KEY_ID && env.APPLE_API_ISSUER) ||
  !!(env.APPLE_ID && env.APPLE_APP_SPECIFIC_PASSWORD && env.APPLE_TEAM_ID);
const macSigned = !!env.CSC_LINK || !!env.CSC_NAME;

/** @type {import("electron-builder").Configuration} */
module.exports = {
  appId: "io.github.kubohiroya.docx-revision-analyzer",
  productName: "Docx Revision Analyzer",
  copyright: "Copyright © Hiroya Kubo",
  directories: { output: "release", buildResources: "build" },
  files: ["dist/**", "extensions/**", "package.json"],
  // 依存はすべて dist/main.js にバンドル済み。electron-updater だけは実行時に node_modules から読む
  asar: true,
  fileAssociations: [{ ext: "docx", name: "Word Document", role: "Viewer", rank: "Alternate" }],
  publish: [{ provider: "github", owner: "kubohiroya", repo: "docx-revision-analyzer", releaseType: "draft" }],
  mac: {
    target: [
      { target: "dmg", arch: ["universal"] },
      // 自動更新 (Squirrel.Mac) には zip が必要
      { target: "zip", arch: ["universal"] },
    ],
    category: "public.app-category.productivity",
    hardenedRuntime: true,
    gatekeeperAssess: false,
    entitlements: "build/entitlements.mac.plist",
    entitlementsInherit: "build/entitlements.mac.plist",
    // 証明書が無ければ署名しない (ad-hoc 署名のみ)
    identity: macSigned ? undefined : null,
    notarize,
  },
  dmg: { sign: false },
  win: {
    target: [{ target: "nsis", arch: ["x64", "arm64"] }, { target: "zip", arch: ["x64", "arm64"] }],
    // 署名した場合、自動更新でインストーラの発行者名を検証する
    ...(env.WIN_PUBLISHER_NAME ? { publisherName: [env.WIN_PUBLISHER_NAME] } : {}),
    verifyUpdateCodeSignature: !!env.WIN_PUBLISHER_NAME,
  },
  nsis: { oneClick: false, perMachine: false, allowToChangeInstallationDirectory: true },
  linux: { target: ["AppImage"], category: "Office" },
};
