// PDF のページを画像にする (pdf.js)。メインプロセスの非表示のウィンドウ (dra-pdf://app/) で動き、
// ネットワークには出られない。PDF は同じオリジンの dra-pdf://app/doc/<番号> (メインプロセスが許可したファイル) から読む。
import * as pdfjs from "./pdf.min.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = "./pdf.worker.min.mjs";

/** 開いた PDF (URL → 文書)。新しいものから2つまで覚えておく */
const docs = new Map();

async function open(url) {
  let doc = docs.get(url);
  if (doc) {
    docs.delete(url);
    docs.set(url, doc);
    return doc;
  }
  doc = await pdfjs.getDocument({
    url,
    cMapUrl: "./cmaps/",
    cMapPacked: true,
    // 埋め込まれていないフォントは OS のフォントで描く
    useSystemFonts: true,
    isEvalSupported: false,
  }).promise;
  docs.set(url, doc);
  while (docs.size > 2) {
    const [oldest, d] = docs.entries().next().value;
    docs.delete(oldest);
    void d.destroy();
  }
  return doc;
}

/** ページ (0 から) を、maxWidth × maxHeight (画素) に収まる大きさの PNG にする */
window.renderPage = async (url, pageIndex, maxWidth, maxHeight) => {
  const doc = await open(url);
  const pageCount = doc.numPages;
  const page = await doc.getPage(Math.min(Math.max(1, pageIndex + 1), pageCount));
  const base = page.getViewport({ scale: 1 });
  const scale = Math.max(0.1, Math.min(maxWidth / base.width, maxHeight / base.height));
  const vp = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(vp.width));
  canvas.height = Math.max(1, Math.floor(vp.height));
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport: vp }).promise;
  page.cleanup();
  return { dataUrl: canvas.toDataURL("image/png"), pageCount, width: base.width, height: base.height };
};

window.closeDoc = (url) => {
  const d = docs.get(url);
  docs.delete(url);
  if (d) void d.destroy();
};

window.ready = true;
