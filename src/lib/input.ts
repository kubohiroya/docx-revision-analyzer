/**
 * input.ts
 *
 * 解析の入口が受け取る docx のバイト列の型。
 * Node.js の Buffer も Uint8Array なのでそのまま渡せる。ブラウザでは File.arrayBuffer() の結果などを渡す。
 */
export type DocxInput = ArrayBuffer | Uint8Array;
