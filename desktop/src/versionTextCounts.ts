import type { AppLang, VersionTextCounts } from "./shared";

const segmenters: Record<AppLang, Intl.Segmenter> = {
  en: new Intl.Segmenter("en", { granularity: "word" }),
  ja: new Intl.Segmenter("ja", { granularity: "word" }),
};

export function countVersionText(paragraphs: readonly { text: string }[], lang: AppLang): VersionTextCounts {
  let characters = 0;
  let words = 0;
  for (const { text } of paragraphs) {
    characters += Array.from(text).length;
    for (const word of segmenters[lang].segment(text)) if (word.isWordLike) words++;
  }
  return { characters, words };
}
