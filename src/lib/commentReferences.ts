export interface CommentLink {
  url: string;
  label: string;
}

const uri = /[a-z][a-z\d+.-]*:\/\/[^\s<>"'「」『』【】、。！？]+/gi;

export function omitCommentUrls(text: string): string {
  return text.replace(uri, "");
}

export function commentLinkUrl(value: unknown): string | undefined {
  if (typeof value !== "string" || !/^https?:\/\//i.test(value)) return undefined;
  try {
    const url = new URL(value);
    if (!url.username && !url.password) return url.href;
  } catch {
    return undefined;
  }
}

export function commentReferences(text: string, wordLinks: CommentLink[] = []): { tags: string[]; links: CommentLink[] } {
  const links = new Map<string, CommentLink>();
  for (const link of wordLinks) {
    const url = commentLinkUrl(link.url);
    if (url) links.set(url, { url, label: link.label || url });
  }
  // Remove URI spans first so that their fragments cannot become tags.
  const withoutLinks = text.replace(uri, (match) => {
    let value = match.replace(/[.,;:!?]+$/, "");
    while (value.endsWith(")") && (value.match(/\)/g)?.length ?? 0) > (value.match(/\(/g)?.length ?? 0)) {
      value = value.slice(0, -1);
    }
    const url = commentLinkUrl(value);
    if (url && !links.has(url)) links.set(url, { url, label: value });
    return " ";
  });
  const tags = new Set<string>();
  for (const [, tag] of withoutLinks.matchAll(/(?:^|[^\p{L}\p{M}\p{N}_#＃])[#＃]([\p{L}\p{N}_][\p{L}\p{M}\p{N}_-]*)/gu)) {
    tags.add(tag);
  }
  return { tags: [...tags], links: [...links.values()] };
}
