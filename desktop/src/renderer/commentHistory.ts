import { commentReferences } from "../../../src/lib/commentReferences";
import type { CommentHistoryDetail, CommentObservation } from "../shared";
import { el, header, modal } from "./extensions";
import type { Strings } from "./strings";

export function showCommentHistory(
  detail: CommentHistoryDetail,
  strings: Strings,
  formatDate: (iso: string) => string,
  openLink: (url: string) => void,
): Promise<void> {
  const labels: Record<CommentObservation["type"], string> = {
    removed: strings.commentHistoryAbsent,
    reappeared: strings.commentHistoryReappeared,
    resolved: strings.commentResolved,
    reopened: strings.commentHistoryReopened,
  };
  return modal<void>((dialog, close) => {
    header(dialog, strings.commentHistoryTitle);
    if (detail.version) {
      dialog.append(el("p", strings.commentHistoryVersion(detail.version.n, formatDate(detail.version.archivedAt))));
    }
    if (detail.versionUnavailable) dialog.append(el("p", strings.commentHistoryMissing, "hint"));
    if (!detail.passage) dialog.append(el("p", strings.commentHistoryUnlocated, "hint"));
    if (detail.locatedByContext) dialog.append(el("p", strings.commentHistoryByContext, "hint"));
    const passage = detail.passage ?? detail.recordedAnchor;
    if (passage) {
      if (!detail.passage) dialog.append(el("h4", strings.commentHistoryQuote));
      const quote = el("blockquote", undefined, "comment-passage");
      const mark = el("mark", passage.quote || "│");
      if (!passage.quote) mark.setAttribute("aria-label", strings.commentHistoryPoint);
      quote.append(passage.prefix, mark, passage.suffix);
      dialog.append(quote);
    }
    if (detail.observations.length) {
      const list = el("ul", undefined, "comment-observations");
      for (const observation of detail.observations) {
        list.append(el("li", strings.commentHistoryObservation(labels[observation.type], formatDate(observation.at), observation.version)));
      }
      dialog.append(list);
    }
    const thread = el("div", undefined, "comment-thread");
    detail.comments.forEach((comment, index) => {
      if (index === 1) thread.append(el("h4", strings.commentHistoryReplies));
      const message = el("div", undefined, "comment-message");
      const date = comment.dateUtc ?? comment.date;
      const byline = [comment.author, date && formatDate(date)].filter(Boolean).join(" · ");
      if (byline) message.append(el("div", byline, "hint"));
      message.append(el("p", comment.text, "comment-text"));
      const { links } = commentReferences(comment.text, comment.links);
      if (links.length) {
        const references = el("ul", undefined, "comment-references");
        references.setAttribute("aria-label", strings.commentHistoryReferences);
        for (const link of links) {
          const item = el("li");
          const button = el("button", link.label, "link");
          button.title = link.url;
          button.onclick = () => openLink(link.url);
          item.append(button);
          references.append(item);
        }
        message.append(references);
      }
      thread.append(message);
    });
    const buttons = el("div", undefined, "buttons");
    const done = el("button", strings.close);
    done.onclick = () => close();
    buttons.append(done);
    dialog.append(thread, buttons);
  }, undefined);
}
