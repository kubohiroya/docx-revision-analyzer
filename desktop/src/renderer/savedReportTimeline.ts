import type { VersionBranchCandidate, VersionBranchEntry, VersionBranchHoldReason } from "../../../src/lib/lineageBranches";
import type { AppLang, SavedLineageVersion } from "../shared";
import { el } from "./extensions";

const LABELS = {
  en: {
    title: "Version branches",
    axis: "Observation time →",
    plottedTime: "Plotted time",
    inferred: "Inferred relation",
    possible: "Candidate · on hold",
    compareInterval: "Compare interval",
    empty: "Select versions to show the timeline.",
    unknown: "Time unknown",
    evidence: "Evidence",
    candidate: "Candidate",
    accepted: "Inferred source",
    matches: (count: number, distinctive: number) => `${count} matching sentences · ${distinctive} specific to this candidate`,
    coverage: (source: string, target: string) => `Text coverage · source ${source} · target ${target}`,
    location: (from: number, to: number) => `Paragraph ${from} → paragraph ${to}`,
    distinctive: "Specific to this candidate",
    archiveNumber: "Archive sequence",
    duplicateOf: "Same file content",
    sameText: "Same visible text",
    timeSources: { archivedAt: "Archived", sourceModified: "Modified", fileCreatedAt: "Created" },
    orders: { earlier: "Earlier", "same-time": "Same time", "unconfirmed-time": "Time order uncertain", "archive-order-conflict": "Archive sequence conflicts with the timestamps" },
    holds: {
      "missing-time": "Time unknown", "creation-time-only": "Only the creation time is available",
      "conflicting-times": "Timestamps conflict", unavailable: "Version could not be read",
      "duplicate-file": "Identical file", "same-visible-text": "Identical visible text",
      "no-earlier-version": "No known earlier version", "no-shared-body-text": "No matching body text",
      "only-common-body-text": "Only text common to all versions matches", "time-order-unconfirmed": "Time order could not be established",
      "competing-candidates": "Multiple candidates have distinctive matches",
      "insufficient-distinctive-text": "Too little distinctive text matches", "insufficient-coverage": "Matching text covers too little of the documents",
    } satisfies Record<VersionBranchHoldReason, string>,
  },
  ja: {
    title: "保存版の分岐",
    axis: "観測日時 →",
    plottedTime: "図で使った日時",
    inferred: "推定した派生関係",
    possible: "候補・判定保留",
    compareInterval: "区間を比較",
    empty: "表示する版を選んでください。",
    unknown: "日時不明",
    evidence: "判定の根拠",
    candidate: "候補",
    accepted: "推定した派生元",
    matches: (count: number, distinctive: number) => `一致 ${count} 文・候補固有 ${distinctive} 文`,
    coverage: (source: string, target: string) => `本文の一致割合 · 派生元側 ${source} · 派生先側 ${target}`,
    location: (from: number, to: number) => `段落 ${from} → 段落 ${to}`,
    distinctive: "候補固有の一致",
    archiveNumber: "保全順",
    duplicateOf: "同一ファイルの内容",
    sameText: "同一の表示本文",
    timeSources: { archivedAt: "保管日時", sourceModified: "更新日時", fileCreatedAt: "作成日時" },
    orders: { earlier: "時刻が先行", "same-time": "同時刻", "unconfirmed-time": "日時の前後が不明", "archive-order-conflict": "保全順と日時が矛盾" },
    holds: {
      "missing-time": "日時不明", "creation-time-only": "作成日時のみ判明",
      "conflicting-times": "日時に矛盾あり", unavailable: "保存版を読み込めない",
      "duplicate-file": "同一のファイル", "same-visible-text": "表示本文が同一",
      "no-earlier-version": "先行する版を確認できない", "no-shared-body-text": "一致する本文なし",
      "only-common-body-text": "全版に共通する本文のみ一致", "time-order-unconfirmed": "日時の前後を確定できない",
      "competing-candidates": "複数の候補に固有の一致がある",
      "insufficient-distinctive-text": "候補固有の一致が少ない", "insufficient-coverage": "本文の一致割合が不足",
    } satisfies Record<VersionBranchHoldReason, string>,
  },
};

interface TimelineNode {
  entry: VersionBranchEntry;
  column: number;
  lane: number;
}
interface TimelineEdge {
  from: TimelineNode;
  to: TimelineNode;
  inferred: boolean;
}

export function layoutSavedVersionTimeline(entries: readonly VersionBranchEntry[]) {
  // Equal timestamps share a column. Unknown dates go in a separate last column.
  const times = [...new Set(entries.map(entry => entry.time.selected?.value))];
  const columns = new Map(times.map((time, index) => [time, index]));
  const nodes = new Map(entries.map(entry =>
    [entry.id, { entry, column: columns.get(entry.time.selected?.value)!, lane: 0 } satisfies TimelineNode]));
  const sources = new Set(entries.flatMap(entry => entry.decision.status === "inferred" ? [entry.decision.parentId] : []));
  const laneEnds: TimelineNode[] = [];
  const edges: TimelineEdge[] = [];
  for (const node of nodes.values()) {
    const { entry } = node;
    if (entry.decision.status === "inferred") {
      const parent = nodes.get(entry.decision.parentId)!;
      node.lane = laneEnds[parent.lane] === parent ? parent.lane : laneEnds.length;
      laneEnds[node.lane] = node;
      edges.push({ from: parent, to: node, inferred: true });
    } else if (sources.has(entry.id)) {
      node.lane = laneEnds.push(node) - 1;
    }
  }
  const isolatedStart = laneEnds.length;
  const isolatedCounts = new Map<number, number>();
  for (const node of nodes.values()) {
    const { entry } = node;
    if (entry.decision.status === "inferred") continue;
    if (!sources.has(entry.id)) {
      const count = isolatedCounts.get(node.column) ?? 0;
      node.lane = isolatedStart + count;
      isolatedCounts.set(node.column, count + 1);
    }
    for (const candidate of entry.candidates.filter(candidate => candidate.order === "earlier").slice(0, 2)) {
      edges.push({ from: nodes.get(candidate.parentId)!, to: node, inferred: false });
    }
  }
  const lanes = isolatedStart + Math.max(0, ...isolatedCounts.values());
  return { times, nodes: [...nodes.values()], edges, lanes };
}

const svgNamespace = "http://www.w3.org/2000/svg";
const columnStep = 138;
const firstColumn = 72;
const laneStep = 42;
const firstLane = 37;

export function savedReportTimeline(
  lang: AppLang, versions: readonly SavedLineageVersion[], name: (index: number) => string,
  choosePath: (key: string) => void, chooseInterval: (after: number) => void,
) {
  const s = LABELS[lang];
  const date = new Intl.DateTimeFormat(lang, { dateStyle: "medium", timeStyle: "long" });
  const tick = new Intl.DateTimeFormat(lang, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const percent = new Intl.NumberFormat(lang, { style: "percent", maximumFractionDigits: 1 });
  const element = el("details", undefined, "report-branch-timeline");
  element.open = true;
  const caption = el("div", undefined, "report-branch-caption");
  caption.append(el("span", s.axis), el("span", s.inferred, "report-branch-legend"),
    el("span", s.possible, "report-branch-legend report-branch-possible"));
  const viewport = el("div", undefined, "report-branch-viewport");
  viewport.setAttribute("role", "group");
  viewport.setAttribute("aria-label", s.title);
  const evidence = el("details", undefined, "report-branch-evidence");
  evidence.hidden = true;
  const evidenceTitle = el("summary", s.evidence);
  const evidenceBody = el("div");
  evidence.append(evidenceTitle, evidenceBody);
  element.append(el("summary", s.title), caption, viewport, evidence);

  let inspected: string | undefined;
  let nodes = new Map<string, { entry: VersionBranchEntry; button: HTMLButtonElement }>();
  let drawnEdges: { edge: TimelineEdge; line: SVGPathElement; button?: HTMLButtonElement }[] = [];
  let disabled = false;

  function showCandidate(entry: VersionBranchEntry, candidate: VersionBranchCandidate): HTMLElement {
    const accepted = entry.decision.status === "inferred" && entry.decision.parentId === candidate.parentId;
    const box = el("details");
    box.append(el("summary", [
      accepted ? s.accepted : s.candidate, name(Number(candidate.parentId)), s.orders[candidate.order],
      s.matches(candidate.matches.length, candidate.distinctiveMatches),
    ].join(" · ")));
    box.addEventListener("toggle", () => {
      if (!box.open || box.childElementCount > 1) return;
      box.append(el("p", s.coverage(percent.format(candidate.sourceCoverage), percent.format(candidate.targetCoverage))));
      const version = versions[Number(entry.id)];
      const paragraphs = new Map(version.unavailable ? [] : version.paragraphs.map(paragraph => [paragraph.index, paragraph.text]));
      const list = el("ul");
      for (const { from, to, distinctive } of candidate.matches) {
        const quote = paragraphs.get(to.paragraph)?.slice(to.start, to.end);
        const location = s.location(from.paragraph + 1, to.paragraph + 1);
        const item = el("li", distinctive ? `${location} · ${s.distinctive}` : location);
        if (quote) item.append(el("blockquote", quote));
        list.append(item);
      }
      box.append(list);
    });
    return box;
  }

  function inspect(entry: VersionBranchEntry, open: boolean): void {
    nodes.get(inspected ?? "")?.button.removeAttribute("data-inspected");
    inspected = entry.id;
    nodes.get(inspected)?.button.setAttribute("data-inspected", "");
    evidence.hidden = false;
    evidence.open = open;
    evidenceTitle.textContent = `${s.evidence} · ${name(Number(entry.id))}`;
    const time = entry.time.selected;
    const decision = entry.decision.status === "inferred"
      ? `${s.accepted} · ${name(Number(entry.decision.parentId))}` : s.holds[entry.decision.reason];
    const dates = el("ul");
    if (time) dates.append(el("li", `${s.plottedTime} · ${s.timeSources[time.source]} · ${date.format(Date.parse(time.value))}`));
    else dates.append(el("li", s.unknown));
    const uncertainty = entry.time.uncertainty;
    if (time && uncertainty && (entry.decision.status !== "held" || entry.decision.reason !== uncertainty)) {
      dates.append(el("li", s.holds[uncertainty]));
    }
    for (const source of ["archivedAt", "sourceModified", "fileCreatedAt"] as const) {
      const value = entry.time.observed[source];
      if (value && source !== time?.source) dates.append(el("li", `${s.timeSources[source]} · ${value}`));
    }
    if (entry.time.archiveNumber !== undefined) dates.append(el("li", `${s.archiveNumber} · ${entry.time.archiveNumber}`));
    if (entry.duplicateOf !== undefined) dates.append(el("li", `${s.duplicateOf} · ${name(Number(entry.duplicateOf))}`));
    if (entry.sameVisibleText.length) dates.append(el("li", `${s.sameText} · ${entry.sameVisibleText.map(id => name(Number(id))).join(" / ")}`));
    evidenceBody.replaceChildren(el("p", decision), dates, ...entry.candidates.map(candidate => showCandidate(entry, candidate)));
  }

  function refresh(entries: readonly VersionBranchEntry[], paths: ReadonlyMap<string, readonly number[]>): void {
    const layout = layoutSavedVersionTimeline(entries);
    const keys = new Map([...paths].map(([key, path]) => [String(path[path.length - 1]), key]));
    nodes = new Map();
    drawnEdges = [];
    const wasOpen = evidence.open;
    if (!layout.times.length) {
      viewport.replaceChildren(el("p", s.empty, "hint"));
      evidence.hidden = true;
      return;
    }
    const x = (column: number) => firstColumn + column * columnStep;
    const y = (lane: number) => firstLane + lane * laneStep;
    const plot = el("div", undefined, "report-branch-plot");
    const width = x(layout.times.length - 1) + firstColumn;
    const height = y(layout.lanes - 1) + laneStep;
    plot.style.width = `${width}px`;
    plot.style.height = `${height}px`;
    const graph = document.createElementNS(svgNamespace, "svg");
    graph.setAttribute("width", String(width));
    graph.setAttribute("height", String(height));
    graph.setAttribute("aria-hidden", "true");
    plot.append(graph);
    layout.times.forEach((time, column) => {
      const mark = el("span", time ? tick.format(Date.parse(time)) : s.unknown, "report-branch-tick");
      mark.style.left = `${x(column)}px`;
      if (time) mark.title = date.format(Date.parse(time));
      plot.append(mark);
    });
    for (const edge of [...layout.edges].sort((a, b) => Number(a.inferred) - Number(b.inferred))) {
      const start = x(edge.from.column), end = x(edge.to.column);
      const from = y(edge.from.lane), to = y(edge.to.lane);
      const middle = (start + end) / 2;
      const line = document.createElementNS(svgNamespace, "path");
      line.setAttribute("d", `M ${start} ${from} C ${middle} ${from}, ${middle} ${to}, ${end} ${to}`);
      line.classList.add(edge.inferred ? "report-branch-inferred" : "report-branch-possible");
      graph.append(line);
      drawnEdges.push({ edge, line });
    }
    for (const { entry, column, lane } of layout.nodes) {
      const label = name(Number(entry.id));
      const button = el("button", undefined, "report-branch-node");
      button.type = "button";
      button.disabled = disabled;
      button.dataset.version = entry.id;
      button.style.left = `${x(column)}px`;
      button.style.top = `${y(lane)}px`;
      const time = entry.time.selected;
      button.title = [label, time ? `${s.timeSources[time.source]} · ${date.format(Date.parse(time.value))}` : s.unknown].join(" · ");
      button.setAttribute("aria-label", entry.decision.status === "held"
        ? `${button.title} · ${s.holds[entry.decision.reason]}` : button.title);
      const key = keys.get(entry.id);
      if (key) button.setAttribute("aria-pressed", "false");
      if (entry.decision.status === "held") button.dataset.held = "";
      button.append(el("span", undefined, "report-branch-dot"), el("span", label, "report-branch-name"));
      button.onclick = () => {
        if (key) choosePath(key);
        inspect(entry, !key || evidence.open);
      };
      plot.append(button);
      nodes.set(entry.id, { entry, button });
    }
    for (const drawn of drawnEdges.filter(({ edge }) => edge.inferred)) {
      const { from, to } = drawn.edge;
      const after = Number(to.entry.id);
      const button = el("button", undefined, "report-branch-edge-button");
      button.type = "button";
      button.hidden = true;
      button.disabled = disabled;
      // Place the control farther down a split to clear the labels on the source lane.
      const t = from.lane === to.lane ? 0.5 : 0.72;
      const dx = 1.5 * t * (1 - t) + t ** 3;
      const dy = t ** 2 * (3 - 2 * t);
      button.style.left = `${x(from.column) + (x(to.column) - x(from.column)) * dx}px`;
      button.style.top = `${y(from.lane) + (y(to.lane) - y(from.lane)) * dy}px`;
      button.title = `${s.compareInterval} · ${name(Number(from.entry.id))} → ${name(after)}`;
      button.setAttribute("aria-label", button.title);
      button.append(el("span", "›"));
      button.onclick = () => chooseInterval(after);
      plot.append(button);
      drawn.button = button;
    }
    viewport.replaceChildren(plot);
    const previous = nodes.get(inspected ?? "");
    if (previous) {
      inspect(previous.entry, wasOpen);
    } else {
      inspected = undefined;
      evidence.hidden = true;
    }
  }

  function select(path: readonly number[] | undefined, pair?: readonly [number, number]): void {
    const active = new Set(path);
    const endpoint = path?.at(-1);
    for (const [id, { button }] of nodes) {
      button.toggleAttribute("data-path", active.has(Number(id)));
      if (button.hasAttribute("aria-pressed")) button.setAttribute("aria-pressed", String(Number(id) === endpoint));
    }
    for (const { edge, line, button } of drawnEdges) {
      const from = Number(edge.from.entry.id), to = Number(edge.to.entry.id);
      const onPath = edge.inferred && active.has(from) && active.has(to);
      const inInterval = onPath && pair?.[0] === from && pair[1] === to;
      line.classList.toggle("report-branch-active", onPath);
      line.classList.toggle("report-branch-interval", inInterval);
      if (button) {
        button.hidden = !onPath;
        button.setAttribute("aria-pressed", String(inInterval));
      }
    }
  }

  function setDisabled(value: boolean): void {
    if (value === disabled) return;
    disabled = value;
    for (const { button } of nodes.values()) button.disabled = value;
    for (const { button } of drawnEdges) if (button) button.disabled = value;
  }

  return { element, refresh, select, setDisabled };
}
