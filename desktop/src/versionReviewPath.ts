import { compareSavedVersions, type LineageLevel } from "../../src/lib/lineage";
import { inferVersionBranches, type VersionBranchEntry } from "../../src/lib/lineageBranches";
import { SavedReportAnalysis } from "./savedReport";
import type { SavedLineageVersion } from "./shared";

/** The review input is one archive history, or a set of local files. */
export function inferSavedVersionTimeline(versions: readonly SavedLineageVersion[], targets: readonly number[]): {
  entries: VersionBranchEntry[];
  paths: Map<string, number[]>;
} {
  const archiveId = versions.every(version => version.source === "archive") ? "archive" : undefined;
  const entries = inferVersionBranches(targets.map(index => ({ ...versions[index], id: String(index), archiveId })));
  const byEndpoint = new Map<string, number[]>();
  for (const entry of entries) {
    if (entry.decision.status === "inferred") {
      const parent = entry.decision.parentId;
      byEndpoint.set(entry.id, [...(byEndpoint.get(parent) ?? [Number(parent)]), Number(entry.id)]);
    }
  }
  return { entries, paths: new Map([...byEndpoint.values()].map(path => [path.join(":"), path])) };
}

/** Build lineage and its comment index from the complete selected ancestry. */
export function compareSavedVersionPath(
  versions: readonly SavedLineageVersion[], path: readonly number[], level: LineageLevel,
): SavedReportAnalysis {
  const selected = path.map((index, position) => {
    const version = versions[index];
    if (!position || version.gapBefore || version.source !== "archive") return version;
    const previous = versions[path[position - 1]];
    if (previous.source !== "archive" || previous.number === undefined || version.number === undefined) return version;
    const first = previous.number;
    const last = version.number;
    // Preserve recorded gaps in this archive even when a snapshot on the edge was skipped.
    const gap = versions.some(snapshot => snapshot.source === "archive" && snapshot.gapBefore &&
      snapshot.number !== undefined && first < snapshot.number && snapshot.number <= last);
    return gap ? { ...version, gapBefore: true } : version;
  });
  return new SavedReportAnalysis(selected, compareSavedVersions(selected, { japaneseWordHints: true }), level);
}
