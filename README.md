# docx-revision-analyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
English | [日本語](./README.ja.md)

Three CLI tools that analyze Word (`.docx`) files edited with Track Changes
enabled:

1. **`docx-revision-chart`** — renders an SVG chart of edit activity over time
2. **`docx-revision-flow`** — for each session of continuous editing, draws
   the document's pages at the start and end of the session as a schematic,
   colors fine-grained edits (green), bulk insertions/replacements (orange) and
   deletions (red), and connects each paragraph/figure/table across the two
3. **`docx-ai-suspicion-score`** — scores how likely it is that a chunk of
   text was pasted in from an external app (e.g. an AI writing tool) rather
   than typed and reviewed inside Word, on a 0–100 scale

They're published as an npm package and also distributed as single, dependency-free executables built with
[Bun](https://bun.sh) (no Node.js required). Licensed under MIT.

> **Prerequisite**: all tools require a `.docx` file that was authored/edited
> with Word's "Track Changes" turned on (Review tab → Track Changes). Files
> edited with Track Changes off don't retain insertion (`w:ins`) / deletion
> (`w:del`) markup, so there's nothing to analyze (the tools exit with an
> error explaining why).

> **What this is for**: a tool for writers to look back on how their own document was written — not a way to police
> others. Tracked changes in a `.docx` saved on a computer are easy to remove or rewrite: turning Track Changes off,
> accepting all changes, or editing the XML to change `w:date` / `w:author`. Other apps and converters also write
> files differently from Word. So the history these tools show can be incomplete or altered, and nothing they output
> is proof of how a document was written. See "[Integrity notes](#integrity-notes)".

---

## Installation

### From npm (Node.js 18+)

```bash
npm install -g docx-revision-analyzer
```

This installs three commands: `docx-revision-chart`, `docx-revision-flow`, and `docx-ai-suspicion-score`.

### Single executables (no Node.js needed)

Download the file for your OS / CPU from [GitHub Releases](https://github.com/kubohiroya/docx-revision-analyzer/releases) and unzip it.

| File | Contents |
|---|---|
| `docx-revision-<chart\|flow>-<os>-<cpu>.zip`, `docx-ai-suspicion-score-<os>-<cpu>.zip` | Standalone executable for each tool (`linux-x64` / `linux-arm64` / `macos-x64` / `macos-arm64` / `windows-x64`) |
| `docx-revision-<chart\|flow>-macos-arm64.app.zip` | Drag-and-drop app for macOS (Apple Silicon): drop `.docx` files onto its icon in Finder and the SVG is written next to them |

The drag-and-drop apps are only ad-hoc signed, so the first time, right-click the app in Finder and choose "Open".
On Windows you can drop `.docx` files straight onto the `.exe`.

For building from source, making the single executables and drag-and-drop apps, the test fixtures, and the directory
layout, see [INSTALL.md](./INSTALL.md).

---

## 1. `docx-revision-chart` — SVG chart of the revision history

```bash
node dist/cli/chart.js <input.docx> [input2.docx ...] [options]
```

Multiple input files can be given at once (each is processed independently,
and a failure on one doesn't stop the rest) — mainly useful for Windows,
where dropping several files onto the executable in Explorer passes them
all as separate arguments to one process invocation. `-o/--output` can only
be used with a single input file.

| Option | Description | Default |
|---|---|---|
| `-o, --output <file.svg>` | Output SVG path (single-file use only) | `<input filename>.svg` |
| `-b, --bucket <spec>` | Time bucket granularity: `auto`\|`second`\|`minute`\|`hour`\|`day`, or a number of seconds | `auto` |
| `-p, --gap-threshold <hours>` | Threshold (in hours) used to tell "periods of continuous editing" apart from "periods with no activity". When set, renders a separate, detailed chart per period and lays them out horizontally (see below) | unset (renders one single chart) |
| `-w, --width <px>` | Image width | `1100` (auto-computed from content when `-p` is used) |
| `-H, --height <px>` | Image height | `550` |
| `-t, --title <text>` | Chart title | `Revision history: <file name> (last modified <last-modified time>)` |
| `--bulk-chars <n>` | Treat insertions by the same author at the same time totalling at least this many characters as a bulk insertion (same rule as `docx-revision-flow`) | `150` |
| `--json [file.json]` | Also write the analysis (insertion windows and their features) as JSON. See "Analysis JSON" below | off (`<output>.json` when given without a name) |
| `--rules <file>` | Rules file that assigns highlight levels to insertion windows (replaces `--bulk-chars`). See "Highlight rules" below | none |
| `--window-seconds <s>` / `--window-chars <n>` / `--window-paras <n>` | Insertion windows: time window Δt, and distance in the document in characters / paragraphs. Overrides the rules' `window` | from the rules |
| `--preserve-history` | If the document removes personal information (tracked-change authors and dates) on save, remove that setting, turn Track Changes on, and save it in place (the original is kept as a backup; fails if the document is open). `--preserveHistory` also works. See "When timestamps are missing" below | off |
| `--check-history-settings` | Don't draw a chart; only check the settings and print `ok` or `needs-fix` on the first line, followed by the confirmation text when `needs-fix` | off |
| `--drop` | Desktop drag-and-drop launch mode. When `-o` isn't given, names each output `<same directory as its input>/<filename>-<that input file's last-modified time>.svg` instead of the plain `<filename>.svg` default. Intended for the macOS Finder droplet or a direct Windows Explorer drop (see [INSTALL.md](./INSTALL.md)) (or any other double-click/drag-drop launch with no terminal attached). On Windows, also shows a native message box summarizing the result | off |
| `--lang <en\|ja>` | Display language for messages and the figure (see [Language](#language)) | OS locale |

### How to read the chart

![Sample docx-revision-chart output](./fixtures/chart-demo.png)

This is the output of the command below for `fixtures/chart-demo.docx` (three writing sessions over two days: typed
in small steps while trimming a draft in deletions of varying size here and there, with one bulk paste at the start of
the second session, 14:50). With `-p 2` the chart is split at idle gaps longer than two hours, and the length of each
gap (4.5 h, 17 h) is shown between the periods.

```bash
node dist/cli/chart.js fixtures/chart-demo.docx -p 2
```


- **X-axis**: time (aggregated per bucket)
- **Upward bars**: characters inserted within that bucket, stacked from the bottom in this order (same rules as
  `docx-revision-flow`):
  - **Green (fine-grained editing)**: insertions that are neither bulk nor moved
  - **Orange (bulk insertion)**: insertions by the same author at the same time totalling at least `--bulk-chars`
    characters — Word records a pasted passage as one insertion per paragraph, all with the same timestamp
  - **Blue (moves/reordering)**: insertions (20+ characters) matching text deleted elsewhere in the document, e.g.
    copy + paste + delete; shown (and listed in the legend) only when present
- **Downward bars (red)**: characters deleted within that bucket (shown as an absolute value)
- **Line (dark gray, right axis)**: the document's total character count at the end of each bucket (the estimated character count before tracking began, plus the running net change since then)

Changes Word recorded as moves (`w:moveFrom` / `w:moveTo`) don't change the character count and aren't included in
the bars.

```bash
node dist/cli/chart.js fixtures/natural-writing.docx -o out.svg
```

Sample outputs are bundled under `fixtures/*.svg` (and `fixtures/*.png` for a
quick visual preview).

### The `-p` option: a detailed view per editing period

Passing `-p <hours>` treats any gap between two consecutive revision events
that exceeds that threshold as an "idle period", and splits the timeline
there. **Each period of continuous editing (a "session") gets its own
independent chart, laid out side by side.** Because each session is bucketed
independently, short bursts of editing that would otherwise be flattened into
a handful of coarse buckets in a single multi-day chart become visible in
detail, session by session.

Between two sessions, a narrow dashed gap — just wide enough for a label like
`"12 h"` (about 4 characters) — is inserted to show how long that idle period
was. The vertical scale for inserted/deleted characters, and the scale for
the total-character-count line (right axis), are shared across all sessions
so activity levels remain comparable. The total-character-count line is kept
as a separate polyline per session rather than one continuous line, so it
never visually bridges a gap where no time actually elapsed.

```bash
node dist/cli/chart.js fixtures/multi-session.docx -p 12 -o out.svg
```

`fixtures/multi-session.docx` (writing spread across 3 days, with 29-hour and
20.5-hour idle gaps) is bundled as a test case; its `-p 12` output is at
`fixtures/multi-session-p12.svg` / `.png`. Compare it against the single-chart
rendering (`fixtures/multi-session-single.svg` / `.png`, generated without
`-p`) to see how much detail session-splitting recovers.

---

## 2. `docx-revision-flow` — edit flow across editing sessions

Splits the Track Changes history into sessions of continuous editing and draws, left to right in time order, the
document at the start of the first session and at the end of every session as columns of page thumbnails (a
schematic, not Word's real layout). Between two columns, bands connect each paragraph/figure/table as it changed during
that session. Nothing changes between sessions, so each session's end column doubles as the next session's start.

```bash
node dist/cli/flow.js fixtures/flow-demo.docx -o flow.svg

# Limit the period and split sessions at idle gaps longer than 2 hours
node dist/cli/flow.js report.docx --from 2026-06-01 --to "2026-06-02 18:00" -p 2
```

![Sample docx-revision-flow output](./fixtures/flow-demo.png)

| Option | Description | Default |
|---|---|---|
| `-o, --output <file.svg>` | Output SVG path | `<input name>-flow.svg` |
| `-p, --gap-threshold <hours>` | Start a new session after an idle gap longer than this | `1` |
| `--from <datetime>` / `--to <datetime>` | Period to analyze (local time, `2026-06-01` or `"2026-06-01 09:30"`; a date-only `--to` includes that whole day) | everything |
| `--bulk-chars <n>` | Treat insertions made at the same time totalling at least this many characters as a bulk insertion | `150` |
| `--json [file.json]` | Also write the analysis (insertion windows and their features) as JSON. See "Analysis JSON" below | off (`<output>.json` when given without a name) |
| `--rules <file>` | Rules file that assigns highlight levels to insertion windows (replaces `--bulk-chars`). See "Highlight rules" below | none |
| `--window-seconds <s>` / `--window-chars <n>` / `--window-paras <n>` | Insertion windows: time window Δt, and distance in the document in characters / paragraphs. Overrides the rules' `window` | from the rules |
| `--page-width <px>` | Width of each page thumbnail | `150` |
| `--slope-width <px>` | Width of the band area between two thumbnail columns | `72` |
| `-t, --title <text>` | Title | `Edit flow: <file name> (last modified <last-modified time>)` |
| `--drop` / `--preserve-history` / `--check-history-settings` | Same as `docx-revision-chart`; `--drop` writes `<file name>-flow-<last-modified time>.svg` | |
| `--lang <en\|ja>` | Display language for messages and the figure (see [Language](#language)) | OS locale |

If there are no timestamped revisions, or none in the requested period, no SVG is written and the tool exits with an error.

### How to read it

- **Each column**: the first is the document just before session 1's first change; the others are the document at the
  end of each session (earlier changes applied, later ones undone), page by page. Gray bars are body lines, dark bars
  headings, bluish bars tables, crossed boxes figures. A column's heading also shows the idle gap that follows it.
- **Bands between two columns** show one session: they connect each paragraph/figure (a whole table counts as one) from its start position
  and height to its end position and height: a band that widens gained content, one that narrows lost it. A red band
  that ends in the middle was deleted during the session; a band that grows out of the middle was added (orange if
  bulk-inserted, green otherwise); gray means unchanged.
- **Blue (moves/reordering)**: a blue band runs from where the text was to where it went; where the order changed it
  crosses the other bands. Paragraphs that received moved text are painted blue.
- **Marks at a page's right edge**: paragraphs/figures deleted (red) or moved away (blue) in the next session.
- **Green (fine-grained editing)**: paragraphs that received many small insertions/deletions in the session.
  Intensity is the larger of (fine-edit characters ÷ paragraph length) and (number of fine edits ÷ 8), capped at 1.
- **Orange (bulk insertion / replacement)**: the share of the paragraph that arrived as a bulk insertion in the
  session. Intensity is bulk-inserted characters ÷ paragraph length.
- Green and orange are painted on the session's end column. A paragraph with both is filled with the stronger color, and
  the other is shown as a thin bar on its left.
- **Below each band area**: session start/end (month/day hour:minute), total inserted/deleted characters, and how much
  of the insertion was bulk.

### How bulk insertions are detected

When you paste several paragraphs, Word records each paragraph as a separate insertion (`w:ins`) with the same
timestamp. Insertions by the same author with the same timestamp are therefore grouped, and a group totalling at least
`--bulk-chars` characters counts as a bulk insertion. Deletions by the same author at the same timestamp as a bulk
insertion are treated as the replaced text and not counted as fine-grained editing.

### How moves and reordering are detected

How Word records reordering depends on how it was done:

- **Cut + paste, or drag and drop**: with "Track moves" on (the default), Word records a move
  (`w:moveFrom` / `w:moveTo`); source and destination are paired by the range name Word assigns.
- **Copy + paste + delete**, and cut + paste that Word didn't record as a move: recorded as an insertion plus a
  deletion. If inserted text (20+ characters) matches text deleted anywhere in the document, it's treated as
  reordering/duplication within the document, and if the matching deletion is in the same session the two are
  connected with a blue band.

Either way, reordered text counts neither as a bulk insertion (orange) nor as fine-grained editing (green). The caption
shows it as "N chars moved/reordered" (the total inserted/deleted characters still include reordering recorded as
insertion + deletion).

### Limitations

- Pages are a schematic, not Word's actual layout. Characters per line come from the page width and font size, lines
  per page from the line pitch, calibrated so the final state matches the page count Word saved (`docProps/app.xml`).
  Page breaks for earlier sessions reuse that calibration and are approximate.
- Word keeps no record when you delete text you inserted yourself, so text written and deleted within a session is
  not counted.
- Word often records timestamps to the minute, so typing more than `--bulk-chars` characters within one minute can
  also be classified as a bulk insertion.
- Formatting-only changes are not colored. Reordering recorded as insertion + deletion is detected only when the text
  matches exactly (ignoring whitespace); anything edited after pasting counts as editing, not reordering. An added figure with no text is
  shown as "added" (green) in the bands, but its paragraph isn't painted. Only the body
  (`word/document.xml`) is analyzed — footnotes, headers, and text boxes are not.
- Bulk-insertion detection is a heuristic, not evidence of misconduct (same caveat as `docx-ai-suspicion-score`).

---

## 3. `docx-ai-suspicion-score` — AI-misuse suspicion score

```bash
node dist/cli/score.js <input.docx> [options]
```

| Option | Description | Default |
|---|---|---|
| `-o, --output <file.json>` | Where to write the result (stdout if omitted) | - |
| `--min-chars <n>` | Insertions smaller than this are always ignored (avoids false positives) | `20` |
| `--burst-low <n>` | Character-count threshold below which the burst score is 0 | `20` |
| `--burst-high <n>` | Character-count threshold above which the burst score is 100 | `300` |
| `--rate-low <cps>` | Insertion-speed threshold (chars/sec) below which the rate score is 0 | `8` |
| `--rate-high <cps>` | Insertion-speed threshold (chars/sec) above which the rate score is 100 | `40` |
| `--max-weight <0-1>` | Weight given to the single most suspicious event in the overall score (the rest is a character-weighted average) | `0.6` |
| `--pretty` | Pretty-print the JSON output | off |
| `--lang <en\|ja>` | Display language for messages and the figure (see [Language](#language)) | OS locale |

```bash
node dist/cli/score.js fixtures/suspicious-paste.docx --pretty
```

### How the score is computed

Word records every inserted run of text as a `w:ins` element, timestamped to
the second. **When someone types and reviews their own text through Word's
UI, a single insertion event rarely grows by more than a few dozen
characters at once.** By contrast, **pasting in a finished passage written
externally (e.g. by an AI writing tool) typically adds several hundred
characters in a single insertion event, within one or two seconds.** This
tool looks for exactly that difference.

For each insertion event, two sub-scores (0–100) are computed, and the larger
of the two becomes that event's suspicion score:

- **Burst score**: how large the event is by itself (0 at or below
  `--burst-low` characters, 100 at or above `--burst-high`, linearly
  interpolated in between)
- **Rate score**: the implied insertion speed (characters/second) since the
  previous insertion event (0 at or below `--rate-low` cps, 100 at or above
  `--rate-high` cps, linearly interpolated)

Insertions smaller than `--min-chars` are always scored 0, to avoid false
positives from ordinary small edits.

The document-level score is a weighted blend:

```
score = max_weight × (score of the single most suspicious event)
      + (1 − max_weight) × (character-weighted average score across all events)
```

with `max_weight = 0.6` by default. Using only the maximum makes the score
too sensitive to a single false positive; using only the average lets one
large paste get diluted by many ordinary edits. Blending the two avoids both
failure modes.

`riskLevel` is a rough guide: `low` (0–19) / `medium` (20–44) / `high`
(45–74) / `very_high` (75–100).

### ⚠️ Limitations — please read before relying on this

This score is a **statistical heuristic — it is not proof of misconduct**.
It can score legitimately high in cases such as:

- Very fast typists, or people using voice dictation
- Someone who drafted in a separate notes app or editor and pasted the
  result into Word (their own writing can still look like "one big instant
  insertion" the moment it's pasted)
- Pasting a large table or a mechanically-generated reference list

Conversely, it can miss real cases, such as:

- AI-generated text that was slowly copied by hand, character by character
- A pasted passage that was subsequently reworked heavily enough to break it
  into many small insertion events

For these reasons, **treat this score only as a signal to look into further,
and always have a human (e.g. an instructor) make the final judgment call in
context.** It is strongly recommended that this never be used to
automatically penalize a student or author.

Other technical constraints:

- Documents saved with Word's "Remove personal information from file properties on save"
  option lose the `w:date` (and author) of every tracked change, so they cannot be analyzed.
  The tools report how many undated revisions were found in that case.
- `w:date` timestamps have one-second resolution, so sub-second gaps can't be
  distinguished.
- `w:moveFrom` / `w:moveTo` (drag-and-drop moves within the same document)
  are not treated specially in this version; depending on the Word version
  and settings, a move may be counted the same as a plain insertion.
- By default, only `word/document.xml` (the document body) is analyzed;
  Track Changes inside headers, footers, comments, or footnotes are not
  included.

---

## Language

Messages, help, and the figures are shown in English or Japanese. The language is chosen in this order:

1. `--lang en` / `--lang ja`
2. `lang` in the settings file (see below)
3. The `DOCX_REVISION_LANG` environment variable
4. The `LC_ALL` / `LC_MESSAGES` / `LANG` environment variables
5. The OS language setting (on macOS, System Settings > General > Language & Region)

Anything other than Japanese falls back to English. The macOS drag-and-drop apps follow the macOS language setting.

---

## Settings file

To change the defaults, put a YAML file named after the tool (`docx-revision-chart.yml`,
`docx-revision-flow.yml`, or `docx-ai-suspicion-score.yml`; `.yaml` also works) in the same folder as the executable.
For the macOS drag-and-drop apps, put it next to the `.app`; for an npm installation, next to the installed command.

```yaml
# docx-revision-flow.yml
gap-threshold: 2
bulk-chars: 200
page-width: 120
slope-width: 60
title: "Research proposal: edit flow"
output: out/flow.svg
lang: en
```

- Keys are the tool's long option names without `--` (for example `output`, `gap-threshold`, `bulk-chars`,
  `page-width`, `slope-width`, `title`, `bucket`, `from`, `to`, `lang`; for `docx-ai-suspicion-score`, `min-chars`,
  `burst-low`, `pretty`, ...). Flags such as `pretty` take `true` / `false`.
- Options given on the command line take precedence over the settings file, which takes precedence over the
  built-in defaults.
- A relative `output` or `rules` path is resolved against the folder containing the settings file. `output` is ignored when
  several files are processed at once.
- Unknown keys and invalid values are reported as warnings and ignored; a file that isn't valid YAML is an error.
  The tool prints which settings file it used.

---

## When timestamps are missing (`--preserve-history`)

"Remove personal information from file properties on save" is a per-document setting, stored as
`<w:removePersonalInformation/>` / `<w:removeDateAndTime/>` in `word/settings.xml`. While it is on, Word strips
the author and date of every tracked change each time the document is saved. Word for Windows can turn it off under
File > Options > Trust Center > Trust Center Settings > Privacy Options, but Word for macOS has no such option in
Settings > Security.

Before analyzing, `docx-revision-chart` and `docx-revision-flow` check this setting and, if it is on (documents that merely have Track Changes off are not flagged):

- **CLI:** when run from a terminal, asks `[y/N]` whether to perform the `--preserve-history` rewrite
  (Enter alone or Ctrl+D means No); when input isn't interactive (e.g. piped), it only prints a warning.
  Run with `--preserve-history` to skip the question and remove the personal-information setting, turn Track
  Changes (`<w:trackRevisions/>`) on, and save the file in place. The original is kept as `<name>.backup-<YYYYMMDD-HHMMSS>.docx` in the same folder.
- **macOS droplet:** shows an OK/Cancel dialog asking whether to enable saving editor names and edit times; OK reruns
  with `--preserve-history`, Cancel analyzes the file without changing it.
- **Windows `--drop`:** shows the same OK/Cancel message box.

Before rewriting, it checks that the document isn't open elsewhere (Word's `~$…` owner file; on Windows, whether the
file can be opened for writing; on macOS/Linux, `lsof`) and fails without touching the file if it is.
`--check-history-settings` only reports the settings (used internally by the droplet).

Edits made after the rewrite are timestamped; timestamps already removed cannot be recovered.

---

## Analysis JSON (`--json`)

With `--json`, `docx-revision-chart` and `docx-revision-flow` also write the analysis as JSON. It contains
*insertion windows*: insertions made within a short time (`seconds`, from the window's first insertion)
and close together in the document (`chars`, and `paras` if given), measured in the final document. The window
settings come from the rules (see below; by default, the same author and the same timestamp) and can be
overridden with `--window-*`. Moves and reordering are excluded. Each window has these features:

| Feature | Meaning |
|---|---|
| `insertedChars` | Total inserted characters |
| `durationSec` / `cps` | Time span (seconds) / insertion speed (chars per second). The span adds the estimated timestamp resolution (`timeResolutionSec`: 60 when Word recorded minutes only; at least 1 second), since an insertion may have happened anywhere within it |
| `spanChars` / `spanParas` | Extent in the final document (characters / paragraphs) |
| `maxSingleInsert` | Largest single `w:ins` (characters) |
| `paraCount` | Number of paragraphs the insertions touch |
| `postEditRatio` | Characters inserted/deleted in that range after the window ÷ `insertedChars` |
| `precededByDeletion` / `precedingDeletedChars` | Whether text in that range was deleted just before or during the window (a replacement), and how much |
| `insertCount` | Number of insertions |

The features don't judge anything by themselves; thresholds are applied separately.

### Annotating parts of the figures (`--annotations`)

`--annotations <file>` (YAML or JSON) adds a mouse-over explanation (`<title>`) and/or a click-through link
(`<a href>`, http/https only) to parts of the SVG. Each part has a key, listed in the `--json` output as
`figureTargets`: chart bars `chart:bar:<category>:<bucket start>`, flow paragraphs `flow:para:<column>:<paragraph>`,
bands `flow:band:<session>:<unit>`, move bands `flow:move:…`, captions `flow:caption:<session>`.

```yaml
annotations:
  - target: "flow:para:2:9"
    tooltip: { en: "Rewritten after feedback", ja: "指摘を受けて書き直した" }
    href: "https://example.com/notes#p9"
```

The desktop app's extensions can add the same annotations (see `desktop/EXTENSIONS.md`). Without annotations the SVG
is unchanged. Library: `resolveAnnotations`, `chartTargets` / `sessionedChartTargets` / `flowTargets`, and the
renderers' `annotations` option.

### Highlight rules (`--rules`)

A rules file (YAML or JSON) assigns *levels* to insertion windows based on their features. Levels are checked
from the top, and the first match is used. Insertions in a window with a level are drawn as bulk insertions
(orange), and the JSON output records each window's `level` and the rules' `ruleSet`; with `--rules`, the
figures also show the `ruleSet` at the bottom right. See [examples/rules.example.yml](examples/rules.example.yml).

```yaml
ruleSet: example-v1        # identifier recorded in the outputs
window: { seconds: 60, chars: 2000 }   # optional: also paras, byAuthor
levels:
  - id: level-2
    label: { ja: 大量の一括挿入, en: Large bulk insertion with little editing afterwards }
    color: "#9A3412"
    when: { all: [ { insertedChars: { gte: 800 } }, { postEditRatio: { lt: 0.05 } } ] }
  - id: level-1
    label: { ja: 一括挿入文字数過多, en: Large bulk insertion }
    color: "#EA6C00"
    when: { insertedChars: { gte: 300 } }
```

- Comparisons: `gte` / `gt` / `lte` / `lt` / `eq` (several in one mapping must all hold); `{ precededByDeletion: true }`
  is short for `eq`. Combine with `all: [...]`, `any: [...]`, `not: ...`. Any number of levels.
- Without `--rules`, the default rule is used: window = same author and same timestamp, one level
  `insertedChars >= --bulk-chars`. This gives exactly the same result as before.
- `--bulk-chars` is still used to decide which large deletions don't count as fine-grained edits in
  `docx-revision-flow`.
- Each level is drawn in its own color and replaces "bulk insertion" in the legends. In the chart, levels are stacked
  between fine-grained edits and moves (lower levels below). In `docx-revision-flow`, a paragraph is filled with the
  highest level that applies to it; other categories are shown by the thin strip at its left.
- So that levels aren't told apart by color alone, each level also gets a pattern (`hatch`, `cross`, `dots` in order;
  set `pattern:` on a level to choose, or `pattern: none`). A warning is shown when a level's color has less than 3:1
  contrast against white.

## Integrity notes

With `--json`, the output also includes `integrity`: a few informational observations about the file, each with
`observed` (`true` if it differs from what Word normally writes, `false` if not, `null` if the file doesn't record
the needed information) and a message. They are **not used for any judgment or warning** and say nothing about
whether the file was altered — Word versions, other apps and converters produce the same differences.

| Item | What is observed |
|---|---|
| `settings` | Whether Track Changes was on when saved, and whether Word removes authors/dates on save |
| `undated` | Changes without a date or author; number of authors |
| `duplicateIds` | Change ids (`w:id`) used more than once (Word gives each change its own id) |
| `idOrder` | Places where change ids decrease in document order (Word renumbers them in document order on save) |
| `futureDates` | Change dates in the future, or later than the file's last-modified time (`docProps/core.xml`) |
| `beforeCreated` | Change dates more than a day before the file's creation time (copying tracked text from another file also does this) |
| `deletedBeforeInserted` | Text whose deletion is dated before its insertion |
| `rsids` | Editing-session ids (rsid) used in the text but missing from the list in `settings.xml` |
| `application` | The app that last saved the file and the recorded total editing time (`docProps/app.xml`) |

The desktop app shows the same list under the Highlights tab. Library: `checkIntegrity(bytes)`.

## Desktop app (preview)

[`desktop/`](desktop/) contains a desktop app (Electron) for people who don't use the command line: open or drop a
.docx to see the chart, the flow and the highlighted insertions. Everything runs locally and the app blocks network
access. Installers aren't published yet.

## Using it as a library

`src/index.ts` re-exports the main functions and types (`extractRevisionsFromFile`,
`computeSuspicionScore`, `renderRevisionChart`, etc.), so beyond the CLIs, the
package can also be used as a library from another Node.js / Bun project via
`import { ... } from "docx-revision-analyzer"`.

### In the browser (`docx-revision-analyzer/core`)

The analysis and drawing code takes the .docx as bytes (`ArrayBuffer` / `Uint8Array`) and doesn't use the file
system or child processes, so it also runs in a browser, a WebView or a sandboxed worker. Import it from
`docx-revision-analyzer/core`; the default entry adds the Node.js-only helpers (`extractRevisionsFromFile`,
`enableHistoryPreservation` that rewrites a file in place, `buildDropOutputPath`, and reading the macOS language
setting).

```ts
import { extractRevisions, parseDocxLayout, buildFlow, renderFlowSvg, preserveHistoryInDocx } from "docx-revision-analyzer/core";

const bytes = new Uint8Array(await file.arrayBuffer());
const data = await extractRevisions(bytes);
const svg = renderFlowSvg(buildFlow(await parseDocxLayout(bytes), { gapThresholdHours: 1, bulkChars: 150 }));
const { output } = await preserveHistoryInDocx(bytes); // the fixed .docx as bytes, if it needed fixing
```

`npm test` bundles the core for the browser with esbuild (failing if it imports a Node.js module) and runs it on
the fixtures in a sandbox without `process` / `require` / `Buffer`, checking that the results match Node.js.

### Classifiers

Highlights are produced by a pipeline of *classifiers*. A classifier (`{ id, version, categories?, classify(ctx) }`)
receives the analysis (`ctx.positioned` events, `ctx.windowsFor(options)` insertion windows with features,
`ctx.sessions` time ranges) and returns highlights: category id, target insertions (`w:id`), time range, document
range, features and a reason. `runClassifiers` runs them in order and, for each insertion, keeps the highlight
whose category has the highest priority (ties: the earlier classifier). The built-in pipeline
(`defaultClassifiers(rules)`) is `builtin.relocation` (reordering/duplication → moves) followed by `builtin.rules`
(the rule levels). The JSON output records the classifiers' ids and versions (`classifiers`) and every highlight
with its reason (`highlights`).

### Color categories

The colors in the figures come from a `CategoryRegistry` (`{ id, role, color, label, priority, pattern }`). The
built-in categories are fine-grained edits, bulk insertion, moves, deletions and unchanged; register more
`highlight` categories and pass the registry as `categories` to `renderRevisionChart` /
`renderSessionedRevisionChart` / `renderFlowSvg`. Insertions are tied to a category through
`RevisionEvent.category` (chart) or `FlowOptions.highlightOf` (flow).

### Where each edit happened

`extractRevisionPositions` returns one event per insertion/deletion (`w:id`) with its time, author, character
count, text, and position: paragraph index, offset within the paragraph, and offset from the start of the
document, both in the final document (all changes accepted) and at the moment of the edit (right after an
insertion / right before a deletion). Offsets index into `finalDocumentText(model)` / `documentTextAt(model, date)`
(paragraphs joined with `"\n"`, counted in UTF-16 code units). Moves (`w:moveFrom` / `w:moveTo`) are included with
`move: true` and paired by `moveName`. `attachRevisionPositions` sets `position` on matching `RevisionEvent`s.

```ts
import { readFileSync } from "fs";
import { parseDocxLayout, extractRevisionPositions } from "docx-revision-analyzer";

const model = await parseDocxLayout(readFileSync("report.docx"));
for (const e of extractRevisionPositions(model)) {
  console.log(e.type, e.date, e.chars, e.final.start.paraIndex, e.final.start.offsetInPara);
}
```

---

## Contributing

Issues and pull requests are welcome — please use
[GitHub Issues](https://github.com/kubohiroya/docx-revision-analyzer/issues)
for bug reports and feature requests.

---

## License

[MIT](./LICENSE) © 2026 Hiroya Kubo
