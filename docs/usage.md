# Usage

English | [日本語](./usage.ja.md) · [Back to the README](../README.md)

Each command's options and how to read the figures it draws. For running it together with Word and Teams, see the [operations guide](operations-guide.md); for the JSON output and highlight rules, see [advanced usage](advanced.md).

**Contents**

- [1. `docx-revision-chart` — SVG chart of the revision history](#1-docx-revision-chart--svg-chart-of-the-revision-history)
- [2. `docx-revision-flow` — edit flow across editing sessions](#2-docx-revision-flow--edit-flow-across-editing-sessions)
- [3. `docx-ai-suspicion-score` — AI-misuse suspicion score](#3-docx-ai-suspicion-score--ai-misuse-suspicion-score)
- [4. `docx-revision-snapshot` — archive and chart submissions over time](#4-docx-revision-snapshot--archive-and-chart-submissions-over-time)
- [Detecting traces of tampering](#detecting-traces-of-tampering)
- [When timestamps are missing (`--preserve-history`)](#when-timestamps-are-missing---preserve-history)
- [Settings file](#settings-file)
- [Language](#language)
- [Desktop app (preview)](#desktop-app-preview)

## 1. `docx-revision-chart` — SVG chart of the revision history

```bash
docx-revision-chart <input.docx> [input2.docx ...] [options]
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
| `--json [file.json]` | Also write the analysis (insertion windows and their features) as JSON. See [Analysis JSON](advanced.md#analysis-json---json) | off (`<output>.json` when given without a name) |
| `--rules <file>` | Rules file that assigns highlight levels to insertion windows (replaces `--bulk-chars`). See [Highlight rules](advanced.md#highlight-rules---rules) | none |
| `--window-seconds <s>` / `--window-chars <n>` / `--window-paras <n>` | Insertion windows: time window Δt, and distance in the document in characters / paragraphs. Overrides the rules' `window` | from the rules |
| `--template <file.docx>` | The template you handed out; each file is checked against it for traces of tampering ([Detecting traces of tampering](#detecting-traces-of-tampering)) | none |
| `--no-tamper-check` | Don't check for traces of tampering (no warning, no `-tampered`) | off |
| `--preserve-history` | If the document removes personal information (tracked-change authors and dates) on save, remove that setting, turn Track Changes on, and save it in place (the original is kept as a backup; fails if the document is open). `--preserveHistory` also works. See [When timestamps are missing](#when-timestamps-are-missing---preserve-history) | off |
| `--check-history-settings` | Don't draw a chart; only check the settings and print `ok` or `needs-fix` on the first line, followed by the confirmation text when `needs-fix` | off |
| `--lock` | Don't draw a chart; lock Track Changes in the given template and overwrite it (the original is kept as a backup). See [Locking with these tools](operations-guide.md#2-lock-tracking-optional) | off |
| `--drop` | Desktop drag-and-drop launch mode. When `-o` isn't given, names each output `<same directory as its input>/<filename>-<that input file's last-modified time>.svg` instead of the plain `<filename>.svg` default. Intended for the macOS Finder droplet or a direct Windows Explorer drop (see [Installation](installation.md#finder-drag-and-drop-app-macos-only)) (or any other double-click/drag-drop launch with no terminal attached). On Windows, also shows a native message box summarizing the result | off |
| `--lang <en\|ja>` | Display language for messages and the figure (see [Language](#language)) | OS locale |

### How to read the chart

![Sample docx-revision-chart output](../fixtures/chart-demo.png)

This is the output of the command below for `fixtures/chart-demo.docx` (three writing sessions over two days: typed
in small steps while trimming a draft in deletions of varying size here and there, with one bulk paste at the start of
the second session, 14:50). With `-p 2` the chart is split at idle gaps longer than two hours, and the length of each
gap (4.5 h, 17 h) is shown between the periods.

```bash
docx-revision-chart fixtures/chart-demo.docx -p 2
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
docx-revision-chart fixtures/natural-writing.docx -o out.svg
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
docx-revision-chart fixtures/multi-session.docx -p 12 -o out.svg
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
docx-revision-flow fixtures/flow-demo.docx -o flow.svg

# Limit the period and split sessions at idle gaps longer than 2 hours
docx-revision-flow report.docx --from 2026-06-01 --to "2026-06-02 18:00" -p 2
```

![Sample docx-revision-flow output](../fixtures/flow-demo.png)

| Option | Description | Default |
|---|---|---|
| `-o, --output <file.svg>` | Output SVG path | `<input name>-flow.svg` |
| `-p, --gap-threshold <hours>` | Start a new session after an idle gap longer than this | `1` |
| `--from <datetime>` / `--to <datetime>` | Period to analyze (local time, `2026-06-01` or `"2026-06-01 09:30"`; a date-only `--to` includes that whole day) | everything |
| `--bulk-chars <n>` | Treat insertions made at the same time totalling at least this many characters as a bulk insertion | `150` |
| `--json [file.json]` | Also write the analysis (insertion windows and their features) as JSON. See [Analysis JSON](advanced.md#analysis-json---json) | off (`<output>.json` when given without a name) |
| `--rules <file>` | Rules file that assigns highlight levels to insertion windows (replaces `--bulk-chars`). See [Highlight rules](advanced.md#highlight-rules---rules) | none |
| `--window-seconds <s>` / `--window-chars <n>` / `--window-paras <n>` | Insertion windows: time window Δt, and distance in the document in characters / paragraphs. Overrides the rules' `window` | from the rules |
| `--template <file.docx>` | The template you handed out; each file is checked against it for traces of tampering ([Detecting traces of tampering](#detecting-traces-of-tampering)) | none |
| `--no-tamper-check` | Don't check for traces of tampering (no warning, no `-tampered`) | off |
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
docx-ai-suspicion-score <input.docx> [options]
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
docx-ai-suspicion-score fixtures/suspicious-paste.docx --pretty
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

## 4. `docx-revision-snapshot` — archive and chart submissions over time

Archives the submissions (snapshots) of a long-running document, such as a thesis handed in at the end of each
sprint, and analyzes them as one history. For the workflow, see
[Long-running writing](operations-guide.md#long-running-writing-such-as-a-thesis-submissions-at-each-sprint).

```bash
# Sprint 1: pass the folder of submissions (the first one is checked against the template you handed out)
docx-revision-snapshot archive/ submissions-sprint1/ --template template.docx

# Later sprints: pass that sprint's folder to the same archive
docx-revision-snapshot archive/ submissions-sprint2/

# Redraw the figures and the index from what is archived
docx-revision-snapshot archive/
```

For each submission it:

1. **Archives** it in the document's folder as `<name>-s01.docx`, `<name>-s02.docx`, … and records each one
   (source, time archived, SHA-256) in `snapshots.json`. A file identical to the previous submission is not archived
   again.
2. **Checks** it against the previous submission for [traces of tampering](#detecting-traces-of-tampering) (the
   first one against `--template`). Because the previous submission is the reference, `untrackedText` (text typed
   outside Track Changes) is checked even without a template. The lock's password is not compared, since the
   teacher re-locking the file to accept changes changes its hash. It also reports `datedBeforePrevious` when changes
   first seen in a submission are dated before the last change of the previous one (a file edited elsewhere swapped
   in, or rewritten dates).
3. **Draws each submission's figures**: a chart (`<name>-s01.svg`) and a flow (`<name>-s01-flow.svg`), with
   `-tampered` added when traces were found.
4. **Charts the whole history** (`<name>-through.svg`): one panel per submission, labeled "#1", "#2", … above
   it (in red when traces were found). Only the changes new in each submission are counted, so changes left in a
   file handed back without accepting them are not counted twice (changes with the same type, author, date and text
   are treated as the same change).
5. **Writes an index** next to the documents: `index.html` (links to every figure) and `summary.csv` (for Excel).

A document is identified by its path relative to the folder you pass (without the extension), so passing the same
folder layout each time keeps each document in the same place. If file names or the folders below change between
submissions, use `--key-depth <n>` to identify documents by only their first n folders (for example
`--key-depth 1` for `<student>/<assignment>/<file>.docx`).

| Option | Description | Default |
|---|---|---|
| `--template <file.docx>` | The template handed out at the start; the first submission is checked against it | none |
| `--key-depth <n>` | Identify a document by only this many leading folders of its path | none (the whole path) |
| `-p, --gap-threshold <hours>` | Session gap for each submission's flow | `1` |
| `--bulk-chars <n>` / `--rules <file>` | Bulk-insertion detection (as in `docx-revision-chart`) | `150` / none |
| `--no-tamper-check` | Don't check for traces of tampering | off |
| `--lang <en\|ja>` | Display language | from the OS |

## Detecting traces of tampering

`docx-revision-chart` and `docx-revision-flow` check each file for traces of tampering with its tracked changes.
When they find any:

- The figure gets a red-bordered warning at the top listing the traces found.
- The output file name ends in `-tampered` (`report.svg` → `report-tampered.svg`, `report-flow.svg` →
  `report-flow-tampered.svg`), and so does the `--json` output. If you set the output path with `-o`, the name is
  left as is.
- The `--json` output lists the traces under `tamperEvidence`.

```bash
# Process a folder of submissions, checking each against the template you handed out
docx-revision-chart submissions/ --template template.docx
```

| Trace | What it means | `--template` |
|---|---|---|
| `trackingOff` | Track Changes was off when the file was saved | Not needed |
| `lockReleased` | The Track Changes lock was released; with a template, also when the template was locked but the file is not | Not needed |
| `authorDateRemoved` | Some changes have no author or date | Not needed |
| `notFromTemplate` | The file was not made from the template (its first editing session, rsidRoot, differs) | Needed |
| `lockChanged` | The lock's password differs from the template's | Needed |
| `untrackedText` | Text typed in editing sessions (rsid) that aren't in the template is outside the tracked changes: it was typed with tracking off, or its changes were accepted | Needed |

`untrackedText` is not checked without a template because it can't be told apart from a legitimate pattern:
creating a document, typing a little, and turning tracking on within the same editing session. In the settings
file, use `template: template.docx`; a relative path is resolved against the settings file's folder. To skip the
check, pass `--no-tamper-check`.

**A trace is not proof of tampering.** Saving with an app other than Word, an old Word version, or a converter can
produce the same marks. How files saved by Word for the web (often used when a Teams assignment is opened) come
out has not been fully checked yet. For example, `notFromTemplate` and `untrackedText` can't
be checked in a file saved by an app that doesn't write editing-session ids (rsid). Equally, finding no trace does
not prove there was no tampering. Check the file and ask the writer before drawing conclusions.

The desktop app's folder batch uses the same warning and file names for the checks that need no template, and
notes the traces in the summary (summary.csv, index.html). Library: `checkTamperEvidence(bytes, { template })` and
`addTamperWarning(svg, report)`.

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

## Language

Messages, help, and the figures are shown in English or Japanese. The language is chosen in this order:

1. `--lang en` / `--lang ja`
2. `lang` in the [settings file](#settings-file)
3. The `DOCX_REVISION_LANG` environment variable
4. The `LC_ALL` / `LC_MESSAGES` / `LANG` environment variables
5. The OS language setting (on macOS, System Settings > General > Language & Region)

Anything other than Japanese falls back to English. The macOS drag-and-drop apps follow the macOS language setting.

## Desktop app (preview)

[`desktop/`](../desktop/) contains a desktop app (Electron) for people who don't use the command line: open or drop a
.docx to see the chart, the flow and the highlighted insertions. Everything runs locally and the app blocks network
access. It can also open a .docx from a OneDrive / SharePoint link (after signing in to Microsoft); the document is
only downloaded for analysis. Installers aren't published yet.

**Folders (batch).** Drop a folder on the app (or its icon) to create a chart (`<name>.svg`) and a flow
(`<name>-flow.svg`) next to every `.docx` in it, including subfolders. Paste a SharePoint / OneDrive *folder* link in
"Open from URL…" — for example the folder where Microsoft Teams collects assignment submissions — and choose (or
create) an output folder: the source folder's hierarchy is recreated there with the SVGs inside (or, if you choose,
everything is put in one folder with the path in the file names). Both also write `summary.csv` (for Excel) and
`index.html` (links to every figure). The CLIs (and the macOS droplets) also accept folders: `docx-revision-flow
submissions/`.

**Lock template**: saves a copy of the opened document with Track Changes locked, through a save dialog (see
[Locking with these tools](operations-guide.md#2-lock-tracking-optional)).
