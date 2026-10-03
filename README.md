# docx-revision-analyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
English | [日本語](./README.ja.md)

Four CLI tools that analyze Word (`.docx`) files edited with Track Changes
enabled:

1. **`docx-revision-chart`** — renders an SVG chart of edit activity over time
2. **`docx-revision-flow`** — for each session of continuous editing, draws
   the document's pages at the start and end of the session as a schematic,
   colors fine-grained edits (green), bulk insertions/replacements (orange) and
   deletions (red), and connects each paragraph/figure/table across the two
3. **`docx-ai-suspicion-score`** — scores how likely it is that a chunk of
   text was pasted in from an external app (e.g. an AI writing tool) rather
   than typed and reviewed inside Word, on a 0–100 scale
4. **`docx-revision-snapshot`** — archives each submission of a long-running
   document (such as a thesis handed in at the end of each sprint) under a
   serial number, and charts the revision history across all of them

They're published as an npm package and also distributed as single, dependency-free executables built with
[Bun](https://bun.sh) (no Node.js required). Licensed under MIT.

> **Prerequisite**: all tools require a `.docx` file that was authored/edited
> with Word's "Track Changes" turned on (Review tab → Track Changes). Files
> edited with Track Changes off don't retain insertion (`w:ins`) / deletion
> (`w:del`) markup, so there's nothing to analyze (the tools exit with an
> error explaining why). See "[The `.docx` files these tools expect](#the-docx-files-these-tools-expect)".

> **What this is for**: a tool for writers to look back on how their own document was written — not a way to police
> others. Tracked changes in a `.docx` saved on a computer are easy to leave out, remove or rewrite: edits made with
> Track Changes off are never recorded, accepting or rejecting all changes erases the ones already recorded, and
> editing the XML can change `w:date` / `w:author`. Other apps and converters also write
> files differently from Word. So the history these tools show can be incomplete or altered, and nothing they output
> is proof of how a document was written. See "[Integrity notes](#integrity-notes)".

---

## The `.docx` files these tools expect

**These tools are meant for `.docx` files in state D below ("Sufficient"): files that hold enough tracked
changes, recorded over a long enough period, for their content to be analyzed.** The tools read the tracked
changes (`w:ins` / `w:del`, with their `w:date` and `w:author`) stored inside a `.docx`, and how much a file can
tell you depends on how much of that history it still holds. A document starts in state A and is always in one of
four states:

![Revision history in four states. From the start, a document is in A (no tracking), then moves to B (tracking on) → C (too sparse) → D (sufficient), the state these tools analyze. Accepting or rejecting some changes moves D back to C; "Accept All" or "Reject All" moves C or D back to B; "Accept All Changes and Stop Tracking" moves them back to A, and none of this can be undone](./fixtures/revision-states.png)

| State | The document | What the tools do |
|---|---|---|
| **A: No tracking** | Track Changes is off, so edits leave no record | Exit with an error saying the file was saved with Track Changes off |
| **B: Tracking on** | Track Changes is on, but nothing has been edited since | Exit with an error saying there are no tracked changes yet |
| **C: Too sparse** | Some edits are recorded, but too few or over too short a time | Run, but the chart and flow show little, and `docx-ai-suspicion-score` warns that fewer than 5 insertion events are statistically weak |
| **D: Sufficient** | Enough edits recorded over a long enough period | **What these tools are for:** produce meaningful results |

A document moves forward, A → B → C → D, as you keep writing with Track Changes on. Turn it on (Review tab →
Track Changes) **before** you start writing; what was typed while it was off is never recorded.

It can also move backward, and every red path in the figure loses tracked changes:

- **D → C:** accepting or rejecting individual changes removes them from the history, which can leave too little
  to analyze.
- **C / D → B:** "Accept All" or "Reject All" removes every tracked change. Track Changes stays on, so the
  document is back to "tracking on, no changes yet".
- **C / D → A:** "Accept All Changes and Stop Tracking" removes every tracked change and turns tracking off.

Turning Track Changes off is not one of these: in state C or D the changes already recorded stay in the file, and
only later edits go unrecorded. In state B it takes the document back to A, but there is nothing to lose there.

**Lost tracked changes can't be brought back** (for files saved on OneDrive, see [below](#files-saved-on-onedrive)),
so to keep the history, analyze — or keep a copy of — the file before accepting or rejecting changes.

Also check that each change keeps its timestamp. If the document is set to remove personal information on save,
Word strips the author and date of every tracked change, so the timeline is lost even in state D; see
"[When timestamps are missing](#when-timestamps-are-missing---preserve-history)".

### Files saved on OneDrive

OneDrive keeps earlier versions of a file in Version History, which adds a way back that a file on your computer
doesn't have:

![The same four states for a file saved on OneDrive. In addition to the paths above, a document in A, B or C can return to D by restoring, from Version History, a version saved while it was in D](./fixtures/revision-states-onedrive.png)

If a version saved while the document was in state D is still in Version History, restoring it brings the document
back to D, whichever state it is in now. This restores an earlier copy of the whole file, so later edits are not in
it, and it isn't guaranteed: OneDrive keeps only a limited number of versions, and a file that was never in state D
has no such version to restore.

### In a class: hand out a template with locked tracking through a Teams assignment

To use these tools on student reports, the reliable way is for the teacher to prepare a template with Track
Changes on and hand each student a copy of it through a Microsoft Teams assignment. The template can also be
protected with Word's built-in Lock Tracking.

#### 1. Make the template

Leave the title, student ID and name blank for students to fill in. If useful, add the report's section headings
or dummy body text for students to write over. Finally, turn Track Changes on (Review tab → Track Changes) and
save. The template is now in state B.

#### 2. Lock tracking (optional)

**How to set it (teacher)**

- **Word for Windows:** Review tab → the Track Changes ▼ → Lock Tracking, enter a password, then OK
- **Word for Windows (another way):** Review tab → Restrict Editing → under "2. Editing restrictions", check "Allow
  only this type of editing in the document" and choose "Tracked changes" → under "3. Start enforcement", click
  "Yes, Start Enforcing Protection" → enter a password
- **Word for Mac:** Review tab → Protect Document (in some versions, Protect → Protect Document), choose "Tracked
  changes" and enter a password

The password is optional, but without one anyone can unlock tracking from the same menu, so set one. To edit the
template later, unlock it from the same menu with the password, make your changes, and lock it again.

**Locking with these tools**

Instead of using Word, you can lock the template with these tools. Both also turn Track Changes on, and turn off
the setting that removes authors and dates on save if it is set.

- **CLI:** pass the template with `--lock`. It asks for the password twice and locks the file in place (the
  original is kept as `<name>.backup-<date>.docx`). No figure is drawn.

  ```bash
  docx-revision-chart --lock template.docx
  ```

  Where there is no terminal (such as when launched by drag and drop), pass the password in the
  `DOCX_LOCK_PASSWORD` environment variable. Without it the file is locked with no password, except that a file
  already locked with a password is left unchanged with an error rather than re-locked without one.
- **Desktop app:** open (drop) the template and click "Lock template". After you enter the password, a dialog asks
  where to save the locked copy (named `<name>-locked.docx` by default). The opened file itself is not changed.

The password hash follows the specification (ECMA-376) and Apache POI's implementation, in the format Word 2013
and later use (SHA-512, 100,000 rounds), so Word's Lock Tracking menu should accept the same password. Before
handing the template out, open it in Word once and check that the password unlocks it.

**What this does**

- Track Changes stays on, and students who don't know the password can't turn it off.
- Accept and Reject are grayed out and can't be used. This closes the ways a student could lose tracked changes
  (state D → C, C / D → B, C / D → A).
- Typing and deleting work as usual, and every edit is recorded as a tracked change.
- Inside the file, `word/settings.xml` gets `<w:documentProtection w:edit="trackedChanges" w:enforcement="1" .../>`
  and `<w:trackRevisions/>`. The lock is a setting of the file, so copies of the template keep it.

**Limits**

The lock is not strong protection. The password is stored only as a hash in `settings.xml`, and unzipping the
`.docx` and deleting that element removes the lock. Selecting all the text and pasting it into a new document
gives a document with neither the lock nor any tracked changes, and apps other than Word may not honor the lock.
Also, deleting text you inserted while tracking removes it outright, leaving no record of the deletion, even with
the lock on (this is how Word works). Treat the lock as a guard against mistakes, and pair it with a way to check
that a submitted file was made from the template. These tools show a warning in the figure when a submitted file
has traces of the lock being removed or of text typed with tracking off (see
[Detecting traces of tampering](#detecting-traces-of-tampering)); pass the template you handed out with
`--template` to also check each file against it.

Before handing it out, check with a test account that the locked file can be edited in the app students will
actually use (Word desktop, Word for the web and so on). Some apps may not let you edit a protected document.

#### 3. Hand it out with a Teams assignment

1. In Teams, open Assignments and create a new assignment
2. Enter the assignment's title and instructions
3. Under Attach, choose the template file from step 1 (and 2)
4. Click "Students can't edit" below the attached file and change it to **"Students edit their own copy"**

When a student opens the assignment, a copy of the template just for them is created automatically. They edit
that copy on OneDrive with AutoSave on, and turn it in as is.

#### 4. Tell students what to do

- Don't use Accept or Reject while writing (with the lock on, they can't).
- If the markup gets in the way, set the display to "No Markup" to write on a clean view; changing the display
  doesn't stop the recording.
- Keep AutoSave (top left of the window) on. Turning it off means fewer saves, so fewer versions are kept in
  Version History.

#### Why this is reliable

- Students don't have to turn Track Changes on themselves. Each copy starts in state B, so nobody ends up writing
  in state A because they forgot. With the lock on, nobody can turn it off midway either.
- Each copy lives on OneDrive and is saved automatically, so the Version History described
  [above](#files-saved-on-onedrive) keeps a fine-grained record.
- Students open their copy, write in it and turn it in as is, which leaves less room to swap in a file rebuilt on
  their own computer.

#### Long-running writing (such as a thesis): submissions at each sprint

In a document written over months, tracked changes get lost in two ways:

- OneDrive's Version History drops old versions, because of limits on their number and age and the
  organization's retention policies.
- Text you typed while tracking is your own pending insertion until it is accepted. Deleting it leaves no record
  of the deletion; it simply disappears. Rewrite the same passage several times and the drafts in between are gone.

So have the writer hand the document in at the end of each sprint (one or two weeks, or each chapter), archive
it, then accept all the changes and hand it back. Rewriting accepted text is recorded as a deletion, so every
revision across sprints is kept; only rewrites within a single sprint are lost.

1. The student hands the document in at the end of the sprint.
2. The teacher archives it with [`docx-revision-snapshot`](#4-docx-revision-snapshot--archive-and-chart-submissions-over-time),
   which numbers it, checks it against the previous submission, and draws each submission's figures and a chart
   across all of them.
3. The teacher opens the same file in Word, unlocks it, clicks "Accept All Changes", locks it again and saves.
4. The teacher hands it back, and the student continues in the same file.

The submitted file sits in the student's OneDrive, where the student can delete it, so keep the archive on the
teacher's side.


If you make these requirements of the assignment, state them in advance, for example in the syllabus's notes for
students. A sample:

> To ensure the integrity of the learning process, assignments in this course that require submitting a file based
> on a template must be worked on in OneDrive, with Track Changes kept on at all times and the Version History
> preserved. Students with a legitimate reason they cannot follow this must tell the instructor in advance.
> Submissions that do not meet these requirements will not be graded.

---

## Installation

### From npm (Node.js 18+)

```bash
npm install -g docx-revision-analyzer
```

This installs four commands: `docx-revision-chart`, `docx-revision-flow`, `docx-ai-suspicion-score`, and `docx-revision-snapshot`.

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
| `--template <file.docx>` | The template you handed out; each file is checked against it for traces of tampering ([Detecting traces of tampering](#detecting-traces-of-tampering)) | none |
| `--no-tamper-check` | Don't check for traces of tampering (no warning, no `-tampered`) | off |
| `--preserve-history` | If the document removes personal information (tracked-change authors and dates) on save, remove that setting, turn Track Changes on, and save it in place (the original is kept as a backup; fails if the document is open). `--preserveHistory` also works. See "When timestamps are missing" below | off |
| `--check-history-settings` | Don't draw a chart; only check the settings and print `ok` or `needs-fix` on the first line, followed by the confirmation text when `needs-fix` | off |
| `--lock` | Don't draw a chart; lock Track Changes in the given template and overwrite it (the original is kept as a backup). See [Locking with these tools](#2-lock-tracking-optional) | off |
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

## 4. `docx-revision-snapshot` — archive and chart submissions over time

Archives the submissions (snapshots) of a long-running document, such as a thesis handed in at the end of each
sprint, and analyzes them as one history. For the workflow, see
[Long-running writing](#long-running-writing-such-as-a-thesis-submissions-at-each-sprint).

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

Only the items in "Detecting traces of tampering" below lead to a warning.

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

## Desktop app (preview)

[`desktop/`](desktop/) contains a desktop app (Electron) for people who don't use the command line: open or drop a
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
[Locking with these tools](#2-lock-tracking-optional)).

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
