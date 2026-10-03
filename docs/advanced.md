# Advanced usage

English | [日本語](./advanced.ja.md) · [Back to the README](../README.md)

The analysis JSON output, annotations on figures, highlight rules, and integrity notes.

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
figures also show the `ruleSet` at the bottom right. See [examples/rules.example.yml](../examples/rules.example.yml).

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

Only the items in [Detecting traces of tampering](usage.md#detecting-traces-of-tampering) lead to a warning.
