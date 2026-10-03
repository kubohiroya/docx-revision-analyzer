# 高度な使い方

[English](./advanced.md) | 日本語 · [README に戻る](../README.ja.md)

解析結果の JSON 出力、図への注釈、ハイライトの判定ルール、整合性についての情報です。

## 解析結果の JSON (`--json`)

`--json` を付けると、`docx-revision-chart` / `docx-revision-flow` は解析結果を JSON にも出力します。
中身は「挿入の窓」です。短い時間 (窓の最初の挿入から `seconds` 秒以内) に、文書の近い範囲
(`chars` 文字以内。`paras` を指定すれば段落数も) へ行われた挿入を1つの窓にまとめます。
距離は最終文書で測り、移動・並べ替えは除きます。窓の設定はルール (下記。既定は同じ作成者・同じ日時) から取り、
`--window-*` で置き換えられます。窓ごとの特徴量は次のとおりです。

| 特徴量 | 意味 |
|---|---|
| `insertedChars` | 挿入文字数の合計 |
| `durationSec` / `cps` | 時間幅 (秒) / 挿入速度 (文字/秒)。時間幅には推定した時刻の解像度 (`timeResolutionSec`。Word が分単位でしか記録していなければ 60、最低1秒) を足す (挿入はその幅のどこかで行われたとみなすため) |
| `spanChars` / `spanParas` | 最終文書での範囲 (文字数 / 段落数) |
| `maxSingleInsert` | 最大の単一の `w:ins` の文字数 |
| `paraCount` | 挿入が及ぶ段落の数 |
| `postEditRatio` | 窓の後にその範囲へ加えられた挿入・削除の文字数 ÷ `insertedChars` |
| `precededByDeletion` / `precedingDeletedChars` | 窓の直前または窓の間に、同じ範囲で削除があったか (置き換え) と、その文字数 |
| `insertCount` | 挿入の件数 |

特徴量そのものは判定をしません (しきい値は別に適用します)。

### 図の部分への注釈 (`--annotations`)

`--annotations <file>` (YAML / JSON) で、SVG の部分ごとに、マウスオーバーで出る説明 (`<title>`) やクリックで開くリンク
(`<a href>`。http / https のみ) を付けられます。部分のキーは `--json` の出力の `figureTargets` にあります: chart の棒
`chart:bar:<カテゴリ>:<バケットの開始時刻>`、flow の段落 `flow:para:<列>:<段落>`、帯 `flow:band:<区間>:<単位>`、移動の帯
`flow:move:…`、キャプション `flow:caption:<区間>`。

```yaml
annotations:
  - target: "flow:para:2:9"
    tooltip: { en: "Rewritten after feedback", ja: "指摘を受けて書き直した" }
    href: "https://example.com/notes#p9"
```

デスクトップアプリの拡張機能も同じ注釈を付けられます (`desktop/EXTENSIONS.md`)。注釈が無ければ SVG は変わりません。
ライブラリでは `resolveAnnotations`、`chartTargets` / `sessionedChartTargets` / `flowTargets` と、描画の `annotations` オプション。

### ハイライトの判定ルール (`--rules`)

ルールファイル (YAML / JSON) で、挿入の窓の特徴量に対する条件と段階 (レベル) を書きます。レベルは上から評価し、
最初に当てはまったものを採用します。レベルの付いた窓の挿入は一括挿入 (オレンジ) として描き、JSON には窓ごとの
`level` とルールの `ruleSet` を記録します。`--rules` を指定した場合は、図の右下にも `ruleSet` を表示します。
例: [examples/rules.example.yml](../examples/rules.example.yml)

```yaml
ruleSet: example-v1        # 出力に記録する識別子
window: { seconds: 60, chars: 2000 }   # 省略可。paras, byAuthor も指定できる
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

- 比較は `gte` / `gt` / `lte` / `lt` / `eq` (1つの中に複数書くとすべてを満たす必要がある)。`{ precededByDeletion: true }` は
  `eq` の省略形。`all: [...]` / `any: [...]` / `not: ...` で組み合わせる。レベルの数は任意。
- `--rules` が無いときは既定ルール (窓 = 同じ作成者・同じ日時、レベル = `insertedChars >= --bulk-chars` の1つ) を使います。
  結果は従来とまったく同じです。
- `--bulk-chars` は、`docx-revision-flow` で大きな削除を細かい編集に数えない判定にも引き続き使います。
- レベルはそれぞれの色で描き、凡例では「一括挿入」の代わりにレベルを載せます。chart では細かい編集と移動の間に
  積み上げます (下のレベルほど下)。`docx-revision-flow` では、段落に当てはまる最も上のレベルで段落を塗り、
  ほかのカテゴリは段落左側の細い帯で示します。
- 色だけで見分けずに済むよう、レベルには模様も付けます (上から順に `hatch` / `cross` / `dots`。レベルに `pattern:` を
  書けば指定でき、`pattern: none` で模様なし)。色の白地とのコントラストが 3:1 未満なら警告を表示します。

## 整合性についての情報

`--json` の出力には `integrity` も含めます。ファイルについてのいくつかの観察で、それぞれに `observed`
(Word が通常書き出す形と違えば `true`、同じなら `false`、必要な情報がファイルに無ければ `null`) と説明文があります。
**判定や警告には使わず**、ファイルが改変されたかどうかも示しません (Word のバージョンの違いや、ほかのアプリ・変換ツールでも
同じ違いは生じます)。

| 項目 | 観察する内容 |
|---|---|
| `settings` | 保存時に変更履歴の記録がオンだったか、保存時に作成者・日時を削除する設定か |
| `undated` | 日時・作成者の無い変更の数、作成者の数 |
| `duplicateIds` | 重複している変更の id (`w:id`)。Word は変更ごとに別の id を付ける |
| `idOrder` | 文書の順に見て変更の id が小さくなる箇所 (Word は保存時に文書の順に番号を振り直す) |
| `futureDates` | 未来の日時や、ファイルの最終更新日時 (`docProps/core.xml`) より後の日時の変更 |
| `beforeCreated` | ファイルの作成日時より1日以上前の日時の変更 (変更履歴の付いた別の文書から写した場合にも起こる) |
| `deletedBeforeInserted` | 削除の日時が挿入の日時より前の文章 |
| `rsids` | 本文で使われているのに `settings.xml` の一覧に無い編集セッションの識別子 (rsid) |
| `application` | 最後に保存したアプリと、記録された編集時間の合計 (`docProps/app.xml`) |

デスクトップアプリでは、ハイライトのタブに同じ一覧を表示します。ライブラリでは `checkIntegrity(bytes)`。

警告を出すのは、[改ざんの痕跡の検出](usage.ja.md#改ざんの痕跡の検出)で調べる項目だけです。
