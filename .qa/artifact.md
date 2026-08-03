# Template execution contract

- Reference: `/Users/cynthiaxii427/Desktop/广发/report_generator/template/总量通用.docx`
- SHA-256: `17747443d040f605759385d8b9b14ba1335caaae8bd6e3d6a720bb1cd4b34ce1`
- Package parts: 60
- Sections: 3; A4 portrait throughout.
- Render evidence: canonical renderer attempted at `/private/tmp/report-generator-qa/template-reference-render`; LibreOffice is unavailable, so no reference page PNGs were produced.
- Structural evidence: `/private/tmp/report-generator-qa/template-style-evidence.json` and the packaged section/content-control audits.

## Page system

- Section 1: 8.27 × 11.69 in; margins L/R/T/B = 0.32/0.32/0.34/0.27 in; new-page start; different first page; independent header/footer.
- Section 2: 8.27 × 11.69 in; margins = 0.39/0.39/0.34/0.27 in; continuous start; header linked, footer independent.
- Section 3: same size/margins as section 2; continuous start; header independent, footer linked.

## Stable slot map

- Cover: `word/document.xml`, first direct `w:body/w:tbl`.
- Cover title: marker `[Table_Title]`, first cell paragraph one direct sibling row later. Subtitle is two rows later and may be cleared.
- Cover summary: marker `[Table_Summary]`, first cell one direct sibling row later; existing paragraph properties are reusable and the cell may grow by cloning its last paragraph pattern.
- Industry/report line: marker `[Table_IndustryAndDate]`, one row later.
- Cover sidebar example areas: direct top-table rows 4–6, third cell; remove example chart/table/research content while preserving row and cell geometry.
- Index region: direct body children after the first paragraph-level section break and before the first paragraph styled `内页一级标题`; replace with a real TOC field and a page break.
- Main body example region: first paragraph styled `内页一级标题` through the direct body child immediately before the second paragraph-level section break. This range is replaceable.
- Preserve-only tail: the second paragraph-level section break onward, including research team, ratings, contact, legal, important notice, interest disclosure, copyright and final section properties.

## Reusable roles

- Heading 1/2/3: `内页一级标题`, `内页二级标题`, `内页三级标题`.
- Body: `内页正文`.
- Figure/table caption: `题注 + 思源黑体 CN Medium 6 磅`.
- Tables and figures from the input are copied as native OOXML blocks; their relationship-backed dependency graph is imported recursively.

## Fidelity gates

- Footer1–3 and base theme bytes must remain unchanged. Header geometry and
  artwork remain template-owned; marker text is filled from the reference
  report's `投资策略 | 专题报告` running-header convention.
- Section count, page size, margins, header/footer linkage and first-page behavior must match the reference.
- No `[Table_*]`, `正文文案`, `图表标题` or report-example placeholder survives in `word/document.xml`.
- Every clean `figure` and `table` must resolve to a manifest asset; no raster fallback is allowed.
- Every internal relationship in the final package must resolve to an existing part.
- `w:updateFields` must be true so the TOC refreshes in Word.

## Layout refinement reference (2026-08-03)

- Reference: `/Users/cynthiaxii427/Desktop/广发/【广发策略】从杠杆繁荣到筹码松动：韩国杠杆去化走到哪一步？(1).docx`
- SHA-256: `b1d02b01d24625d91855ce300767166728d1d829cfa08874da4d890a3b50d18f`
- Figure position token: borderless fixed-layout table, `tblW=8051 dxa`,
  `tblInd=2689 dxa`, one `8051 dxa` cell; caption and native drawing are
  non-splitting rows and the drawing paragraph is centered.
- TOC token: `TOC \\o "1-2" \\h \\z \\u`; Microsoft Word was used to update
  the complete field result after pagination.
- Final Word pagination: 22 pages. Visual checks covered the cover, refreshed
  contents page, first chart page, and retained page furniture.
- Final structure: 20 editable chart references, 27 source visuals inside 25
  Word-normalized layout tables, 7 TOC `PAGEREF` fields, and no template
  markers (`[Table_*]`, `XXXX`, `XX报告`).
- Final SHA-256: `db47bec30a92cef60443433cf2da699463b1ee449e17fb8603eba8b8615bd36d`.

## Final index, pairing, and caption pass (2026-08-03)

- Main contents field: `TOC \\o "1-2" \\h \\z \\u`; figure and table indexes:
  `TOC \\h \\z \\c "图"` and `TOC \\h \\z \\c "表"`. All three complete
  field results were refreshed and saved in Microsoft Word after pagination.
- Source automatic numbering is resolved during parsing, so hidden Word list
  numbering contributes the missing level-2 and level-3 headings to the TOC.
- Source editorial pairing directives create two fixed double-column visual
  tables (`10296 dxa`, two `5148 dxa` columns): one chart/chart pair and one
  editable image-table/chart pair.
- Every generated figure/table caption uses the reference paragraph separator:
  a black single bottom border (`w:sz=4`, `w:space=1`) with `40` twips after
  spacing. There are 24 caption paragraphs with this border (16 figures and 8
  tables), including independent left/right borders in paired layouts.
- Package and field audit: 20 editable chart references, 24 `SEQ` fields, 35
  `PAGEREF` fields, 3 sections, 22 pages, and no ZIP integrity errors.
- Visual Word checks covered the refreshed main TOC, figure/table indexes,
  single-chart caption separator, and paired-chart caption separators.
- Canonical PNG rendering remains unavailable in this runtime because its
  rendering dependency is absent; Microsoft Word was used for final visual QA.
- Final SHA-256: `a614b155ecd2864f1bcee37ca7be97a0d80d9fa7b81cd3858f13bd8ab0af4931`.

## Fixed closing pages and local corrections (2026-08-03)

- Fixed back-matter authority: `/Users/cynthiaxii427/Desktop/广发/report_generator/examples/output/【广发策略】从杠杆繁荣到筹码松动：韩国杠杆去化走到哪一步？(1).docx`, SHA-256 `b1d02b01d24625d91855ce300767166728d1d829cfa08874da4d890a3b50d18f`.
- The second section break through the document end is copied as a relationship-aware package graph, preserving the strategy research team, analyst/contact portraits and details, ratings, addresses, legal notices, disclosures, and copyright page.
- All 24 generated captions have neither `w:pStyle` nor `w:numPr`; the inherited square list marker is therefore removed while explicit caption typography and separator rules remain.
- The cover risk paragraph uses the retained `首页正文` style (`styleId=25`), with the label in `思源黑体 CN Medium` and remaining text in `思源黑体 CN Light`.
- All 20 editable chart parts have no-fill chart/plot outlines; all 7 imported picture frames have no-fill outlines. Pixel-baked borders are intentionally untouched.
- Package ZIP validation passed. Canonical PNG rendering was attempted but remains unavailable because `pdf2image` is missing.
- Final SHA-256: `fa3f53737b0fba4d94f652fdf255746a9db6310bf9cecf2fc3926502391d645a`.

## Caption recovery, paired alignment, and risk section (2026-08-03)

- The retained example contains 22 editable-chart captions. Raw editable charts are sequence-aligned to this list through exact neighboring caption anchors, allowing image-only charts in the example. Four missing raw chart captions are restored: chart assets 004, 005, 012, and 013.
- All 20 editable chart wrappers now contain a `SEQ 图` caption; total generated `SEQ` captions are 28 and no editable chart wrapper is untitled.
- Double-visual layout tables use `w:jc="right"`; the two equal-width caption rules and visual cells therefore share a common right boundary.
- Every cover-summary paragraph uses the reference `首页正文` style (`styleId=25`); ordinary runs remain Light and emphasized runs Medium.
- The report ending now includes the reference-formatted `五、风险提示` section with four source-derived risk paragraphs. Its last paragraph carries the continuous section break, followed by the fixed analyst/legal back matter.
- Structural checks: three sections, all four restored titles present, complete four-item risk text present both on the cover and in the ending section, and no ZIP errors.
- Canonical PNG rendering remains unavailable because `pdf2image` is absent; package-level and relationship validation were used.
- Final SHA-256: `5dd2c7578feae58a3d36b36ece228a67d16306bb45161991432ce751506b1896`.

## Word repair-warning correction (2026-08-03)

- Root cause: generator-created run, paragraph, table, cell, picture, and chart
  property children were semantically correct but some were emitted outside the
  strict OOXML schema sequence. Microsoft Word therefore offered to recover the
  package and reported a style repair.
- The generator now schema-orders all edited `w:rPr`, `w:pPr`, `w:tblPr`, and
  `w:tcPr` blocks across the Word package. DrawingML picture/chart outline and
  text-property children are also inserted in schema order.
- Style merging maps same-name/source styles to retained template IDs, permits
  only one default style per type, and rewrites imported header/footer style
  references. Full-package audit found no dangling style IDs.
- Validation: Python compilation passed; the full pipeline regenerated the
  report with 20 editable native charts, zero unmatched native objects, zero
  out-of-order audited Word property blocks, and no ZIP integrity errors.
- Final SHA-256: `a7fc45662c846cc05231368d06d6143a93aa152e8764fe66a155bc028c9f7907`.

## Body/data-source alignment correction (2026-08-03)

- Screenshot diagnosis: the body starts at `2688 dxa`, while the figure wrapper
  starts at `2689 dxa`; the visible drift came from Word's implicit table-cell
  left padding on top of that wrapper indent.
- Generated single- and double-figure cells now set explicit left/right margins
  to `0 dxa` (vertical margin remains `57 dxa`). This removes renderer-dependent
  default padding while preserving the specified 14.2 cm figure width and 4.74
  cm wrapper indent.
- Structural audit: all 29 generated data-source cells have `0 dxa` left
  margins. The screenshot-equivalent single-figure source starts at `2689 dxa`
  and its following body paragraph at `2688 dxa`, a non-renderable `1 dxa`
  rounding difference.
- Full pipeline passed with 20 native editable charts and zero unmatched native
  objects; ZIP integrity passed.
- Canonical PNG rendering was attempted but remains unavailable because the
  runtime lacks `pdf2image`.
- Final SHA-256: `cf19696fc1a719318d1488a07eae5cd1a1ede1057911446c7638823dfa908635`.

## Risk typography and paired-caption rule isolation (2026-08-03)

- All four ending risk-item paragraphs are explicitly formatted as
  `思源黑体 CN Medium`, `17` half-points (8.5 pt). The existing `五、风险提示`
  heading remains Medium 11 pt.
- In each of the two double-visual tables, only the inner edge of each caption
  paragraph is inset by `57 dxa` (left caption: right indent; right caption:
  left indent). This creates a `114 dxa` center gap between the two independent
  bottom rules without moving their outer edges, visuals, or data sources.
- The prior body/data-source alignment remains unchanged at `2689 dxa` for the
  source line versus `2688 dxa` for the adjacent body text.
- Pipeline, 20-chart native-object audit, and ZIP integrity passed. Canonical
  PNG rendering remains unavailable because `pdf2image` is absent.
- Final SHA-256: `f1618d87bdd1fb53918020757ce4e49b9d3a200c7eb4533dc82088936ad596e7`.

## Word unreadable-content repair (2026-08-03)

- Microsoft Word reproduced the blocking `发现无法读取的内容` warning on the
  generated output. A Word-recovered/save-as copy was compared with the exact
  package that triggered it.
- Root causes:
  1. replacing the template index range retained bookmark start `_Toc518314372`
     while deleting its bookmark end;
  2. the final risk paragraph was deep-copied with section footer reference
     `rId56`, but its source document relationship was not imported.
- The generator now removes bookmark endpoints orphaned by replaced template
  ranges and clones the entire risk section through `PackagePartImporter`, so
  its section-level header/footer relationships and dependent parts are copied
  and remapped.
- `validate_package` now fails generation when any `r:id`, `r:embed`, or
  `r:link` points to a relationship absent from that part's `.rels`, or when
  bookmark start/end IDs are unbalanced.
- Acceptance proof: a byte-identical `/private/tmp/rg_relationship_fixed.docx`
  was opened in Microsoft Word from a fresh path. Word showed no recovery or
  unreadable-content prompt; only the expected editable-chart external-field
  update prompt appeared. Selecting `No` opened the 27-page report normally.
- Regression audit: 20 editable charts, 3 sections, 37 resolved relationship
  references, balanced bookmarks, Medium 8.5 pt risk items, two separated
  paired-caption rule groups, retained `2689 dxa` source alignment, and no ZIP
  errors.
- Canonical PNG rendering remains unavailable because `pdf2image` is absent;
  the opened cover was visually inspected directly in Microsoft Word.
- Final SHA-256: `c15813012c2187a6366fc67e8358e4e1130f3e2561bbea8b8adf5e86cb5af5ec`.
