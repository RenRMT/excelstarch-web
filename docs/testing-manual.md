# Manual sideload testing — the interop layers Jest cannot cover

The pure `config/` and `logic/` layers are unit-tested with Jest (`npm test`). Everything that
touches the live Office host is **not** unit-tested and must be verified by sideloading the add-in into a real
Excel. This document is the canonical checklist for that.

> **Why manual?** `excel/` and `chrome/` call `Excel.run` and manipulate the workbook
> object model, charts, and worksheet shapes. Mocking the entire Office object model is brittle and
> proves nothing about real host behaviour — so we sideload instead. The decision boundary: if a
> function imports `Excel`/`Office`, it is sideload-tested, not Jested.

## Scope by phase

Sections 3–5 map to the delivery phases of the original port. Section 6 covers the follow-up work
from [implementation-plan.md](implementation-plan.md), one subsection per phase of that plan.

| Phase | What to sideload-test | Section |
|-------|-----------------------|---------|
| 0 | Nothing — no interop yet. `npm test` green is the whole gate. | — |
| 1 | Bar + column creation, branded styling, the worksheet-chrome group, chart-text panel. | [§3](#3-phase-1--chart-creation-styling-and-chrome), [§4](#4-phase-1--chart-text-chrome-panel) |
| 2 | Remaining chart types, full colour tooling, per-element fill, last-used persistence. | [§5](#5-phase-2--colour-tooling-and-breadth) |
| follow-up | Fixes and features from the implementation plan. | [§6](#6-follow-up-port) |
| every | Requirement-set gating, cross-platform smoke. | [§7](#7-requirement-set-gating), [§8](#8-cross-platform-smoke) |

---

## 1. Prerequisites

- A clean checkout with dependencies installed (`npm install`).
- A trusted dev certificate. The first `npm start` provisions one via `office-addin-dev-certs`; on
  Windows accept the prompt to install the CA. If the task pane shows a blank/red page, the cert is
  the usual cause — re-run `npx office-addin-dev-certs install`.
- Excel: at least one of **Excel on the web** (Microsoft 365), **Excel for Windows**, or
  **Excel for Mac**. Cross-platform parity is a project requirement, so test more than one when you can.

## 2. Sideloading

```sh
npm start          # builds dev bundle, provisions cert, sideloads into the host in package.json "config"
```

- `package.json` → `config.app_to_debug` selects the host (`excel`) and `app_type_to_debug`
  (`desktop` by default; set to `web` for Excel on the web).
- To target the web explicitly: `npm start -- --app excel` after setting the debug type, or sideload
  `manifest.xml` manually via **Insert → Add-ins → Upload My Add-in** in Excel on the web.
- Stop and clean up the sideload registration with `npm stop`.
- After editing interop code, `npm run watch` rebuilds on save; reload the task pane (right-click →
  Reload, or close/reopen the pane) to pick up changes.

**Sanity before any feature test:** the add-in loads, the ribbon button appears, and clicking it
opens the task pane with no console errors (F12 / Edge DevTools on the web; the in-pane runtime
console on desktop).

---

## 3. Phase 1 — chart creation, styling, and chrome

**Fixture.** A worksheet with a small labelled data range — one category column and 1–3 numeric
series columns (e.g. 5 rows × 3 series). Keep a copy with **8 series** and one with **a single
series** for the colouring edge cases later.

### 3.1 Create a branded chart

1. Select the data range (including header row/column).
2. In the task pane, choose **Bar** (then repeat for **Column**).
3. Verify, against the VBA output side by side:
   - [ ] A chart is created on the worksheet (embedded, not a chart sheet).
   - [ ] Canvas is **600 × 600 points** (right-click chart → Format → Size; or trust the layout —
         it should be a visibly square canvas).
   - [ ] Series use the **brand data palette** in order (ROOS Hemelblauw, Oranje-750, Violet-600, … = `colorData1..8`).
   - [ ] Gridlines match the type: **column → Y-gridlines only**, **bar → X-gridlines only**.
   - [ ] Axes/legend match the chart-type defaults (bar/column: both axes shown, legend off by default).

### 3.2 The worksheet-chrome group

Office.js charts cannot host in-chart shapes, so chrome is built as **worksheet shapes grouped with
each other** (generalizing `modEngineExChrome.bas`). The chart itself cannot join the group — charts
and shapes are separate object models in Office.js — so it is sized into the canvas band. Verify:

- [ ] A white **canvas rectangle** sits behind the chart (sent to back).
- [ ] **Title, subtitle, source/notes** text boxes appear at the branded positions with the correct
      fonts/sizes/colours (all Verdana; title 28pt Lintblauw `colorBrand1`, subtitle 22pt `colorBrand2`, source
      14pt `colorBrand3`). The title sits at the very top of the canvas; there is **no figure-number
      box and no logo**.
- [ ] The bottom of the chart lines up with the top of the source/notes (footer) box.
- [ ] Selecting a chrome element and dragging moves the **chrome group together** (canvas + text),
      named `ESChromeGroup_<chartName>`. The **chart does not move with it** (expected); re-running
      the builder on the chart snaps it back onto the canvas.
- [ ] Placeholder text matches the spec (`"Title in 28pt sentence case"`,
      `"Source: Source text goes here."`, etc.).

### 3.3 Re-run / reuse-canvas behaviour

1. With a branded chart present, select its range and create the chart again (same type).
2. Verify:
   - [ ] The chrome **snaps back to the same canvas position** (the reuse-canvas-origin trick), rather
         than stacking a second offset group.
   - [ ] No orphaned shapes are left behind.

### 3.4 Partial-failure safety

- [ ] If a shape fails to create (rare; force by, e.g., a protected sheet), the add-in **groups only
      the successfully-created shapes and shows a non-blocking warning** (Fluent `MessageBar`) — it
      does not crash or leave the workbook half-built.

---

## 4. Phase 1 — chart-text chrome panel

The VBA `Selection`-driven click-to-edit model is gone; chrome text is edited from the task pane.

1. With a branded chart present (and selected, if the resolver requires it), open the chart-text panel.
2. Type into the **Title / Subtitle / Y-axis / Source / Notes** fields.
3. Verify:
   - [ ] Each field writes to the matching shape's `textFrame.textRange.text` — the chart chrome
         updates live (or on apply, per the panel's design).
   - [ ] Empty field restores/keeps the placeholder, it does not crash.
   - [ ] Editing one field does not disturb the others or ungroup the chrome.
   - [ ] Italic uses `font.italic = true`, **not** a named italic font family (cross-platform).

---

## 5. Phase 2 — colour tooling and breadth

### 5.1 Remaining chart types

**All chart types shipped:** bar, column, line, area, scatter, pie (classic) plus treemap and box &
whisker (chartex). Lollipop and the other chartex types (sunburst, histogram, waterfall, funnel,
pareto) remain a follow-up.

For each classic type (line, area, scatter, pie):

- [ ] Creates a 600×600 branded chart with the type-appropriate defaults (see `chartDefaults`).
- [ ] Chrome group builds correctly; the chart insets into the canvas band (re-run snaps it back).
- [ ] Gridlines/axes match the defaults: **line/area → Y-gridlines, both axes**; **scatter → both
      gridlines, both axes**; **pie → no axes/gridlines, legend ON**.
- [ ] **Pie point-colouring (HIGH-risk path):** a pie's slices are *points* of one series, so the
      colourer fills **points** in brand-palette order (`colorByPalette` with a `point` target), not the single series.
      Confirm slices are distinctly coloured (not one flat colour). Try a fixture with 3 and with 8+
      categories (>8 falls back to neutral per `getPaletteColor`).
- [ ] **Pie band/legend layout:** legend sits below the subtitle without overlap; the no-axes band
      (`chartBandPosition(false,false,true)`) leaves the title block above and the source box below.
- [ ] **Scatter axis/marker (HIGH-risk path):** scatter has two value axes (no category axis).
      Confirm `applyChartStyle` does not throw and the chart shows the intended X/Y gridlines + axes,
      and that markers pick up the brand colour. If markers don't colour, note the `IsLineTarget`
      seam in `seriesRecolorer`/`seriesColorer` (see §5.3) as the follow-up.
- [ ] **Retype between kinds:** with a chart selected, create a different kind (e.g. bar → pie → line)
      and confirm it retypes in place and the chrome snaps back to its canvas origin.

For the chartex types (treemap, box & whisker — both ExcelApi 1.9, below the add-in's minimum):

- [ ] **Treemap (HIGH-risk path):** select a categorical range → Treemap → branded 600×600 chrome,
      **no axis errors** (`applyChartStyle` now gates axis styling on `axisDisplay`, so it never
      touches the treemap's non-existent value/category axes), **no Y-axis title box**. Tiles are
      coloured **per point** in brand-palette order (the `point` colour target, pie's path) — confirm
      distinct tile colours, not one flat fill. Try 3 and 8+ categories (>8 → neutral).
- [ ] **Box & Whisker:** select multi-series numeric data → Box & Whisker → branded chart with
      **Y-gridlines + both axes** and the **worksheet Y-axis title box present** (`showYAxisTitle:
      true` → `chromeBuilder.addYAxisTitle`). Series coloured **per series** in palette order (not
      per point). Chrome group builds; re-run snaps it back. Confirm the worksheet Y-axis title box
      does **not** visually collide with the chart's own rendered value-axis tick labels (box &
      whisker keeps `axisDisplay: Both`, so the chart still draws its value axis).
- [ ] **Axis-guard regression:** after the `applyChartStyle` change, re-confirm **pie** (distinct
      slice colours, legend, no axis error) and **scatter** (gridlines/axes/markers) still create
      correctly. Retype an existing chart across the classic↔chartex boundary (e.g. column → treemap
      → box & whisker) and confirm no axis error on the no-axis (treemap) leg and chrome snaps back.

### 5.2 Palette / ramps / diverging

Use the multi-series fixtures. The **pure** ordering math is already Jest-proven; here you confirm it
is wired to the host correctly.

- [ ] **Palette** recolours series in brand order; the **Rainbow/Contrasting toggle** swaps slots 2
      and 6 (matches `getPaletteColor` with/without alt order).
- [ ] **Single-hue ramp** (e.g. Hemelblauw): series go **darkest-first** (series 1 darkest), n=1..6.
      Spot-check n=3 → steps `[6,4,2]` of the ramp; n=6 → all six ROOS tints, base→150.
- [ ] **Diverging ramp** (e.g. `Rood|Groen`): dark→light on the left, light→dark on the right; **odd series
      count gets a grey centre** (`colorBrand4`), even count does not. Try n=5 (grey middle) and n=8.
- [ ] **Invert** reverses the current fill assignment across all series.
- [ ] Over-limit guard: >6 series for a single ramp, or >13 for diverging, shows the non-blocking
      "too many series" message instead of mis-colouring.

### 5.3 Per-element fill (element selector)

- [ ] The element selector lists the chart's series; choosing one and applying a fill colours **only
      that series** (replaces the lost VBA Selection model). "All series" applies to every series.
- [ ] A colour applies (`DATA1..8`, `NEUTRAL2/4`); `NONE`/`NOFILL`/`OFF` removes the fill.
- [ ] **Transparency is not offered** for series fills — Office.js `ChartFill` has no transparency
      API (the VBA `DATA1|0.5` case doesn't apply to chart series). Confirm there is no transparency
      control and a solid colour is applied.
- [ ] Line/scatter series colour the **line**, not a fill (the `IsLineTarget` distinction) — **N/A
      until line/scatter ship** (chart-types breadth PR); the apply seam is in `seriesRecolorer`.

### 5.4 Last-used persistence

The Color Picker remembers the last-used **single-hue ramp, diverging tag, fill colour, and
palette-order toggle** per workbook (`persist/settings.ts` over `Office.context.document.settings`,
the web replacement for VBA `CustomDocumentProperties`). Values are saved **after a successful
apply** and restored as the pane's **pre-selected dropdown/switch defaults** when it next mounts
(the user still clicks Apply — there is no auto-re-apply). Document-global scope (one value per
setting, like the VBA).

- [ ] Select a chart; in the Color Picker pick a non-default ramp (e.g. Coral) and **Apply ramp** →
      success. Switch to **Rainbow order** and **Apply palette**. Pick a diverging tag (Apply) and a
      fill colour (Apply fill).
- [ ] **Reopen the pane** (close the Color Picker, open it again): the ramp, diverging, and fill
      dropdowns and the order switch show the **last-used** choices, not the hard defaults.
- [ ] **Close and reopen the workbook** (save first): the same choices are restored — proving the
      values persisted to `document.settings`, not just React state.
- [ ] **New/clean workbook:** the pane opens with the **hard defaults** (first ramp, first diverging
      tag, DATA1 fill, Contrasting order) and does not error when no settings are stored.
- [ ] **Failed/cancelled apply does not overwrite** last-used: trigger an over-limit error (e.g. a
      single ramp on >6 series), then reopen the pane — the previously-saved ramp is still shown.
- [ ] Persistence is **best-effort/silent**: a save failure never shows an error or blocks the apply.

---

## 6. Follow-up port

Checks for the work in [implementation-plan.md](implementation-plan.md). None of this could be
sideloaded when it was written, so every item here is a first verification, not a regression check.

### 6.0 Plan phase 0 — foundations

**Colour lives in the right place**

- [ ] **Line** chart: each series line is a brand colour, 2 pt. No stray fill or marker changes.
- [ ] **Scatter** chart: each series' markers are a brand colour.
- [ ] Color Picker on a line chart: palette, single ramp, diverging ramp and Invert all change the
      **line** colour. "No fill (remove)" hides the line.
- [ ] Color Picker on a scatter chart: the same tools change the **marker** colour.
- [ ] Color Picker on a **pie** and a **treemap**: the tools recolour the **slices / tiles**, the
      Element list shows the points (category names, or "Point n"), and filling one element
      colours one slice.
- [ ] Color Picker on a chart the add-in did not create (insert a native line-with-markers chart):
      line and markers both take the colour.

**Text survives a rebuild**

- [ ] Type a title, subtitle, source and notes. Select the chart and choose a different chart type.
      All four texts are kept; the chrome is still grouped and in the same place.
- [ ] Clear a field with `-`, then restyle: the box stays empty (it does not come back as the
      placeholder).

**Chart-text panel**

- [ ] Close and reopen the Chart Builder pane, then select a styled chart: the fields fill with the
      chart's current text and are editable without rebuilding the chart.
- [ ] Select a different styled chart: the fields switch to that chart within a moment, with no
      Refresh. Click a cell: the panel keeps editing the last chart.
- [ ] Select a native, unstyled chart: the panel disables itself.
- [ ] With two worksheets each holding a "Chart 1", editing text changes the chart on the sheet
      that chart lives on, not the active sheet's namesake.
- [ ] The Y-axis title field is enabled for box & whisker and disabled for pie.
- [ ] Typing into a field and blurring writes once; tabbing through untouched fields writes nothing
      (in particular, it does not restore placeholders into cleared boxes).

**Selection tracking and legend**

- [ ] Color Picker: selecting and deselecting a chart enables and disables the controls without
      pressing Refresh. Adding a new worksheet and selecting a chart there also works.
- [ ] Pie: the legend is a row at the top of the chart object, in the axis font size, and does not
      overlap the subtitle. There is no empty strip between the subtitle and the chart.

### 6.1 Plan phase 1 — export

The Export section of the Chart Builder pane composes the chart and its chrome into one image in
the task pane (`excel/exportFlow` → `taskpane/export/compose`). Three things in this path could not
be confirmed without a live host; check them first and note the results in the PR.

**Spikes**

- [ ] **Group render.** Export a styled chart. If the title text looks exactly like Excel's (same
      font rendering, same wrapping), the chrome came from `shape.getAsImage` on the group. If it is
      slightly different, the group render failed and the pane drew the text itself — note which
      host did which.
- [ ] **Chart transparency.** The exported chart area shows the white canvas, with no black or
      grey box behind the plot.
- [ ] **Delivery.** For each platform, record which of **Download**, **Copy** and **Insert on
      sheet** work. (Copy is hidden when the browser has no image clipboard.)

**Behaviour**

- [ ] PNG at 1×, 2× and 3× gives 800, 1600 and 2400 px square images for a styled chart; the
      success message reports the size.
- [ ] The image matches the sheet: title, subtitle, source/notes, y-axis title (box & whisker), and
      the chart sitting in the same band. Try bar, line, pie and treemap.
- [ ] JPEG has a white background (no black areas where the PNG is transparent).
- [ ] The download is named after the chart title; with the placeholder title it is named after
      the chart (e.g. `Chart 1.png`). Titles with `/`, `:` or `?` give a valid file name.
- [ ] **Insert on sheet** places a picture of the same size to the right of the canvas, and it
      can be copied into PowerPoint as one object.
- [ ] A native chart with no ExcelStarch chrome exports as the chart alone.
- [ ] No chart selected: a clear "No chart selected" message, nothing exported.
- [ ] Format and resolution are remembered after closing and reopening the pane.

### 6.2 Plan phase 2 — chart variants, finishing, Apply Chart Style

**Fixtures:** the 3-series fixture; an X/Y/size range for bubble; a single-column category range
for pie/donut.

**New variants** — each creates a branded chart of the right type:

- [ ] Stacked column, 100% column, stacked bar, 100% bar: segments are flush (no gap between
      stacked segments), bars are wide (gap width 33%).
- [ ] Area is now **stacked** (no series hidden behind another); 100% area fills to the top.
- [ ] Donut: per-slice colours, hole about half the diameter, legend at the top.
- [ ] Bubble: per-series fills (opaque — transparency is not available in Office JS).

**Per-type finishing** (compare with the VBA output side by side):

- [ ] Clustered column/bar: a small gap between series in a group (overlap −5), no shadows.
- [ ] Line: tick marks outside, the line starts on the y-axis (no half-category gap), 2 pt lines.
- [ ] Area: no tick marks, the area starts on the y-axis.
- [ ] Scatter: markers cycle circle → square → triangle → X per series, size 9, tick marks outside.
- [ ] Pie and donut: the plot is a centred circle that does not touch the legend or the source box.
- [ ] Treemap and box & whisker: data labels (if shown) use the chart font. No warning is shown if
      the host refuses this.
- [ ] If any of the above is refused by a host, the chart is still created and a warning names
      what was skipped. Record any such warning per host.

**Y-axis title on classic charts**

- [ ] Column, bar, line, area, scatter and bubble now get the italic "Y axis title (unit)" box
      in the strip under the subtitle; pie, donut and treemap do not.
- [ ] Restyling an older chart (made before this change) adds the box with the placeholder and
      keeps the existing title text.

**Apply style to selected chart**

- [ ] On a branded column chart: nothing about the type changes; chrome is rebuilt in place and the
      text is kept.
- [ ] On a native Excel chart of each builder type (e.g. insert a stacked bar from Excel's own
      ribbon): it ends up identical to the builder's version of that type.
- [ ] On a native line-with-markers chart and an XY scatter-with-lines chart: brand colours on lines
      **and** markers, Y gridlines (line) or both gridlines (scatter).
- [ ] On a 3-D column, a waterfall and a sunburst: canvas and chrome are added, colours are applied
      if the type allows, and **the axes and legend are left as they were**. No error.
- [ ] No chart selected: "No chart selected" message.

### 6.3 Plan phase 3 — chart tools (toggles)

The **Chart tools** section of the Chart Builder pane. Each button acts on the selected chart and
the status line names the new state. Use a branded 3-series column chart unless stated.

- [ ] **Gridlines** cycles none → value axis → category axis → both → none. Gridlines that appear
      are light grey, 1 pt. On a bar chart, "value axis" gridlines run vertically.
- [ ] **Axis labels** cycles none → X axis → Y axis → both → none. Labels that reappear use the
      chart font, size 18.
- [ ] **Axis titles** cycles both → Y → X → none → both. The X title is centred under the chart.
      The chart grows into the freed space and shrinks to make room; it never overlaps the title,
      subtitle or source box. Text typed into a title box that stays is kept.
- [ ] Axis titles on a funnel only toggles the X title; on a pie it says the chart has no axes.
- [ ] Axis titles on a native, unstyled chart says to create or apply style first.
- [ ] After **Axis titles**, the X-axis title field in the Chart text panel enables and disables to
      match, and the chrome is still one group (drag it).
- [ ] **Legend** turns the legend on (top row, axis font) and off. On a pie or donut the plot
      re-centres. On a single-series column chart and a treemap it explains why not.
- [ ] **Data labels** with "All series" cycles none → outside end → inside centre → none.
      Outside labels are dark; inside labels are white on dark fills and dark on light fills
      (needs ExcelApi 1.16; on older hosts they stay dark).
- [ ] Data labels on a stacked column cycles none ↔ inside only.
- [ ] Data labels with one series picked changes only that series. The series list follows the
      selected chart without a refresh.
- [ ] Data labels on a pie: inside labels contrast per slice.
- [ ] **Spike:** gridlines on an axis hidden with Axis labels → none: record whether they can still
      be toggled. (The VBA had to show the axis temporarily.)
- [ ] A rebuild (choose the same type again) keeps an X-axis title that was switched on.

### 6.4 Plan phase 4 — annotation, point fill, logo

**Z-order (affects every chart).** The chrome group is now sent to the back after grouping.

- [ ] On a new chart of each family, the chart is fully visible in front of the white canvas, and
      the title, subtitle and source text are visible. Record the result per host: if any host
      puts the group in front of the chart, the canvas would hide the plot.

**Annotation** (Chart tools → Annotation at → Add annotation)

- [ ] "Plot centre" adds an "Annotation" box centred on the plot, in the axis font. Click it on the
      sheet to type.
- [ ] Picking a series selects its last point by default; the box lands just above-right of that
      point. Pick another point: the box follows it. Points without a data label before still have
      none afterwards.
- [ ] Repeated clicks stack boxes numbered Annotation1, Annotation2, … ; none replaces another.
- [ ] Annotations stay in front of the chart and are **not** in the chrome group. Restyling or
      toggling axis titles leaves them where they are.
- [ ] Export includes the annotations on top of the chart.
- [ ] Treemap: "Plot centre" lands near the middle of the chart (plot area may be unreadable).

**Point fill** (Color Picker)

- [ ] On a column chart, choosing a series shows a Point list ("Whole series" + categories). A
      colour applied to one point changes only that bar; "No fill" clears just that bar.
- [ ] On a scatter or line-with-markers chart, a point's **marker** takes the colour.
- [ ] On a plain line chart and on pie/treemap there is no Point list (pie already lists slices).

**Logo** (off by default)

- [ ] With `logoEnabled = false` nothing changes.
- [ ] Set `logoEnabled = true` and paste a small PNG as base64 into `src/config/logo.ts`, rebuild
      and restyle a chart: the logo sits bottom-right, does not overlap the source text, is part
      of the chrome group, and appears in the export.

### 6.5 Plan phase 5 — ribbon tab

The two Home-tab buttons are replaced by an **ExcelStarch** tab. Ribbon buttons run commands from
a hidden page (`commands.html`); the panes are unchanged. Re-register or re-upload the manifest
first: Office caches ribbons, so an old one can linger until the add-in is removed and re-added.

- [ ] The ExcelStarch tab shows four groups: Charts, Chart tools, Colours, Output. Every button has
      an icon at the normal and the large ribbon size.
- [ ] **Charts**: each family menu lists its types; each creates or retypes a chart like the
      matching pane button. **Apply style** matches the pane's button.
- [ ] **Chart tools**: each toggle gives the same result as the pane's button. Data labels act on
      all series; Annotation goes to the plot centre.
- [ ] **Colours**: in the Color Picker, apply a non-default ramp, diverging ramp and Rainbow
      palette. Close the pane. **Ramp**, **Diverging** and **Palette** re-apply those choices on
      another chart (the command re-reads the workbook settings). **Invert** reverses the colours.
- [ ] **Output → Picture** inserts the composed picture next to the selected chart.
- [ ] **Chart Builder** and **Color Picker** open their panes.
- [ ] With no chart selected, a toggle opens a small dialog saying so; **OK** closes it. The
      button does not stay busy.
- [ ] A command that fails leaves Excel responsive (no spinning button) — every command completes
      its event.

---

## 7. Requirement-set gating

The add-in requires **ExcelApi 1.14** (worksheet shapes are 1.9, but the chrome is looked up with
`shapes.getItemOrNullObject`, which is 1.14). The value lives in `MIN_EXCEL_API` in
`src/excel/session.ts` and in the `<Requirements>` element of `manifest.xml`; keep them equal.
Anything newer must be gated with `isSetSupported`, not assumed — today that is reading a fill
colour back (ExcelApi 1.16), used by Invert on fill- and point-coloured charts.

- [ ] On a host **with** ExcelApi 1.14: shape/chrome features work.
- [ ] On a host with 1.14 or 1.15: **Invert** on a column or pie chart shows the "needs a newer
      Excel" message and changes nothing; Invert on a line chart works.
- [ ] On a host **without** ExcelApi 1.14 (an older Excel, or simulate by temporarily gating a feature
      off): the add-in shows the **graceful unsupported-host message** (Fluent `MessageBar`) and does
      **not** throw or build a broken chart.

## 8. Cross-platform smoke

The add-in must work on Excel on the web, Windows, and Mac. For each platform you have access to,
run at least §2 (loads) + §3.1–3.2 (create + chrome):

- [ ] **Excel on the web** — loads, creates a branded grouped chart.
- [ ] **Excel for Windows** — same.
- [ ] **Excel for Mac** — same; double-check fonts/italic render (named italic families differ on Mac,
      which is why we use `font.italic`).

## 9. Reporting in the PR

Per the testing rule, **interop changes ship with documented sideload steps in the PR description.**
When a PR touches `excel/` or `chrome/`:

- State which sections of this guide you ran and on which platform(s).
- Paste/attach a screenshot of the resulting chart.
- Note any checklist item that could not be verified (e.g. "Mac not available").

`npm test` green is **necessary but not sufficient** when interop changed — sideload too.
