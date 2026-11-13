# Implementation plan: remaining portable features

*Written 2026-11-13, against commit `c6a9a13`. Background and evidence are in
[porting-assessment.md](porting-assessment.md); section references like "§5.1" point there.*

## Status

| # | Phase | State |
|---|---|---|
| 0 | Fix the foundations | Built |
| 1 | Export and copy as picture | Built |
| 2 | Chart variants, finishing, Apply Chart Style | Built |
| 3 | Toggles | Built |
| 4 | Annotation, point fill, logo | Built |
| 5 | Ribbon tab | Built, as function commands without a shared runtime (see below) |
| 6 | Lollipop and other experiments | Not started: waiting for demand and a live-Excel spike |

"Built" means type-checked, unit-tested where pure, and passing the production build. None of the
interop has been sideloaded yet; the checks for each phase are in
[testing-manual.md](testing-manual.md) section 6. The spikes listed below are part of those checks.

Changes from the plan made during the build:

- **Phase 5 has no shared runtime.** A shared runtime needs the task pane and the commands to be
  one page, and the add-in has two panes. Ribbon commands therefore run from a hidden commands
  page, report failures in a small dialog, and cannot open a pane with a message.
- **Annotations are not in the chrome group.** The group is sent behind the chart so its white
  canvas cannot cover the plot; an annotation over the plot must stay in front, so it stands alone
  and export draws it as its own layer.
- **No chart band strip for the legend or for x-axis labels.** Both are drawn inside the chart
  object. The band reserves strips only for the axis-title boxes that exist.
- **Realign (phase 6)** needs no command: Apply style already snaps a moved chart back onto its
  canvas.

The plan covers every VBA feature that the assessment found portable, in priority order. Priority
is user value weighed against certainty: defects first, then the feature that closes the largest
functional hole, then breadth, then convenience.

| # | Phase | Why here | Size |
|---|---|---|---|
| 0 | Fix the foundations | Current charts are wrong or fragile | S–M |
| 1 | Export and copy as picture | Without it a chart cannot leave Excel with its title | M–L, spike first |
| 2 | Chart variants, finishing, Apply Chart Style | Completes the core "make a chart" job | M |
| 3 | Toggles | The VBA's main post-creation tools | M |
| 4 | Annotation, point fill, logo | Useful, lower frequency | S–M |
| 5 | Ribbon tab | Restores one-click use; depends on 2 and 3 | M |
| 6 | Lollipop and other experiments | Uncertain or low demand | S each, spike first |

Sizes: S is about a day, M a few days, L a week or more, for someone who knows the codebase.

**Rules that apply to every phase** (they are the existing conventions; keep them):

- Decisions go in `src/logic/` or `src/config/` as pure functions with Jest tests. Host calls go
  in `src/excel/` or `src/chrome/`.
- One `Excel.run` per user action, no `sync` inside loops, return a `RunResult`.
- Every API newer than the declared minimum is gated with `isSetSupported` and degrades with a
  message.
- Each phase adds its checks to `docs/testing-manual.md`. Items marked **Spike** are questions
  only a live Excel can answer; do them first and on Windows desktop, Mac and web where possible.

---

## Phase 0 — Fix the foundations

Small changes that make the existing features correct. Later phases build on all of them.

### 0.1 Describe each chart kind in one table

Chart behaviour is currently spread over three places: a `switch` in `chartFactory.ts`, a record
in `chartFlow.ts`, and an inline `kind === "pie" || kind === "treemap"` test. Phase 2 nearly
doubles the number of kinds, so consolidate first.

- Add `src/config/chartSpecs.ts` (pure): one entry per `ChartKind` with
  - `excelType`: the chart type as a string literal, so the file stays free of `Excel` imports
  - `colourTarget`: `"fill"`, `"line"`, `"marker"` or `"point"`
  - `defaults`: the existing `ChartDefaults`
  - `finishing`: optional per-type settings, filled in during Phase 2
- Make `chartFactory.ts` and `chartFlow.ts` read from it. Delete the three scattered copies.
- Jest: every `ChartKind` has a spec; gridlines are a subset of displayed axes (the invariant
  already documented in `chartDefaults.ts`).

### 0.2 Colour line and scatter series correctly (§5.1)

- In `seriesColorer.ts` and `seriesRecolorer.ts`, branch on `colourTarget`:
  - `fill` → `format.fill.setSolidColor` (today's behaviour)
  - `line` → `format.line.color` and `format.line.weight` (2 pt, as in the VBA)
  - `marker` → `markerBackgroundColor` and `markerForegroundColor`
  - `point` → per-point fill (today's pie and treemap path)
- The Color Picker works on any selected chart, not only ones this add-in made. Add a pure
  `src/logic/chartClass.ts` that maps an Excel chart-type string to a colour target, mirroring
  the VBA's `GetStyleColorMode` and `IsPieChartType`. `colourFlow.ts` loads `chart.chartType`
  and uses it. Phase 2's Apply Chart Style reuses the same function.
- Invert must read and write the same property it colours. `format.line.color` is readable
  without `getSolidColor`.
- Jest: classification table. Sideload: line, scatter, and each colour tool on each.

### 0.3 Declare the real minimum version (§5.2)

- Decide the minimum. Recommended: ExcelApi **1.14**, which makes `getItemOrNullObject` on shapes
  legitimate. Check Microsoft's requirement-set table for which Office versions that excludes. If
  1.9 must stay, replace the three `shapes.getItemOrNullObject` uses with a load of
  `shapes.items/name` and a lookup.
- Update `manifest.xml`, `session.ts` and the README to the chosen value.
- Gate Invert on 1.16 with a clear message, or remove the need to read fills: the add-in assigns
  every colour itself, so it can store the last applied fill list per chart in document settings
  and reverse that.

### 0.4 Keep the user's text across rebuilds (§5.3)

- In `chartFlow.ts`, load the text of the existing title, subtitle, source and Y-axis boxes in the
  sync that already checks for the old canvas. Pass the values to `buildChrome`, which uses them
  in place of the placeholders.
- Do the same for any chrome shape added later (X-axis title, annotations). The simplest rule:
  rebuilds move and restyle shapes that exist and create only the ones that are missing. Consider
  switching to that rule here; it removes the delete-and-recreate cycle entirely.

### 0.5 Smaller fixes (§5.4)

- **Chart-text panel target.** Resolve the target from the selected chart when the pane opens and
  on `charts.onActivated`, falling back to the last created chart. Load the current text into the
  fields.
- **Right worksheet.** `writeChromeText` should find the chart by name across worksheets, or take
  the chart's worksheet from the resolver above.
- **Y-axis field.** Enable it when the `_YAxisTitle` shape exists.
- **Auto-refresh.** Register `worksheet.charts.onActivated` / `onDeactivated` in the Color Picker
  and drop the manual Refresh (keep the button as a fallback).
- **Legend.** When the legend is on, set its position (top), font size and colour, as the VBA
  does. Stop reserving a legend strip in the chart band: the legend is inside the chart object, so
  the strip is dead space.
- **Docs.** Correct `testing-manual.md` §3.2 (the chart is not in the group) and the comment in
  `chromeLayout.ts` that says there is no plot-area object.

**Done when:** line and scatter charts are brand-coloured, a retype keeps the title, the add-in
refuses to load (or degrades cleanly) below its real minimum, and the text panel works on an
existing chart after reopening the pane.

---

## Phase 1 — Export and copy as picture

The largest hole in the port: a finished chart cannot be taken out of Excel with its title and
source (assessment §1, §3.1). Target is a composite PNG or JPEG of canvas, chart and chrome.

### 1.0 Spikes (do these before any design)

1. **Group render.** Does `shape.getAsImage("PNG")` on the chrome group return the white canvas
   with the text, as Excel draws it? At what pixel size for a 600 pt group?
2. **Chart render.** Does `chart.getImage(w, h)` keep the transparent chart area, or flatten it to
   white or black? Is the output sharp at 2× and 3×?
3. **Delivery.** From the task pane, on Windows desktop, Mac and web: does a blob download work?
   Does `navigator.clipboard.write` with a PNG work?
4. **Fonts.** If text has to be drawn by the pane (fallback below), is the chart font available
   in the task pane's browser on each platform?

### 1.1 Compose

- `src/logic/exportLayout.ts` (pure): given the canvas size, the chart band and a scale factor,
  return pixel rectangles for each layer. Jest-tested.
- `src/excel/exportFlow.ts`: resolve the selected chart, find its chrome group by name, request
  both images, return them with the geometry. If the chart has no chrome (not made by this
  add-in), export the chart image alone.
- `src/taskpane/export/compose.ts`: draw the layers onto an HTML canvas and return a blob.
- **Fallback if spike 1 fails:** draw the chrome in the pane from the shapes' text and the known
  geometry and fonts. Needs a small word-wrap routine and gives slightly different text
  rendering from Excel's.
- Annotations (Phase 4) are worksheet shapes too; if they are added to the group they come along
  for free.

### 1.2 Deliver

Build what the spikes allow, in this order of preference:

1. **Download** as PNG or JPEG, with scale 1×, 2× or 3×. Default file name from the chart title.
2. **Copy to clipboard**, for pasting into Word or PowerPoint.
3. **Insert as picture on the sheet** (`shapes.addImage`), next to the chart. Always possible. The
   user can then copy it or use Excel's own *Save as Picture*.

Remember the last format and scale in the existing settings bag.

### 1.3 UI

An Export section in the Chart Builder pane: format, scale, and one button per available delivery
method. Say plainly in the UI and README that SVG and PDF are not available (§3.3).

**Done when:** a branded chart of each type can be turned into one image that matches what is on
the sheet, on every platform the add-in supports, by at least one delivery method.

---

## Phase 2 — Chart variants, finishing, Apply Chart Style

### 2.1 Variants

Add to `ChartKind` and `chartSpecs.ts`. All chart types exist in ExcelApi 1.9 or earlier.

| New kind | Excel type | Notes |
|---|---|---|
| Stacked column, 100% stacked column | `columnStacked`, `columnStacked100` | overlap 100 |
| Stacked bar, 100% stacked bar | `barStacked`, `barStacked100` | overlap 100 |
| Stacked area (replaces today's "area"), 100% area | `areaStacked`, `areaStacked100` | matches the VBA; fixes §5.4 |
| Donut | `doughnut` | `colourTarget: "point"`, hole size 50% |
| Bubble | `bubble` | opaque; document the missing transparency (§3.4) |

UI: eight buttons become fifteen. Group them by family (Column, Bar, Line & Area, Pie, Scatter,
Complex) with a default and a small menu each, mirroring the VBA ribbon groups.

### 2.2 Per-type finishing

Extend `finishing` in the spec and apply it in `chartStyle.ts`. Values come from `modConfig.bas`.

| Setting | API | Applies to |
|---|---|---|
| Gap width 33, overlap −5 (100 when stacked) | `series.gapWidth`, `series.overlap` | bar, column |
| Tick marks none | `axis.majorTickMark`, `minorTickMark` | bar, column, area |
| Tick marks outside | same | line, scatter, bubble |
| Start on the axis | `categoryAxis.isBetweenCategories = false` | line, area |
| No shadow | `series.showShadow = false` | bar, column, line |
| Marker shape cycle (circle, square, triangle, X), size 9 | `series.markerStyle`, `markerSize` | scatter |
| Donut hole 50% | `series.doughnutHoleSize` | donut |
| Data-label font | `series.dataLabels.format.font` | treemap, box & whisker |
| Square centred plot | `chart.plotArea` position and size | pie, donut |

**Spike:** which of these the treemap family rejects. The VBA wraps them in error suppression for
that reason. In Office JS one bad queued write fails the whole batch, so apply the uncertain ones
in their own guarded sync.

### 2.3 Y-axis title on classic charts

Set `showYAxisTitle: true` for every kind with a value axis. The band already reserves the strip
(§4.1), so this is a one-line change per spec plus enabling the panel field (done in 0.5).

### 2.4 Apply Chart Style

Restyle the selected chart without changing its type (§4.3).

- Add `styleChartInPlace()` to `chartFlow.ts`: the existing flow minus `retypeChart`. Load
  `chart.chartType`, classify it with `logic/chartClass.ts` (from 0.2) into pie, filled-with-axes,
  line/scatter or other, and pick defaults the way the VBA's `ClassifyChart` does.
- For the "other" class (radar, stock, waterfall, funnel, sunburst, histogram and anything
  unknown): chrome, font and size only, then best-effort colouring in a guarded sync.
- UI: an "Apply style to selected chart" button above the type buttons.

**Done when:** every VBA chart button except Lollipop has a working equivalent, each chart matches
the VBA output side by side apart from colours and fonts, and any existing chart can be restyled
in place.

---

## Phase 3 — Toggles

All five VBA toggles (§4.5). Each is one small `Excel.run` on the selected chart.

### 3.1 Shared groundwork

- `src/logic/toggleCycles.ts` (pure, Jest): the four state machines, copied from
  `modEngineToggles.bas`:
  - gridlines: none → horizontal → vertical → both → none
  - axis labels: none → X → Y → both → none
  - axis titles: both → Y → X → none → both, clamped when an axis is absent
  - data labels: none → outside → inside → none; none ↔ inside for stacked types
- `src/excel/layoutState.ts`: read the current layout from the workbook, not from stored state,
  as the VBA does: which axis-title shapes exist, whether the legend is visible, the chart type.
- A `relayout(chart, state)` helper that recomputes the chart band from `plotAreaTopFor` /
  `plotAreaHeightFor` and moves existing shapes. This relies on 0.4: a relayout must never reset
  text.

### 3.2 The toggles

| Toggle | Implementation | Watch for |
|---|---|---|
| Gridlines | `majorGridlines.visible` plus line colour and weight on value and category axes | **Spike:** reading and writing gridlines when the axis itself is hidden. The VBA re-enables the axis temporarily. |
| Axis labels | `tickLabelPosition` `"None"` / `"NextToAxis"`, restyle font when shown | No-op on charts without axes |
| Legend | `legend.visible`, then position and font (from 0.5) | Message for single-series charts and treemap, as in the VBA |
| Axis titles | Create or delete `_YAxisTitle` and a new `_XAxisTitle` shape, regroup, relayout | X title is centred (`textFrame.horizontalAlignment`); add the suffix to `chromeNames.ts` |
| Data labels | `hasDataLabels`, `dataLabels.position`, font | Scope and contrast, below |

Data labels need two decisions:

- **Scope.** The VBA used the clicked series. Offer "All series" plus the series dropdown the
  Color Picker already has.
- **Contrast colour for inside labels.** `contrastColorForFill` is already ported and tested.
  It needs the fill colour: use `getSolidColor` where ExcelApi 1.16 is available, per point for
  pie and donut. Below 1.16, fall back to the dark label colour.

### 3.3 UI

A "Chart tools" section with five buttons. Each shows the state it moved to ("Gridlines:
horizontal"), since a cycle is hard to follow blind. These buttons are the main candidates for
the ribbon in Phase 5.

**Done when:** each toggle cycles through the VBA's states on every chart family, typed text
survives every toggle, and toggling never leaves the chart outside its canvas.

---

## Phase 4 — Annotation, point fill, logo

### 4.1 Annotation (§4.6)

- New chrome shape `<chart>_Annotation<n>`: a transparent text box, axis font size and colour,
  default text "Annotation". Runs stack, as in the VBA.
- Placement:
  - default: centre of the plot (`chart.left + plotArea.insideLeft + insideWidth / 2`, same for
    top). **Spike:** the treemap family may not report inside dimensions; fall back to the centre
    of the chart object.
  - next to a point: series and point dropdowns, defaulting to the last point. Read the point's
    data-label `left` / `top`, enabling the label for one sync and restoring it, as the VBA does.
    **Spike:** whether the position is reliable straight after enabling the label.
- Add the box to the chrome group so it moves and exports with the chrome. Rebuilds and relayouts
  must keep it (0.4).

### 4.2 Fill a single data point

Add a point dropdown to the element-fill row, shown when a series is chosen, writing
`points.getItemAt(i).format.fill`. For pie, donut and treemap the "series" level is meaningless,
so show points directly.

### 4.3 Logo (optional)

Removed deliberately in the rebrand, so build it behind a config switch.

- `src/config/logo.ts`: enabled flag, base64 PNG, aspect ratio, height as a share of the canvas.
- `chromeShapes.ts`: `addLogo` with `shapes.addImage`, bottom-right, `lockAspectRatio`.
  `addImage` takes PNG or JPEG only; convert an SVG logo at build time.
- Geometry: the source box is 80% of the canvas width, leaving 120 pt for the logo. Add a test
  that the two do not overlap for the configured sizes.
- `_LogoImage` is already listed as a legacy suffix; move it back to the live list when enabled.

---

## Phase 5 — Ribbon tab

Restores the VBA's one-click workflow (§4.8). Do it after Phases 2 and 3 so there are commands
worth a button.

1. **Shared runtime.** Add the runtime to `manifest.xml` and a `src/commands/commands.ts` entry
   that registers functions with `Office.actions.associate`. Both panes and the commands then
   share one JavaScript context. **Spike:** confirm both existing panes still behave under the
   shared runtime on all three platforms; it changes when the page loads and unloads.
2. **Custom tab** named ExcelStarch, with groups that follow the VBA layout:
   - one menu per chart family
   - Apply Chart Style
   - the five toggles
   - Export
   - Reapply last palette, ramp and diverging ramp
   - buttons that open the Chart Builder and Color Picker panes
3. **Icons.** The VBA `icons/` folder has a matching PNG for every command. Each needs 16, 32 and
   80 px versions under `assets/`.
4. **Feedback.** A function command has no message surface. On failure, open the task pane with
   the status message (`Office.addin.showAsTaskpane()`). On success, stay silent.
5. Keep in the panes everything that needs a parameter: chart text, ramp choice, series and point
   pickers, export options.

Check the current limits on controls per group and menu items in Microsoft's manifest reference
before laying out the tab.

---

## Phase 6 — Lollipop and other experiments

Only if there is demand. Each starts with a spike and may end there.

- **Lollipop (§4.2).** The VBA's arrowhead technique is not available. Try a vertical version:
  `lineMarkers` chart, line hidden, round markers in the palette colour, `yErrorBars` at minus
  100% without caps as the sticks. If error-bar formatting holds up, ship it as "Lollipop
  (vertical)". A horizontal version needs helper columns on the sheet and is not recommended.
- **Dedicated builders for sunburst, waterfall, funnel, histogram, pareto.** The VBA has none;
  Apply Chart Style (2.4) covers them. Add buttons only if users ask.
- **Keep chart and chrome together.** A "Realign" command that snaps the chart back to its canvas
  after either was moved is cheap (the build flow already does it). Listening for moves is not
  possible; there is no shape-moved event.

---

## Not planned

These cannot be ported (assessment §3) and should be stated in the README so users of the VBA
version know what to expect:

- text or logo inside the chart object, and chart-plus-chrome as one movable, copyable object
- acting on the series or point the user clicked
- SVG, PDF, GIF and BMP export, and a Save As dialog
- transparent series fills, including semi-transparent bubbles
- the horizontal lollipop chart as the VBA builds it
- chart sheets
