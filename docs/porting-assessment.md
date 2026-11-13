# Porting assessment: ExcelStarch VBA → Office JS

*Written 2026-11-13. Compares the VBA add-in (`C:\projects\ExcelStarch`, 25 modules, ~5,200 lines)
with this repository at commit `c6a9a13`.*

This document answers one question: which parts of the VBA add-in can be rebuilt in the Office JS
port, which cannot, and why. The follow-up work is ordered in
[implementation-plan.md](implementation-plan.md).

**How this was checked.** Both codebases were read in full. Every statement about what Office JS
can do was checked against the current `@types/office-js` definitions (member names and the
requirement set each one needs). Nothing was run in Excel. Where behaviour can only be confirmed
in a live host, the text says "needs a spike".

Colour, font and palette differences between the two (teal palette and Calibri versus ROOS and
Verdana) are configuration and are ignored here.

---

## 1. Summary

- **The core is ported; the tools around it are not.** Chart creation for 8 of the VBA's 16 chart
  buttons, the title/subtitle/source text, palette and ramp recolouring, per-series fill and
  last-used persistence work. The five toggles, export, annotation, the type-agnostic restyle,
  the other eight chart variants and the ribbon are not built.
- **Most of what is missing can be ported.** All five toggles, the type-agnostic restyle, every
  missing chart variant except lollipop, annotation, the logo and a ribbon tab are covered by
  APIs that exist today.
- **Four things cannot be ported, and they shape everything else:**
  1. Office JS cannot put text boxes or pictures *inside* a chart, and cannot group a chart with
     shapes. A chart and its title/source text are therefore always separate objects.
  2. Office JS cannot see which series or data point the user has clicked. Anything the VBA did
     "to the selected series" needs a dropdown instead.
  3. Office JS cannot write files or export vector formats. Export is limited to PNG/JPEG built in
     the task pane; SVG and PDF are gone.
  4. Chart fills have no transparency, and chart lines have no arrowheads. Semi-transparent
     bubbles and the VBA lollipop technique cannot be reproduced.
- **The biggest practical gap is not on the VBA feature list.** Because of point 1, a chart made
  by the web add-in cannot leave Excel with its title and source attached: copying the chart to
  PowerPoint copies only the plot. In the VBA version this worked for all classic charts. A
  composite "export / copy as picture" is the only fix, which makes export more important in the
  port than it was in the original.
- **The port has three defects that should be fixed before adding features**: line and scatter
  series are probably not brand-coloured, the declared minimum Excel version is lower than what
  the code needs, and restyling a chart wipes the title text the user typed. Details in §5.

---

## 2. How the two add-ins are built

| | VBA add-in | Office JS port |
|---|---|---|
| Runs on | Excel for Windows only, macros enabled | Windows, Mac and web; no macros |
| User interface | Custom ribbon tab, 10 groups, ~60 one-click buttons | Two task panes opened from two Home-tab buttons |
| How it finds its target | `ActiveChart` and `Selection` (chart, series or point) | The selected chart only (`getActiveChartOrNullObject`) |
| Where title/source/logo live | Inside the chart (`cht.Shapes`) for classic charts; worksheet shapes grouped with the chart for treemap and box & whisker | Worksheet shapes for every chart type, grouped with each other but not with the chart |
| Layout model | Chart is the 600×600 canvas; the plot area is positioned inside it | A white 600×600 rectangle is the canvas; the whole chart object is shrunk into a band on top of it |
| State | Module variables, document properties, Windows registry | `Office.context.document.settings` per workbook |
| Tests | `Debug.Assert` harness for pure functions | Jest for `config/` and `logic/`; manual sideload checklist for the rest |

The port took the VBA's second pipeline (`modEngineExChrome`, written for treemap) and made it
the only pipeline. That was the right call: the first pipeline (`modEngineBuilder`, in-chart
shapes) has no Office JS equivalent at all.

The port's layering is sound. `config/` and `logic/` are pure and tested, `excel/` and `chrome/`
hold all host calls, each action is one `Excel.run` with explicit sync points, and nothing throws
into the UI. New features fit into this structure without changing it.

---

## 3. Hard limits of Office JS

These are platform limits, not gaps in the port. No amount of work removes them.

### 3.1 No shapes inside a chart, and no grouping a chart with shapes

`Excel.Chart` has no shapes collection. A chart is not a member of `worksheet.shapes`, so
`shapes.addGroup` cannot include it.

Consequences:

- **The chart and its text move separately.** Dragging the chrome group leaves the chart behind,
  and the reverse. Re-running the builder snaps the chart back onto the canvas.
- **Copy and paste loses the chrome.** Copying the chart into Word or PowerPoint brings only the
  plot. The user has to multi-select the chart and the group, and they paste as two objects.
- **`chart.getImage()` returns the plot only**, without title, subtitle or source.
- **Chart sheets are out.** Office JS cannot reach charts on chart sheets, and worksheet shapes
  need a worksheet. The VBA supported chart sheets for classic charts.

### 3.2 No access to the selection inside a chart

Office JS can tell that a chart is selected. It cannot tell whether a series, a data point, the
plot area or the legend is selected within it. `getActiveShapeOrNullObject` exists for worksheet
shapes, but only from ExcelApi 1.19.

Consequences for VBA features that read `Selection`:

| VBA behaviour | What the port can do instead |
|---|---|
| Fill colour applies to the clicked series or point | Series chosen from a dropdown (already built). Points need a second dropdown. |
| Data-label toggle applies to the clicked series only | "All series" or a series from a dropdown |
| Annotation appears next to the clicked point | Series and point chosen from dropdowns |

### 3.3 No file access and no vector export

The VBA export offered PNG, GIF, JPG, BMP, SVG and PDF through a Save As dialog.

- Office JS returns a chart as a base64 **PNG** and nothing else. A shape can be returned as PNG,
  JPEG, GIF, BMP or SVG, but a chart is not a shape.
- There is no Save As dialog and no file system access. A file can only be offered as a browser
  download from the task pane. That works in Excel on the web. Whether it works in the desktop
  task pane on Windows and Mac **needs a spike**; it has been unreliable.
- **SVG and PDF export of a chart cannot be ported.** GIF and BMP are pointless to rebuild.
- The registry setting that remembered the last format becomes a workbook or browser setting.

One thing improves: `chart.getImage(width, height)` accepts a pixel size, so the port can export
at a higher resolution than the VBA's screen-resolution path for treemap groups.

### 3.4 Missing formatting properties

| Missing in Office JS | VBA feature affected | Verdict |
|---|---|---|
| Transparency on chart series fills (`ChartFill` has only set, get and clear) | Bubble charts at 50% transparency; `Fill:Name\|0.5` tags | Cannot be ported. Bubbles will be opaque. |
| Arrowheads on chart lines (`ChartLineFormat` has only colour, style and weight) | Lollipop: the dot is an oval arrowhead on an error bar | Cannot be ported as built. See §4.2 for the alternative. |
| Reading a series fill colour before ExcelApi 1.16 | Invert ramp; contrast-coloured data labels | Portable on current Excel only |

### 3.5 Smaller differences that do not matter

- `Application.ScreenUpdating` has no equivalent and is not needed; batched syncs do the same job.
- `MsgBox` becomes a non-blocking message bar. This is an improvement.
- Split buttons do not exist in add-in ribbons. Menus do.
- A ribbon button's icon cannot change to show the last-used colour.

---

## 4. Feature-by-feature status

**Status** is what the port does today. **Verdict** is whether the gap can be closed.

### 4.1 Chart creation and restyle

| VBA feature | Status | Verdict |
|---|---|---|
| Create from selected range, or retype the selected chart | Ported | — |
| Single selected cell expands to its data region | Ported (the VBA relied on Excel doing this) | — |
| 600×600 canvas, font, no border, no native title | Ported | — |
| Title, subtitle, source/notes placeholders | Ported as worksheet text boxes | — |
| Y-axis title box on every chart with a value axis | Built for box & whisker only. The space for it is reserved on all axis charts but left empty. | Portable, small |
| Logo bottom-right | Removed on purpose in the rebrand | Portable (`shapes.addImage`, PNG or JPEG only, so an SVG logo must be rasterised first) |
| Per-type gridline and axis defaults | Ported | — |
| Series coloured from the palette | Ported for fills. Line and marker colours are not set (see §5.1). | Portable |
| Legend placed in a strip under the subtitle, styled | Only shown or hidden. Position and font are never set. | Portable (`legend.position`, `left`, `top`, `format.font`) |
| Chart-sheet support | Not ported | **Not portable** |

### 4.2 Chart types

| VBA chart | Status | Verdict |
|---|---|---|
| Column, bar (clustered) | Ported | — |
| Stacked and 100% stacked column and bar | Missing | Portable. The chart types exist. |
| Lollipop | Missing | **Not portable as built.** A vertical version (line chart with hidden line, round markers and error-bar sticks) is possible and needs a spike. A horizontal version would need helper data on the sheet. |
| Line | Ported, probably in Excel's default colours | Portable |
| Area (stacked), 100% area | The port creates a *non-stacked* area chart; 100% is missing | Portable |
| Pie | Ported | — |
| Donut with configurable hole size | Missing | Portable (`series.doughnutHoleSize`) |
| Scatter with per-series marker shapes | Ported without marker styling | Portable (`markerStyle`, `markerSize`, marker colours) |
| Bubble | Missing | Portable, but opaque (§3.4) |
| Treemap, box & whisker | Ported | — |

Per-type finishing that the VBA applies and the port does not, all portable:

- gap width and overlap on bar and column (`series.gapWidth`, `series.overlap`)
- tick marks off or outside per type (`axis.majorTickMark`, `minorTickMark`)
- line and area starting on the axis (`axis.isBetweenCategories`)
- shadows removed (`series.showShadow`)
- data-label font on treemap and box & whisker
- square, centred plot area for pie and donut (`chart.plotArea`, ExcelApi 1.8)

A comment in `chromeLayout.ts` says Office JS has no plot-area object. That is out of date:
`chart.plotArea` has had position and size since ExcelApi 1.8. The band design is still the right
default, because treemap-family charts ignore plot-area settings, but the plot area is available
where it helps.

### 4.3 Apply Chart Style (restyle any chart without changing its type)

**Status:** missing. The port restyles a chart only by picking a type, which also retypes it.

**Verdict:** portable. `chart.chartType` is readable, so the VBA's four-way classification (pie,
filled-with-axes, line/scatter, other) carries over. This also covers sunburst, waterfall, funnel
and histogram, which the VBA handled only through this button.

### 4.4 Colour tools

| VBA feature | Status | Verdict |
|---|---|---|
| Palette order toggle (Contrasting / Rainbow) | Ported, and now remembered per workbook | — |
| Single-hue ramps | Ported | — |
| Diverging ramps with neutral centre | Ported. The port offers every ordered pair of ramps; the VBA offered nine. | — |
| Invert ramp | Ported, but needs ExcelApi 1.16 (§5.2) | — |
| Fill one series | Ported with a dropdown | — |
| Fill one data point | Missing | Portable with a point dropdown |
| Line and scatter series take the colour on the line | Missing. Every colour tool writes fills only. | Portable |
| Fill transparency | Missing | **Not portable** |
| Last-used colour, ramp and diverging pair | Ported | — |

### 4.5 Toggles

None of the five toggles is built. All are portable.

| Toggle | Office JS API | Note |
|---|---|---|
| Gridlines: none → horizontal → vertical → both | `axis.majorGridlines.visible` and line format | Reading gridlines on a hidden axis needs a spike |
| Axis labels: none → X → Y → both | `axis.tickLabelPosition` | Direct |
| Legend on/off | `chart.legend.visible`, position, font | Simpler than the VBA: the legend is inside the chart object |
| Axis titles: both → Y → X → none | Worksheet text boxes plus chart band resize | Same approach as the VBA: state is read from which boxes exist |
| Data labels: none → outside → inside | `series.hasDataLabels`, `dataLabels.position`, font | "Selected series" becomes a dropdown. Contrast colour needs ExcelApi 1.16 or a recomputed colour. |

### 4.6 Annotation

**Status:** missing. **Verdict:** portable with a changed workflow.

The box becomes a worksheet text box over the chart. Placement at the plot centre is direct
(`chart.plotArea.insideLeft` and friends). Placement next to a data point can reuse the VBA's
trick of reading the point's data-label position, which Office JS exposes; this needs a spike.
The point is chosen from dropdowns, not by clicking (§3.2).

### 4.7 Export

**Status:** missing. **Verdict:** partly portable; see §3.3.

A PNG or JPEG of the full chart with its chrome can be assembled in the task pane: render the
chrome group with `shape.getAsImage`, render the chart with `chart.getImage`, and draw one over
the other on an HTML canvas. Delivery is a download on the web. On desktop it may have to be
"insert as a picture on the sheet" or "copy to clipboard". Three spikes decide the design: image
quality of the group render, transparency of the chart render, and download support in the desktop
task pane.

### 4.8 Ribbon

**Status:** two Home-tab buttons that open task panes.

**Verdict:** largely portable. Add-in manifests support a custom tab, groups, buttons and
dropdown menus, and a button can run a function without opening a pane. A tab close to the VBA's
is possible: chart-type menus, Apply Style, the five toggles, Export, and "reapply last palette /
ramp". Limits:

- no split buttons, and no icon that changes to the last-used colour
- commands that need a parameter the ribbon cannot carry (a series, some text) stay in the pane
- a function command cannot show a message; errors must open the task pane to be seen
- sharing state between ribbon commands and the pane needs the shared runtime

This is a user-experience choice more than a porting question. The task panes work. The ribbon
restores the VBA's one-click feel and matters most for the toggles, which are meant to be clicked
repeatedly.

### 4.9 Infrastructure

| VBA part | Status |
|---|---|
| Config constants and derived geometry | Ported, tested |
| Pure logic (contrast, ramps, palette order, fill tags, plot geometry) | Ported, tested. `contrastColorForFill` is ported but nothing calls it yet. |
| Test harness | Replaced by Jest |
| Messages | Replaced by message bars |
| Embedded logo, base64 decode to temp file | Not needed; an image can be added directly from base64 |
| Build and signing of the `.xlam`, Trust Center guidance | Replaced by static hosting and a manifest |

---

## 5. Defects in the current port

Found by reading the code. Each should be confirmed by sideloading.

### 5.1 Line and scatter series are probably not brand-coloured

`colorSeriesByPalette` and `applySeriesFills` call `series.format.fill.setSolidColor` for every
chart type. A line series carries its colour on the line, and a scatter series on its markers.
The VBA handled this with a separate LINE mode and explicit marker colours. The port's own
comments mark this as an unfinished seam. Line and scatter charts, and every colour tool applied
to them, likely show Excel's default colours.

### 5.2 The declared minimum Excel version is too low

The manifest and `shapesSupported()` require ExcelApi 1.9. The code also calls:

- `shapes.getItemOrNullObject` — ExcelApi **1.14** — on every chart creation and text edit
- `fill.getSolidColor` — ExcelApi **1.16** — for Invert

On an Excel between 1.9 and 1.13 the add-in loads, passes its own check, and then fails on the
first chart. Volume-licensed perpetual Office versions are the likely place to meet this, and
those are common in the locked-down environments the port exists for. Check the exact version
table before choosing the new minimum.

### 5.3 Restyling wipes the user's text

Every build deletes and recreates the chrome with placeholder text. Changing a column chart to a
bar chart resets the title, subtitle and source. The VBA did the same, but there it was one of
many paths; here, the retype path is the only restyle path, so users will hit it constantly. Once
toggles exist, every layout change would trigger it.

### 5.4 Smaller issues

- The chart-text panel only knows the chart created in the current pane session. Reopen the pane
  and the fields are disabled until a chart is rebuilt (which wipes its text, §5.3).
- `writeChromeText` looks for shapes on the active worksheet, not the chart's worksheet.
- The Y-axis title field is permanently disabled, including for box & whisker, which has the box.
- The Color Picker needs a manual Refresh after selecting a different chart.
  `charts.onActivated` (ExcelApi 1.8) can do this automatically.
- "Area" creates an overlapping area chart; the VBA creates a stacked one. Later series hide
  earlier ones.
- `docs/testing-manual.md` §3.2 says dragging any element moves the chart and chrome together.
  The code and §3.1 above say otherwise.

---

## 6. What the port gains

To keep the comparison fair, the port is not only a subset:

- runs on Mac and in the browser, and where macros are blocked
- updates reach every user on the next load, with no file to redistribute
- one layout pipeline for all chart types, where the VBA has two
- real unit tests in CI
- non-blocking errors, per-workbook settings, and a remembered palette order
- export, once built, can exceed screen resolution

---

## 7. Verdict

A complete port is not possible. A port that covers everyday use is, with these permanent
differences:

1. Chart and chrome are two objects. Moving and copying need care, and a composite export is the
   answer for getting charts out of Excel.
2. Series and points are picked from dropdowns, not by clicking.
3. Export is PNG/JPEG only.
4. No transparency, no VBA-style lollipop, no chart sheets.

Everything else on the VBA feature list can be rebuilt on the existing architecture. The order to
do it in is in [implementation-plan.md](implementation-plan.md).
