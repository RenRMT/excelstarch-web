# Manual sideload testing — the interop layers Jest cannot cover

The pure `config/` and `logic/` layers are unit-tested with Jest (`npm test`) — see
[`testing-with-jest`](../.claude/skills/testing-with-jest/SKILL.md). Everything that touches the
live Office host is **not** unit-tested and must be verified by sideloading the add-in into a real
Excel. This document is the canonical checklist for that.

> **Why manual?** `excel/`, `chrome/`, and `export/` call `Excel.run` and manipulate the workbook
> object model, charts, and worksheet shapes. Mocking the entire Office object model is brittle and
> proves nothing about real host behaviour — so we sideload instead. The decision boundary: if a
> function imports `Excel`/`Office`, it is sideload-tested, not Jested.

## Scope by phase

This guide is written so each section maps to a phased delivery exit in
[`conversion-plan.md`](conversion-plan.md). Run only the sections whose phase has shipped.

| Phase | What to sideload-test | Section |
|-------|-----------------------|---------|
| 0 | Nothing — no interop yet. `npm test` green is the whole gate. | — |
| 1 | Bar + column creation, branded styling, the worksheet-chrome group, chart-text panel. | [§3](#3-phase-1--chart-creation-styling-and-chrome), [§4](#4-phase-1--chart-text-chrome-panel) |
| 2 | Remaining chart types, full colour tooling, per-element fill, last-used persistence. | [§5](#5-phase-2--colour-tooling-and-breadth) |
| 3 | Toggles, restyle, export (canvas composite), annotation. | [§6](#6-phase-3--toggles-export-annotation) |
| every | Requirement-set gating, cross-platform smoke. | [§7](#7-requirement-set-gating), [§8](#8-cross-platform-smoke) |

---

## 1. Prerequisites

- A clean checkout with dependencies installed (`npm install`).
- A trusted dev certificate. The first `npm start` provisions one via `office-addin-dev-certs`; on
  Windows accept the prompt to install the CA. If the task pane shows a blank/red page, the cert is
  the usual cause — re-run `npx office-addin-dev-certs install`.
- Excel: at least one of **Excel on the web** (Microsoft 365), **Excel for Windows**, or
  **Excel for Mac**. Cross-platform parity is a project requirement, so test more than one when you can.
- The original VBA add-in (`ExcelStarch.xlsm` on the pinned `demo/inso-brand-colors` build) open in a
  second window for **side-by-side visual comparison** — the fidelity bar is "matches the VBA output."

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
   - [ ] Series use the **brand data palette** in order (Ocean, Coral, Sky, … = `colorData1..8`).
   - [ ] Gridlines match the type: **column → Y-gridlines only**, **bar → X-gridlines only**.
   - [ ] Axes/legend match the chart-type defaults (bar/column: both axes shown, legend off by default).

### 3.2 The worksheet-chrome group

Office.js charts cannot host in-chart shapes, so chrome is built as **worksheet shapes grouped with
the chart** (generalizing `modEngineExChrome.bas`). Verify:

- [ ] A white **canvas rectangle** sits behind the chart (sent to back).
- [ ] **Title, subtitle, figure-number, source/notes** text boxes appear at the branded positions
      with the correct fonts/sizes/colours (title 28pt `colorBrand1`, subtitle 22pt `colorBrand2`,
      figure 18pt `colorBrand3`, source 14pt `colorBrand3`).
- [ ] The **logo** appears bottom-right at the branded margin.
- [ ] Selecting any one element and dragging moves the **whole group together** (chart + canvas +
      text + logo) — i.e. they are grouped, named `ESChromeGroup_<chartName>`.
- [ ] Placeholder text matches the spec (`"Title in 28pt sentence case"`, `"Figure XX (optional)"`,
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
2. Type into the **Title / Subtitle / Figure / Y-axis / Source / Notes** fields.
3. Verify:
   - [ ] Each field writes to the matching shape's `textFrame.textRange.text` — the chart chrome
         updates live (or on apply, per the panel's design).
   - [ ] Empty field restores/keeps the placeholder, it does not crash.
   - [ ] Editing one field does not disturb the others or ungroup the chrome.
   - [ ] Italic uses `font.italic = true`, **not** a named italic font family (cross-platform).

---

## 5. Phase 2 — colour tooling and breadth

### 5.1 Remaining chart types

For each newly added type (area, scatter, pie, line, treemap, box & whisker, lollipop):

- [ ] Creates a 600×600 branded chart with the type-appropriate defaults (see `chartDefaults`).
- [ ] Chrome group builds correctly (box & whisker also gets the optional **Y-axis title box**).
- [ ] Pie/donut: legend sits below the subtitle without overlap; plot geometry matches the pie spec.

### 5.2 Palette / ramps / diverging

Use the multi-series fixtures. The **pure** ordering math is already Jest-proven; here you confirm it
is wired to the host correctly.

- [ ] **Palette** recolours series in brand order; the **Rainbow/Contrasting toggle** swaps slot 2↔
      Lavender and slot 6↔Coral (matches `getPaletteColor` with/without alt order).
- [ ] **Single-hue ramp** (e.g. Ocean): series go **darkest-first** (series 1 darkest), n=1..10.
      Spot-check n=3 → steps `[6,4,2]` of the ramp; n=10 → all ten darkest→lightest.
- [ ] **Diverging ramp** (e.g. `A|B`): dark→light on the left, light→dark on the right; **odd series
      count gets a grey centre** (`colorBrand4`), even count does not. Try n=5 (grey middle) and n=8.
- [ ] **Invert** reverses the current fill assignment across all series.
- [ ] Over-limit guard: >10 series for a single ramp, or >21 for diverging, shows the non-blocking
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

- [ ] After applying a ramp/fill, the **LASTUSED** path re-applies the same choice on the next run,
      and persists across close/reopen of the workbook (via `Office.context.document.settings`, the
      replacement for VBA `CustomDocumentProperties`).

---

## 6. Phase 3 — toggles, export, annotation

### 6.1 Toggles & restyle

- [ ] Each of the 5 toggles flips its target on the resolved chart (gridlines, axes, legend, …)
      without ungrouping chrome.
- [ ] **Restyle / ApplyChartStyle** re-applies branding to an existing chart.

### 6.2 Export (group-image composite — verify the primary path first)

`chart.getImage()` is **bare-chart JPEG only** (no chrome). The primary export rasterizes the chart,
temporarily adds it into the chrome group, images the **group** with `Shape.getAsImage` (ExcelApi 1.9,
native PNG/JPEG), then restores the sheet. An HTML-`<canvas>` composite is the fallback if a host's
group `getAsImage` doesn't capture child shapes. (Background: the chart can't be a real group member —
see the chrome-overlay skill.)

**Decide the path first (decision #0):**

- [ ] Export PNG and open the file: it shows the **chart pixels + white canvas + title/subtitle/figure/
      source/logo** all composited. If only the chrome (or an empty/partial image) appears, the host's
      group `getAsImage` doesn't capture children → the **canvas fallback** must be the primary path
      on that host; record it.

**Then verify the rest:**

- [ ] **Z-order:** chart sits over the white canvas with the text/logo legible on top, nothing clipped.
- [ ] **Cleanup is exact:** after export the on-sheet chart + chrome are **unchanged** — the group is
      still named `ESChromeGroup_<chartName>`, all members present, and **no leftover
      `<chartName>_ExportChartPic`**. Then re-run create/restyle and confirm it still finds and cleans
      the group (proves the group name was restored).
- [ ] **Failure path:** if `getAsImage` fails, the sheet is restored (chart intact, no temp picture)
      and a non-blocking error `MessageBar` shows — no crash.
- [ ] **PNG and JPEG** both download and open; the composite visually matches the on-sheet chrome.
- [ ] **Per host:** the download actually **saves** on Excel web (Edge/Chrome), Windows (WebView2), and
      Mac (WKWebView — most restrictive; note if a dialog fallback is needed).

### 6.3 Annotation (degraded — plot-centre box)

- [ ] No per-point anchor (the web API has no point pixel coords); annotation appears as a
      **plot-centre draggable box** that can be repositioned, per the degraded-feature decision.

---

## 7. Requirement-set gating

The add-in requires **ExcelApi 1.9** (worksheet shapes). Features must be gated with
`isSetSupported`, not assumed.

- [ ] On a host **with** ExcelApi 1.9: shape/chrome features work.
- [ ] On a host **without** ExcelApi 1.9 (an older Excel, or simulate by temporarily gating a feature
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
When a PR touches `excel/`, `chrome/`, or `export/`:

- State which sections of this guide you ran and on which platform(s).
- Paste/attach a screenshot of the resulting chart next to the VBA reference where fidelity matters.
- Note any checklist item that could not be verified (e.g. "Mac not available").

`npm test` green is **necessary but not sufficient** when interop changed — sideload too.
