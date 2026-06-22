# Convert ExcelStarch VBA Add-in → Office.js Web Add-in

> Reconstructed from the approved plan. The machine-local plan file was later reused for a
> separate task, so this in-repo copy is the canonical record of the conversion design.

## Repository structure & bootstrap (Phases 0–1)

### Where the new project lives
A **new, standalone git repo** at `c:/projects/excelstarch-web` (sibling of
`c:/projects/ExcelStarch`). Not a branch, submodule, or subfolder of the VBA repo — its own
history and remote. The VBA repo remains the upstream source of truth and is never modified.

### How the assistant reads the VBA source (sibling dir + pinned snapshot)
A **pinned, read-only snapshot** of the relevant `.bas` files lives in `reference/vba-source/`
so the port has a stable, narrow, self-contained input (no wandering the live VBA tree):

```
excelstarch-web/
├── reference/
│   ├── PROVENANCE.md     # source repo + exact commit/branch the snapshot came from
│   ├── spec.json         # extracted, reviewed values — the port's real input
│   └── vba-source/       # verbatim .bas copies needed for the port
└── src/ … test/ …
```

**Source-of-truth pin:** snapshot taken from the **`demo/inso-brand-colors`** branch @
`c6fd5df9b04ffac334ee87631228d80493d1601c` (INSO palette + logo). `modConfig.bas` and
`modEmbeddedImages.bas` differ materially between `main` and that branch — porting from
`main` would bake in the wrong palette/logo.

### Intermediate spec artifact: `reference/spec.json`
Before any TypeScript, VBA constants are extracted into one reviewed `spec.json`, then
`config/`+`logic/` are ported **from the JSON** (not by re-transcribing `.bas`). Contains:
brand colours (BGR `Long` + converted `#RRGGBB`), ramps (A–H × 10), diverging tags,
geometry, fonts/sizes, placeholder text, palette-order maps, enums. The BGR→hex conversion
is auditable in one place and was cross-checked against the VBA comment hex (99/99 match).

### Bootstrap sequence
1. `git init` the new repo; create its remote.
2. Populate `reference/vba-source/` from `demo/inso-brand-colors`; write `PROVENANCE.md`.
3. Extract `reference/spec.json` (BGR→hex here); **review pause** before porting.
4. Scaffold with `yo office` (TS / React / Fluent UI / Webpack).
5. Port `config/` + `logic/` from `spec.json`; write the Jest suite (Phase 0 exit).
6. Phase-1 MVP (bar + column + generalized chrome + chart-text panel), using
   `reference/vba-source/modEngineExChrome.bas` as the chrome reference.

---

## Context

ExcelStarch is a mature VBA add-in (25 modules) that turns a selected data range into a
house-styled 600×600pt chart: brand colours/fonts, gridline styling, and a chrome layer of
title / subtitle / figure-number / y-axis-label / source-notes text boxes plus a
bottom-right logo. Windows-only (registry prefs, file-save-dialog export).

Goal: a **cross-platform Office.js web add-in** (manifest, task pane + small ribbon,
TypeScript) preserving the branded charting identity and as much of the feature set as the
web platform allows.

**The one fact that shapes the whole port:** all chrome is built as shapes *inside the chart
object* via `cht.Shapes.AddTextbox`/`AddPicture` (modEngineBuilder.bas). **The Office.js
`Excel.Chart` API has no shape collection.** The escape hatch already exists:
`modEngineExChrome.bas` builds the same chrome as *worksheet* shapes grouped with the chart
(for chartex types). **The conversion promotes that worksheet-overlay pipeline to the only
chrome pipeline, for every chart type.**

Verified Office.js facts:
- Worksheet shapes ARE supported (`addTextBox`/`addImage`/`addGeometricShape`/`addGroup`/
  `setZOrder`) — ExcelApi **1.9+**.
- `chart.getImage()` returns base64 **JPEG only**, recently ignores resize args, and captures
  the bare chart — **not** worksheet chrome. No native PNG/GIF/BMP/SVG/PDF, no Save-As.
- Treemap and box & whisker chart types are exposed in the JS chart API.
- No `ActiveChart`/`Selection` "clicked series/point" model.
- Shape positions are in **points**, so the geometry transfers with no unit conversion.

## Decisions (confirmed with user)
- **Chrome fidelity:** worksheet-overlay group (closest to current look).
- **Priority features:** full colour tooling (palette / 8 ramps / diverging) **and**
  per-element fill/annotation (panel-driven, since the Selection model is gone).
- **Stack:** Yeoman `yo office`, TypeScript, React, Fluent UI v9, Webpack.

## Target architecture

Strict layering, dependency direction downward; bottom two layers never import `Excel`/`Office`:

```
taskpane (React/Fluent UI) ─┐
commands (ribbon onAction) ─┼─► excel/ + chrome/ + export/ ─► logic/ + config/ (PURE)
                            └─► persist/
```

```
src/
├── config/   brand.ts (BGR→hex), geometry.ts, chartDefaults.ts, enums.ts   (← spec.json)
├── logic/    colorRamp.ts, colorContrast.ts, colorFill.ts, colorSeries.ts
├── excel/    session, chartFactory, chartStyle, ramps, fill, toggles
├── chrome/   chromeBuilder, chromeShapes, chromeGroup, chromeLayout
├── export/   exportImage.ts (getImage + client-side canvas composite)
├── persist/  settings.ts (Office.context.document.settings)
├── taskpane/ React + Fluent UI components
└── commands/ commands.ts (Office.actions.associate)
manifest.xml  add-in commands + task pane; min ExcelApi 1.9
```

**UI split.** Ribbon stays tiny (Show Task Pane + optional one-click creators). Everything
galleried/stateful (chart gallery, palette, ramp + diverging pickers, 5 toggles, chart-text
panel, element selector, export) lives in the task pane. **Batching:** one `Excel.run` per
action, minimal `context.sync()`.

## Module-by-module conversion (summary)

- **Port as-is (pure):** modConfig → `config/brand.ts`; modConfigDerived → `geometry.ts` +
  `chartDefaults.ts`; modColorContrast → `logic/colorContrast.ts`.
- **Split (pure math + async apply):** modColorRamp, modColorFill, modColorSeries.
- **Two fates:** modEngineBuilder — non-shape format → `excel/chartStyle.ts`; in-chart shape
  chrome **deleted**, rebuilt on the worksheet in `chrome/`. `GetTargetChart` → `chartFactory.ts`.
- **Promote to core engine:** modEngineExChrome → `chrome/*` (the universal chrome pipeline).
- **Rearchitect:** modEngineStyle (annotation degraded — no per-point anchor); modEngineToggles
  (async, resolve chart by name); modColorFill target detection (panel-driven, no Selection).
- **Drop:** modAppState (no ScreenUpdating on web).
- **Simplify:** modEmbeddedImages → bundled `assets/logo.svg` + `addImage(base64)`.
- **Replace:** modRibbonHandlers → `commands.ts` + React handlers; modMessages → Fluent
  `MessageBar`; modExport → `export/exportImage.ts`; modTestHarness → Jest.

## Chrome-overlay strategy (core)

Recreate each chrome element as a worksheet shape over the chart, then group with the chart —
generalizing `BuildChartExChrome` to all types. Reuse `ChartExCanvasOrigin` (reuse canvas
position on re-run) and `PositionChartExChart` (inset chart into the plot band) as
`chromeLayout.ts`. Per action, in one `Excel.run`: create/size chart → add white canvas
(`setZOrder(sendToBack)`) → add text boxes + logo at canvas-offset positions → optional
y-axis-title box → sync → `addGroup([...])` named `ESChromeGroup_<chartName>` → sync.

**Chrome text editing moves to the task pane** (Title/Subtitle/Figure/Y-axis/Source/Notes
fields writing `shape.textFrame.textRange.text`). Risks & mitigations: group only
successfully-created shapes, warn on partial failure; keep the reuse-canvas-position trick so
re-runs snap chrome back; lock at 600×600 (offer "reset size" = re-run); gate ExcelApi 1.9
with `isSetSupported`; use `font.italic = true` not the named italic family.

## Phased delivery

- **Phase 0 — Foundation.** Scaffold; port `config/`+`logic/` with BGR→hex; Jest suite
  reproducing `modTestHarness`. *Exit:* `npm test` green; math proven identical to VBA.
- **Phase 1 — MVP.** `chartFactory` + `chartStyle` + generalized `chrome/` for **bar and
  column** only; ribbon Show-Task-Pane + task pane with those creators + chart-text panel.
  *Exit:* select range → Bar → branded 600×600 grouped chart with editable chrome text.
- **Phase 2 — Breadth + colour tooling (priority).** Remaining classic + chartex types; full
  colour tooling + per-element fill via element selector; last-used via `persist/settings.ts`.
- **Phase 3 — Toggles, restyle, export, annotation.** Toggles; `ApplyChartStyle`; export via
  client-side canvas composite; annotation as a plot-centre draggable box.

## Degraded features (with fallbacks)

1. **Export formats — severely degraded.** No PNG/GIF/BMP/SVG/PDF/Save-As; `getImage()` is
   bare-chart JPEG only. **Fallback (chosen):** reconstruct the composite client-side —
   `getImage()` for the plot, draw chrome+logo onto an HTML `<canvas>` using `geometry.ts`,
   export the canvas (PNG; PDF/SVG via a JS lib). Higher fidelity than VBA.
2. **Annotation on a point — degraded.** No per-point pixel coords → plot-centre box + drag.
3. **Selection-dependent ops — rearchitected.** Resolve target chart by enumerating
   `worksheet.charts`; replace click-to-select with a task-pane element selector.
4. **Ribbon richness — lost.** Nested menus / split galleries / 21-combo diverging → task pane.
5. **Modal messaging — gone.** `MsgBox` → non-blocking Fluent `MessageBar`.
6. **PDF/`ExportAsFixedFormat` + chart sheets — gone.** PDF only via the canvas route;
   embedded charts only.

## Verification

- **Pure layer:** `npm test` — Jest mirrors `modTestHarness` (ramp ordering, diverging
  layout, tag/payload parse, contrast thresholds, plot-area geometry).
- **Interop / chrome:** `npm start` sideload (web + desktop); per phase, create each chart
  type and confirm 600×600 canvas, brand colours, gridlines, and a grouped chrome layer
  matching the VBA output side-by-side. Full step-by-step checklists are in
  [`testing-manual.md`](testing-manual.md).
- **Colour tooling / toggles / export / host gating / cross-platform smoke:** per the phased
  exits above and the matching sections of [`testing-manual.md`](testing-manual.md).
