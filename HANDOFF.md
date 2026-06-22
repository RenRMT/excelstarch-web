# Handoff — ExcelStarch Office.js conversion

This repo is the in-progress conversion of the **ExcelStarch VBA add-in**
(`https://github.com/RenRMT/ExcelStarch.git`) into a cross-platform **Office.js web add-in**
(TypeScript / React + Fluent UI / Webpack / Jest).

The full design is in [`docs/conversion-plan.md`](docs/conversion-plan.md). Read it first.

## Status

| Step | State |
|------|-------|
| New repo initialized (`.gitignore`, git) | ✅ done |
| Pinned VBA snapshot in `reference/vba-source/` + `reference/PROVENANCE.md` | ✅ done |
| `reference/spec.json` extracted (BGR→hex, 99/99 cross-checked vs VBA comment hex) | ✅ done |
| Scaffold with `yo office` (TS / React / Fluent UI / Webpack) | ✅ done |
| Port `config/` + `logic/` from `spec.json` | ✅ done |
| Jest suite reproducing `modTestHarness` (Phase 0 exit gate) | ✅ done — **Phase 0 complete; reviewed pause point** |
| Phase 1 MVP (bar + column + worksheet chrome + chart-text panel) | ⬜ not started |

**Phase 0 is complete.** The add-in is scaffolded; the pure `config/` and `logic/` layers are
ported from `spec.json`; the Jest suite mirrors every `modTestHarness` assertion. Exit gate met:
`tsc --noEmit`, `npm run lint`, and `npm test` (33 tests) are all green. The port was code-reviewed
for value/behavior fidelity — no transcription errors, all functions verified equivalent to the VBA.
The pause point is now **Phase 1 MVP not started.**

## Source-of-truth decisions (do not re-litigate without reason)

- **VBA snapshot pinned to** branch `demo/inso-brand-colors` @
  `c6fd5df9b04ffac334ee87631228d80493d1601c` (INSO palette + logo). See
  [`reference/PROVENANCE.md`](reference/PROVENANCE.md). `modConfig.bas` / `modEmbeddedImages.bas`
  differ from `main` — do **not** port from `main`.
- **Port from [`reference/spec.json`](reference/spec.json)**, not by re-reading `.bas` files.
  It is the reviewed, syntax-decoupled input; the BGR→hex conversion is auditable there.
- **Chrome is built as worksheet shapes grouped with the chart** (generalizing
  `reference/vba-source/modEngineExChrome.bas`) — Office.js charts cannot host in-chart shapes.
- **Stack:** Yeoman `yo office`, TypeScript (strict), React, Fluent UI v9, Webpack, Jest.

## Next steps (Phase 1 MVP)

Phase 0 (scaffold + pure `config/`/`logic/` + Jest) is done — see the status table. Next:

1. `excel/`: `chartFactory` (resolve/create a 600×600 chart by name; replace VBA `GetTargetChart`)
   + `chartStyle` (non-shape formatting from `modEngineBuilder.bas`) for **bar and column** only.
2. `chrome/`: generalize the worksheet-overlay group (canvas + text boxes + logo, grouped with the
   chart) from `reference/vba-source/modEngineExChrome.bas` — the universal chrome pipeline.
3. `taskpane/` (React + Fluent UI): bar/column creators + the chart-text panel
   (Title/Subtitle/Figure/Y-axis/Source/Notes → `shape.textFrame.textRange.text`).
4. `commands/` + `manifest.xml`: ribbon Show-Task-Pane.
   **Phase 1 exit:** select range → Bar → branded 600×600 grouped chart with editable chrome text
   (manual sideload — interop is not unit-tested; see `docs/conversion-plan.md` → Verification).

The interop layers consume the pure layer already shipped. Two contract notes from the port:
`parseDivergingTag` and the fill-tag parsers do NOT upper-case/trim — the interop caller must
normalize input as the VBA `UCase$`/`Trim$` did. `getPaletteColor(i, useAltOrder)` takes the
palette-order flag as a parameter (the VBA module-level `m_useAltOrder`); the toggle lives in `excel/`.

## Repo layout

- `reference/` — committed; the auditable input to the port (snapshot + `spec.json` + provenance).
- `docs/conversion-plan.md` — the full approved design and phased plan.
- `src/config/`, `src/logic/` — the ported pure layer (no `Excel`/`Office` imports); `test/` mirrors it.
- `src/taskpane/`, `src/commands/`, `manifest.xml` — scaffold defaults from `yo office`; Phase 1
  reshapes these into the strict layering. `excel/`, `chrome/`, `export/`, `persist/` are not created yet.

## Notes

- A local Claude Code config exists at `.claude/` but is **git-ignored** and not part of this
  handoff.
- Nothing is published to a remote yet; this is the initial commit.
