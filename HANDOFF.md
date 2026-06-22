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
| `reference/spec.json` extracted (BGR→hex, 99/99 cross-checked vs VBA comment hex) | ✅ done — **reviewed pause point** |
| Scaffold with `yo office` (TS / React / Fluent UI / Webpack) | ⬜ not started |
| Port `config/` + `logic/` from `spec.json` | ⬜ not started |
| Jest suite reproducing `modTestHarness` (Phase 0 exit gate) | ⬜ not started |
| Phase 1 MVP (bar + column + worksheet chrome + chart-text panel) | ⬜ not started |

The work stopped at the planned review pause: **`spec.json` is ready; nothing has been
ported to TypeScript yet.** This is a clean place to take over.

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

## Next steps (Phase 0 → 1)

1. `yo office` scaffold into this repo (TS + React).
2. Port `src/config/` (brand, geometry, chartDefaults, enums) and `src/logic/` (colorRamp,
   colorContrast, colorFill, colorSeries) **from `spec.json`**.
3. Write the Jest suite mirroring `modTestHarness` (cases listed in `docs/conversion-plan.md`
   → Verification). **Phase 0 exit:** `npm test` green.
4. Phase 1 MVP: `chartFactory` + `chartStyle` + generalized `chrome/` for **bar and column**,
   ribbon Show-Task-Pane + task pane with those creators and the chart-text panel.

## Repo layout

- `reference/` — committed; the auditable input to the port (snapshot + `spec.json` + provenance).
- `docs/conversion-plan.md` — the full approved design and phased plan.
- `src/`, `test/`, `manifest.xml` — to be created by the scaffold/port (see the plan for the
  target structure and strict layering).

## Notes

- A local Claude Code config exists at `.claude/` but is **git-ignored** and not part of this
  handoff.
- Nothing is published to a remote yet; this is the initial commit.
