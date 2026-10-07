# Reference snapshot provenance

The files under `reference/vba-source/` are a **pinned, read-only snapshot** of the
original ExcelStarch VBA add-in. They are the source of truth this project ports from.
**Do not edit them** — they exist only as the auditable input to the TypeScript port.
If the upstream brand palette or logic changes, re-snapshot from a new pinned commit and
update this file rather than hand-editing the copies.

## Source

| Field | Value |
|-------|-------|
| Source repo | https://github.com/RenRMT/ExcelStarch.git |
| Local path  | `c:/projects/ExcelStarch` |
| Branch      | `demo/inso-brand-colors` |
| Commit SHA  | `c6fd5df9b04ffac334ee87631228d80493d1601c` |
| Snapshot of | `modules/*.bas` (subset needed for the port) |

## Why this branch (not `main`)

`modConfig.bas` (brand palette) and `modEmbeddedImages.bas` (logo) differ materially
between `main` and `demo/inso-brand-colors`. The INSO palette + INSO logo on
`demo/inso-brand-colors` is the canonical source for this port (confirmed decision).
Porting from `main` would bake in the wrong palette and logo.

## Files in this snapshot

| File | Role in the port |
|------|------------------|
| `modConfig.bas` | Brand colours, fonts, sizes, placeholder text → `src/config/brand.ts` |
| `modConfigDerived.bas` | Derived geometry + `ChartDefaults` → `src/config/geometry.ts`, `chartDefaults.ts` |
| `modColorRamp.bas` | Ramp math, `ParseDivergingTag`, `LoadPalette` → `src/logic/colorRamp.ts` |
| `modColorContrast.bas` | WCAG luminance → `src/logic/colorContrast.ts` |
| `modColorFill.bas` | `ParseFillPayload` / `ColorFromName` → `src/logic/colorFill.ts` |
| `modColorSeries.bas` | `GetPaletteColor` / palette order → `src/logic/colorSeries.ts` |
| `modEngineBuilder.bas` | Non-shape format reference; identifies which chrome is deleted |
| `modEngineExChrome.bas` | Worksheet-chrome pipeline to generalize (Phase 1 chrome engine) |
| `modEmbeddedImages.bas` | Base64 logo source → `assets/logo.svg` |
| `modTestHarness.bas` | Assertions to reproduce in the Jest suite |

## Intentional overrides of the upstream snapshot

`spec.json` is normally a faithful transcription of the pinned VBA. One field is a deliberate
**product override**, not a transcription of upstream:

| Field | Upstream value | spec.json value | Why |
|-------|----------------|-----------------|-----|
| `orgName` | `"COMPANY"` | `"INSO"` | Upstream `modConfig.bas` on `demo/inso-brand-colors` still carries the `"COMPANY"` placeholder; the product name is INSO. Set here (and in `src/config/brand.ts`) rather than waiting on an upstream re-snapshot. The `spec-parity` test still passes because both sides agree. |

If upstream later sets `orgName` itself, re-snapshot per below and drop this row.

## Re-snapshot procedure

```sh
# from a clean checkout of the source repo at the desired commit
SRC=c:/projects/ExcelStarch
DST=c:/projects/excelstarch-web
for f in modConfig modConfigDerived modColorRamp modColorContrast modColorFill \
         modColorSeries modEngineBuilder modEngineExChrome modEmbeddedImages modTestHarness; do
  cp "$SRC/modules/$f.bas" "$DST/reference/vba-source/$f.bas"
done
# then update the Commit SHA / Branch fields above
```

## ROOS palette (`roos-palette.json`)

The colours and the typeface no longer come from the VBA snapshot. `roos-palette.json` holds the
values copied verbatim from the ROOS (RVO Open Ontwerp Systeem) colour picker, and is the source
of every colour:

- **Data colours** (`colorData1..8`) are its `seriesPalette`, in order.
- **Ramps** are its nine `ramps` (six tints each, 150 → base). They replace the eight
  ten-step VBA ramps.
- **Font** for all chart text is its `font` (Verdana), replacing Calibri.
- **Chart title** colour is Lintblauw.
- **Every other colour** is the nearest ROOS token (CIEDE2000) to the colour the VBA port used,
  with neutrals kept on the ROOS grey scale.

`test/config/roos-parity.test.ts` reads this file and checks the ported values against it.
The `colors`, `ramps`, `rampNames`, `divergingTags`, `fonts.primary`, logo and figure-box entries
in `spec.json` are kept for history only and are no longer used.
