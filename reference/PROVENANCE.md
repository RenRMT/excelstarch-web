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
