/**
 * The brand logo as raw base64, ready for Office.js `worksheet.shapes.addImage` (which takes
 * base64, not a URL). The logo is a committed PNG asset (`assets/logo.png`); webpack inlines it
 * to a `data:image/png;base64,...` URI at build time (the `?inline` query → `asset/inline` rule).
 * We strip the data-URI prefix here because `addImage` wants the bare base64 payload.
 *
 * This is the cleaner web-native replacement for the VBA `modEmbeddedImages.bas` hand-maintained
 * base64 string (which existed only because an `.xlsm` cannot reference an external file).
 *
 * INTEROP-adjacent: imports a bundled asset but does not call Excel; the shape builders use it.
 */
import logoDataUri from "../../assets/logo.png?inline";

/** Raw base64 PNG payload (no `data:` prefix) for `shape.addImage`. */
export const logoBase64: string = logoDataUri.replace(/^data:image\/[a-z]+;base64,/, "");
