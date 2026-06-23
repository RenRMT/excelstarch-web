/**
 * Pure resolution of what text a chrome shape should hold, given the user's field input.
 *
 * The rule (replaces the lost VBA Selection-driven edit model):
 *   - blank field            → restore the placeholder (the helpful default)
 *   - the clear sentinel `-`  → clear to truly empty ("")
 *   - any other text          → that text verbatim
 *
 * Kept pure so the three-way rule is unit-proven without a host; the `excel/` layer applies the
 * result to `textFrame.textRange.text`.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import { clearFieldToken } from "../config/text";

/** Resolve the literal text to write to a chrome shape for a given field value + placeholder. */
export function resolveChromeText(value: string, placeholder: string): string {
  const trimmed = value.trim();
  if (trimmed === clearFieldToken) {
    return "";
  }
  return trimmed.length > 0 ? value : placeholder;
}
