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

/**
 * Split the shared source/notes box text into its two lines. Text frames report the in-shape line
 * break as \r, \n or \v depending on host, so split on any run of them (otherwise editing one field
 * could wipe the other). A missing second line yields an empty notes string.
 */
export function splitSourceNotes(text: string): { source: string; notes: string } {
  const [source = "", ...rest] = text.split(/[\r\n\v]+/);
  return { source, notes: rest.join(" ") };
}

/**
 * What to show in a task-pane field for a box's current text — the inverse of `resolveChromeText`:
 * the placeholder shows as an empty field (so the field's own hint appears), a cleared box shows
 * as the clear sentinel, and anything else verbatim. Applying the shown value unchanged therefore
 * writes back exactly what the box already holds.
 */
export function fieldValueFor(text: string, placeholder: string): string {
  if (text === placeholder) return "";
  if (text === "") return clearFieldToken;
  return text;
}
