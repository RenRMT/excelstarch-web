/**
 * Default chrome placeholder text — ported from `reference/spec.json` (`placeholders.*`),
 * mirroring `modConfig.bas`. Shown in a freshly-built chrome box until the user edits it via
 * the chart-text panel.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
export const figurePlaceholder = "Figure XX (optional)";
export const titlePlaceholder = "Title in 28pt sentence case";
export const subtitlePlaceholder = "Subtitle in 22pt sentence case";
export const yAxisPlaceholder = "Y axis title (unit)";
export const xAxisPlaceholder = "X axis title (unit)";
export const sourcePlaceholder = "Source: Source text goes here.";
export const notesPlaceholder = "Notes: Notes text goes here.";
export const annotationPlaceholder = "Annotation";

/**
 * Sentinel a user types into a chrome-text field to clear it to *truly empty* (no placeholder).
 * A blank field restores the placeholder; this token writes an empty string instead.
 */
export const clearFieldToken = "-";
