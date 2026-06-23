/**
 * Chart-kind discriminated union shared by the task pane and the interop layer, so the UI never
 * has to import `Excel.ChartType`. The `excel/` layer maps `ChartKind` to the concrete Office.js
 * chart type. Phase 1 ships bar + column; later phases extend this union.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
export type ChartKind = "bar" | "column";
