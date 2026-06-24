/**
 * Chart-kind discriminated union shared by the task pane and the interop layer, so the UI never
 * has to import `Excel.ChartType`. The `excel/` layer maps `ChartKind` to the concrete Office.js
 * chart type. Phase 2 ships the classic types (bar, column, line, area, scatter, pie); the chartex
 * types (treemap, box & whisker) extend this union in a later slice.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
export type ChartKind = "bar" | "column" | "line" | "area" | "scatter" | "pie";
