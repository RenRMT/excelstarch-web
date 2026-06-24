/**
 * Chart-kind discriminated union shared by the task pane and the interop layer, so the UI never
 * has to import `Excel.ChartType`. The `excel/` layer maps `ChartKind` to the concrete Office.js
 * chart type. Phase 2 ships the classic types (bar, column, line, area, scatter, pie) plus the
 * chartex types (treemap, box & whisker) — completing the chart-type breadth.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
export type ChartKind =
  | "bar"
  | "column"
  | "line"
  | "area"
  | "scatter"
  | "pie"
  | "treemap"
  | "boxwhisker";
