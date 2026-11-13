/**
 * Chart-kind type shared by the task pane and the interop layer, so the UI never has to import
 * `Excel.ChartType`. The kinds and what each one maps to are defined in `config/chartSpecs`.
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
export type { ChartKind } from "../config/chartSpecs";
