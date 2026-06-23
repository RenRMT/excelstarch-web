/**
 * A transient status message rendered in a non-blocking Fluent `MessageBar` (never a modal).
 * Shared by both task-pane roots (Chart Builder `App` and Color Picker `ColorPickerApp`) and the
 * panels they compose, so neither root has to import the other.
 */
export interface Status {
  intent: "success" | "warning" | "error";
  title: string;
  body?: string;
}
