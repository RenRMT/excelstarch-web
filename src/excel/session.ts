/* global Excel Office OfficeExtension */
/**
 * The single entry point for host interop: a typed `Excel.run` wrapper plus requirement-set
 * gating. Everything that touches the Excel object model goes through `runExcel`, which catches
 * `OfficeExtension.Error` and returns a typed `RunResult` so nothing throws across the UI boundary
 * — the task pane renders a non-blocking MessageBar from the result.
 *
 * INTEROP: only excel/ + chrome/ call into Office.js.
 */

/** Result of an interop action — never throws to the caller. */
export type RunResult<T> =
  | { ok: true; value: T }
  | { ok: false; kind: "unsupported" | "error"; message: string };

/** Whether the host supports the worksheet-shape/grouping APIs the chrome overlay needs. */
export function shapesSupported(): boolean {
  return Office.context.requirements.isSetSupported("ExcelApi", "1.9");
}

const UNSUPPORTED_MESSAGE =
  "This Excel version doesn't support the chart styling features (requires ExcelApi 1.9). " +
  "Try Excel on the web or a newer desktop build.";

/** A typed `RunResult` for the unsupported-host case (shared by all gated actions). */
export function unsupported<T>(): RunResult<T> {
  return { ok: false, kind: "unsupported", message: UNSUPPORTED_MESSAGE };
}

/**
 * Run an Excel.run callback, normalizing success/failure into a `RunResult`. Does NOT gate the
 * requirement set — callers that need shapes check `shapesSupported()` first and return
 * `unsupported()` so the UI can show the specific guidance.
 */
export async function runExcel<T>(
  fn: (context: Excel.RequestContext) => Promise<T>
): Promise<RunResult<T>> {
  try {
    const value = await Excel.run(fn);
    return { ok: true, value };
  } catch (error) {
    return { ok: false, kind: "error", message: describeError(error) };
  }
}

function describeError(error: unknown): string {
  if (error instanceof OfficeExtension.Error) {
    return `${error.code}: ${error.message}`;
  }
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}
