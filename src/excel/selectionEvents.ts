/* global Excel window */
/**
 * Tell a task pane when the chart selection may have changed, so it can re-resolve the active
 * chart without a manual Refresh. Office.js has no single "active chart changed" event, so this
 * listens to chart activation/deactivation on every worksheet, re-subscribes when a sheet is
 * added, and also fires when the pane regains focus (a cheap catch-all).
 *
 * Best-effort: if event registration fails on a host, the pane still works through the focus
 * fallback and its Refresh button. Notifications are debounced, since deactivating one chart and
 * activating another fires twice.
 *
 * INTEROP: Office.js events.
 */

const DEBOUNCE_MS = 150;

/**
 * Call `onChange` whenever the selected chart may have changed. Returns a function that stops the
 * focus fallback and silences further notifications (the host-side handlers live as long as the
 * pane does, which is the lifetime they are needed for).
 */
export function watchChartSelection(onChange: () => void): () => void {
  let stopped = false;
  let timer: number | undefined;

  const notify = () => {
    if (stopped) return;
    if (timer !== undefined) window.clearTimeout(timer);
    timer = window.setTimeout(onChange, DEBOUNCE_MS);
  };
  const handler = async () => notify();

  const subscribeSheet = (sheet: Excel.Worksheet) => {
    sheet.charts.onActivated.add(handler);
    sheet.charts.onDeactivated.add(handler);
  };

  void Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items");
    await ctx.sync();
    sheets.items.forEach(subscribeSheet);
    sheets.onAdded.add(async (event) => {
      await Excel.run(async (innerCtx) => {
        subscribeSheet(innerCtx.workbook.worksheets.getItem(event.worksheetId));
        await innerCtx.sync();
      });
    });
    sheets.onActivated.add(handler);
    await ctx.sync();
  }).catch(() => {
    // Events unavailable — the focus fallback below still works.
  });

  window.addEventListener("focus", notify);

  return () => {
    stopped = true;
    if (timer !== undefined) window.clearTimeout(timer);
    window.removeEventListener("focus", notify);
  };
}
