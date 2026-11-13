/* global document, window, Office, URLSearchParams */
/**
 * The dialog a ribbon command opens to report a problem. Ribbon commands run in a hidden page and
 * cannot show anything themselves, so they open this page with `displayDialogAsync`, passing the
 * message in the query string. OK asks the command page to close the dialog.
 */
const params = new URLSearchParams(window.location.search);
const titleEl = document.getElementById("title");
const bodyEl = document.getElementById("body");
if (titleEl) titleEl.textContent = params.get("title") ?? "ExcelStarch";
if (bodyEl) bodyEl.textContent = params.get("body") ?? "";

Office.onReady(() => {
  document.getElementById("close")?.addEventListener("click", () => {
    Office.context.ui.messageParent("close");
  });
});
