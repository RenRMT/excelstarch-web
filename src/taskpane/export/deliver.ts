/* global document, window, navigator, Blob, URL, atob, Uint8Array */
/**
 * Ways to hand a composed image to the user from the task pane. Office.js has no file access and
 * no Save As dialog, so these are browser mechanisms and each can be unavailable in a given host:
 * downloads work in Excel on the web but not reliably in the desktop task pane, and the clipboard
 * needs the async clipboard API with image support. Inserting the picture on the sheet (see
 * `excel/exportFlow.insertExportPicture`) is the route that always works.
 *
 * DOM code: runs in the task pane's browser.
 */
import type { ComposedImage } from "./compose";

function toBlob(image: ComposedImage): Blob {
  const binary = atob(image.base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: image.mimeType });
}

/** Offer the image as a browser download. Cannot report whether the host actually saved it. */
export function downloadImage(image: ComposedImage, fileName: string): void {
  const url = URL.createObjectURL(toBlob(image));
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  // Give the browser time to start the download before the URL is released.
  window.setTimeout(() => URL.revokeObjectURL(url), 10000);
}

type ClipboardItemCtor = new (items: Record<string, Blob>) => unknown;
type ImageClipboard = { write?: (items: unknown[]) => Promise<void> };

function clipboardItemCtor(): ClipboardItemCtor | undefined {
  return (window as unknown as { ClipboardItem?: ClipboardItemCtor }).ClipboardItem;
}

function imageClipboard(): ImageClipboard | undefined {
  return typeof navigator === "undefined"
    ? undefined
    : (navigator.clipboard as unknown as ImageClipboard | undefined);
}

/** Whether this browser exposes the async clipboard API with image support. */
export function clipboardImageSupported(): boolean {
  const clipboard = imageClipboard();
  return (
    clipboard !== undefined &&
    typeof clipboard.write === "function" &&
    clipboardItemCtor() !== undefined
  );
}

/**
 * Copy the image to the clipboard. Browsers only accept PNG here, so compose as PNG for this route.
 * Rejects when the host blocks clipboard access.
 */
export async function copyImageToClipboard(image: ComposedImage): Promise<void> {
  const Item = clipboardItemCtor();
  const clipboard = imageClipboard();
  if (Item === undefined || clipboard === undefined || typeof clipboard.write !== "function") {
    throw new Error("This Excel does not allow add-ins to copy images to the clipboard.");
  }
  await clipboard.write([new Item({ [image.mimeType]: toBlob(image) })]);
}
