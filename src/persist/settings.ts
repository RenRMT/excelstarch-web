/* global Office */
/**
 * Last-used colour choices, persisted per-workbook via `Office.context.document.settings` — the
 * web equivalent of the VBA `ThisWorkbook.CustomDocumentProperties` last-used tags
 * (`modColorRamp.bas` / `modColorFill.bas`). Saved after a successful apply, restored when the Color
 * Picker pane mounts so its dropdowns pre-select the user's previous choices.
 *
 * Persistence is **best-effort and silent** (mirrors the VBA `On Error Resume Next`): a save/load
 * failure must never disrupt the colour action or the pane, so these helpers never throw. Loaded
 * values are untrusted (a workbook can carry stale or hand-edited settings) — the caller validates
 * each field against the current option lists before applying it.
 *
 * INTEROP: reads/writes `Office.context.document.settings`. Needs no ExcelApi 1.9 gate — settings is
 * available once `Office.onReady` has fired (both panes await it before rendering). Sibling of
 * `excel/`; must not import `excel/`, and the pure `logic/`+`config/` layers must not import this.
 */

/** One bag of last-used choices. All optional — an absent field means "never set, use the UI default". */
export interface LastUsedColours {
  /** Palette order: Rainbow (true) vs Contrasting (false). */
  paletteAltOrder?: boolean;
  /** Single-hue ramp name (a `RampName` key, stored as string). */
  rampName?: string;
  /** Diverging tag, `"LEFT|RIGHT"`. */
  divergingTag?: string;
  /** Per-element fill value (a `FILL_OPTIONS` value, e.g. `"DATA1"` / `"NONE"`). */
  fillValue?: string;
}

/** Single namespaced key so the add-in's settings don't collide with anything else in the workbook. */
const KEY = "ExcelStarch:lastUsedColours";

/**
 * In-memory copy of the last-known-good bag. Seeded lazily from `settings` on first read and updated
 * synchronously on every save, so back-to-back fire-and-forget saves merge over each other's patches
 * (rather than each re-reading the same pre-save snapshot and clobbering the other's field). `null`
 * until first loaded.
 */
let cache: LastUsedColours | null = null;

/** Keep only well-typed fields, so a stale or hand-edited setting can't carry garbage into state/storage. */
function sanitize(raw: unknown): LastUsedColours {
  if (!raw || typeof raw !== "object") {
    return {};
  }
  const r = raw as Record<string, unknown>;
  const clean: LastUsedColours = {};
  if (typeof r.paletteAltOrder === "boolean") clean.paletteAltOrder = r.paletteAltOrder;
  if (typeof r.rampName === "string") clean.rampName = r.rampName;
  if (typeof r.divergingTag === "string") clean.divergingTag = r.divergingTag;
  if (typeof r.fillValue === "string") clean.fillValue = r.fillValue;
  return clean;
}

/**
 * Read the stored last-used choices. Synchronous (`settings.get` is sync) and never throws — returns
 * an empty bag if nothing is stored or the host/settings is unavailable. Fields are sanitized by
 * type; the caller still validates each against its current option list before applying it.
 */
export function loadLastUsedColours(): LastUsedColours {
  if (cache !== null) {
    return cache;
  }
  try {
    cache = sanitize(Office.context.document.settings.get(KEY));
  } catch {
    cache = {};
  }
  return cache;
}

/**
 * Merge `patch` over the stored choices and persist. Best-effort: resolves regardless of the
 * `saveAsync` status and swallows any failure (the colour action already succeeded; persistence is a
 * nicety). The in-memory `cache` is updated synchronously before the async write, so rapid
 * fire-and-forget saves accumulate their patches instead of racing (last-write-wins only per field).
 */
export async function saveLastUsedColours(patch: Partial<LastUsedColours>): Promise<void> {
  try {
    const settings = Office.context.document.settings;
    cache = { ...loadLastUsedColours(), ...sanitize(patch) };
    settings.set(KEY, cache);
    await new Promise<void>((resolve) => {
      // saveAsync is callback-style, not promise-returning; resolve no matter the status.
      settings.saveAsync(() => resolve());
    });
  } catch {
    // Silent — never disrupt the UI on a persistence failure.
  }
}
