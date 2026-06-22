/**
 * Pure fill payload/tag parsing — ported from the unit-testable functions in `modColorFill.bas`
 * (`ParseFillPayload`, `IsRemoveFillPayload`, `ColorFromName`). The object-model apply loops
 * (`ApplyFill`, `RemoveFill`, target detection) are NOT here — they belong to the Phase-1
 * `excel/` layer, which the task pane drives (the VBA Selection model is gone).
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { Hex } from "../config/brand";
import {
  colorData1,
  colorData2,
  colorData3,
  colorData4,
  colorData5,
  colorData6,
  colorData7,
  colorData8,
  colorNeutral2,
  colorNeutral4,
} from "../config/brand";

/** True when a fill payload (the part after `FILL:`) means "remove the fill" (case-insensitive). */
export function isRemoveFillPayload(payload: string): boolean {
  switch (payload.trim().toUpperCase()) {
    case "NONE":
    case "NOFILL":
    case "OFF":
      return true;
    default:
      return false;
  }
}

/**
 * Split a resolved fill payload `"Name"` or `"Name|transparency"` into its colour name and a
 * transparency in [0, 1]. A missing, non-numeric, or out-of-range transparency falls back to 0
 * (opaque).
 *
 * Parsed with `parseFloat`, NOT `Number`/locale parsing: like VBA's `Val`, `parseFloat` always
 * treats `.` as the decimal point, so a comma-decimal locale (e.g. nl-NL) cannot misread
 * `"0.5"`. `parseFloat` returns `NaN` for non-numeric input, which falls back to 0.
 * (VBA `Val` also ignores embedded whitespace, e.g. `Val("1 2")=12`; that quirk is out of scope —
 * fill-tag payloads are always whitespace-free fixed strings like `"DATA1|0.5"`.)
 */
export function parseFillPayload(payload: string): { name: string; transparency: number } {
  const parts = payload.split("|");
  const name = parts[0];
  let transparency = 0;

  if (parts.length >= 2) {
    let t = parseFloat(parts[1]); // non-numeric → NaN; "." is always the decimal point
    if (Number.isNaN(t)) t = 0;
    if (t < 0) t = 0;
    if (t > 1) t = 1;
    transparency = t;
  }

  return { name, transparency };
}

/**
 * Resolve a colour name to its palette hex (case-insensitive). Returns `null` for an unknown
 * name — the VBA original used a `-1` sentinel; in TypeScript the absence is modelled as `null`.
 */
export function colorFromName(name: string): Hex | null {
  switch (name.toUpperCase()) {
    case "DATA1":
      return colorData1;
    case "DATA2":
      return colorData2;
    case "DATA3":
      return colorData3;
    case "DATA4":
      return colorData4;
    case "DATA5":
      return colorData5;
    case "DATA6":
      return colorData6;
    case "DATA7":
      return colorData7;
    case "DATA8":
      return colorData8;
    case "NEUTRAL2":
      return colorNeutral2;
    case "NEUTRAL4":
      return colorNeutral4;
    default:
      return null;
  }
}
