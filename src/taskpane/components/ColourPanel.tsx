import * as React from "react";
import {
  Button,
  Dropdown,
  Option,
  Switch,
  Field,
  Title3,
  Body1,
  Spinner,
  makeStyles,
} from "@fluentui/react-components";
import type { Status } from "./status";
import type { RampName } from "../../config/ramps";
import { ramps, divergingTags, rampNames } from "../../config/ramps";
import {
  recolourSeries,
  applyElementFill,
  listSeries,
  hasActiveChart,
  type RecolourMode,
} from "../../excel/colourFlow";
import type { RunResult } from "../../excel/session";
import { loadLastUsedColours, saveLastUsedColours } from "../../persist/settings";

interface ColourPanelProps {
  onStatus: (status: Status) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  row: { display: "flex", columnGap: "8px", alignItems: "end" },
  hint: { color: "#605e5c" },
  // Fluent v9 Dropdown defaults to min-width 250px; shrink it to fit its text so the task pane
  // stays narrow. fit-content + min-width 0 lets the trigger size to the selected option label.
  narrowDropdown: { minWidth: "0", width: "fit-content" },
});

// H (Steel) is a neutral grey ramp, defined but intentionally not offered in the ramp menu (it's a
// neutral, not a sequential brand hue) — matching the VBA. Diverging tags already exclude it.
const SINGLE_RAMP_NAMES = (Object.keys(ramps) as RampName[]).filter((name) => name !== "H");

/** Friendly label for a diverging tag, e.g. "A|B" → "Ocean to Coral". */
function divergingLabel(tag: string): string {
  const [left, right] = tag.split("|") as RampName[];
  return `${rampNames[left]} to ${rampNames[right]}`;
}

/** Fill colours offered for per-element fill (matches colorFromName's known names) plus remove. */
const FILL_OPTIONS: { value: string; label: string }[] = [
  { value: "DATA1", label: "Data 1 (Ocean)" },
  { value: "DATA2", label: "Data 2 (Coral)" },
  { value: "DATA3", label: "Data 3 (Sky)" },
  { value: "DATA4", label: "Data 4 (Pine)" },
  { value: "DATA5", label: "Data 5 (Gold)" },
  { value: "DATA6", label: "Data 6 (Rust)" },
  { value: "DATA7", label: "Data 7 (Lavender)" },
  { value: "DATA8", label: "Data 8 (Grey)" },
  { value: "NEUTRAL2", label: "Neutral grey" },
  { value: "NEUTRAL4", label: "White" },
  { value: "NONE", label: "No fill (remove)" },
];

/**
 * Recolour the ACTIVE (selected) chart's series — palette (Contrasting/Rainbow), single-hue ramp,
 * diverging ramp, invert — and apply a per-element fill to one chosen series. Lives in the Color
 * Picker task pane, which has no shared React state with the builder, so it resolves the chart the
 * user has selected in Excel (via the interop's `getActiveChartOrNullObject`). The element list
 * loads on mount and via the Refresh button (re-read after selecting a different chart). Maps each
 * typed RunResult to a MessageBar; over-limit (>10 single / >21 diverging) surfaces as a non-blocking
 * error.
 */
const ColourPanel: React.FC<ColourPanelProps> = ({ onStatus }) => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);
  const [hasChart, setHasChart] = React.useState(false);
  const [useAltOrder, setUseAltOrder] = React.useState(false);
  const [rampName, setRampName] = React.useState<RampName>(SINGLE_RAMP_NAMES[0]);
  const [divergingTag, setDivergingTag] = React.useState<string>(divergingTags[0]);
  const [seriesNames, setSeriesNames] = React.useState<string[]>([]);
  const [elementIndex, setElementIndex] = React.useState<"all" | number>("all");
  const [fillValue, setFillValue] = React.useState<string>(FILL_OPTIONS[0].value);

  const disabled = !hasChart || busy;

  /** Re-resolve the active chart + its series (on mount and via Refresh). */
  const refresh = React.useCallback(async () => {
    const present = await hasActiveChart();
    const chartPresent = present.ok && present.value;
    setHasChart(chartPresent);
    if (!chartPresent) {
      setSeriesNames([]);
      return;
    }
    const result = await listSeries();
    if (result.ok) {
      setSeriesNames(result.value);
      setElementIndex("all");
    }
  }, []);

  React.useEffect(() => {
    void refresh();
  }, [refresh]);

  // Seed the controls from the workbook's last-used choices (persist/settings). Each value is
  // validated against the current option lists — a stale or hand-edited setting falls back to the
  // hard default rather than selecting something the dropdown can't show. elementIndex is not
  // persisted (it's a per-session selection tied to the live chart's series).
  React.useEffect(() => {
    const last = loadLastUsedColours();
    if (typeof last.paletteAltOrder === "boolean") {
      setUseAltOrder(last.paletteAltOrder);
    }
    if (last.rampName && SINGLE_RAMP_NAMES.some((n) => n === last.rampName)) {
      setRampName(last.rampName as RampName);
    }
    if (last.divergingTag && divergingTags.includes(last.divergingTag)) {
      setDivergingTag(last.divergingTag);
    }
    if (last.fillValue && FILL_OPTIONS.some((o) => o.value === last.fillValue)) {
      setFillValue(last.fillValue);
    }
  }, []);

  const report = (result: RunResult<void>, failTitle: string) => {
    if (result.ok) {
      onStatus({ intent: "success", title: "Colours applied." });
    } else {
      onStatus({
        intent: result.kind === "unsupported" ? "warning" : "error",
        title: result.kind === "unsupported" ? "Unsupported Excel version" : failTitle,
        body: result.message,
      });
    }
  };

  const runRecolour = async (mode: RecolourMode) => {
    setBusy(true);
    try {
      const result = await recolourSeries(mode);
      report(result, "Couldn't recolour the series");
      // Persist the choice only after a successful apply (matches the VBA last-used behaviour).
      // Invert has no choice to remember. Fire-and-forget — never block the UI on the save.
      if (result.ok) {
        if (mode.kind === "palette") {
          void saveLastUsedColours({ paletteAltOrder: mode.useAltOrder });
        } else if (mode.kind === "single") {
          void saveLastUsedColours({ rampName: mode.rampName });
        } else if (mode.kind === "diverging") {
          void saveLastUsedColours({ divergingTag: mode.tag });
        }
      }
    } finally {
      setBusy(false);
    }
  };

  const runElementFill = async () => {
    setBusy(true);
    try {
      const result = await applyElementFill(elementIndex, fillValue);
      report(result, "Couldn't apply the fill");
      if (result.ok) {
        void saveLastUsedColours({ fillValue });
      }
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Colours</Title3>
      <div className={styles.row}>
        <Body1 className={styles.hint}>
          {hasChart
            ? "Recolouring the selected chart."
            : "Select a chart on the sheet, then click Refresh."}
        </Body1>
        <Button size="small" disabled={busy} onClick={() => void refresh()}>
          Refresh
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>

      <Switch
        checked={useAltOrder}
        disabled={disabled}
        label={useAltOrder ? "Rainbow order" : "Contrasting order"}
        onChange={(_, d) => setUseAltOrder(d.checked)}
      />
      <div className={styles.row}>
        <Button
          appearance="primary"
          disabled={disabled}
          onClick={() => runRecolour({ kind: "palette", useAltOrder })}
        >
          Apply palette
        </Button>
        <Button disabled={disabled} onClick={() => runRecolour({ kind: "invert" })}>
          Invert
        </Button>
      </div>

      <div className={styles.row}>
        <Field label="Single-hue ramp">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={disabled}
            value={rampNames[rampName]}
            selectedOptions={[rampName]}
            onOptionSelect={(_, d) => d.optionValue && setRampName(d.optionValue as RampName)}
          >
            {SINGLE_RAMP_NAMES.map((name) => (
              <Option key={name} value={name}>
                {rampNames[name]}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button disabled={disabled} onClick={() => runRecolour({ kind: "single", rampName })}>
          Apply ramp
        </Button>
      </div>

      <div className={styles.row}>
        <Field label="Diverging ramp">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={disabled}
            value={divergingLabel(divergingTag)}
            selectedOptions={[divergingTag]}
            onOptionSelect={(_, d) => d.optionValue && setDivergingTag(d.optionValue)}
          >
            {divergingTags.map((tag) => (
              <Option key={tag} value={tag}>
                {divergingLabel(tag)}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button
          disabled={disabled}
          onClick={() => runRecolour({ kind: "diverging", tag: divergingTag })}
        >
          Apply diverging
        </Button>
      </div>

      <div className={styles.row}>
        <Field label="Element">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={disabled}
            value={elementIndex === "all" ? "All series" : seriesNames[elementIndex] ?? "Series"}
            selectedOptions={[String(elementIndex)]}
            onOptionSelect={(_, d) =>
              setElementIndex(d.optionValue === "all" ? "all" : Number(d.optionValue))
            }
          >
            <Option value="all">All series</Option>
            {seriesNames.map((name, i) => (
              <Option key={i} value={String(i)}>
                {name}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Field label="Fill">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={disabled}
            value={FILL_OPTIONS.find((o) => o.value === fillValue)?.label ?? fillValue}
            selectedOptions={[fillValue]}
            onOptionSelect={(_, d) => d.optionValue && setFillValue(d.optionValue)}
          >
            {FILL_OPTIONS.map((o) => (
              <Option key={o.value} value={o.value}>
                {o.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button disabled={disabled} onClick={runElementFill}>
          Apply fill
        </Button>
      </div>
    </section>
  );
};

export default ColourPanel;
