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
import type { Status } from "./App";
import type { RampName } from "../../config/ramps";
import { ramps, divergingTags } from "../../config/ramps";
import {
  recolourSeries,
  applyElementFill,
  listSeries,
  type RecolourMode,
} from "../../excel/colourFlow";
import type { RunResult } from "../../excel/session";

interface ColourPanelProps {
  targetChartName: string | null;
  onStatus: (status: Status) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  row: { display: "flex", columnGap: "8px", alignItems: "end" },
  hint: { color: "#605e5c" },
});

const RAMP_NAMES = Object.keys(ramps) as RampName[];

/** Fill colours offered for per-element fill (matches colorFromName's known names) plus remove. */
const FILL_OPTIONS: { value: string; label: string }[] = [
  { value: "DATA1", label: "Data 1 (Ocean)" },
  { value: "DATA2", label: "Data 2 (Coral)" },
  { value: "DATA3", label: "Data 3 (Sky)" },
  { value: "DATA4", label: "Data 4 (Teal)" },
  { value: "DATA5", label: "Data 5 (Gold)" },
  { value: "DATA6", label: "Data 6 (Rust)" },
  { value: "DATA7", label: "Data 7 (Lavender)" },
  { value: "DATA8", label: "Data 8 (Grey)" },
  { value: "NEUTRAL2", label: "Neutral grey" },
  { value: "NEUTRAL4", label: "White" },
  { value: "NONE", label: "No fill (remove)" },
];

/**
 * Recolour the most-recently created/restyled chart's series — palette (Contrasting/Rainbow),
 * single-hue ramp, diverging ramp, invert — and apply a per-element fill to one chosen series.
 * Replaces the VBA ribbon colour buttons + Selection model. Maps each typed RunResult to a
 * MessageBar; over-limit (>10 single / >21 diverging) surfaces as a non-blocking error.
 */
const ColourPanel: React.FC<ColourPanelProps> = ({ targetChartName, onStatus }) => {
  const styles = useStyles();
  const disabled = targetChartName === null;
  const [busy, setBusy] = React.useState(false);
  const [useAltOrder, setUseAltOrder] = React.useState(false);
  const [rampName, setRampName] = React.useState<RampName>(RAMP_NAMES[0]);
  const [divergingTag, setDivergingTag] = React.useState<string>(divergingTags[0]);
  const [seriesNames, setSeriesNames] = React.useState<string[]>([]);
  const [elementIndex, setElementIndex] = React.useState<"all" | number>("all");
  const [fillValue, setFillValue] = React.useState<string>(FILL_OPTIONS[0].value);

  // Refresh the element selector whenever the target chart changes.
  React.useEffect(() => {
    if (targetChartName === null) {
      setSeriesNames([]);
      return;
    }
    let cancelled = false;
    void (async () => {
      const result = await listSeries(targetChartName);
      if (!cancelled && result.ok) {
        setSeriesNames(result.value);
        setElementIndex("all");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [targetChartName]);

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
    if (targetChartName === null) return;
    setBusy(true);
    try {
      report(await recolourSeries(targetChartName, mode), "Couldn't recolour the series");
    } finally {
      setBusy(false);
    }
  };

  const runElementFill = async () => {
    if (targetChartName === null) return;
    setBusy(true);
    try {
      report(
        await applyElementFill(targetChartName, elementIndex, fillValue),
        "Couldn't apply the fill"
      );
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Colours</Title3>
      <Body1 className={styles.hint}>
        {disabled
          ? "Create a chart first, then recolour its series here."
          : `Recolouring "${targetChartName}".`}
      </Body1>

      <Switch
        checked={useAltOrder}
        disabled={disabled || busy}
        label={useAltOrder ? "Rainbow order" : "Contrasting order"}
        onChange={(_, d) => setUseAltOrder(d.checked)}
      />
      <div className={styles.row}>
        <Button
          appearance="primary"
          disabled={disabled || busy}
          onClick={() => runRecolour({ kind: "palette", useAltOrder })}
        >
          Apply palette
        </Button>
        <Button disabled={disabled || busy} onClick={() => runRecolour({ kind: "invert" })}>
          Invert
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>

      <div className={styles.row}>
        <Field label="Single-hue ramp">
          <Dropdown
            disabled={disabled || busy}
            value={rampName}
            selectedOptions={[rampName]}
            onOptionSelect={(_, d) => d.optionValue && setRampName(d.optionValue as RampName)}
          >
            {RAMP_NAMES.map((name) => (
              <Option key={name} value={name}>
                {name}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button
          disabled={disabled || busy}
          onClick={() => runRecolour({ kind: "single", rampName })}
        >
          Apply ramp
        </Button>
      </div>

      <div className={styles.row}>
        <Field label="Diverging ramp">
          <Dropdown
            disabled={disabled || busy}
            value={divergingTag}
            selectedOptions={[divergingTag]}
            onOptionSelect={(_, d) => d.optionValue && setDivergingTag(d.optionValue)}
          >
            {divergingTags.map((tag) => (
              <Option key={tag} value={tag}>
                {tag}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button
          disabled={disabled || busy}
          onClick={() => runRecolour({ kind: "diverging", tag: divergingTag })}
        >
          Apply diverging
        </Button>
      </div>

      <div className={styles.row}>
        <Field label="Element">
          <Dropdown
            disabled={disabled || busy}
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
            disabled={disabled || busy}
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
        <Button disabled={disabled || busy} onClick={runElementFill}>
          Apply fill
        </Button>
      </div>
    </section>
  );
};

export default ColourPanel;
