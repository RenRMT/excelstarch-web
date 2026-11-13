import * as React from "react";
import {
  Button,
  Dropdown,
  Option,
  Field,
  Title3,
  Body1,
  Spinner,
  makeStyles,
} from "@fluentui/react-components";
import type { Status } from "./status";
import type { RunResult } from "../../excel/session";
import {
  toggleGridlines,
  toggleAxisLabels,
  toggleLegend,
  toggleAxisTitles,
  toggleDataLabels,
  listSeriesNames,
  LabelScope,
} from "../../excel/toggleFlow";
import { watchChartSelection } from "../../excel/selectionEvents";
import { addAnnotation, AnnotationAnchor } from "../../excel/annotationFlow";
import { listPointNames } from "../../excel/chartPoints";

interface ChartToolsPanelProps {
  onStatus: (status: Status) => void;
  /** Called after a toggle that changes the chrome (axis titles), so the text panel can refresh. */
  onChromeChanged: () => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  row: { display: "flex", flexWrap: "wrap", columnGap: "6px", rowGap: "6px", alignItems: "end" },
  hint: { color: "#475569" }, // ROOS Grijs-600
  narrowDropdown: { minWidth: "0", width: "fit-content" },
});

/**
 * The VBA "Toggles" ribbon group: each button steps the selected chart through one cycle and the
 * status line says where it landed. Data labels act on all series or the series picked here (Excel
 * add-ins cannot see which series is clicked in the chart); annotations go at the plot centre or
 * beside a point picked here.
 */
const ChartToolsPanel: React.FC<ChartToolsPanelProps> = ({ onStatus, onChromeChanged }) => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);
  const [seriesNames, setSeriesNames] = React.useState<string[]>([]);
  const [scope, setScope] = React.useState<LabelScope>("all");
  // Annotation anchor: the plot centre, or a point of a series (default: the series' last point).
  const [anchorSeries, setAnchorSeries] = React.useState<"centre" | number>("centre");
  const [anchorPoints, setAnchorPoints] = React.useState<string[]>([]);
  const [anchorPoint, setAnchorPoint] = React.useState(0);

  React.useEffect(() => {
    if (anchorSeries === "centre") {
      setAnchorPoints([]);
      return;
    }
    let cancelled = false;
    void listPointNames(anchorSeries).then((result) => {
      if (cancelled) return;
      const names = result.ok ? result.value : [];
      setAnchorPoints(names);
      setAnchorPoint(Math.max(0, names.length - 1)); // the VBA anchored a series on its last point
    });
    return () => {
      cancelled = true;
    };
  }, [anchorSeries, seriesNames]);

  const anchor: AnnotationAnchor =
    anchorSeries === "centre" || anchorPoints.length === 0
      ? { kind: "centre" }
      : { kind: "point", series: anchorSeries, point: anchorPoint };

  const refreshSeries = React.useCallback(async () => {
    const result = await listSeriesNames();
    const names = result.ok ? result.value : [];
    setSeriesNames(names);
    setScope((s) => (s === "all" || s < names.length ? s : "all"));
    setAnchorSeries((s) => (s === "centre" || s < names.length ? s : "centre"));
  }, []);

  React.useEffect(() => {
    void refreshSeries();
    return watchChartSelection(() => void refreshSeries());
  }, [refreshSeries]);

  const run = async (toggle: () => Promise<RunResult<string>>, changesChrome = false) => {
    setBusy(true);
    try {
      const result = await toggle();
      if (result.ok) {
        onStatus({ intent: "success", title: result.value });
        if (changesChrome) {
          onChromeChanged();
        }
      } else {
        onStatus({
          intent: result.kind === "unsupported" ? "warning" : "error",
          title: result.kind === "unsupported" ? "Unsupported Excel version" : "Couldn't change the chart",
          body: result.message,
        });
      }
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Chart tools</Title3>
      <Body1 className={styles.hint}>Select a chart. Each click moves to the next setting.</Body1>
      <div className={styles.row}>
        <Button size="small" disabled={busy} onClick={() => void run(toggleGridlines)}>
          Gridlines
        </Button>
        <Button size="small" disabled={busy} onClick={() => void run(toggleAxisLabels)}>
          Axis labels
        </Button>
        <Button size="small" disabled={busy} onClick={() => void run(toggleAxisTitles, true)}>
          Axis titles
        </Button>
        <Button size="small" disabled={busy} onClick={() => void run(toggleLegend)}>
          Legend
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>
      <div className={styles.row}>
        <Field label="Data labels on">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={busy}
            value={scope === "all" ? "All series" : seriesNames[scope] ?? "Series"}
            selectedOptions={[String(scope)]}
            onOptionSelect={(_, d) =>
              setScope(d.optionValue === "all" ? "all" : Number(d.optionValue))
            }
          >
            <Option value="all">All series</Option>
            {seriesNames.map((name, i) => (
              <Option key={i} value={String(i)}>
                {name || `Series ${i + 1}`}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Button size="small" disabled={busy} onClick={() => void run(() => toggleDataLabels(scope))}>
          Data labels
        </Button>
      </div>
      <div className={styles.row}>
        <Field label="Annotation at">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={busy}
            value={
              anchorSeries === "centre"
                ? "Plot centre"
                : seriesNames[anchorSeries] || `Series ${anchorSeries + 1}`
            }
            selectedOptions={[String(anchorSeries)]}
            onOptionSelect={(_, d) =>
              setAnchorSeries(d.optionValue === "centre" ? "centre" : Number(d.optionValue))
            }
          >
            <Option value="centre">Plot centre</Option>
            {seriesNames.map((name, i) => (
              <Option key={i} value={String(i)}>
                {name || `Series ${i + 1}`}
              </Option>
            ))}
          </Dropdown>
        </Field>
        {anchorSeries !== "centre" && anchorPoints.length > 0 && (
          <Field label="Point">
            <Dropdown
              className={styles.narrowDropdown}
              disabled={busy}
              value={anchorPoints[anchorPoint] ?? ""}
              selectedOptions={[String(anchorPoint)]}
              onOptionSelect={(_, d) => setAnchorPoint(Number(d.optionValue))}
            >
              {anchorPoints.map((name, i) => (
                <Option key={i} value={String(i)}>
                  {name}
                </Option>
              ))}
            </Dropdown>
          </Field>
        )}
        <Button size="small" disabled={busy} onClick={() => void run(() => addAnnotation(anchor))}>
          Add annotation
        </Button>
      </div>
    </section>
  );
};

export default ChartToolsPanel;
