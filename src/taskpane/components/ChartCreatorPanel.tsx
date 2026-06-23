import * as React from "react";
import { Button, Title3, Body1, Spinner, makeStyles } from "@fluentui/react-components";
import type { Status } from "./App";
import type { ChartKind } from "../../logic/chartType";
import { createBrandedChart } from "../../excel/chartFlow";

interface ChartCreatorPanelProps {
  onStatus: (status: Status) => void;
  onChartReady: (chartName: string) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "8px" },
  buttons: { display: "flex", columnGap: "8px" },
});

/**
 * Creates a branded chart from the current selection. With a chart selected, the create call
 * restyles it in place; otherwise it creates a new one. Maps the typed RunResult to a MessageBar.
 */
const ChartCreatorPanel: React.FC<ChartCreatorPanelProps> = ({ onStatus, onChartReady }) => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);

  const handleCreate = async (kind: ChartKind) => {
    setBusy(true);
    try {
      const result = await createBrandedChart(kind);
      if (result.ok) {
        onChartReady(result.value.chartName);
        if (result.value.warnings.length > 0) {
          onStatus({
            intent: "warning",
            title: `Created "${result.value.chartName}" with warnings`,
            body: result.value.warnings.join(" "),
          });
        } else {
          onStatus({ intent: "success", title: `Created "${result.value.chartName}".` });
        }
      } else {
        onStatus({
          intent: result.kind === "unsupported" ? "warning" : "error",
          title: result.kind === "unsupported" ? "Unsupported Excel version" : "Couldn't create the chart",
          body: result.message,
        });
      }
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Create a branded chart</Title3>
      <Body1>Select your data range, then choose a chart type. A selected chart is restyled in place.</Body1>
      <div className={styles.buttons}>
        <Button appearance="primary" disabled={busy} onClick={() => handleCreate("bar")}>
          Bar
        </Button>
        <Button appearance="primary" disabled={busy} onClick={() => handleCreate("column")}>
          Column
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>
    </section>
  );
};

export default ChartCreatorPanel;
