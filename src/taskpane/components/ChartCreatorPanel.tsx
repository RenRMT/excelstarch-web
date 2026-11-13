import * as React from "react";
import {
  Button,
  Title3,
  Body1,
  Caption1Strong,
  Spinner,
  makeStyles,
} from "@fluentui/react-components";
import type { Status } from "./status";
import type { ChartKind } from "../../logic/chartType";
import { createBrandedChart, styleSelectedChart, CreateChartResult } from "../../excel/chartFlow";
import type { RunResult } from "../../excel/session";
import type { ChartRef } from "../../excel/chromeText";
import { chartFamilies, chartSpecs } from "../../config/chartSpecs";

interface ChartCreatorPanelProps {
  onStatus: (status: Status) => void;
  onChartReady: (chart: ChartRef) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "8px" },
  family: { display: "flex", flexDirection: "column", rowGap: "4px" },
  buttons: { display: "flex", flexWrap: "wrap", columnGap: "6px", rowGap: "6px" },
  hint: { color: "#475569" }, // ROOS Grijs-600
});

/**
 * Creates a branded chart from the current selection, one button per chart type grouped by family
 * (the VBA ribbon groups; the first button of each family is its default). With a chart selected,
 * a type button retypes and restyles it; "Apply style" restyles it without changing its type.
 */
const ChartCreatorPanel: React.FC<ChartCreatorPanelProps> = ({ onStatus, onChartReady }) => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);

  const run = async (action: () => Promise<RunResult<CreateChartResult>>, verb: string) => {
    setBusy(true);
    try {
      const result = await action();
      if (result.ok) {
        const { chartName, worksheetId, warnings } = result.value;
        onChartReady({ worksheetId, chartName });
        if (warnings.length > 0) {
          onStatus({
            intent: "warning",
            title: `${verb} "${chartName}" with warnings`,
            body: warnings.join(" "),
          });
        } else {
          onStatus({ intent: "success", title: `${verb} "${chartName}".` });
        }
      } else {
        onStatus({
          intent: result.kind === "unsupported" ? "warning" : "error",
          title:
            result.kind === "unsupported" ? "Unsupported Excel version" : "Couldn't style the chart",
          body: result.message,
        });
      }
    } finally {
      setBusy(false);
    }
  };

  const handleCreate = (kind: ChartKind) => run(() => createBrandedChart(kind), "Styled");
  const handleApplyStyle = () => run(styleSelectedChart, "Styled");

  return (
    <section className={styles.panel}>
      <Title3>Create a branded chart</Title3>
      <Body1 className={styles.hint}>
        Select your data range, then choose a chart type. With a chart selected, a type button
        changes it to that type; Apply style keeps its type.
      </Body1>
      <div className={styles.buttons}>
        <Button appearance="primary" disabled={busy} onClick={() => void handleApplyStyle()}>
          Apply style to selected chart
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>
      {chartFamilies.map(({ family, kinds }) => (
        <div key={family} className={styles.family}>
          <Caption1Strong>{family}</Caption1Strong>
          <div className={styles.buttons}>
            {kinds.map((kind, i) => (
              <Button
                key={kind}
                size="small"
                appearance={i === 0 ? "primary" : "secondary"}
                disabled={busy}
                onClick={() => void handleCreate(kind)}
              >
                {chartSpecs[kind].label}
              </Button>
            ))}
          </div>
        </div>
      ))}
    </section>
  );
};

export default ChartCreatorPanel;
