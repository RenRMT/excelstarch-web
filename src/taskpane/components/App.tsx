import * as React from "react";
import { makeStyles, MessageBar, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import Header from "./Header";
import ChartCreatorPanel from "./ChartCreatorPanel";
import ChromeTextPanel from "./ChromeTextPanel";
import ExportPanel from "./ExportPanel";
import type { Status } from "./status";
import { resolveChromeTarget, ChartRef, ChromeTarget } from "../../excel/chromeText";
import { watchChartSelection } from "../../excel/selectionEvents";
import logoUrl from "../../../assets/logo_empty_dark.png";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: { minHeight: "100vh" },
  body: {
    display: "flex",
    flexDirection: "column",
    rowGap: "16px",
    padding: "16px",
  },
  status: { marginBottom: "4px" },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status | null>(null);
  // The chart the chart-text panel edits: the chart selected in Excel, else the last one it had.
  const [target, setTarget] = React.useState<ChromeTarget | null>(null);
  const targetRef = React.useRef<ChartRef | null>(null);

  /** Re-resolve the text panel's chart; `preferred` is used when no chart is selected. */
  const refreshTarget = React.useCallback(async (preferred?: ChartRef) => {
    const result = await resolveChromeTarget(preferred ?? targetRef.current);
    if (result.ok) {
      targetRef.current = result.value;
      // Keep the same object when nothing changed, so the panel doesn't reset a field mid-typing.
      setTarget((prev) =>
        JSON.stringify(prev) === JSON.stringify(result.value) ? prev : result.value
      );
    }
  }, []);

  React.useEffect(() => {
    void refreshTarget();
    return watchChartSelection(() => void refreshTarget());
  }, [refreshTarget]);

  return (
    <div className={styles.root}>
      <Header logo={logoUrl} title={props.title} message="Excel chart builder" />
      <main className={styles.body}>
        {status && (
          <MessageBar key={status.title} intent={status.intent} className={styles.status}>
            <MessageBarBody>
              <MessageBarTitle>{status.title}</MessageBarTitle>
              {status.body}
            </MessageBarBody>
          </MessageBar>
        )}
        <ChartCreatorPanel onStatus={setStatus} onChartReady={(ref) => void refreshTarget(ref)} />
        <ChromeTextPanel target={target} onStatus={setStatus} />
        <ExportPanel onStatus={setStatus} />
      </main>
    </div>
  );
};

export default App;
