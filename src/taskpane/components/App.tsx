import * as React from "react";
import { makeStyles, MessageBar, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import Header from "./Header";
import ChartCreatorPanel from "./ChartCreatorPanel";
import ChromeTextPanel from "./ChromeTextPanel";
import ColourPanel from "./ColourPanel";
import { orgName } from "../../config/brand";
import logoUrl from "../../../assets/logo.png";

interface AppProps {
  title: string;
}

/** A transient status message rendered in a non-blocking MessageBar (never a modal). */
export interface Status {
  intent: "success" | "warning" | "error";
  title: string;
  body?: string;
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
  // The chart most recently created/restyled — the chart-text panel writes to its shapes.
  const [targetChartName, setTargetChartName] = React.useState<string | null>(null);

  return (
    <div className={styles.root}>
      <Header logo={logoUrl} title={props.title} message={`${orgName} chart styles`} />
      <main className={styles.body}>
        {status && (
          <MessageBar key={status.title} intent={status.intent} className={styles.status}>
            <MessageBarBody>
              <MessageBarTitle>{status.title}</MessageBarTitle>
              {status.body}
            </MessageBarBody>
          </MessageBar>
        )}
        <ChartCreatorPanel
          onStatus={setStatus}
          onChartReady={setTargetChartName}
        />
        <ChromeTextPanel targetChartName={targetChartName} onStatus={setStatus} />
        <ColourPanel targetChartName={targetChartName} onStatus={setStatus} />
      </main>
    </div>
  );
};

export default App;
