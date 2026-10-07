import * as React from "react";
import { makeStyles, MessageBar, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import Header from "./Header";
import ChartCreatorPanel from "./ChartCreatorPanel";
import ChromeTextPanel from "./ChromeTextPanel";
import type { Status } from "./status";
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
  // The chart most recently created/restyled — the chart-text panel writes to its shapes.
  const [targetChartName, setTargetChartName] = React.useState<string | null>(null);

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
        <ChartCreatorPanel
          onStatus={setStatus}
          onChartReady={setTargetChartName}
        />
        <ChromeTextPanel targetChartName={targetChartName} onStatus={setStatus} />
      </main>
    </div>
  );
};

export default App;
