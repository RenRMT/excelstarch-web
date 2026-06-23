import * as React from "react";
import { makeStyles, MessageBar, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import Header from "../taskpane/components/Header";
import ColourPanel from "../taskpane/components/ColourPanel";
import type { Status } from "../taskpane/components/status";
import { orgName } from "../config/brand";
import logoUrl from "../../assets/logo.png";

interface ColorPickerAppProps {
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

/**
 * Root of the Color Picker task pane (its own ribbon button + HTML entry). It hosts only the colour
 * tooling; with no shared React state with the Chart Builder pane, `ColourPanel` resolves the chart
 * the user has selected in Excel (see colourFlow). Mirrors the builder `App` shell (Header +
 * non-blocking MessageBar).
 */
const ColorPickerApp: React.FC<ColorPickerAppProps> = (props: ColorPickerAppProps) => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status | null>(null);

  return (
    <div className={styles.root}>
      <Header logo={logoUrl} title={props.title} message={`${orgName} colour picker`} />
      <main className={styles.body}>
        {status && (
          <MessageBar key={status.title} intent={status.intent} className={styles.status}>
            <MessageBarBody>
              <MessageBarTitle>{status.title}</MessageBarTitle>
              {status.body}
            </MessageBarBody>
          </MessageBar>
        )}
        <ColourPanel onStatus={setStatus} />
      </main>
    </div>
  );
};

export default ColorPickerApp;
