import * as React from "react";
import { Field, Input, Textarea, Title3, Body1, makeStyles } from "@fluentui/react-components";
import type { Status } from "./App";
import { writeChromeText, ChromeField } from "../../excel/chromeText";

interface ChromeTextPanelProps {
  targetChartName: string | null;
  onStatus: (status: Status) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  hint: { color: "#605e5c" },
});

interface FieldDef {
  field: ChromeField;
  label: string;
  multiline?: boolean;
  /** Disabled for bar/column (the y-axis title box isn't built for them). */
  unavailable?: boolean;
}

const FIELDS: FieldDef[] = [
  { field: "figure", label: "Figure number" },
  { field: "title", label: "Title" },
  { field: "subtitle", label: "Subtitle" },
  { field: "source", label: "Source", multiline: true },
  { field: "notes", label: "Notes", multiline: true },
  { field: "yAxis", label: "Y-axis title", unavailable: true },
];

/**
 * Edits the chrome text of the most recently created/restyled chart by writing to the named
 * worksheet shapes (no chart selection required). Applies on blur; blank restores the placeholder,
 * "-" clears the field completely.
 */
const ChromeTextPanel: React.FC<ChromeTextPanelProps> = ({ targetChartName, onStatus }) => {
  const styles = useStyles();
  const disabled = targetChartName === null;

  const handleApply = async (field: ChromeField, value: string) => {
    if (targetChartName === null) {
      return;
    }
    const result = await writeChromeText(targetChartName, field, value);
    if (!result.ok) {
      onStatus({
        intent: result.kind === "unsupported" ? "warning" : "error",
        title: "Couldn't update the chart text",
        body: result.message,
      });
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Chart text</Title3>
      <Body1 className={styles.hint}>
        {disabled
          ? "Create a chart first, then edit its title, subtitle, and notes here."
          : `Editing "${targetChartName}". Leave a field blank to restore its placeholder, or type "-" to clear it completely.`}
      </Body1>
      {FIELDS.map(({ field, label, multiline, unavailable }) => (
        <Field
          key={field}
          label={unavailable ? `${label} (not used for bar/column)` : label}
        >
          {multiline ? (
            <Textarea
              disabled={disabled || unavailable}
              onBlur={(e) => handleApply(field, e.target.value)}
            />
          ) : (
            <Input
              disabled={disabled || unavailable}
              onBlur={(e) => handleApply(field, e.target.value)}
            />
          )}
        </Field>
      ))}
    </section>
  );
};

export default ChromeTextPanel;
