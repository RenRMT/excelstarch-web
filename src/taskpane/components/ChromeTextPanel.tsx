import * as React from "react";
import { Field, Input, Textarea, Title3, Body1, makeStyles } from "@fluentui/react-components";
import type { Status } from "./status";
import { writeChromeText, ChromeField, ChromeTarget } from "../../excel/chromeText";
import { fieldValueFor } from "../../logic/chromeText";
import {
  titlePlaceholder,
  subtitlePlaceholder,
  yAxisPlaceholder,
  xAxisPlaceholder,
  sourcePlaceholder,
  notesPlaceholder,
} from "../../config/text";

interface ChromeTextPanelProps {
  /** The chart being edited and its current text, or null when there is none. */
  target: ChromeTarget | null;
  onStatus: (status: Status) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  hint: { color: "#475569" }, // ROOS Grijs-600
});

interface FieldDef {
  field: ChromeField;
  label: string;
  placeholder: string;
  multiline?: boolean;
}

const FIELDS: FieldDef[] = [
  { field: "title", label: "Title", placeholder: titlePlaceholder },
  { field: "subtitle", label: "Subtitle", placeholder: subtitlePlaceholder },
  { field: "yAxis", label: "Y-axis title", placeholder: yAxisPlaceholder },
  { field: "xAxis", label: "X-axis title", placeholder: xAxisPlaceholder },
  { field: "source", label: "Source", placeholder: sourcePlaceholder, multiline: true },
  { field: "notes", label: "Notes", placeholder: notesPlaceholder, multiline: true },
];

type FieldValues = Record<ChromeField, string>;

/** The field values to show for a target: each box's text, with the placeholder shown as blank. */
function valuesFor(target: ChromeTarget | null): FieldValues {
  const values = {} as FieldValues;
  for (const { field, placeholder } of FIELDS) {
    const text = target ? target.texts[field] : null;
    values[field] = text === null ? "" : fieldValueFor(text, placeholder);
  }
  return values;
}

/**
 * Edits the chrome text of the selected chart (or the one most recently created) by writing to the
 * named worksheet shapes. The fields show the chart's current text and apply on blur; blank
 * restores the placeholder, "-" clears the field completely. A field is disabled when the chart
 * has no such box (e.g. no y-axis title on a pie; axis titles are switched on and off with the
 * Axis titles tool).
 */
const ChromeTextPanel: React.FC<ChromeTextPanelProps> = ({ target, onStatus }) => {
  const styles = useStyles();
  const [values, setValues] = React.useState<FieldValues>(() => valuesFor(target));
  // What the chart holds, in field form — a blur only writes when the field differs from this.
  const applied = React.useRef<FieldValues>(valuesFor(target));

  React.useEffect(() => {
    const fresh = valuesFor(target);
    applied.current = fresh;
    setValues(fresh);
  }, [target]);

  const handleApply = async (field: ChromeField) => {
    const value = values[field];
    if (target === null || value === applied.current[field]) {
      return;
    }
    const result = await writeChromeText(target, field, value);
    if (result.ok) {
      applied.current = { ...applied.current, [field]: value };
    } else {
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
        {target === null
          ? "Create a chart, or select one made with ExcelStarch, to edit its title, subtitle and notes here."
          : `Editing "${target.chartName}". Leave a field blank to restore its placeholder, or type "-" to clear it completely.`}
      </Body1>
      {FIELDS.map(({ field, label, placeholder, multiline }) => {
        const disabled = target === null || target.texts[field] === null;
        const common = {
          disabled,
          value: values[field],
          placeholder,
          onBlur: () => void handleApply(field),
        };
        return (
          <Field key={field} label={label}>
            {multiline ? (
              <Textarea
                {...common}
                onChange={(_, d) => setValues((v) => ({ ...v, [field]: d.value }))}
              />
            ) : (
              <Input
                {...common}
                onChange={(_, d) => setValues((v) => ({ ...v, [field]: d.value }))}
              />
            )}
          </Field>
        );
      })}
    </section>
  );
};

export default ChromeTextPanel;
