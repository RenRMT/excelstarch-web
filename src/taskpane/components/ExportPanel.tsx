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
import { readExportLayers, insertExportPicture, ExportLayers } from "../../excel/exportFlow";
import { composeExport, ComposedImage } from "../export/compose";
import { downloadImage, copyImageToClipboard, clipboardImageSupported } from "../export/deliver";
import {
  EXPORT_SCALES,
  exportFileName,
  ExportFormat,
  ExportScale,
} from "../../logic/exportLayout";
import { titlePlaceholder } from "../../config/text";
import { loadExportPrefs, saveExportPrefs } from "../../persist/settings";

interface ExportPanelProps {
  onStatus: (status: Status) => void;
}

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", rowGap: "10px" },
  row: { display: "flex", flexWrap: "wrap", columnGap: "8px", rowGap: "8px", alignItems: "end" },
  hint: { color: "#475569" }, // ROOS Grijs-600
  narrowDropdown: { minWidth: "0", width: "fit-content" },
});

const FORMATS: { value: ExportFormat; label: string }[] = [
  { value: "png", label: "PNG" },
  { value: "jpeg", label: "JPEG" },
];

type Delivery = "download" | "copy" | "insert";

/**
 * Exports the selected chart together with its title, subtitle and source as one image. The chart
 * and its chrome are separate objects in Excel, so this is the way to take a finished chart out of
 * the workbook in one piece. Three delivery routes are offered because not every Excel host allows
 * every one: download, copy to clipboard, or insert as a picture next to the chart.
 */
const ExportPanel: React.FC<ExportPanelProps> = ({ onStatus }) => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);
  const [format, setFormat] = React.useState<ExportFormat>("png");
  const [scale, setScale] = React.useState<ExportScale>(2);
  const canCopy = React.useMemo(() => clipboardImageSupported(), []);

  // Seed from the workbook's last-used export choices, validated against the current options.
  React.useEffect(() => {
    const prefs = loadExportPrefs();
    const savedFormat = FORMATS.find((f) => f.value === prefs.format);
    if (savedFormat) {
      setFormat(savedFormat.value);
    }
    const savedScale = EXPORT_SCALES.find((s) => s === prefs.scale);
    if (savedScale !== undefined) {
      setScale(savedScale);
    }
  }, []);

  const deliver = async (
    delivery: Delivery,
    layers: ExportLayers,
    image: ComposedImage
  ): Promise<string> => {
    if (delivery === "download") {
      const fileName = exportFileName(
        layers.texts.title ?? null,
        titlePlaceholder,
        layers.chartName,
        format
      );
      downloadImage(image, fileName);
      return `Download started: ${fileName}. If nothing was saved, use "Insert on sheet" instead.`;
    }
    if (delivery === "copy") {
      await copyImageToClipboard(image);
      return "Image copied. Paste it into your document.";
    }
    const inserted = await insertExportPicture(layers.worksheetId, layers.frame, image.base64);
    if (!inserted.ok) {
      throw new Error(inserted.message);
    }
    return "Picture inserted next to the chart. Right-click it to copy or save it.";
  };

  const run = async (delivery: Delivery) => {
    setBusy(true);
    try {
      const layers = await readExportLayers(scale);
      if (!layers.ok) {
        onStatus({
          intent: layers.kind === "unsupported" ? "warning" : "error",
          title: layers.kind === "unsupported" ? "Unsupported Excel version" : "Couldn't export",
          body: layers.message,
        });
        return;
      }
      // The clipboard only takes PNG, whatever format is selected for files.
      const image = await composeExport(layers.value, scale, delivery === "copy" ? "png" : format);
      const message = await deliver(delivery, layers.value, image);
      onStatus({
        intent: "success",
        title: `Exported "${layers.value.chartName}" (${image.width} × ${image.height} px).`,
        body: message,
      });
      void saveExportPrefs({ format, scale });
    } catch (error) {
      onStatus({
        intent: "error",
        title: "Couldn't export",
        body: error instanceof Error ? error.message : String(error),
      });
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className={styles.panel}>
      <Title3>Export</Title3>
      <Body1 className={styles.hint}>
        Select a chart to export it with its title and source as one image. Excel add-ins can only
        produce PNG and JPEG.
      </Body1>
      <div className={styles.row}>
        <Field label="Format">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={busy}
            value={FORMATS.find((f) => f.value === format)?.label}
            selectedOptions={[format]}
            onOptionSelect={(_, d) => d.optionValue && setFormat(d.optionValue as ExportFormat)}
          >
            {FORMATS.map((f) => (
              <Option key={f.value} value={f.value}>
                {f.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Field label="Resolution">
          <Dropdown
            className={styles.narrowDropdown}
            disabled={busy}
            value={`${scale}×`}
            selectedOptions={[String(scale)]}
            onOptionSelect={(_, d) => {
              const next = EXPORT_SCALES.find((s) => String(s) === d.optionValue);
              if (next !== undefined) {
                setScale(next);
              }
            }}
          >
            {EXPORT_SCALES.map((s) => (
              <Option key={s} value={String(s)}>
                {`${s}×`}
              </Option>
            ))}
          </Dropdown>
        </Field>
      </div>
      <div className={styles.row}>
        <Button appearance="primary" disabled={busy} onClick={() => void run("download")}>
          Download
        </Button>
        {canCopy && (
          <Button disabled={busy} onClick={() => void run("copy")}>
            Copy
          </Button>
        )}
        <Button disabled={busy} onClick={() => void run("insert")}>
          Insert on sheet
        </Button>
        {busy && <Spinner size="tiny" label="Working…" />}
      </div>
    </section>
  );
};

export default ExportPanel;
