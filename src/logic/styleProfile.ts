/**
 * Everything the build flow needs to style one chart: formatting defaults, where its colour lives,
 * per-type finishing, and how cautious to be. A profile comes either from a builder kind (creating
 * or retyping a chart) or from an existing chart's Excel type (Apply Chart Style, which styles a
 * chart in place without changing its type — the VBA `ApplyChartStyle`).
 *
 * PURE: this module must never import `Excel`/`Office`.
 */
import type { ChartDefaults } from "../config/chartDefaults";
import {
  columnChartDefaults,
  lineChartDefaults,
  scatterChartDefaults,
  pieChartDefaults,
} from "../config/chartDefaults";
import { Axis } from "../config/enums";
import {
  chartSpecs,
  kindForExcelType,
  ChartKind,
  ColourTarget,
  Finishing,
} from "../config/chartSpecs";
import {
  colourTargetForChartType,
  styleClassForChartType,
  isScatterChartType,
  StyleClass,
} from "./chartClass";

export interface StyleProfile {
  defaults: ChartDefaults;
  colourTarget: ColourTarget;
  finishing: Finishing;
  /**
   * Leave the chart's axes and legend as they are. For chart types the add-in knows nothing about
   * — styling an axis that does not exist fails the whole batch.
   */
  keepAxesAndLegend: boolean;
  /**
   * Colour in a separate guarded step and skip silently on failure, as the VBA did for chart types
   * that may reject standard colouring.
   */
  bestEffortColour: boolean;
}

/** The profile for a chart the builder creates or retypes as `kind`. */
export function profileForKind(kind: ChartKind): StyleProfile {
  const spec = chartSpecs[kind];
  return {
    defaults: spec.defaults(),
    colourTarget: spec.colourTarget,
    finishing: spec.finishing,
    keepAxesAndLegend: false,
    bestEffortColour: false,
  };
}

/** Chrome only: no axes touched, no legend change, no value-axis title box. */
const otherDefaults: ChartDefaults = {
  gridlines: Axis.None,
  axisDisplay: Axis.None,
  legend: false,
  showYAxisTitle: false,
};

function defaultsForClass(styleClass: StyleClass, chartType: string): ChartDefaults {
  switch (styleClass) {
    case "pie":
      return pieChartDefaults();
    case "axisFill":
      return columnChartDefaults();
    case "lineScatter":
      return isScatterChartType(chartType) ? scatterChartDefaults() : lineChartDefaults();
    case "other":
      return { ...otherDefaults };
  }
}

/**
 * The profile for styling an existing chart of Excel type `chartType` in place. A type the builder
 * makes gets that kind's full profile; any other type gets the defaults of its VBA style class and
 * no type-specific finishing.
 */
export function profileForChartType(chartType: string): StyleProfile {
  const kind = kindForExcelType(chartType);
  if (kind !== undefined) {
    return profileForKind(kind);
  }

  const styleClass = styleClassForChartType(chartType);
  return {
    defaults: defaultsForClass(styleClass, chartType),
    colourTarget: colourTargetForChartType(chartType),
    finishing: styleClass === "pie" ? { squarePlot: true } : {},
    keepAxesAndLegend: styleClass === "other",
    bestEffortColour: styleClass === "other",
  };
}
