/**
 * Unit tests for the pure export maths: layer placement, file naming and word wrap.
 */
import {
  exportLayout,
  exportFileName,
  wrapText,
  PX_PER_POINT,
} from "../../src/logic/exportLayout";

describe("exportLayout", () => {
  const canvas = { left: 100, top: 50, width: 600, height: 600 };
  const chart = { left: 103, top: 146, width: 594, height: 430 };

  it("sizes the image from the frame at 96 px per 72 pt, times the scale", () => {
    expect(exportLayout(canvas, chart, canvas, 1)).toMatchObject({ width: 800, height: 800 });
    expect(exportLayout(canvas, chart, canvas, 3)).toMatchObject({ width: 2400, height: 2400 });
  });

  it("places the chart relative to the frame origin", () => {
    const layout = exportLayout(canvas, chart, canvas, 2);
    const k = PX_PER_POINT * 2;
    expect(layout.pxPerPoint).toBeCloseTo(k, 10);
    expect(layout.chart.left).toBeCloseTo(3 * k, 10);
    expect(layout.chart.top).toBeCloseTo(96 * k, 10);
    expect(layout.chart.width).toBeCloseTo(594 * k, 10);
    expect(layout.chart.height).toBeCloseTo(430 * k, 10);
  });

  it("places a chrome group that extends beyond the canvas at a negative offset", () => {
    // e.g. an annotation box dragged left of the canvas widens the group.
    const group = { left: 80, top: 50, width: 620, height: 600 };
    const layout = exportLayout(canvas, chart, group, 1);
    expect(layout.chrome?.left).toBeCloseTo(-20 * PX_PER_POINT, 10);
    expect(layout.chrome?.width).toBeCloseTo(620 * PX_PER_POINT, 10);
    expect(layout.width).toBe(800); // the frame, not the group, sets the image size
  });

  it("has no chrome layer for an unstyled chart framed on itself", () => {
    const layout = exportLayout(chart, chart, null, 1);
    expect(layout.chrome).toBeNull();
    expect(layout.chart.left).toBeCloseTo(0, 10);
    expect(layout.chart.top).toBeCloseTo(0, 10);
    expect(layout.width).toBe(Math.round(594 * PX_PER_POINT));
  });
});

describe("exportFileName", () => {
  const placeholder = "Title in 28pt sentence case";

  it("uses a real title", () => {
    expect(exportFileName("Sales by region", placeholder, "Chart 1", "png")).toBe(
      "Sales by region.png"
    );
  });

  it("falls back to the chart name for a placeholder, empty or missing title", () => {
    expect(exportFileName(placeholder, placeholder, "Chart 1", "png")).toBe("Chart 1.png");
    expect(exportFileName("  ", placeholder, "Chart 1", "png")).toBe("Chart 1.png");
    expect(exportFileName(null, placeholder, "Chart 1", "png")).toBe("Chart 1.png");
  });

  it("strips characters that are illegal in file names and collapses whitespace", () => {
    expect(exportFileName('Q1/Q2: "growth"?\nby <region>', placeholder, "Chart 1", "png")).toBe(
      "Q1 Q2 growth by region.png"
    );
  });

  it("drops trailing dots, caps the length and never returns an empty base", () => {
    expect(exportFileName("Totals...", placeholder, "Chart 1", "png")).toBe("Totals.png");
    expect(exportFileName("x".repeat(200), placeholder, "Chart 1", "png")).toBe(
      `${"x".repeat(80)}.png`
    );
    expect(exportFileName("???", placeholder, "***", "png")).toBe("chart.png");
  });

  it("uses .jpg for JPEG", () => {
    expect(exportFileName("Sales", placeholder, "Chart 1", "jpeg")).toBe("Sales.jpg");
  });
});

describe("wrapText", () => {
  const measure = (s: string) => s.length; // one unit per character

  it("keeps short text on one line", () => {
    expect(wrapText("short text", 20, measure)).toEqual(["short text"]);
  });

  it("wraps at word boundaries", () => {
    expect(wrapText("one two three four", 9, measure)).toEqual(["one two", "three", "four"]);
  });

  it("keeps explicit line breaks, including empty lines", () => {
    expect(wrapText("Source: a\nNotes: b", 40, measure)).toEqual(["Source: a", "Notes: b"]);
    expect(wrapText("a\n\nb", 40, measure)).toEqual(["a", "", "b"]);
  });

  it("gives an over-long word its own line instead of splitting it", () => {
    expect(wrapText("a verylongword b", 5, measure)).toEqual(["a", "verylongword", "b"]);
  });
});
