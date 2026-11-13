/**
 * Annotation naming and placement (modEngineStyle.bas AnnotateButton).
 */
import {
  annotationPrefix,
  isAnnotationName,
  nextAnnotationName,
  annotationBoxAtCentre,
  annotationBoxAtPoint,
} from "../../src/logic/annotation";
import {
  annotationBoxWidth,
  annotationBoxHeight,
  annotationOffsetX,
  annotationOffsetY,
} from "../../src/config/geometry";

describe("annotation names", () => {
  it("numbers from 1 and stacks after the highest existing number", () => {
    expect(nextAnnotationName("Chart 1", [])).toBe("Chart 1_Annotation1");
    expect(
      nextAnnotationName("Chart 1", ["Chart 1_Annotation1", "Chart 1_Annotation3", "Chart 1_TitleBox"])
    ).toBe("Chart 1_Annotation4");
  });

  it("ignores other charts, including ones whose name starts the same", () => {
    expect(nextAnnotationName("Chart 1", ["Chart 10_Annotation5", "Chart 2_Annotation2"])).toBe(
      "Chart 1_Annotation1"
    );
  });

  it("recognises annotation names", () => {
    expect(annotationPrefix("Chart 1")).toBe("Chart 1_Annotation");
    expect(isAnnotationName("Chart 1_Annotation2")).toBe(true);
    expect(isAnnotationName("Chart 1_TitleBox")).toBe(false);
    expect(isAnnotationName("Chart 1_AnnotationX")).toBe(false);
  });
});

describe("annotation placement", () => {
  it("centres the box on an area", () => {
    const box = annotationBoxAtCentre({ left: 100, top: 200, width: 400, height: 300 });
    expect(box.left + box.width / 2).toBeCloseTo(300, 10);
    expect(box.top + box.height / 2).toBeCloseTo(350, 10);
    expect(box.width).toBe(annotationBoxWidth);
    expect(box.height).toBe(annotationBoxHeight);
  });

  it("offsets the box from a data label", () => {
    const box = annotationBoxAtPoint(250, 120);
    expect(box.left).toBe(250 + annotationOffsetX);
    expect(box.top).toBe(120 + annotationOffsetY);
  });
});
