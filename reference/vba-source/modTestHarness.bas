Attribute VB_Name = "modTestHarness"
'==== Module: modTestHarness ====
' Lightweight unit-test harness for the project's PURE logic - the deterministic
' functions that take values in and return values out with no dependency on the
' Excel object model (charts, Selection, Shapes). Object-model-bound procedures
' (chart builders, toggles, export) are NOT covered here; see docs/testing_manual.md
' for their manual test steps.
'
' How to run
' ----------
'   1. Open the VBE Immediate Window (Ctrl+G).
'   2. Type:  RunAllTests   and press Enter.
'   3. Each test prints a "PASS:" line. A failing Debug.Assert halts execution on
'      the offending line in the IDE (assertions are no-ops in a compiled add-in,
'      so this module is harmless to ship).
'
' Scope (high-ROI targets only - bugs here are both likely and invisible to eye):
'   modColorContrast  RelativeLuminance, ContrastColorForFill   (WCAG maths)
'   modColorRamp           OrderedRampSteps, DivergingSideCount/HasMiddle, ParseDivergingTag
'   modColorFill     ParseFillPayload, IsRemoveFillPayload, ColorFromName
'   modColorSeries   GetPaletteColor                          (palette-order map)
'   modEngineBuilder   PlotAreaTopFor, PlotAreaHeightFor        (plot-area geometry)
'
' Trivial lookups (LoadPalette, chart-type classifiers, ChartDefaults factories)
' are deliberately NOT tested - a test there only re-states the constants.
Option Explicit


Public Sub RunAllTests()
    Debug.Print "=== Test run: " & Now & " ==="

    TestRelativeLuminance
    TestContrastColorForFill
    TestOrderedRampSteps
    TestDivergingLayout
    TestParseDivergingTag
    TestParseFillPayload
    TestIsRemoveFillPayload
    TestColorFromName
    TestGetPaletteColor
    TestPlotAreaGeometry

    Debug.Print "=== All tests passed ==="
End Sub


' ------------------------------------------------------------
'   modColorContrast
' ------------------------------------------------------------

Private Sub TestRelativeLuminance()
    ' Black and white are exact endpoints of the WCAG luminance scale.
    Debug.Assert RelativeLuminance(RGB(0, 0, 0)) = 0#
    Debug.Assert RelativeLuminance(RGB(255, 255, 255)) = 1#

    ' A mid colour must land strictly between the endpoints.
    Dim mid As Double
    mid = RelativeLuminance(RGB(0, 119, 187))   ' arbitrary mid-tone blue
    Debug.Assert mid > 0# And mid < 1#

    ' Green is weighted far more heavily than blue (0.7152 vs 0.0722).
    Debug.Assert RelativeLuminance(RGB(0, 255, 0)) > RelativeLuminance(RGB(0, 0, 255))

    Debug.Print "  PASS: TestRelativeLuminance"
End Sub

Private Sub TestContrastColorForFill()
    ' Dark fills -> white text; light fills -> dark brand text.
    Debug.Assert ContrastColorForFill(RGB(0, 0, 0)) = colorWhite
    Debug.Assert ContrastColorForFill(RGB(255, 255, 255)) = colorBrand3

    Debug.Print "  PASS: TestContrastColorForFill"
End Sub


' ------------------------------------------------------------
'   modColorRamp - step ordering
' ------------------------------------------------------------

Private Sub TestOrderedRampSteps()
    ' Priority is [6,2,4,3,5,7,8,1,9,10]; OrderedRampSteps takes the first n,
    ' sorts ascending, then reverses so the darkest step comes first.

    ' n=1: the single darkest priority step.
    AssertArrayEqual OrderedRampSteps(1), Array(6), "OrderedRampSteps(1)"

    ' n=2: priority [6,2] -> sort asc [2,6] -> reverse (darkest first) [6,2].
    AssertArrayEqual OrderedRampSteps(2), Array(6, 2), "OrderedRampSteps(2)"

    ' n=3: priority [6,2,4] -> sort asc [2,4,6] -> reverse (darkest first) [6,4,2].
    AssertArrayEqual OrderedRampSteps(3), Array(6, 4, 2), "OrderedRampSteps(3)"

    ' n=5: priority [6,2,4,3,5] -> sort asc [2,3,4,5,6] -> reverse [6,5,4,3,2].
    AssertArrayEqual OrderedRampSteps(5), Array(6, 5, 4, 3, 2), "OrderedRampSteps(5)"

    ' n=10: full set, darkest (10) to lightest (1).
    AssertArrayEqual OrderedRampSteps(10), Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 1), "OrderedRampSteps(10)"

    Debug.Print "  PASS: TestOrderedRampSteps"
End Sub

Private Sub TestDivergingLayout()
    ' Even count: equal sides, no grey middle.
    Debug.Assert DivergingSideCount(8) = 4
    Debug.Assert DivergingHasMiddle(8) = False

    ' Odd count: floor(n/2) per side plus a grey middle.
    Debug.Assert DivergingSideCount(5) = 2
    Debug.Assert DivergingHasMiddle(5) = True

    ' Single series: no side, no middle pairing (n\2 = 0).
    Debug.Assert DivergingSideCount(1) = 0
    Debug.Assert DivergingHasMiddle(1) = True

    Debug.Print "  PASS: TestDivergingLayout"
End Sub

Private Sub TestParseDivergingTag()
    Dim l As String, r As String

    ' Valid "A|B" splits into the two ramp names.
    Debug.Assert ParseDivergingTag("A|B", l, r) = True
    Debug.Assert l = "A" And r = "B"

    ' Missing separator or empty side is invalid.
    Debug.Assert ParseDivergingTag("A", l, r) = False
    Debug.Assert ParseDivergingTag("A|", l, r) = False
    Debug.Assert ParseDivergingTag("|B", l, r) = False

    Debug.Print "  PASS: TestParseDivergingTag"
End Sub


' ------------------------------------------------------------
'   modColorFill - payload/tag parsing
' ------------------------------------------------------------

Private Sub TestParseFillPayload()
    Dim nm As String, tr As Double

    ' Name only -> opaque (0 transparency).
    ParseFillPayload "DATA1", nm, tr
    Debug.Assert nm = "DATA1" And tr = 0#

    ' Name + transparency.
    ParseFillPayload "DATA1|0.5", nm, tr
    Debug.Assert nm = "DATA1" And tr = 0.5

    ' Out-of-range transparency clamps to [0, 1].
    ParseFillPayload "DATA2|1.7", nm, tr
    Debug.Assert nm = "DATA2" And tr = 1#

    ' Non-numeric transparency falls back to 0.
    ParseFillPayload "DATA3|abc", nm, tr
    Debug.Assert nm = "DATA3" And tr = 0#

    Debug.Print "  PASS: TestParseFillPayload"
End Sub

Private Sub TestIsRemoveFillPayload()
    Debug.Assert IsRemoveFillPayload("NONE") = True
    Debug.Assert IsRemoveFillPayload("nofill") = True   ' case-insensitive
    Debug.Assert IsRemoveFillPayload("OFF") = True
    Debug.Assert IsRemoveFillPayload("DATA1") = False

    Debug.Print "  PASS: TestIsRemoveFillPayload"
End Sub

Private Sub TestColorFromName()
    ' Known names resolve to their palette constants.
    Debug.Assert ColorFromName("DATA1") = colorData1
    Debug.Assert ColorFromName("data1") = colorData1     ' case-insensitive
    Debug.Assert ColorFromName("DATA8") = colorData8
    Debug.Assert ColorFromName("NEUTRAL4") = colorNeutral4

    ' Unknown name returns the -1 sentinel.
    Debug.Assert ColorFromName("BOGUS") = -1

    Debug.Print "  PASS: TestColorFromName"
End Sub


' ------------------------------------------------------------
'   modColorSeries - palette order
' ------------------------------------------------------------

Private Sub TestGetPaletteColor()
    ' Default (Contrasting) order: slot i maps to colorData(i) for 1..8.
    ' Note: the alt (Rainbow) branch depends on the module-level m_useAltOrder
    ' flag, which only the object-model TogglePaletteOrder sets, so it is not
    ' covered here - this exercises the default branch and the fallback.
    Debug.Assert GetPaletteColor(1) = colorData1
    Debug.Assert GetPaletteColor(3) = colorData3   ' Sky slot - guards the recolour
    Debug.Assert GetPaletteColor(8) = colorData8

    ' Out-of-range indices fall back to the neutral (Steel).
    Debug.Assert GetPaletteColor(9) = colorNeutral2
    Debug.Assert GetPaletteColor(0) = colorNeutral2

    Debug.Print "  PASS: TestGetPaletteColor"
End Sub


' ------------------------------------------------------------
'   modEngineBuilder - plot-area geometry
' ------------------------------------------------------------

Private Sub TestPlotAreaGeometry()
    ' Assert against the formula composed from the named constants (not magic
    ' numbers) so the test validates which bands are included per flag combo and
    ' stays valid if a constant is retuned.

    ' Top band = titles, plus legend strip and/or y-axis-label strip when shown.
    Debug.Assert PlotAreaTopFor(False, False) = calcTitlesHeight
    Debug.Assert PlotAreaTopFor(False, True) = calcTitlesHeight + LegendHeight
    Debug.Assert PlotAreaTopFor(True, False) = calcTitlesHeight + yAxisLabelHeight + yAxisLabelPad
    Debug.Assert PlotAreaTopFor(True, True) = _
        calcTitlesHeight + LegendHeight + yAxisLabelHeight + yAxisLabelPad

    ' Height = canvas less the top band, the x-title strip (when shown), and the
    ' always-reserved bottom margin + logo. Check the composition identity holds.
    Debug.Assert PlotAreaHeightFor(True, True, True) = _
        chartHeight - PlotAreaTopFor(True, True) - xAxisLabelHeight _
        - plotAreaBottomMargin - logoHeight
    Debug.Assert PlotAreaHeightFor(False, False, False) = _
        chartHeight - PlotAreaTopFor(False, False) - plotAreaBottomMargin - logoHeight

    ' Sanity: the minimal-chrome plot area must still be positive.
    Debug.Assert PlotAreaHeightFor(False, False, False) > 0

    Debug.Print "  PASS: TestPlotAreaGeometry"
End Sub


' ------------------------------------------------------------
'   Assertion helpers
' ------------------------------------------------------------

' Asserts two 0-based arrays hold the same values. VBA's = cannot compare arrays,
' so this checks bounds then walks element by element. label appears in the
' Immediate Window if the assert halts, to identify which case failed.
Private Sub AssertArrayEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal label As String)
    Dim okBounds As Boolean
    okBounds = (LBound(actual) = LBound(expected)) And (UBound(actual) = UBound(expected))
    If Not okBounds Then Debug.Print "  FAIL (length): " & label
    Debug.Assert okBounds

    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If actual(i) <> expected(i) Then Debug.Print "  FAIL (elem " & i & "): " & label
        Debug.Assert actual(i) = expected(i)
    Next i
End Sub
