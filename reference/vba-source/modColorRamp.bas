Attribute VB_Name = "modColorRamp"
'==== Module: modColorRamp ====
' Applies colour ramps to the data series of the active chart.
'
' ApplyColorRamp      - single-hue ramp, steps assigned in spread order 6,2,4,3,5,7,8,1,9,10.
' InvertColorRamp     - reverses the current fill colour assignment across all series.
' ApplyDivergingRamp  - two-hue diverging ramp: dark→light on the left, light→dark on
'                       the right, with an optional grey centre for odd series counts.
'
' Step selection for both single and diverging ramps follows the same priority sequence
' [6,2,4,3,5,7,8,1,9,10]. For diverging ramps the selected steps are then sorted
' numerically (1 = lightest, 10 = darkest) before being assigned as a gradient.
'
' Maximum series: 10 (single), 21 (diverging: 10 + grey + 10).
'
' The step-ordering decisions (OrderedRampSteps, DivergingSideCount/HasMiddle) and
' the tag parser (ParseDivergingTag) are pure functions with no chart dependency,
' kept separate from the object-model fill loops so they can be unit-tested from
' modTestHarness.
Option Explicit

Private Const LASTUSED_RAMP_KEY As String = "LastUsedRampTag"
Private Const LASTUSED_DIV_KEY  As String = "LastUsedDivergingTag"

' ============================================================
'   PUBLIC ENTRY POINTS
' ============================================================

Public Sub InvertColorRamp()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If

    Dim n As Long
    n = cht.SeriesCollection.Count
    If n < 2 Then GoTo CleanExit

    ' Snapshot current fill colors
    Dim colors() As Long
    ReDim colors(1 To n)

    Dim i As Long
    For i = 1 To n
        colors(i) = cht.SeriesCollection(i).Format.Fill.ForeColor.RGB
    Next i

    ' Re-apply in reverse order
    For i = 1 To n
        With cht.SeriesCollection(i).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = colors(n - i + 1)
        End With
    Next i
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "InvertColorRamp"
End Sub

Public Sub ApplyColorRamp(ByVal rampName As String)
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If

    rampName = UCase$(Trim$(rampName))

    If rampName = "LASTUSED" Then
        rampName = GetLastUsedRampTag()
        If rampName = "" Then rampName = "A"   ' default to Ocean
    End If

    BuildColorRamp cht, rampName
    SaveLastUsedRampTag rampName
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "ApplyColorRamp"
End Sub

Public Sub ApplyDivergingRampFromTag(ByVal tagValue As String)
    tagValue = UCase$(Trim$(tagValue))

    If tagValue = "LASTUSED" Then
        tagValue = GetLastUsedDivergingTag()
        If tagValue = "" Then tagValue = "A|B"  ' default to Ocean - Coral
    End If

    Dim leftRamp As String, rightRamp As String
    If Not ParseDivergingTag(tagValue, leftRamp, rightRamp) Then
        MsgInvalidDivergingTag
        Exit Sub
    End If

    ApplyDivergingRamp leftRamp, rightRamp
    SaveLastUsedDivergingTag tagValue
End Sub

' Parses a diverging-ramp tag "LEFT|RIGHT" (e.g. "A|B") into its two ramp names.
' Pure (no UI): returns False when the pipe separator or either side is missing,
' so the caller decides how to report the error. outLeft/outRight are set only on
' success.
Public Function ParseDivergingTag(ByVal tagValue As String, _
                                  ByRef outLeft As String, _
                                  ByRef outRight As String) As Boolean
    Dim parts() As String
    parts = Split(tagValue, "|")
    If UBound(parts) < 1 Then Exit Function          ' no separator → invalid
    If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Then Exit Function

    outLeft = parts(0)
    outRight = parts(1)
    ParseDivergingTag = True
End Function

Public Sub ApplyDivergingRamp(ByVal leftRamp As String, ByVal rightRamp As String)
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If
    BuildDivergingRamp cht, UCase$(Trim$(leftRamp)), UCase$(Trim$(rightRamp))
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "ApplyDivergingRamp"
End Sub


' ============================================================
'   PRIVATE BUILDERS
' ============================================================

Private Sub BuildColorRamp(cht As Chart, ByVal rampName As String)
    Dim n As Long
    n = cht.SeriesCollection.Count
    If n = 0 Then Exit Sub

    If n > 10 Then
        MsgRampTooManySeries
        Exit Sub
    End If

    Dim palette(1 To 10) As Long
    If Not LoadPalette(rampName, palette) Then Exit Sub

    ' Pure decision: the 1-based step index to give each series, darkest first.
    Dim steps As Variant
    steps = OrderedRampSteps(n)

    ' Object-model application: series 1 = darkest, series N = lightest.
    Dim i As Long
    For i = 1 To n
        With cht.SeriesCollection(i).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = palette(steps(i - 1))
        End With
    Next i
End Sub

Private Sub BuildDivergingRamp(cht As Chart, ByVal leftRamp As String, ByVal rightRamp As String)
    Dim n As Long
    n = cht.SeriesCollection.Count
    If n = 0 Then Exit Sub

    If n > 21 Then
        MsgDivergingTooManySeries
        Exit Sub
    End If

    Dim leftPalette(1 To 10) As Long
    Dim rightPalette(1 To 10) As Long
    If Not LoadPalette(leftRamp, leftPalette) Then Exit Sub
    If Not LoadPalette(rightRamp, rightPalette) Then Exit Sub

    ' Pure decision: side size, odd-series middle flag, and the side step indices
    ' sorted ascending (1 = lightest .. 10 = darkest).
    Dim sideCount As Long, hasMiddle As Boolean
    Dim sideSteps As Variant
    sideCount = DivergingSideCount(n)
    hasMiddle = DivergingHasMiddle(n)
    sideSteps = PriorityStepsSorted(sideCount)

    Dim i As Long
    Dim seriesIdx As Long
    seriesIdx = 1

    ' Left side: descending through the sorted steps (dark → light)
    For i = sideCount To 1 Step -1
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = leftPalette(sideSteps(i - 1))
        End With
        seriesIdx = seriesIdx + 1
    Next i

    ' Centre: grey if odd series count
    If hasMiddle Then
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = colorBrand4
        End With
        seriesIdx = seriesIdx + 1
    End If

    ' Right side: ascending through the sorted steps (light → dark)
    For i = 1 To sideCount
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = rightPalette(sideSteps(i - 1))
        End With
        seriesIdx = seriesIdx + 1
    Next i
End Sub


' ============================================================
'   PURE STEP-ORDERING LOGIC (no chart dependency — unit-testable)
' ============================================================
' These functions encode the ramp colour-ordering decisions independently of the
' Excel object model, so they can be exercised from modTestHarness. The Build*
' subs above consume their output and only perform the fill writes.
' All returned arrays are 0-based Variants (from Array()); index with (k - 1) when
' walking a 1-based series/step counter.

' The fixed step-selection priority [6,2,4,3,5,7,8,1,9,10]: which palette steps to use,
' and in what preference order, as the series count grows.
Private Function StepPriority() As Variant
    StepPriority = Array(6, 2, 4, 3, 5, 7, 8, 1, 9, 10)
End Function

' Returns the first `count` priority steps, sorted ascending (1 = lightest ..
' 10 = darkest). Shared by single and diverging ramps. count must be 0..10.
Private Function PriorityStepsSorted(ByVal count As Long) As Variant
    Dim pr As Variant
    pr = StepPriority()

    Dim s() As Integer
    If count <= 0 Then
        PriorityStepsSorted = Array()
        Exit Function
    End If
    ReDim s(0 To count - 1)

    Dim i As Long, j As Long, tmp As Integer
    For i = 0 To count - 1
        s(i) = pr(i)
    Next i

    ' Bubble sort ascending
    For i = 0 To count - 2
        For j = 0 To count - 2 - i
            If s(j) > s(j + 1) Then
                tmp = s(j): s(j) = s(j + 1): s(j + 1) = tmp
            End If
        Next j
    Next i

    PriorityStepsSorted = s
End Function

' Single-hue ramp: the 1-based palette step for each series, darkest first.
' Series 1 gets the darkest selected step, series n the lightest. Returns an
' n-element 0-based array; element (i-1) is the step for series i.
Public Function OrderedRampSteps(ByVal n As Long) As Variant
    Dim asc As Variant
    asc = PriorityStepsSorted(n)        ' ascending: lightest .. darkest
    If n <= 0 Then
        OrderedRampSteps = Array()
        Exit Function
    End If

    Dim out() As Integer
    ReDim out(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        out(i) = asc(n - 1 - i)         ' reverse → darkest first
    Next i
    OrderedRampSteps = out
End Function

' Diverging ramp: number of series on each side (floor(n / 2)).
Public Function DivergingSideCount(ByVal n As Long) As Long
    DivergingSideCount = n \ 2
End Function

' Diverging ramp: True when n is odd, so a grey centre series is inserted.
Public Function DivergingHasMiddle(ByVal n As Long) As Boolean
    DivergingHasMiddle = (n Mod 2 = 1)
End Function


' ============================================================
'   SHARED HELPERS
' ============================================================

' Returns ActiveChart if a chart is in edit mode, or the chart from a selected
' ChartObject (single-click). Handles ribbon buttons deactivating the chart before
' the onAction callback fires.
Public Function ResolveActiveChart() As Chart
    If Not ActiveChart Is Nothing Then
        Set ResolveActiveChart = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set ResolveActiveChart = Selection.Chart
    End If
End Function

' ============================================================
'   LAST USED TRACKING
' ============================================================

Private Sub SaveLastUsedRampTag(ByVal rampName As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(LASTUSED_RAMP_KEY).Value = rampName
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add LASTUSED_RAMP_KEY, , msoPropertyTypeString, rampName
    End If
    On Error GoTo 0
End Sub

Private Function GetLastUsedRampTag() As String
    On Error Resume Next
    Dim val As String
    val = ThisWorkbook.CustomDocumentProperties(LASTUSED_RAMP_KEY).Value
    If Err.Number = 0 Then GetLastUsedRampTag = val
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SaveLastUsedDivergingTag(ByVal tagValue As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(LASTUSED_DIV_KEY).Value = tagValue
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add LASTUSED_DIV_KEY, , msoPropertyTypeString, tagValue
    End If
    On Error GoTo 0
End Sub

Private Function GetLastUsedDivergingTag() As String
    On Error Resume Next
    Dim val As String
    val = ThisWorkbook.CustomDocumentProperties(LASTUSED_DIV_KEY).Value
    If Err.Number = 0 Then GetLastUsedDivergingTag = val
    Err.Clear
    On Error GoTo 0
End Function


' ============================================================
'   SHARED HELPERS
' ============================================================

' Fills a 1-to-10 Long array with the ramp constants for rampName.
' Returns False and shows an error if the name is unrecognised.
Private Function LoadPalette(ByVal rampName As String, palette() As Long) As Boolean
    Select Case rampName
        Case "A"
            palette(1) = rampA1: palette(2) = rampA2: palette(3) = rampA3
            palette(4) = rampA4: palette(5) = rampA5: palette(6) = rampA6
            palette(7) = rampA7: palette(8) = rampA8: palette(9) = rampA9
            palette(10) = rampA10
        Case "B"
            palette(1) = rampB1: palette(2) = rampB2: palette(3) = rampB3
            palette(4) = rampB4: palette(5) = rampB5: palette(6) = rampB6
            palette(7) = rampB7: palette(8) = rampB8: palette(9) = rampB9
            palette(10) = rampB10
        Case "C"
            palette(1) = rampC1: palette(2) = rampC2: palette(3) = rampC3
            palette(4) = rampC4: palette(5) = rampC5: palette(6) = rampC6
            palette(7) = rampC7: palette(8) = rampC8: palette(9) = rampC9
            palette(10) = rampC10
        Case "D"
            palette(1) = rampD1: palette(2) = rampD2: palette(3) = rampD3
            palette(4) = rampD4: palette(5) = rampD5: palette(6) = rampD6
            palette(7) = rampD7: palette(8) = rampD8: palette(9) = rampD9
            palette(10) = rampD10
        Case "E"
            palette(1) = rampE1: palette(2) = rampE2: palette(3) = rampE3
            palette(4) = rampE4: palette(5) = rampE5: palette(6) = rampE6
            palette(7) = rampE7: palette(8) = rampE8: palette(9) = rampE9
            palette(10) = rampE10
        Case "F"
            palette(1) = rampF1: palette(2) = rampF2: palette(3) = rampF3
            palette(4) = rampF4: palette(5) = rampF5: palette(6) = rampF6
            palette(7) = rampF7: palette(8) = rampF8: palette(9) = rampF9
            palette(10) = rampF10
        Case "G"
            palette(1) = rampG1: palette(2) = rampG2: palette(3) = rampG3
            palette(4) = rampG4: palette(5) = rampG5: palette(6) = rampG6
            palette(7) = rampG7: palette(8) = rampG8: palette(9) = rampG9
            palette(10) = rampG10
        Case "H"
            palette(1) = rampH1: palette(2) = rampH2: palette(3) = rampH3
            palette(4) = rampH4: palette(5) = rampH5: palette(6) = rampH6
            palette(7) = rampH7: palette(8) = rampH8: palette(9) = rampH9
            palette(10) = rampH10
        Case Else
            MsgUnknownRamp rampName
            LoadPalette = False
            Exit Function
    End Select
    LoadPalette = True
End Function
