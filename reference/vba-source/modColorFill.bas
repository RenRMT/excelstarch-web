Attribute VB_Name = "modColorFill"
'==== Module: modColorFill ====
Option Explicit

Private Const LASTUSED_KEY As String = "LastUsedFillTag"

'   TAG DISPATCHER
' Called from modRibbonHandlers. Parses the ribbon button tag and calls ApplyFill or RemoveFill.
' Tag format: "FILL:ColorName" | "FILL:ColorName|0.3" | "FILL:NONE" | "FILL:LASTUSED"
Public Sub ApplyFillFromTag(ByVal tagValue As String)
    tagValue = Trim$(tagValue)

    If InStr(1, tagValue, ":", vbTextCompare) = 0 Then
        MsgInvalidFillTag
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(tagValue, ":")

    Dim payload As String: payload = UCase$(parts(1))

    If IsRemoveFillPayload(payload) Then
        RemoveFill
        Exit Sub
    End If

    If payload = "LASTUSED" Then
        Dim lastTag As String: lastTag = GetLastUsedFillTag()
        If lastTag = "" Then
            ' No prior selection - default to Ocean
            payload = "DATA1"
        Else
            payload = lastTag
        End If
    End If

    ' Pure parse: split "Name|transparency" into its parts.
    Dim colorName As String, transparency As Double
    ParseFillPayload payload, colorName, transparency

    Dim colorRGB As Long: colorRGB = ColorFromName(colorName)
    If colorRGB = -1 Then
        MsgUnknownColor colorName
        Exit Sub
    End If

    ApplyFill colorRGB, transparency
    SaveLastUsedFillTag colorName
End Sub


' True when a fill payload (the part after "FILL:") means "remove the fill".
' Pure; accepts the documented NONE/NOFILL/OFF spellings (case-insensitive).
Public Function IsRemoveFillPayload(ByVal payload As String) As Boolean
    Select Case UCase$(Trim$(payload))
        Case "NONE", "NOFILL", "OFF": IsRemoveFillPayload = True
    End Select
End Function

' Splits a resolved fill payload "Name" or "Name|transparency" into its colour
' name and a transparency in [0, 1]. Pure (no chart/UI). A missing, non-numeric,
' or out-of-range transparency falls back to 0 (opaque). outName/outTransparency
' are always set.
'
' The transparency is parsed with Val, NOT CDbl/IsNumeric: the payload is always
' a fixed period-decimal string (e.g. "DATA1|0.5") from ribbon Tag attributes,
' whereas CDbl/IsNumeric honour the regional decimal separator. On a comma-decimal
' locale (e.g. nl-NL) CDbl("0.5") reads the "." as a group separator and returns 5.
' Val always treats "." as the decimal point, so parsing is locale-invariant.
Public Sub ParseFillPayload(ByVal payload As String, _
                            ByRef outName As String, _
                            ByRef outTransparency As Double)
    Dim subp() As String
    subp = Split(payload, "|")

    outName = subp(0)
    outTransparency = 0

    If UBound(subp) >= 1 Then
        Dim t As Double: t = Val(subp(1))   ' non-numeric -> 0; "." is always the decimal point
        If t < 0 Then t = 0
        If t > 1 Then t = 1
        outTransparency = t
    End If
End Sub


' ColorFromName: Public so modTestHarness can exercise the name -> RGB lookup.
Public Function ColorFromName(ByVal name As String) As Long
    Select Case UCase$(name)
        Case "DATA1":    ColorFromName = colorData1
        Case "DATA2":    ColorFromName = colorData2
        Case "DATA3":    ColorFromName = colorData3
        Case "DATA4":    ColorFromName = colorData4
        Case "DATA5":    ColorFromName = colorData5
        Case "DATA6":    ColorFromName = colorData6
        Case "DATA7":    ColorFromName = colorData7
        Case "DATA8":    ColorFromName = colorData8
        Case "NEUTRAL2": ColorFromName = colorNeutral2
        Case "NEUTRAL4": ColorFromName = colorNeutral4
        Case Else:       ColorFromName = -1
    End Select
End Function


'   FILL
Public Sub ApplyFill(ByVal colorRGB As Long, Optional ByVal transparency As Single = 0!)
    ' transparency: 0 = opaque, 1 = fully transparent (chart model)
    ' For line/scatter series the colour lives on .Format.Line; transparency is not applicable.
    On Error GoTo CleanFail
    AppFast

    Dim tgt As Object
    Set tgt = GetFillTarget()

    If tgt Is Nothing Then
        ' No specific element selected - apply to all series in the active chart.
        Dim cht As Chart
        If Not ActiveChart Is Nothing Then
            Set cht = ActiveChart
        ElseIf TypeName(Selection) = "ChartObject" Then
            Set cht = Selection.Chart
        End If

        If cht Is Nothing Then
            MsgSelectTarget
            GoTo CleanExit
        End If

        Dim i As Long
        For i = 1 To cht.SeriesCollection.Count
            Dim srs As Series
            Set srs = cht.SeriesCollection(i)
            If IsLineTarget(srs) Then
                With srs.Format.Line
                    .Visible = msoTrue
                    .ForeColor.RGB = colorRGB
                End With
            Else
                With srs.Format.Fill
                    .Visible = msoTrue
                    .Solid
                    .ForeColor.RGB = colorRGB
                    If transparency < 0 Then transparency = 0
                    If transparency > 1 Then transparency = 1
                    .transparency = transparency
                End With
            End If
        Next i
        GoTo CleanExit
    End If

    If IsLineTarget(tgt) Then
        With tgt.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = colorRGB
        End With
    Else
        With tgt.Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = colorRGB

            ' Charts: .Transparency 0-1
            ' Shapes:   .Transparency 0-1 (Excel normalizes)
            If transparency < 0 Then transparency = 0
            If transparency > 1 Then transparency = 1
            .transparency = transparency
        End With
    End If

CleanExit:
    AppRestore
    Exit Sub

CleanFail:
    AppRestore
    MsgError "ApplyFill"
End Sub


Public Sub RemoveFill()
    On Error GoTo CleanFail
    AppFast

    Dim tgt As Object
    Set tgt = GetFillTarget()

    If tgt Is Nothing Then
        ' No specific element selected - remove fill from all series in the active chart.
        Dim cht As Chart
        If Not ActiveChart Is Nothing Then
            Set cht = ActiveChart
        ElseIf TypeName(Selection) = "ChartObject" Then
            Set cht = Selection.Chart
        End If

        If cht Is Nothing Then
            MsgSelectTarget
            GoTo CleanExit
        End If

        Dim i As Long
        For i = 1 To cht.SeriesCollection.Count
            Dim srs As Series
            Set srs = cht.SeriesCollection(i)
            If IsLineTarget(srs) Then
                srs.Format.Line.Visible = msoFalse
            Else
                srs.Format.Fill.Visible = msoFalse
            End If
        Next i
        GoTo CleanExit
    End If

    If IsLineTarget(tgt) Then
        tgt.Format.Line.Visible = msoFalse
    Else
        tgt.Format.Fill.Visible = msoFalse
    End If

CleanExit:
    AppRestore
    Exit Sub

CleanFail:
    AppRestore
    MsgError "RemoveFill"
End Sub

'   TARGET DETECTION HELPERS

Private Function IsLineTarget(ByVal tgt As Object) As Boolean
    ' Returns True if tgt is a series whose chart type is line or scatter.
    ' Used to decide whether to write to .Format.Line rather than .Format.Fill.
    Dim srs As Series
    On Error Resume Next
    Set srs = tgt
    On Error GoTo 0
    If srs Is Nothing Then Exit Function    ' not a series - use FILL

    Dim ct As Long
    On Error Resume Next
    ct = srs.chartType
    On Error GoTo 0

    Select Case ct
        Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
             xlLineStacked100, xlLineMarkersStacked100, _
             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            IsLineTarget = True
    End Select
End Function


Private Function GetFillTarget() As Object
    ' Resolve a chart even when a ribbon button click deactivated it before
    ' onAction fired - in that case ActiveChart is Nothing but the ChartObject
    ' remains selected at the worksheet level.
    Dim cht As Chart
    If Not ActiveChart Is Nothing Then
        Set cht = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set cht = Selection.Chart
    End If

    If Not cht Is Nothing Then
        If Not Selection Is Nothing Then
            If IsSeriesOrPoint(Selection) Then
                Set GetFillTarget = Selection
                Exit Function
            End If
        End If
        ' Selection is a chart background element (plot area, chart area, etc.) -
        ' return Nothing so the caller applies fill to all series instead.
        Exit Function
    End If

    If Not Selection Is Nothing Then
        If IsSeriesOrPoint(Selection) Then Set GetFillTarget = Selection
    End If
End Function


Private Function IsSeriesOrPoint(ByVal o As Object) As Boolean
    ' Returns True only for Series and Point objects - the elements the user
    ' intends to colour. Chart area, plot area, walls, etc. are deliberately
    ' excluded so clicks on those background elements fall through to the
    ' "apply to all series" branch in ApplyFill.
    Dim srs As Series
    Dim pt As Point
    On Error Resume Next
    Set srs = o
    If Err.Number = 0 Then IsSeriesOrPoint = Not (srs Is Nothing): Exit Function
    Err.Clear
    Set pt = o
    If Err.Number = 0 Then IsSeriesOrPoint = Not (pt Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function


'   LAST USED TRACKING

Private Sub SaveLastUsedFillTag(ByVal colorName As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(LASTUSED_KEY).Value = colorName
    ' If the property doesn't exist, create it
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add LASTUSED_KEY, , msoPropertyTypeString, colorName
    End If
    On Error GoTo 0
End Sub


Private Function GetLastUsedFillTag() As String
    On Error Resume Next
    Dim val As String
    val = ThisWorkbook.CustomDocumentProperties(LASTUSED_KEY).Value
    If Err.Number = 0 And val <> "" Then
        GetLastUsedFillTag = val
    End If
    Err.Clear
    On Error GoTo 0
End Function
