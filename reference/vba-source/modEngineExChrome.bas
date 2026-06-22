Attribute VB_Name = "modEngineExChrome"
'==== Module: modEngineExChrome ====
' Worksheet-shape chrome for "chartex" charts (the Excel 2016+ family: treemap,
' sunburst, waterfall, funnel, box & whisker, histogram).
'
' Why this exists
' ---------------
' A chartex chart rejects cht.Shapes.AddTextbox / AddPicture with error 1004 - the
' chartex file schema has no userShapes slot, so it cannot own the title/subtitle/
' logo/source overlay boxes that classic charts carry inside cht.Shapes. The chrome
' is therefore created on the HOST WORKSHEET (cht.Parent.Parent.Shapes), positioned
' over the chart, and grouped with the ChartObject so the group exports as one image.
'
' This is deliberately a separate pipeline from the classic in-chart chrome in
' modEngineBuilder (which is left untouched). The chrome is built from a white 600x600
' Canvas behind everything, then FigureBox, TitleBox, SubTitleBox, SourceBox and
' LogoImage on top of it. An optional YAxisTitle box is added for chartex types that
' have a value axis (e.g. box & whisker) via defaults.ShowYAxisTitle; types with no
' value axis (treemap, sunburst, funnel) leave it off. The text boxes are transparent
' with no border; the Canvas supplies the white background.
'
' Per-type differences (has value-axis title? plot-band geometry?) are passed in via
' the ChartDefaults struct and the Position* arguments, so each chartex builder
' (modChartTreemap, modChartBoxWhisker, ...) configures the shared chrome rather than
' forking it.
'
' Coordinate model
' ----------------
' Classic chrome uses chart-relative coordinates on the 600x600 canvas. Worksheet
' shapes use absolute sheet coordinates, so every position is offset by the chart's
' position on the sheet (cht.Parent.Left / cht.Parent.Top). The chart is forced to
' chartWidth x chartHeight before chrome is built, so the same modConfig geometry
' constants apply.
'
' Naming
' ------
' Shape members are named "<ChartObjectName><suffix>" (e.g. "Chart 1_TitleBox") so
' that several chartex charts on one sheet do not collide. The group is named
' "ESTreemapGroup_<ChartObjectName>" so it can be found again on re-run/export. (The
' prefix literal is kept as "ESTreemapGroup_" for backward-compat with groups exported
' by earlier versions; see chartExGroupPrefix.)
'
' Known limitations (prototype)
' -----------------------------
'   - Moving the chart after creation leaves the chrome behind once the group is
'     ungrouped (re-styling ungroups). Re-run the builder to realign.
'   - Re-running with the GROUP selected (rather than the chart) cannot retype the
'     chart; ActiveChart is Nothing and the caller shows MsgSelectRangeOrChart.
'     Click the chart itself, not the group.
'   - Export rasterises the group at ~screen resolution (see modExport), softer than
'     the classic Chart.Export path.
'
' Z-order note / future refinement
' --------------------------------
' The ChartObject is created by the shared GetTargetChart BEFORE this module runs, so
' the white Canvas (added here) lands in front of it and is pushed behind with
' ZOrder msoSendToBack. A cleaner design would create the Canvas FIRST and add the
' chart + chrome on top, making the z-order correct by construction - but that needs a
' chartex-specific creation flow rather than the shared GetTargetChart.
Option Explicit

' Public so modExport can recognise chartex groups by name (single source of truth).
' The literal stays "ESTreemapGroup_" so groups exported by earlier versions still
' resolve; only the constant name is generalised.
Public Const chartExGroupPrefix As String = "ESTreemapGroup_"


' ============================================================
'   PUBLIC ENTRY
' ============================================================

' Builds the chrome shapes on the host worksheet at the given canvas origin, and
' groups them with the ChartObject. Removes any prior chrome/group first so re-runs
' don't duplicate. baseLeft/baseTop are the canvas top-left (resolved by
' ChartExCanvasOrigin before the chart was repositioned). defaults.ShowYAxisTitle adds
' the optional worksheet Y-axis title box. Caller has already shrunk the chart into the
' plot band and coloured the series/tiles.
Public Sub BuildChartExChrome(cht As Chart, ByVal baseLeft As Double, ByVal baseTop As Double, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then
        MsgChartExNeedsEmbedded
        Exit Sub
    End If

    ' Clear any chrome from a previous run before rebuilding.
    RemoveExistingChartExChrome cht

    Dim baseName As String
    baseName = cht.Parent.name

    ' Up to 7 chrome shapes: Canvas + Figure/Title/SubTitle/Source/Logo + optional
    ' YAxisTitle. Members that are not built stay vbNullString and are skipped at group
    ' time, so a chart without a Y-axis title (e.g. treemap) groups exactly 6.
    Dim chromeNames(1 To 7) As String
    Dim shp As Shape

    ' White canvas first so it sits at the back of the z-order; everything else
    ' (chart series + chrome) renders on top of it.
    Set shp = AddChartExCanvas(ws, baseName, baseLeft, baseTop): chromeNames(1) = shp.name

    Set shp = AddChartExFigureBox(ws, baseName, baseLeft, baseTop): chromeNames(2) = shp.name
    Set shp = AddChartExTitleBox(ws, baseName, baseLeft, baseTop): chromeNames(3) = shp.name
    Set shp = AddChartExSubtitleBox(ws, baseName, baseLeft, baseTop): chromeNames(4) = shp.name
    Set shp = AddChartExSourceBox(ws, baseName, baseLeft, baseTop): chromeNames(5) = shp.name

    ' Logo may fail to decode; build the other shapes regardless.
    Set shp = AddChartExLogo(ws, baseName, baseLeft, baseTop)
    If shp Is Nothing Then
        chromeNames(6) = vbNullString
    Else
        chromeNames(6) = shp.name
    End If

    ' Optional value-axis title for chartex types that have one (box & whisker).
    If defaults.ShowYAxisTitle Then
        Set shp = AddChartExYAxisTitle(ws, baseName, baseLeft, baseTop): chromeNames(7) = shp.name
    Else
        chromeNames(7) = vbNullString
    End If

    GroupChartExChrome ws, cht, chromeNames

    Exit Sub
CleanFail:
    MsgError "BuildChartExChrome"
End Sub


' ============================================================
'   SHAPE BUILDERS (worksheet-targeted, offset by chart position)
' ============================================================
' Bodies mirror the classic Create*Box / InsertSource / InsertLogo helpers in
' modEngineBuilder, but target ws.Shapes and add (baseLeft, baseTop) to every
' position. baseName is the ChartObject name, used to make member names unique.

' A borderless white 600x600 rectangle behind the chart and chrome, so the whole
' group exports on a solid white canvas.
Private Function AddChartExCanvas(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=baseLeft, Top:=baseTop, _
                    Width:=chartWidth, Height:=chartHeight)

    With shp
        .name = baseName & "_Canvas"
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = colorWhite
        .Line.Visible = msoFalse
        ' The ChartObject already exists (created before chrome), so a just-added
        ' shape sits in front of it. Send the canvas to the back so the chart series
        ' and chrome render on top of the white background, not behind it.
        .ZOrder msoSendToBack
    End With

    Set AddChartExCanvas = shp
End Function


Private Function AddChartExFigureBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + figureBoxTop, _
                    Width:=titleBoxWidth, Height:=figureBoxHeight)

    With shp
        .name = baseName & "_FigureBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = figureBoxDefaultText
        With .TextFrame2.TextRange.Font
            .Size = figureFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = figureFontColor
            .Bold = msoFalse
        End With
    End With

    Set AddChartExFigureBox = shp
End Function


Private Function AddChartExTitleBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + titleBoxTop, _
                    Width:=titleBoxWidth, Height:=titleBoxHeight)

    With shp
        .name = baseName & "_TitleBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = titleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = titleFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = titleFontColor
            .Bold = msoTrue
        End With
    End With

    Set AddChartExTitleBox = shp
End Function


Private Function AddChartExSubtitleBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + subtitleBoxTop, _
                    Width:=titleBoxWidth, Height:=subtitleBoxHeight)

    With shp
        .name = baseName & "_SubTitleBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = subtitleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = subTitleFontSize
            .Fill.ForeColor.RGB = subTitleFontColor
            .name = fontPrimary
            .Bold = msoFalse
        End With
    End With

    Set AddChartExSubtitleBox = shp
End Function


Private Function AddChartExSourceBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    ' Sit the source box INSIDE the canvas: its bottom aligns with the canvas bottom
    ' edge (baseTop + chartHeight - sourceBoxHeight) and its left with the canvas left.
    ' The source box and the bottom-right logo share the bottom band: they clear each
    ' other only because sourceBoxWidth (0.8 * canvas) ends left of the logo column.
    ' Keep sourceBoxWidthProportion well below ~0.88 (or shrink the logo) or the source
    ' text will slide under the logo.
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    baseLeft, baseTop + chartHeight - sourceBoxHeight, sourceBoxWidth, sourceBoxHeight)

    With shp
        .name = baseName & "_SourceBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = sourceDefaultText & vbNewLine & notesDefaultText
        .TextFrame.Characters.Font.Size = sourceTextFontSize
        .TextFrame.Characters.Font.name = fontPrimary
        .TextFrame.VerticalAlignment = xlVAlignBottom
        .IncrementLeft -sourceBoxLeftNudge
    End With

    Set AddChartExSourceBox = shp
End Function


' Returns Nothing (and shows MsgLogoDecodeFailed) if the embedded logo can't be
' decoded - the rest of the chrome is still built.
Private Function AddChartExLogo(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    On Error GoTo Fail

    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\logo_temp.svg"

    If Not Base64ToFile(LogoPNG_Base64, tmpPath) Then
        MsgLogoDecodeFailed
        Set AddChartExLogo = Nothing
        Exit Function
    End If

    Dim logoShape As Shape
    Set logoShape = ws.Shapes.AddPicture( _
                Filename:=tmpPath, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=baseLeft, Top:=baseTop, _
                Width:=-1, Height:=-1)

    logoShape.name = baseName & "_LogoImage"

    ' Scale against the fixed canvas (chart is forced to chartWidth x chartHeight).
    Dim TargetHeight As Single, TargetWidth As Single
    TargetHeight = chartHeight * logoHeightScale
    TargetWidth = TargetHeight * logoAspectRatio

    logoShape.LockAspectRatio = msoFalse
    logoShape.Height = TargetHeight
    logoShape.Width = TargetWidth

    ' Position bottom-right of the chart, in worksheet coordinates.
    logoShape.Left = baseLeft + chartWidth - logoShape.Width - logoMarginRight
    logoShape.Top = baseTop + chartHeight - logoShape.Height - logoMarginBottom

    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0

    Set AddChartExLogo = logoShape
    Exit Function

Fail:
    MsgLogoDecodeFailed
    Set AddChartExLogo = Nothing
End Function


' Worksheet equivalent of modEngineBuilder.CreateYAxisLabelBox: a horizontal title box
' over the value axis. Only added for chartex types with a value axis (box & whisker)
' via defaults.ShowYAxisTitle. Mirrors the no-legend top position and italic styling
' of the classic box, offset to the canvas origin.
Private Function AddChartExYAxisTitle(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + yAxisLabelTop_noLegend, _
                    Width:=titleBoxWidth, Height:=yAxisLabelHeight)

    With shp
        .name = baseName & "_YAxisTitle"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = yAxisDefaultText
        With .TextFrame2.TextRange.Font
            .Size = axisFontSize
            .name = fontPrimaryItalic
            .Bold = msoFalse
            .Italic = msoTrue
        End With
        .IncrementLeft -titleBoxNudge
    End With

    Set AddChartExYAxisTitle = shp
End Function


' ============================================================
'   GROUPING & LIFECYCLE
' ============================================================

' Groups the chrome shapes + the ChartObject into one named group. On failure the
' shapes are left in place (ungrouped) and the user is told.
Private Function GroupChartExChrome(ws As Worksheet, cht As Chart, ByRef chromeNames() As String) As Shape
    On Error GoTo Fail

    ' Build the member-name list: the ChartObject's own shape + each chrome shape that
    ' was actually created (logo and the optional Y-axis title may be absent). Upper
    ' bound 8 = ChartObject + canvas + 4 text boxes + logo + optional Y-axis title.
    ' Typed as Variant because Shapes.Range expects its index packaged in a Variant - a
    ' typed String() array can raise type-mismatch (error 13) on some Excel builds.
    Dim names() As Variant
    Dim n As Long
    ReDim names(1 To 8)

    n = n + 1: names(n) = cht.Parent.name

    Dim i As Long
    For i = LBound(chromeNames) To UBound(chromeNames)
        If Len(chromeNames(i)) > 0 Then
            n = n + 1: names(n) = chromeNames(i)
        End If
    Next i

    ReDim Preserve names(1 To n)

    Dim grp As Shape
    Set grp = ws.Shapes.Range(names).Group
    grp.name = ChartExGroupName(cht)

    Set GroupChartExChrome = grp
    Exit Function

Fail:
    MsgChartExGroupFailed
    Set GroupChartExChrome = Nothing
End Function


' Removes a prior chartex group and its chrome members so a re-run doesn't duplicate.
' Ungrouping leaves the ChartObject intact (the freshly-retyped chart reference stays
' valid); only the chrome members are deleted.
Private Sub RemoveExistingChartExChrome(cht As Chart)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    baseName = cht.Parent.name

    ' Ungroup the prior group if present (ChartObject survives the ungroup).
    Dim grp As Shape
    Set grp = ws.Shapes(ChartExGroupName(cht))
    If Not grp Is Nothing Then
        If grp.Type = msoGroup Then grp.Ungroup
    End If
    Set grp = Nothing

    ' Delete the prefixed chrome members by name.
    SafeDeleteSheetShape ws, baseName & "_Canvas"
    SafeDeleteSheetShape ws, baseName & "_FigureBox"
    SafeDeleteSheetShape ws, baseName & "_TitleBox"
    SafeDeleteSheetShape ws, baseName & "_SubTitleBox"
    SafeDeleteSheetShape ws, baseName & "_SourceBox"
    SafeDeleteSheetShape ws, baseName & "_LogoImage"
    SafeDeleteSheetShape ws, baseName & "_YAxisTitle"

    On Error GoTo 0
End Sub


' ============================================================
'   LAYOUT (public - called by the builder before chrome)
' ============================================================

' Resolves the canvas top-left (the origin all chrome and the chart band are laid
' out from). On a re-run an existing "<chart>_Canvas" shape exists, so reuse its
' position - this keeps the layout stable even if the group was moved. On first run
' there is no canvas yet, so fall back to the chart's current position.
Public Sub ChartExCanvasOrigin(cht As Chart, ByRef outLeft As Double, ByRef outTop As Double)
    outLeft = cht.Parent.Left
    outTop = cht.Parent.Top

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then Exit Sub

    Dim canvas As Shape
    On Error Resume Next
    Set canvas = ws.Shapes(cht.Parent.name & "_Canvas")
    On Error GoTo 0
    If Not canvas Is Nothing Then
        outLeft = canvas.Left
        outTop = canvas.Top
    End If
End Sub


' Positions the chartex ChartObject as a band inside the canvas, leaving the title
' block above and the logo/source below. Mirrors the standard plot-area geometry via
' the shared helpers: inset left/right by plotAreaLeft, top/height from
' PlotAreaTopFor/PlotAreaHeightFor - all relative to the canvas origin. showY/showX/
' hasLegend let each chartex type reserve the right bands (treemap: True/False/False -
' value-axis-title band above, no category strip below; box & whisker adds showX).
Public Sub PositionChartExChart(cht As Chart, ByVal baseLeft As Double, ByVal baseTop As Double, _
                                ByVal showY As Boolean, ByVal showX As Boolean, ByVal hasLegend As Boolean)
    With cht.Parent
        .Left = baseLeft + plotAreaLeft
        .Top = baseTop + PlotAreaTopFor(showY, hasLegend)
        .Width = chartWidth - 2 * plotAreaLeft
        .Height = PlotAreaHeightFor(showY, showX, hasLegend)
    End With
End Sub


' ============================================================
'   HELPERS
' ============================================================

' The host worksheet for an embedded chart, or Nothing for a chart sheet
' (cht.Parent is the Workbook) - chart sheets are out of scope.
Private Function HostSheet(cht As Chart) As Worksheet
    On Error Resume Next
    If TypeName(cht.Parent) = "ChartObject" Then
        Set HostSheet = cht.Parent.Parent
    End If
    On Error GoTo 0
End Function


' Deterministic group name for a given chart.
' Precondition: cht.Parent is a ChartObject (embedded chart). All callers reach this
' only after BuildChartExChrome's HostSheet guard has passed.
Private Function ChartExGroupName(cht As Chart) As String
    ChartExGroupName = chartExGroupPrefix & cht.Parent.name
End Function


' Deletes a worksheet shape by name if present (mirrors SafeDeleteShape, which only
' searches cht.Shapes).
Private Sub SafeDeleteSheetShape(ws As Worksheet, ByVal nm As String)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
End Sub
