Attribute VB_Name = "modEngineBuilder"
Option Explicit

' Classic-Chart in-chart formatting pipeline: builds the title/logo/source chrome
' inside cht.Shapes. ChartEx types (treemap, box & whisker, ...) cannot host in-chart
' shapes and use the separate worksheet-chrome pipeline in modEngineExChrome instead.
'
' Shared formatting pipeline applied to every classic chart type.
' colorMode: "FILL" for bar/column charts; "LINE" for line/slope/scatter charts.
' defaults: ChartDefaults UDT containing all five formatting options.
'
' Step order matters:
'   1. OuterFormat    - sets chart size and plot area geometry first; everything else depends on it
'   2. InsertLogo     - anchored to chart bottom-right; independent of plot area
'   3. InsertSource   - anchored to chart bottom-left; must exist before FormatTitle so boxes don't overlap
'   4. FormatTitle    - adds title/subtitle/y-axis label text boxes at top-left
'   5. FormatGridlines - applies major gridline style to value axis
'   6. FormatXAxis    - sizes and colors axis tick labels; runs after gridlines to avoid selection conflicts
'   7. FormatSeriesColors - applied last so series exist and pipeline hasn't altered their format
'
' Chart types that skip steps (slope, dot plot, scatter) call individual functions directly.
Public Sub ApplyChartPipeline(cht As Chart, ByVal colorMode As String, ByRef defaults As ChartDefaults)
    Call OuterFormat(cht, defaults)
    Call InsertLogo(cht)
    Call InsertSource(cht)
    Call FormatTitle(cht)
    Call FormatGridlines(cht)
    Call FormatXAxis(cht)
    Call FormatSeriesColors(cht, UCase$(colorMode))
    Call ApplyDefaultFormatting(cht, defaults)
End Sub


Function OuterFormat(cht As Chart, ByRef defaults As ChartDefaults) As Boolean
    On Error GoTo Fail

    Dim SeriesCount As Long

    'Font
    cht.ChartArea.Font.name = fontPrimary

    'Format axis lines to white
    If cht.HasAxis(xlValue, xlPrimary) Then FormatAxisLineWhite cht.Axes(xlValue, xlPrimary)
    If cht.HasAxis(xlCategory, xlPrimary) Then FormatAxisLineWhite cht.Axes(xlCategory, xlPrimary)

    'Remove axis titles
    RemoveAxisTitles cht

    'Chart size
    If TypeName(cht.Parent) = "ChartObject" Then
        With cht.Parent
            .Width = chartWidth
            .Height = chartHeight
        End With
    End If

    'Remove border
    cht.ChartArea.Border.LineStyle = xlNone

    'Series count
    SeriesCount = cht.SeriesCollection.Count
    If SeriesCount = 0 Then
        OuterFormat = True
        Exit Function
    End If

    'Plot area adjustments
    ApplyPlotAreaGeometry cht, SeriesCount, defaults.Legend

    OuterFormat = True
    Exit Function

Fail:
    OuterFormat = False
End Function


Public Sub FormatAxisLineWhite(ax As Axis)
    'Format axis line to white using Format.Line (no Select required).
    On Error Resume Next
    With ax.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = colorWhite
        .Weight = axisLineWeight
    End With
    On Error GoTo 0
End Sub


Private Sub RemoveAxisTitles(cht As Chart)
    If cht.HasAxis(xlValue) Then
        If cht.Axes(xlValue).HasTitle Then cht.Axes(xlValue).AxisTitle.Delete
    End If
    If cht.HasAxis(xlCategory) Then
        If cht.Axes(xlCategory).HasTitle Then cht.Axes(xlCategory).AxisTitle.Delete
    End If
End Sub


Private Sub ApplyPlotAreaGeometry(cht As Chart, ByVal SeriesCount As Long, ByVal ShowLegend As Boolean)
    Dim pa As PlotArea
    Set pa = cht.PlotArea

    Dim HasMultipleSeries As Boolean, HasLegend As Boolean
    HasMultipleSeries = (SeriesCount > 1)
    HasLegend = HasMultipleSeries And ShowLegend And cht.hasLegend

    'Remove legend if single series or legend disabled in defaults
    If Not HasLegend And cht.hasLegend Then
        cht.Legend.Delete
    End If

    'Position legend and adjust plot area
    If HasLegend Then
        ' xlLegendPositionBottom lays entries out in a horizontal row and lets Excel
        ' auto-size the legend to its content width; setting Top/Left then places it
        ' (and avoids assigning xlLegendPositionCustom directly, which can fail on a
        ' freshly-created legend). Mirrors ToggleLegendStandard.
        With cht.Legend
            .Position = xlLegendPositionBottom
            .Top = legendTop
            .Left = legendLeftPad
            .Font.Color = legendFontColor
            .Font.Size = axisFontSize
        End With

        With pa
            ' Built charts carry a y-title but no x-title (showY=True, showX=False).
            ' Same source of truth as the axis-title toggle, so the two never disagree.
            .Top = PlotAreaTopFor(True, True)
            .Height = PlotAreaHeightFor(True, False, True)
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    Else
        With pa
            .Top = PlotAreaTopFor(True, False)
            .Height = PlotAreaHeightFor(True, False, False)
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    End If
End Sub


' Pure geometry helpers used by the axis-title toggle to recompute the plot area
' for any combination of (y-title present, x-title present, legend present).
' Kept pure (no chart access) so they can be checked from the Immediate window:
'   ? PlotAreaTopFor(True, True)   ? PlotAreaHeightFor(True, True, False)
' The plot-area top depends only on the top bands (legend + y-title strip).
Public Function PlotAreaTopFor(ByVal showY As Boolean, ByVal HasLegend As Boolean) As Double
    PlotAreaTopFor = calcTitlesHeight _
                   + IIf(HasLegend, LegendHeight, 0) _
                   + IIf(showY, yAxisLabelHeight + yAxisLabelPad, 0)
End Function

' Plot-area height is the canvas below the top bands, less the bottom bands
' (x-title strip when present, plus the always-reserved logo + bottom margin).
Public Function PlotAreaHeightFor(ByVal showY As Boolean, ByVal showX As Boolean, ByVal HasLegend As Boolean) As Double
    PlotAreaHeightFor = chartHeight - PlotAreaTopFor(showY, HasLegend) _
                      - IIf(showX, xAxisLabelHeight, 0) _
                      - plotAreaBottomMargin - logoHeight
End Function


' silent: when True, a failure returns False without showing a message. Used by
' the chart-type-agnostic styler, where some chart types (e.g. waterfall,
' sunburst, treemap, funnel) don't support added shapes and the step is meant to
' be skipped quietly. The creation pipeline calls without it and keeps messaging.
Public Function InsertLogo(cht As Chart, Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Fail

    'Decode Base64 to temp file
    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\logo_temp.svg"

    If Not Base64ToFile(LogoPNG_Base64, tmpPath) Then
        MsgLogoDecodeFailed
        InsertLogo = False
        Exit Function
    End If

    'Remove existing logo to avoid duplicates
    SafeDeleteShape cht, "LogoImage"

    'Insert logo at native size
    Dim logoShape As Shape
    Set logoShape = cht.Shapes.AddPicture( _
                Filename:=tmpPath, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=0, Top:=0, _
                Width:=-1, Height:=-1)

    logoShape.name = "LogoImage"

    'Scale to target dimensions. cht.Parent is a ChartObject for embedded charts
    '(has .Width/.Height) but the Workbook for a chart sheet - fall back to the
    'canvas constants so chart sheets still get a correctly-sized logo.
    Dim ChartWidth As Single, ChartHeight As Single
    If TypeName(cht.Parent) = "ChartObject" Then
        ChartWidth = cht.Parent.Width
        ChartHeight = cht.Parent.Height
    Else
        ChartWidth = modConfig.chartWidth
        ChartHeight = modConfig.chartHeight
    End If

    Dim TargetHeight As Single, TargetWidth As Single
    TargetHeight = ChartHeight * logoHeightScale
    TargetWidth = TargetHeight * logoAspectRatio

    logoShape.LockAspectRatio = msoFalse
    logoShape.Height = TargetHeight
    logoShape.Width = TargetWidth

    'Position bottom right
    logoShape.Left = ChartWidth - logoShape.Width - logoMarginRight
    logoShape.Top = ChartHeight - logoShape.Height - logoMarginBottom

    'Clean up temp file
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0

    InsertLogo = True
    Exit Function

Fail:
    InsertLogo = False
    If Not silent Then MsgError "InsertLogo"
End Function


' silent: see InsertLogo - suppresses the failure message for the agnostic styler.
Function InsertSource(cht As Chart, Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Fail

    SafeDeleteShape cht, "SourceBox"

    'Add textbox at bottom-left. cht.Parent is a ChartObject for embedded charts
    'but the Workbook for a chart sheet - fall back to the canvas constant so the
    'source box still positions on a chart sheet.
    Dim SourceBox As Shape
    Dim ChartHeight As Long
    If TypeName(cht.Parent) = "ChartObject" Then
        ChartHeight = cht.Parent.Height
    Else
        ChartHeight = modConfig.chartHeight
    End If

    Set SourceBox = cht.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    0, ChartHeight, sourceBoxWidth, sourceBoxHeight)

    With SourceBox
        .name = "SourceBox"
        .TextFrame.Characters.Text = sourceDefaultText & vbNewLine & notesDefaultText
        .TextFrame.Characters.Font.Size = sourceTextFontSize
        .TextFrame.Characters.Font.name = fontPrimary
        .TextFrame.VerticalAlignment = xlVAlignBottom
        .IncrementLeft -sourceBoxLeftNudge
    End With

    InsertSource = True
    Exit Function

Fail:
    InsertSource = False
    If Not silent Then MsgError "InsertSource"
End Function



Function FormatTitle(cht As Chart) As Boolean
    On Error GoTo Fail

    'Delete existing title-related boxes
    SafeDeleteShape cht, "FigureBox"
    SafeDeleteShape cht, "TitleBox"
    SafeDeleteShape cht, "SubTitleBox"
    SafeDeleteShape cht, "YAxisLabelBox"

    'Remove built-in chart title
    If cht.HasTitle Then cht.ChartTitle.Delete

    'Create all title boxes
    CreateFigureBox cht
    CreateTitleBox cht
    CreateSubtitleBox cht
    CreateYAxisLabelBox cht, cht.hasLegend

    FormatTitle = True
    Exit Function

Fail:
    FormatTitle = False
End Function


Private Sub CreateFigureBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=figureBoxTop, Width:=titleBoxWidth, Height:=figureBoxHeight)

    With shp
        .name = "FigureBox"
        .TextFrame2.TextRange.Text = figureBoxDefaultText
        With .TextFrame2.TextRange.Font
            .Size = figureFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = figureFontColor
            .Bold = msoFalse
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Private Sub CreateTitleBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=titleBoxTop, Width:=titleBoxWidth, Height:=titleBoxHeight)

    With shp
        .name = "TitleBox"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = titleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = titleFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = titleFontColor
            .Bold = msoTrue
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Private Sub CreateSubtitleBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=subtitleBoxTop, Width:=titleBoxWidth, Height:=subtitleBoxHeight)

    With shp
        .name = "SubTitleBox"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = subtitleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = subTitleFontSize
            .Fill.ForeColor.RGB = subTitleFontColor
            .name = fontPrimary
            .Bold = msoFalse
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Public Sub CreateYAxisLabelBox(cht As Chart, ByVal HasLegend As Boolean)
    If Not cht.HasAxis(xlValue) Then Exit Sub

    Dim shp As Shape
    Dim yAxisTop As Single

    yAxisTop = IIf(HasLegend, yAxisLabelTop, yAxisLabelTop_noLegend)

    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=yAxisTop, Width:=titleBoxWidth, Height:=yAxisLabelHeight)

    With shp
        .name = "YAxisLabelBox"
        .TextFrame2.TextRange.Text = yAxisDefaultText
        With .TextFrame2.TextRange.Font
            .Size = axisFontSize
            .name = fontPrimaryItalic
            .Bold = msoFalse
            .Italic = msoTrue
        End With
        .Left = .Left - titleBoxNudge
    End With
End Sub


' Creates the x-axis title box ("XAxisLabelBox") in the band below the plot area,
' mirroring CreateYAxisLabelBox styling. Unlike the y-axis box this is NOT created by
' the pipeline; only the axis-title toggle adds it. Idempotent via SafeDeleteShape.
Public Sub CreateXAxisLabelBox(cht As Chart)
    If Not cht.HasAxis(xlCategory) Then Exit Sub

    SafeDeleteShape cht, "XAxisLabelBox"

    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=xAxisLabelTop, Width:=titleBoxWidth, Height:=xAxisLabelHeight)

    With shp
        .name = "XAxisLabelBox"
        .TextFrame2.TextRange.Text = xAxisDefaultText
        ' Centre the title within the full-width box (the box spans the chart width,
        ' so centred text sits over the middle of the plot rather than hugging the left).
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextFrame2.TextRange.Font
            .Size = axisFontSize
            .name = fontPrimaryItalic
            .Bold = msoFalse
            .Italic = msoTrue
        End With
        .Left = .Left - titleBoxNudge
    End With
End Sub


Function FormatGridlines(cht As Chart) As Boolean
    On Error GoTo Fail

    If Not cht.HasAxis(xlValue) Then
        FormatGridlines = True
        Exit Function
    End If

    Dim ax As Axis
    Set ax = cht.Axes(xlValue)

    ' Add gridlines if missing
    If Not ax.HasMajorGridlines Then
        cht.SetElement msoElementPrimaryValueGridLinesMajor
    End If

    ' Apply major gridline formatting
    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .Weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.RGB = colorNeutral2
    End With

    FormatGridlines = True
    Exit Function

Fail:
    FormatGridlines = False
End Function


Function FormatXAxis(cht As Chart) As Boolean
    On Error GoTo Fail

    'Format category (X) axis
    If cht.HasAxis(xlCategory) Then
        With cht.Axes(xlCategory)
            .TickLabels.Font.Size = axisFontSize
            .TickLabels.Font.Color = legendFontColor
        End With
        FormatCategoryAxisLine cht.Axes(xlCategory)
    End If

    'Format value (Y) axis
    If cht.HasAxis(xlValue) Then
        With cht.Axes(xlValue)
            .TickLabels.Font.Size = axisFontSize
            .TickLabels.Font.Color = axisFontColor
        End With
    End If

    FormatXAxis = True
    Exit Function

Fail:
    FormatXAxis = False
End Function


Private Sub FormatCategoryAxisLine(ax As Axis)
    FormatAxisLineWhite ax
End Sub


Function RemoveShadow(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim seriescount As Long

    ' Get series count safely
    seriescount = cht.SeriesCollection.Count
    If seriescount = 0 Then
        RemoveShadow = True
        Exit Function
    End If

    ' Remove shadow directly
    For i = 1 To seriescount
        With cht.SeriesCollection(i).Format.Shadow
            .Visible = msoFalse
        End With
    Next i

    RemoveShadow = True
    Exit Function

Fail:
    RemoveShadow = False
End Function


Public Sub SafeDeleteShape(cht As Chart, ByVal nm As String)
    On Error Resume Next
    cht.Shapes(nm).Delete
    On Error GoTo 0
End Sub


' Applies defaults after the full pipeline completes.
' Undoes gridlines added by FormatGridlines and removes axes per defaults.AxisDisplay.
' Called as the final step of ApplyChartPipeline so earlier steps can still access axes.
Private Sub ApplyDefaultFormatting(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' --- Gridlines: remove value-axis gridlines unless Y or Both are requested
    If defaults.Gridlines = axisNone Or defaults.Gridlines = axisX Then
        If cht.HasAxis(xlValue) Then
            If cht.Axes(xlValue).HasMajorGridlines Then
                cht.Axes(xlValue).MajorGridlines.Delete
            End If
        End If
    End If

    ' --- Axis display: show/hide each axis per defaults parameter
    Dim showX As Boolean, showY As Boolean
    showX = (defaults.AxisDisplay = axisX Or defaults.AxisDisplay = axisBoth)
    showY = (defaults.AxisDisplay = axisY Or defaults.AxisDisplay = axisBoth)

    cht.HasAxis(xlValue) = showY
    cht.HasAxis(xlCategory) = showX

    Exit Sub
CleanFail:
    MsgError "ApplyDefaultFormatting"
End Sub


'Returns the chart to style, modifying it in-place. Two entry paths:
'  1. A chart is already active  → retype it to chartType; return it.
'  2. A range is selected        → create a new chart of chartType; return it.
'Returns Nothing on any other selection state or error.
Public Function GetTargetChart(ByVal chartType As Long) As Chart
    On Error GoTo Fail

    Set GetTargetChart = Nothing

    'Path 1: Retype active chart
    If Not ActiveChart Is Nothing Then
        ActiveChart.chartType = chartType
        Set GetTargetChart = ActiveChart
        Exit Function
    End If

    'Path 2: Create new chart from selected range
    If TypeName(Selection) <> "Range" Then
        MsgSelectRangeOrChart
        Exit Function
    End If

    On Error Resume Next
    ActiveSheet.Shapes.AddChart2(-1, chartType).Select
    If Not ActiveChart Is Nothing Then
        Set GetTargetChart = ActiveChart
    End If
    On Error GoTo 0

    Exit Function

Fail:
    'Already set to Nothing on entry
End Function
