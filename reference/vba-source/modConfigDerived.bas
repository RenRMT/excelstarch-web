Attribute VB_Name = "modConfigDerived"
Option Explicit
'==== Module: modConfigDerived ====
' Computed constants and chart-engine defaults. These are derived from the user
' settings in modConfig and ensure the chart layout responds correctly to any
' changes made there. Do not edit these directly unless you are changing the chart
' engine itself - adjust the source settings in modConfig instead.
'
' Contents:
'   - Derived geometry (logo, title/plot/source bands) computed from modConfig.
'   - Chart formatting defaults: axis enums, the ChartDefaults type, and the
'     per-chart-type factory functions consumed by the chart pipeline.

' +---------------------------------------------------------+
' |  DERIVED CONSTANTS                                      |
' |  Computed from modConfig. Do not edit directly.         |
' +---------------------------------------------------------+
'=== Logo geometry ===
Public Const logoHeight As Double = chartHeight * logoHeightScale
Public Const logoMarginBottom As Double = chartHeight * logoMarginBottomProp
Public Const logoTop As Double = chartHeight - logoHeight - logoMarginBottom
Public Const logoMarginRight As Double = chartWidth * logoMarginRightProp

'=== Plot area bottom margin ===
Public Const plotAreaBottomMargin As Double = chartHeight * plotAreaBottomMarginProp


'=== Title area ===
' Figure number box aligns with the top of the chart area (no canvas margin).
Public Const figureBoxTop As Double = 0
Public Const figureBoxHeight As Double = chartHeight * FigureBoxHeightProportion
Public Const titleBoxTop As Double = figureBoxHeight
Public Const titleBoxHeight As Double = chartHeight * titleBoxHeightProportion
Public Const subtitleBoxTop As Double = titleBoxTop + titleBoxHeight
Public Const subtitleBoxHeight As Double = chartHeight * subtitleBoxHeightProportion
Public Const titleBoxWidth As Double = chartWidth * titleBoxWidthProportion
Public Const titleBoxNudge As Double = chartWidth * titleBoxNudgeProportion
Public Const calcTitlesHeight As Double = figureBoxHeight + titleBoxHeight + subtitleBoxHeight ' For calculation only
' Pie/donut legend sits just below the subtitle box to avoid overlapping it.
Public Const pieLegendTop As Double = subtitleBoxTop + subtitleBoxHeight + pieLegendGap

'=== Plot area ===
'Legend
Public Const legendTop As Double = calcTitlesHeight
Public Const legendLeftPad As Double = chartWidth * legendLeftPadProportion
Public Const LegendHeight As Double = chartHeight * legendHeightProportion
'Y axis
Public Const yAxisLabelTop As Double = calcTitlesHeight + LegendHeight
Public Const yAxisLabelHeight As Double = chartHeight * yAxisLabelHeightProportion
Public Const yAxisLabelTop_noLegend As Double = calcTitlesHeight
'X axis title (mirrors the Y strip; sits in a band just above the logo/source band)
Public Const xAxisLabelHeight As Double = yAxisLabelHeight
Public Const xAxisLabelTop As Double = chartHeight - logoHeight - plotAreaBottomMargin - xAxisLabelHeight
'Plot area
' Plot-area Top/Height are no longer fixed constants: they depend on which title
' bands (y-title strip, x-title strip) and the legend are present, so they are
' computed at runtime by PlotAreaTopFor / PlotAreaHeightFor in modEngineBuilder.
' Width/Left are still fixed.
Public Const plotAreaWidth As Double = chartWidth
Public Const plotAreaLeft As Double = chartWidth * plotAreaLeftProportion

'=== Source box ===
Public Const sourceBoxWidth As Double = chartWidth * sourceBoxWidthProportion
Public Const sourceBoxLeftNudge As Double = chartWidth * sourceBoxNudgeProportion
Public Const sourceBoxHeight As Double = chartHeight * sourceBoxHeightProportion

'=== Export ===
Public Const exportAppName As String = orgName & " Chart Styles"


' +---------------------------------------------------------+
' |  DEFAULT CHART FORMATTING                               |
' |  Controls pipeline defaults for new charts.             |
' |  Axis constants: 0=none, 1=X only, 2=Y only, 3=both    |
' +---------------------------------------------------------+
'=== Axis selection values (used by defaultGridlines, defaultAxisDisplay, etc.) ===
Public Const axisNone As Long = 0
Public Const axisX As Long = 1
Public Const axisY As Long = 2
Public Const axisBoth As Long = 3

'=== Data label contrast settings ===
Public Const wcagLuminanceThreshold As Double = 0.179   ' WCAG threshold for black vs white label text

'=== Scatter / bubble styling ===
Public Const scatterMarkerSize As Long = 9              ' marker point size (points) for scatter charts
Public Const bubbleTransparency As Single = 0.5         ' bubble fill transparency (0 = opaque, 1 = clear)

'=== Default formatting for new/reformatted charts ===
Public Const defaultGridlines As Long = axisNone        ' gridline visibility
Public Const defaultAxisDisplay As Long = axisNone      ' axis visibility (HasAxis)
Public Const defaultAxisLines As Long = axisNone        ' axis line visibility
Public Const defaultAxisLabels As Long = axisNone       ' tick label visibility
Public Const defaultLegend As Boolean = False           ' False = no legend

'=== ChartDefaults User-Defined Type ===
'Bundles formatting options into a single parameter for chart pipeline.
'Only Gridlines, AxisDisplay, and Legend are currently consumed by ApplyDefaultFormatting.
'ShowYAxisTitle is consumed only by the chartex chrome (modEngineExChrome) - it adds the
'optional worksheet Y-axis title box for chartex types with a value axis (box & whisker);
'classic factories leave it False and ignore it.
'AxisLines and AxisLabels are reserved for future use (phase 6+).
Public Type ChartDefaults
    Gridlines As Long       ' axisNone, axisX, axisY, axisBoth (controls gridline visibility)
    AxisDisplay As Long     ' axisNone, axisX, axisY, axisBoth (controls axis visibility)
    Legend As Boolean       ' True = show legend, False = hide
    ShowYAxisTitle As Boolean  ' chartex only: add a worksheet Y-axis title box
End Type

'=== Factory function for global defaults ===
Public Function DefaultChartDefaults() As ChartDefaults
    With DefaultChartDefaults
        .Gridlines = defaultGridlines
        .AxisDisplay = defaultAxisDisplay
        .Legend = defaultLegend
        .ShowYAxisTitle = False     ' chartex-only; classic charts ignore it
    End With
End Function

'=== Chart-type-specific profile factories ===
Public Function LineChartDefaults() As ChartDefaults
    With LineChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function BarChartDefaults() As ChartDefaults
    With BarChartDefaults
        .Gridlines = axisX          ' X-gridlines only (vertical lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function ColumnChartDefaults() As ChartDefaults
    With ColumnChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function AreaChartDefaults() As ChartDefaults
    With AreaChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function ScatterChartDefaults() As ChartDefaults
    With ScatterChartDefaults
        .Gridlines = axisBoth       ' Both gridlines for reference grid
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function PieChartDefaults() As ChartDefaults
    With PieChartDefaults
        .Gridlines = axisNone       ' No gridlines (pie has no axes)
        .AxisDisplay = axisNone     ' No axes for pie charts
        .Legend = True              ' Pie typically shows legend for slice labels
    End With
End Function

Public Function TreemapChartDefaults() As ChartDefaults
    With TreemapChartDefaults
        .Gridlines = axisNone       ' No gridlines (treemap has no axes)
        .AxisDisplay = axisNone     ' No axes for treemaps
        .Legend = defaultLegend     ' Use global default (tile labels usually suffice)
        .ShowYAxisTitle = False     ' Treemap has no value axis - no Y-axis title box
    End With
End Function

Public Function BoxWhiskerChartDefaults() As ChartDefaults
    With BoxWhiskerChartDefaults
        .Gridlines = axisY          ' Y-gridlines (horizontal value scale)
        .AxisDisplay = axisBoth     ' Box & whisker has a value axis and a category axis
        .Legend = defaultLegend     ' Use global default
        .ShowYAxisTitle = True      ' Has a value axis - add the worksheet Y-axis title box
    End With
End Function
