Attribute VB_Name = "modConfig"
Option Explicit
'==== Module: modConfig ====
' Brand and user-editable settings - the values you adapt to your own house style.
' Everything here can be safely changed: brand/data colours and ramps, chart canvas
' dimensions, font types & sizes, font colours, logo size, margins, series spacing,
' and the default title/subtitle placeholder texts. The chart layout is generally
' responsive to changes you make here (you might need to test font sizes a little).
'
' Computed/derived values and the chart-engine defaults live in modConfigDerived;
' you should not need to touch that module unless you are changing the chart engine
' itself.


' +---------------------------------------------------------+
' |  BRAND COLOURS                                          |
' |  Edit these to match your house palette.                |
' +---------------------------------------------------------+
'=== Brand colors ===
' Note: colorBrand1, colorBrand2 are defined for completeness
' but are not currently referenced in code. Reserved for future ribbon buttons.
Public Const colorBrand1 As Long = 10963739    'Primary Blue #1B4BA7 RGB(27, 75, 167)
Public Const colorBrand2 As Long = 2888711     'Dark Blue    #07142C RGB(7, 20, 44)
Public Const colorBrand3 As Long = 655874      'Black        #02020A RGB(2, 2, 10)
' Brand Light Grey (#F9F9F9) - reserved for future ribbon buttons.
Public Const colorBrandLightGrey As Long = 16382457 'Light Grey #F9F9F9 RGB(249, 249, 249)
' colorBrand4 is the diverging-ramp neutral centre (#F9F9F9) so the centre series
' reads as near-white. In this palette it equals the brand Light Grey above (both
' #F9F9F9). Used by modColorRamp.BuildDivergingRamp.
Public Const colorBrand4 As Long = 16382457    'Neutral centre #F9F9F9 RGB(249, 249, 249)

'=== Neutral colors ===
' Note: colorNeutral3 is defined for completeness but not currently referenced in code.
Public Const colorNeutral1 As Long = 14540253  'Silver RGB(221, 221, 221) #DDDDDD
Public Const colorNeutral2 As Long = 12303291  'Steel  RGB(187, 187, 187) #BBBBBB
Public Const colorNeutral3 As Long = 10263708  'Ash    RGB(156, 156, 156) - reserved
Public Const colorNeutral4 As Long = 16777215  'White  RGB(255, 255, 255) - used by name-lookup palette (modColorFill)
Public Const colorWhite As Long = colorNeutral4  'Semantic alias - use this for axis/border white styling

'=== Data colors (contrasting order) ===
' Slot 8 (Steel) is intentionally the same neutral grey as colorNeutral2 (the
' >8-series fallback): the 8th categorical series reads as grey by design.
Public Const colorData1 As Long = 12285696     'Ocean    #0077BB RGB(0, 119, 187)
Public Const colorData2 As Long = 6719743      'Coral    #FF8866 RGB(255, 136, 102)
Public Const colorData3 As Long = 16764023     'Sky      #77CCFF RGB(119, 204, 255)
Public Const colorData4 As Long = 8952064      'Pine     #009988 RGB(0, 153, 136)
Public Const colorData5 As Long = 3399167      'Gold     #FFDD33 RGB(255, 221, 51)
Public Const colorData6 As Long = 17578        'Rust     #AA4400 RGB(170, 68, 0)
Public Const colorData7 As Long = 15636906     'Lavender #AA99EE RGB(170, 153, 238)
Public Const colorData8 As Long = 12303291     'Steel    #BBBBBB RGB(187, 187, 187)

'== Color ramp ==
' Each ramp is a 10-step sequential palette (1 = lightest .. 10 = darkest).
' rampA = Ocean
Public Const rampA1 As Long = 16314854  '#E6F1F8 RGB(230, 241, 248)
Public Const rampA2 As Long = 15852748  '#CCE4F1 RGB(204, 228, 241)
Public Const rampA3 As Long = 14993817  '#99C9E4 RGB(153, 201, 228)
Public Const rampA4 As Long = 14069094  '#66ADD6 RGB(102, 173, 214)
Public Const rampA5 As Long = 13210163  '#3392C9 RGB(51, 146, 201)
Public Const rampA6 As Long = 12285696  '#0077BB RGB(0, 119, 187)
Public Const rampA7 As Long = 9854720   '#005F96 RGB(0, 95, 150)
Public Const rampA8 As Long = 7358208   '#004770 RGB(0, 71, 112)
Public Const rampA9 As Long = 4927488   '#00304B RGB(0, 48, 75)
Public Const rampA10 As Long = 2430976  '#001825 RGB(0, 24, 37)

' rampB = Coral
Public Const rampB1 As Long = 15791103  '#FFF3F0 RGB(255, 243, 240)
Public Const rampB2 As Long = 14739455  '#FFE7E0 RGB(255, 231, 224)
Public Const rampB3 As Long = 12767231  '#FFCFC2 RGB(255, 207, 194)
Public Const rampB4 As Long = 10729727  '#FFB8A3 RGB(255, 184, 163)
Public Const rampB5 As Long = 8757503   '#FFA085 RGB(255, 160, 133)
Public Const rampB6 As Long = 6719743   '#FF8866 RGB(255, 136, 102)
Public Const rampB7 As Long = 5402060   '#CC6D52 RGB(204, 109, 82)
Public Const rampB8 As Long = 4018841   '#99523D RGB(153, 82, 61)
Public Const rampB9 As Long = 2700902   '#663629 RGB(102, 54, 41)
Public Const rampB10 As Long = 1317683  '#331B14 RGB(51, 27, 20)

' rampC = Sky
Public Const rampC1 As Long = 16775921  '#F1FAFF RGB(241, 250, 255)
Public Const rampC2 As Long = 16774628  '#E4F5FF RGB(228, 245, 255)
Public Const rampC3 As Long = 16772041  '#C9EBFF RGB(201, 235, 255)
Public Const rampC4 As Long = 16769197  '#ADE0FF RGB(173, 224, 255)
Public Const rampC5 As Long = 16766610  '#92D6FF RGB(146, 214, 255)
Public Const rampC6 As Long = 16764023  '#77CCFF RGB(119, 204, 255)
Public Const rampC7 As Long = 13411167  '#5FA3CC RGB(95, 163, 204)
Public Const rampC8 As Long = 10058311  '#477A99 RGB(71, 122, 153)
Public Const rampC9 As Long = 6705712   '#305266 RGB(48, 82, 102)
Public Const rampC10 As Long = 3352856  '#182933 RGB(24, 41, 51)

' rampD = Pine
Public Const rampD1 As Long = 15988198  '#E6F5F3 RGB(230, 245, 243)
Public Const rampD2 As Long = 15199180  '#CCEBE7 RGB(204, 235, 231)
Public Const rampD3 As Long = 13620889  '#99D6CF RGB(153, 214, 207)
Public Const rampD4 As Long = 12108390  '#66C2B8 RGB(102, 194, 184)
Public Const rampD5 As Long = 10530099  '#33ADA0 RGB(51, 173, 160)
Public Const rampD6 As Long = 8952064   '#009988 RGB(0, 153, 136)
Public Const rampD7 As Long = 7174656   '#007A6D RGB(0, 122, 109)
Public Const rampD8 As Long = 5397504   '#005C52 RGB(0, 92, 82)
Public Const rampD9 As Long = 3554560   '#003D36 RGB(0, 61, 54)
Public Const rampD10 As Long = 1777408  '#001F1B RGB(0, 31, 27)

' rampE = Gold
Public Const rampE1 As Long = 15465727  '#FFFCEB RGB(255, 252, 235)
Public Const rampE2 As Long = 14088447  '#FFF8D6 RGB(255, 248, 214)
Public Const rampE3 As Long = 11399679  '#FFF1AD RGB(255, 241, 173)
Public Const rampE4 As Long = 8776703   '#FFEB85 RGB(255, 235, 133)
Public Const rampE5 As Long = 6087935   '#FFE45C RGB(255, 228, 92)
Public Const rampE6 As Long = 3399167   '#FFDD33 RGB(255, 221, 51)
Public Const rampE7 As Long = 2732492   '#CCB129 RGB(204, 177, 41)
Public Const rampE8 As Long = 2065817   '#99851F RGB(153, 133, 31)
Public Const rampE9 As Long = 1333350   '#665814 RGB(102, 88, 20)
Public Const rampE10 As Long = 666675   '#332C0A RGB(51, 44, 10)

' rampF = Rust
Public Const rampF1 As Long = 15133943  '#F7ECE6 RGB(247, 236, 230)
Public Const rampF2 As Long = 13425390  '#EEDACC RGB(238, 218, 204)
Public Const rampF3 As Long = 10073309  '#DDB499 RGB(221, 180, 153)
Public Const rampF4 As Long = 6721484   '#CC8F66 RGB(204, 143, 102)
Public Const rampF5 As Long = 3369403   '#BB6933 RGB(187, 105, 51)
Public Const rampF6 As Long = 17578     '#AA4400 RGB(170, 68, 0)
Public Const rampF7 As Long = 13960     '#883600 RGB(136, 54, 0)
Public Const rampF8 As Long = 10598     '#662900 RGB(102, 41, 0)
Public Const rampF9 As Long = 6980      '#441B00 RGB(68, 27, 0)
Public Const rampF10 As Long = 3618     '#220E00 RGB(34, 14, 0)

' rampG = Lavender
Public Const rampG1 As Long = 16643575  '#F7F5FD RGB(247, 245, 253)
Public Const rampG2 As Long = 16575470  '#EEEBFC RGB(238, 235, 252)
Public Const rampG3 As Long = 16307933  '#DDD6F8 RGB(221, 214, 248)
Public Const rampG4 As Long = 16106188  '#CCC2F5 RGB(204, 194, 245)
Public Const rampG5 As Long = 15838651  '#BBADF1 RGB(187, 173, 241)
Public Const rampG6 As Long = 15636906  '#AA99EE RGB(170, 153, 238)
Public Const rampG7 As Long = 12483208  '#887ABE RGB(136, 122, 190)
Public Const rampG8 As Long = 9395302   '#665C8F RGB(102, 92, 143)
Public Const rampG9 As Long = 6241604   '#443D5F RGB(68, 61, 95)
Public Const rampG10 As Long = 3153698  '#221F30 RGB(34, 31, 48)

' rampH = Steel (neutral grey ramp; defined for completeness, not exposed in the
' ramp menu since Steel is a neutral rather than a sequential brand hue).
Public Const rampH1 As Long = 16316664  '#F8F8F8 RGB(248, 248, 248)
Public Const rampH2 As Long = 15856113  '#F1F1F1 RGB(241, 241, 241)
Public Const rampH3 As Long = 15000804  '#E4E4E4 RGB(228, 228, 228)
Public Const rampH4 As Long = 14079702  '#D6D6D6 RGB(214, 214, 214)
Public Const rampH5 As Long = 13224393  '#C9C9C9 RGB(201, 201, 201)
Public Const rampH6 As Long = 12303291  '#BBBBBB RGB(187, 187, 187)
Public Const rampH7 As Long = 9868950   '#969696 RGB(150, 150, 150)
Public Const rampH8 As Long = 7368816   '#707070 RGB(112, 112, 112)
Public Const rampH9 As Long = 4934475   '#4B4B4B RGB(75, 75, 75)
Public Const rampH10 As Long = 2434341  '#252525 RGB(37, 37, 37)


' +---------------------------------------------------------+
' |  USER SETTINGS                                          |
' |  Edit these constants to customise the chart style.     |
' +---------------------------------------------------------+

'=== Identity settings ===
Public Const orgName As String = "COMPANY"

'=== Canvas settings ===
' Canvas is measured in Excel points (1pt = 1/72" or about 13/360cm).
' Origin is the top-left corner of the chart area.
Public Const chartWidth As Double = 600         ' 20cm canvas width
Public Const chartHeight As Double = 600        ' 20cm canvas height

'=== Placeholder text settings ===
' Chart text boxes
' all chart text elements come pre-filled with placeholder text.
' use this text to convey standards relating to chart text,
' and to show the standard font colors for these texts.
Public Const figureBoxDefaultText As String = "Figure XX (optional)"
Public Const titleDefaultText    As String = "Title in 28pt sentence case"
Public Const subtitleDefaultText As String = "Subtitle in 22pt sentence case"
Public Const yAxisDefaultText    As String = "Y axis title (unit)"
Public Const xAxisDefaultText    As String = "X axis title (unit)"
Public Const sourceDefaultText   As String = "Source: Source text goes here."
Public Const notesDefaultText    As String = "Notes: Notes text goes here."

' Export
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"


'=== Font settings ===
' Font family
' - fontPrimary: font used for most text boxes, including title, legend, source box.
' - fontPrimaryItalic: by default, italic font is only used for Y-/X-axis labels.
'   Set to same value as fontPrimary if you don't want to use italic font.
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"

' Font sizes
' font sizes expressed in Excel points
' - generalFontSize: is currently not used for anything but kept for compatibility
Public Const titleFontSize As Double = 28
Public Const subTitleFontSize As Double = 22
Public Const figureFontSize As Double = 18
Public Const axisFontSize As Double = 18
Public Const sourceTextFontSize As Double = 14
Public Const generalFontSize As Double = 18

'Font colors
' colors as defined in the BRAND COLOURS section above.
' - generalFontColor: is currently not used for anything but kept for compatibility
Public Const titleFontColor As Long = colorBrand1
Public Const subTitleFontColor As Long = colorBrand2
Public Const figureFontColor As Long = colorBrand3

Public Const axisFontColor As Long = colorBrand3
Public Const legendFontColor As Long = colorBrand3
Public Const sourceFontColor As Long = colorBrand3
Public Const generalFontColor As Long = colorBrand3

'=== Chart Data settings ===
' General data series settings
' - seriesGapWidth: amount of horizontal space between data series, expressed as a
'   percentage of the series width.
' - seriesOverlap: amount of overlap between data series. negative values create distance,
'   positive values create overlap. Setting to 0 makes data series touch. Note that
'   this setting is overriden for stacked bar/column charts.
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = -5

' Lollipop chart settings
' Lollipop charts are generated a bit differently and require their own settings.
' - lollipopGapWidth: set width between lollipop chart series.
' - lollipopStickWeight: the size of the lollipop sticks expressed in Excel points
Public Const lollipopGapWidth As Double = 150
Public Const lollipopStickWeight As Single = 2

' Pie chart settings
Public Const pieplotAreaSize_legend As Long = 400   ' width and height (square) when legend present
Public Const pieplotAreaSize_noLegend As Long = 447 ' width and height (square) without legend
Public Const pieplotAreaLeft As Long = 131
Public Const pieplotAreaTop As Long = 53
Public Const piePlotTopRatio As Double = 0.75   ' vertical centering ratio
Public Const pieLegendGap As Double = 6          ' gap between subtitle box and legend
' Note: pieLegendTop is a derived value in modConfigDerived, since it derives
' from subtitleBoxTop/subtitleBoxHeight (VBA Const cannot forward-reference).

' Weights
'   - gridLineWeight:  weight of chart gridlines expressed in Excel points
'   - axisLineWeight:weight of chart axis lines expressed in Excel points
Public Const gridlineWeight As Double = 1
Public Const axisLineWeight As Double = 1

'=== Chart Actions settings ===
'Annotation box
Public Const annotationDefaultText As String = "Annotation"
Public Const annotationBoxWidth As Double = 120
Public Const annotationBoxHeight As Double = 30
Public Const annotationOffsetX As Double = 8     ' box left = point + offset
Public Const annotationOffsetY As Double = -8    ' box top  = point - offset (sit above)
Public Const annotationFontSize As Double = axisFontSize   ' reuse existing axis size
Public Const annotationFontColor As Long = axisFontColor   ' reuse existing axis colour


' === Box sizes ===
' Box sizes expressed as proportion of chart width / height
Public Const FigureBoxHeightProportion As Double = 0.04
Public Const titleBoxHeightProportion As Double = 0.07
Public Const subtitleBoxHeightProportion As Double = 0.05
Public Const yAxisLabelHeightProportion As Double = 0.04
Public Const legendHeightProportion As Double = 0.04
Public Const titleBoxWidthProportion As Double = 1 'keep this as 1 unless you are moving logo to the top
Public Const titleBoxNudgeProportion As Double = 0
'bottom
Public Const sourceBoxWidthProportion As Double = 0.8
Public Const sourceBoxHeightProportion As Double = 0.08
Public Const sourceBoxNudgeProportion As Double = 0.01

'padding
Public Const legendLeftPadProportion As Double = 0
Public Const plotAreaLeftProportion As Double = 0.005
Public Const yAxisLabelPad As Double = 10

'=== Layout and logo ===
' - logoFileType: Embedded logo accepts PNG or SVG files
' - logoHeightScale: as proportion of chart height.
' - logoAspectRatio: the keep aspect ratio setting in Excel does not work properly.
'   setting it here prevents your logo from getting distorted on chart resize.
' - plotAreaBottomMarginProp: reserved space below plot area to prevent X-axis labels
'   from overlapping the logo. Adjust based on axis label height and desired spacing.
Public Const logoFileType As String = "svg"
Public Const logoHeightScale As Double = 0.1        ' logo height as fraction of chart height
Public Const logoAspectRatio As Double = 2.08          ' logo width = aspectRatio x height
Public Const logoMarginRightProp As Double = 0.01 'chartWidth * 0.01
Public Const logoMarginBottomProp As Double = 0.01 'chartHeight * 0.01
Public Const plotAreaBottomMarginProp As Double = 0.03 ' clearance between x-axis labels and top logo, as fraction of chart height
