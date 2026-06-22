Attribute VB_Name = "modColorContrast"
'==== Module: modColorContrast ====
' Pure colour-contrast utilities, independent of any chart object.
'
' Used to choose readable data-label text (black vs white) against a coloured
' fill, following WCAG relative-luminance guidance. Kept free of Excel
' object-model calls so the functions can be unit-tested directly from the
' Immediate Window, e.g.
'   ? RelativeLuminance(RGB(28, 103, 88))   ' -> value in [0, 1]
'   ? ContrastColorForFill(RGB(28, 103, 88)) ' -> colorWhite or colorBrand3
'
' Depends only on constants: wcagLuminanceThreshold (modConfigDerived),
' colorWhite and colorBrand3 (modConfig).
Option Explicit


' Returns white or dark brand text for best contrast against the given fill
' colour, using WCAG relative luminance: white text on dark fills, dark text
' on light fills.
Public Function ContrastColorForFill(ByVal fillRGB As Long) As Long
    If RelativeLuminance(fillRGB) < wcagLuminanceThreshold Then
        ContrastColorForFill = colorWhite
    Else
        ContrastColorForFill = colorBrand3
    End If
End Function

' Calculates WCAG relative luminance of an RGB color.
' Input: clr as Long (Excel RGB format: R + G*256 + B*65536)
' Returns: Double in range [0, 1], where 0 is black and 1 is white
' The thresholds (0.03928), exponent (2.4) and channel weights (0.2126 / 0.7152
' / 0.0722) below are the standard sRGB-linearization and relative-luminance
' constants from the WCAG 2.x definition of relative luminance.
Public Function RelativeLuminance(ByVal clr As Long) As Double
    Dim R As Double, G As Double, B As Double
    Dim Rs As Double, Gs As Double, Bs As Double

    ' Extract 8-bit RGB components from Long
    R = clr Mod 256
    G = (clr \ 256) Mod 256
    B = (clr \ 65536) Mod 256

    ' Normalize to 0-1
    Rs = R / 255#
    Gs = G / 255#
    Bs = B / 255#

    ' Convert sRGB to linear RGB
    If Rs <= 0.03928 Then
        Rs = Rs / 12.92
    Else
        Rs = ((Rs + 0.055) / 1.055) ^ 2.4
    End If

    If Gs <= 0.03928 Then
        Gs = Gs / 12.92
    Else
        Gs = ((Gs + 0.055) / 1.055) ^ 2.4
    End If

    If Bs <= 0.03928 Then
        Bs = Bs / 12.92
    Else
        Bs = ((Bs + 0.055) / 1.055) ^ 2.4
    End If

    ' WCAG luminance formula
    RelativeLuminance = 0.2126 * Rs + 0.7152 * Gs + 0.0722 * Bs
End Function
