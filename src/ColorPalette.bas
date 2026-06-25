Attribute VB_Name = "ColorPalette"
Option Explicit

' =====================================================================
'  Color Palette  -  think-cell-style quick colour applicator
'  Windows PowerPoint VBA add-in
'
'  - Swatches are read LIVE from the active presentation's theme
'  - Toggle between Font / Fill / Outline, then click a swatch
'  - Transparent / No Fill / No Line
'  - Custom hex swatches that persist between sessions
' =====================================================================

Public gMode As String                      ' "FONT" | "FILL" | "LINE"

' ---- Launcher (assign this to a ribbon / QAT button) ---------------
Public Sub ShowPalette()
    gMode = "FILL"
    On Error Resume Next
    Unload frmPalette                       ' rebuild from the current theme
    On Error GoTo 0
    frmPalette.Show vbModeless
End Sub

' ---- Theme colours from the ACTIVE presentation -------------------
Public Function ThemeRGB(ByVal idx As Long) As Long
    On Error Resume Next
    ThemeRGB = ActivePresentation.SlideMaster.Theme.ThemeColorScheme(idx).RGB
End Function

' ---- Tint / shade helpers (same maths PowerPoint uses) ------------
Public Function Lighter(ByVal c As Long, ByVal pct As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&: g = (c \ &H100&) And &HFF&: b = (c \ &H10000) And &HFF&
    r = r + (255 - r) * pct
    g = g + (255 - g) * pct
    b = b + (255 - b) * pct
    Lighter = RGB(r, g, b)
End Function

Public Function Darker(ByVal c As Long, ByVal pct As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&: g = (c \ &H100&) And &HFF&: b = (c \ &H10000) And &HFF&
    r = r * (1 - pct): g = g * (1 - pct): b = b * (1 - pct)
    Darker = RGB(r, g, b)
End Function

' ---- Apply a colour to the current selection ----------------------
Public Sub ApplyColor(ByVal c As Long)
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Sub
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    Select Case sel.Type
        Case ppSelectionText
            If gMode = "FONT" Then
                sel.TextRange.Font.Color.RGB = c
            Else
                ApplyToShapes sel.ShapeRange, c
            End If
        Case ppSelectionShapes
            ApplyToShapes sel.ShapeRange, c
        Case Else
            ' nothing selected - ignore quietly
    End Select
End Sub

Private Sub ApplyToShapes(ByVal shp As ShapeRange, ByVal c As Long)
    Dim s As Shape
    For Each s In shp
        Select Case gMode
            Case "FILL"
                s.Fill.Visible = msoTrue
                s.Fill.Solid
                s.Fill.ForeColor.RGB = c
            Case "LINE"
                s.Line.Visible = msoTrue
                s.Line.ForeColor.RGB = c
            Case "FONT"
                If s.HasTextFrame Then s.TextFrame.TextRange.Font.Color.RGB = c
        End Select
    Next s
End Sub

' ---- Transparent / No Fill / No Line ------------------------------
Public Sub ApplyNoColor()
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Sub
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    Dim shp As ShapeRange
    Select Case sel.Type
        Case ppSelectionShapes: Set shp = sel.ShapeRange
        Case ppSelectionText:   Set shp = sel.ShapeRange
        Case Else: Exit Sub
    End Select
    Dim s As Shape
    For Each s In shp
        If gMode = "FILL" Then s.Fill.Visible = msoFalse
        If gMode = "LINE" Then s.Line.Visible = msoFalse
        ' FONT has no "transparent" - ignored
    Next s
End Sub

' ---- Custom swatch persistence (registry) -------------------------
Public Function LoadCustom() As String
    LoadCustom = GetSetting("PPColorPalette", "Swatches", "Custom", "")
End Function

Public Sub SaveCustom(ByVal csv As String)
    SaveSetting "PPColorPalette", "Swatches", "Custom", csv
End Sub
