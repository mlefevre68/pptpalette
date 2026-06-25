' =====================================================================
'  frmPalette  -  paste ALL of this into the blank UserForm's code.
'  (Insert > UserForm, rename it frmPalette, double-click it, paste.)
' =====================================================================
Option Explicit

Private mCtrls As Collection                 ' keeps clsSwatch instances alive
Private mUID As Long                         ' unique control-name counter
Public CurrentMode As String
Private mLblFont As MSForms.Label
Private mLblFill As MSForms.Label
Private mLblLine As MSForms.Label
Private mCustX As Single, mCustY As Single

Private Const SW As Single = 16              ' swatch size (pixels)
Private Const GAP As Single = 2

Private Sub UserForm_Initialize()
    Set mCtrls = New Collection
    Me.Caption = "Colour Palette"
    Me.BackColor = RGB(245, 245, 245)
    BuildUI
    SetMode "FILL"
End Sub

Private Sub BuildUI()
    Dim x As Single, y As Single, col As Long, baseC As Long

    ' ---- mode toggles ----
    y = 6
    AddModeButton "FONT", "Font", 6, y, 40
    AddModeButton "FILL", "Fill", 48, y, 40
    AddModeButton "LINE", "Line", 90, y, 40

    ' ---- theme colours + tints / shades ----
    y = 32
    Dim themeIdx As Variant
    themeIdx = Array(msoThemeColorDark1, msoThemeColorLight1, _
                     msoThemeColorDark2, msoThemeColorLight2, _
                     msoThemeColorAccent1, msoThemeColorAccent2, _
                     msoThemeColorAccent3, msoThemeColorAccent4, _
                     msoThemeColorAccent5, msoThemeColorAccent6)
    For col = 0 To 9
        baseC = ColorPalette.ThemeRGB(CLng(themeIdx(col)))
        x = 6 + col * (SW + GAP)
        AddSwatch x, y, baseC
        AddSwatch x, y + 1 * (SW + GAP), ColorPalette.Lighter(baseC, 0.8)
        AddSwatch x, y + 2 * (SW + GAP), ColorPalette.Lighter(baseC, 0.5)
        AddSwatch x, y + 3 * (SW + GAP), ColorPalette.Lighter(baseC, 0.25)
        AddSwatch x, y + 4 * (SW + GAP), ColorPalette.Darker(baseC, 0.25)
        AddSwatch x, y + 5 * (SW + GAP), ColorPalette.Darker(baseC, 0.5)
    Next col

    ' ---- standard colours (PowerPoint's fixed row) ----
    y = y + 6 * (SW + GAP) + 8
    Dim sc(0 To 9) As Long
    sc(0) = RGB(192, 0, 0):  sc(1) = RGB(255, 0, 0):   sc(2) = RGB(255, 192, 0)
    sc(3) = RGB(255, 255, 0): sc(4) = RGB(146, 208, 80): sc(5) = RGB(0, 176, 80)
    sc(6) = RGB(0, 176, 240): sc(7) = RGB(0, 112, 192): sc(8) = RGB(0, 32, 96)
    sc(9) = RGB(112, 48, 160)
    For col = 0 To 9
        AddSwatch 6 + col * (SW + GAP), y, sc(col)
    Next col

    ' ---- transparent / no fill ----
    y = y + SW + 10
    AddNoneButton 6, y

    ' ---- custom hex entry ----
    y = y + SW + 10
    Dim tb As MSForms.TextBox
    Set tb = Me.Controls.Add("Forms.TextBox.1", "txtHex", True)
    tb.Left = 6: tb.Top = y: tb.Width = 70: tb.Height = 18
    tb.Text = "#RRGGBB"
    AddTextButton "Set", 80, y, "HEXSET"
    AddTextButton "+", 116, y, "HEXADD"

    ' ---- saved custom swatches ----
    y = y + SW + 16
    BuildCustomRow 6, y

    Me.Width = 6 + 10 * (SW + GAP) + 24
    Me.Height = y + SW + 44
End Sub

' ---------------------------------------------------------------------
Private Function NewLabel(ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.Label
    mUID = mUID + 1
    Dim lb As MSForms.Label
    Set lb = Me.Controls.Add("Forms.Label.1", nm & "_" & mUID, True)
    lb.Left = l: lb.Top = t: lb.Width = w: lb.Height = h
    lb.BorderStyle = fmBorderStyleSingle
    Set NewLabel = lb
End Function

Private Sub AddSwatch(ByVal l As Single, ByVal t As Single, ByVal c As Long)
    Dim sw As clsSwatch: Set sw = New clsSwatch
    sw.Kind = "COLOR": sw.ColorVal = c
    Set sw.Lbl = NewLabel("sw", l, t, SW, SW)
    sw.Lbl.BackColor = c
    mCtrls.Add sw
End Sub

Private Sub AddModeButton(ByVal modeVal As String, ByVal cap As String, _
                          ByVal l As Single, ByVal t As Single, ByVal w As Single)
    Dim sw As clsSwatch: Set sw = New clsSwatch
    sw.Kind = "MODE": sw.ModeVal = modeVal
    Dim lb As MSForms.Label
    Set lb = NewLabel("mode", l, t, w, 20)
    lb.Caption = cap: lb.TextAlign = fmTextAlignCenter
    Set sw.Lbl = lb
    mCtrls.Add sw
    Select Case modeVal
        Case "FONT": Set mLblFont = lb
        Case "FILL": Set mLblFill = lb
        Case "LINE": Set mLblLine = lb
    End Select
End Sub

Private Sub AddNoneButton(ByVal l As Single, ByVal t As Single)
    Dim sw As clsSwatch: Set sw = New clsSwatch
    sw.Kind = "NONE"
    Dim lb As MSForms.Label
    Set lb = NewLabel("none", l, t, 10 * (SW + GAP) - GAP, SW)
    lb.Caption = "Transparent / No Fill / No Line"
    lb.TextAlign = fmTextAlignCenter
    lb.BackColor = RGB(255, 255, 255)
    Set sw.Lbl = lb
    mCtrls.Add sw
End Sub

Private Sub AddTextButton(ByVal cap As String, ByVal l As Single, _
                          ByVal t As Single, ByVal kindVal As String)
    Dim sw As clsSwatch: Set sw = New clsSwatch
    sw.Kind = kindVal
    Dim lb As MSForms.Label
    Set lb = NewLabel("btn", l, t, 30, 18)
    lb.Caption = cap: lb.TextAlign = fmTextAlignCenter
    lb.BackColor = RGB(220, 220, 220)
    Set sw.Lbl = lb
    mCtrls.Add sw
End Sub

Private Sub BuildCustomRow(ByVal l As Single, ByVal t As Single)
    mCustX = l: mCustY = t
    Dim cap As MSForms.Label
    Set cap = NewLabel("custcap", l, t - 14, 80, 12)
    cap.Caption = "Custom:": cap.BorderStyle = fmBorderStyleNone
    Dim csv As String: csv = ColorPalette.LoadCustom()
    If Len(csv) = 0 Then Exit Sub
    Dim parts() As String: parts = Split(csv, ",")
    Dim i As Long
    For i = 0 To UBound(parts)
        If IsNumeric(parts(i)) Then
            AddSwatch mCustX, mCustY, CLng(parts(i))
            mCustX = mCustX + SW + GAP
        End If
    Next i
End Sub

' ---- public callbacks used by clsSwatch ----------------------------
Public Sub SetMode(ByVal m As String)
    ColorPalette.gMode = m
    CurrentMode = m
    Hilite mLblFont, (m = "FONT")
    Hilite mLblFill, (m = "FILL")
    Hilite mLblLine, (m = "LINE")
End Sub

Private Sub Hilite(ByVal lb As MSForms.Label, ByVal bOn As Boolean)
    If lb Is Nothing Then Exit Sub
    If bOn Then
        lb.BackColor = RGB(0, 32, 96): lb.ForeColor = vbWhite
    Else
        lb.BackColor = RGB(230, 230, 230): lb.ForeColor = vbBlack
    End If
End Sub

Public Sub HexSet()
    Dim c As Long
    If ParseHex(Me.Controls("txtHex").Text, c) Then ColorPalette.ApplyColor c
End Sub

Public Sub HexAdd()
    Dim c As Long
    If Not ParseHex(Me.Controls("txtHex").Text, c) Then Exit Sub
    Dim csv As String: csv = ColorPalette.LoadCustom()
    If Len(csv) > 0 Then csv = csv & ","
    csv = csv & CStr(c)
    ColorPalette.SaveCustom csv
    AddSwatch mCustX, mCustY, c
    mCustX = mCustX + SW + GAP
    ColorPalette.ApplyColor c
End Sub

Private Function ParseHex(ByVal s As String, ByRef outC As Long) As Boolean
    s = Replace(Trim$(s), "#", "")
    If Len(s) <> 6 Then Exit Function
    On Error GoTo bad
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    outC = RGB(r, g, b)
    ParseHex = True
    Exit Function
bad:
End Function
