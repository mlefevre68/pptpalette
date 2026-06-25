VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Wraps one dynamically-created label and routes its Click event.
Public WithEvents Lbl As MSForms.Label
Attribute Lbl.VB_VarHelpID = -1
Public Kind As String          ' "COLOR" | "NONE" | "MODE" | "HEXSET" | "HEXADD"
Public ColorVal As Long
Public ModeVal As String

Private Sub Lbl_Click()
    Select Case Kind
        Case "COLOR":  ColorPalette.ApplyColor ColorVal
        Case "NONE":   ColorPalette.ApplyNoColor
        Case "MODE":   frmPalette.SetMode ModeVal
        Case "HEXSET": frmPalette.HexSet
        Case "HEXADD": frmPalette.HexAdd
    End Select
End Sub
