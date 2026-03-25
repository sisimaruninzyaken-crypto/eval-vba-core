VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TxtImeHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'=== TxtImeHook.cls : TextBox へフォーカス時にIMEをひらがなへ ===
Option Explicit
Public WithEvents t As MSForms.TextBox
Attribute t.VB_VarHelpID = -1

Public Sub Init(tb As MSForms.TextBox)
    Set t = tb
End Sub

Private Sub T_Enter()
    On Error Resume Next
    If LCase$(t.name) <> "txtbirth" Then
        t.IMEMode = fmIMEModeHiragana
    Else
        t.IMEMode = fmIMEModeOff
    End If
End Sub

