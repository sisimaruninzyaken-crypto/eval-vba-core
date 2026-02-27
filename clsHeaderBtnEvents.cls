VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHeaderBtnEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public tag As String

Private Sub btn_Click()
#If APP_DEBUG Then
    Debug.Print "[HeaderBtn] Click Tag=" & tag & " Name=" & btn.name
#End If

    Select Case tag
        Case "Save"
            frmEval.DoSaveGlobal
        Case "Clear"
            frmEval.DoClearGlobal
        Case "Close"
            frmEval.DoCloseForm
    End Select
End Sub


