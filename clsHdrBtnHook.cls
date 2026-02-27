VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHdrBtnHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public tag As String
Public owner As frmEval

Private Sub btn_Click()
   
    Select Case LCase$(Trim$(tag))
        Case "loadprev"
            If Not owner Is Nothing Then
                owner.HandleHdrLoadPrevClick
            End If
        Case Else
            ArchiveAndDelete_EvalData_ByName
    End Select
End Sub



