VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDailyLogList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lb As MSForms.ListBox
Attribute lb.VB_VarHelpID = -1

Public Event DblClicked()

Public Event Clicked(ByVal rowIndex As Long)

Private Function CurrentRowIndex() As Long
    On Error Resume Next
    CurrentRowIndex = lb.listIndex
    On Error GoTo 0
End Function

Private Sub lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClicked
End Sub

Private Sub lb_Click()
    RaiseEvent Clicked(CurrentRowIndex())
End Sub
