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

Private Sub lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClicked
End Sub

