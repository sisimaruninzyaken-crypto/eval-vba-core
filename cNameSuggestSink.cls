VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNameSuggestSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lb As MSForms.ListBox
Attribute lb.VB_VarHelpID = -1

Public Sub Hook(ByVal x As MSForms.ListBox)
    Set lb = x
End Sub

Private Sub lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lb.ListIndex < 0 Then Exit Sub

    ' Ž–¼‚¾‚¯”½‰fiƒwƒbƒ_{— Œûj
    frmEval.controls("frHeader").controls("txtHdrName").text = lb.List(lb.ListIndex, 0)
    frmEval.controls("txtName").text = frmEval.controls("frHeader").controls("txtHdrName").text

    lb.Visible = False
End Sub

