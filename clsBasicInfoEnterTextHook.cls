VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBasicInfoEnterTextHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Private mOwner As frmEval

Public Sub Init(ByVal owner As frmEval, ByVal target As MSForms.TextBox)
    Set mOwner = owner
    Set tb = target
End Sub

Private Sub tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mOwner Is Nothing Then Exit Sub

    KeyCode = 0
    mOwner.MoveBasicInfoFocusNext tb
End Sub

