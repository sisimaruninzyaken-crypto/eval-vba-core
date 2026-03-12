VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBasicInfoEnterComboHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents cb As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1

Private mOwner As frmEval

Public Sub Init(ByVal owner As frmEval, ByVal target As MSForms.ComboBox)
    Set mOwner = owner
    Set cb = target
End Sub

Private Sub cb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mOwner Is Nothing Then Exit Sub

    KeyCode = 0
    mOwner.MoveBasicInfoFocusNext cb
End Sub
