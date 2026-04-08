VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFacilityButtonHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1


Private mOwner As frmFacilitySettings
Private mActionName As String

Public Sub Init(ByVal owner As frmFacilitySettings, ByVal target As MSForms.CommandButton, ByVal actionName As String)
    Set mOwner = owner
    Set Btn = target
    mActionName = actionName
End Sub

Private Sub Btn_Click()
    If mOwner Is Nothing Then Exit Sub
    mOwner.HandleButtonClick mActionName
End Sub

