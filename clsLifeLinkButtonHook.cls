VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLifeLinkButtonHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private mOwner As frmLifeLink
Private mActionName As String

Public Sub Init(ByVal owner As frmLifeLink, ByVal target As MSForms.CommandButton, ByVal actionName As String)
    Set mOwner = owner
    Set btn = target
    mActionName = actionName
End Sub

Private Sub Btn_Click()
    If mOwner Is Nothing Then Exit Sub
    mOwner.HandleButtonClick mActionName
End Sub
