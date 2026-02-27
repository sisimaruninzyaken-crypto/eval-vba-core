VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CboBIHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== CboBIHook.cls: BI項目コンボのChangeを親フォームへ通知 ===
Option Explicit
Public WithEvents c As MSForms.ComboBox
Attribute c.VB_VarHelpID = -1
Private pForm As frmEval

Public Sub Init(owner As frmEval, target As MSForms.ComboBox)
    Set pForm = owner
    Set c = target
End Sub

Private Sub C_Change()
    On Error Resume Next
    pForm.RecalcBI
End Sub

Private Sub C_Click()
    On Error Resume Next
    pForm.RecalcBI
End Sub

