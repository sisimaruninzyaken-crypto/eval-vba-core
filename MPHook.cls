VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MPHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== MPHooK.cls : MultiPage のタブ切り替え時にIME再適用 =========
Option Explicit
Public WithEvents mp As MSForms.MultiPage
Attribute mp.VB_VarHelpID = -1
Private owner As frmEval

Public Sub Init(o As frmEval, target As MSForms.MultiPage)
    Set owner = o
    Set mp = target
End Sub

Private Sub MP_Change()
    On Error Resume Next
    owner.ApplyImeToIADLNote
End Sub

