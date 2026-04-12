VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDailyDateHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1

Private mHost As Object

Public Sub Hook(ByVal target As MSForms.TextBox, ByVal host As Object)
    Set tb = target
    Set mHost = host
End Sub

Private Sub tb_Change()
    NotifyHost
End Sub

Private Sub tb_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    NotifyHost
End Sub

Private Sub NotifyHost()
    On Error Resume Next
    CallByName mHost, "OnDailyDateChanged", VbMethod
    On Error GoTo 0
End Sub

