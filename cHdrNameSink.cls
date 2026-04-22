VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHdrNameSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1

Public Sub Hook(ByVal t As MSForms.TextBox)
    Set tb = t
End Sub
Private Sub tb_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim lb As Object

    On Error Resume Next
    Set lb = frmEval.controls("lstNameSuggest")
    On Error GoTo 0

    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If Not lb Is Nothing Then
            If lb.Visible And lb.ListCount > 0 Then
                If lb.listIndex < 0 Then lb.listIndex = 0

                If KeyCode = vbKeyDown Then
                    If lb.listIndex < lb.ListCount - 1 Then lb.listIndex = lb.listIndex + 1
                Else
                    If lb.listIndex > 0 Then lb.listIndex = lb.listIndex - 1
                End If
            End If
        End If
        Exit Sub
    End If

    If KeyCode = vbKeyReturn Then
        If Not lb Is Nothing Then
            If lb.Visible And lb.ListCount > 0 Then
                If lb.listIndex < 0 Then lb.listIndex = 0
                frmEval.CommitNameSuggestSelection lb.listIndex
            End If
        End If
        Exit Sub
    End If

    frmEval.controls("txtName").text = tb.text
    frmEval.EnsureNameSuggestList
    frmEval.UpdateNameSuggest
End Sub
