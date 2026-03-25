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
    
    If KeyCode = vbKeyReturn Then
    Dim lb As Object
    On Error Resume Next
    Set lb = frmEval.controls("lstNameSuggest")
    On Error GoTo 0
    
    
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    Dim lb2 As Object
    On Error Resume Next
    Set lb2 = frmEval.controls("lstNameSuggest")
    On Error GoTo 0

    If Not lb2 Is Nothing Then
        If lb2.Visible And lb2.ListCount > 0 Then
            If lb2.ListIndex < 0 Then lb2.ListIndex = 0

            If KeyCode = vbKeyDown Then
                If lb2.ListIndex < lb2.ListCount - 1 Then lb2.ListIndex = lb2.ListIndex + 1
            Else
                If lb2.ListIndex > 0 Then lb2.ListIndex = lb2.ListIndex - 1
            End If
        End If
    End If
    Exit Sub
End If

    
    

    If Not lb Is Nothing Then
        If lb.Visible Then
            If lb.ListCount > 0 Then
                If lb.ListIndex < 0 Then lb.ListIndex = 0

                frmEval.controls("frHeader").controls("txtHdrName").text = lb.List(lb.ListIndex, 0)
                frmEval.controls("txtName").text = frmEval.controls("frHeader").controls("txtHdrName").text
                lb.Visible = False
            End If
        End If
    End If
    Exit Sub
End If
        
    ' ヘッダ → 裏口同期（既存ロジック駆動）
    frmEval.controls("txtName").text = tb.text

    ' 候補BOXを確実に出す
    frmEval.EnsureNameSuggestList
    frmEval.UpdateNameSuggest
    
    
End Sub

