Attribute VB_Name = "modUiQuickInspect"
Public Sub ListFrmEvalControls()
    Dim c As Control
    For Each c In frmEval.Controls
        Debug.Print TypeName(c), c.name, "Top=" & c.top, "Left=" & c.Left
    Next
End Sub

