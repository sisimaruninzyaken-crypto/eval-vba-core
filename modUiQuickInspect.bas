Attribute VB_Name = "modUiQuickInspect"
Public Sub ListFrmEvalControls()
    Dim c As Control
    For Each c In frmEval.controls
        Debug.Print typeName(c), c.name, "Top=" & c.top, "Left=" & c.Left
    Next
End Sub

