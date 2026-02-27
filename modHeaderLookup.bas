Attribute VB_Name = "modHeaderLookup"
'=== ƒwƒbƒ_ŒŸõEì¬ ==================
Public Function BuildHeaderLookup(ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim key As String: key = Trim$(CStr(ws.Cells(1, c).value))
        If Len(key) > 0 Then dict(key) = c
    Next
    Set BuildHeaderLookup = dict
End Function

Public Function ResolveColumn(look As Object, key As String) As Long
    If look.exists(key) Then
        ResolveColumn = look(key)
    Else
        ResolveColumn = 0
    End If
End Function

Public Function EnsureHeaderColumn(ws As Worksheet, look As Object, key As String) As Long
    Dim c As Long: c = ResolveColumn(look, key)
    If c = 0 Then
        c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, c).value = key
        look(key) = c
    End If
    EnsureHeaderColumn = c
End Function

Public Function ResolveColOrCreate(ws As Worksheet, look As Object, ParamArray keys()) As Long
    Dim i As Long, c As Long
    For i = LBound(keys) To UBound(keys)
        c = ResolveColumn(look, CStr(keys(i)))
        If c > 0 Then ResolveColOrCreate = c: Exit Function
    Next
    ResolveColOrCreate = EnsureHeaderColumn(ws, look, CStr(keys(LBound(keys))))
End Function

Public Function EnsureEvalData() As Worksheet
    On Error Resume Next
    Set EnsureEvalData = ThisWorkbook.Sheets("EvalData")
    On Error GoTo 0
    If EnsureEvalData Is Nothing Then
        Set EnsureEvalData = ThisWorkbook.Sheets.Add
        EnsureEvalData.name = "EvalData"
    End If
End Function

