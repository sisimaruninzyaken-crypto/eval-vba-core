Attribute VB_Name = "Module_PostureOnce"
' ==== 一回きり：姿勢タブの主要コントロール列挙 ====
Public Sub Posture_ListOnce()
    Dim uf As Object: Set uf = frmEval
    Dim mp As Object, pg As Object
    Dim c As Object, y As Object

    On Error Resume Next

    ' どれかのMultiPageを見つける（名前に依存しない）
    For Each c In uf.controls
        If typeName(c) = "MultiPage" Then Set mp = c: Exit For
        If c.controls.count >= 0 Then
            For Each y In c.controls
                If typeName(y) = "MultiPage" Then Set mp = y: Exit For
            Next
            If Not mp Is Nothing Then Exit For
        End If
    Next
    If mp Is Nothing Then Debug.Print "[ERR] MultiPage not found": Exit Sub

    Set pg = mp.Pages(mp.value)
    Debug.Print "=== Controls on current page (type | name | caption) ==="

    ' ページ直下と1階層内側（Frameなど）を列挙
    For Each c In pg.controls
        Debug.Print typeName(c), "|", SafeName1(c), "|", SafeCap1(c)
        If c.controls.count >= 0 Then
            For Each y In c.controls
                Debug.Print "  -", typeName(y), "|", SafeName1(y), "|", SafeCap1(y)
            Next
        End If
    Next
End Sub

Private Function SafeCap1(o As Object) As String
    On Error Resume Next: SafeCap1 = o.caption
End Function

Private Function SafeName1(o As Object) As String
    On Error Resume Next: SafeName1 = o.name
End Function

