Attribute VB_Name = "Module_PostureOnce"
' ==== 一回きり：姿勢タブの主要コントロール列挙 ====
Public Sub Posture_ListOnce()
    Dim uf As Object: Set uf = frmEval
    Dim mp As Object, pg As Object
    Dim c As Object, Y As Object

    On Error Resume Next

    ' どれかのMultiPageを見つける（名前に依存しない）
    For Each c In uf.Controls
        If TypeName(c) = "MultiPage" Then Set mp = c: Exit For
        If c.Controls.Count >= 0 Then
            For Each Y In c.Controls
                If TypeName(Y) = "MultiPage" Then Set mp = Y: Exit For
            Next
            If Not mp Is Nothing Then Exit For
        End If
    Next
    If mp Is Nothing Then Debug.Print "[ERR] MultiPage not found": Exit Sub

    Set pg = mp.Pages(mp.value)
    Debug.Print "=== Controls on current page (type | name | caption) ==="

    ' ページ直下と1階層内側（Frameなど）を列挙
    For Each c In pg.Controls
        Debug.Print TypeName(c), "|", SafeName1(c), "|", SafeCap1(c)
        If c.Controls.Count >= 0 Then
            For Each Y In c.Controls
                Debug.Print "  -", TypeName(Y), "|", SafeName1(Y), "|", SafeCap1(Y)
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

