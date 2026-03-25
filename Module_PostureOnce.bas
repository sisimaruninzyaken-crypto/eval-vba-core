Attribute VB_Name = "Module_PostureOnce"
' ==== 荳蝗槭″繧奇ｼ壼ｧｿ蜍｢繧ｿ繝悶・荳ｻ隕√さ繝ｳ繝医Ο繝ｼ繝ｫ蛻玲嫌 ====
Public Sub Posture_ListOnce()
    Dim uf As Object: Set uf = frmEval
    Dim mp As Object, pg As Object
    Dim c As Object, y As Object

    On Error Resume Next

    ' 縺ｩ繧後°縺ｮMultiPage繧定ｦ九▽縺代ｋ・亥錐蜑阪↓萓晏ｭ倥＠縺ｪ縺・ｼ・
    For Each c In uf.controls
        If TypeName(c) = "MultiPage" Then Set mp = c: Exit For
        If c.controls.count >= 0 Then
            For Each y In c.controls
                If TypeName(y) = "MultiPage" Then Set mp = y: Exit For
            Next
            If Not mp Is Nothing Then Exit For
        End If
    Next
    If mp Is Nothing Then Debug.Print "[ERR] MultiPage not found": Exit Sub

    Set pg = mp.Pages(mp.value)
    Debug.Print "=== Controls on current page (type | name | caption) ==="

    ' 繝壹・繧ｸ逶ｴ荳九→1髫主ｱ､蜀・・・・rame縺ｪ縺ｩ・峨ｒ蛻玲嫌
    For Each c In pg.controls
        Debug.Print TypeName(c), "|", SafeName1(c), "|", SafeCap1(c)
        If c.controls.count >= 0 Then
            For Each y In c.controls
                Debug.Print "  -", TypeName(y), "|", SafeName1(y), "|", SafeCap1(y)
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

