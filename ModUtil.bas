Attribute VB_Name = "ModUtil"
Option Explicit

Public Const TRACE_ON As Boolean = False

Public Sub Trace(ByVal msg As String, Optional ByVal tag As String = "")
    If Not TRACE_ON Then Exit Sub
    If Len(tag) > 0 Then
        Debug.Print Format(Now, "hh:nn:ss"), "[" & tag & "]", msg
    Else
        Debug.Print Format(Now, "hh:nn:ss"), msg
    End If
End Sub

' ===== Deep control search & helpers =====

Public Function FindCtlDeep(ByVal container As Object, ByVal ctlName As String) As MSForms.Control
    Dim hit As Object
   
    Set hit = modCommonUtil.SafeGetControl(container, ctlName)
    If Not hit Is Nothing Then
        Set FindCtlDeep = hit
    End If

End Function

Public Function FindCtlByTagDeep(ByVal container As Object, ByVal targetTag As String) As MSForms.Control
    Dim c As Object
    Dim hit As MSForms.Control
    Dim pg As MSForms.page

    On Error Resume Next

    If Not container Is Nothing Then
        If StrComp(CStr(container.tag), targetTag, vbTextCompare) = 0 Then
            Set FindCtlByTagDeep = container
            Exit Function
        End If
    End If

    For Each c In container.controls
        If StrComp(CStr(c.tag), targetTag, vbTextCompare) = 0 Then
            Set FindCtlByTagDeep = c
            Exit Function
        End If

        Set hit = FindCtlByTagDeep(c, targetTag)
        If Not hit Is Nothing Then
            Set FindCtlByTagDeep = hit
            Exit Function
        End If
    Next

    If TypeName(container) = "MultiPage" Then
        For Each pg In container.pages
            Set hit = FindCtlByTagDeep(pg, targetTag)
            If Not hit Is Nothing Then
                Set FindCtlByTagDeep = hit
                Exit Function
            End If
        Next
    End If
    On Error GoTo 0
End Function

Public Function GetCtlText(ByVal owner As Object, ByVal ctlName As String) As String
    Dim ctl As MSForms.Control
    Set ctl = FindCtlDeep(owner, ctlName)
    If Not ctl Is Nothing Then On Error Resume Next: GetCtlText = ctl.value
End Function

Public Function GetCtlCheck(ByVal owner As Object, ByVal ctlName As String) As String
    Dim ctl As MSForms.Control
    Set ctl = FindCtlDeep(owner, ctlName)
    If Not ctl Is Nothing Then On Error Resume Next: GetCtlCheck = IIf(ctl.value = True, "有", "無")  ' ← 半角ダブルクォーテーション
End Function
