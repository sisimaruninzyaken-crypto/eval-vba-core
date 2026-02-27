Attribute VB_Name = "ModUtil"
Option Explicit

Public Const TRACE_ON As Boolean = True

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
    Dim c As Object, pg As MSForms.Page
    On Error Resume Next

    For Each c In container.Controls
        If StrComp(c.name, ctlName, vbTextCompare) = 0 Then
            Set FindCtlDeep = c
            Exit Function
        End If

        If TypeOf c Is MSForms.Frame Or TypeOf c Is MSForms.Page Then
            Set FindCtlDeep = FindCtlDeep(c, ctlName)   ' ← ctlName に統一
            If Not FindCtlDeep Is Nothing Then Exit Function
        End If

        If TypeOf c Is MSForms.MultiPage Then
            For Each pg In c.Pages
                Set FindCtlDeep = FindCtlDeep(pg, ctlName)  ' ← ctlName に統一
                If Not FindCtlDeep Is Nothing Then Exit Function
            Next
        End If
    Next
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
