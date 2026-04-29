Attribute VB_Name = "modLifeCsvWriter"
Option Explicit

Public Sub LifeCsv_WriteUtf8BomLines(ByVal filePath As String, ByVal lines As Variant)
    Dim textBody As String
    textBody = Join(lines, vbCrLf)
    LifeCsv_WriteUtf8BomText filePath, textBody
End Sub

Public Sub LifeCsv_WriteUtf8BomText(ByVal filePath As String, ByVal textBody As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText textBody
    stm.SaveToFile filePath, 2
    stm.Close
End Sub

Public Function LifeCsv_JoinRow(ByVal values As Variant) As String
    Dim i As Long
    Dim parts() As String

    ReDim parts(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        parts(i) = LifeCsv_Escape(CStr(values(i)))
    Next i

    LifeCsv_JoinRow = Join(parts, ",")
End Function

Public Function LifeCsv_Escape(ByVal valueText As String) As String
    Dim escaped As String
    escaped = Replace$(valueText, """", """""")

    If InStr(escaped, ",") > 0 _
       Or InStr(escaped, vbCr) > 0 _
       Or InStr(escaped, vbLf) > 0 _
       Or InStr(escaped, """") > 0 Then
        LifeCsv_Escape = """" & escaped & """"
    Else
        LifeCsv_Escape = escaped
    End If
End Function

Public Sub LifeCsv_EnsureFolderExists(ByVal folderPath As String)
    If LenB(Trim$(folderPath)) = 0 Then Exit Sub
    If LenB(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

