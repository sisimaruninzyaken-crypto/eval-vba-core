Attribute VB_Name = "modExtractLayer"
Option Explicit

' frmEval と EvalData から Basic生成で使う最小データのみを抽出する。
Public Function ExtractBasicSourceData(ByVal patientName As String) As Object
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    data("PatientName") = patientName
    data("CareLevelRaw") = ReadFrmEvalControlText("cboCare")
    data("LivingTypeRaw") = ReadFrmEvalControlText("txtLiving")
    data("BITotalRaw") = ReadFrmEvalControlText("txtBITotal")
    data("NeedPatientRaw") = ReadLatestEvalTextByHeader(patientName, "本人Needs")
    data("NeedFamilyRaw") = ReadLatestEvalTextByHeader(patientName, "家族Needs")
    data("MMT_IO_Raw") = ReadLatestEvalTextByHeader(patientName, "MMT_IO")

    Set ExtractBasicSourceData = data
End Function

Private Function ReadFrmEvalControlText(ByVal controlName As String) As String
    On Error GoTo EH

    Dim frm As Object

    If VBA.UserForms.count > 0 Then
        For Each frm In VBA.UserForms
            If TypeName(frm) = "frmEval" Then
                ReadFrmEvalControlText = Trim$(CStr(frm.Controls(controlName).value))
                Exit Function
            End If
        Next frm
    End If

    ReadFrmEvalControlText = vbNullString
    Exit Function
EH:
    ReadFrmEvalControlText = vbNullString
End Function

Private Function ReadLatestEvalTextByHeader(ByVal patientName As String, ByVal headerName As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Dim latestRow As Long

    If LenB(Trim$(patientName)) = 0 Then Exit Function

    Set ws = ThisWorkbook.Worksheets("EvalData")
    latestRow = FindLatestRowByName(ws, patientName)
    If latestRow <= 0 Then Exit Function

    ReadLatestEvalTextByHeader = Trim$(ReadStr_Compat(headerName, latestRow, ws))
    Exit Function
EH:
    ReadLatestEvalTextByHeader = vbNullString
End Function
