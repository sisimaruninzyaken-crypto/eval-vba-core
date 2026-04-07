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
    data("TrunkROMRaw") = ReadTrunkROMRaw(patientName)
    data("EvalTestNoteRaw") = ReadLatestEvalTextByHeader(patientName, "TestEval_Note")
    data("InterestNowRaw") = ReadLatestEvalTextByHeader(patientName, "Interest_Now")
    data("InterestPastRaw") = ReadLatestEvalTextByHeader(patientName, "Interest_Past")
    data("InterestWantRaw") = ReadLatestEvalTextByHeader(patientName, "Interest_Want")
    data("InterestSocialRaw") = ReadLatestEvalTextByHeader(patientName, "Interest_Social")

    Set ExtractBasicSourceData = data
End Function


Private Function ReadTrunkROMRaw(ByVal patientName As String) As String
    Dim keys As Variant
    Dim labels As Variant
    Dim i As Long
    Dim v As String
    Dim chunks As Collection
    Dim arr() As String

    keys = Array("ROM_Trunk_Flex", "ROM_Trunk_Ext", "ROM_Trunk_Rot_R", "ROM_Trunk_Rot_L", "ROM_Trunk_LatFlex_R", "ROM_Trunk_LatFlex_L")
    labels = Array("Trunk_Flex", "Trunk_Ext", "Trunk_Rot_R", "Trunk_Rot_L", "Trunk_LatFlex_R", "Trunk_LatFlex_L")

    Set chunks = New Collection
    For i = LBound(keys) To UBound(keys)
        v = ReadLatestEvalTextByHeader(patientName, CStr(keys(i)))
        If LenB(v) > 0 Then chunks.Add CStr(labels(i)) & "=" & v
    Next i

    If chunks.count = 0 Then Exit Function

    ReDim arr(1 To chunks.count)
    For i = 1 To chunks.count
        arr(i) = CStr(chunks(i))
    Next i

    ReadTrunkROMRaw = Join(arr, "|")
End Function


Private Function ReadFrmEvalControlText(ByVal controlName As String) As String
    On Error GoTo EH

    Dim frm As Object

    If VBA.UserForms.count > 0 Then
        For Each frm In VBA.UserForms
            If TypeName(frm) = "frmEval" Then
                ReadFrmEvalControlText = Trim$(CStr(frm.controls(controlName).value))
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
