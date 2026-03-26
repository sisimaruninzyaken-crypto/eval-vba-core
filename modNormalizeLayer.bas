Attribute VB_Name = "modNormalizeLayer"
Option Explicit

' 抽出データを判定可能な形に正規化する。
Public Function NormalizeBasicSourceData(ByVal extracted As Object) As Object
    Dim normalized As Object

    Set normalized = CreateObject("Scripting.Dictionary")
    normalized("PatientName") = TrimValue(extracted, "PatientName")
    normalized("CareLevelBand") = NormalizeCareLevel(TrimValue(extracted, "CareLevelRaw"))
    normalized("LivingType") = NormalizeLivingType(TrimValue(extracted, "LivingTypeRaw"))
    normalized("BITotal") = NormalizeNumeric(TrimValue(extracted, "BITotalRaw"), -1)
    normalized("NeedPatient") = TrimValue(extracted, "NeedPatientRaw")
    normalized("NeedFamily") = TrimValue(extracted, "NeedFamilyRaw")
    normalized("MMT_IO") = TrimValue(extracted, "MMT_IO_Raw")
    normalized("TrunkROMLimitTags") = NormalizeTrunkROMLimitTags(TrimValue(extracted, "TrunkROMRaw"))
    normalized("EvalTestNoteRaw") = TrimValue(extracted, "EvalTestNoteRaw")
    normalized("EvalTestCriticalFindings") = ExtractImportantEvalFindings(normalized("EvalTestNoteRaw"))

    Set NormalizeBasicSourceData = normalized
End Function

Private Function NormalizeTrunkROMLimitTags(ByVal trunkRomRaw As String) As String
    Dim parts() As String
    Dim i As Long
    Dim kv() As String
    Dim key As String
    Dim v As Double
    Dim tags As Collection
    Dim arr() As String

    trunkRomRaw = Trim$(trunkRomRaw)
    If LenB(trunkRomRaw) = 0 Then Exit Function

    Set tags = New Collection
    parts = Split(trunkRomRaw, "|")

    For i = LBound(parts) To UBound(parts)
        kv = Split(CStr(parts(i)), "=")
        If UBound(kv) <> 1 Then GoTo ContinueLoop

        key = Trim$(CStr(kv(0)))
        If Not TryParseNumber(CStr(kv(1)), v) Then GoTo ContinueLoop

        Select Case key
            Case "Trunk_Flex": If v <= 40# Then tags.Add key
            Case "Trunk_Ext": If v <= 20# Then tags.Add key
            Case "Trunk_Rot_R", "Trunk_Rot_L": If v <= 30# Then tags.Add key
            Case "Trunk_LatFlex_R", "Trunk_LatFlex_L": If v <= 30# Then tags.Add key
        End Select
ContinueLoop:
    Next i

    If tags.count = 0 Then Exit Function

    ReDim arr(1 To tags.count)
    For i = 1 To tags.count
        arr(i) = CStr(tags(i))
    Next i

    NormalizeTrunkROMLimitTags = Join(arr, ",")
End Function

Private Function TryParseNumber(ByVal value As String, ByRef outVal As Double) As Boolean
    value = Trim$(value)
    If LenB(value) = 0 Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    outVal = CDbl(value)
    TryParseNumber = True
End Function


Private Function TrimValue(ByVal data As Object, ByVal key As String) As String
    If data Is Nothing Then Exit Function
    If Not data.exists(key) Then Exit Function
    TrimValue = Trim$(CStr(data(key)))
End Function

Private Function NormalizeCareLevel(ByVal value As String) As String
    If LenB(value) = 0 Then
        NormalizeCareLevel = "unknown"
    Else
        NormalizeCareLevel = LCase$(value)
    End If
End Function

Private Function NormalizeLivingType(ByVal value As String) As String
    If LenB(value) = 0 Then
        NormalizeLivingType = "unknown"
    Else
        NormalizeLivingType = value
    End If
End Function

Private Function NormalizeNumeric(ByVal value As String, ByVal defaultValue As Long) As Long
    If IsNumeric(value) Then
        NormalizeNumeric = CLng(value)
    Else
        NormalizeNumeric = defaultValue
    End If
End Function
