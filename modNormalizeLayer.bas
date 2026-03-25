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

    Set NormalizeBasicSourceData = normalized
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
