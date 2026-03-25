Attribute VB_Name = "modJudgeLayer"
Option Explicit

' 蛻､螳壹・縺吶∋縺ｦVBA縺ｧ螳溯｡後☆繧具ｼ・I蟋碑ｭｲ遖∵ｭ｢・峨・
Public Function JudgeBasicPlanInputs(ByVal normalized As Object) As Object
    Dim judged As Object

    Set judged = CreateObject("Scripting.Dictionary")
    judged("ActivityCandidate") = JudgeActivityCandidate(normalized)
    judged("MainCause") = JudgeMainCause(normalized)
    judged("FunctionCandidate") = JudgeFunctionCandidate(normalized)
    judged("NeedPatient") = GetValue(normalized, "NeedPatient")
    judged("NeedFamily") = GetValue(normalized, "NeedFamily")
    judged("MMT_IO") = GetValue(normalized, "MMT_IO")

    Set JudgeBasicPlanInputs = judged
End Function

Private Function JudgeActivityCandidate(ByVal normalized As Object) As String
    Dim biTotal As Long
    biTotal = ToLong(GetValue(normalized, "BITotal"), -1)

    Select Case biTotal
        Case Is < 40: JudgeActivityCandidate = "襍ｷ螻・ｧｻ蜍・
        Case 40 To 69: JudgeActivityCandidate = "螻句・豁ｩ陦・
        Case Is >= 70: JudgeActivityCandidate = "螻句､匁ｭｩ陦・
        Case Else: JudgeActivityCandidate = "荳肴・"
    End Select
End Function

Private Function JudgeMainCause(ByVal normalized As Object) As String
    If InStr(1, GetValue(normalized, "MMT_IO"), "逍ｼ逞・, vbTextCompare) > 0 Then
        JudgeMainCause = "逍ｼ逞・
    ElseIf InStr(1, GetValue(normalized, "MMT_IO"), "遲句鴨", vbTextCompare) > 0 Then
        JudgeMainCause = "遲句鴨菴惹ｸ・
    Else
        JudgeMainCause = "閠蝉ｹ・ｧ菴惹ｸ・
    End If
End Function

Private Function JudgeFunctionCandidate(ByVal normalized As Object) As String
    Select Case GetValue(normalized, "LivingType")
        Case "迢ｬ螻・: JudgeFunctionCandidate = "遘ｻ荵怜ｮ牙ｮ壽ｧ"
        Case "蜷悟ｱ・: JudgeFunctionCandidate = "豁ｩ陦梧戟荵・ｧ"
        Case Else: JudgeFunctionCandidate = "蝓ｺ譛ｬ蜍穂ｽ・
    End Select
End Function

Private Function GetValue(ByVal data As Object, ByVal key As String) As Variant
    If data Is Nothing Then Exit Function
    If Not data.exists(key) Then Exit Function
    GetValue = data(key)
End Function

Private Function ToLong(ByVal value As Variant, ByVal defaultValue As Long) As Long
    If IsNumeric(value) Then
        ToLong = CLng(value)
    Else
        ToLong = defaultValue
    End If
End Function
