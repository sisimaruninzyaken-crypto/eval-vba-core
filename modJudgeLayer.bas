Attribute VB_Name = "modJudgeLayer"
Option Explicit

' 判定はすべてVBAで実行する（AI委譲禁止）。
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
        Case Is < 40: JudgeActivityCandidate = "起居移動"
        Case 40 To 69: JudgeActivityCandidate = "屋内歩行"
        Case Is >= 70: JudgeActivityCandidate = "屋外歩行"
        Case Else: JudgeActivityCandidate = "不明"
    End Select
End Function

Private Function JudgeMainCause(ByVal normalized As Object) As String
    If InStr(1, GetValue(normalized, "MMT_IO"), "疼痛", vbTextCompare) > 0 Then
        JudgeMainCause = "疼痛"
    ElseIf InStr(1, GetValue(normalized, "MMT_IO"), "筋力", vbTextCompare) > 0 Then
        JudgeMainCause = "筋力低下"
    Else
        JudgeMainCause = "耐久性低下"
    End If
End Function

Private Function JudgeFunctionCandidate(ByVal normalized As Object) As String
    Select Case GetValue(normalized, "LivingType")
        Case "独居": JudgeFunctionCandidate = "移乗安定性"
        Case "同居": JudgeFunctionCandidate = "歩行持久性"
        Case Else: JudgeFunctionCandidate = "基本動作"
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
