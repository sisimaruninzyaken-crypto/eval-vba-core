Attribute VB_Name = "modBasicPipeline"
Option Explicit

' Basic生成のオーケストレーター
' 判定はすべてVBAで実行し、AIには文章生成のみを委譲する。
Public Function GenerateBasicPlan(ByVal patientName As String) As Object
    Dim extracted As Object
    Dim normalized As Object
    Dim judged As Object
    Dim planStructure As Object
    Dim aiDraft As Object
    Dim output As Object

    Set extracted = ExtractBasicSourceData(patientName)
    Set normalized = NormalizeBasicSourceData(extracted)
    Set judged = JudgeBasicPlanInputs(normalized)
    Set planStructure = BuildBasicPlanStructureFromJudge(judged)
    Set aiDraft = GenerateBasicPlanNarrative(planStructure)

    Set output = CreateObject("Scripting.Dictionary")
    output("Extract") = extracted
    output("Normalize") = normalized
    output("Judge") = judged
    output("Structure") = planStructure
    output("AIDraft") = aiDraft

    Set GenerateBasicPlan = output
End Function

Public Sub ReflectBasicPlanToReport(ByVal result As Object)
    If result Is Nothing Then Exit Sub
    If Not result.exists("AIDraft") Then Exit Sub

    ' 帳票反映処理は既存の modEvalPrintPack / modEvalReportPrint 側に実装する。
    ' ここではオーケストレーション境界のみを定義。
End Sub
