Attribute VB_Name = "modBasicPipeline"
Option Explicit

' Basic逕滓・縺ｮ繧ｪ繝ｼ繧ｱ繧ｹ繝医Ξ繝ｼ繧ｿ繝ｼ
' 蛻､螳壹・縺吶∋縺ｦVBA縺ｧ螳溯｡後＠縲、I縺ｫ縺ｯ譁・ｫ逕滓・縺ｮ縺ｿ繧貞ｧ碑ｭｲ縺吶ｋ縲・
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

    ' 蟶ｳ逾ｨ蜿肴丐蜃ｦ逅・・譌｢蟄倥・ modEvalPrintPack / modEvalReportPrint 蛛ｴ縺ｫ螳溯｣・☆繧九・
    ' 縺薙％縺ｧ縺ｯ繧ｪ繝ｼ繧ｱ繧ｹ繝医Ξ繝ｼ繧ｷ繝ｧ繝ｳ蠅・阜縺ｮ縺ｿ繧貞ｮ夂ｾｩ縲・
End Sub
