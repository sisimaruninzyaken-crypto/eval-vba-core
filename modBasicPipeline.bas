Attribute VB_Name = "modBasicPipeline"
Option Explicit

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
    Set output("Extract") = extracted
    Set output("Normalize") = normalized
    Set output("Judge") = judged
    Set output("Structure") = planStructure
    Set output("AIDraft") = aiDraft

    Set GenerateBasicPlan = output
End Function

Public Sub RunBasicPlan()
    Dim frm As Object
    Dim patientName As String
    Dim result As Object
    Dim prevSnap As Object
    Dim changeIssue As Object

    Set frm = TryGetOwnerForm()
    If Not frm Is Nothing Then
        On Error Resume Next
        patientName = Trim$(CStr(frm.Controls("txtName").value))
        Err.Clear
        On Error GoTo 0
    End If
    Set result = GenerateBasicPlan(patientName)
    Debug.Print "[RunBasicPlan] patientName=[" & patientName & "]"

    Set prevSnap = GetPreviousEvalSnapshot(frm)
    If Not prevSnap Is Nothing Then
        Set changeIssue = GenerateChangeAndIssue(result("Structure"), prevSnap)
        If Not changeIssue Is Nothing Then Set result("ChangeIssue") = changeIssue
    End If

    ReflectBasicPlanToReport result, patientName, frm
End Sub

Public Sub ReflectBasicPlanToReport(ByVal result As Object, ByVal patientName As String, Optional ByVal owner As Object = Nothing)
    Dim planData As Object
    If result Is Nothing Then Debug.Print "[Reflect] result=Nothing": Exit Sub
    If Not result.exists("AIDraft") Then Debug.Print "[Reflect] AIDraft key missing": Exit Sub
    If owner Is Nothing Then Set owner = TryGetOwnerForm()
    Set planData = BuildPlanDataFromResult(result)
    If planData Is Nothing Then
        Debug.Print "[Reflect] planDataCount=-1"
    Else
        Debug.Print "[Reflect] planDataCount=" & planData.count
    End If
    ExportPlanAsXlsx patientName, owner, planData
End Sub

Private Function TryGetOwnerForm() As Object
    Dim i As Integer
    On Error Resume Next
    For i = 0 To VBA.UserForms.count - 1
        If StrComp(VBA.UserForms(i).name, "frmEval", vbTextCompare) = 0 Then
            Set TryGetOwnerForm = VBA.UserForms(i)
            Exit Function
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Function

Private Sub ExportPlanAsXlsx(ByVal patientName As String, ByVal owner As Object, ByVal planData As Object)
    Dim fso As Object
    Dim baseDir As String
    Dim SafeName As String
    Dim outputDir As String
    Dim tmpl As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim dateStr As String
    Dim fileName As String
    Dim savePath As String
    Const TEMPLATE_NAME As String = "�ʋ@�\�P���v�揑""
    On Error GoTo EH

    Set fso = CreateObject("Scripting.FileSystemObject")
    baseDir = ThisWorkbook.path & "\KojinPlan"
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir

    SafeName = patientName
    SafeName = Replace(SafeName, "/", "")
    SafeName = Replace(SafeName, "\", "")
    SafeName = Replace(SafeName, "[", "")
    SafeName = Replace(SafeName, "]", "")
    SafeName = Replace(SafeName, "*", "")
    SafeName = Replace(SafeName, "?", "")
    SafeName = Replace(SafeName, ":", "")
    If LenB(Trim$(SafeName)) = 0 Then SafeName = "kanja"

    outputDir = baseDir & "\" & SafeName
    If Not fso.FolderExists(outputDir) Then fso.CreateFolder outputDir
    
    
    Debug.Print "[WB] name=" & ThisWorkbook.name & " | sheets=" & ThisWorkbook.Worksheets.count & " | hasTemplate=" & CStr(Not ThisWorkbook.Worksheets("�ʋ@�\�P���v�揑") Is Nothing)

    Set tmpl = ThisWorkbook.Worksheets("�ʋ@�\�P���v�揑")
    If tmpl Is Nothing Then
        Debug.Print "[ExportPlan] template not found"
        Exit Sub
    End If
    Debug.Print "[ExportPlan] template=" & tmpl.name

    tmpl.Copy
    Set newWb = ActiveWorkbook
    Set newWs = newWb.Worksheets(1)

    modEvalPlanSheetOutput.WriteEvalPlanSheet newWs, owner, planData

    dateStr = Format$(Now(), "YYYYMMDD")
    fileName = SafeName & "_" & dateStr
    savePath = outputDir & "\" & fileName & ".xlsx"

    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False

    MsgBox "saved: " & savePath, vbInformation, "done"
    Exit Sub

EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "error"
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
End Sub


Private Function ResolvePlanTemplateSheet(ByVal wb As Workbook, ByVal primaryName As String) As Worksheet
    Dim ws As Worksheet
    Dim candidates As Variant
    Dim i As Long

    If wb Is Nothing Then Exit Function

    candidates = Array(primaryName, "個別機能訓練計画書", "kojinkinokunren")

    On Error Resume Next
    For i = LBound(candidates) To UBound(candidates)
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(candidates(i)))
        If Not ws Is Nothing Then
            Set ResolvePlanTemplateSheet = ws
            On Error GoTo 0
            Exit Function
        End If
    Next i
    On Error GoTo 0

    For Each ws In wb.Worksheets
        If InStr(1, ws.name, "個別機能訓練計画書", vbTextCompare) > 0 Then
            Set ResolvePlanTemplateSheet = ws
            Exit Function
        End If
        If InStr(1, ws.name, "計画書", vbTextCompare) > 0 Then
            Set ResolvePlanTemplateSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function BuildPlanDataFromResult(ByVal result As Object) As Object
    Dim d As Object
    Dim structure As Object
    Dim k As Variant
    Dim aiDraft As Object
    Dim pi As Long
    Dim pKey As String
    Dim goalKey As Variant
    Dim ci As Object

    Set d = CreateObject("Scripting.Dictionary")

    If result.exists("Structure") Then
        Set structure = result("Structure")
        If Not structure Is Nothing Then
            For Each k In Array("Activity_Long", "Activity_Short", "Function_Long", "Function_Short", "Participation_Long", "Participation_Short", "MainCause")
                If structure.exists(CStr(k)) Then d(CStr(k)) = structure(CStr(k))
            Next k
        End If
    End If

    If result.exists("AIDraft") Then
        Set aiDraft = result("AIDraft")
        If Not aiDraft Is Nothing Then
            If aiDraft.exists("MonitoringText") Then d("Monitoring.Change") = aiDraft("MonitoringText")
            If aiDraft.exists("HomeExercise") Then d("HomeExercise") = aiDraft("HomeExercise")
            For pi = 1 To 5
                pKey = "Program" & pi & "Content"
                If aiDraft.exists(pKey) Then
                    If Len(Trim$(CStr(aiDraft(pKey)))) > 0 Then d(pKey) = aiDraft(pKey)
                End If
            Next pi
            For Each goalKey In Array("Function_Long", "Function_Short", "Activity_Long", "Activity_Short", "Participation_Long", "Participation_Short")
                If aiDraft.exists(CStr(goalKey)) Then
                    If Len(Trim$(CStr(aiDraft(CStr(goalKey))))) > 0 Then d(CStr(goalKey)) = aiDraft(CStr(goalKey))
                End If
            Next goalKey
        End If
    End If

    If result.exists("ChangeIssue") Then
        Set ci = result("ChangeIssue")
        If Not ci Is Nothing Then
            If ci.exists("Change") Then d("Monitoring.Change") = ci("Change")
            If ci.exists("Issue") Then d("Monitoring.Issue") = ci("Issue")
        End If
    End If

    Set BuildPlanDataFromResult = d
End Function

Private Function GetPreviousEvalSnapshot(ByVal owner As Object) As Object
    Dim ws As Worksheet
    Dim firstD As String
    Dim latestD As String
    Dim prevD As String
    Dim recCnt As Long
    Dim latestRow As Long
    Dim snap As Object
    Dim col As Long
    Dim v As String
    On Error GoTo EH

    If Not modEvalIOEntry.TryGetUserHistorySheet(owner, ws) Then Exit Function
    modEvalIOEntry.GetUserEvalDateStats ws, firstD, latestD, prevD, recCnt
    If LenB(Trim$(latestD)) = 0 Then Exit Function

    latestRow = FindLatestEvalRow(ws)
    If latestRow = 0 Then Exit Function

    Set snap = CreateObject("Scripting.Dictionary")
    snap("EvalDate") = latestD

    col = FindSheetColByHeader(ws, "BITotal")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("BITotal") = v
    col = FindSheetColByHeader(ws, "Test_TUG_sec")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("Test_TUG_sec") = v
    col = FindSheetColByHeader(ws, "Test_10MWalk_sec")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("Test_10MWalk_sec") = v
    col = FindSheetColByHeader(ws, "Test_Grip_R_kg")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("Test_Grip_R_kg") = v
    col = FindSheetColByHeader(ws, "Test_Grip_L_kg")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("Test_Grip_L_kg") = v
    col = FindSheetColByHeader(ws, "Test_5xSitStand_sec")
    If col > 0 Then v = Trim$(CStr(ws.Cells(latestRow, col).value)): If LenB(v) > 0 Then snap("Test_5xSitStand_sec") = v

    Set GetPreviousEvalSnapshot = snap
    Exit Function
EH:
    Err.Clear
End Function

Private Function FindLatestEvalRow(ByVal ws As Worksheet) As Long
    Dim colEvalDate As Long
    Dim r As Long
    Dim latestRow As Long
    Dim latestDate As Date
    Dim hasDate As Boolean
    Dim d As Date
    On Error GoTo EH

    colEvalDate = FindSheetColByHeader(ws, "Basic.EvalDate")
    If colEvalDate = 0 Then Exit Function
    For r = 2 To ws.UsedRange.rows.count
        If IsDate(ws.Cells(r, colEvalDate).value) Then
            d = DateValue(CDate(ws.Cells(r, colEvalDate).value))
            If Not hasDate Or d > latestDate Then
                latestDate = d
                latestRow = r
                hasDate = True
            End If
        End If
    Next r
    FindLatestEvalRow = latestRow
    Exit Function
EH:
    Err.Clear
End Function

Private Function FindSheetColByHeader(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.count
        If Trim$(CStr(ws.Cells(1, c).value)) = header Then
            FindSheetColByHeader = c
            Exit Function
        End If
    Next c
End Function
