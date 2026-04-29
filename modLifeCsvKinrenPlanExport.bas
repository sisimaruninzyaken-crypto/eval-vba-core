Attribute VB_Name = "modLifeCsvKinrenPlanExport"
Option Explicit

Private Const LIFE_KIND As String = "INDIVIDUAL_FUNCTION_TRAINING_PLAN_2024"

Private Function LifePlanHeader() As String
    LifePlanHeader = _
        "care_facility_id,service_code,insurer_no,insured_no,external_system_management_number,trinity_attempt,evaluate_date,last_date,first_date,care_level," & _
        "impaired_elderly_independence_degree,dementia_elderly_independence_degree,user_request,user_family_request,social_participation,housing_environment,disease_name_code," & _
        "onset_date_year,onset_date_month,onset_date_day,latest_admission_date_year,latest_admission_date_month,latest_admission_date_day,latest_discharge_date_year,latest_discharge_date_month,latest_discharge_date_day,progress," & _
        "complications_control_cerebrovascular_disease,complications_control_fracture,complications_control_pneumonia,complications_control_congestive_heart_failure,complications_control_urinary_tract_infection,complications_control_diabetes," & _
        "complications_control_hypertension,complications_control_osteoporosis,complications_control_articular_rheumatism,complications_control_cancer,complications_control_depression_state,complications_control_dementia," & _
        "complications_control_decubitus,complications_control_nervous_disease,complications_control_motor_disorder,complications_control_respiratory_disease,complications_control_circulatory_disease,complications_control_digestive_system_disease," & _
        "complications_control_kidney_disease,complications_control_endocrine_disease,complications_control_skin_disease,complications_control_neurological_disease,complications_control_other,implementation_status," & _
        "short_goal_functional_training_mind_and_body_function1,short_goal_functional_training_mind_and_body_function2,short_goal_functional_training_mind_and_body_function3,short_goal_functional_training_mind_and_body_function_contents," & _
        "short_goal_functional_training_activity1,short_goal_functional_training_activity2,short_goal_functional_training_activity3,short_goal_functional_training_activity_contents," & _
        "short_goal_functional_training_activity_join1,short_goal_functional_training_activity_join2,short_goal_functional_training_activity_join3,short_goal_functional_training_activity_join_contents,short_goal_achievement," & _
        "long_goal_functional_training_mind_and_body_function1,long_goal_functional_training_mind_and_body_function2,long_goal_functional_training_mind_and_body_function3,long_goal_functional_training_mind_and_body_function_contents," & _
        "long_goal_functional_training_activity1,long_goal_functional_training_activity2,long_goal_functional_training_activity3,long_goal_functional_training_activity_contents," & _
        "long_goal_functional_training_activity_join1,long_goal_functional_training_activity_join2,long_goal_functional_training_activity_join3,long_goal_functional_training_activity_join_contents,long_goal_achievement," & _
        "function_training_content_01,function_training_note_01,function_training_frequency_times_01,function_training_date_01,function_training_personnel_01," & _
        "function_training_content_02,function_training_note_02,function_training_frequency_times_02,function_training_date_02,function_training_personnel_02," & _
        "function_training_content_03,function_training_note_03,function_training_frequency_times_03,function_training_date_03,function_training_personnel_03," & _
        "function_training_content_04,function_training_note_04,function_training_frequency_times_04,function_training_date_04,function_training_personnel_04," & _
        "program_planner,himself_or_family_routine_work,functional_training_notice,functional_training_summary,subject_and_factors,version"
End Function

Public Sub ExportLifeKinrenPlanCsvActive()
    ExportLifeKinrenPlanCsvForOwner frmEval
End Sub

Public Sub ExportLifeKinrenPlanCsvForOwner(Optional ByVal owner As Object)
    Dim headers As Variant
    Dim values() As String
    Dim ws As Worksheet
    Dim rowNo As Long
    Dim facilityNo As String
    Dim evalDate As String
    Dim insuredNo As String
    Dim outDir As String
    Dim outPath As String

    headers = Split(LifePlanHeader(), ",")
    ReDim values(LBound(headers) To UBound(headers))

    ResolveActiveEvaluation owner, ws, rowNo

    facilityNo = GetConfigValue("FacilityNo")
    If Len(facilityNo) = 0 Then facilityNo = GetConfigValue("care_facility_id")
    evalDate = ToLifeDate(GetSourceValue(owner, ws, rowNo, Array("txtEDate"), Array("Basic.EvalDate", "EvalDate")))
    insuredNo = ResolveInsuredNo(owner, ws, rowNo)

    PutLifeKinrenPlanValues headers, values, owner, ws, rowNo, facilityNo, evalDate, insuredNo

    outDir = ThisWorkbook.path & Application.PathSeparator & "CSV" & Application.PathSeparator & IIf(Len(facilityNo) > 0, facilityNo, "LIFE")
    EnsureFolderPath outDir
    outPath = outDir & Application.PathSeparator & "17_" & LIFE_KIND & "_" & IIf(Len(facilityNo) > 0, facilityNo, "UNKNOWN") & "_" & Left$(evalDate & Format(Date, "yyyymm"), 6) & "_" & Format(Now, "yyyymmddhhnnss") & ".csv"

    WriteUtf8BomText outPath, LIFE_KIND & vbCrLf & LifePlanHeader() & vbCrLf & CsvLine(values)
    MsgBox "LIFE CSV exported:" & vbCrLf & outPath, vbInformation
End Sub

Public Sub ExportLifeKinrenPlanCsvBatch()
    Dim headers As Variant
    Dim sourceWs As Worksheet
    Dim facilityNo As String
    Dim outDir As String
    Dim outPath As String
    Dim logPath As String
    Dim body As String
    Dim logText As String
    Dim processed As Long
    Dim skipped As Long

    headers = Split(LifePlanHeader(), ",")
    facilityNo = GetConfigValue("FacilityNo")
    If Len(facilityNo) = 0 Then facilityNo = GetConfigValue("care_facility_id")

    Set sourceWs = ResolveBatchSourceSheet()
    If sourceWs Is Nothing Then
        MsgBox "Batch source sheet was not found.", vbExclamation
        Exit Sub
    End If

    body = LIFE_KIND & vbCrLf & LifePlanHeader()
    logText = "source,row,reason"

    If FindHeaderCol(sourceWs, "SheetName") > 0 Then
        AppendBatchRowsFromIndex headers, sourceWs, facilityNo, body, logText, processed, skipped
    Else
        AppendBatchRowsFromWorksheet headers, sourceWs, facilityNo, body, logText, processed, skipped
    End If

    outDir = ThisWorkbook.path & Application.PathSeparator & "CSV" & Application.PathSeparator & IIf(Len(facilityNo) > 0, facilityNo, "LIFE")
    EnsureFolderPath outDir
    outPath = outDir & Application.PathSeparator & "17_" & LIFE_KIND & "_" & IIf(Len(facilityNo) > 0, facilityNo, "UNKNOWN") & "_batch_" & Format(Now, "yyyymmddhhnnss") & ".csv"
    logPath = outDir & Application.PathSeparator & "17_" & LIFE_KIND & "_batch_log_" & Format(Now, "yyyymmddhhnnss") & ".csv"

    If processed = 0 Then
        If skipped > 0 Then WriteUtf8BomText logPath, logText
        MsgBox "No LIFE CSV rows were exported." & IIf(skipped > 0, vbCrLf & "Log: " & logPath, ""), vbExclamation
        Exit Sub
    End If

    WriteUtf8BomText outPath, body
    If skipped > 0 Then WriteUtf8BomText logPath, logText
    MsgBox "LIFE batch CSV exported:" & vbCrLf & outPath & vbCrLf & "Rows: " & processed & vbCrLf & "Skipped: " & skipped & IIf(skipped > 0, vbCrLf & "Log: " & logPath, ""), vbInformation
End Sub

Private Sub PutLifeKinrenPlanValues(ByVal headers As Variant, ByRef values() As String, ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long, ByVal facilityNo As String, ByVal evalDate As String, ByVal insuredNo As String)
    PutCsvValue headers, values, "care_facility_id", facilityNo                         ' Source: AppConfig.FacilityNo.
    PutCsvValue headers, values, "service_code", GetLifeSettingValue("SERVICE_CODE_PLAN", "15") ' Source: fixed day-service default; LIFE setting can override.
    PutCsvValue headers, values, "insurer_no", GetRowValueAny(ws, rowNo, Array("InsurerNo")) ' Source: saved evaluation row when present.
    PutCsvValue headers, values, "insured_no", insuredNo                              ' Source: saved evaluation row first; active form fallback.
    PutCsvValue headers, values, "external_system_management_number", ExternalKey(owner, ws, rowNo, evalDate) ' Source: saved ExternalSystemKey or local fallback.
    PutCsvValue headers, values, "trinity_attempt", "0"
    PutCsvValue headers, values, "evaluate_date", evalDate                              ' Source: active form txtEDate / saved Basic.EvalDate.
    PutCsvValue headers, values, "last_date", PreviousEvaluationDate(ws, rowNo)         ' Source: previous row in the same per-user EV_* sheet.
    PutCsvValue headers, values, "first_date", FirstEvaluationDate(owner, ws, evalDate) ' Source: EvalIndex.FirstEvalDate; evaluate_date fallback.
    PutCsvValue headers, values, "care_level", ToLifeCareLevel(GetSourceValue(owner, ws, rowNo, Array("cboCare"), Array("Basic.CareLevel", "CareLevel"))) ' Source: active care-level control / saved row.
    PutCsvValue headers, values, "impaired_elderly_independence_degree", ToLifeImpairedElderlyDegree(GetSourceValue(owner, ws, rowNo, Empty, Array("Basic.ImpairedElderlyADL", "áŠQ‚—îŽÒ‚Ì“úí¶ŠˆŽ©—§“x", "‚—îŽÒ‚Ì“úí¶ŠˆŽ©—§“x")))
    PutCsvValue headers, values, "dementia_elderly_independence_degree", ToLifeDementiaElderlyDegree(GetSourceValue(owner, ws, rowNo, Array("cboDementia"), Array("Basic.DementiaADL", "Basic.DementiaLevel", "”F’mÇ‚—îŽÒ‚Ì“úí¶ŠˆŽ©—§“x"))) ' Source: active dementia ADL control / saved row.
    PutCsvValue headers, values, "user_request", GetSourceValue(owner, ws, rowNo, Array("txtNeedsPt"), Array("Basic.Needs.Patient", "Needs.Patient")) ' Source: active needs field / saved row.
    PutCsvValue headers, values, "user_family_request", GetSourceValue(owner, ws, rowNo, Array("txtNeedsFam"), Array("Basic.Needs.Family", "Needs.Family")) ' Source: active family-needs field / saved row.
    PutCsvValue headers, values, "social_participation", GetSourceValue(owner, ws, rowNo, Array("txtLiving"), Array("BI.SocialParticipation", "Basic.LifeStatus")) ' Source: active living/social field / saved row.
    PutCsvValue headers, values, "housing_environment", GetSourceValue(owner, ws, rowNo, Empty, Array("Basic.HomeEnv.Note", "Basic.HomeEnv.Checks", "Basic.Living"))
    PutCsvValue headers, values, "disease_name_code", TemporaryDiseaseNameCode(owner, ws, rowNo)
    PutDateParts headers, values, "onset_date", GetSourceValue(owner, ws, rowNo, Array("txtOnset"), Array("Basic.OnsetDate"))
    PutDateParts headers, values, "latest_admission_date", GetSourceValue(owner, ws, rowNo, Array("txtAdmDate"), Array("Basic.Medical.AdmitDate"))
    PutDateParts headers, values, "latest_discharge_date", GetSourceValue(owner, ws, rowNo, Array("txtDisDate"), Array("Basic.Medical.DischargeDate"))
    PutCsvValue headers, values, "progress", GetSourceValue(owner, ws, rowNo, Array("txtTxCourse"), Array("Basic.Medical.CourseNote"))
    PutComplicationFlags headers, values, ResolveComplicationsText(owner, ws, rowNo) ' Source: primary diagnosis, history, complication memo, and medical notes.
    PutCsvValue headers, values, "implementation_status", ResolveImplementationStatus(owner, ws, rowNo)

    PutPlanSheetValues headers, values                                                   ' Source: current generated plan worksheet.

    PutCsvValue headers, values, "program_planner", GetSourceValue(owner, ws, rowNo, Array("txtEvaluator"), Array("Basic.Evaluator"))
    PutCsvValue headers, values, "himself_or_family_routine_work", PlanSheetText(45, 2, 62)
    PutCsvValue headers, values, "functional_training_notice", PlanSheetText(19, 2, 62)
    PutCsvValue headers, values, "functional_training_summary", PlanSheetText(49, 2, 62)
    PutCsvValue headers, values, "subject_and_factors", GetSourceValue(owner, ws, rowNo, Empty, Array("EvalTestCriticalFindings", "IO_Mental_Note", "IO_Cog_DementiaNote"))
    PutCsvValue headers, values, "version", "0310"
End Sub

Private Function ResolveBatchSourceSheet() As Worksheet
    On Error Resume Next
    Set ResolveBatchSourceSheet = ActiveSheet
    On Error GoTo 0
    If Not ResolveBatchSourceSheet Is Nothing Then
        If LastUsedRow(ResolveBatchSourceSheet) >= 2 Then
            If FindHeaderCol(ResolveBatchSourceSheet, "SheetName") > 0 Or FindHeaderCol(ResolveBatchSourceSheet, "Basic.EvalDate") > 0 Or FindHeaderCol(ResolveBatchSourceSheet, "EvalDate") > 0 Then Exit Function
        End If
    End If
    On Error Resume Next
    Set ResolveBatchSourceSheet = ThisWorkbook.Worksheets("EvalIndex")
    On Error GoTo 0
End Function

Private Sub AppendBatchRowsFromIndex(ByVal headers As Variant, ByVal indexWs As Worksheet, ByVal facilityNo As String, ByRef body As String, ByRef logText As String, ByRef processed As Long, ByRef skipped As Long)
    Dim r As Long
    Dim cSheet As Long
    Dim cEval As Long
    Dim sheetName As String
    Dim evalDate As String
    Dim ws As Worksheet
    Dim rowNo As Long

    cSheet = FindHeaderCol(indexWs, "SheetName")
    cEval = FindHeaderCol(indexWs, "Basic.EvalDate")
    If cEval = 0 Then cEval = FindHeaderCol(indexWs, "EvalDate")
    If cEval = 0 Then cEval = FindHeaderCol(indexWs, "LastEvalDate")

    For r = 2 To LastUsedRow(indexWs)
        sheetName = Trim$(CStr(indexWs.Cells(r, cSheet).value))
        Set ws = WorksheetByName(sheetName)
        If ws Is Nothing Then
            AppendBatchLog logText, indexWs.name, r, "sheet not found"
            skipped = skipped + 1
        Else
            evalDate = ""
            If cEval > 0 Then evalDate = ToLifeDate(CStr(indexWs.Cells(r, cEval).value))
            rowNo = FindRowByEvalDate(ws, evalDate)
            If rowNo = 0 Then rowNo = LastUsedRow(ws)
            AppendBatchRow headers, ws, rowNo, facilityNo, body, logText, processed, skipped
        End If
    Next r
End Sub

Private Sub AppendBatchRowsFromWorksheet(ByVal headers As Variant, ByVal ws As Worksheet, ByVal facilityNo As String, ByRef body As String, ByRef logText As String, ByRef processed As Long, ByRef skipped As Long)
    Dim r As Long
    For r = 2 To LastUsedRow(ws)
        AppendBatchRow headers, ws, r, facilityNo, body, logText, processed, skipped
    Next r
End Sub

Private Sub AppendBatchRow(ByVal headers As Variant, ByVal ws As Worksheet, ByVal rowNo As Long, ByVal facilityNo As String, ByRef body As String, ByRef logText As String, ByRef processed As Long, ByRef skipped As Long)
    Dim values() As String
    Dim evalDate As String
    Dim insuredNo As String
    Dim missing As String

    ReDim values(LBound(headers) To UBound(headers))
    evalDate = ToLifeDate(GetSourceValue(Nothing, ws, rowNo, Empty, Array("Basic.EvalDate", "EvalDate")))
    insuredNo = ResolveInsuredNo(Nothing, ws, rowNo)
    missing = MissingLifePlanRequiredFields(facilityNo, insuredNo, evalDate)
    If Len(missing) > 0 Then
        AppendBatchLog logText, ws.name, rowNo, missing
        skipped = skipped + 1
        Exit Sub
    End If

    PutLifeKinrenPlanValues headers, values, Nothing, ws, rowNo, facilityNo, evalDate, insuredNo
    body = body & vbCrLf & CsvLine(values)
    processed = processed + 1
End Sub

Private Function MissingLifePlanRequiredFields(ByVal facilityNo As String, ByVal insuredNo As String, ByVal evalDate As String) As String
    If Len(facilityNo) = 0 Then MissingLifePlanRequiredFields = AppendMissingField(MissingLifePlanRequiredFields, "care_facility_id")
    If Len(insuredNo) = 0 Then MissingLifePlanRequiredFields = AppendMissingField(MissingLifePlanRequiredFields, "insured_no")
    If Len(evalDate) = 0 Then MissingLifePlanRequiredFields = AppendMissingField(MissingLifePlanRequiredFields, "evaluate_date")
End Function

Private Function AppendMissingField(ByVal currentValue As String, ByVal fieldName As String) As String
    AppendMissingField = currentValue & IIf(Len(currentValue) > 0, ";", "") & fieldName
End Function

Private Sub AppendBatchLog(ByRef logText As String, ByVal sourceName As String, ByVal rowNo As Long, ByVal reason As String)
    logText = logText & vbCrLf & CsvEscape(sourceName) & "," & CStr(rowNo) & "," & CsvEscape(reason)
    Debug.Print "[LIFE_PLAN_CSV][BATCH] " & sourceName & " row " & rowNo & ": " & reason
End Sub

Private Function WorksheetByName(ByVal sheetName As String) As Worksheet
    If Len(sheetName) = 0 Then Exit Function
    On Error Resume Next
    Set WorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub ResolveActiveEvaluation(ByVal owner As Object, ByRef ws As Worksheet, ByRef rowNo As Long)
    Dim targetName As String
    Dim targetID As String
    Dim targetDate As String
    Dim indexWs As Worksheet
    Dim sheetName As String

    targetName = GetControlTextAny(owner, Array("txtHdrName", "txtName"))
    targetID = GetControlTextAny(owner, Array("txtHdrPID", "txtPID"))
    targetDate = ToLifeDate(GetControlTextAny(owner, Array("txtEDate")))

    On Error Resume Next
    Set indexWs = ThisWorkbook.Worksheets("EvalIndex")
    On Error GoTo 0

    If Not indexWs Is Nothing Then
        sheetName = FindEvalSheetName(indexWs, targetID, targetName)
        If Len(sheetName) > 0 Then
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(sheetName)
            On Error GoTo 0
        End If
    End If

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("EvalData")
        On Error GoTo 0
    End If

    If ws Is Nothing Then Exit Sub
    rowNo = FindRowByEvalDate(ws, targetDate)
    If rowNo = 0 Then rowNo = LastUsedRow(ws)
End Sub

Private Function FindEvalSheetName(ByVal ws As Worksheet, ByVal targetID As String, ByVal targetName As String) As String
    Dim r As Long
    Dim lastR As Long
    Dim cID As Long
    Dim cName As Long
    Dim cSheet As Long

    cID = FindHeaderCol(ws, "UserID")
    cName = FindHeaderCol(ws, "Name")
    cSheet = FindHeaderCol(ws, "SheetName")
    If cSheet = 0 Then Exit Function

    lastR = LastUsedRow(ws)
    For r = 2 To lastR
        If Len(targetID) > 0 And cID > 0 Then
            If Trim$(CStr(ws.Cells(r, cID).value)) = targetID Then
                FindEvalSheetName = Trim$(CStr(ws.Cells(r, cSheet).value))
                Exit Function
            End If
        End If
        If Len(targetName) > 0 And cName > 0 Then
            If Trim$(CStr(ws.Cells(r, cName).value)) = targetName Then
                FindEvalSheetName = Trim$(CStr(ws.Cells(r, cSheet).value))
                Exit Function
            End If
        End If
    Next r
End Function

Private Function FindRowByEvalDate(ByVal ws As Worksheet, ByVal lifeDate As String) As Long
    Dim c As Long
    Dim r As Long
    If Len(lifeDate) = 0 Then Exit Function
    c = FindHeaderCol(ws, "Basic.EvalDate")
    If c = 0 Then c = FindHeaderCol(ws, "EvalDate")
    If c = 0 Then Exit Function

    For r = LastUsedRow(ws) To 2 Step -1
        If ToLifeDate(CStr(ws.Cells(r, c).value)) = lifeDate Then
            FindRowByEvalDate = r
            Exit Function
        End If
    Next r
End Function

Private Function GetSourceValue(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long, ByVal controlNames As Variant, ByVal headerNames As Variant) As String
    Dim v As String
    v = GetControlTextAny(owner, controlNames)
    If Len(v) > 0 Then
        GetSourceValue = v
    Else
        GetSourceValue = GetRowValueAny(ws, rowNo, headerNames)
    End If
End Function

Private Function GetControlTextAny(ByVal owner As Object, ByVal names As Variant) As String
    Dim i As Long
    Dim v As String
    On Error GoTo Done
    If IsMissingOrEmpty(names) Then Exit Function
    For i = LBound(names) To UBound(names)
        v = GetControlTextRecursive(owner, CStr(names(i)))
        If Len(v) > 0 Then GetControlTextAny = v: Exit Function
    Next i
Done:
End Function

Private Function GetControlTextRecursive(ByVal parent As Object, ByVal controlName As String) As String
    Dim c As Object
    Dim pg As Object
    Dim v As String

    If parent Is Nothing Then Exit Function
    On Error Resume Next
    v = Trim$(CStr(parent.controls(controlName).value))
    If Len(v) = 0 Then v = Trim$(CStr(parent.controls(controlName).text))
    On Error GoTo 0
    If Len(v) > 0 Then GetControlTextRecursive = v: Exit Function

    On Error Resume Next
    For Each c In parent.controls
        If TypeName(c) = "MultiPage" Then
            For Each pg In c.pages
                v = GetControlTextRecursive(pg, controlName)
                If Len(v) > 0 Then GetControlTextRecursive = v: Exit Function
            Next pg
        Else
            v = GetControlTextRecursive(c, controlName)
            If Len(v) > 0 Then GetControlTextRecursive = v: Exit Function
        End If
    Next c
    On Error GoTo 0
End Function

Private Function GetRowValueAny(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headers As Variant) As String
    Dim i As Long
    Dim c As Long
    If ws Is Nothing Or rowNo < 2 Or IsMissingOrEmpty(headers) Then Exit Function
    For i = LBound(headers) To UBound(headers)
        c = FindHeaderCol(ws, CStr(headers(i)))
        If c > 0 Then
            GetRowValueAny = Trim$(CStr(ws.Cells(rowNo, c).value))
            If Len(GetRowValueAny) > 0 Then Exit Function
        End If
    Next i
End Function

Private Function ResolveInsuredNo(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long) As String
    ResolveInsuredNo = GetRowValueAny(ws, rowNo, Array("InsuredNo"))
    If Len(ResolveInsuredNo) = 0 Then
        ResolveInsuredNo = GetControlTextAny(owner, Array("txtInsuredNo", "txtInsuredNumber", "txtCareInsuredNo", "txtHdrPID", "txtPID"))
    End If
    If Len(ResolveInsuredNo) = 0 Then Debug.Print "[LIFE_PLAN_CSV][WARN] insured_no is blank"
End Function

Private Function TemporaryDiseaseNameCode(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long) As String
    TemporaryDiseaseNameCode = ""
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastC As Long
    Dim c As Long
    If ws Is Nothing Then Exit Function
    lastC = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastC
        If Trim$(CStr(ws.Cells(1, c).value)) = headerName Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Exit Function
    LastUsedRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If LastUsedRow < 2 Then LastUsedRow = ws.UsedRange.rows(ws.UsedRange.rows.count).row
End Function

Private Function GetConfigValue(ByVal keyName As String) As String
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("AppConfig")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = 1 To LastUsedRow(ws)
        If Trim$(CStr(ws.Cells(r, 1).value)) = keyName Then
            GetConfigValue = Trim$(CStr(ws.Cells(r, 2).value))
            Exit Function
        End If
    Next r
End Function

Private Function GetLifeSettingValue(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    On Error GoTo 0
    If ws Is Nothing Then GetLifeSettingValue = defaultValue: Exit Function
    For r = 1 To LastUsedRow(ws)
        If Trim$(CStr(ws.Cells(r, 1).value)) = keyName Then
            GetLifeSettingValue = Trim$(CStr(ws.Cells(r, 2).value))
            Exit Function
        End If
    Next r
    GetLifeSettingValue = defaultValue
End Function

Private Function ExternalKey(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long, ByVal evalDate As String) As String
    ExternalKey = GetRowValueAny(ws, rowNo, Array("ExternalSystemKey", "external_system_management_number"))
    If Len(ExternalKey) = 0 Then
        ExternalKey = GetRowValueAny(ws, rowNo, Array("InsuredNo"))
        If Len(ExternalKey) = 0 Then ExternalKey = GetControlTextAny(owner, Array("txtHdrPID", "txtPID"))
        If Len(ExternalKey) > 0 And Len(evalDate) > 0 Then ExternalKey = ExternalKey & "_" & evalDate
    End If
End Function

Private Function FirstEvaluationDate(ByVal owner As Object, ByVal activeWs As Worksheet, ByVal evalDate As String) As String
    Dim ws As Worksheet
    Dim targetName As String
    Dim targetID As String
    Dim r As Long
    Dim cID As Long
    Dim cName As Long
    Dim cFirst As Long

    targetName = GetControlTextAny(owner, Array("txtHdrName", "txtName"))
    targetID = GetControlTextAny(owner, Array("txtHdrPID", "txtPID"))

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalIndex")
    On Error GoTo 0
    If ws Is Nothing Then FirstEvaluationDate = evalDate: Exit Function
    cID = FindHeaderCol(ws, "UserID")
    cName = FindHeaderCol(ws, "Name")
    cFirst = FindHeaderCol(ws, "FirstEvalDate")
    If cFirst = 0 Then FirstEvaluationDate = evalDate: Exit Function

    For r = 2 To LastUsedRow(ws)
        If Len(targetID) > 0 And cID > 0 And Trim$(CStr(ws.Cells(r, cID).value)) = targetID Then
            FirstEvaluationDate = ToLifeDate(CStr(ws.Cells(r, cFirst).value))
            Exit Function
        End If
        If Len(targetName) > 0 And cName > 0 And Trim$(CStr(ws.Cells(r, cName).value)) = targetName Then
            FirstEvaluationDate = ToLifeDate(CStr(ws.Cells(r, cFirst).value))
            Exit Function
        End If
    Next r
    FirstEvaluationDate = evalDate
End Function

Private Function PreviousEvaluationDate(ByVal ws As Worksheet, ByVal rowNo As Long) As String
    Dim c As Long
    If ws Is Nothing Or rowNo <= 2 Then Exit Function
    c = FindHeaderCol(ws, "Basic.EvalDate")
    If c = 0 Then c = FindHeaderCol(ws, "EvalDate")
    If c > 0 Then PreviousEvaluationDate = ToLifeDate(CStr(ws.Cells(rowNo - 1, c).value))
End Function

Private Function ToLifeDate(ByVal rawValue As String) As String
    Dim d As Date
    rawValue = Trim$(rawValue)
    If Len(rawValue) = 0 Then Exit Function
    If IsDate(rawValue) Then
        d = CDate(rawValue)
        ToLifeDate = Format$(d, "yyyymmdd")
    Else
        rawValue = Replace(Replace(Replace(rawValue, "/", ""), "-", ""), ".", "")
        If Len(rawValue) >= 8 And IsNumeric(Left$(rawValue, 8)) Then ToLifeDate = Left$(rawValue, 8)
    End If
End Function

Private Sub PutDateParts(ByVal headers As Variant, ByRef values() As String, ByVal baseName As String, ByVal rawDate As String)
    Dim s As String
    s = ToLifeDate(rawDate)
    If Len(s) <> 8 Then Exit Sub
    PutCsvValue headers, values, baseName & "_year", Left$(s, 4)
    PutCsvValue headers, values, baseName & "_month", Mid$(s, 5, 2)
    PutCsvValue headers, values, baseName & "_day", Right$(s, 2)
End Sub

Private Function ToLifeCareLevel(ByVal rawValue As String) As String
    rawValue = NormalizeLifeText(rawValue)
    If Len(rawValue) = 0 Then Exit Function
    If rawValue = "01" Or rawValue = "06" Or rawValue = "11" _
            Or rawValue = "12" Or rawValue = "13" _
            Or rawValue = "21" Or rawValue = "22" Or rawValue = "23" _
            Or rawValue = "24" Or rawValue = "25" Then
        ToLifeCareLevel = rawValue
        Exit Function
    End If
    Select Case rawValue
        Case "”ñŠY“–": ToLifeCareLevel = "01"
        Case "Ž–‹Æ‘ÎÛŽÒ": ToLifeCareLevel = "06"
        Case "—vŽx‰‡iŒo‰ß“I—v‰îŒìj", "—vŽx‰‡(Œo‰ß“I—v‰îŒì)": ToLifeCareLevel = "11"
        Case "—vŽx‰‡1": ToLifeCareLevel = "12"
        Case "—vŽx‰‡2": ToLifeCareLevel = "13"
        Case "—v‰îŒì1": ToLifeCareLevel = "21"
        Case "—v‰îŒì2": ToLifeCareLevel = "22"
        Case "—v‰îŒì3": ToLifeCareLevel = "23"
        Case "—v‰îŒì4": ToLifeCareLevel = "24"
        Case "—v‰îŒì5": ToLifeCareLevel = "25"
    End Select
End Function

Private Function ToLifeImpairedElderlyDegree(ByVal rawValue As String) As String
    rawValue = UCase$(NormalizeLifeText(rawValue))
    Select Case rawValue
        Case "1", "Ž©—§": ToLifeImpairedElderlyDegree = "1"
        Case "2", "J1": ToLifeImpairedElderlyDegree = "2"
        Case "3", "J2": ToLifeImpairedElderlyDegree = "3"
        Case "4", "A1": ToLifeImpairedElderlyDegree = "4"
        Case "5", "A2": ToLifeImpairedElderlyDegree = "5"
        Case "6", "B1": ToLifeImpairedElderlyDegree = "6"
        Case "7", "B2": ToLifeImpairedElderlyDegree = "7"
        Case "8", "C1": ToLifeImpairedElderlyDegree = "8"
        Case "9", "C2": ToLifeImpairedElderlyDegree = "9"
    End Select
End Function

Private Function ToLifeDementiaElderlyDegree(ByVal rawValue As String) As String
    rawValue = UCase$(NormalizeLifeText(rawValue))
    rawValue = Replace(rawValue, "‡T", "I")
    rawValue = Replace(rawValue, "‡U", "II")
    rawValue = Replace(rawValue, "‡V", "III")
    rawValue = Replace(rawValue, "‡W", "IV")
    Select Case rawValue
        Case "1", "Ž©—§": ToLifeDementiaElderlyDegree = "1"
        Case "2", "I": ToLifeDementiaElderlyDegree = "2"
        Case "3", "IIA": ToLifeDementiaElderlyDegree = "3"
        Case "4", "IIB": ToLifeDementiaElderlyDegree = "4"
        Case "5", "IIIA": ToLifeDementiaElderlyDegree = "5"
        Case "6", "IIIB": ToLifeDementiaElderlyDegree = "6"
        Case "7", "IV": ToLifeDementiaElderlyDegree = "7"
        Case "8", "M": ToLifeDementiaElderlyDegree = "8"
    End Select
End Function

Private Sub PutComplicationFlags(ByVal headers As Variant, ByRef values() As String, ByVal note As String)
    PutComplicationFlag headers, values, "complications_control_cerebrovascular_disease", note, Array("”][Ç", "”]oŒŒ", "‚­‚à–Œ‰º")
    PutComplicationFlag headers, values, "complications_control_fracture", note, Array("œÜ")
    PutComplicationFlag headers, values, "complications_control_pneumonia", note, Array("”x‰Š")
    PutComplicationFlag headers, values, "complications_control_congestive_heart_failure", note, Array("S•s‘S")
    PutComplicationFlag headers, values, "complications_control_urinary_tract_infection", note, Array("”A˜HŠ´õ")
    PutComplicationFlag headers, values, "complications_control_diabetes", note, Array("“œ”A•a")
    PutComplicationFlag headers, values, "complications_control_hypertension", note, Array("‚ŒŒˆ³")
    PutComplicationFlag headers, values, "complications_control_osteoporosis", note, Array("œ‘eé Ç")
    PutComplicationFlag headers, values, "complications_control_dementia", note, Array("”F’mÇ")
End Sub

Private Sub PutComplicationFlag(ByVal headers As Variant, ByRef values() As String, ByVal fieldName As String, ByVal note As String, ByVal keywords As Variant)
    PutCsvValue headers, values, fieldName, ToComplicationFlag(note, keywords)
End Sub

Private Function ResolveComplicationsText(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long) As String
    ResolveComplicationsText = JoinNonBlank(Array( _
        GetSourceValue(owner, ws, rowNo, Array("txtDx"), Array("Basic.PrimaryDx", "Žåf’f", "Žå•a–¼")), _
        GetSourceValue(owner, ws, rowNo, Array("txtHistory", "txtHx", "txtPastHistory"), Array("Basic.Medical.History", "Basic.PastHistory", "Šù‰—ð")), _
        GetSourceValue(owner, ws, rowNo, Empty, Array("Basic.Medical.ComplicationNote", "‡•¹Çƒƒ‚", "‡•¹Ç")), _
        GetSourceValue(owner, ws, rowNo, Array("txtMedicalNote", "txtMedNote"), Array("Basic.Medical.Note", "Medical.Note", "ˆã—Ã”õl")) _
    ))
End Function

Private Function JoinNonBlank(ByVal parts As Variant) As String
    Dim i As Long
    Dim v As String
    For i = LBound(parts) To UBound(parts)
        v = Trim$(CStr(parts(i)))
        If Len(v) > 0 Then JoinNonBlank = JoinNonBlank & IIf(Len(JoinNonBlank) > 0, " ", "") & v
    Next i
End Function

Private Function ResolveImplementationStatus(ByVal owner As Object, ByVal ws As Worksheet, ByVal rowNo As Long) As String
    ResolveImplementationStatus = ToImplementationStatusCode(JoinNonBlank(Array( _
        GetSourceValue(owner, ws, rowNo, Empty, Array("EvalTestCriticalFindings", "EvaluationMemo", "•]‰¿ƒƒ‚", "TONE_NOTE", "SENSE_NOTE")), _
        GetSourceValue(owner, ws, rowNo, Empty, Array("Plan.Status", "PlanStatus", "Œv‰æ‘ó‘Ô")), _
        GetSourceValue(owner, ws, rowNo, Empty, Array("ImplementationStatus", "Status", "Basic.Status", "OtherStatus", "ƒXƒe[ƒ^ƒX")) _
    )))
End Function

Private Function ToImplementationStatusCode(ByVal rawValue As String) As String
    Dim s As String
    Dim isActive As Boolean
    Dim isEnded As Boolean
    s = NormalizeLifeText(rawValue)
    If Len(s) = 0 Then Exit Function
    isActive = (InStr(1, s, "Œp‘±", vbTextCompare) > 0 Or InStr(1, s, "ŽÀŽ{’†", vbTextCompare) > 0)
    isEnded = (InStr(1, s, "I—¹", vbTextCompare) > 0 Or InStr(1, s, "’†Ž~", vbTextCompare) > 0)
    If isActive = isEnded Then Exit Function
    If isActive Then
        ToImplementationStatusCode = "1"
    ElseIf isEnded Then
        ToImplementationStatusCode = "2"
    End If
End Function

Private Function ToComplicationFlag(ByVal note As String, ByVal keywords As Variant) As String
    Dim i As Long
    Dim n As String
    n = NormalizeLifeText(note)
    If Len(n) = 0 Then Exit Function
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, n, NormalizeLifeText(CStr(keywords(i))), vbTextCompare) > 0 Then
            ToComplicationFlag = "1"
            Exit Function
        End If
    Next i
End Function

Private Function IsClearlyNone(ByVal s As String) As Boolean
    IsClearlyNone = (s = "‚È‚µ" Or s = "–³‚µ" Or s = "–³" Or s = "“Á‚É‚È‚µ" Or s = "‚È‚µB")
End Function

Private Function HasExplicitYesNear(ByVal textValue As String, ByVal keyword As String) As Boolean
    Dim p As Long
    Dim part As String
    p = InStr(1, textValue, keyword, vbTextCompare)
    If p = 0 Then Exit Function
    part = Mid$(textValue, p, Len(keyword) + 12)
    HasExplicitYesNear = (InStr(part, "‚ ‚è") > 0 Or InStr(part, "—L") > 0 _
        Or InStr(part, "1") > 0 Or InStr(part, "›") > 0 Or InStr(part, "Z") > 0)
End Function

Private Function HasExplicitNoNear(ByVal textValue As String, ByVal keyword As String) As Boolean
    Dim p As Long
    Dim part As String
    p = InStr(1, textValue, keyword, vbTextCompare)
    If p = 0 Then Exit Function
    part = Mid$(textValue, p, Len(keyword) + 12)
    HasExplicitNoNear = (InStr(part, "‚È‚µ") > 0 Or InStr(part, "–³‚µ") > 0 _
        Or InStr(part, "–³") > 0 Or InStr(part, "0") > 0)
End Function

Private Sub PutPlanSheetValues(ByVal headers As Variant, ByRef values() As String)
    Dim shortMindBody As String
    Dim shortActivity As String
    Dim shortActivityJoin As String
    Dim longMindBody As String
    Dim longActivity As String
    Dim longActivityJoin As String

    shortMindBody = PlanSheetText(24, 2, 31)
    shortActivity = PlanSheetText(25, 2, 31)
    shortActivityJoin = PlanSheetText(26, 2, 31)
    longMindBody = PlanSheetText(24, 32, 62)
    longActivity = PlanSheetText(25, 32, 62)
    longActivityJoin = PlanSheetText(26, 32, 62)

    PutGoalCategory headers, values, "short_goal_functional_training_mind_and_body_function", shortMindBody
    PutGoalCategory headers, values, "short_goal_functional_training_activity", shortActivity
    PutGoalCategory headers, values, "short_goal_functional_training_activity_join", shortActivityJoin
    PutGoalCategory headers, values, "long_goal_functional_training_mind_and_body_function", longMindBody
    PutGoalCategory headers, values, "long_goal_functional_training_activity", longActivity
    PutGoalCategory headers, values, "long_goal_functional_training_activity_join", longActivityJoin
    PutCsvValue headers, values, "short_goal_achievement", ToAchievementCode(PlanSheetText(23, 2, 31))
    PutCsvValue headers, values, "long_goal_achievement", ToAchievementCode(PlanSheetText(23, 32, 62))
    PutProgram headers, values, 1, 29
    PutProgram headers, values, 2, 32
    PutProgram headers, values, 3, 35
    PutProgram headers, values, 4, 38
End Sub

Private Sub PutGoalCategory(ByVal headers As Variant, ByRef values() As String, ByVal fieldPrefix As String, ByVal contents As String)
    PutCsvValue headers, values, fieldPrefix & "1", IIf(Len(contents) > 0, "1", "")
    PutCsvValue headers, values, fieldPrefix & "2", ""
    PutCsvValue headers, values, fieldPrefix & "3", ""
    PutCsvValue headers, values, fieldPrefix & "_contents", contents
End Sub

Private Sub PutProgram(ByVal headers As Variant, ByRef values() As String, ByVal indexNo As Long, ByVal rowNo As Long)
    Dim suffix As String
    suffix = Format$(indexNo, "00")
    PutCsvValue headers, values, "function_training_content_" & suffix, PlanSheetText(rowNo, 2, 40)
    PutCsvValue headers, values, "function_training_note_" & suffix, PlanSheetText(rowNo + 1, 2, 40)
    PutCsvValue headers, values, "function_training_frequency_times_" & suffix, ExtractWeeklyFrequency(PlanSheetText(rowNo, 41, 50))
    PutCsvValue headers, values, "function_training_date_" & suffix, ExtractTrainingMinutes(PlanSheetText(rowNo, 51, 56))
    PutCsvValue headers, values, "function_training_personnel_" & suffix, FirstNumberText(PlanSheetText(rowNo, 57, 62))
End Sub

Private Function PlanSheetText(ByVal rowNo As Long, ByVal firstCol As Long, ByVal lastCol As Long) As String
    Dim ws As Worksheet
    Dim c As Long
    Dim s As String
    Dim v As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(7)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For c = firstCol To lastCol
        v = Trim$(CStr(ws.Cells(rowNo, c).value))
        If Len(v) > 0 Then
            If InStr(1, s, v, vbTextCompare) = 0 Then s = s & IIf(Len(s) > 0, " ", "") & v
        End If
    Next c
    PlanSheetText = s
End Function

Private Function FirstNumberText(ByVal rawValue As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(rawValue)
        ch = Mid$(rawValue, i, 1)
        If ch Like "#" Then FirstNumberText = FirstNumberText & ch
    Next i
End Function

Private Function ExtractWeeklyFrequency(ByVal rawValue As String) As String
    Dim s As String
    s = NormalizeLifeText(rawValue)
    If Len(s) = 0 Then Exit Function
    If InStr(1, s, "–ˆ“ú", vbTextCompare) > 0 Then
        ExtractWeeklyFrequency = "7"
        Exit Function
    End If
    ExtractWeeklyFrequency = ExtractNumberBetween(s, "T", "‰ñ")
End Function

Private Function ExtractTrainingMinutes(ByVal rawValue As String) As String
    Dim s As String
    Dim hoursText As String
    s = NormalizeLifeText(rawValue)
    If Len(s) = 0 Then Exit Function
    ExtractTrainingMinutes = ExtractNumberBefore(s, "•ª")
    If Len(ExtractTrainingMinutes) > 0 Then Exit Function
    hoursText = ExtractNumberBefore(s, "ŽžŠÔ")
    If Len(hoursText) > 0 Then ExtractTrainingMinutes = CStr(CLng(hoursText) * 60)
End Function

Private Function ExtractNumberBetween(ByVal textValue As String, ByVal prefixText As String, ByVal suffixText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(1, textValue, prefixText, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(prefixText)
    endPos = InStr(startPos, textValue, suffixText, vbTextCompare)
    If endPos = 0 Then Exit Function
    ExtractNumberBetween = DigitsOnly(Mid$(textValue, startPos, endPos - startPos))
End Function

Private Function ExtractNumberBefore(ByVal textValue As String, ByVal suffixText As String) As String
    Dim suffixPos As Long
    Dim i As Long
    Dim ch As String
    suffixPos = InStr(1, textValue, suffixText, vbTextCompare)
    If suffixPos = 0 Then Exit Function
    For i = suffixPos - 1 To 1 Step -1
        ch = Mid$(textValue, i, 1)
        If Not ch Like "#" Then Exit For
        ExtractNumberBefore = ch & ExtractNumberBefore
    Next i
End Function

Private Function DigitsOnly(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If Not ch Like "#" Then
            DigitsOnly = ""
            Exit Function
        End If
        DigitsOnly = DigitsOnly & ch
    Next i
End Function

Private Function ToAchievementCode(ByVal rawValue As String) As String
    rawValue = NormalizeLifeText(rawValue)
    Select Case rawValue
        Case "1", "’B¬": ToAchievementCode = "1"
        Case "2", "ˆê•”": ToAchievementCode = "2"
        Case "3", "–¢’B": ToAchievementCode = "3"
    End Select
End Function

Private Function NormalizeLifeText(ByVal rawValue As String) As String
    Dim i As Long
    NormalizeLifeText = Trim$(CStr(rawValue))
    NormalizeLifeText = Replace(NormalizeLifeText, "@", "")
    NormalizeLifeText = Replace(NormalizeLifeText, " ", "")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚P", "1")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚Q", "2")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚R", "3")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚S", "4")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚T", "5")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚U", "6")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚V", "7")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚W", "8")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚X", "9")
    NormalizeLifeText = Replace(NormalizeLifeText, "‚O", "0")
    For i = 1 To 2
        NormalizeLifeText = Replace(NormalizeLifeText, vbCr, "")
        NormalizeLifeText = Replace(NormalizeLifeText, vbLf, "")
        NormalizeLifeText = Replace(NormalizeLifeText, vbTab, "")
    Next i
End Function

Private Sub PutCsvValue(ByVal headers As Variant, ByRef values() As String, ByVal fieldName As String, ByVal fieldValue As String)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If CStr(headers(i)) = fieldName Then
            values(i) = fieldValue
            Exit Sub
        End If
    Next i
End Sub

Private Function CsvLine(ByRef values() As String) As String
    Dim i As Long
    Dim parts() As String
    ReDim parts(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        parts(i) = CsvEscape(values(i))
    Next i
    CsvLine = Join(parts, ",")
End Function

Private Function CsvEscape(ByVal s As String) As String
    If InStr(s, """") > 0 Then s = Replace(s, """", """""")
    If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Or InStr(s, """") > 0 Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function

Private Sub WriteUtf8BomText(ByVal filePath As String, ByVal textBody As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText textBody
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Sub EnsureFolderPath(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then EnsureFolderPath parentPath
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function IsMissingOrEmpty(ByVal v As Variant) As Boolean
    On Error GoTo EH
    If IsEmpty(v) Then IsMissingOrEmpty = True: Exit Function
    If UBound(v) < LBound(v) Then IsMissingOrEmpty = True
    Exit Function
EH:
    IsMissingOrEmpty = True
End Function
