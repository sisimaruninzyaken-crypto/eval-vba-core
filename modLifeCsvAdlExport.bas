Attribute VB_Name = "modLifeCsvAdlExport"
Option Explicit

Private Const ADL_CSV_DOC_TYPE As String = "ADL_MAINTENANCE_ADDITION_2024"
Private Const ADL_CSV_SERVICE_CODE As String = "15"

Private Const LIFE_KEY_VERSION_ADL As String = "VERSION_ADL"
Private Const LIFE_KEY_STATUS_DEFAULT As String = "STATUS_DEFAULT"

Private Const BI_CODE_ZERO As String = "0"
Private Const BI_CODE_ONE As String = "1"
Private Const BI_CODE_TWO As String = "2"
Private Const BI_CODE_THREE As String = "3"

Private Const FIRST_MONTH_DEFAULT As String = "0"
Private Const SIXTH_MONTH_DEFAULT As String = "1"
Public Sub ExportLifeAdlCsvFromActiveEval()
    Dim owner As Object
    Set owner = FindActiveEvalForm()
    If owner Is Nothing Then
        MsgBox "frmEval is not open.", vbExclamation
        Exit Sub
    End If

    ExportLifeAdlCsvFromOwner owner
End Sub

Public Sub ExportLifeAdlCsvFromOwner(ByVal owner As Object)
    On Error GoTo EH

    Dim tempWb As Workbook
    Dim tempWs As Worksheet
    Dim record As Object
    Dim outputPath As String
    Dim headerRow As Variant
    Dim dataRow As Variant
    Dim csvLines(0 To 2) As String

    Set tempWs = CreateFilledLifeFuncCheckSheet(owner, tempWb)
    Set record = BuildAdlCsvRecord(tempWs, owner)
    outputPath = BuildAdlCsvOutputPath(record)

    headerRow = BuildAdlCsvHeaderRow()
    dataRow = BuildAdlCsvDataRow(record)

    csvLines(0) = ADL_CSV_DOC_TYPE
    csvLines(1) = modLifeCsvWriter.LifeCsv_JoinRow(headerRow)
    csvLines(2) = modLifeCsvWriter.LifeCsv_JoinRow(dataRow)

    modLifeCsvWriter.LifeCsv_WriteUtf8BomLines outputPath, csvLines

    If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False

    MsgBox "ADL CSV saved:" & vbCrLf & outputPath, vbInformation
    Exit Sub

EH:
    On Error Resume Next
    If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
    MsgBox "ADL CSV export error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Public Sub ExportLifeAdlBatchCsvFromCandidates(ByVal candidates As Collection)
    On Error GoTo EH

    Dim csvLines As Collection
    Dim records As Collection
    Dim failures As Collection
    Dim item As Object
    Dim record As Object
    Dim outputPath As String
    Dim i As Long
    Dim lineArray() As String

    If candidates Is Nothing Or candidates.count = 0 Then
        MsgBox "No users are selected.", vbExclamation
        Exit Sub
    End If

    Set records = New Collection
    Set failures = New Collection

    For Each item In candidates
        On Error GoTo ItemEH
        Set record = BuildAdlCsvRecordFromBatchCandidate(item)
        If Not record Is Nothing Then records.Add record
        On Error GoTo EH
        GoTo ContinueItem
ItemEH:
        failures.Add BatchCandidateFailureText(item, Err.Description)
        Err.Clear
        On Error GoTo EH
ContinueItem:
    Next item

    If records.count = 0 Then
        MsgBox "No ADL CSV records were created." & BuildBatchFailureMessage(failures), vbExclamation
        Exit Sub
    End If

    Set csvLines = New Collection
    csvLines.Add ADL_CSV_DOC_TYPE
    csvLines.Add modLifeCsvWriter.LifeCsv_JoinRow(BuildAdlCsvHeaderRow())
    For i = 1 To records.count
        csvLines.Add modLifeCsvWriter.LifeCsv_JoinRow(BuildAdlCsvDataRow(records(i)))
    Next i

    outputPath = BuildAdlCsvOutputPath(records(1))
    ReDim lineArray(0 To csvLines.count - 1)
    For i = 1 To csvLines.count
        lineArray(i - 1) = CStr(csvLines(i))
    Next i
    modLifeCsvWriter.LifeCsv_WriteUtf8BomLines outputPath, lineArray

    MsgBox "ADL CSV saved:" & vbCrLf & outputPath & vbCrLf & _
           "Records: " & CStr(records.count) & BuildBatchFailureMessage(failures), vbInformation
    Exit Sub

EH:
    MsgBox "ADL batch CSV export error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Function BuildAdlCsvRecordFromBatchCandidate(ByVal item As Object) As Object
    On Error GoTo EH

    Dim wsHistory As Worksheet
    Dim historyRow As Long
    Dim tempWb As Workbook
    Dim tempWs As Worksheet
    Dim eligibility As Object

    If item Is Nothing Then Err.Raise vbObjectError + 510, , "Candidate is empty."
    If Not item.exists("SheetName") Then Err.Raise vbObjectError + 511, , "Candidate sheet name is missing."
    If Not item.exists("HistoryRow") Then Err.Raise vbObjectError + 512, , "Candidate history row is missing."

    Set wsHistory = ThisWorkbook.Worksheets(CStr(item("SheetName")))
    historyRow = CLng(item("HistoryRow"))
    If historyRow < 2 Then Err.Raise vbObjectError + 513, , "Candidate latest row is invalid."

    Set tempWs = CreateFilledLifeFuncCheckSheetFromHistoryRow(wsHistory, historyRow, tempWb)
    Set eligibility = modLifeAdlEligibility.BuildAdlEligibilityFromHistoryRow(wsHistory, historyRow)
    Set BuildAdlCsvRecordFromBatchCandidate = BuildAdlCsvRecordCore( _
        tempWs, _
        HistoryText(wsHistory, historyRow, Array("InsuredNo")), _
        HistoryText(wsHistory, historyRow, Array("InsurerNo")), _
        HistoryText(wsHistory, historyRow, Array("ExternalSystemKey")), _
        eligibility)

CleanUp:
    On Error Resume Next
    If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

EH:
    Dim errText As String
    errText = Err.Description
    Resume CleanUpRaise

CleanUpRaise:
    On Error Resume Next
    If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 514, , errText
End Function

Private Function CreateFilledLifeFuncCheckSheetFromHistoryRow(ByVal wsHistory As Worksheet, _
                                                              ByVal historyRow As Long, _
                                                              ByRef tempWb As Workbook) As Worksheet
    Dim templateWs As Worksheet
    Set templateWs = ThisWorkbook.Worksheets(LifeFuncTemplateSheetName())

    templateWs.Copy
    Set tempWb = ActiveWorkbook
    Set CreateFilledLifeFuncCheckSheetFromHistoryRow = tempWb.Worksheets(1)

    WriteBatchBasicInfoToLifeSheet CreateFilledLifeFuncCheckSheetFromHistoryRow, wsHistory, historyRow
    WriteBatchBiToLifeSheet CreateFilledLifeFuncCheckSheetFromHistoryRow, wsHistory, historyRow
End Function

Private Sub WriteBatchBasicInfoToLifeSheet(ByVal ws As Worksheet, ByVal wsHistory As Worksheet, ByVal historyRow As Long)
    WriteMergedValue ws, "E3:N3", HistoryText(wsHistory, historyRow, Array("Basic.Name", ChrW$(&H6C0F) & ChrW$(&H540D)))
    WriteMergedValue ws, "R3:W3", HistoryText(wsHistory, historyRow, Array("Basic.BirthDate", ChrW$(&H751F) & ChrW$(&H5E74) & ChrW$(&H6708) & ChrW$(&H65E5)))
    WriteMergedValue ws, "Y3:Z3", HistoryText(wsHistory, historyRow, Array("Basic.Sex", ChrW$(&H6027) & ChrW$(&H5225)))
    WriteMergedValue ws, "E4:R4", BuildBatchEvalDateText(HistoryValue(wsHistory, historyRow, Array("Basic.EvalDate", ChrW$(&H8A55) & ChrW$(&H4FA1) & ChrW$(&H65E5))))
    WriteMergedValue ws, "V4:Z4", HistoryText(wsHistory, historyRow, Array("Basic.CareLevel", ChrW$(&H8981) & ChrW$(&H4ECB) & ChrW$(&H8B77) & ChrW$(&H5EA6)))
    WriteMergedValue ws, "I6:Z6", HistoryText(wsHistory, historyRow, Array("Basic.LifeStatus", ChrW$(&H969C) & ChrW$(&H5BB3) & ChrW$(&H9AD8) & ChrW$(&H9F62) & ChrW$(&H8005) & ChrW$(&H306E) & ChrW$(&H65E5) & ChrW$(&H5E38) & ChrW$(&H751F) & ChrW$(&H6D3B) & ChrW$(&H81EA) & ChrW$(&H7ACB) & ChrW$(&H5EA6), ChrW$(&H9AD8) & ChrW$(&H9F62) & ChrW$(&H8005) & ChrW$(&H306E) & ChrW$(&H65E5) & ChrW$(&H5E38) & ChrW$(&H751F) & ChrW$(&H6D3B) & ChrW$(&H81EA) & ChrW$(&H7ACB) & ChrW$(&H5EA6)))
    WriteMergedValue ws, "I7:Z7", HistoryText(wsHistory, historyRow, Array("Basic.DementiaADL", ChrW$(&H8A8D) & ChrW$(&H77E5) & ChrW$(&H75C7) & ChrW$(&H9AD8) & ChrW$(&H9F62) & ChrW$(&H8005) & ChrW$(&H306E) & ChrW$(&H65E5) & ChrW$(&H5E38) & ChrW$(&H751F) & ChrW$(&H6D3B) & ChrW$(&H81EA) & ChrW$(&H7ACB) & ChrW$(&H5EA6)))
End Sub

Private Sub WriteBatchBiToLifeSheet(ByVal ws As Worksheet, ByVal wsHistory As Worksheet, ByVal historyRow As Long)
    Dim adlText As String
    Dim biRows As Variant
    Dim i As Long
    Dim biKey As String
    Dim scoreText As String

    adlText = HistoryText(wsHistory, historyRow, Array("IO_ADL"))
    biRows = BuildBiSourceMap()
    For i = LBound(biRows) To UBound(biRows)
        biKey = CStr(biRows(i)(0))
        scoreText = ExtractIoPairValue(adlText, biKey)
        If LenB(Trim$(scoreText)) > 0 Then WriteMergedValue ws, CStr(biRows(i)(2)), scoreText
    Next i
End Sub

Private Function BuildBatchEvalDateText(ByVal rawValue As Variant) As String
    If IsDate(rawValue) Then
        BuildBatchEvalDateText = Format$(CDate(rawValue), "yyyy/mm/dd") & " 13:00" & ChrW$(65374) & "15:00"
    ElseIf LenB(Trim$(CStr(rawValue))) > 0 Then
        BuildBatchEvalDateText = Trim$(CStr(rawValue)) & " 13:00" & ChrW$(65374) & "15:00"
    End If
End Function

Private Function ExtractIoPairValue(ByVal ioText As String, ByVal keyName As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim p As String
    Dim pos As Long

    If LenB(Trim$(ioText)) = 0 Then Exit Function
    parts = Split(ioText, "|")
    For i = LBound(parts) To UBound(parts)
        p = CStr(parts(i))
        pos = InStr(1, p, "=", vbBinaryCompare)
        If pos > 0 Then
            If StrComp(Left$(p, pos - 1), keyName, vbTextCompare) = 0 Then
                ExtractIoPairValue = Mid$(p, pos + 1)
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub WriteMergedValue(ByVal ws As Worksheet, ByVal addressText As String, ByVal textValue As String)
    ws.Range(addressText).Cells(1, 1).value = textValue
End Sub

Private Function HistoryText(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headers As Variant) As String
    HistoryText = Trim$(CStr(HistoryValue(ws, rowNo, headers)))
End Function

Private Function HistoryValue(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headers As Variant) As Variant
    Dim i As Long
    Dim colNo As Long
    Dim rawValue As Variant

    For i = LBound(headers) To UBound(headers)
        colNo = FindHistoryHeaderCol(ws, CStr(headers(i)))
        If colNo > 0 Then
            rawValue = ws.Cells(rowNo, colNo).value
            If LenB(Trim$(CStr(rawValue))) > 0 Then
                HistoryValue = rawValue
                Exit Function
            End If
        End If
    Next i
    HistoryValue = vbNullString
End Function

Private Function FindHistoryHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colNo As Long

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For colNo = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, colNo).value)), headerName, vbTextCompare) = 0 Then
            FindHistoryHeaderCol = colNo
            Exit Function
        End If
    Next colNo
End Function
Private Function BatchCandidateFailureText(ByVal item As Object, ByVal reasonText As String) As String
    Dim nameText As String
    If Not item Is Nothing Then
        On Error Resume Next
        If item.exists("Name") Then nameText = CStr(item("Name"))
        On Error GoTo 0
    End If
    If LenB(Trim$(nameText)) = 0 Then nameText = "(unknown)"
    BatchCandidateFailureText = nameText & ": " & reasonText
End Function

Private Function BuildBatchFailureMessage(ByVal failures As Collection) As String
    Dim i As Long
    Dim msg As String
    If failures Is Nothing Then Exit Function
    If failures.count = 0 Then Exit Function

    msg = vbCrLf & vbCrLf & "Skipped:" & vbCrLf
    For i = 1 To failures.count
        msg = msg & "- " & CStr(failures(i)) & vbCrLf
    Next i
    BuildBatchFailureMessage = msg
End Function
Public Function BuildAdlCsvHeaderRow() As Variant
    BuildAdlCsvHeaderRow = Array( _
        "care_facility_id", _
        "service_code", _
        "insurer_no", _
        "insured_no", _
        "external_system_management_number", _
        "care_level", _
        "impaired_elderly_independence_degree", _
        "dementia_elderly_independence_degree", _
        "evaluate_date", _
        "status", _
        "barthel_index_meal", _
        "barthel_transfer", _
        "barthel_index_personal_hygiene_and_adjustment", _
        "barthel_index_toilet_activity", _
        "barthel_index_bathing", _
        "barthel_index_flat_ground_walking", _
        "barthel_index_stair_movement", _
        "barthel_index_changing_clothes", _
        "barthel_index_defecation_manage", _
        "barthel_index_urination_manage", _
        "first_month", _
        "sixth_month", _
        "version")
End Function

Public Function BuildAdlCsvDataRow(ByVal record As Object) As Variant
    BuildAdlCsvDataRow = Array( _
        Nz(record, "care_facility_id"), _
        Nz(record, "service_code"), _
        Nz(record, "insurer_no"), _
        Nz(record, "insured_no"), _
        Nz(record, "external_system_management_number"), _
        Nz(record, "care_level"), _
        Nz(record, "impaired_elderly_independence_degree"), _
        Nz(record, "dementia_elderly_independence_degree"), _
        Nz(record, "evaluate_date"), _
        Nz(record, "status"), _
        Nz(record, "barthel_index_meal"), _
        Nz(record, "barthel_transfer"), _
        Nz(record, "barthel_index_personal_hygiene_and_adjustment"), _
        Nz(record, "barthel_index_toilet_activity"), _
        Nz(record, "barthel_index_bathing"), _
        Nz(record, "barthel_index_flat_ground_walking"), _
        Nz(record, "barthel_index_stair_movement"), _
        Nz(record, "barthel_index_changing_clothes"), _
        Nz(record, "barthel_index_defecation_manage"), _
        Nz(record, "barthel_index_urination_manage"), _
        Nz(record, "first_month"), _
        Nz(record, "sixth_month"), _
        Nz(record, "version"))
End Function
Public Function BuildAdlCsvRecord(ByVal ws As Worksheet, ByVal owner As Object) As Object
    Set BuildAdlCsvRecord = BuildAdlCsvRecordCore( _
        ws, _
        GetOwnerLifeValue(owner, "InsuredNo"), _
        GetOwnerLifeValue(owner, "InsurerNo"), _
        GetOwnerLifeValue(owner, "ExternalSystemKey"), _
        modLifeAdlEligibility.BuildAdlEligibility(owner))
End Function

Private Function BuildAdlCsvRecordCore(ByVal ws As Worksheet, _
                                       ByVal insuredNo As String, _
                                       ByVal insurerNo As String, _
                                       ByVal externalSystemKey As String, _
                                       ByVal eligibility As Object) As Object
    Dim d As Object
    Dim facilityName As String
    Dim facilityNo As String
    Dim facilityAddress As String
    Dim facilityPhone As String
    Dim biItems As Variant
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    modAppConfig.LoadFacilitySettings facilityName, facilityNo, facilityAddress, facilityPhone

    d("care_facility_id") = Trim$(facilityNo)
    d("service_code") = ADL_CSV_SERVICE_CODE
    d("insured_no") = Trim$(insuredNo)
    d("insurer_no") = Trim$(insurerNo)
    d("external_system_management_number") = Trim$(externalSystemKey)
    d("care_level") = MapCareLevelToCsvCode(GetMergedText(ws, "V4:Z4"))
    d("impaired_elderly_independence_degree") = MapImpairedElderlyDegreeToCsvCode(GetMergedText(ws, "I6:Z6"))
    d("dementia_elderly_independence_degree") = MapDementiaDegreeToCsvCode(GetMergedText(ws, "I7:Z7"))
    d("evaluate_date") = NormalizeSheetDateToYmd(GetMergedText(ws, "E4:R4"))
    d("status") = modLifeSettings.GetLifeSetting(LIFE_KEY_STATUS_DEFAULT)
    d("version") = modLifeSettings.GetLifeSetting(LIFE_KEY_VERSION_ADL)
    d("first_month") = GetEligibilityFlagValue(eligibility, "FirstMonthFlag")
    d("sixth_month") = GetEligibilityFlagValue(eligibility, "SixthMonthFlag")

    biItems = BuildBiSourceMap()
    For i = LBound(biItems) To UBound(biItems)
        d(CStr(biItems(i)(1))) = MapBiDisplayToCsvCode(CStr(biItems(i)(0)), GetMergedText(ws, CStr(biItems(i)(2))))
    Next i

    Set BuildAdlCsvRecordCore = d
End Function
Private Function BuildBiSourceMap() As Variant
    BuildBiSourceMap = Array( _
        Array("BI_0", "barthel_index_meal", "G13:N14"), _
        Array("BI_1", "barthel_transfer", "G15:N16"), _
        Array("BI_2", "barthel_index_personal_hygiene_and_adjustment", "G17:N18"), _
        Array("BI_3", "barthel_index_toilet_activity", "G19:N20"), _
        Array("BI_4", "barthel_index_bathing", "G21:N22"), _
        Array("BI_5", "barthel_index_flat_ground_walking", "G23:N24"), _
        Array("BI_6", "barthel_index_stair_movement", "G25:N26"), _
        Array("BI_7", "barthel_index_changing_clothes", "G27:N28"), _
        Array("BI_8", "barthel_index_defecation_manage", "G29:N30"), _
        Array("BI_9", "barthel_index_urination_manage", "G31:N32"))
End Function

Private Function CreateFilledLifeFuncCheckSheet(ByVal owner As Object, ByRef tempWb As Workbook) As Worksheet
    Dim templateWs As Worksheet
    Set templateWs = ThisWorkbook.Worksheets(LifeFuncTemplateSheetName())

    templateWs.Copy
    Set tempWb = ActiveWorkbook
    Set CreateFilledLifeFuncCheckSheet = tempWb.Worksheets(1)

    modLifeFuncCheckSheetOutput.WriteLifeFuncCheckSheet CreateFilledLifeFuncCheckSheet, owner
End Function

Public Function BuildAdlCsvOutputPath(ByVal record As Object) As String
    Dim rootDir As String
    Dim facilityDir As String
    Dim facilityNo As String
    Dim evalMonth As String
    Dim stamp As String
    Dim fileName As String

    facilityNo = Nz(record, "care_facility_id")
    If LenB(facilityNo) = 0 Then facilityNo = "unknown"

    evalMonth = Left$(Nz(record, "evaluate_date"), 6)
    If LenB(evalMonth) = 0 Then evalMonth = Format$(Date, "yyyymm")

    stamp = Format$(Now, "yyyymmddhhnnss")

    rootDir = ThisWorkbook.path & Application.PathSeparator & "CSV"
    modLifeCsvWriter.LifeCsv_EnsureFolderExists rootDir

    facilityDir = rootDir & Application.PathSeparator & facilityNo
    modLifeCsvWriter.LifeCsv_EnsureFolderExists facilityDir

    fileName = "26_ADL_MAINTENANCE_ADDITION_2024_" & facilityNo & "_" & evalMonth & "_" & stamp & ".csv"
    BuildAdlCsvOutputPath = facilityDir & Application.PathSeparator & fileName
End Function

Private Function GetMergedText(ByVal ws As Worksheet, ByVal addressText As String) As String
    Dim rng As Range
    Set rng = ws.Range(addressText)
    GetMergedText = Trim$(CStr(rng.Cells(1, 1).value))
End Function

Private Function GetOwnerLifeValue(ByVal owner As Object, ByVal fieldKey As String) As String
    On Error GoTo EH
    GetOwnerLifeValue = Trim$(CStr(CallByName(owner, "GetLifeLinkFieldValue", VbMethod, fieldKey)))
    Exit Function
EH:
    Err.Clear
End Function

Private Function MapBiDisplayToCsvCode(ByVal biKey As String, ByVal displayText As String) As String
    Dim scoreText As String
    scoreText = ExtractScoreFromBiDisplay(displayText)

    Select Case UCase$(Trim$(biKey))
        Case "BI_0"
            MapBiDisplayToCsvCode = MapBiCode_BI0(scoreText)
        Case "BI_1"
            MapBiDisplayToCsvCode = MapBiCode_BI1(scoreText)
        Case "BI_2"
            MapBiDisplayToCsvCode = MapBiCode_BI2(scoreText)
        Case "BI_3"
            MapBiDisplayToCsvCode = MapBiCode_BI3(scoreText)
        Case "BI_4"
            MapBiDisplayToCsvCode = MapBiCode_BI4(scoreText)
        Case "BI_5"
            MapBiDisplayToCsvCode = MapBiCode_BI5(scoreText)
        Case "BI_6"
            MapBiDisplayToCsvCode = MapBiCode_BI6(scoreText)
        Case "BI_7"
            MapBiDisplayToCsvCode = MapBiCode_BI7(scoreText)
        Case "BI_8"
            MapBiDisplayToCsvCode = MapBiCode_BI8(scoreText)
        Case "BI_9"
            MapBiDisplayToCsvCode = MapBiCode_BI9(scoreText)
    End Select
End Function

Private Function MapBiCode_BI0(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI0 = BI_CODE_ZERO
        Case "5": MapBiCode_BI0 = BI_CODE_ONE
        Case "10": MapBiCode_BI0 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI1(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI1 = BI_CODE_ZERO
        Case "5": MapBiCode_BI1 = BI_CODE_ONE
        Case "10": MapBiCode_BI1 = BI_CODE_TWO
        Case "15": MapBiCode_BI1 = BI_CODE_THREE
    End Select
End Function

Private Function MapBiCode_BI2(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI2 = BI_CODE_ZERO
        Case "5": MapBiCode_BI2 = BI_CODE_ONE
    End Select
End Function

Private Function MapBiCode_BI3(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI3 = BI_CODE_ZERO
        Case "5": MapBiCode_BI3 = BI_CODE_ONE
        Case "10": MapBiCode_BI3 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI4(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI4 = BI_CODE_ZERO
        Case "5": MapBiCode_BI4 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI5(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI5 = BI_CODE_ZERO
        Case "5": MapBiCode_BI5 = BI_CODE_ONE
        Case "10": MapBiCode_BI5 = BI_CODE_TWO
        Case "15": MapBiCode_BI5 = BI_CODE_THREE
    End Select
End Function

Private Function MapBiCode_BI6(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI6 = BI_CODE_ZERO
        Case "5": MapBiCode_BI6 = BI_CODE_ONE
        Case "10": MapBiCode_BI6 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI7(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI7 = BI_CODE_ZERO
        Case "5": MapBiCode_BI7 = BI_CODE_ONE
        Case "10": MapBiCode_BI7 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI8(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI8 = BI_CODE_ZERO
        Case "5": MapBiCode_BI8 = BI_CODE_ONE
        Case "10": MapBiCode_BI8 = BI_CODE_TWO
    End Select
End Function

Private Function MapBiCode_BI9(ByVal scoreText As String) As String
    Select Case scoreText
        Case "0": MapBiCode_BI9 = BI_CODE_ZERO
        Case "5": MapBiCode_BI9 = BI_CODE_ONE
        Case "10": MapBiCode_BI9 = BI_CODE_TWO
    End Select
End Function

Private Function ExtractScoreFromBiDisplay(ByVal displayText As String) As String
    Dim openPos As Long
    Dim closePos As Long
    Dim scoreText As String

    openPos = InStrRev(displayText, ChrW(&HFF08))
    If openPos = 0 Then openPos = InStrRev(displayText, "(")
    closePos = InStrRev(displayText, ChrW(&HFF09))
    If closePos = 0 Then closePos = InStrRev(displayText, ")")

    If openPos > 0 And closePos > openPos Then
        scoreText = Mid$(displayText, openPos + 1, closePos - openPos - 1)
    Else
        scoreText = ExtractTrailingDigits(displayText)
    End If

    ExtractScoreFromBiDisplay = Trim$(scoreText)
End Function

Private Function ExtractTrailingDigits(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        ElseIf LenB(digits) > 0 Then
            ExtractTrailingDigits = digits
            digits = vbNullString
        End If
    Next i

    If LenB(digits) > 0 Then ExtractTrailingDigits = digits
End Function

Private Function MapCareLevelToCsvCode(ByVal src As String) As String
    Select Case NormalizeSymbolText(src)
        Case NormalizeSymbolText(BuildWordNeedSupport(1))
            MapCareLevelToCsvCode = "12"
        Case NormalizeSymbolText(BuildWordNeedSupport(2))
            MapCareLevelToCsvCode = "13"
        Case NormalizeSymbolText(BuildWordNeedCare(1))
            MapCareLevelToCsvCode = "21"
        Case NormalizeSymbolText(BuildWordNeedCare(2))
            MapCareLevelToCsvCode = "22"
        Case NormalizeSymbolText(BuildWordNeedCare(3))
            MapCareLevelToCsvCode = "23"
        Case NormalizeSymbolText(BuildWordNeedCare(4))
            MapCareLevelToCsvCode = "24"
        Case NormalizeSymbolText(BuildWordNeedCare(5))
            MapCareLevelToCsvCode = "25"
    End Select
End Function

Private Function MapImpairedElderlyDegreeToCsvCode(ByVal src As String) As String
    Select Case UCase$(NormalizeSymbolText(src))
        Case "J1"
            MapImpairedElderlyDegreeToCsvCode = "1"
        Case "J2"
            MapImpairedElderlyDegreeToCsvCode = "2"
        Case "A1"
            MapImpairedElderlyDegreeToCsvCode = "3"
        Case "A2"
            MapImpairedElderlyDegreeToCsvCode = "4"
        Case "B1"
            MapImpairedElderlyDegreeToCsvCode = "5"
        Case "B2"
            MapImpairedElderlyDegreeToCsvCode = "6"
        Case "C1"
            MapImpairedElderlyDegreeToCsvCode = "7"
        Case "C2"
            MapImpairedElderlyDegreeToCsvCode = "8"
        Case UCase$(NormalizeSymbolText(BuildWordIndependent()))
            MapImpairedElderlyDegreeToCsvCode = BI_CODE_ZERO
    End Select
End Function

Private Function MapDementiaDegreeToCsvCode(ByVal src As String) As String
    Select Case UCase$(NormalizeSymbolText(ConvertRomanNumerals(src)))
        Case UCase$(NormalizeSymbolText(BuildWordIndependent()))
            MapDementiaDegreeToCsvCode = "1"
        Case "I"
            MapDementiaDegreeToCsvCode = "2"
        Case "IIA"
            MapDementiaDegreeToCsvCode = "3"
        Case "IIB"
            MapDementiaDegreeToCsvCode = "4"
        Case "IIIA"
            MapDementiaDegreeToCsvCode = "5"
        Case "IIIB"
            MapDementiaDegreeToCsvCode = "6"
        Case "IV"
            MapDementiaDegreeToCsvCode = "7"
        Case "M"
            MapDementiaDegreeToCsvCode = "8"
    End Select
End Function

Private Function ConvertRomanNumerals(ByVal src As String) As String
    Dim s As String
    s = src
    s = Replace$(s, ChrW(&H2160), "I")
    s = Replace$(s, ChrW(&H2161), "II")
    s = Replace$(s, ChrW(&H2162), "III")
    s = Replace$(s, ChrW(&H2163), "IV")
    ConvertRomanNumerals = s
End Function

Private Function NormalizeSheetDateToYmd(ByVal src As String) As String
    Dim s As String
    Dim y As Long
    Dim m As Long
    Dim d As Long

    s = Trim$(src)
    If LenB(s) = 0 Then Exit Function

    If IsDate(s) Then
        NormalizeSheetDateToYmd = Format$(CDate(s), "yyyymmdd")
        Exit Function
    End If

    s = Replace$(s, vbTab, " ")
    s = Replace$(s, ChrW(&H3000), " ")
    If InStr(s, " ") > 0 Then s = Split(s, " ")(0)
    If InStr(s, "(") > 0 Then s = Left$(s, InStr(s, "(") - 1)
    If InStr(s, ChrW(&HFF08)) > 0 Then s = Left$(s, InStr(s, ChrW(&HFF08)) - 1)
    If InStr(s, ChrW(&HFF5E)) > 0 Then s = Left$(s, InStr(s, ChrW(&HFF5E)) - 1)

    If TryParseSlashDate(s, y, m, d) Then
        NormalizeSheetDateToYmd = Format$(DateSerial(y, m, d), "yyyymmdd")
        Exit Function
    End If

    If TryParseWarekiDate(s, y, m, d) Then
        NormalizeSheetDateToYmd = Format$(DateSerial(y, m, d), "yyyymmdd")
    End If
End Function

Private Function TryParseSlashDate(ByVal src As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim parts As Variant
    Dim normalized As String

    normalized = Replace$(src, ".", "/")
    normalized = Replace$(normalized, "-", "/")

    parts = Split(normalized, "/")
    If UBound(parts) <> 2 Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If Not IsNumeric(parts(2)) Then Exit Function

    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))
    TryParseSlashDate = (y > 0 And m > 0 And d > 0)
End Function

Private Function TryParseWarekiDate(ByVal src As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim eraBase As Long
    Dim eraPos As Long
    Dim yearText As String
    Dim monthText As String
    Dim dayText As String
    Dim yearPos As Long
    Dim monthPos As Long

    If InStr(src, BuildWordReiwa()) > 0 Then
        eraBase = 2018
        eraPos = InStr(src, BuildWordReiwa()) + Len(BuildWordReiwa())
    ElseIf InStr(src, BuildWordHeisei()) > 0 Then
        eraBase = 1988
        eraPos = InStr(src, BuildWordHeisei()) + Len(BuildWordHeisei())
    ElseIf InStr(src, BuildWordShowa()) > 0 Then
        eraBase = 1925
        eraPos = InStr(src, BuildWordShowa()) + Len(BuildWordShowa())
    Else
        Exit Function
    End If

    yearPos = InStr(src, BuildWordYear())
    monthPos = InStr(src, BuildWordMonth())
    If yearPos = 0 Or monthPos = 0 Then Exit Function

    yearText = ExtractDigitsUntilToken(Mid$(src, eraPos), BuildWordYear())
    monthText = ExtractDigitsUntilToken(Mid$(src, yearPos + Len(BuildWordYear())), BuildWordMonth())
    dayText = ExtractDigitsUntilToken(Mid$(src, monthPos + Len(BuildWordMonth())), BuildWordDay())

    If LenB(yearText) = 0 Or LenB(monthText) = 0 Or LenB(dayText) = 0 Then Exit Function

    y = eraBase + CLng(yearText)
    m = CLng(monthText)
    d = CLng(dayText)
    TryParseWarekiDate = True
End Function

Private Function ExtractDigitsUntilToken(ByVal src As String, ByVal endToken As String) As String
    Dim stopPos As Long
    Dim i As Long
    Dim ch As String

    stopPos = InStr(src, endToken)
    If stopPos = 0 Then Exit Function

    For i = 1 To stopPos - 1
        ch = Mid$(src, i, 1)
        If ch Like "#" Then ExtractDigitsUntilToken = ExtractDigitsUntilToken & ch
    Next i
End Function

Private Function NormalizeSymbolText(ByVal src As String) As String
    Dim s As String
    s = Trim$(src)
    s = Replace$(s, " ", vbNullString)
    s = Replace$(s, ChrW(&H3000), vbNullString)
    NormalizeSymbolText = s
End Function

Private Function FindActiveEvalForm() As Object
    Dim i As Long
    On Error Resume Next
    For i = 0 To VBA.UserForms.count - 1
        If StrComp(VBA.UserForms(i).name, "frmEval", vbTextCompare) = 0 Then
            Set FindActiveEvalForm = VBA.UserForms(i)
            Exit Function
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Function

Private Function Nz(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.exists(key) Then Exit Function
    Nz = Trim$(CStr(d(key)))
End Function

Private Function LifeFuncTemplateSheetName() As String
    LifeFuncTemplateSheetName = BuildWordLife() & BuildWordFunction() & BuildWordCheckSheet()
End Function

Private Function BuildWordLife() As String
    BuildWordLife = ChrW(&H751F) & ChrW(&H6D3B)
End Function

Private Function BuildWordFunction() As String
    BuildWordFunction = ChrW(&H6A5F) & ChrW(&H80FD)
End Function

Private Function BuildWordCheckSheet() As String
    BuildWordCheckSheet = ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8)
End Function

Private Function BuildWordNeedSupport(ByVal levelNum As Long) As String
    BuildWordNeedSupport = ChrW(&H8981) & ChrW(&H652F) & ChrW(&H63F4) & CStr(levelNum)
End Function

Private Function BuildWordNeedCare(ByVal levelNum As Long) As String
    BuildWordNeedCare = ChrW(&H8981) & ChrW(&H4ECB) & ChrW(&H8B77) & CStr(levelNum)
End Function

Private Function BuildWordIndependent() As String
    BuildWordIndependent = ChrW(&H81EA) & ChrW(&H7ACB)
End Function

Private Function BuildWordReiwa() As String
    BuildWordReiwa = ChrW(&H4EE4) & ChrW(&H548C)
End Function

Private Function BuildWordHeisei() As String
    BuildWordHeisei = ChrW(&H5E73) & ChrW(&H6210)
End Function

Private Function BuildWordShowa() As String
    BuildWordShowa = ChrW(&H662D) & ChrW(&H548C)
End Function

Private Function BuildWordYear() As String
    BuildWordYear = ChrW(&H5E74)
End Function

Private Function BuildWordMonth() As String
    BuildWordMonth = ChrW(&H6708)
End Function

Private Function BuildWordDay() As String
    BuildWordDay = ChrW(&H65E5)
End Function



Private Function GetEligibilityFlagValue(ByVal eligibility As Object, ByVal keyName As String) As String
    If eligibility Is Nothing Then
        GetEligibilityFlagValue = "0"
        Exit Function
    End If
    If Not eligibility.exists(keyName) Then
        GetEligibilityFlagValue = "0"
        Exit Function
    End If
    GetEligibilityFlagValue = Trim$(CStr(eligibility(keyName)))
    If LenB(GetEligibilityFlagValue) = 0 Then GetEligibilityFlagValue = "0"
End Function


