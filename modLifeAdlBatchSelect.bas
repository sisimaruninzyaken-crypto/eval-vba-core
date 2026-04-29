Attribute VB_Name = "modLifeAdlBatchSelect"
Option Explicit

Private Const EVAL_INDEX_SHEET_NAME As String = "EvalIndex"
Private Const HDR_USER_ID As String = "UserID"
Private Const HDR_NAME As String = "Name"
Private Const HDR_SHEET As String = "SheetName"

Public Sub ShowLifeAdlBatchSelect()
    frmLifeAdlBatchSelect.Show vbModal
End Sub

Public Function BuildLifeAdlBatchCandidates() As Collection
    Dim result As Collection
    Dim wsIndex As Worksheet
    Dim rowNo As Long
    Dim lastRow As Long
    Dim colName As Long

    Set result = New Collection
    If Not TryGetWorksheet(EVAL_INDEX_SHEET_NAME, wsIndex) Then
        Set BuildLifeAdlBatchCandidates = result
        Exit Function
    End If

    colName = FindHeaderCol(wsIndex, HDR_NAME)
    If colName = 0 Then
        Set BuildLifeAdlBatchCandidates = result
        Exit Function
    End If

    lastRow = wsIndex.Cells(wsIndex.rows.count, colName).End(xlUp).row
    For rowNo = 2 To lastRow
        If Len(Trim$(CStr(wsIndex.Cells(rowNo, colName).value))) > 0 Then
            result.Add BuildLifeAdlBatchCandidate(wsIndex, rowNo)
        End If
    Next rowNo

    Set BuildLifeAdlBatchCandidates = result
End Function

Public Function LifeAdlBatchShouldSelect(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case modLifeAdlEligibility.ADL_ELIGIBILITY_STATUS_FIRST, modLifeAdlEligibility.ADL_ELIGIBILITY_STATUS_DUE
            LifeAdlBatchShouldSelect = True
    End Select
End Function

Private Function BuildLifeAdlBatchCandidate(ByVal wsIndex As Worksheet, ByVal indexRow As Long) As Object
    Dim item As Object
    Dim wsHistory As Worksheet
    Dim sheetName As String
    Dim latestRow As Long
    Dim eligibility As Object
    Dim evalDateValue As Variant

    Set item = CreateObject("Scripting.Dictionary")
    item("Selected") = False
    item("Name") = CellTextByHeader(wsIndex, indexRow, HDR_NAME)
    item("UserID") = CellTextByHeader(wsIndex, indexRow, HDR_USER_ID)
    item("SheetName") = CellTextByHeader(wsIndex, indexRow, HDR_SHEET)
    item("HistoryRow") = 0
    item("EvaluateDate") = vbNullString
    item("Status") = modLifeAdlEligibility.ADL_ELIGIBILITY_STATUS_INSUFFICIENT
    item("InsurerNo") = vbNullString
    item("InsuredNo") = vbNullString
    item("ExternalSystemKey") = vbNullString
    item("MissingReason") = vbNullString

    sheetName = Trim$(CStr(item("SheetName")))
    If LenB(sheetName) = 0 Or Not TryGetWorksheet(sheetName, wsHistory) Then
        item("MissingReason") = "History sheet was not found."
        Set BuildLifeAdlBatchCandidate = item
        Exit Function
    End If

    latestRow = FindLatestEvaluateRow(wsHistory)
    item("HistoryRow") = latestRow

    If latestRow < 2 Then
        item("MissingReason") = "Latest evaluate row was not found."
        Set BuildLifeAdlBatchCandidate = item
        Exit Function
    End If

    If LenB(Trim$(CStr(item("Name")))) = 0 Then
        item("Name") = FirstNonBlankHeaderText(wsHistory, latestRow, Array("Basic.Name", ChrW$(&H6C0F) & ChrW$(&H540D)))
    End If

    evalDateValue = FirstNonBlankHeaderValue(wsHistory, latestRow, Array("Basic.EvalDate", ChrW$(&H8A55) & ChrW$(&H4FA1) & ChrW$(&H65E5)))
    item("EvaluateDate") = FormatDisplayDate(evalDateValue)
    item("InsurerNo") = FirstNonBlankHeaderText(wsHistory, latestRow, Array("InsurerNo"))
    item("InsuredNo") = FirstNonBlankHeaderText(wsHistory, latestRow, Array("InsuredNo"))
    item("ExternalSystemKey") = FirstNonBlankHeaderText(wsHistory, latestRow, Array("ExternalSystemKey"))

    Set eligibility = modLifeAdlEligibility.BuildAdlEligibilityFromHistoryRow(wsHistory, latestRow)
    If Not eligibility Is Nothing Then
        item("Status") = CStr(eligibility("Status"))
        item("Selected") = LifeAdlBatchShouldSelect(CStr(eligibility("Status")))
        If eligibility.exists("CurrentEvaluateDate") Then
            If IsDate(eligibility("CurrentEvaluateDate")) Then item("EvaluateDate") = Format$(CDate(eligibility("CurrentEvaluateDate")), "yyyy/mm/dd")
        End If
        If eligibility.exists("MissingReason") Then item("MissingReason") = CStr(eligibility("MissingReason"))
    End If

    Set BuildLifeAdlBatchCandidate = item
End Function

Private Function FindLatestEvaluateRow(ByVal ws As Worksheet) As Long
    Dim evalCol As Long
    Dim lastRow As Long
    Dim rowNo As Long
    Dim parsedDate As Date
    Dim bestDate As Date
    Dim hasBest As Boolean

    evalCol = FindHeaderColAny(ws, Array("Basic.EvalDate", ChrW$(&H8A55) & ChrW$(&H4FA1) & ChrW$(&H65E5), ChrW$(&H8A18) & ChrW$(&H9332) & ChrW$(&H65E5), "EvalDate"))
    If evalCol = 0 Then Exit Function

    lastRow = ws.Cells(ws.rows.count, evalCol).End(xlUp).row
    For rowNo = 2 To lastRow
        If TryParseBatchDate(ws.Cells(rowNo, evalCol).value, parsedDate) Then
            If (Not hasBest) Or parsedDate > bestDate Then
                bestDate = parsedDate
                FindLatestEvaluateRow = rowNo
                hasBest = True
            End If
        End If
    Next rowNo
End Function

Private Function FormatDisplayDate(ByVal rawValue As Variant) As String
    Dim parsedDate As Date
    If TryParseBatchDate(rawValue, parsedDate) Then
        FormatDisplayDate = Format$(parsedDate, "yyyy/mm/dd")
    Else
        FormatDisplayDate = Trim$(CStr(rawValue))
    End If
End Function

Private Function TryParseBatchDate(ByVal rawValue As Variant, ByRef parsedDate As Date) As Boolean
    Dim s As String
    On Error GoTo EH

    If IsDate(rawValue) Then
        parsedDate = dateValue(CDate(rawValue))
        TryParseBatchDate = True
        Exit Function
    End If

    s = Trim$(CStr(rawValue))
    If LenB(s) = 0 Then Exit Function
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo EH
    s = Replace$(s, ChrW$(&H5E74), "/")
    s = Replace$(s, ChrW$(&H6708), "/")
    s = Replace$(s, ChrW$(&H65E5), "")
    s = Replace$(s, ".", "/")
    s = Replace$(s, "-", "/")
    s = Replace$(s, ChrW$(&HFF0F), "/")

    If IsDate(s) Then
        parsedDate = dateValue(CDate(s))
        TryParseBatchDate = True
    End If
    Exit Function
EH:
    TryParseBatchDate = False
End Function

Private Function CellTextByHeader(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerName As String) As String
    Dim colNo As Long
    colNo = FindHeaderCol(ws, headerName)
    If colNo > 0 Then CellTextByHeader = Trim$(CStr(ws.Cells(rowNo, colNo).value))
End Function

Private Function FirstNonBlankHeaderText(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headers As Variant) As String
    FirstNonBlankHeaderText = Trim$(CStr(FirstNonBlankHeaderValue(ws, rowNo, headers)))
End Function

Private Function FirstNonBlankHeaderValue(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headers As Variant) As Variant
    Dim i As Long
    Dim colNo As Long
    Dim rawValue As Variant

    For i = LBound(headers) To UBound(headers)
        colNo = FindHeaderCol(ws, CStr(headers(i)))
        If colNo > 0 Then
            rawValue = ws.Cells(rowNo, colNo).value
            If LenB(Trim$(CStr(rawValue))) > 0 Then
                FirstNonBlankHeaderValue = rawValue
                Exit Function
            End If
        End If
    Next i
    FirstNonBlankHeaderValue = vbNullString
End Function

Private Function TryGetWorksheet(ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    TryGetWorksheet = Not ws Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindHeaderColAny(ByVal ws As Worksheet, ByVal headers As Variant) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        FindHeaderColAny = FindHeaderCol(ws, CStr(headers(i)))
        If FindHeaderColAny > 0 Then Exit Function
    Next i
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colNo As Long

    If ws Is Nothing Then Exit Function
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For colNo = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, colNo).value)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = colNo
            Exit Function
        End If
    Next colNo
End Function
