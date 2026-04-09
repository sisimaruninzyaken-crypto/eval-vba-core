Attribute VB_Name = "modDailyCommonRecord"


Option Explicit

Private Const COMMON_RECORD_SHEET_NAME As String = "DailyCommonRecordConfig"
Private Const COL_WEEKDAY As Long = 1
Private Const COL_TEXT As Long = 2

Public Function GetCommonRecordByWeekday(ByVal weekday As Long) As String
    Dim ws As Worksheet
    Dim rowIndex As Long

    Set ws = EnsureCommonRecordSheet()
    rowIndex = FindWeekdayRow(ws, weekday)

    GetCommonRecordByWeekday = CStr(ws.Cells(rowIndex, COL_TEXT).value)
End Function

Public Sub SaveCommonRecordByWeekday(ByVal weekday As Long, ByVal text As String)
    Dim ws As Worksheet
    Dim rowIndex As Long

    Set ws = EnsureCommonRecordSheet()
    rowIndex = FindWeekdayRow(ws, weekday)

    ws.Cells(rowIndex, COL_TEXT).value = CStr(text)
End Sub

Public Function MergeDailyLog(ByVal commonText As String, ByVal tokhenText As String) As String
    Dim normalizedCommon As String
    Dim normalizedTokhen As String

    normalizedCommon = Trim$(CStr(commonText))
    normalizedTokhen = Trim$(CStr(tokhenText))

    If Len(normalizedCommon) = 0 Then
        MergeDailyLog = normalizedTokhen
    ElseIf Len(normalizedTokhen) = 0 Then
        MergeDailyLog = normalizedCommon
    Else
        MergeDailyLog = normalizedCommon & vbCrLf & normalizedTokhen
    End If
End Function

Private Function EnsureCommonRecordSheet() As Worksheet
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COMMON_RECORD_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = COMMON_RECORD_SHEET_NAME

        ws.Cells(1, COL_WEEKDAY).value = "Weekday"
        ws.Cells(1, COL_TEXT).value = "CommonRecord"

        For i = 1 To 7
            ws.Cells(i + 1, COL_WEEKDAY).value = i
            ws.Cells(i + 1, COL_TEXT).value = vbNullString
        Next i
    End If

    ws.Visible = xlSheetVeryHidden
    Set EnsureCommonRecordSheet = ws
End Function

Private Function NormalizeWeekday(ByVal weekdayValue As Long) As Long
    If weekdayValue < 1 Or weekdayValue > 7 Then
        NormalizeWeekday = weekday(Date, vbSunday)
    Else
        NormalizeWeekday = weekdayValue
    End If
End Function

Private Function FindWeekdayRow(ByVal ws As Worksheet, ByVal weekday As Long) As Long
    Dim normalized As Long
    Dim lastRow As Long
    Dim r As Long

    normalized = NormalizeWeekday(weekday)
    lastRow = ws.Cells(ws.rows.count, COL_WEEKDAY).End(xlUp).row

    For r = 2 To lastRow
        If CLng(val(ws.Cells(r, COL_WEEKDAY).value)) = normalized Then
            FindWeekdayRow = r
            Exit Function
        End If
    Next r

    FindWeekdayRow = lastRow + 1
    ws.Cells(FindWeekdayRow, COL_WEEKDAY).value = normalized
    ws.Cells(FindWeekdayRow, COL_TEXT).value = vbNullString
End Function

