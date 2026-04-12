Attribute VB_Name = "modClientMasterWeekdaySelector"


Option Explicit

Private Const CLIENT_MASTER_SHEET_NAME As String = "ClientMaster"
Private Const HDR_USER_ID As String = "UserID"
Private Const HDR_NAME As String = "Name"
Private Const HDR_KANA As String = "Kana"
Private Const HDR_BIRTH_DATE As String = "BirthDate"
Private Const HDR_GENDER As String = "Gender"
Private Const HDR_CARE_LEVEL As String = "CareLevel"
Private Const HDR_USE_WEEKDAY_MON As String = "UseWeekday_Mon"
Private Const HDR_USE_WEEKDAY_TUE As String = "UseWeekday_Tue"
Private Const HDR_USE_WEEKDAY_WED As String = "UseWeekday_Wed"
Private Const HDR_USE_WEEKDAY_THU As String = "UseWeekday_Thu"
Private Const HDR_USE_WEEKDAY_FRI As String = "UseWeekday_Fri"
Private Const HDR_USE_WEEKDAY_SAT As String = "UseWeekday_Sat"

Public Function ResolveClientUseWeekdayHeaderByDate(ByVal targetDate As Date) As String
    ResolveClientUseWeekdayHeaderByDate = ResolveClientUseWeekdayHeaderByWeekday(weekday(targetDate, vbMonday))
End Function

Public Function ResolveClientUseWeekdayHeaderByWeekday(ByVal weekdayMonStart As Long) As String
    Select Case weekdayMonStart
        Case 1: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_MON
        Case 2: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_TUE
        Case 3: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_WED
        Case 4: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_THU
        Case 5: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_FRI
        Case 6: ResolveClientUseWeekdayHeaderByWeekday = HDR_USE_WEEKDAY_SAT
        Case Else
            ResolveClientUseWeekdayHeaderByWeekday = vbNullString
    End Select
End Function


Public Function BuildClientTargetsFromDateValue(ByVal dateValue As Variant) As Collection
    Dim targetDate As Date

    If IsDate(dateValue) Then
        targetDate = CDate(dateValue)
    Else
        targetDate = Date
    End If

    Set BuildClientTargetsFromDateValue = BuildClientTargetsByDate(targetDate)
End Function

Public Function BuildClientTargetsByDate(ByVal targetDate As Date) As Collection
    Dim ws As Worksheet
    Dim weekdayHeader As String

    weekdayHeader = ResolveClientUseWeekdayHeaderByDate(targetDate)
    If Len(weekdayHeader) = 0 Then
        Set BuildClientTargetsByDate = New Collection
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CLIENT_MASTER_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set BuildClientTargetsByDate = New Collection
        Exit Function
    End If

    Set BuildClientTargetsByDate = BuildClientTargetsByWeekdayHeader(ws, weekdayHeader, targetDate)
End Function

Private Function BuildClientTargetsByWeekdayHeader(ByVal ws As Worksheet, ByVal weekdayHeader As String, ByVal targetDate As Date) As Collection
    Dim result As Collection
    Dim lastRow As Long
    Dim colWeekday As Long
    Dim colUserID As Long
    Dim colName As Long
    Dim colKana As Long
    Dim colBirthDate As Long
    Dim colGender As Long
    Dim colCareLevel As Long
    Dim r As Long

    Set result = New Collection

    colWeekday = FindHeaderCol(ws, weekdayHeader)
    If colWeekday <= 0 Then
        Set BuildClientTargetsByWeekdayHeader = result
        Exit Function
    End If

    colUserID = FindHeaderCol(ws, HDR_USER_ID)
    colName = FindHeaderCol(ws, HDR_NAME)
    colKana = FindHeaderCol(ws, HDR_KANA)
    colBirthDate = FindHeaderCol(ws, HDR_BIRTH_DATE)
    colGender = FindHeaderCol(ws, HDR_GENDER)
    colCareLevel = FindHeaderCol(ws, HDR_CARE_LEVEL)

   lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    For r = 2 To lastRow
        If IsTruthyLikeValue(ws.Cells(r, colWeekday).value) Then
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")

            item("Row") = r
            item("TargetDate") = targetDate
            item("TargetDateText") = Format$(targetDate, "yyyy/mm/dd")
            item("WeekdayMonStart") = weekday(targetDate, vbMonday)
            item("WeekdayHeader") = weekdayHeader
            item("UserID") = CellText(ws, r, colUserID)
            item("Name") = CellText(ws, r, colName)
            item("Kana") = CellText(ws, r, colKana)
            item("BirthDate") = CellText(ws, r, colBirthDate)
            item("Gender") = CellText(ws, r, colGender)
            item("CareLevel") = CellText(ws, r, colCareLevel)

            result.Add item
        End If
    Next r

    Set BuildClientTargetsByWeekdayHeader = result
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function CellText(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal colNo As Long) As String
    If colNo <= 0 Then Exit Function
    CellText = Trim$(CStr(ws.Cells(rowNo, colNo).value))
End Function

Private Function IsTruthyLikeValue(ByVal rawValue As Variant) As Boolean
    Dim normalized As String

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function

    Select Case VarType(rawValue)
        Case vbBoolean
            IsTruthyLikeValue = CBool(rawValue)
            Exit Function
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsTruthyLikeValue = (CDbl(rawValue) <> 0)
            Exit Function
    End Select

    normalized = LCase$(Trim$(CStr(rawValue)))
    IsTruthyLikeValue = (normalized = "1" Or normalized = "true" Or normalized = "yes" Or normalized = "y")
End Function




End Sub
