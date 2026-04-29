Attribute VB_Name = "modLifeAdlEligibility"
Option Explicit

Public Const ADL_ELIGIBILITY_STATUS_FIRST As String = "FIRST"
Public Const ADL_ELIGIBILITY_STATUS_DUE As String = "DUE"
Public Const ADL_ELIGIBILITY_STATUS_SOON As String = "SOON"
Public Const ADL_ELIGIBILITY_STATUS_NOT_DUE As String = "NOT_DUE"
Public Const ADL_ELIGIBILITY_STATUS_INSUFFICIENT As String = "INSUFFICIENT"

Public Function BuildAdlEligibility(ByVal owner As Object) As Object
    Dim result As Object
    Dim currentEvaluateDate As Date
    Dim previousInfo As Object
    Dim wsHistory As Worksheet
    Dim missingReason As String
    Dim previousEvaluateDate As Date

    Set result = CreateObject("Scripting.Dictionary")
    InitializeAdlEligibilityResult result

    If Not TryResolveCurrentEvaluateDate(owner, currentEvaluateDate, missingReason) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = missingReason
        Set BuildAdlEligibility = result
        Exit Function
    End If

    result("CurrentEvaluateDate") = currentEvaluateDate

    If Not TryResolveHistorySheetForAdl(owner, wsHistory) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_FIRST, 1, 0
        Set BuildAdlEligibility = result
        Exit Function
    End If

    Set previousInfo = FindPreviousEvaluateDateInfo(wsHistory, currentEvaluateDate)
    If CBool(previousInfo("HasError")) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = CStr(previousInfo("MissingReason"))
        Set BuildAdlEligibility = result
        Exit Function
    End If

    If Not CBool(previousInfo("HasPrevious")) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_FIRST, 1, 0
        Set BuildAdlEligibility = result
        Exit Function
    End If

    previousEvaluateDate = CDate(previousInfo("PreviousEvaluateDate"))
    result("PreviousEvaluateDate") = previousEvaluateDate
    result("MonthsSincePrevious") = WholeMonthsBetween(previousEvaluateDate, currentEvaluateDate)

    If currentEvaluateDate >= DateAdd("m", 6, previousEvaluateDate) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_DUE, 0, 1
    ElseIf currentEvaluateDate >= DateAdd("m", 5, previousEvaluateDate) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_SOON, 0, 0
    Else
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_NOT_DUE, 0, 0
    End If

    Set BuildAdlEligibility = result
End Function

Public Function BuildAdlEligibilityFromActiveEval() As Object
    Dim owner As Object
    Set owner = FindActiveEvalForm()
    If owner Is Nothing Then Exit Function
    Set BuildAdlEligibilityFromActiveEval = BuildAdlEligibility(owner)
End Function

Public Function BuildAdlEligibilityFromHistoryRow(ByVal wsHistory As Worksheet, ByVal targetRow As Long) As Object
    Dim result As Object
    Dim currentEvaluateDate As Date
    Dim previousInfo As Object
    Dim missingReason As String
    Dim previousEvaluateDate As Date
    Dim evalCol As Long

    Set result = CreateObject("Scripting.Dictionary")
    InitializeAdlEligibilityResult result

    If wsHistory Is Nothing Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = "History sheet was not found."
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    If targetRow < 2 Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = "Latest evaluate row was not found."
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    evalCol = ResolveHistoryEvaluateDateColumn(wsHistory)
    If evalCol = 0 Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = "History evaluate date column was not found."
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    If Not TryParseFlexibleDateValue(wsHistory.Cells(targetRow, evalCol).value, currentEvaluateDate) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = "Current evaluate date could not be parsed."
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    result("CurrentEvaluateDate") = currentEvaluateDate

    Set previousInfo = FindPreviousEvaluateDateInfo(wsHistory, currentEvaluateDate)
    If CBool(previousInfo("HasError")) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, 0, 0
        result("MissingReason") = CStr(previousInfo("MissingReason"))
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    If Not CBool(previousInfo("HasPrevious")) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_FIRST, 1, 0
        Set BuildAdlEligibilityFromHistoryRow = result
        Exit Function
    End If

    previousEvaluateDate = CDate(previousInfo("PreviousEvaluateDate"))
    result("PreviousEvaluateDate") = previousEvaluateDate
    result("MonthsSincePrevious") = WholeMonthsBetween(previousEvaluateDate, currentEvaluateDate)

    If currentEvaluateDate >= DateAdd("m", 6, previousEvaluateDate) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_DUE, 0, 1
    ElseIf currentEvaluateDate >= DateAdd("m", 5, previousEvaluateDate) Then
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_SOON, 0, 0
    Else
        ApplyAdlEligibilityStatus result, ADL_ELIGIBILITY_STATUS_NOT_DUE, 0, 0
    End If

    Set BuildAdlEligibilityFromHistoryRow = result
End Function
Private Sub InitializeAdlEligibilityResult(ByVal result As Object)
    result("Status") = ADL_ELIGIBILITY_STATUS_INSUFFICIENT
    result("CurrentEvaluateDate") = vbNullString
    result("PreviousEvaluateDate") = vbNullString
    result("MonthsSincePrevious") = vbNullString
    result("FirstMonthFlag") = 0
    result("SixthMonthFlag") = 0
    result("MissingReason") = vbNullString
End Sub

Private Sub ApplyAdlEligibilityStatus(ByVal result As Object, ByVal statusText As String, ByVal firstMonthFlag As Long, ByVal sixthMonthFlag As Long)
    result("Status") = statusText
    result("FirstMonthFlag") = firstMonthFlag
    result("SixthMonthFlag") = sixthMonthFlag
    If StrComp(statusText, ADL_ELIGIBILITY_STATUS_INSUFFICIENT, vbTextCompare) <> 0 Then
        result("MissingReason") = vbNullString
    End If
End Sub

Private Function TryResolveCurrentEvaluateDate(ByVal owner As Object, ByRef resultDate As Date, ByRef missingReason As String) As Boolean
    Dim rawValue As String

    rawValue = GetControlTextSafe(owner, "txtEDate")
    If LenB(rawValue) = 0 Then
        missingReason = "Current evaluate date is blank."
        Exit Function
    End If

    If Not TryParseFlexibleDateValue(rawValue, resultDate) Then
        missingReason = "Current evaluate date could not be parsed."
        Exit Function
    End If

    TryResolveCurrentEvaluateDate = True
End Function

Private Function TryResolveHistorySheetForAdl(ByVal owner As Object, ByRef wsHistory As Worksheet) As Boolean
    On Error Resume Next
    TryResolveHistorySheetForAdl = modEvalIOEntry.TryGetUserHistorySheet(owner, wsHistory)
    If Err.Number <> 0 Then
        TryResolveHistorySheetForAdl = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function FindPreviousEvaluateDateInfo(ByVal wsHistory As Worksheet, ByVal currentEvaluateDate As Date) As Object
    Dim result As Object
    Dim evalCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim rawValue As Variant
    Dim rowDate As Date
    Dim bestDate As Date
    Dim hasPrevious As Boolean
    Dim sawNonEmpty As Boolean
    Dim sawInvalid As Boolean

    Set result = CreateObject("Scripting.Dictionary")
    result("HasPrevious") = False
    result("PreviousEvaluateDate") = vbNullString
    result("HasError") = False
    result("MissingReason") = vbNullString

    If wsHistory Is Nothing Then
        Set FindPreviousEvaluateDateInfo = result
        Exit Function
    End If

    evalCol = ResolveHistoryEvaluateDateColumn(wsHistory)
    If evalCol = 0 Then
        result("HasError") = True
        result("MissingReason") = "History evaluate date column was not found."
        Set FindPreviousEvaluateDateInfo = result
        Exit Function
    End If

    lastRow = wsHistory.Cells(wsHistory.rows.count, evalCol).End(xlUp).row
    If lastRow < 2 Then
        Set FindPreviousEvaluateDateInfo = result
        Exit Function
    End If

    For r = 2 To lastRow
        rawValue = wsHistory.Cells(r, evalCol).value
        If IsCellValueBlank(rawValue) Then GoTo ContinueRow

        sawNonEmpty = True
        If TryParseFlexibleDateValue(rawValue, rowDate) Then
            If rowDate < currentEvaluateDate Then
                If (Not hasPrevious) Or (rowDate > bestDate) Then
                    bestDate = rowDate
                    hasPrevious = True
                End If
            End If
        Else
            sawInvalid = True
        End If

ContinueRow:
    Next r

    If hasPrevious Then
        result("HasPrevious") = True
        result("PreviousEvaluateDate") = bestDate
    ElseIf sawNonEmpty And sawInvalid Then
        result("HasError") = True
        result("MissingReason") = "Previous evaluate date could not be parsed from history."
    End If

    Set FindPreviousEvaluateDateInfo = result
End Function
Private Function ResolveHistoryEvaluateDateColumn(ByVal wsHistory As Worksheet) As Long
    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Basic.EvalDate", _
        ChrW(&H8A55) & ChrW(&H4FA1) & ChrW(&H65E5), _
        ChrW(&H8A18) & ChrW(&H9332) & ChrW(&H65E5), _
        ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H65E5), _
        ChrW(&H4F5C) & ChrW(&H6210) & ChrW(&H65E5), _
        "EvalDate")
    For i = LBound(headers) To UBound(headers)
        ResolveHistoryEvaluateDateColumn = FindHeaderColumnExact(wsHistory, CStr(headers(i)))
        If ResolveHistoryEvaluateDateColumn > 0 Then Exit Function
    Next i
End Function

Private Function FindHeaderColumnExact(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindHeaderColumnExact = c
            Exit Function
        End If
    Next c
End Function

Private Function WholeMonthsBetween(ByVal previousDate As Date, ByVal currentDate As Date) As Long
    WholeMonthsBetween = DateDiff("m", previousDate, currentDate)
    If DateAdd("m", WholeMonthsBetween, previousDate) > currentDate Then
        WholeMonthsBetween = WholeMonthsBetween - 1
    End If
    If WholeMonthsBetween < 0 Then WholeMonthsBetween = 0
End Function

Private Function TryParseFlexibleDateValue(ByVal src As Variant, ByRef resultDate As Date) As Boolean
    Dim s As String
    Dim y As Long
    Dim m As Long
    Dim d As Long

    On Error GoTo EH

    If IsDate(src) Then
        resultDate = dateValue(CDate(src))
        TryParseFlexibleDateValue = True
        Exit Function
    End If

    s = NormalizeDateText(CStr(src))
    If LenB(s) = 0 Then Exit Function

    If IsDate(s) Then
        resultDate = dateValue(CDate(s))
        TryParseFlexibleDateValue = True
        Exit Function
    End If

    If TryParseSlashDateLike(s, y, m, d) Then
        resultDate = DateSerial(y, m, d)
        TryParseFlexibleDateValue = True
        Exit Function
    End If

    If TryParseCompactYmd(s, y, m, d) Then
        resultDate = DateSerial(y, m, d)
        TryParseFlexibleDateValue = True
        Exit Function
    End If

    If TryParseWarekiDateLike(s, y, m, d) Then
        resultDate = DateSerial(y, m, d)
        TryParseFlexibleDateValue = True
    End If
    Exit Function

EH:
    Err.Clear
End Function

Private Function NormalizeDateText(ByVal src As String) As String
    Dim s As String

    s = Trim$(src)
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    s = Replace$(s, vbTab, " ")
    s = Replace$(s, ChrW(&H3000), " ")
    If InStr(s, " ") > 0 Then s = Split(s, " ")(0)
    If InStr(s, "(") > 0 Then s = Left$(s, InStr(s, "(") - 1)
    If InStr(s, ChrW(&HFF08)) > 0 Then s = Left$(s, InStr(s, ChrW(&HFF08)) - 1)
    If InStr(s, ChrW(&HFF5E)) > 0 Then s = Left$(s, InStr(s, ChrW(&HFF5E)) - 1)

    NormalizeDateText = Trim$(s)
End Function

Private Function TryParseSlashDateLike(ByVal src As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim normalized As String
    Dim parts As Variant

    normalized = Replace$(src, ".", "/")
    normalized = Replace$(normalized, "-", "/")
    normalized = Replace$(normalized, ChrW(&HFF0F), "/")

    parts = Split(normalized, "/")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If Not IsNumeric(parts(2)) Then Exit Function

    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))
    TryParseSlashDateLike = (y > 0 And m > 0 And d > 0)
End Function

Private Function TryParseCompactYmd(ByVal src As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim digits As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(src)
        ch = Mid$(src, i, 1)
        If ch Like "#" Then digits = digits & ch
    Next i

    If Len(digits) <> 8 Then Exit Function

    y = CLng(Left$(digits, 4))
    m = CLng(Mid$(digits, 5, 2))
    d = CLng(Right$(digits, 2))
    TryParseCompactYmd = (y > 0 And m > 0 And d > 0)
End Function

Private Function TryParseWarekiDateLike(ByVal src As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim eraBase As Long
    Dim eraPos As Long
    Dim yearPos As Long
    Dim monthPos As Long
    Dim yearText As String
    Dim monthText As String
    Dim dayText As String

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
    TryParseWarekiDateLike = True
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

Private Function IsCellValueBlank(ByVal src As Variant) As Boolean
    If IsEmpty(src) Then
        IsCellValueBlank = True
    ElseIf LenB(Trim$(CStr(src))) = 0 Then
        IsCellValueBlank = True
    End If
End Function

Private Function GetControlTextSafe(ByVal owner As Object, ByVal controlName As String) As String
    Dim ctl As Object
    Set ctl = FindControlDeep(owner, controlName)
    If ctl Is Nothing Then Exit Function

    On Error Resume Next
    GetControlTextSafe = Trim$(CStr(ctl.value))
    If Err.Number <> 0 Then
        Err.Clear
        GetControlTextSafe = Trim$(CStr(ctl.text))
    End If
    On Error GoTo 0
End Function

Private Function FindControlDeep(ByVal container As Object, ByVal controlName As String) As Object
    Dim ctl As Object

    On Error Resume Next
    Set FindControlDeep = container.controls(controlName)
    If Err.Number = 0 Then
        If Not FindControlDeep Is Nothing Then Exit Function
    End If
    Err.Clear

    For Each ctl In container.controls
        Set FindControlDeep = FindControlDeep(ctl, controlName)
        If Not FindControlDeep Is Nothing Then Exit Function
    Next ctl
    On Error GoTo 0
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

