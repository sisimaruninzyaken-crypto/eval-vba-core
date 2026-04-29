Attribute VB_Name = "modLifeSettings"
Option Explicit

Private Const LIFE_SETTINGS_HEADER_KEY As String = "Key"
Private Const LIFE_SETTINGS_HEADER_VALUE As String = "Value"

Private mLifeSettingsReady As Boolean

Public Sub EnsureLifeSettingsReady()
    On Error GoTo EH

    If mLifeSettingsReady Then Exit Sub

    Dim ws As Worksheet
    Set ws = EnsureLifeSettingsSheet()
    EnsureLifeSettingsDefaults ws
    mLifeSettingsReady = True
    Exit Sub

EH:
#If APP_DEBUG Then
    Debug.Print "[LifeSettings][ERR]", Err.Number, Err.Description
#End If
End Sub

Public Function GetLifeSetting(ByVal key As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim valueCol As Long

    Set ws = EnsureLifeSettingsSheet()
    EnsureLifeSettingsDefaults ws

    rowIndex = FindLifeSettingRow(ws, key)
    If rowIndex = 0 Then Exit Function

    valueCol = FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_VALUE)
    If valueCol = 0 Then Exit Function

    GetLifeSetting = Trim$(CStr(ws.Cells(rowIndex, valueCol).value))
    Exit Function

EH:
#If APP_DEBUG Then
    Debug.Print "[LifeSettings][Get][ERR]", Err.Number, Err.Description
#End If
End Function

Private Function EnsureLifeSettingsSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LifeSettingsSheetName())
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = LifeSettingsSheetName()
    End If

    EnsureLifeSettingsHeaders ws
    ws.Visible = xlSheetVeryHidden

    Set EnsureLifeSettingsSheet = ws
End Function

Private Sub EnsureLifeSettingsDefaults(ByVal ws As Worksheet)
    EnsureLifeSettingValue ws, "VERSION_PLAN", "2024"
    EnsureLifeSettingValue ws, "VERSION_ADL", "0310"
    EnsureLifeSettingValue ws, "STATUS_DEFAULT", "2"
End Sub

Private Sub EnsureLifeSettingsHeaders(ByVal ws As Worksheet)
    If FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_KEY) = 0 Then
        AppendLifeSettingHeader ws, LIFE_SETTINGS_HEADER_KEY
    End If

    If FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_VALUE) = 0 Then
        AppendLifeSettingHeader ws, LIFE_SETTINGS_HEADER_VALUE
    End If
End Sub

Private Sub AppendLifeSettingHeader(ByVal ws As Worksheet, ByVal headerText As String)
    Dim nextCol As Long
    nextCol = NextLifeSettingsHeaderColumn(ws)
    ws.Cells(1, nextCol).value = headerText
End Sub

Private Function FindLifeSettingColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = LastLifeSettingsHeaderColumn(ws)
    If lastCol = 0 Then Exit Function

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerText, vbTextCompare) = 0 Then
            FindLifeSettingColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function FindLifeSettingRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim keyCol As Long
    Dim lastRow As Long
    Dim r As Long

    keyCol = FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_KEY)
    If keyCol = 0 Then Exit Function

    lastRow = LastLifeSettingsDataRow(ws, keyCol)
    If lastRow < 2 Then Exit Function

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, keyCol).value)), Trim$(key), vbTextCompare) = 0 Then
            FindLifeSettingRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub EnsureLifeSettingValue(ByVal ws As Worksheet, ByVal key As String, ByVal targetValue As String)
    Dim rowIndex As Long
    Dim keyCol As Long
    Dim valueCol As Long

    keyCol = FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_KEY)
    valueCol = FindLifeSettingColumn(ws, LIFE_SETTINGS_HEADER_VALUE)
    If keyCol = 0 Or valueCol = 0 Then Exit Sub

    rowIndex = FindLifeSettingRow(ws, key)
    If rowIndex = 0 Then
        rowIndex = NextLifeSettingDataRow(ws, keyCol)
        ws.Cells(rowIndex, keyCol).value = key
    End If

    ws.Cells(rowIndex, valueCol).NumberFormat = "@": ws.Cells(rowIndex, valueCol).value = targetValue
End Sub

Private Function LastLifeSettingsHeaderColumn(ByVal ws As Worksheet) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    If lastCol = 1 Then
        If Len(Trim$(CStr(ws.Cells(1, 1).value))) = 0 Then
            LastLifeSettingsHeaderColumn = 0
            Exit Function
        End If
    End If

    LastLifeSettingsHeaderColumn = lastCol
End Function

Private Function NextLifeSettingsHeaderColumn(ByVal ws As Worksheet) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = LastLifeSettingsHeaderColumn(ws)
    If lastCol = 0 Then
        NextLifeSettingsHeaderColumn = 1
        Exit Function
    End If

    For c = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(1, c).value))) = 0 Then
            NextLifeSettingsHeaderColumn = c
            Exit Function
        End If
    Next c

    NextLifeSettingsHeaderColumn = lastCol + 1
End Function

Private Function LastLifeSettingsDataRow(ByVal ws As Worksheet, ByVal keyCol As Long) As Long
    LastLifeSettingsDataRow = ws.Cells(ws.rows.count, keyCol).End(xlUp).row
    If LastLifeSettingsDataRow < 1 Then LastLifeSettingsDataRow = 1
End Function

Private Function NextLifeSettingDataRow(ByVal ws As Worksheet, ByVal keyCol As Long) As Long
    Dim lastRow As Long
    lastRow = LastLifeSettingsDataRow(ws, keyCol)
    If lastRow < 2 Then
        NextLifeSettingDataRow = 2
    Else
        NextLifeSettingDataRow = lastRow + 1
    End If
End Function

Private Function LifeSettingsSheetName() As String
    LifeSettingsSheetName = "LIFE" & ChrW(&H8A2D) & ChrW(&H5B9A)
End Function
