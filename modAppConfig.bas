Attribute VB_Name = "modAppConfig"

Option Explicit

Private Const CONFIG_SHEET_NAME As String = "AppConfig"
Private Const KEY_FACILITY_NAME As String = "FacilityName"
Private Const KEY_FACILITY_NO As String = "FacilityNo"
Private Const KEY_FACILITY_ADDRESS As String = "FacilityAddress"
Private Const KEY_FACILITY_PHONE As String = "FacilityPhone"
Private Const KEY_IS_INITIALIZED As String = "IsInitialized"

Private Const COL_KEY As Long = 1
Private Const COL_VALUE As Long = 2

Public Function EnsureFacilitySetupOnStartup() As Boolean
    If IsFacilityInitialized() Then
        EnsureFacilitySetupOnStartup = True
        Exit Function
    End If

    Dim frm As frmFacilitySettings
    Set frm = New frmFacilitySettings
    frm.Show vbModal

    EnsureFacilitySetupOnStartup = frm.IsSaved
    Unload frm
    Set frm = Nothing
End Function

Public Function IsFacilityInitialized() As Boolean
    Dim flagValue As String
    flagValue = GetConfigValue(KEY_IS_INITIALIZED)

    If StrComp(flagValue, "1", vbBinaryCompare) = 0 Then
        IsFacilityInitialized = True
        Exit Function
    End If

    Dim facilityName As String
    Dim facilityNo As String
    Dim facilityAddress As String
    Dim facilityPhone As String

    LoadFacilitySettings facilityName, facilityNo, facilityAddress, facilityPhone
    IsFacilityInitialized = (LenB(facilityName) > 0 And LenB(facilityNo) > 0 And LenB(facilityAddress) > 0 And LenB(facilityPhone) > 0)
End Function

Public Sub LoadFacilitySettings(ByRef facilityName As String, _
                                ByRef facilityNo As String, _
                                ByRef facilityAddress As String, _
                                ByRef facilityPhone As String)
    facilityName = GetConfigValue(KEY_FACILITY_NAME)
    facilityNo = GetConfigValue(KEY_FACILITY_NO)
    facilityAddress = GetConfigValue(KEY_FACILITY_ADDRESS)
    facilityPhone = GetConfigValue(KEY_FACILITY_PHONE)
End Sub

Public Sub SaveFacilitySettings(ByVal facilityName As String, _
                                ByVal facilityNo As String, _
                                ByVal facilityAddress As String, _
                                ByVal facilityPhone As String)
    SetConfigValue KEY_FACILITY_NAME, Trim$(facilityName)
    SetConfigValue KEY_FACILITY_NO, Trim$(facilityNo)
    SetConfigValue KEY_FACILITY_ADDRESS, Trim$(facilityAddress)
    SetConfigValue KEY_FACILITY_PHONE, Trim$(facilityPhone)
    SetConfigValue KEY_IS_INITIALIZED, "1"
End Sub

Private Function EnsureConfigSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = CONFIG_SHEET_NAME

        ws.Cells(1, COL_KEY).value = KEY_FACILITY_NAME
        ws.Cells(2, COL_KEY).value = KEY_FACILITY_NO
        ws.Cells(3, COL_KEY).value = KEY_FACILITY_ADDRESS
        ws.Cells(4, COL_KEY).value = KEY_FACILITY_PHONE
        ws.Cells(5, COL_KEY).value = KEY_IS_INITIALIZED
        ws.Cells(6, COL_KEY).value = "UpdatedAt"
    End If

    ws.Visible = xlSheetVeryHidden
    Set EnsureConfigSheet = ws
End Function

Private Function FindConfigRow(ByVal ws As Worksheet, ByVal configKey As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, COL_KEY).End(xlUp).row

    Dim r As Long
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, COL_KEY).value)), configKey, vbTextCompare) = 0 Then
            FindConfigRow = r
            Exit Function
        End If
    Next r

    FindConfigRow = lastRow + 1
    ws.Cells(FindConfigRow, COL_KEY).value = configKey
End Function

Private Function GetConfigValue(ByVal configKey As String) As String
    Dim ws As Worksheet
    Set ws = EnsureConfigSheet()

    Dim rowIndex As Long
    rowIndex = FindConfigRow(ws, configKey)

    GetConfigValue = Trim$(CStr(ws.Cells(rowIndex, COL_VALUE).value))
End Function

Private Sub SetConfigValue(ByVal configKey As String, ByVal configValue As String)
    Dim ws As Worksheet
    Set ws = EnsureConfigSheet()

    Dim rowIndex As Long
    rowIndex = FindConfigRow(ws, configKey)

    ws.Cells(rowIndex, COL_VALUE).value = configValue

    Dim updatedRow As Long
    updatedRow = FindConfigRow(ws, "UpdatedAt")
    ws.Cells(updatedRow, COL_VALUE).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub


