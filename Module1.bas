Attribute VB_Name = "Module1"


Option Explicit

Private Const PLAN_TEMPLATE_SHEET As String = "個別機能訓練計画書"
Private Const PLAN_OUTPUT_DIR As String = "KojinPlan"
Private Const UNKNOWN_NAME As String = "kanja"

Public Sub ExportEvalPlanSheet(ByVal owner As Object, ByVal planData As Object, Optional ByVal patientName As String = "")
    On Error GoTo EH

    Dim templateWs As Worksheet
    On Error Resume Next
    Set templateWs = ThisWorkbook.Worksheets(PLAN_TEMPLATE_SHEET)
    On Error GoTo EH
    If templateWs Is Nothing Then
        MsgBox "個別機能訓練計画書のテンプレートシートが見つかりません。", vbExclamation
        Exit Sub
    End If

    If LenB(Trim$(patientName)) = 0 Then
        patientName = GetControlTextSafe(owner, "txtName")
    End If

    Dim safePatientName As String
    safePatientName = SanitizeFileToken(patientName, UNKNOWN_NAME)

    Dim evalDateToken As String
    evalDateToken = BuildEvalDateToken(owner)

    Dim outputDir As String
    outputDir = EnsureOutputDirectory(safePatientName)

    Dim outputPath As String
    outputPath = BuildUniquePath(outputDir, safePatientName & "_" & evalDateToken, "xlsx")

    Dim newWb As Workbook
    Dim newWs As Worksheet

    templateWs.Copy
    Set newWb = ActiveWorkbook
    Set newWs = newWb.Worksheets(1)

    modEvalPlanSheetOutput.WriteEvalPlanSheet newWs, owner, planData

    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False

    MsgBox "saved: " & outputPath, vbInformation, "done"
    Exit Sub
EH:
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "error"
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
End Sub

Public Function BuildEvalPlanSheetPathPreview(ByVal owner As Object) As String
    On Error GoTo EH

    Dim patientName As String
    patientName = SanitizeFileToken(GetControlTextSafe(owner, "txtName"), UNKNOWN_NAME)

    Dim evalDateToken As String
    evalDateToken = BuildEvalDateToken(owner)

    BuildEvalPlanSheetPathPreview = ThisWorkbook.path & Application.PathSeparator & _
                                    PLAN_OUTPUT_DIR & Application.PathSeparator & _
                                    patientName & Application.PathSeparator & _
                                    patientName & "_" & evalDateToken & "_01.xlsx"
    Exit Function
EH:
    Err.Clear
End Function

Private Function EnsureOutputDirectory(ByVal patientName As String) As String
    Dim rootDir As String
    rootDir = ThisWorkbook.path & Application.PathSeparator & PLAN_OUTPUT_DIR
    EnsureFolderExists rootDir

    Dim patientDir As String
    patientDir = rootDir & Application.PathSeparator & patientName
    EnsureFolderExists patientDir

    EnsureOutputDirectory = patientDir
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If LenB(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function BuildUniquePath(ByVal folderPath As String, ByVal fileBaseName As String, ByVal ext As String) As String
    Dim seq As Long
    Dim candidate As String

    seq = 1
    Do
        candidate = folderPath & Application.PathSeparator & fileBaseName & "_" & Format$(seq, "00") & "." & ext
        If LenB(Dir$(candidate, vbNormal)) = 0 Then
            BuildUniquePath = candidate
            Exit Function
        End If
        seq = seq + 1
    Loop
End Function

Private Function BuildEvalDateToken(ByVal owner As Object) As String
    Dim rawDate As String
    rawDate = Trim$(GetControlTextSafe(owner, "txtEDate"))

    If LenB(rawDate) = 0 Then
        BuildEvalDateToken = Format$(Date, "yyyymmdd")
        Exit Function
    End If

    If IsDate(rawDate) Then
        BuildEvalDateToken = Format$(CDate(rawDate), "yyyymmdd")
        Exit Function
    End If

    BuildEvalDateToken = SanitizeFileToken(rawDate, Format$(Date, "yyyymmdd"))
End Function

Private Function SanitizeFileToken(ByVal src As String, ByVal fallbackValue As String) As String
    Dim token As String
    token = Trim$(src)

    Dim ng As Variant
    For Each ng In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "[", "]")
        token = Replace$(token, CStr(ng), "_")
    Next ng

    token = Replace$(token, vbTab, " ")
    Do While InStr(token, "  ") > 0
        token = Replace$(token, "  ", " ")
    Loop

    token = Trim$(token)
    If LenB(token) = 0 Then token = fallbackValue

    SanitizeFileToken = token
End Function

Private Function GetControlTextSafe(ByVal owner As Object, ByVal controlName As String) As String
    On Error GoTo EH
    If owner Is Nothing Then Exit Function
    GetControlTextSafe = Trim$(CStr(owner.Controls(controlName).value))
    Exit Function
EH:
    Err.Clear
End Function


