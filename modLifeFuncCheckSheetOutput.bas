Attribute VB_Name = "modLifeFuncCheckSheetOutput"

Option Explicit

Private Const LIFE_FUNC_TEMPLATE_SHEET As String = "生活機能チェックシート"
Private Const LIFE_FUNC_OUTPUT_DIR As String = "LifeFuncCheckSheet"
Private Const UNKNOWN_NAME As String = "unknown"

Private Const LIFE_FUNC_TRACE_TAG As String = "[LifeFuncTrace]"

Private Sub TraceLifeFunc(ByVal message As String)
#If APP_DEBUG Then
    Debug.Print LIFE_FUNC_TRACE_TAG & " " & message
#End If
End Sub


Public Sub ExportLifeFuncCheckSheet(ByVal owner As Object)
    On Error GoTo EH
    TraceLifeFunc "ExportLifeFuncCheckSheet entered"

    Dim templateWs As Worksheet
    On Error Resume Next
    Set templateWs = ThisWorkbook.Worksheets(LIFE_FUNC_TEMPLATE_SHEET)
    On Error GoTo EH
    If templateWs Is Nothing Then
        MsgBox "生活機能チェックシートのテンプレシートが見つかりません。", vbExclamation
        Exit Sub
        Exit Sub
    End If

    Dim patientName As String
    patientName = SanitizeFileToken(GetControlTextSafe(owner, "txtName"), UNKNOWN_NAME)

    Dim evalDateToken As String
    evalDateToken = BuildEvalDateToken(owner)

    Dim outputDir As String
    outputDir = EnsureOutputDirectory(patientName)

    Dim fileBaseName As String
    fileBaseName = patientName & "_" & evalDateToken

    Dim outputPath As String
    outputPath = BuildUniquePath(outputDir, fileBaseName, "xlsx")

    Dim newWb As Workbook
    Dim newWs As Worksheet

    templateWs.Copy
    Set newWb = ActiveWorkbook
    Set newWs = newWb.Worksheets(1)

    WriteLifeFuncCheckSheetContent newWs, owner

    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False

    MsgBox "生活機能チェックシートを保存しました:" & vbCrLf & outputPath, vbInformation
    Exit Sub
EH:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    MsgBox "生活機能チェックシート保存エラー " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Public Function CopyLifeFuncTemplateSheetToWorkbook(ByVal destWb As Workbook) As Worksheet
    On Error GoTo EH
    If destWb Is Nothing Then Exit Function

    Dim templateWs As Worksheet
    Set templateWs = ThisWorkbook.Worksheets("生活機能チェックシート")
    templateWs.Copy After:=destWb.Worksheets(destWb.Worksheets.count)
    Set CopyLifeFuncTemplateSheetToWorkbook = destWb.Worksheets(destWb.Worksheets.count)

    Exit Function
EH:
    On Error Resume Next

    Err.Clear
End Function

Public Sub WriteLifeFuncCheckSheet(ByVal ws As Worksheet, ByVal owner As Object)
    If ws Is Nothing Then Exit Sub
    WriteBasicInfo ws, owner
    WriteEnvironment ws, owner
    WriteLevelTaskComment ws, owner
End Sub



Public Function BuildLifeFuncCheckSheetPathPreview(ByVal owner As Object) As String
    On Error GoTo EH

    Dim patientName As String
    patientName = SanitizeFileToken(GetControlTextSafe(owner, "txtName"), UNKNOWN_NAME)

    Dim evalDateToken As String
    evalDateToken = BuildEvalDateToken(owner)

    BuildLifeFuncCheckSheetPathPreview = ThisWorkbook.path & Application.PathSeparator & _
                                         LIFE_FUNC_OUTPUT_DIR & Application.PathSeparator & _
                                         patientName & Application.PathSeparator & _
                                         patientName & "_" & evalDateToken & "_01.xlsx"
    Exit Function
EH:
    Err.Clear
End Function



Private Function EnsureOutputDirectory(ByVal patientName As String) As String
    Dim rootDir As String
    rootDir = ThisWorkbook.path & Application.PathSeparator & LIFE_FUNC_OUTPUT_DIR
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



Private Sub WriteLifeFuncCheckSheetContent(ByVal ws As Worksheet, ByVal owner As Object)
    On Error GoTo EH

    WriteBasicInfo ws, owner
    WriteEnvironment ws, owner
    WriteLevelTaskComment ws, owner

    Exit Sub
EH:
    MsgBox "LifeFuncCheckSheet write error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub WriteBasicInfo(ByVal ws As Worksheet, ByVal owner As Object)
    Dim sexValue As String
    Dim y3AfterWrite As String

    sexValue = GetControlTextSafe(owner, "cboSex")
    
    WriteMerged ws, "E3:N3", GetControlTextSafe(owner, "txtName")
    WriteMerged ws, "R3:W3", GetControlTextSafe(owner, "txtBirth")
    WriteMerged ws, "Y3:Z3", GetControlTextSafe(owner, "cboSex")
    y3AfterWrite = CStr(ws.Range("Y3").value)
 

    If LenB(Trim$(sexValue)) > 0 And LenB(Trim$(y3AfterWrite)) = 0 Then
        If ws.Range("Y3").MergeCells Then
            ws.Range("Y3").MergeArea.Cells(1, 1).value = sexValue
            TraceLifeFunc "WriteBasicInfo sex fallback applied target=[Y3.MergeArea.LeftTopCell] value=[" & sexValue & "]"
        Else
            ws.Range("Y3").value = sexValue
            TraceLifeFunc "WriteBasicInfo sex fallback applied target=[Y3] value=[" & sexValue & "]"
        End If
    End If
    WriteMerged ws, "E4:R4", BuildEvalDateWithFixedTime(owner)
    WriteMerged ws, "V4:Z4", GetControlTextSafe(owner, "cboCare")
    WriteMerged ws, "E5:N5", GetControlTextSafe(owner, "txtEvaluator")
    WriteMerged ws, "R5:Z5", GetControlTextSafe(owner, "txtEvaluatorJob")
    WriteMerged ws, "I6:Z6", GetControlTextSafe(owner, "cboElder")
    WriteMerged ws, "I7:Z7", GetControlTextSafe(owner, "cboDementia")

End Sub

Private Sub WriteLevelTaskComment(ByVal ws As Worksheet, ByVal owner As Object)
    Dim rows As Variant
    rows = BuildLifeFuncRowMap()

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim srcKey As String
        Dim levelText As String
        Dim scoreText As String
        Dim commentText As String
        Dim taskText As String

        srcKey = CStr(rows(i)(0))
        levelText = ResolveLevelText(owner, srcKey)
        scoreText = GetControlTextSafe(owner, Replace$(srcKey, "BI_", "cmbBI_"))
        taskText = BuildTaskText(levelText)
        levelText = FormatLevelDisplayText(srcKey, levelText, scoreText)
        commentText = ResolveCommentText(owner, srcKey)

        WriteMerged ws, CStr(rows(i)(1)), levelText
        WriteMerged ws, CStr(rows(i)(2)), taskText

        If LenB(Trim$(commentText)) = 0 Then
            TraceLifeFunc "WriteLevelTaskComment skip comment srcKey=[" & srcKey & "] address=[" & CStr(rows(i)(3)) & "]"
        Else
            WriteMerged ws, CStr(rows(i)(3)), commentText
        End If
    Next i
End Sub

Private Function BuildLifeFuncRowMap() As Variant
    BuildLifeFuncRowMap = Array( _
        Array("BI_0", "G13:N14", "O13:P14", "Q13:Z14"), _
        Array("BI_1", "G15:N16", "O15:P16", "Q15:Z16"), _
        Array("BI_2", "G17:N18", "O17:P18", "Q17:Z18"), _
        Array("BI_3", "G19:N20", "O19:P20", "Q19:Z20"), _
        Array("BI_4", "G21:N22", "O21:P22", "Q21:Z22"), _
        Array("BI_5", "G23:N24", "O23:P24", "Q23:Z24"), _
        Array("BI_6", "G25:N26", "O25:P26", "Q25:Z26"), _
        Array("BI_7", "G27:N28", "O27:P28", "Q27:Z28"), _
        Array("BI_8", "G29:N30", "O29:P30", "Q29:Z30"), _
        Array("BI_9", "G31:N32", "O31:P32", "Q31:Z32"), _
        Array("IADL_0", "G33:N34", "O33:P34", "Q33:Z34"), _
        Array("IADL_1", "G35:N36", "O35:P36", "Q35:Z36"), _
        Array("IADL_2", "G37:N38", "O37:P38", "Q37:Z38"), _
        Array("Kyo_Roll", "G40:N41", "O40:P41", "Q40:Z41"), _
        Array("Kyo_SitUp", "G42:N43", "O42:P43", "Q42:Z43"), _
        Array("Kyo_SitHold", "G44:N45", "O44:P45", "Q44:Z45"), _
        Array("Kyo_StandUp", "G46:N47", "O46:P47", "Q46:Z47"), _
        Array("Kyo_StandHold", "G48:N49", "O48:P49", "Q48:Z49") _
    )
End Function

Private Function ResolveLevelText(ByVal owner As Object, ByVal srcKey As String) As String
    Select Case srcKey
        Case "BI_0", "BI_1", "BI_2", "BI_3", "BI_4", "BI_5", "BI_6", "BI_7", "BI_8", "BI_9"
            ResolveLevelText = LFM_BIWordLevel(srcKey, GetControlTextSafe(owner, Replace$(srcKey, "BI_", "cmbBI_")), vbNullString)
        Case "IADL_0", "IADL_1", "IADL_2"
            ResolveLevelText = GetControlTextSafe(owner, Replace$(srcKey, "IADL_", "cmbIADL_"))
        Case "Kyo_Roll"
            ResolveLevelText = GetControlTextSafe(owner, "cmbKyo_Roll")
        Case "Kyo_SitUp"
            ResolveLevelText = GetControlTextSafe(owner, "cmbKyo_SitUp")
        Case "Kyo_SitHold"
            ResolveLevelText = GetControlTextSafe(owner, "cmbKyo_SitHold")
        Case "Kyo_StandUp"
            ResolveLevelText = ResolveKyoUnnamedComboText(owner, BuildWordKyoStandUpLabel())
        Case "Kyo_StandHold"
            ResolveLevelText = ResolveKyoUnnamedComboText(owner, BuildWordKyoStandHoldLabel())
    End Select
End Function

Private Function ResolveCommentText(ByVal owner As Object, ByVal srcKey As String) As String
    If Left$(srcKey, 3) = "BI_" Then
        ResolveCommentText = JoinNonEmpty( _
            GetControlTextSafeAny(owner, "txtBICheck_" & Mid$(srcKey, 4), "txtBIChk_" & Mid$(srcKey, 4)), _
            GetControlTextSafeAny(owner, "txtBIRemark_" & Mid$(srcKey, 4), "txtBINote_" & Mid$(srcKey, 4)), _
            " / ")
        Exit Function
    End If

    If Left$(srcKey, 5) = "IADL_" Then
        ResolveCommentText = GetControlTextSafeAny(owner, "txtIADLNote")
        Exit Function
    End If

    ResolveCommentText = GetControlTextSafeAny(owner, "txtKyoNote")
End Function

Private Function BuildTaskText(ByVal levelText As String) As String
    Dim normalized As Long
    normalized = LFM_NormalizeAssistLevel(levelText, -1)
    If normalized = 2 Then
        BuildTaskText = BuildWordMu()
    ElseIf LenB(Trim$(levelText)) > 0 Then
        BuildTaskText = BuildWordYu()
    End If
End Function

Private Function FormatLevelDisplayText(ByVal srcKey As String, ByVal levelText As String, ByVal biScoreText As String) As String
    Dim s As String
    s = NormalizeLevelDisplayText(levelText)

    If Left$(srcKey, 3) <> "BI_" Then
        FormatLevelDisplayText = s
        Exit Function
    End If

    biScoreText = Trim$(biScoreText)
    If LenB(s) = 0 Or LenB(biScoreText) = 0 Then
        FormatLevelDisplayText = s
    Else
        FormatLevelDisplayText = s & BuildWordParenOpen() & biScoreText & BuildWordParenClose()
    End If
End Function

Private Function NormalizeLevelDisplayText(ByVal src As String) As String
    Dim s As String
    s = Trim$(src)
    If LenB(s) = 0 Then Exit Function

    s = Replace$(s, BuildWordMimamoriKanshika(), BuildWordMimamori())
    s = Replace$(s, BuildWordZaiiHoji(), BuildWordZaii())
    s = Replace$(s, BuildWordRitsuiHoji(), BuildWordRitsui())
    NormalizeLevelDisplayText = s
End Function

Private Function BuildEvalDateWithFixedTime(ByVal owner As Object) As String
    Dim d As String
    d = Trim$(GetControlTextSafe(owner, "txtEDate"))
    If LenB(d) = 0 Then Exit Function
    BuildEvalDateWithFixedTime = d & " " & "13:00" & ChrW$(65374) & "15:00"
End Function

Private Function GetControlTextSafeAny(ByVal owner As Object, ParamArray names() As Variant) As String
    Dim i As Long
    For i = LBound(names) To UBound(names)
        GetControlTextSafeAny = GetControlTextSafe(owner, CStr(names(i)))
        If LenB(GetControlTextSafeAny) > 0 Then Exit Function
    Next i
End Function

Private Sub WriteMerged(ByVal ws As Worksheet, ByVal addressText As String, ByVal textValue As String)
    On Error GoTo EH
    Dim rng As Range
    Set rng = ws.Range(addressText)
    If ShouldTraceEnvRange(rng) Then
        TraceLifeFunc "WriteMerged before ws=[" & ws.name & "] range=[" & rng.Address(False, False) & "] value=[" & CStr(rng.Cells(1, 1).value) & "] mergeArea=[" & rng.Cells(1, 1).MergeArea.Address(False, False) & "]"
    End If

    rng.Cells(1, 1).value = textValue

    If ShouldTraceEnvRange(rng) Then
        TraceLifeFunc "WriteMerged after ws=[" & ws.name & "] range=[" & rng.Address(False, False) & "] value=[" & CStr(rng.Cells(1, 1).value) & "] mergeArea=[" & rng.Cells(1, 1).MergeArea.Address(False, False) & "]"
    End If
    Exit Sub
EH:
    TraceLifeFunc "WriteMerged error ws=[" & IIf(ws Is Nothing, "", ws.name) & "] range=[" & addressText & "] err=" & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

Private Function ShouldTraceEnvRange(ByVal rng As Range) As Boolean
    On Error GoTo EH
    If rng Is Nothing Then Exit Function
    ShouldTraceEnvRange = Not Intersect(rng, rng.Worksheet.Range("Q13:U32")) Is Nothing
    Exit Function
EH:
    Err.Clear
End Function


Private Function JoinNonEmpty(ByVal leftText As String, ByVal rightText As String, ByVal sep As String) As String
    leftText = Trim$(leftText)
    rightText = Trim$(rightText)

    If LenB(leftText) = 0 Then
        JoinNonEmpty = rightText
    ElseIf LenB(rightText) = 0 Then
        JoinNonEmpty = leftText
    Else
        JoinNonEmpty = leftText & sep & rightText
    End If
End Function

Private Function BuildWordMimamoriKanshika() As String
    BuildWordMimamoriKanshika = ChrW$(35211) & ChrW$(23432) & ChrW$(12426) & ChrW$(65288) & ChrW$(30435) & ChrW$(35222) & ChrW$(19979) & ChrW$(65289)
End Function

Private Function BuildWordMimamori() As String
    BuildWordMimamori = ChrW$(35211) & ChrW$(23432) & ChrW$(12426)
End Function

Private Function BuildWordZaiiHoji() As String
    BuildWordZaiiHoji = ChrW$(24231) & ChrW$(20301) & ChrW$(20445) & ChrW$(25345)
End Function

Private Function BuildWordZaii() As String
    BuildWordZaii = ChrW$(24231) & ChrW$(20301)
End Function

Private Function BuildWordRitsuiHoji() As String
    BuildWordRitsuiHoji = ChrW$(31435) & ChrW$(20301) & ChrW$(20445) & ChrW$(25345)
End Function

Private Function BuildWordRitsui() As String
    BuildWordRitsui = ChrW$(31435) & ChrW$(20301)
End Function

Private Function BuildWordYu() As String
    BuildWordYu = ChrW$(26377)
End Function

Private Function BuildWordMu() As String
    BuildWordMu = ChrW$(28961)
End Function

Private Function BuildWordParenOpen() As String
    BuildWordParenOpen = ChrW$(65288)
End Function

Private Function BuildWordParenClose() As String
    BuildWordParenClose = ChrW$(65289)
End Function

Private Sub WriteEnvironment(ByVal ws As Worksheet, ByVal owner As Object)
    Dim envText As String
    envText = BuildHomeEnvText(owner)
    If LenB(envText) = 0 Then
        TraceLifeFunc "WriteEnvironment skipped: envText is empty"
        Exit Sub
    End If

    Dim envAddress As String
    envAddress = ResolveEnvironmentAddress(ws)
    If LenB(envAddress) = 0 Then
        TraceLifeFunc "WriteEnvironment skipped: envAddress is empty"
        Exit Sub
    End If

    TraceLifeFunc "WriteEnvironment write address=[" & envAddress & "]"

    WriteMerged ws, envAddress, envText
End Sub

Private Function BuildHomeEnvText(ByVal owner As Object) As String
    Dim labels As Collection
    Set labels = New Collection
    

    CollectHomeEnvCheckedCaptions owner, labels

    Dim text As String
    text = JoinCollection(labels, BuildWordDot())

    Dim note As String
    note = GetControlTextSafeAnyDeep(owner, "txtBIHomeEnvNote", "txtHomeNote")


    If LenB(note) > 0 Then
        If LenB(text) > 0 Then
            text = text & BuildWordKuten() & BuildWordBikoLabel() & note
        Else
            text = BuildWordBikoLabel() & note
        End If
    End If

    BuildHomeEnvText = text
 
    
End Function



Private Sub CollectHomeEnvCheckedCaptions(ByVal container As Object, ByVal labels As Collection)
    
    On Error GoTo EH
    If container Is Nothing Then Exit Sub
    
    Dim pagesObj As Object
    Set pagesObj = TryGetObjectMember(container, "Pages")
    If Not pagesObj Is Nothing Then
        Dim pg As Object
        For Each pg In pagesObj
            CollectHomeEnvCheckedCaptions pg, labels
        Next pg
    End If

    Dim controlsObj As Object
    Set controlsObj = TryGetObjectMember(container, "Controls")
    If controlsObj Is Nothing Then Exit Sub

    Dim ctl As Object
    For Each ctl In controlsObj
        If IsHomeEnvCheckControl(ctl) Then
            If GetCheckValueSafe(ctl) Then AddUniqueText labels, GetControlCaptionSafe(ctl)
        End If
        CollectHomeEnvCheckedCaptions ctl, labels
    Next ctl
    Exit Sub
EH:
    Err.Clear
    
End Sub

Private Function IsHomeEnvCheckControl(ByVal ctl As Object) As Boolean
    If ctl Is Nothing Then Exit Function
    If StrComp(TypeName(ctl), "CheckBox", vbTextCompare) <> 0 Then Exit Function

    Dim tagText As String
    tagText = GetControlTagSafe(ctl)
    If Len(tagText) < Len("BI.HomeEnv.") Then Exit Function
    IsHomeEnvCheckControl = (StrComp(Left$(tagText, Len("BI.HomeEnv.")), "BI.HomeEnv.", vbTextCompare) = 0)
End Function

Private Function ResolveEnvironmentAddress(ByVal ws As Worksheet) As String
    Dim fixedAddress As String
    fixedAddress = "Q13:U32"

    ResolveEnvironmentAddress = fixedAddress
End Function

Private Function TryGetObjectMember(ByVal obj As Object, ByVal memberName As String) As Object
    On Error GoTo EH
    Select Case memberName
        Case "Controls"
            Set TryGetObjectMember = obj.Controls
        Case "Pages"
            Set TryGetObjectMember = obj.Pages
    End Select
    Exit Function
EH:
    Err.Clear
End Function


Private Function GetControlNameSafe(ByVal ctl As Object) As String
    On Error GoTo EH
    GetControlNameSafe = CStr(ctl.name)
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetParentNameSafe(ByVal ctl As Object) As String
    On Error GoTo EH
    If ctl Is Nothing Then Exit Function
    If ctl.parent Is Nothing Then Exit Function
    GetParentNameSafe = CStr(ctl.parent.name)
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetControlValueAsTextSafe(ByVal ctl As Object) As String
    On Error GoTo EH
    If ctl Is Nothing Then Exit Function
    GetControlValueAsTextSafe = Trim$(CStr(ctl.value))
    Exit Function
EH:
    Err.Clear
End Function




Private Function GetControlTextSafeAnyDeep(ByVal owner As Object, ParamArray names() As Variant) As String
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim ctl As Object
        Set ctl = FindControlByNameDeep(owner, CStr(names(i)))
        If Not ctl Is Nothing Then
            On Error Resume Next
            GetControlTextSafeAnyDeep = Trim$(CStr(ctl.value))
            Err.Clear
            On Error GoTo 0
            If LenB(GetControlTextSafeAnyDeep) > 0 Then Exit Function
        End If
    Next i
End Function

Private Function FindControlByNameDeep(ByVal container As Object, ByVal ctrlName As String) As Object
    On Error GoTo EH
    If container Is Nothing Then Exit Function

    Dim thisName As String
    On Error Resume Next
    thisName = CStr(container.name)
    Err.Clear
    On Error GoTo EH
    If LenB(thisName) > 0 Then
        If StrComp(thisName, ctrlName, vbTextCompare) = 0 Then
            Set FindControlByNameDeep = container
            Exit Function
        End If
    End If

    Dim pagesObj As Object
    Set pagesObj = TryGetObjectMember(container, "Pages")
    If Not pagesObj Is Nothing Then
        Dim pg As Object
        For Each pg In pagesObj
            Set FindControlByNameDeep = FindControlByNameDeep(pg, ctrlName)
            If Not FindControlByNameDeep Is Nothing Then Exit Function
        Next pg
    End If

    Dim controlsObj As Object
    Set controlsObj = TryGetObjectMember(container, "Controls")
    If controlsObj Is Nothing Then Exit Function

    Dim ctl As Object
    For Each ctl In controlsObj
        Set FindControlByNameDeep = FindControlByNameDeep(ctl, ctrlName)
        If Not FindControlByNameDeep Is Nothing Then Exit Function
    Next ctl
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetCheckValueSafe(ByVal ctl As Object) As Boolean
    On Error GoTo EH
    GetCheckValueSafe = CBool(ctl.value)
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetControlCaptionSafe(ByVal ctl As Object) As String
    On Error GoTo EH
    GetControlCaptionSafe = Trim$(CStr(ctl.caption))
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetControlTagSafe(ByVal ctl As Object) As String
    On Error GoTo EH
    GetControlTagSafe = Trim$(CStr(ctl.tag))
    Exit Function
EH:
    Err.Clear
End Function

Private Sub AddUniqueText(ByVal col As Collection, ByVal textValue As String)
    textValue = Trim$(textValue)
    If LenB(textValue) = 0 Then Exit Sub

    On Error Resume Next
    col.Add textValue, textValue
    Err.Clear
    On Error GoTo 0
End Sub

Private Function JoinCollection(ByVal col As Collection, ByVal sep As String) As String
    Dim i As Long
    For i = 1 To col.count
        If i > 1 Then JoinCollection = JoinCollection & sep
        JoinCollection = JoinCollection & CStr(col(i))
    Next i
End Function

Private Function BuildWordEnvironmentHeader() As String
    BuildWordEnvironmentHeader = ChrW$(29872) & ChrW$(22659) & ChrW$(65288) & ChrW$(23455) & ChrW$(26045) & ChrW$(22580) & ChrW$(25152) & ChrW$(12539) & ChrW$(35036) & ChrW$(21161) & ChrW$(20855) & ChrW$(31561) & ChrW$(65289)
End Function

Private Function BuildWordDot() As String
    BuildWordDot = ChrW$(12539)
End Function

Private Function BuildWordKuten() As String
    BuildWordKuten = ChrW$(12290)
End Function

Private Function BuildWordBikoLabel() As String
    BuildWordBikoLabel = ChrW$(20633) & ChrW$(32771) & ChrW$(65306)
End Function

Private Function BuildWordKyoStandUpLabel() As String
    BuildWordKyoStandUpLabel = ChrW$(31435) & ChrW$(12385) & ChrW$(19978) & ChrW$(12364) & ChrW$(12426)
End Function

Private Function BuildWordKyoStandHoldLabel() As String
    BuildWordKyoStandHoldLabel = ChrW$(31435) & ChrW$(20301) & ChrW$(20445) & ChrW$(25345)
End Function

Private Function ResolveKyoUnnamedComboText(ByVal owner As Object, ByVal labelCaption As String) As String
    Dim cmb As Object
    Set cmb = FindRightComboByLabelCaptionDeep(owner, labelCaption)
    If Not cmb Is Nothing Then ResolveKyoUnnamedComboText = Trim$(CStr(cmb.value))
End Function

Private Function FindRightComboByLabelCaptionDeep(ByVal container As Object, ByVal labelCaption As String) As Object
    Dim targetLabel As Object
    Set targetLabel = FindLabelByCaptionDeep(container, labelCaption)
    If targetLabel Is Nothing Then Exit Function
    Set FindRightComboByLabelCaptionDeep = FindNearestRightComboOnSameRow(targetLabel.parent, targetLabel)
End Function

Private Function FindLabelByCaptionDeep(ByVal container As Object, ByVal labelCaption As String) As Object
    On Error GoTo EH

    Dim c As Object
    For Each c In container.Controls
        If TypeName(c) = "Label" Then
            If StrComp(Trim$(CStr(c.caption)), labelCaption, vbBinaryCompare) = 0 Then
                Set FindLabelByCaptionDeep = c
                Exit Function
            End If
        End If
        If HasControls(c) Then
            Set FindLabelByCaptionDeep = FindLabelByCaptionDeep(c, labelCaption)
            If Not FindLabelByCaptionDeep Is Nothing Then Exit Function
        End If
    Next c
    Exit Function
EH:
    Err.Clear
End Function

Private Function FindNearestRightComboOnSameRow(ByVal container As Object, ByVal targetLabel As Object) As Object
    On Error GoTo EH

    Dim best As Object
    Dim bestDx As Double
    bestDx = 1E+30

    Dim c As Object
    For Each c In container.Controls
        If TypeName(c) = "ComboBox" Then
            Dim dy As Double
            Dim dx As Double
            dy = Abs(CDbl(c.top) - CDbl(targetLabel.top))
            dx = CDbl(c.Left) - CDbl(targetLabel.Left)
            If dy <= 6 And dx > 0 Then
                If dx < bestDx Then
                    bestDx = dx
                    Set best = c
                End If
            End If
        End If
    Next c

    Set FindNearestRightComboOnSameRow = best
    Exit Function
EH:
    Err.Clear
End Function

Private Function HasControls(ByVal obj As Object) As Boolean
    On Error GoTo EH
    Dim n As Long
    n = obj.Controls.count
    HasControls = (n >= 0)
    Exit Function
EH:
    Err.Clear
End Function

