Attribute VB_Name = "modEvalPlanSheetOutput"



Option Explicit

Public Sub WriteEvalPlanSheet(ByVal ws As Worksheet, ByVal owner As Object, Optional ByVal planData As Object = Nothing)
    
    On Error GoTo EH
    
    If ws Is Nothing Then Exit Sub

    Dim eraName As String
    Dim birthBody As String
    SplitWarekiBirthParts GetCtrlTextSafe(owner, "txtBirth"), GetCtrlTextSafe(owner, "txtAge"), eraName, birthBody

    WriteMerged ws, "A2:U2", BuildHeaderDate("作成日", FormatWarekiFull(GetCtrlTextSafe(owner, "txtEDate")))
    WriteMerged ws, "V2:AP2", BuildHeaderDate("前回作成日", FormatWarekiFull(GetPreviousCreatedDateText(owner)))
    WriteMerged ws, "AQ2:BJ2", BuildHeaderDate("初回作成", FormatWarekiFull(GetFirstCreatedDateText(owner)))


    WriteMerged ws, "E3:Q3", GetCtrlTextSafe(owner, "txtHdrKana")
    WriteMerged ws, "V3:AK3", eraName
    WriteMerged ws, "E4:Q4", GetCtrlTextSafe(owner, "txtName")
    WriteMerged ws, "V4:AK4", birthBody
    WriteMerged ws, "R4:U4", GetCtrlTextSafe(owner, "cboSex")
    WriteMerged ws, "AL4:AP4", GetCtrlTextSafe(owner, "cboCare")
    WriteMerged ws, "AQ3:BJ3", "計画作成者：" & GetCtrlTextSafe(owner, "txtEvaluator")
    WriteMerged ws, "AQ4:BJ4", "職種：" & GetCtrlTextSafe(owner, "txtEvaluatorJob")
    WriteMerged ws, "O5:AE5", GetCtrlTextSafe(owner, "cboElder")
    WriteMerged ws, "AS5:BJ5", GetCtrlTextSafe(owner, "cboDementia")

    WriteMerged ws, "A8:AE9", GetCtrlTextSafe(owner, "txtNeedsPt")
    WriteMerged ws, "AF8:BJ9", GetCtrlTextSafe(owner, "txtNeedsFam")
    WriteMerged ws, "A11:AE12", GetCtrlTextSafe(owner, "txtLiving")
    WriteMerged ws, "AF11:BJ12", BuildHomeEnvText(owner)

    WriteMerged ws, "D14:T14", GetCtrlTextSafe(owner, "txtDx")
    WriteMerged ws, "U14:BJ14", BuildMedicalDatesText(owner)

    Dim dbgTx As String: dbgTx = GetCtrlTextSafe(owner, "txtTxCourse")
    Dim dbgCp As String: dbgCp = GetCtrlTextSafe(owner, "txtComplications")
    Dim dbgLv As String: dbgLv = GetCtrlTextSafe(owner, "txtLiving")
    Debug.Print "[WriteEvalPlanSheet] owner=" & typeName(owner) & _
        " | txtLiving=[" & dbgLv & "]" & _
        " | txtTxCourse=[" & dbgTx & "]" & _
        " | txtComplications=[" & dbgCp & "]"
    WriteMerged ws, "A16:BJ16", dbgTx
    WriteMerged ws, "A18:BJ18", dbgCp
    WriteMerged ws, "A20:BJ20", GetPlanTextWithFallback(planData, owner, Array("TrainingPrecaution", "Medical.TrainingPrecaution", "機能訓練実施上の留意事項"), Array("txtTrainingNote", "txtRehabNote", "txtPrecaution"))

    WriteProgramBlocks ws, planData

    WriteMerged ws, "A50:AE51", GetPlanText(planData, Array("Monitoring.Change", "monitoring.change", "MonitoringChange", "changeText"))
    WriteMerged ws, "AF50:BJ51", GetPlanText(planData, Array("Monitoring.Issue", "monitoring.issue", "MonitoringIssue", "issueText"))


    Exit Sub
EH:
    Err.Clear


End Sub

' デバッグ用：テンプレートシートのラベル位置を Immediate に出力
Public Sub DebugScanPlanSheetLabels(ByVal ws As Worksheet)
    Dim keywords As Variant
    keywords = Array("性別", "要介護度", "障害高齢者", "認知症高齢者", "自立度", "介護度")
    Dim cell As Range
    Dim lastRow As Long: lastRow = 30
    Dim c As Long, r As Long
    For r = 1 To lastRow
        For c = 1 To 62 ' A to BJ
            On Error Resume Next
            Dim v As String: v = CStr(ws.Cells(r, c).value)
            On Error GoTo 0
            Dim i As Long
            For i = LBound(keywords) To UBound(keywords)
                If InStr(v, CStr(keywords(i))) > 0 Then
                    Debug.Print "Row=" & r & " Col=" & c & " (" & ws.Cells(r, c).Address(False, False) & ") = [" & v & "]"
                End If
            Next i
        Next c
    Next r
End Sub

Private Function GetPreviousCreatedDateText(ByVal owner As Object) As String
    On Error GoTo EH

    Dim wsEval As Worksheet
    If modEvalIOEntry.TryGetUserHistorySheet(owner, wsEval) Then
        GetPreviousCreatedDateText = modEvalIOEntry.GetPreviousEvalDateText(wsEval)
    End If
    Exit Function
EH:
    Err.Clear
End Function


Private Function GetFirstCreatedDateText(ByVal owner As Object) As String
    On Error GoTo EH

    Dim wsEval As Worksheet

    Dim firstEvalDate As String
    Dim latestEvalDate As String
    Dim previousEvalDate As String
    Dim recordCount As Long

        If modEvalIOEntry.TryGetUserHistorySheet(owner, wsEval) Then
        modEvalIOEntry.GetUserEvalDateStats wsEval, firstEvalDate, latestEvalDate, previousEvalDate, recordCount
        GetFirstCreatedDateText = firstEvalDate
    End If

   
    Exit Function
EH:
    Err.Clear
End Function



Private Sub WriteProgramBlocks(ByVal ws As Worksheet, ByVal planData As Object)
    Dim i As Long
    Dim startRow As Long
    Dim item As Variant

    For i = 1 To 5
        startRow = 29 + (i - 1) * 3
        item = GetProgramItem(planData, i)

        WriteMerged ws, "C" & startRow & ":AE" & (startRow + 2), GetProgramField(planData, item, i, Array("Content", "Program", "ProgramContent", "programContent"), Array("Program" & i & "Content"))
        WriteMerged ws, "AF" & startRow & ":AR" & (startRow + 2), GetProgramField(planData, item, i, Array("Note", "Notes", "Caution", "Consideration"), Array("Program" & i & "Note"))
        WriteMerged ws, "AS" & startRow & ":AX" & (startRow + 2), GetProgramField(planData, item, i, Array("Frequency", "Freq", "frequency"), Array("Program" & i & "Frequency"))
        WriteMerged ws, "AY" & startRow & ":BD" & (startRow + 2), GetProgramField(planData, item, i, Array("Time", "Duration", "time"), Array("Program" & i & "Time"))
        WriteMerged ws, "BE" & startRow & ":BJ" & (startRow + 2), GetProgramField(planData, item, i, Array("Performer", "Staff", "Executor", "staff"), Array("Program" & i & "Performer"))
    Next i
End Sub

Private Function GetProgramField(ByVal planData As Object, ByVal item As Variant, ByVal idx As Long, ByVal itemKeys As Variant, ByVal rootKeys As Variant) As String
    Dim s As String
    s = GetTextByKeys(item, itemKeys)
    If Len(s) > 0 Then
        GetProgramField = s
        Exit Function
    End If
    GetProgramField = GetTextByKeys(planData, rootKeys)
End Function

Private Function GetProgramItem(ByVal planData As Object, ByVal idx As Long) As Variant
    Dim programs As Variant
    programs = ResolvePath(planData, "Programs")
    If IsEmpty(programs) Then programs = ResolvePath(planData, "programs")
    If IsEmpty(programs) Then programs = ResolvePath(planData, "ProgramItems")
    If IsEmpty(programs) Then programs = ResolvePath(planData, "programItems")

    If IsEmpty(programs) Then Exit Function
    
    GetProgramItem = GetIndexValue(programs, idx)
End Function

Private Function GetIndexValue(ByVal src As Variant, ByVal idx As Long) As Variant
    On Error GoTo EH
    If IsObject(src) Then
        Dim t As String
        t = typeName(src)
        If StrComp(t, "Collection", vbTextCompare) = 0 Then
            If idx >= 1 And idx <= CLng(CallByName(src, "Count", VbGet)) Then
                GetIndexValue = CallByName(src, "Item", VbGet, idx)
            End If
            Exit Function
        End If
        
        GetIndexValue = CallByName(src, "Item", VbGet, idx)
        Exit Function
    End If

    If IsArray(src) Then
        Dim arrIdx As Long
        arrIdx = LBound(src) + idx - 1
        If arrIdx >= LBound(src) And arrIdx <= UBound(src) Then GetIndexValue = src(arrIdx)
    End If
    Exit Function
EH:
    Err.Clear
End Function

Private Function GetPlanTextWithFallback(ByVal planData As Object, ByVal owner As Object, ByVal planKeys As Variant, ByVal ctrlNames As Variant) As String
    Dim s As String
    s = GetPlanText(planData, planKeys)
    If Len(s) > 0 Then
        GetPlanTextWithFallback = s
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ctrlNames) To UBound(ctrlNames)
        s = GetCtrlTextSafe(owner, CStr(ctrlNames(i)))
        If Len(s) > 0 Then
            GetPlanTextWithFallback = s
            Exit Function
        End If
    Next i
End Function

Private Function BuildHeaderDate(ByVal labelText As String, ByVal formattedDate As String) As String
    If Len(Trim$(formattedDate)) = 0 Then Exit Function
    BuildHeaderDate = labelText & "：" & formattedDate
End Function

Private Function BuildMedicalDatesText(ByVal owner As Object) As String
    Dim onsetText As String
    Dim admText As String
    Dim disText As String

    onsetText = FormatDateForSentence(GetCtrlTextSafe(owner, "txtOnset"))
    admText = FormatDateForSentence(GetCtrlTextSafeAny(owner, "txtAdmDate", "txtHosp"))
    disText = FormatDateForSentence(GetCtrlTextSafeAny(owner, "txtDisDate", "txtDischarge"))

    BuildMedicalDatesText = "発症日・受傷日：" & onsetText & "  直近の入院日：" & admText & "  直近の退院日：" & disText
End Function

Private Function BuildHomeEnvText(ByVal owner As Object) As String
    Dim names As Variant
    names = TryGetHomeEnvControlNames()
    If IsEmpty(names) Then names = CollectHomeEnvCheckNames(owner)
    
    Dim labels As Collection
    Set labels = New Collection

    Dim i As Long
    If Not IsEmpty(names) Then
        For i = LBound(names) To UBound(names)
            Dim ctl As Object
            Set ctl = FindControlByName(owner, CStr(names(i)))
            If Not ctl Is Nothing Then
                If GetCheckValueSafe(ctl) Then
                    AddUniqueText labels, GetControlCaptionSafe(ctl)
                End If
            End If
        Next i
    End If
    

    If labels.count = 0 Then
        CollectHomeEnvCheckedCaptions owner, labels
    End If
    
    
    Dim text As String
    text = JoinCollection(labels, "、")

    Dim note As String
    note = GetCtrlTextSafeAny(owner, "txtBIHomeEnvNote", "txtHomeNote")
    If Len(note) > 0 Then
        If Len(text) > 0 Then
            text = text & "。備考：" & note
        Else
            text = "備考：" & note
        End If
    End If

    BuildHomeEnvText = text
End Function

Private Function TryGetHomeEnvControlNames() As Variant
    On Error Resume Next
    TryGetHomeEnvControlNames = Application.Run("HomeEnvControlNames")
    If Err.Number <> 0 Then
        Err.Clear
        TryGetHomeEnvControlNames = Application.Run("modEvalIOEntry.HomeEnvControlNames")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        TryGetHomeEnvControlNames = Empty
    End If
    On Error GoTo 0
End Function

Private Function CollectHomeEnvCheckNames(ByVal owner As Object) As Variant
    Dim names As Collection
    Set names = New Collection
    CollectHomeEnvCheckNamesFromContainer owner, names
    If names.count = 0 Then Exit Function

    Dim arr() As String

    Dim i As Long
    ReDim arr(0 To names.count - 1)
    For i = 1 To names.count
        arr(i - 1) = CStr(names(i))
    Next i
    CollectHomeEnvCheckNames = arr
End Function

Private Sub CollectHomeEnvCheckNamesFromContainer(ByVal container As Object, ByVal names As Collection)
    If ObjectIsNothingSafe(container) Then Exit Sub

    Dim controlsObj As Object
    Set controlsObj = GetControlsSafe(container)
    If controlsObj Is Nothing Then Exit Sub
    
    
    Dim ctl As Object
    For Each ctl In controlsObj
        If IsHomeEnvCheckControl(ctl) Then
            On Error Resume Next
            names.Add NzTextSafe(CallByName(ctl, "Name", VbGet)), NzTextSafe(CallByName(ctl, "Name", VbGet))
            Err.Clear
            On Error GoTo 0
        End If

        CollectHomeEnvCheckNamesFromContainer ctl, names
    Next ctl
End Sub

Private Sub CollectHomeEnvCheckedCaptions(ByVal container As Object, ByVal labels As Collection)
    If ObjectIsNothingSafe(container) Then Exit Sub
    
    Dim controlsObj As Object
    Set controlsObj = GetControlsSafe(container)
    If controlsObj Is Nothing Then Exit Sub

    Dim ctl As Object
    For Each ctl In controlsObj
        If IsHomeEnvCheckControl(ctl) Then
            If GetCheckValueSafe(ctl) Then AddUniqueText labels, GetControlCaptionSafe(ctl)
        End If
        CollectHomeEnvCheckedCaptions ctl, labels
    Next ctl
End Sub

Private Function IsHomeEnvCheckControl(ByVal ctl As Object) As Boolean
    If ObjectIsNothingSafe(ctl) Then Exit Function
    If StrComp(typeName(ctl), "CheckBox", vbTextCompare) <> 0 Then Exit Function

    Dim tagText As String
    tagText = GetControlTagSafe(ctl)
    If Len(tagText) < Len("BI.HomeEnv.") Then Exit Function
    IsHomeEnvCheckControl = (StrComp(Left$(tagText, Len("BI.HomeEnv.")), "BI.HomeEnv.", vbTextCompare) = 0)
End Function


Private Function FormatWarekiFull(ByVal dateText As String) As String
    Dim dt As Date
    If Not TryParseDate(dateText, dt) Then Exit Function

    Dim era As String
    Dim eraYear As Long
    ToWareki dt, era, eraYear
    If Len(era) = 0 Then Exit Function

    FormatWarekiFull = era & CStr(eraYear) & "年" & Month(dt) & "月" & day(dt) & "日"
End Function

Private Sub SplitWarekiBirthParts(ByVal birthText As String, ByVal ageText As String, ByRef eraName As String, ByRef bodyText As String)
    eraName = vbNullString
    bodyText = vbNullString

    Dim era As String, y As Long, m As Long, d As Long
    If ParseWarekiInput(birthText, era, y, m, d) Then
        eraName = era
        bodyText = CStr(y) & "年" & CStr(m) & "月" & CStr(d) & "日生"
    ElseIf IsDate(Trim$(birthText)) Then
        Dim dt As Date
        dt = CDate(Trim$(birthText))
        Dim eraY As Long
        ToWareki dt, eraName, eraY
        bodyText = CStr(eraY) & "年" & Month(dt) & "月" & day(dt) & "日生"
    Else
        eraName = ExtractEraName(birthText)
        bodyText = Trim$(RemoveEraPrefix(birthText))
        If Len(bodyText) > 0 Then bodyText = bodyText & "生"
    End If

    If Len(Trim$(ageText)) > 0 Then
        If Len(bodyText) > 0 Then
            bodyText = bodyText & "（" & Trim$(ageText) & "歳）"
        Else
            bodyText = "（" & Trim$(ageText) & "歳）"
        End If
    End If
End Sub

Private Function ParseWarekiInput(ByVal src As String, ByRef era As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim s As String
    s = Trim$(NzTextSafe(src))
    If Len(s) = 0 Then Exit Function

    era = ExtractEraName(s)
    If Len(era) = 0 Then Exit Function

    Dim nums As Variant
    nums = ExtractNumbers(s)
    On Error GoTo EH
    If UBound(nums) < 2 Then Exit Function

    y = CLng(nums(0))
    m = CLng(nums(1))
    d = CLng(nums(2))
    ParseWarekiInput = True
    Exit Function
EH:
    Err.Clear
End Function

Private Function ExtractNumbers(ByVal s As String) As Variant
    Dim values() As Long
    Dim count As Long
    Dim i As Long
    Dim buf As String

    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then
            buf = buf & ch
            
        ElseIf Len(buf) > 0 Then
            ReDim Preserve values(0 To count)
            values(count) = CLng(buf)
            count = count + 1
            buf = vbNullString
        End If
    Next i

    If Len(buf) > 0 Then
        ReDim Preserve values(0 To count)
        values(count) = CLng(buf)
    End If

    If count = 0 And Len(buf) = 0 Then
        ExtractNumbers = Array()
    Else
        ExtractNumbers = values
    End If
End Function

Private Function ExtractEraName(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))

    If InStr(1, s, "令和", vbTextCompare) = 1 Or Left$(t, 1) = "R" Then
        ExtractEraName = "令和"
    ElseIf InStr(1, s, "平成", vbTextCompare) = 1 Or Left$(t, 1) = "H" Then
        ExtractEraName = "平成"
    ElseIf InStr(1, s, "昭和", vbTextCompare) = 1 Or Left$(t, 1) = "S" Then
        ExtractEraName = "昭和"
    ElseIf InStr(1, s, "大正", vbTextCompare) = 1 Or Left$(t, 1) = "T" Then
        ExtractEraName = "大正"
    ElseIf InStr(1, s, "明治", vbTextCompare) = 1 Or Left$(t, 1) = "M" Then
        ExtractEraName = "明治"
    End If
End Function

Private Function RemoveEraPrefix(ByVal s As String) As String
    s = Trim$(NzTextSafe(s))
    Dim era As String
    era = ExtractEraName(s)

    s = Trim$(NzTextSafe(s))

    If Len(s) > 0 Then
        Dim head As String
        head = UCase$(Left$(s, 1))
        If head = "R" Or head = "H" Or head = "S" Or head = "T" Or head = "M" Then
            RemoveEraPrefix = Trim$(Mid$(s, 2))
            Exit Function
        End If
    End If

    RemoveEraPrefix = s
End Function

Private Sub ToWareki(ByVal dt As Date, ByRef era As String, ByRef eraYear As Long)
    If dt >= DateSerial(2019, 5, 1) Then
        era = "令和": eraYear = Year(dt) - 2018
    ElseIf dt >= DateSerial(1989, 1, 8) Then
        era = "平成": eraYear = Year(dt) - 1988
    ElseIf dt >= DateSerial(1926, 12, 25) Then
        era = "昭和": eraYear = Year(dt) - 1925
    ElseIf dt >= DateSerial(1912, 7, 30) Then
        era = "大正": eraYear = Year(dt) - 1911
    ElseIf dt >= DateSerial(1868, 1, 25) Then
        era = "明治": eraYear = Year(dt) - 1867
    Else
        era = vbNullString: eraYear = 0
    End If
End Sub

Private Function TryParseDate(ByVal src As String, ByRef dt As Date) As Boolean
    Dim s As String
    s = Trim$(NzTextSafe(src))
    If Len(s) = 0 Then Exit Function

    On Error Resume Next
    dt = CDate(s)
    TryParseDate = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FormatDateForSentence(ByVal src As String) As String
    Dim dt As Date
    If TryParseDate(src, dt) Then
        FormatDateForSentence = Year(dt) & "年" & Month(dt) & "月" & day(dt) & "日"
    Else
        FormatDateForSentence = Trim$(NzTextSafe(src))
    End If
End Function

Private Function GetCtrlTextSafeAny(ByVal owner As Object, ParamArray names() As Variant) As String
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim s As String
        s = GetCtrlTextSafe(owner, NzTextSafe(names(i)))
        If Len(s) > 0 Then
            GetCtrlTextSafeAny = s
            Exit Function
        End If
    Next i
End Function

Private Function GetCtrlTextSafe(ByVal owner As Object, ByVal ctrlName As String) As String
    Dim ctl As Object
    Set ctl = FindControlByName(owner, ctrlName)
    If ctl Is Nothing Then Exit Function
    
    GetCtrlTextSafe = GetControlTextSafe(ctl)
End Function


Private Function GetControlTextSafe(ByVal ctl As Object) As String
    If ObjectIsNothingSafe(ctl) Then Exit Function
    
    On Error Resume Next
    GetControlTextSafe = NzTextSafe(CallByName(ctl, "Text", VbGet))
    If Len(GetControlTextSafe) = 0 Then GetControlTextSafe = NzTextSafe(CallByName(ctl, "Value", VbGet))
    If Len(GetControlTextSafe) = 0 Then GetControlTextSafe = NzTextSafe(CallByName(ctl, "Caption", VbGet))
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlValueSafe(ByVal ctl As Object) As Variant
    If ObjectIsNothingSafe(ctl) Then Exit Function
    On Error Resume Next
    GetControlValueSafe = CallByName(ctl, "Value", VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        GetControlValueSafe = Empty
    End If
    On Error GoTo 0
End Function

Private Function FindControlByName(ByVal container As Object, ByVal ctrlName As String) As Object
    On Error GoTo SafeExit
    If ObjectIsNothingSafe(container) Then Exit Function

    Dim thisName As String
    thisName = NzTextSafe(GetMemberValue(container, "Name"))
    If Len(thisName) > 0 Then
        If StrComp(thisName, ctrlName, vbTextCompare) = 0 Then
            Set FindControlByName = container
            Exit Function
        End If
    End If


    Dim pagesObj As Object
    Set pagesObj = GetPagesSafe(container)
    If Not pagesObj Is Nothing Then
        Dim pg As Object
        For Each pg In pagesObj
            Set FindControlByName = FindControlByName(pg, ctrlName)
            If Not FindControlByName Is Nothing Then Exit Function
        Next pg
    End If
    Err.Clear

    Dim controlsObj As Object
    Set controlsObj = GetControlsSafe(container)
    If controlsObj Is Nothing Then Exit Function

    Dim c As Variant
    Dim ctl As Object
    For Each ctl In controlsObj
        Set FindControlByName = FindControlByName(ctl, ctrlName)
        If Not FindControlByName Is Nothing Then Exit Function
    Next ctl

SafeExit:
    Err.Clear
End Function

Private Function GetControlsSafe(ByVal container As Object) As Object
    If ObjectIsNothingSafe(container) Then Exit Function
    On Error Resume Next
    Set GetControlsSafe = CallByName(container, "Controls", VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        Set GetControlsSafe = Nothing
    End If
    On Error GoTo 0
End Function

Private Function GetPagesSafe(ByVal container As Object) As Object
    If ObjectIsNothingSafe(container) Then Exit Function
    On Error Resume Next
    Set GetPagesSafe = CallByName(container, "Pages", VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        Set GetPagesSafe = Nothing
    End If
    On Error GoTo 0
End Function


Private Function GetCheckValueSafe(ByVal ctl As Object) As Boolean
    Dim v As Variant
    v = GetControlValueSafe(ctl)

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    On Error Resume Next
    GetCheckValueSafe = CBool(v)
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlCaptionSafe(ByVal ctl As Object) As String
    If ObjectIsNothingSafe(ctl) Then Exit Function
    On Error Resume Next
    GetControlCaptionSafe = NzTextSafe(CallByName(ctl, "Caption", VbGet))
    If Len(GetControlCaptionSafe) = 0 Then GetControlCaptionSafe = NzTextSafe(CallByName(ctl, "Name", VbGet))
    Err.Clear
    On Error GoTo 0
End Function


Private Function GetControlTagSafe(ByVal ctl As Object) As String
    If ObjectIsNothingSafe(ctl) Then Exit Function
    On Error Resume Next
    GetControlTagSafe = NzTextSafe(CallByName(ctl, "Tag", VbGet))
    Err.Clear
    On Error GoTo 0
End Function

Private Sub AddUniqueText(ByVal col As Collection, ByVal s As String)
    s = Trim$(NzTextSafe(s))
    If Len(s) = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To col.count
        If StrComp(CStr(col(i)), s, vbTextCompare) = 0 Then Exit Sub
    Next i
    col.Add s
End Sub

Private Function JoinCollection(ByVal col As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To col.count
        If Len(JoinCollection) > 0 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(col(i))
    Next i
End Function

Private Sub WriteMerged(ByVal ws As Worksheet, ByVal addressText As String, ByVal text As String)
    ' 指定範囲の先頭行以降に先頭セルがある結合セルを探して書き込む
    ' （例: A15:A16 が結合済みの場合、A16 から始まる範囲では B16 以降に書く）
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(addressText)
    If rng Is Nothing Then GoTo done

    Dim baseRow As Long: baseRow = rng.Cells(1, 1).row
    Dim cell As Range
    For Each cell In rng.Cells
        Dim top As Range
        Set top = cell
        If cell.MergeCells Then Set top = cell.MergeArea.Cells(1, 1)
        If top.row >= baseRow Then
            top.value = NzTextSafe(text)
            GoTo done
        End If
    Next cell
done:
    Err.Clear
    On Error GoTo 0
End Sub

Private Function NzTextSafe(ByVal v As Variant, Optional ByVal fallback As String = vbNullString) As String
    On Error GoTo EH
    If IsError(v) Then
        NzTextSafe = fallback
    ElseIf IsNull(v) Then
        NzTextSafe = fallback
    ElseIf IsEmpty(v) Then
        NzTextSafe = fallback
    ElseIf IsObject(v) Then
        If ObjectIsNothingSafe(v) Then
            NzTextSafe = fallback
        Else
            NzTextSafe = fallback
        End If
    Else

        NzTextSafe = CStr(v)
    End If
    Exit Function
EH:
    NzTextSafe = fallback
    Err.Clear
End Function

Private Function GetPlanText(ByVal planData As Object, ByVal paths As Variant) As String
    If ObjectIsNothingSafe(planData) Then Exit Function

    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        Dim v As Variant
        v = ResolvePath(planData, NzTextSafe(paths(i)))
        If Not IsEmpty(v) Then
            GetPlanText = NzTextSafe(v)
            If Len(GetPlanText) > 0 Then Exit Function
        End If
    Next i
End Function

Private Function GetTextByKeys(ByVal source As Variant, ByVal keys As Variant) As String
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim v As Variant
        v = ResolvePath(source, NzTextSafe(keys(i)))
        If Not IsEmpty(v) Then
            GetTextByKeys = NzTextSafe(v)
            If Len(GetTextByKeys) > 0 Then Exit Function
        End If
    Next i
End Function

Private Function ResolvePath(ByVal source As Variant, ByVal path As String) As Variant
    If Len(Trim$(path)) = 0 Then Exit Function
    If IsObject(source) Then
        If ObjectIsNothingSafe(source) Then Exit Function
    ElseIf IsEmpty(source) Or IsNull(source) Or IsError(source) Then
        Exit Function
    End If
    
    
    Dim cur As Variant
    cur = source

    Dim parts() As String
    parts = Split(path, ".")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        cur = GetMemberValue(cur, parts(i))
        If IsEmpty(cur) Then Exit Function
    Next i

    ResolvePath = cur
End Function

Private Function GetMemberValue(ByVal source As Variant, ByVal memberName As String) As Variant
    If Len(Trim$(memberName)) = 0 Then Exit Function
    If IsError(source) Or IsNull(source) Or IsEmpty(source) Then Exit Function
    If Not IsObject(source) Then Exit Function
    If ObjectIsNothingSafe(source) Then Exit Function

    On Error Resume Next
    GetMemberValue = CallByName(source, memberName, VbGet)

    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If


    Err.Clear
    GetMemberValue = CallByName(source, "Item", VbGet, memberName)
    If Err.Number <> 0 Then
        Err.Clear
        GetMemberValue = Empty
    End If
    On Error GoTo 0
End Function

Private Function ObjectIsNothingSafe(ByVal obj As Object) As Boolean
    On Error GoTo EH
    ObjectIsNothingSafe = (obj Is Nothing)
    Exit Function
EH:
    ObjectIsNothingSafe = True
    Err.Clear
End Function
