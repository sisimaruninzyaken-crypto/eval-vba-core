Attribute VB_Name = "modPlanGen"


Option Explicit

' [蠢・・ BASIC_KEYS_V1 縺ｮ譌｢蟄假ｼ医≠縺ｪ縺溘・迴ｾ陦後・縺ｾ縺ｾ・・
Public Function BasicKeysV1() As Variant
 BasicKeysV1 = Split("care_level_band|primary_condition_cat|comorbidity_cat_list|history_flags|living_type|support_availability|bi_total|bi_low_items|iadl_limits|bed_mobility_band|rom_limit_tags|strength_band|pain_band|pain_site_tags|needs_patient|needs_family", "|")
End Function

' [蠢・・ BuildBasicInputV1・・ed_mobility_band 繧定ｿｽ蜉貂医∩・・
Public Function BuildBasicInputV1() As String
    Dim keys As Variant
    Dim values() As String
    Dim lines() As String
    Dim i As Long

    Dim careLevelBand As String
    Dim primaryConditionCat As String
    Dim livingType As String
    Dim bedMobilityBand As String
    Dim painBand As String
    Dim painSiteTags As String
    Dim iadlLimits As String
    Dim biLowItems As String
    Dim romLimitTags As String
    Dim strengthBand As String
    Dim needsPatient As String
    Dim needsFamily As String

    keys = BasicKeysV1()
    ReDim values(LBound(keys) To UBound(keys))
    ReDim lines(LBound(keys) To UBound(keys))

    careLevelBand = GetFrmEvalControlText("cboCare")
    primaryConditionCat = GetPrimaryConditionCat()
    livingType = GetFrmEvalControlText("txtLiving")
    iadlLimits = GetIADLLimits()
    biLowItems = GetBILowItems()
    romLimitTags = GetRomLimitTags()

    ' 笘・ｿｽ蜉・喘ed_mobility_band 縺ｯ1蝗槭□縺題ｨ育ｮ・
    bedMobilityBand = GetBedMobilityBand()
    painBand = GetPainBand()
    painSiteTags = GetPainSiteTags()
    strengthBand = GetStrengthBand()
    needsPatient = GetLatestBasicTextByHeader("謔｣閠・eeds")
    needsFamily = GetLatestBasicTextByHeader("螳ｶ譌蒐eeds")


    For i = LBound(keys) To UBound(keys)
        Select Case CStr(keys(i))
            Case "care_level_band"
                values(i) = careLevelBand

            Case "living_type"
                values(i) = livingType
                
            Case "primary_condition_cat"
                values(i) = primaryConditionCat

            Case "bi_total"
                values(i) = GetFrmEvalControlText("txtBITotal")
            Case "iadl_limits"
                values(i) = iadlLimits
            Case "bi_low_items"
                values(i) = biLowItems


            ' 笘・ｿｽ蜉・喘ed_mobility_band
            Case "bed_mobility_band"
                values(i) = bedMobilityBand
            Case "pain_band"
                values(i) = painBand
            Case "pain_site_tags"
                values(i) = painSiteTags
            Case "rom_limit_tags"
                values(i) = romLimitTags
            Case "strength_band"
                values(i) = strengthBand
            Case "needs_patient"
                values(i) = needsPatient
            Case "needs_family"
                values(i) = needsFamily

            Case Else
                values(i) = vbNullString
        End Select

        lines(i) = CStr(keys(i)) & ": " & values(i)
    Next i


BuildBasicInputV1 = Join(lines, vbCrLf)
End Function

Private Function GetLatestBasicTextByHeader(ByVal headerName As String) As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    GetLatestBasicTextByHeader = Trim$(ReadStr_Compat(headerName, rLatest, ws))
End Function



Private Function GetPrimaryConditionCat() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    GetPrimaryConditionCat = Trim$(ReadStr_Compat("荳ｻ險ｺ譁ｭ", rLatest, ws))
End Function

Private Function GetStrengthBand() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long
    Dim mmtIO As String

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then
        GetStrengthBand = "unknown"
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then
        GetStrengthBand = "unknown"
        Exit Function
    End If

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then
        GetStrengthBand = "unknown"
        Exit Function
    End If

    mmtIO = ReadStr_Compat("MMT_IO", rLatest, ws)
    GetStrengthBand = ComputeStrengthBandFromMMTIO(mmtIO)
End Function

Public Function ComputeStrengthBandFromMMTIO(ByVal mmtIO As String) As String
    Dim targetItems As Object
    Dim itemScores As Object
    Dim records As Variant
    Dim fields As Variant
    Dim rec As Variant
    Dim itemName As String
    Dim partFlag As String
    Dim score As Double

    Set targetItems = CreateObject("Scripting.Dictionary")
    targetItems.CompareMode = vbTextCompare
    targetItems("閧｡螻域峇") = True
    targetItems("閧｡莨ｸ螻・) = True
    targetItems("閧｡螟冶ｻ｢]") = True
    targetItems("閹昜ｼｸ螻・) = True
    targetItems("雜ｳ髢｢遽閭悟ｱ・") = True
    targetItems("雜ｳ髢｢遽蠎募ｱ・) = True
    targetItems("豈崎ｶｾ莨ｸ螻・) = True

    Set itemScores = CreateObject("Scripting.Dictionary")
    itemScores.CompareMode = vbTextCompare

    If LenB(Trim$(mmtIO)) = 0 Then
        ComputeStrengthBandFromMMTIO = "unknown"
        Exit Function
    End If

    records = Split(mmtIO, ";")
    For Each rec In records
        rec = Trim$(CStr(rec))
        If LenB(rec) = 0 Then GoTo ContinueRecord

        fields = Split(CStr(rec), "|")
        If UBound(fields) < 3 Then GoTo ContinueRecord

        partFlag = Trim$(CStr(fields(0)))
        If StrComp(partFlag, "1", vbBinaryCompare) <> 0 Then GoTo ContinueRecord

        itemName = Trim$(CStr(fields(1)))
        If Not targetItems.exists(itemName) Then GoTo ContinueRecord

        If TryGetLowerScore(fields, score) Then
            itemScores(itemName) = score
        End If

ContinueRecord:
    Next rec

    ComputeStrengthBandFromMMTIO = ScoreToStrengthBand(itemScores)
End Function

Private Function TryGetLowerScore(ByVal fields As Variant, ByRef score As Double) As Boolean
    Dim hasR As Boolean
    Dim hasL As Boolean
    Dim rVal As Double
    Dim lVal As Double

    hasR = ParseNumericScore(CStr(fields(2)), rVal)
    hasL = ParseNumericScore(CStr(fields(3)), lVal)

    If hasR And hasL Then
        score = IIf(rVal <= lVal, rVal, lVal)
        TryGetLowerScore = True
    ElseIf hasR Then
        score = rVal
        TryGetLowerScore = True
    ElseIf hasL Then
        score = lVal
        TryGetLowerScore = True
    End If
End Function

Private Function ParseNumericScore(ByVal src As String, ByRef outVal As Double) As Boolean
    src = Trim$(src)
    If LenB(src) = 0 Then Exit Function
    If Not IsNumeric(src) Then Exit Function

    outVal = CDbl(src)
    ParseNumericScore = True
End Function

Private Function ScoreToStrengthBand(ByVal itemScores As Object) As String
    Dim k As Variant
    Dim total As Double
    Dim avg As Double

    If itemScores Is Nothing Then
        ScoreToStrengthBand = "unknown"
        Exit Function
    End If
    If itemScores.count = 0 Then
        ScoreToStrengthBand = "unknown"
        Exit Function
    End If

    For Each k In itemScores.keys
        total = total + CDbl(itemScores(k))
    Next k

    avg = total / CDbl(itemScores.count)

    If avg < 2# Then
        ScoreToStrengthBand = "severe"
    ElseIf avg < 3# Then
        ScoreToStrengthBand = "moderate"
    ElseIf avg < 4# Then
        ScoreToStrengthBand = "mild"
    ElseIf avg <= 5# Then
        ScoreToStrengthBand = "normal"
    Else
        ScoreToStrengthBand = "unknown"
    End If
End Function


Private Function GetRomLimitTags() As String
    Dim ws As Worksheet
    Dim look As Object
    Dim nm As String
    Dim rLatest As Long
    Dim tags As Collection

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    Set look = BuildHeaderLookup(ws)
    Set tags = New Collection

    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Ankle_Dorsi_R", 10, "Ankle_DF_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Ankle_Dorsi_L", 10, "Ankle_DF_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Ankle_Plantar_R", 30, "Ankle_PF_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Ankle_Plantar_L", 30, "Ankle_PF_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Knee_Flex_R", 115, "Knee_Flex_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Knee_Flex_L", 115, "Knee_Flex_L", tags
    AddRomLimitTagGE ws, look, rLatest, "ROM_Lower_Knee_Ext_R", 10, "Knee_Ext_R", tags
    AddRomLimitTagGE ws, look, rLatest, "ROM_Lower_Knee_Ext_L", 10, "Knee_Ext_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Flex_R", 100, "Hip_Flex_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Flex_L", 100, "Hip_Flex_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Ext_R", 5, "Hip_Ext_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Ext_L", 5, "Hip_Ext_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Abd_R", 30, "Hip_Abd_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Abd_L", 30, "Hip_Abd_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Add_R", 20, "Hip_Add_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Lower_Hip_Add_L", 20, "Hip_Add_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_Flex_R", 150, "Shoulder_Flex_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_Flex_L", 150, "Shoulder_Flex_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_Abd_R", 150, "Shoulder_Abd_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_Abd_L", 150, "Shoulder_Abd_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_ER_R", 45, "Shoulder_ER_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_ER_L", 45, "Shoulder_ER_L", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_IR_R", 50, "Shoulder_IR_R", tags
    AddRomLimitTag ws, look, rLatest, "ROM_Upper_Shoulder_IR_L", 50, "Shoulder_IR_L", tags
    
    If tags.count = 0 Then Exit Function

    Dim arr() As String
    Dim i As Long
    ReDim arr(0 To tags.count - 1)
    For i = 1 To tags.count
        arr(i - 1) = CStr(tags(i))
    Next i

    GetRomLimitTags = Join(arr, ", ")
End Function

Private Sub AddRomLimitTag( _
    ByVal ws As Worksheet, _
    ByVal look As Object, _
    ByVal rowNum As Long, _
    ByVal headerKey As String, _
    ByVal threshold As Double, _
    ByVal tag As String, _
    ByRef tags As Collection)

    Dim col As Long
    Dim v As Variant

    col = ResolveColumn(look, headerKey)
    If col <= 0 Then Exit Sub

    v = ws.Cells(rowNum, col).value
    If Not IsNumeric(v) Then Exit Sub
    If CDbl(v) <= threshold Then tags.Add tag
End Sub

Private Sub AddRomLimitTagGE( _
    ByVal ws As Worksheet, _
    ByVal look As Object, _
    ByVal rowNum As Long, _
    ByVal headerKey As String, _
    ByVal threshold As Double, _
    ByVal tag As String, _
    ByRef tags As Collection)

    Dim col As Long
    Dim v As Variant

    col = ResolveColumn(look, headerKey)
    If col <= 0 Then Exit Sub

    v = ws.Cells(rowNum, col).value
    If Not IsNumeric(v) Then Exit Sub
    If CDbl(v) >= threshold Then tags.Add tag
End Sub


Private Function GetPainSiteTags() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long
    Dim io As String
    Dim s As String
    Dim rawTags As Variant
    Dim tags() As String
    Dim i As Long
    Dim n As Long
    Dim v As String

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    io = ReadStr_Compat("IO_Pain", rLatest, ws)
    s = Trim$(IO_GetVal(io, "PainSite"))
    If LenB(s) = 0 Then Exit Function

    rawTags = Split(s, "/")
    ReDim tags(LBound(rawTags) To UBound(rawTags))

    n = -1
    For i = LBound(rawTags) To UBound(rawTags)
        v = Trim$(CStr(rawTags(i)))
        If LenB(v) > 0 Then
            n = n + 1
            tags(n) = v
        End If
    Next i

    If n < 0 Then Exit Function

    ReDim Preserve tags(0 To n)
    GetPainSiteTags = Join(tags, ", ")
End Function



Private Function GetPainBand() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long
    Dim io As String
    Dim vasStr As String
    Dim vas As Double

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    io = ReadStr_Compat("IO_Pain", rLatest, ws)
    vasStr = Trim$(IO_GetVal(io, "VAS"))

    If LenB(vasStr) = 0 Then Exit Function
    If Not IsNumeric(vasStr) Then Exit Function

    vas = CDbl(vasStr)
    If vas < 0 Or vas > 100 Then Exit Function

    Select Case vas
        Case 0
            GetPainBand = "縺ｪ縺・
        Case 1 To 30
            GetPainBand = "霆ｽ蠎ｦ"
        Case 31 To 60
            GetPainBand = "荳ｭ遲牙ｺｦ"
        Case Else
            GetPainBand = "驥榊ｺｦ"
    End Select
End Function


' IADL_0`IADL_8 ?uv?OA??
Private Function GetIADLLimits() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long
    Dim io As String
    Dim parts() As String
    Dim i As Long
    Dim n As Long
    Dim v As String

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EvalData")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    io = ReadStr_Compat("IO_ADL", rLatest, ws)

    ReDim parts(0 To 8)
    n = -1
   For i = 0 To 8
    v = Trim$(IO_GetVal(io, "IADL_" & CStr(i)))
    If LenB(v) > 0 Then
        If StrComp(v, "閾ｪ遶・, vbTextCompare) = 0 Then GoTo NextI
        n = n + 1
        parts(n) = "IADL_" & CStr(i) & "=" & v
    End If
NextI:
Next i


    If n < 0 Then Exit Function

    ReDim Preserve parts(0 To n)
    GetIADLLimits = Join(parts, ", ")

End Function



Private Function GetBILowItems() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim rLatest As Long
    Dim io As String
    Dim fullScores(0 To 9) As Double
    Dim i As Long
    Dim v As String
    Dim parts As Collection

    fullScores(0) = 10
    fullScores(1) = 15
    fullScores(2) = 5
    fullScores(3) = 10
    fullScores(4) = 5
    fullScores(5) = 15
    fullScores(6) = 10
    fullScores(7) = 10
    fullScores(8) = 10
    fullScores(9) = 10

    nm = Trim$(GetFrmEvalControlText("txtName"))
    If Len(nm) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EVAL_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    rLatest = FindLatestRowByName(ws, nm)
    If rLatest <= 0 Then Exit Function

    io = ReadStr_Compat("IO_ADL", rLatest, ws)
    If Len(io) = 0 Then Exit Function

    Set parts = New Collection

    For i = 0 To 9
        v = Trim$(IO_GetVal(io, "BI_" & CStr(i)))
        If Len(v) = 0 Then GoTo NextI

        If IsNumeric(v) Then
            If CDbl(v) = fullScores(i) Then GoTo NextI
        End If

        parts.Add "BI_" & CStr(i) & "=" & v
NextI:
    Next i

    If parts.count = 0 Then Exit Function

    Dim arr() As String
    ReDim arr(0 To parts.count - 1)
    For i = 1 To parts.count
        arr(i - 1) = CStr(parts(i))
    Next i

    GetBILowItems = Join(arr, ", ")
End Function

' 笘・ｽｮ謠幢ｼ唔O_ADL・・yo_*・峨°繧・band 逕滓・・域怙驥榊ｺｦ蜆ｪ蜈茨ｼ・
Private Function GetBedMobilityBand() As String
    Dim ws As Worksheet
    Dim nm As String
    Dim r As Long
    Dim raw As String
    Dim v As String

    Set ws = ThisWorkbook.Worksheets("EvalData")

    nm = Trim$(frmEval.controls("txtName").text)
    If LenB(nm) = 0 Then Exit Function

    ' 窶ｻ縺ゅ↑縺溘・譌｢蟄・FindLatestRowByName 繧剃ｽｿ縺・燕謠撰ｼ・ormalizeName蜷ｫ繧迚医〒OK・・
    r = FindLatestRowByName(ws, nm)
    If r <= 1 Then Exit Function

    ' 窶ｻ譌｢蟄倥・ ReadStr_Compat / IO_GetVal 繧剃ｽｿ逕ｨ
    raw = ReadStr_Compat("IO_ADL", r, ws)
    If LenB(raw) = 0 Then Exit Function

    v = IO_GetVal(raw, "Kyo_Roll") & "|" & _
        IO_GetVal(raw, "Kyo_SitUp") & "|" & _
        IO_GetVal(raw, "Kyo_SitHold") & "|" & _
        IO_GetVal(raw, "Kyo_StandUp") & "|" & _
        IO_GetVal(raw, "Kyo_StandHold")

    If InStr(v, "蜈ｨ莉句勧") > 0 Then
        GetBedMobilityBand = "蜈ｨ莉句勧"
    ElseIf InStr(v, "荳驛ｨ莉句勧") > 0 Then
        GetBedMobilityBand = "荳驛ｨ莉句勧"
    ElseIf InStr(v, "隕句ｮ医ｊ") > 0 Then
        GetBedMobilityBand = "隕句ｮ医ｊ"
    ElseIf Len(Replace$(v, "|", vbNullString)) > 0 Then
        GetBedMobilityBand = "閾ｪ遶・
    End If
End Function

' --- 莉･荳九・譌｢蟄倥・縺ｾ縺ｾ・医≠縺ｪ縺溘・迴ｾ陦後さ繝ｼ繝峨ｒ邯ｭ謖・ｼ・---

Private Function GetFrmEvalControlText(ByVal controlName As String) As String
    Dim target As Object

    Set target = FindControlRecursive(frmEval, controlName)
    If target Is Nothing Then Exit Function

    Select Case TypeName(target)
        Case "TextBox", "ComboBox"
            If LenB(CStr(target.value)) > 0 Then
                GetFrmEvalControlText = CStr(target.value)
            End If
        Case Else
            GetFrmEvalControlText = vbNullString
    End Select
End Function

Public Function FindControlRecursive(ByVal container As Object, ByVal targetName As String) As Object
    Dim c As Object
    Dim child As Object

    If container Is Nothing Then Exit Function

    On Error Resume Next
    Set FindControlRecursive = container.controls(targetName)
    On Error GoTo 0
    If Not FindControlRecursive Is Nothing Then Exit Function

    ' 窶ｻ縺薙％莉･髯阪ｂ縺ゅ↑縺溘・譌｢蟄伜ｮ溯｣・・縺ｾ縺ｾ
    ' ・・asChildControls 繧剃ｽｿ縺・沿縺ｪ繧峨◎縺ｮ縺ｾ縺ｾ邯ｭ謖√〒OK・・
    ' ...
End Function

' UniqueLinesInOrder / FindLatestRowByName / ReadStr_Compat / IO_GetVal 縺ｪ縺ｩ繧よ里蟄倥・縺ｾ縺ｾ


