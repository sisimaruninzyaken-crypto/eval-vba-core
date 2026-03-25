Attribute VB_Name = "modHeaderMap"

'=== Legacy/New Header Resolver (隱ｭ縺ｿ霎ｼ縺ｿ莠呈鋤繝槭ャ繝・ ===
Public Function ResolveLegacyHeader(ByVal wantName As String) As String
    ' 隱ｭ縺ｿ霎ｼ縺ｿ蛛ｴ縺梧爾縺呎眠蜷・竊・迴ｾ陦御ｿ晏ｭ倥・譌ｧ蜷・縺ｸ繝槭ャ繝・
    ' 隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ縺昴・縺ｾ縺ｾ霑斐☆・亥酔蜷堺ｿ晏ｭ倥↓蟇ｾ蠢懶ｼ・
    Select Case LCase$(wantName)
        Case "io_sensory":   ResolveLegacyHeader = "IO_Sensory"
        Case "io_testeval": ResolveLegacyHeader = "IO_TestEval"
        Case "io_mmt":       ResolveLegacyHeader = "MMT_IO"
        Case "io_rom":       ResolveLegacyHeader = "ROM_*"      '窶ｻ隍・焚蛻励ょｾ檎ｶ壹〒繝ｯ繧､繝ｫ繝峨き繝ｼ繝牙ｱ暮幕
        Case "io_adl":       ResolveLegacyHeader = "IO_ADL"
        Case "io_tone":      ResolveLegacyHeader = "TONE_IO"
        Case Else:           ResolveLegacyHeader = wantName
    End Select
End Function




'=== Header Column Resolver (譁ｰ譌ｧ繝倥ャ繝蜷阪ｒ蜷ｸ蜿弱＠縺ｦ蛻礼分蜿ｷ繧定ｿ斐☆) ===
Public Function HeaderCol(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim h As String: h = ResolveLegacyHeader(wantName)
    On Error Resume Next
    HeaderCol = Application.Match(h, ws.rows(1), 0)
    On Error GoTo 0
End Function



'=== Read String by Header (譁ｰ譌ｧ繝倥ャ繝蟇ｾ蠢懊〒繧ｻ繝ｫ譁・ｭ怜・蜿門ｾ・ ===
Public Function ReadStr(ByVal wantName As String, ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim c As Long: c = HeaderCol(wantName, ws)
    If c > 0 Then ReadStr = CStr(ws.Cells(r, c).value) Else ReadStr = vbNullString
End Function



'=== Sensory IO accessor (菫晏ｭ伜､縺ｮ蜿門ｾ暦ｼ壽眠譌ｧ繝倥ャ繝蜷ｸ蜿取ｸ・ ===
Public Function GetSavedSensoryIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    GetSavedSensoryIO = ReadStr("IO_Sensory", r, ws)
End Function




'=== ROM IO accessor (菫晏ｭ伜､縺ｮ蜿門ｾ暦ｼ夊ｦ句・縺・"ROM_*" 繧呈ｨｪ譁ｭ) ===
Public Function GetSavedROMIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim buf As String
    Dim m As Variant
    Dim c As Long
    Dim h As String, v As String

    ' 1) IO_ROM 蛻励′縺ゅｌ縺ｰ縲√◎繧後ｒ譛蜆ｪ蜈医〒菴ｿ縺・
    m = Application.Match("IO_ROM", ws.rows(1), 0)
    If Not IsError(m) Then
        buf = CStr(ws.Cells(r, CLng(m)).value)
        If Len(buf) > 0 Then
            GetSavedROMIO = buf
            Exit Function
        End If
    End If

    ' 2) 繝ｬ繧ｬ繧ｷ繝ｼ莠呈鋤・啌OM_* 蛻励°繧臥ｵ・∩遶九※・育ｯ・峇繧・160?213 縺ｫ髯仙ｮ夲ｼ・
    '    窶ｻ 214莉･髯阪・驥崎､⑲OM蛻励・辟｡隕悶＠縺ｦ證ｴ襍ｰ騾｣邨舌ｒ髦ｲ縺・
    For c = 160 To 213
        h = CStr(ws.Cells(1, c).value)
        If LCase$(Left$(h, 4)) = "rom_" Then
            v = CStr(ws.Cells(r, c).value)
            If Len(v) > 0 Then
                If Len(buf) > 0 Then buf = buf & "|"
                buf = buf & h & "=" & v
            End If
        End If
    Next c

    GetSavedROMIO = buf   ' 蟇ｾ雎｡縺檎┌縺代ｌ縺ｰ遨ｺ譁・ｭ・
End Function




'=== ADL IO accessor (菫晏ｭ伜､縺ｮ蜿門ｾ暦ｼ壽眠譌ｧ繝倥ャ繝蜷ｸ蜿取ｸ・ ===
Public Function GetSavedADLIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    GetSavedADLIO = ReadStr("IO_ADL", r, ws)
End Function



'=== Latest row resolver (謖・ｮ壹・繝・ム縺ｮ譛邨り｡後ｒ霑斐☆) ===
Public Function LatestRowByHeader(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim c As Long: c = HeaderCol(wantName, ws)
    If c = 0 Then LatestRowByHeader = 0: Exit Function
    LatestRowByHeader = ws.Cells(ws.rows.count, c).End(xlUp).row
End Function



'=== ADL loader (Raw)・壽怙譁ｰ陦後・IO_ADL譁・ｭ怜・繧定ｿ斐☆ ===
Public Function LoadLatestADLNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_ADL", ws)
    If r <= 0 Then Exit Function
    LoadLatestADLNow_Raw = GetSavedADLIO(r, ws)
End Function


'=== ADL loader (Get by Key)・壽怙譁ｰIO_ADL縺九ｉ key 縺ｮ蛟､繧定ｿ斐☆ ===
Public Function ADL_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestADLNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long, klen As Long
    parts = Split(raw, "|")
    klen = Len(key) + 1 ' "key=" 縺ｮ髟ｷ縺・
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), klen) = key & "=" Then
            ADL_Get = Mid$(parts(i), klen + 1) ' "="縺ｮ谺｡縺ｮ譁・ｭ励°繧画忰蟆ｾ
            Exit Function
        End If
    Next i
End Function


'=== Sensory loader (Raw)・壽怙譁ｰ陦後・SENSE_IO譁・ｭ怜・繧定ｿ斐☆ ===
Public Function LoadLatestSensoryNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_Sensory", ws)
    If r <= 0 Then Exit Function
    LoadLatestSensoryNow_Raw = GetSavedSensoryIO(r, ws)
End Function



'=== Sensory loader (Get by Key)・售ENSE_IO縺九ｉ key 縺ｫ荳閾ｴ縺吶ｋ驛ｨ蛻・ｒ霑斐☆ ===
Public Function Sensory_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestSensoryNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long
    parts = Split(raw, "|")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key)) = key Then
            Sensory_Get = parts(i)
            Exit Function
        End If
    Next i
End Function



'=== ROM loader (Raw)・壽怙譁ｰ陦後・ROM_*鄒､繧偵∪縺ｨ繧√※霑斐☆ ===
Public Function LoadLatestROMNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("ROM_Upper_Shoulder_Flex_R", ws)
    If r <= 0 Then Exit Function
    LoadLatestROMNow_Raw = GetSavedROMIO(r, ws)
End Function



'=== ROM loader (Get by Key)・啌OM_*鄒､縺九ｉ key 縺ｮ蛟､繧定ｿ斐☆ ===
Public Function ROM_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestROMNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long, eqPos As Long
    parts = Split(raw, "|")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key) + 1) = key & "=" Then
            eqPos = InStr(parts(i), "=")
            If eqPos > 0 Then ROM_Get = Mid$(parts(i), eqPos + 1)
            Exit Function
        End If
    Next i
End Function



'=== MMT loader (Raw)・壽怙譁ｰ陦後・MMT_IO繧定ｿ斐☆ ===
Public Function LoadLatestMMTNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_MMT", ws)
    If r <= 0 Then Exit Function
    LoadLatestMMTNow_Raw = ReadStr("IO_MMT", r, ws)
End Function

'=== MMT loader (Get by Key)・唔O_MMT譁・ｭ怜・縺九ｉ key 縺ｮ蛟､繧定ｿ斐☆ ===
Public Function MMT_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestMMTNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long, eqPos As Long
    parts = Split(raw, "|")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key) + 1) = key & "=" Then
            eqPos = InStr(parts(i), "=")
            If eqPos > 0 Then MMT_Get = Mid$(parts(i), eqPos + 1)
            Exit Function
        End If
    Next i
End Function

'=== Tone loader (Raw)・壽怙譁ｰ陦後・TONE_IO繧定ｿ斐☆ ===
Public Function LoadLatestToneNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_Tone", ws)
    If r <= 0 Then Exit Function
    LoadLatestToneNow_Raw = ReadStr("IO_Tone", r, ws)
End Function

'=== Tone loader (Get by Key)・唔O_Tone譁・ｭ怜・縺九ｉ key 縺ｮ蛟､繧定ｿ斐☆ ===
Public Function Tone_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestToneNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long, eqPos As Long
    parts = Split(raw, "|")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key) + 1) = key & "=" Then
            eqPos = InStr(parts(i), "=")
            If eqPos > 0 Then Tone_Get = Mid$(parts(i), eqPos + 1)
            Exit Function
        End If
    Next i
End Function






'=== Recent rows by ID・壽欠螳唔D縺ｮ逶ｴ霑鮮莉ｶ縺ｮ陦檎分蜿ｷ繧呈眠竊呈立縺ｮ鬆・〒霑斐☆ ===
Public Function RecentRowsByID(ByVal targetID As Variant, Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet) As Variant
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim cID As Long: cID = HeaderCol("ID", ws)
    If cID = 0 Or n <= 0 Then RecentRowsByID = Array(): Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, cID).End(xlUp).row
    Dim rowsOut() As Long, r As Long, hit As Long
    ReDim rowsOut(0 To 0): hit = 0
    For r = lastRow To 2 Step -1
        If ws.Cells(r, cID).value = targetID Then
            If hit > 0 Then ReDim Preserve rowsOut(0 To hit)
            rowsOut(hit) = r
            hit = hit + 1
            If hit >= n Then Exit For
        End If
    Next
    If hit = 0 Then RecentRowsByID = Array() Else RecentRowsByID = rowsOut
End Function






'=== Preview: 謖・ｮ唔D縺ｮ逶ｴ霑鮮莉ｶ繧貞ｮ牙・縺ｫ荳隕ｧ蜃ｺ蜉幢ｼ・I蜑阪・譛邨ら｢ｺ隱搾ｼ・===
Public Sub Preview_RecentEvalRows(ByVal targetID As Variant, ByVal n As Long, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim arr As Variant, i As Long, r As Long
    arr = RecentRowsByID(targetID, n, ws)

    ' 遨ｺ驟榊・縺ｫ螳牙・蟇ｾ蠢・
    On Error Resume Next
    i = UBound(arr)
    If Err.Number <> 0 Then
        Debug.Print "=== [Recent] ID=" & targetID & " | none ==="
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "=== [Recent] ID=" & targetID & " ==="
    For i = LBound(arr) To UBound(arr)
        r = arr(i)
        Debug.Print "r=" & r & _
            " | ROM=" & Len(GetSavedROMIO(r, ws)) & _
            " | SENSE=" & Len(GetSavedSensoryIO(r, ws)) & _
            " | MMT=" & Len(ReadStr("IO_MMT", r, ws)) & _
            " | TONE=" & Len(ReadStr("IO_Tone", r, ws)) & _
            " | ADL=" & Len(GetSavedADLIO(r, ws)) & _
            " | PAIN=" & Len(ReadStr("IO_Pain", r, ws))
    Next i
    Debug.Print "=== /Recent ==="
End Sub



'=== Unified dispatcher: 驕ｸ謚櫁｡後・隧穂ｾ｡繝・・繧ｿ繧剃ｸ諡ｬ隱ｭ霎ｼ ===
Public Sub LoadSelectedEvalRow(ByVal r As Long, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

End Sub



'=== Latest row by ID・壽欠螳唔D縺ｮ譛譁ｰ陦・1莉ｶ)繧定ｿ斐☆縲ら┌縺代ｌ縺ｰ0 ===
Public Function LatestRowByID(ByVal targetID As Variant, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim arr As Variant, ub As Long
    arr = RecentRowsByID(targetID, 1, ws)  ' 譁ｰ竊呈立縺ｮ鬆・〒譛螟ｧ1莉ｶ
    On Error Resume Next
    ub = UBound(arr)
    If Err.Number <> 0 Then
        LatestRowByID = 0
        Err.Clear
    Else
        LatestRowByID = arr(0)
    End If
    On Error GoTo 0
End Function


'=== Select & Load Recent by ID・・I蜑阪・螳牙・繧ｻ繝ｬ繧ｯ繧ｿ・・==
Public Sub SelectAndLoadRecentByID(ByVal targetID As Variant, Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim rowsArr As Variant, i As Long, ub As Long, pick As Variant, idx As Long
    rowsArr = RecentRowsByID(targetID, n, ws)

    ' 遨ｺ驟榊・蟇ｾ蠢・
    On Error Resume Next
    ub = UBound(rowsArr)
    If Err.Number <> 0 Then
        Debug.Print "=== [Recent] ID=" & targetID & " | none ==="
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' 逡ｪ蜿ｷ莉倥″縺ｧ荳隕ｧ陦ｨ遉ｺ・・mmediate・・
    Debug.Print "=== [Recent5] ID=" & targetID & " ==="
    For i = LBound(rowsArr) To UBound(rowsArr)
        Debug.Print (i - LBound(rowsArr) + 1) & ":" & rowsArr(i) & _
            " | ROM=" & Len(GetSavedROMIO(rowsArr(i), ws)) & _
            " | SENSE=" & Len(GetSavedSensoryIO(rowsArr(i), ws)) & _
            " | MMT=" & Len(ReadStr("IO_MMT", rowsArr(i), ws)) & _
            " | TONE=" & Len(ReadStr("IO_Tone", rowsArr(i), ws)) & _
            " | ADL=" & Len(GetSavedADLIO(rowsArr(i), ws)) & _
            " | PAIN=" & Len(ReadStr("IO_Pain", rowsArr(i), ws))
    Next i
    Debug.Print "=== /Recent5 ==="

    ' 逡ｪ蜿ｷ蜈･蜉幢ｼ・?莉ｶ謨ｰ・・
    'pick = Application.InputBox(Prompt:="隱ｭ縺ｿ霎ｼ繧逡ｪ蜿ｷ繧貞・蜉幢ｼ・・・ & (UBound(rowsArr) - LBound(rowsArr) + 1) & "・・, Type:=1)
    If VarType(pick) = vbBoolean And pick = False Then Exit Sub 'Cancel
    idx = CLng(pick)

    ' 遽・峇繝√ぉ繝・け
    If idx < 1 Or idx > (UBound(rowsArr) - LBound(rowsArr) + 1) Then
        Debug.Print "[SelectRecent] 遽・峇螟・ " & idx
        Exit Sub
    End If

    ' 螳溯｡・
    LoadSelectedEvalRow rowsArr(UBound(rowsArr)), ws

End Sub



'=== UI hook・壹い繧ｯ繝・ぅ繝冶｡後・ID縺ｧ逶ｴ霑鮮莉ｶ繧ｻ繝ｬ繧ｯ繝遺・荳諡ｬ隱ｭ霎ｼ ===
Public Sub Run_LoadRecentForActiveID(Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim cID As Long: cID = HeaderCol("ID", ws)
    If cID = 0 Then Debug.Print "[Run_LoadRecent] ID header not found": Exit Sub
    If ActiveCell.row < 2 Then Debug.Print "[Run_LoadRecent] ActiveCell row invalid": Exit Sub
    Dim targetID As Variant: targetID = ws.Cells(ActiveCell.row, cID).value
    If Len(CStr(targetID)) = 0 Then Debug.Print "[Run_LoadRecent] ID empty at row " & ActiveCell.row: Exit Sub

    SelectAndLoadRecentByID targetID, n, ws
End Sub


' LEGACY: 譛ｬ逡ｪ遖∵ｭ｢
'=== Pick only: 逶ｴ霑鮮莉ｶ縺九ｉ逡ｪ蜿ｷ驕ｸ謚樞・陦檎分蜿ｷ繧定ｿ斐☆・・I縺ｯ隗ｦ繧峨↑縺・ｼ・===
Public Function PickRecentRowByID(ByVal targetID As Variant, Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim rowsArr As Variant, i As Long, ub As Long, pick As Variant, idx As Long
    
    rowsArr = RecentRowsByID(targetID, n, ws)
    On Error Resume Next
    ub = UBound(rowsArr)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        PickRecentRowByID = 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' 荳隕ｧ・・mmediate・・
    Debug.Print "=== [Recent5] ID=" & targetID & " ==="
    For i = LBound(rowsArr) To UBound(rowsArr)
        Debug.Print (i - LBound(rowsArr) + 1) & ":" & rowsArr(i)
    Next i
    Debug.Print "=== /Recent5 ==="
    
    'pick = Application.InputBox(Prompt:="隱ｭ縺ｿ霎ｼ繧逡ｪ蜿ｷ繧貞・蜉幢ｼ・・・ & (UBound(rowsArr) - LBound(rowsArr) + 1) & "・・, Type:=1)
    If VarType(pick) = vbBoolean And pick = False Then Exit Function 'Cancel
    idx = CLng(pick)
    If idx < 1 Or idx > (UBound(rowsArr) - LBound(rowsArr) + 1) Then Exit Function
    
    PickRecentRowByID = rowsArr(LBound(rowsArr) + idx - 1)
    
    PickRecentRowByID = rowsArr(UBound(rowsArr))

    
End Function




'=== ResolveLegacyCol・啌esolveLegacyHeader繧剃ｽｿ縺｣縺ｦ蛻礼分蜿ｷ繧定ｿ斐☆ ===
Public Function ResolveLegacyCol(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim colName As String
    colName = ResolveLegacyHeader(wantName)
    ResolveLegacyCol = HeaderCol(colName, ws)
End Function

Public Function CompatHeaderNames(ByVal wantName As String) As Variant
    Dim mapped As String
    mapped = ResolveLegacyHeader(wantName)

    Select Case LCase$(mapped)
        Case "io_cog_judgment", "io_cog_judgement"
            CompatHeaderNames = Array("IO_Cog_Judgment", "IO_Cog_Judgement")
        Case Else
            CompatHeaderNames = Array(mapped)
    End Select
End Function


'=== HeaderCol_Compat・壽眠譌ｧ繝倥ャ繝蜷阪ｒ蜷ｸ蜿弱＠縺ｦ蛻礼分蜿ｷ繧定ｿ斐☆ ===
Public Function HeaderCol_Compat(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim headers As Variant
    Dim i As Long

    headers = CompatHeaderNames(wantName)
    For i = LBound(headers) To UBound(headers)
        HeaderCol_Compat = HeaderCol(CStr(headers(i)), ws)
        If HeaderCol_Compat > 0 Then Exit Function
    Next i
End Function


'=== ReadStr_Compat・壽眠譌ｧ繝倥ャ繝蜷阪ｒ蜷ｸ蜿弱＠縺ｦ r 陦後・蛟､繧定ｿ斐☆・・OM縺ｯ迚ｹ蛻･謇ｱ縺・ｼ・===
Public Function ReadStr_Compat(ByVal wantName As String, ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet

    ' 陦檎分蜿ｷ縺御ｸ肴ｭ｣縺ｪ繧牙叉遨ｺ霑斐＠・・ickRecentRowByID 縺・0 縺ｮ蝣ｴ蜷医↑縺ｩ・・
    If r <= 0 Then
        ReadStr_Compat = vbNullString
        Exit Function
    End If

    ' ROM 縺ｯ隍・焚蛻暦ｼ・OM_*・峨↑縺ｮ縺ｧ迚ｹ蛻･繝ｫ繝ｼ繝医〒騾｣邨仙叙蠕・
    If StrComp(wantName, "IO_ROM", vbTextCompare) = 0 Then
        ReadStr_Compat = GetSavedROMIO(r, ws)
        Exit Function
    End If

    ' 縺昴ｌ莉･螟悶・繝倥ャ繝莠呈鋤縺ｧ蜊倅ｸ繧ｻ繝ｫ隱ｭ蜿・
    Dim c As Long: c = HeaderCol_Compat(wantName, ws)
    If c > 0 Then
        ReadStr_Compat = CStr(ws.Cells(r, c).value)
    Else
        ReadStr_Compat = vbNullString
    End If
End Function



