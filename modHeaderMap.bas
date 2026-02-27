Attribute VB_Name = "modHeaderMap"

'=== Legacy/New Header Resolver (読み込み互換マップ) ===
Public Function ResolveLegacyHeader(ByVal wantName As String) As String
    ' 読み込み側が探す新名 → 現行保存の旧名 へマップ
    ' 見つからなければそのまま返す（同名保存に対応）
    Select Case LCase$(wantName)
        Case "io_sensory":   ResolveLegacyHeader = "IO_Sensory"
        Case "io_testeval": ResolveLegacyHeader = "IO_TestEval"
        Case "io_mmt":       ResolveLegacyHeader = "MMT_IO"
        Case "io_rom":       ResolveLegacyHeader = "ROM_*"      '※複数列。後続でワイルドカード展開
        Case "io_adl":       ResolveLegacyHeader = "IO_ADL"
        Case "io_tone":      ResolveLegacyHeader = "TONE_IO"
        Case Else:           ResolveLegacyHeader = wantName
    End Select
End Function




'=== Header Column Resolver (新旧ヘッダ名を吸収して列番号を返す) ===
Public Function HeaderCol(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim h As String: h = ResolveLegacyHeader(wantName)
    On Error Resume Next
    HeaderCol = Application.Match(h, ws.rows(1), 0)
    On Error GoTo 0
End Function



'=== Read String by Header (新旧ヘッダ対応でセル文字列取得) ===
Public Function ReadStr(ByVal wantName As String, ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim c As Long: c = HeaderCol(wantName, ws)
    If c > 0 Then ReadStr = CStr(ws.Cells(r, c).value) Else ReadStr = vbNullString
End Function



'=== Sensory IO accessor (保存値の取得：新旧ヘッダ吸収済) ===
Public Function GetSavedSensoryIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    GetSavedSensoryIO = ReadStr("IO_Sensory", r, ws)
End Function




'=== ROM IO accessor (保存値の取得：見出し "ROM_*" を横断) ===
Public Function GetSavedROMIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim buf As String
    Dim M As Variant
    Dim c As Long
    Dim h As String, v As String

    ' 1) IO_ROM 列があれば、それを最優先で使う
    M = Application.Match("IO_ROM", ws.rows(1), 0)
    If Not IsError(M) Then
        buf = CStr(ws.Cells(r, CLng(M)).value)
        If Len(buf) > 0 Then
            GetSavedROMIO = buf
            Exit Function
        End If
    End If

    ' 2) レガシー互換：ROM_* 列から組み立て（範囲を 160?213 に限定）
    '    ※ 214以降の重複ROM列は無視して暴走連結を防ぐ
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

    GetSavedROMIO = buf   ' 対象が無ければ空文字
End Function




'=== ADL IO accessor (保存値の取得：新旧ヘッダ吸収済) ===
Public Function GetSavedADLIO(ByVal r As Long, Optional ByVal ws As Worksheet) As String
    GetSavedADLIO = ReadStr("IO_ADL", r, ws)
End Function



'=== Latest row resolver (指定ヘッダの最終行を返す) ===
Public Function LatestRowByHeader(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim c As Long: c = HeaderCol(wantName, ws)
    If c = 0 Then LatestRowByHeader = 0: Exit Function
    LatestRowByHeader = ws.Cells(ws.rows.Count, c).End(xlUp).row
End Function



'=== ADL loader (Raw)：最新行のIO_ADL文字列を返す ===
Public Function LoadLatestADLNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_ADL", ws)
    If r <= 0 Then Exit Function
    LoadLatestADLNow_Raw = GetSavedADLIO(r, ws)
End Function


'=== ADL loader (Get by Key)：最新IO_ADLから key の値を返す ===
Public Function ADL_Get(ByVal key As String, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim raw As String: raw = LoadLatestADLNow_Raw(ws)
    If Len(raw) = 0 Then Exit Function
    Dim parts() As String, i As Long, klen As Long
    parts = Split(raw, "|")
    klen = Len(key) + 1 ' "key=" の長さ
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), klen) = key & "=" Then
            ADL_Get = Mid$(parts(i), klen + 1) ' "="の次の文字から末尾
            Exit Function
        End If
    Next i
End Function


'=== Sensory loader (Raw)：最新行のSENSE_IO文字列を返す ===
Public Function LoadLatestSensoryNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_Sensory", ws)
    If r <= 0 Then Exit Function
    LoadLatestSensoryNow_Raw = GetSavedSensoryIO(r, ws)
End Function



'=== Sensory loader (Get by Key)：SENSE_IOから key に一致する部分を返す ===
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



'=== ROM loader (Raw)：最新行のROM_*群をまとめて返す ===
Public Function LoadLatestROMNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("ROM_Upper_Shoulder_Flex_R", ws)
    If r <= 0 Then Exit Function
    LoadLatestROMNow_Raw = GetSavedROMIO(r, ws)
End Function



'=== ROM loader (Get by Key)：ROM_*群から key の値を返す ===
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



'=== MMT loader (Raw)：最新行のMMT_IOを返す ===
Public Function LoadLatestMMTNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_MMT", ws)
    If r <= 0 Then Exit Function
    LoadLatestMMTNow_Raw = ReadStr("IO_MMT", r, ws)
End Function

'=== MMT loader (Get by Key)：IO_MMT文字列から key の値を返す ===
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

'=== Tone loader (Raw)：最新行のTONE_IOを返す ===
Public Function LoadLatestToneNow_Raw(Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long: r = LatestRowByHeader("IO_Tone", ws)
    If r <= 0 Then Exit Function
    LoadLatestToneNow_Raw = ReadStr("IO_Tone", r, ws)
End Function

'=== Tone loader (Get by Key)：IO_Tone文字列から key の値を返す ===
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






'=== Recent rows by ID：指定IDの直近N件の行番号を新→旧の順で返す ===
Public Function RecentRowsByID(ByVal targetID As Variant, Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet) As Variant
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim cID As Long: cID = HeaderCol("ID", ws)
    If cID = 0 Or n <= 0 Then RecentRowsByID = Array(): Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, cID).End(xlUp).row
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






'=== Preview: 指定IDの直近N件を安全に一覧出力（UI前の最終確認） ===
Public Sub Preview_RecentEvalRows(ByVal targetID As Variant, ByVal n As Long, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim arr As Variant, i As Long, r As Long
    arr = RecentRowsByID(targetID, n, ws)

    ' 空配列に安全対応
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



'=== Unified dispatcher: 選択行の評価データを一括読込 ===
Public Sub LoadSelectedEvalRow(ByVal r As Long, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

End Sub



'=== Latest row by ID：指定IDの最新行(1件)を返す。無ければ0 ===
Public Function LatestRowByID(ByVal targetID As Variant, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim arr As Variant, ub As Long
    arr = RecentRowsByID(targetID, 1, ws)  ' 新→旧の順で最大1件
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


'=== Select & Load Recent by ID（UI前の安全セレクタ）===
Public Sub SelectAndLoadRecentByID(ByVal targetID As Variant, Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim rowsArr As Variant, i As Long, ub As Long, pick As Variant, idx As Long
    rowsArr = RecentRowsByID(targetID, n, ws)

    ' 空配列対応
    On Error Resume Next
    ub = UBound(rowsArr)
    If Err.Number <> 0 Then
        Debug.Print "=== [Recent] ID=" & targetID & " | none ==="
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' 番号付きで一覧表示（Immediate）
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

    ' 番号入力（1?件数）
    'pick = Application.InputBox(Prompt:="読み込む番号を入力（1～" & (UBound(rowsArr) - LBound(rowsArr) + 1) & "）", Type:=1)
    If VarType(pick) = vbBoolean And pick = False Then Exit Sub 'Cancel
    idx = CLng(pick)

    ' 範囲チェック
    If idx < 1 Or idx > (UBound(rowsArr) - LBound(rowsArr) + 1) Then
        Debug.Print "[SelectRecent] 範囲外: " & idx
        Exit Sub
    End If

    ' 実行
    LoadSelectedEvalRow rowsArr(UBound(rowsArr)), ws

End Sub



'=== UI hook：アクティブ行のIDで直近N件セレクト→一括読込 ===
Public Sub Run_LoadRecentForActiveID(Optional ByVal n As Long = 5, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim cID As Long: cID = HeaderCol("ID", ws)
    If cID = 0 Then Debug.Print "[Run_LoadRecent] ID header not found": Exit Sub
    If ActiveCell.row < 2 Then Debug.Print "[Run_LoadRecent] ActiveCell row invalid": Exit Sub
    Dim targetID As Variant: targetID = ws.Cells(ActiveCell.row, cID).value
    If Len(CStr(targetID)) = 0 Then Debug.Print "[Run_LoadRecent] ID empty at row " & ActiveCell.row: Exit Sub

    SelectAndLoadRecentByID targetID, n, ws
End Sub


' LEGACY: 本番禁止
'=== Pick only: 直近N件から番号選択→行番号を返す（UIは触らない） ===
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
    
    ' 一覧（Immediate）
    Debug.Print "=== [Recent5] ID=" & targetID & " ==="
    For i = LBound(rowsArr) To UBound(rowsArr)
        Debug.Print (i - LBound(rowsArr) + 1) & ":" & rowsArr(i)
    Next i
    Debug.Print "=== /Recent5 ==="
    
    'pick = Application.InputBox(Prompt:="読み込む番号を入力（1～" & (UBound(rowsArr) - LBound(rowsArr) + 1) & "）", Type:=1)
    If VarType(pick) = vbBoolean And pick = False Then Exit Function 'Cancel
    idx = CLng(pick)
    If idx < 1 Or idx > (UBound(rowsArr) - LBound(rowsArr) + 1) Then Exit Function
    
    PickRecentRowByID = rowsArr(LBound(rowsArr) + idx - 1)
    
    PickRecentRowByID = rowsArr(UBound(rowsArr))

    
End Function




'=== ResolveLegacyCol：ResolveLegacyHeaderを使って列番号を返す ===
Public Function ResolveLegacyCol(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim colName As String
    colName = ResolveLegacyHeader(wantName)
    ResolveLegacyCol = HeaderCol(colName, ws)
End Function


'=== HeaderCol_Compat：新旧ヘッダ名を吸収して列番号を返す ===
Public Function HeaderCol_Compat(ByVal wantName As String, Optional ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim mapped As String
    mapped = ResolveLegacyHeader(wantName) ' 例: "IO_Sensory"→"SENSE_IO"
    HeaderCol_Compat = HeaderCol(mapped, ws)
End Function


'=== ReadStr_Compat：新旧ヘッダ名を吸収して r 行の値を返す（ROMは特別扱い） ===
Public Function ReadStr_Compat(ByVal wantName As String, ByVal r As Long, Optional ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = ActiveSheet

    ' 行番号が不正なら即空返し（PickRecentRowByID が 0 の場合など）
    If r <= 0 Then
        ReadStr_Compat = vbNullString
        Exit Function
    End If

    ' ROM は複数列（ROM_*）なので特別ルートで連結取得
    If StrComp(wantName, "IO_ROM", vbTextCompare) = 0 Then
        ReadStr_Compat = GetSavedROMIO(r, ws)
        Exit Function
    End If

    ' それ以外はヘッダ互換で単一セル読取
    Dim c As Long: c = HeaderCol_Compat(wantName, ws)
    If c > 0 Then
        ReadStr_Compat = CStr(ws.Cells(r, c).value)
    Else
        ReadStr_Compat = vbNullString
    End If
End Function



