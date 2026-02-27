Attribute VB_Name = "modADL_Probe"


'=== Util: 指定Captionの右隣コンボを返す（同一行±6ptで最も近い） ===
Private Function GetRightComboByLabelCaptionIn(p As MSForms.Page, ByVal cap As String) As MSForms.ComboBox
    Dim i As Long, lb As MSForms.label, c As Control, best As MSForms.ComboBox
    Dim dy As Double, bestDx As Double: bestDx = 1E+30
    ' 1) Caption一致ラベルを探す
    For i = 0 To p.Controls.Count - 1
        If TypeName(p.Controls(i)) = "Label" Then
            Set lb = p.Controls(i)
            If lb.caption = cap Then
                ' 2) 同じ行(±6pt)で右側にある最短距離のComboBox
                For Each c In p.Controls
                    If TypeName(c) = "ComboBox" Then
                        dy = Abs(c.Top - lb.Top)
                        If dy <= 6 And c.Left > lb.Left Then
                            If (c.Left - lb.Left) < bestDx Then
                                Set best = c
                                bestDx = (c.Left - lb.Left)
                            End If
                        End If
                    End If
                Next c
                Exit For
            End If
        End If
    Next i
    If Not best Is Nothing Then Set GetRightComboByLabelCaptionIn = best
End Function





'=== Resolve: 起居動作の無名Combo（立ち上がり／立位保持）を取得 ===
Private Sub ResolveKyoUnnamedCombos(ByRef cmbStandUp As MSForms.ComboBox, ByRef cmbStandHold As MSForms.ComboBox)
    Dim mp As MSForms.MultiPage, p As MSForms.Page, c As Control
    ' mpADL 取得
    For Each c In frmEval.Controls
        If TypeName(c) = "MultiPage" Then
            If c.name = "mpADL" Then Set mp = c: Exit For
        End If
    Next c
    If mp Is Nothing Then Exit Sub
    Set p = mp.Pages(2) ' 起居動作
    Set cmbStandUp = GetRightComboByLabelCaptionIn(p, "立ち上がり")
    Set cmbStandHold = GetRightComboByLabelCaptionIn(p, "立位保持")
End Sub

'=== Snapshot: ADL（BI/IADL/起居動作）を固定順でシリアライズ表示 ===
Public Sub Snapshot_ADL_Once()
    Dim mp As MSForms.MultiPage, p As MSForms.Page, ctl As Control
    Dim i As Long, s As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    ' mpADL 取得
    For Each ctl In frmEval.Controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Debug.Print "[ERR] mpADL not found": Exit Sub

    ' --- BI (#0) ---
    Set p = mp.Pages(0)
    s = ""
    v = p.Controls("txtBITotal").Text: s = s & "BITotal=" & v & "|"
    For i = 0 To 9
        v = p.Controls("cmbBI_" & i).Text
        s = s & "BI_" & i & "=" & v & "|"
    Next i

    ' --- IADL (#1) ---
    Set p = mp.Pages(1)
    For i = 0 To 8
        v = p.Controls("cmbIADL_" & i).Text
        s = s & "IADL_" & i & "=" & v & "|"
    Next i
    v = p.Controls("txtIADLNote").Text
    s = s & "IADLNote=" & v & "|"

    ' --- 起居動作 (#2) ---
    Set p = mp.Pages(2)
    s = s & "Kyo_Roll=" & p.Controls("cmbKyo_Roll").Text & "|"
    s = s & "Kyo_SitUp=" & p.Controls("cmbKyo_SitUp").Text & "|"
    s = s & "Kyo_SitHold=" & p.Controls("cmbKyo_SitHold").Text & "|"

    Call ResolveKyoUnnamedCombos(cmbSU, cmbSH)
    If Not cmbSU Is Nothing Then s = s & "Kyo_StandUp=" & cmbSU.Text & "|" Else Debug.Print "[WARN] 立ち上がり 未解決"
    If Not cmbSH Is Nothing Then s = s & "Kyo_StandHold=" & cmbSH.Text & "|" Else Debug.Print "[WARN] 立位保持 未解決"

    s = s & "Kyo_Note=" & p.Controls("txtKyoNote").Text

    Debug.Print "[ADL.IO] "; s
    Debug.Print "[ADL.IO.Len] "; Len(s)
End Sub




'=== ADL IO Builder: フォーム上のADL値を固定順で連結して返す（副作用なし） ===
Public Function Build_ADL_IO() As String
    Dim mp As MSForms.MultiPage, p As MSForms.Page, ctl As Control
    Dim i As Long, s As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    ' mpADL 取得
    For Each ctl In frmEval.Controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Function

    ' --- BI (#0) ---
    Set p = mp.Pages(0)
    s = ""
    v = p.Controls("txtBITotal").Text: s = s & "BITotal=" & v & "|"
    For i = 0 To 9
        v = p.Controls("cmbBI_" & i).Text
        s = s & "BI_" & i & "=" & v & "|"
    Next i

    ' --- IADL (#1) ---
    Set p = mp.Pages(1)
    For i = 0 To 8
        v = p.Controls("cmbIADL_" & i).Text
        s = s & "IADL_" & i & "=" & v & "|"
    Next i
    v = p.Controls("txtIADLNote").Text
    s = s & "IADLNote=" & v & "|"

    ' --- 起居動作 (#2) ---
    Set p = mp.Pages(2)
    s = s & "Kyo_Roll=" & p.Controls("cmbKyo_Roll").Text & "|"
    s = s & "Kyo_SitUp=" & p.Controls("cmbKyo_SitUp").Text & "|"
    s = s & "Kyo_SitHold=" & p.Controls("cmbKyo_SitHold").Text & "|"

    ' 無名コンボ解決（立ち上がり／立位保持）
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "立ち上がり")
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "立位保持")
    If Not cmbSU Is Nothing Then s = s & "Kyo_StandUp=" & cmbSU.Text & "|"
    If Not cmbSH Is Nothing Then s = s & "Kyo_StandHold=" & cmbSH.Text & "|"

    s = s & "Kyo_Note=" & p.Controls("txtKyoNote").Text

    Build_ADL_IO = s
End Function




'=== Save: ADL（BI/IADL/起居動作）を EvalData に1行追記（IO_ADL列） ===

Public Sub Save_ADL_Once()
    Dim ws As Worksheet, look As Object
    Dim s As String, r As Long, c As Long
    Dim lastCol As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")            ' 既存ヘルパ（PainIOと同じ想定）
    

    c = EnsureHeader(ws, "IO_ADL")


    ' 追記行を決定（ヘッダの次行から開始）
    r = ws.Cells(ws.rows.Count, c).End(xlUp).row: If r < 2 Then r = 2 Else r = r + 1


    ' IO生成 → 書き込み
    s = Build_ADL_IO()
    Debug.Print "[Chk]"; TypeName(ws); r; c; TypeName(ws.Cells(r, c))

ws.Cells(r, c).Value2 = CStr(s)




End Sub

'=== Helper: 見出し列を保証して列番号を返す（無ければ1行目の末尾に作成） ===
Public Function EnsureHeader(ws As Worksheet, ByVal header As String) As Long

    Dim M As Variant, lastCol As Long
    M = Application.Match(header, ws.rows(1), 0)
    If IsError(M) Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = header
        EnsureHeader = lastCol + 1
    Else
        EnsureHeader = CLng(M)
    End If
End Function







'=== Load: EvalDataの IO_ADL 最新行を読み込み、フォームに反映 ===
Public Sub Load_ADL_Latest()
    Dim ws As Worksheet, mp As MSForms.MultiPage, p As MSForms.Page, ctl As Control
    Dim c As Long, r As Long, s As String
    Dim parts As Variant, i As Long, n As Long
    Dim k As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.Count, c).End(xlUp).row
    If r < 2 Then Exit Sub    ' データなし

    s = ReadStr_Compat("IO_ADL", r, ws)
    parts = Split(s, "|")

    ' mpADL 取得
    For Each ctl In frmEval.Controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Sub

    ' 無名コンボ（起居：立ち上がり／立位保持）を解決
    Set p = mp.Pages(2) ' 起居動作
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "立ち上がり")
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "立位保持")

    ' ペアを順次反映
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextI
        If InStr(1, parts(i), "=") = 0 Then GoTo NextI
        k = Left$(parts(i), InStr(1, parts(i), "=") - 1)
        v = Mid$(parts(i), InStr(1, parts(i), "=") + 1)

        Select Case k
            
    ' --- BI (#0) ---
    Case "BITotal":                 mp.Pages(0).Controls("txtBITotal").Text = v
    Case "BI_0":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_0"), v
    Case "BI_1":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_1"), v
    Case "BI_2":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_2"), v
    Case "BI_3":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_3"), v
    Case "BI_4":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_4"), v
    Case "BI_5":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_5"), v
    Case "BI_6":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_6"), v
    Case "BI_7":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_7"), v
    Case "BI_8":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_8"), v
    Case "BI_9":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_9"), v

    ' --- IADL (#1) ---
    Case "IADL_0":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_0"), v
    Case "IADL_1":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_1"), v
    Case "IADL_2":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_2"), v
    Case "IADL_3":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_3"), v
    Case "IADL_4":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_4"), v
    Case "IADL_5":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_5"), v
    Case "IADL_6":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_6"), v
    Case "IADL_7":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_7"), v
    Case "IADL_8":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_8"), v
    Case "IADLNote":                mp.Pages(1).Controls("txtIADLNote").Text = v

    ' --- 起居動作 (#2) ---
    Case "Kyo_Roll":                SafeSetComboValue mp.Pages(2).Controls("cmbKyo_Roll"), v
    Case "Kyo_SitUp":               SafeSetComboValue mp.Pages(2).Controls("cmbKyo_SitUp"), v
    Case "Kyo_SitHold":             SafeSetComboValue mp.Pages(2).Controls("cmbKyo_SitHold"), v
    Case "Kyo_StandUp":             If Not cmbSU Is Nothing Then SafeSetComboValue cmbSU, v
    Case "Kyo_StandHold":           If Not cmbSH Is Nothing Then SafeSetComboValue cmbSH, v
    Case "Kyo_Note":                mp.Pages(2).Controls("txtKyoNote").Text = v
End Select

        n = n + 1
NextI:
    Next i

    Debug.Print "[ADL.Load] Row=" & r & " | Pairs=" & n & " | Len=" & Len(s)
End Sub



'=== Save→Load: ADL を一発検証（EvalDataに追記→直後にフォームへ反映） ===
Public Sub SaveAndReload_ADL()
    Dim ws As Worksheet, c As Long, r As Long, s As String
    Call Save_ADL_Once
    Call Load_ADL_Latest

    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.Count, c).End(xlUp).row
    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[ADL.SaveLoad] Row=" & r & " Col=" & c & " | Len=" & Len(s)
End Sub





'=== Checklist: ADL 保存/読込の健全性を一発確認 ===
Public Sub PreRelease_ADL_Checklist()
    Dim ws As Worksheet, c As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.Count, c).End(xlUp).row
    If r < 2 Then Debug.Print "[ADL.Check] データなし": Exit Sub

    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[ADL.Check] Col=" & c & " Row=" & r & " | Len=" & Len(s)

    ' 冪等チェック：保存→読込→長さ
    Call SaveAndReload_ADL
    s = Build_ADL_IO
    Debug.Print "[ADL.Check] AfterReload Len=" & Len(s)
End Sub



'

Private Sub WalkCtrlPaths(host As Object, ByVal path As String)

    ' MultiPage は Controls ではなく Pages を走査
If TypeName(host) = "MultiPage" Then
    Dim pg As MSForms.Page
    For Each pg In host.Pages
        WalkCtrlPaths pg, path & "/" & pg.caption & ":Page"
    Next pg
    Exit Sub
End If


    Dim c As Control, t As String, p As String
    For Each c In host.Controls
        t = TypeName(c)
        p = path & "/" & c.name & ":" & t
        If c.name = "Frame33" Then Debug.Print "[HIT] "; p
        Select Case t
            Case "Frame", "MultiPage", "Page" '子を持つ可能性があるものだけ潜る
                WalkCtrlPaths c, p
        End Select
    Next c
End Sub




'=== Save: ADL を「指定行 r」に書き込む（行は外部で決定） ===
Public Sub Save_ADL_AtRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Long, s As String
    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    c = EnsureHeader(ws, "IO_ADL")   ' 見出し確保して列番号取得（同名が他にある場合は、その関数を使用しているモジュールのものでもOK）
    s = Build_ADL_IO                 ' 現在のフォーム値をIO化（固定順）

    ws.Cells(r, c).Value2 = CStr(s)  ' 指定行に上書き保存（追記は呼び出し側でrを進める）
    Debug.Print "[ADL.Save@Row] Row=" & r & " Col=" & c & " | Len=" & Len(s)
End Sub











Private Function ADLKeyNormalize(ByVal tag As String) As String
    ' UIタグ → 保存キー互換
    ' 例：BI.摂食→BI_0 / IADL.調理→IADL_0 / BI.Total→BITotal
    Dim M As Object, k As String
    k = Replace(tag, ".", "_")
    If k = "BI_Total" Then ADLKeyNormalize = "BITotal": Exit Function
    
    ' 日本語→番号の最小マップ（必要に応じて次の手で拡張）
    ' BI（バーサル）
    Set M = CreateObject("Scripting.Dictionary")
    M.CompareMode = 1
    M("BI_摂食") = "BI_0"
    M("BI_車いす-ベッド移乗") = "BI_1"
    M("BI_整容") = "BI_2"
    M("BI_トイレ動作") = "BI_3"
    M("BI_入浴") = "BI_4"
    M("BI_歩行/車いす移動") = "BI_5"
    M("BI_階段昇降") = "BI_6"
    M("BI_更衣") = "BI_7"
    M("BI_排便コントロール") = "BI_8"
    M("BI_排尿コントロール") = "BI_9"
    
    ' IADL
    M("IADL_調理") = "IADL_0"
    M("IADL_洗濯") = "IADL_1"
    M("IADL_掃除") = "IADL_2"
    M("IADL_買い物") = "IADL_3"
    M("IADL_金銭管理") = "IADL_4"
    M("IADL_服薬管理") = "IADL_5"
    M("IADL_趣味・余暇活動") = "IADL_6"
    M("IADL_社会参加（外出・地域活動）") = "IADL_7"
    M("IADL_コミュニケーション（電話・会話）") = "IADL_8"
    
    If M.exists(k) Then
        ADLKeyNormalize = M(k)
    Else
        ADLKeyNormalize = k ' 未定義はそのまま（次の手で補完）
    End If
End Function








Public Function FindADLControlByKey(ByVal key As String) As Control
    ' 例：key="BI_0" や "IADL_7" や "BITotal"
    Dim p As Object, pg As Object, ctl As Control, t As String, tag As String
    On Error Resume Next
    Set p = frmEval.Controls("mpADL")
    On Error GoTo 0
    If p Is Nothing Then Exit Function

    For Each pg In p.Pages
        For Each ctl In pg.Controls
            t = TypeName(ctl)
            If t = "TextBox" Or t = "ComboBox" Or t = "CheckBox" Then
                On Error Resume Next
                tag = ctl.tag
                On Error GoTo 0
                If Len(tag) > 0 Then
                    If ADLKeyNormalize(tag) = key Then
                        Set FindADLControlByKey = ctl
                        Exit Function
                    End If
                End If
            End If
        Next ctl
    Next pg
End Function








Private Function ComboHasItem(cmb As MSForms.ComboBox, val As String) As Boolean
    Dim i As Long
    If cmb Is Nothing Then Exit Function
    For i = 0 To cmb.ListCount - 1
        If CStr(cmb.List(i)) = CStr(val) Then ComboHasItem = True: Exit Function
    Next i
End Function




Private Sub SafeSetComboValue(cmb As MSForms.ComboBox, val As String)
    On Error Resume Next
    If ComboHasItem(cmb, val) Then
        cmb.value = val
    ElseIf cmb.MatchRequired = False Then
        cmb.AddItem val
        cmb.value = val
    End If
    On Error GoTo 0
End Sub




'=== Load: read IO_ADL from a specified row and apply to owner form ===
Public Sub Load_ADL_FromRow(ws As Worksheet, r As Long, owner As Object)
    Dim mp As MSForms.MultiPage, p As MSForms.Page, ctl As Control
    Dim c As Long, s As String
    Dim parts As Variant, i As Long, n As Long
    Dim k As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    If ws Is Nothing Then Exit Sub
    If owner Is Nothing Then Exit Sub
    If r < 2 Then Exit Sub

    c = EnsureHeader(ws, "IO_ADL")
    If c < 1 Then Exit Sub

    s = ReadStr_Compat("IO_ADL", r, ws)
    If Len(s) = 0 Then Exit Sub
    parts = Split(s, "|")

    ' mpADL 取得（ownerから）
    For Each ctl In owner.Controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Sub

    ' 無名コンボ（起居：立ち上がり／立位保持）を解決
    Set p = mp.Pages(2) ' 起居動作
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "立ち上がり")
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "立位保持")

    ' ペアを順次反映
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextI
        If InStr(1, parts(i), "=") = 0 Then GoTo NextI
        k = Left$(parts(i), InStr(1, parts(i), "=") - 1)
        v = Mid$(parts(i), InStr(1, parts(i), "=") + 1)

        Select Case k
            ' --- BI (#0) ---
            Case "BITotal":                 mp.Pages(0).Controls("txtBITotal").Text = v
            Case "BI_0":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_0"), v
            Case "BI_1":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_1"), v
            Case "BI_2":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_2"), v
            Case "BI_3":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_3"), v
            Case "BI_4":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_4"), v
            Case "BI_5":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_5"), v
            Case "BI_6":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_6"), v
            Case "BI_7":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_7"), v
            Case "BI_8":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_8"), v
            Case "BI_9":                    SafeSetComboValue mp.Pages(0).Controls("cmbBI_9"), v

            ' --- IADL (#1) ---
            Case "IADL_0":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_0"), v
            Case "IADL_1":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_1"), v
            Case "IADL_2":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_2"), v
            Case "IADL_3":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_3"), v
            Case "IADL_4":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_4"), v
            Case "IADL_5":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_5"), v
            Case "IADL_6":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_6"), v
            Case "IADL_7":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_7"), v
            Case "IADL_8":                  SafeSetComboValue mp.Pages(1).Controls("cmbIADL_8"), v
            Case "IADLNote":                mp.Pages(1).Controls("txtIADLNote").Text = v

            ' --- 起居動作 (#2) ---
            Case "Kyo_Roll":                SafeSetComboValue mp.Pages(2).Controls("cmbKyo_Roll"), v
            Case "Kyo_SitUp":               SafeSetComboValue mp.Pages(2).Controls("cmbKyo_SitUp"), v
            Case "Kyo_SitHold":             SafeSetComboValue mp.Pages(2).Controls("cmbKyo_SitHold"), v
            Case "Kyo_StandUp":             If Not cmbSU Is Nothing Then SafeSetComboValue cmbSU, v
            Case "Kyo_StandHold":           If Not cmbSH Is Nothing Then SafeSetComboValue cmbSH, v
            Case "Kyo_Note":                mp.Pages(2).Controls("txtKyoNote").Text = v
        End Select

        n = n + 1
NextI:
    Next i

    Debug.Print "[ADL.Load] Row=" & r & " | Pairs=" & n & " | Len=" & Len(s)
End Sub

