Attribute VB_Name = "modADL_Probe"


'=== Util: 謖・ｮ咾aption縺ｮ蜿ｳ髫｣繧ｳ繝ｳ繝懊ｒ霑斐☆・亥酔荳陦個ｱ6pt縺ｧ譛繧りｿ代＞・・===
Private Function GetRightComboByLabelCaptionIn(p As MSForms.page, ByVal cap As String) As MSForms.ComboBox
    Dim i As Long, lb As MSForms.label, c As Control, best As MSForms.ComboBox
    Dim dy As Double, bestDx As Double: bestDx = 1E+30
    ' 1) Caption荳閾ｴ繝ｩ繝吶Ν繧呈爾縺・
    For i = 0 To p.controls.count - 1
        If TypeName(p.controls(i)) = "Label" Then
            Set lb = p.controls(i)
            If lb.caption = cap Then
                ' 2) 蜷後§陦・ﾂｱ6pt)縺ｧ蜿ｳ蛛ｴ縺ｫ縺ゅｋ譛遏ｭ霍晞屬縺ｮComboBox
                For Each c In p.controls
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





'=== Resolve: 襍ｷ螻・虚菴懊・辟｡蜷垢ombo・育ｫ九■荳翫′繧奇ｼ冗ｫ倶ｽ堺ｿ晄戟・峨ｒ蜿門ｾ・===
Private Sub ResolveKyoUnnamedCombos(ByRef cmbStandUp As MSForms.ComboBox, ByRef cmbStandHold As MSForms.ComboBox)
    Dim mp As MSForms.MultiPage, p As MSForms.page, c As Control
    ' mpADL 蜿門ｾ・
    For Each c In frmEval.controls
        If TypeName(c) = "MultiPage" Then
            If c.name = "mpADL" Then Set mp = c: Exit For
        End If
    Next c
    If mp Is Nothing Then Exit Sub
    Set p = mp.Pages(2) ' 襍ｷ螻・虚菴・
    Set cmbStandUp = GetRightComboByLabelCaptionIn(p, "遶九■荳翫′繧・)
    Set cmbStandHold = GetRightComboByLabelCaptionIn(p, "遶倶ｽ堺ｿ晄戟")
End Sub

'=== Snapshot: ADL・・I/IADL/襍ｷ螻・虚菴懶ｼ峨ｒ蝗ｺ螳夐・〒繧ｷ繝ｪ繧｢繝ｩ繧､繧ｺ陦ｨ遉ｺ ===
Public Sub Snapshot_ADL_Once()
    Dim mp As MSForms.MultiPage, p As MSForms.page, ctl As Control
    Dim i As Long, s As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    ' mpADL 蜿門ｾ・
    For Each ctl In frmEval.controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Debug.Print "[ERR] mpADL not found": Exit Sub

    ' --- BI (#0) ---
    Set p = mp.Pages(0)
    s = ""
    v = p.controls("txtBITotal").text: s = s & "BITotal=" & v & "|"
    For i = 0 To 9
        v = p.controls("cmbBI_" & i).text
        s = s & "BI_" & i & "=" & v & "|"
    Next i
    

If mp.Pages(0).controls("chkBIHomeEnv_Entrance").value Then
    s = s & "BI_HomeEnv_0=1|"
Else
    s = s & "BI_HomeEnv_0=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Genkan").value Then
    s = s & "BI_HomeEnv_1=1|"
Else
    s = s & "BI_HomeEnv_1=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_IndoorStep").value Then
    s = s & "BI_HomeEnv_2=1|"
Else
    s = s & "BI_HomeEnv_2=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Stairs").value Then
    s = s & "BI_HomeEnv_3=1|"
Else
    s = s & "BI_HomeEnv_3=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Handrail").value Then
    s = s & "BI_HomeEnv_4=1|"
Else
    s = s & "BI_HomeEnv_4=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Slope").value Then
    s = s & "BI_HomeEnv_5=1|"
Else
    s = s & "BI_HomeEnv_5=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_NarrowPath").value Then
    s = s & "BI_HomeEnv_6=1|"
Else
    s = s & "BI_HomeEnv_6=0|"
End If

s = s & "BI_HomeEnv_Note=" & mp.Pages(0).controls("txtBIHomeEnvNote").text & "|"

    ' --- IADL (#1) ---
    Set p = mp.Pages(1)
    For i = 0 To 8
        v = p.controls("cmbIADL_" & i).text
        s = s & "IADL_" & i & "=" & v & "|"
    Next i
    v = p.controls("txtIADLNote").text
    s = s & "IADLNote=" & v & "|"

    ' --- 襍ｷ螻・虚菴・(#2) ---
    Set p = mp.Pages(2)
    s = s & "Kyo_Roll=" & p.controls("cmbKyo_Roll").text & "|"
    s = s & "Kyo_SitUp=" & p.controls("cmbKyo_SitUp").text & "|"
    s = s & "Kyo_SitHold=" & p.controls("cmbKyo_SitHold").text & "|"

    Call ResolveKyoUnnamedCombos(cmbSU, cmbSH)
    If Not cmbSU Is Nothing Then s = s & "Kyo_StandUp=" & cmbSU.text & "|" Else Debug.Print "[WARN] 遶九■荳翫′繧・譛ｪ隗｣豎ｺ"
    If Not cmbSH Is Nothing Then s = s & "Kyo_StandHold=" & cmbSH.text & "|" Else Debug.Print "[WARN] 遶倶ｽ堺ｿ晄戟 譛ｪ隗｣豎ｺ"

    s = s & "Kyo_Note=" & p.controls("txtKyoNote").text

    Debug.Print "[ADL.IO] "; s
    Debug.Print "[ADL.IO.Len] "; Len(s)
End Sub




'=== ADL IO Builder: 繝輔か繝ｼ繝荳翫・ADL蛟､繧貞崋螳夐・〒騾｣邨舌＠縺ｦ霑斐☆・亥憶菴懃畑縺ｪ縺暦ｼ・===
Public Function Build_ADL_IO() As String
    Dim mp As MSForms.MultiPage, p As MSForms.page, ctl As Control
    Dim i As Long, s As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    ' mpADL 蜿門ｾ・
    For Each ctl In frmEval.controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Function

    ' --- BI (#0) ---
    Set p = mp.Pages(0)
    s = ""
    v = p.controls("txtBITotal").text: s = s & "BITotal=" & v & "|"
    For i = 0 To 9
        v = p.controls("cmbBI_" & i).text
        s = s & "BI_" & i & "=" & v & "|"
    Next i
    

If mp.Pages(0).controls("chkBIHomeEnv_Entrance").value Then
    s = s & "BI_HomeEnv_0=1|"
Else
    s = s & "BI_HomeEnv_0=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Genkan").value Then
    s = s & "BI_HomeEnv_1=1|"
Else
    s = s & "BI_HomeEnv_1=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_IndoorStep").value Then
    s = s & "BI_HomeEnv_2=1|"
Else
    s = s & "BI_HomeEnv_2=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Stairs").value Then
    s = s & "BI_HomeEnv_3=1|"
Else
    s = s & "BI_HomeEnv_3=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Handrail").value Then
    s = s & "BI_HomeEnv_4=1|"
Else
    s = s & "BI_HomeEnv_4=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_Slope").value Then
    s = s & "BI_HomeEnv_5=1|"
Else
    s = s & "BI_HomeEnv_5=0|"
End If

If mp.Pages(0).controls("chkBIHomeEnv_NarrowPath").value Then
    s = s & "BI_HomeEnv_6=1|"
Else
    s = s & "BI_HomeEnv_6=0|"
End If

s = s & "BI_HomeEnv_Note=" & mp.Pages(0).controls("txtBIHomeEnvNote").text & "|"


    ' --- IADL (#1) ---
    Set p = mp.Pages(1)
    For i = 0 To 8
        v = p.controls("cmbIADL_" & i).text
        s = s & "IADL_" & i & "=" & v & "|"
    Next i
    v = p.controls("txtIADLNote").text
    s = s & "IADLNote=" & v & "|"

    ' --- 襍ｷ螻・虚菴・(#2) ---
    Set p = mp.Pages(2)
    s = s & "Kyo_Roll=" & p.controls("cmbKyo_Roll").text & "|"
    s = s & "Kyo_SitUp=" & p.controls("cmbKyo_SitUp").text & "|"
    s = s & "Kyo_SitHold=" & p.controls("cmbKyo_SitHold").text & "|"

    ' 辟｡蜷阪さ繝ｳ繝懆ｧ｣豎ｺ・育ｫ九■荳翫′繧奇ｼ冗ｫ倶ｽ堺ｿ晄戟・・
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "遶九■荳翫′繧・)
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "遶倶ｽ堺ｿ晄戟")
    If Not cmbSU Is Nothing Then s = s & "Kyo_StandUp=" & cmbSU.text & "|"
    If Not cmbSH Is Nothing Then s = s & "Kyo_StandHold=" & cmbSH.text & "|"

    s = s & "Kyo_Note=" & p.controls("txtKyoNote").text

    Build_ADL_IO = s
End Function




'=== Save: ADL・・I/IADL/襍ｷ螻・虚菴懶ｼ峨ｒ EvalData 縺ｫ1陦瑚ｿｽ險假ｼ・O_ADL蛻暦ｼ・===

Public Sub Save_ADL_Once()
    Dim ws As Worksheet, look As Object
    Dim s As String, r As Long, c As Long
    Dim lastCol As Long

    Set ws = ThisWorkbook.Worksheets("EvalData")            ' 譌｢蟄倥・繝ｫ繝托ｼ・ainIO縺ｨ蜷後§諠ｳ螳夲ｼ・
    

    c = EnsureHeader(ws, "IO_ADL")


    ' 霑ｽ險倩｡後ｒ豎ｺ螳夲ｼ医・繝・ム縺ｮ谺｡陦後°繧蛾幕蟋具ｼ・
    r = ws.Cells(ws.rows.count, c).End(xlUp).row: If r < 2 Then r = 2 Else r = r + 1


    ' IO逕滓・ 竊・譖ｸ縺崎ｾｼ縺ｿ
    s = Build_ADL_IO()
    Debug.Print "[Chk]"; TypeName(ws); r; c; TypeName(ws.Cells(r, c))

ws.Cells(r, c).Value2 = CStr(s)




End Sub

'=== Helper: 隕句・縺怜・繧剃ｿ晁ｨｼ縺励※蛻礼分蜿ｷ繧定ｿ斐☆・育┌縺代ｌ縺ｰ1陦檎岼縺ｮ譛ｫ蟆ｾ縺ｫ菴懈・・・===
Public Function EnsureHeader(ws As Worksheet, ByVal header As String) As Long

    Dim m As Variant, lastCol As Long
    m = Application.Match(header, ws.rows(1), 0)
    If IsError(m) Then
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1
        ws.Cells(1, lastCol + 1).value = header
        EnsureHeader = lastCol + 1
    Else
        EnsureHeader = CLng(m)
    End If
End Function







'=== Load: EvalData縺ｮ IO_ADL 譛譁ｰ陦後ｒ隱ｭ縺ｿ霎ｼ縺ｿ縲√ヵ繧ｩ繝ｼ繝縺ｫ蜿肴丐 ===
Public Sub Load_ADL_Latest()
    Dim ws As Worksheet, mp As MSForms.MultiPage, p As MSForms.page, ctl As Control
    Dim c As Long, r As Long, s As String
    Dim parts As Variant, i As Long, n As Long
    Dim k As String, v As String
    Dim cmbSU As MSForms.ComboBox, cmbSH As MSForms.ComboBox

    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.count, c).End(xlUp).row
    If r < 2 Then Exit Sub    ' 繝・・繧ｿ縺ｪ縺・

    s = ReadStr_Compat("IO_ADL", r, ws)
    parts = Split(s, "|")

    ' mpADL 蜿門ｾ・
    For Each ctl In frmEval.controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Sub

    ' 辟｡蜷阪さ繝ｳ繝懶ｼ郁ｵｷ螻・ｼ夂ｫ九■荳翫′繧奇ｼ冗ｫ倶ｽ堺ｿ晄戟・峨ｒ隗｣豎ｺ
    Set p = mp.Pages(2) ' 襍ｷ螻・虚菴・
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "遶九■荳翫′繧・)
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "遶倶ｽ堺ｿ晄戟")

    ' 繝壹い繧帝・ｬ｡蜿肴丐
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextI
        If InStr(1, parts(i), "=") = 0 Then GoTo NextI
        k = Left$(parts(i), InStr(1, parts(i), "=") - 1)
        v = Mid$(parts(i), InStr(1, parts(i), "=") + 1)

        Select Case k
            
    ' --- BI (#0) ---
    Case "BITotal":                 mp.Pages(0).controls("txtBITotal").text = v
    Case "BI_0":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_0"), v
    Case "BI_1":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_1"), v
    Case "BI_2":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_2"), v
    Case "BI_3":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_3"), v
    Case "BI_4":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_4"), v
    Case "BI_5":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_5"), v
    Case "BI_6":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_6"), v
    Case "BI_7":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_7"), v
    Case "BI_8":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_8"), v
    Case "BI_9":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_9"), v

    Case "BI_HomeEnv_0":            mp.Pages(0).controls("chkBIHomeEnv_Entrance").value = (v = "1")
    Case "BI_HomeEnv_1":            mp.Pages(0).controls("chkBIHomeEnv_Genkan").value = (v = "1")
    Case "BI_HomeEnv_2":            mp.Pages(0).controls("chkBIHomeEnv_IndoorStep").value = (v = "1")
    Case "BI_HomeEnv_3":            mp.Pages(0).controls("chkBIHomeEnv_Stairs").value = (v = "1")
    Case "BI_HomeEnv_4":            mp.Pages(0).controls("chkBIHomeEnv_Handrail").value = (v = "1")
    Case "BI_HomeEnv_5":            mp.Pages(0).controls("chkBIHomeEnv_Slope").value = (v = "1")
    Case "BI_HomeEnv_6":            mp.Pages(0).controls("chkBIHomeEnv_NarrowPath").value = (v = "1")
    Case "BI_HomeEnv_Note":         mp.Pages(0).controls("txtBIHomeEnvNote").text = v


    ' --- IADL (#1) ---
    Case "IADL_0":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_0"), v
    Case "IADL_1":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_1"), v
    Case "IADL_2":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_2"), v
    Case "IADL_3":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_3"), v
    Case "IADL_4":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_4"), v
    Case "IADL_5":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_5"), v
    Case "IADL_6":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_6"), v
    Case "IADL_7":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_7"), v
    Case "IADL_8":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_8"), v
    Case "IADLNote":                mp.Pages(1).controls("txtIADLNote").text = v

    ' --- 襍ｷ螻・虚菴・(#2) ---
    Case "Kyo_Roll":                SafeSetComboValue mp.Pages(2).controls("cmbKyo_Roll"), v
    Case "Kyo_SitUp":               SafeSetComboValue mp.Pages(2).controls("cmbKyo_SitUp"), v
    Case "Kyo_SitHold":             SafeSetComboValue mp.Pages(2).controls("cmbKyo_SitHold"), v
    Case "Kyo_StandUp":             If Not cmbSU Is Nothing Then SafeSetComboValue cmbSU, v
    Case "Kyo_StandHold":           If Not cmbSH Is Nothing Then SafeSetComboValue cmbSH, v
    Case "Kyo_Note":                mp.Pages(2).controls("txtKyoNote").text = v
End Select

        n = n + 1
NextI:
    Next i

    Debug.Print "[ADL.Load] Row=" & r & " | Pairs=" & n & " | Len=" & Len(s)
End Sub



'=== Save竊鱈oad: ADL 繧剃ｸ逋ｺ讀懆ｨｼ・・valData縺ｫ霑ｽ險倪・逶ｴ蠕後↓繝輔か繝ｼ繝縺ｸ蜿肴丐・・===
Public Sub SaveAndReload_ADL()
    Dim ws As Worksheet, c As Long, r As Long, s As String
    Call Save_ADL_Once
    Call Load_ADL_Latest

    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.count, c).End(xlUp).row
    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[ADL.SaveLoad] Row=" & r & " Col=" & c & " | Len=" & Len(s)
End Sub





'=== Checklist: ADL 菫晏ｭ・隱ｭ霎ｼ縺ｮ蛛･蜈ｨ諤ｧ繧剃ｸ逋ｺ遒ｺ隱・===
Public Sub PreRelease_ADL_Checklist()
    Dim ws As Worksheet, c As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets("EvalData")
    c = EnsureHeader(ws, "IO_ADL")
    r = ws.Cells(ws.rows.count, c).End(xlUp).row
    If r < 2 Then Debug.Print "[ADL.Check] 繝・・繧ｿ縺ｪ縺・: Exit Sub

    s = ReadStr_Compat("IO_Sensory", r, ws)
    Debug.Print "[ADL.Check] Col=" & c & " Row=" & r & " | Len=" & Len(s)

    ' 蜀ｪ遲峨メ繧ｧ繝・け・壻ｿ晏ｭ倪・隱ｭ霎ｼ竊帝聞縺・
    Call SaveAndReload_ADL
    s = Build_ADL_IO
    Debug.Print "[ADL.Check] AfterReload Len=" & Len(s)
End Sub



'

Private Sub WalkCtrlPaths(host As Object, ByVal path As String)

    ' MultiPage 縺ｯ Controls 縺ｧ縺ｯ縺ｪ縺・Pages 繧定ｵｰ譟ｻ
If TypeName(host) = "MultiPage" Then
    Dim pg As MSForms.page
    For Each pg In host.Pages
        WalkCtrlPaths pg, path & "/" & pg.caption & ":Page"
    Next pg
    Exit Sub
End If


    Dim c As Control, t As String, p As String
    For Each c In host.controls
        t = TypeName(c)
        p = path & "/" & c.name & ":" & t
        If c.name = "Frame33" Then Debug.Print "[HIT] "; p
        Select Case t
            Case "Frame", "MultiPage", "Page" '蟄舌ｒ謖√▽蜿ｯ閭ｽ諤ｧ縺後≠繧九ｂ縺ｮ縺縺第ｽ懊ｋ
                WalkCtrlPaths c, p
        End Select
    Next c
End Sub




'=== Save: ADL 繧偵梧欠螳夊｡・r縲阪↓譖ｸ縺崎ｾｼ繧・郁｡後・螟夜Κ縺ｧ豎ｺ螳夲ｼ・===
Public Sub Save_ADL_AtRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Long, s As String
    If ws Is Nothing Then Exit Sub
    If r < 2 Then r = 2

    c = EnsureHeader(ws, "IO_ADL")   ' 隕句・縺礼｢ｺ菫昴＠縺ｦ蛻礼分蜿ｷ蜿門ｾ暦ｼ亥酔蜷阪′莉悶↓縺ゅｋ蝣ｴ蜷医・縲√◎縺ｮ髢｢謨ｰ繧剃ｽｿ逕ｨ縺励※縺・ｋ繝｢繧ｸ繝･繝ｼ繝ｫ縺ｮ繧ゅ・縺ｧ繧０K・・
    s = Build_ADL_IO                 ' 迴ｾ蝨ｨ縺ｮ繝輔か繝ｼ繝蛟､繧棚O蛹厄ｼ亥崋螳夐・ｼ・

    ws.Cells(r, c).Value2 = CStr(s)  ' 謖・ｮ夊｡後↓荳頑嶌縺堺ｿ晏ｭ假ｼ郁ｿｽ險倥・蜻ｼ縺ｳ蜃ｺ縺怜・縺ｧr繧帝ｲ繧√ｋ・・
    Debug.Print "[ADL.Save@Row] Row=" & r & " Col=" & c & " | Len=" & Len(s)
End Sub











Private Function ADLKeyNormalize(ByVal tag As String) As String
    ' UI繧ｿ繧ｰ 竊・菫晏ｭ倥く繝ｼ莠呈鋤
    ' 萓具ｼ咤I.鞫る｣溪・BI_0 / IADL.隱ｿ逅・・IADL_0 / BI.Total竊達ITotal
    Dim m As Object, k As String
    k = Replace(tag, ".", "_")
    If k = "BI_Total" Then ADLKeyNormalize = "BITotal": Exit Function
    
    ' 譌･譛ｬ隱樞・逡ｪ蜿ｷ縺ｮ譛蟆上・繝・・・亥ｿ・ｦ√↓蠢懊§縺ｦ谺｡縺ｮ謇九〒諡｡蠑ｵ・・
    ' BI・医ヰ繝ｼ繧ｵ繝ｫ・・
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = 1
    m("BI_鞫る｣・) = "BI_0"
    m("BI_霆翫＞縺・繝吶ャ繝臥ｧｻ荵・) = "BI_1"
    m("BI_謨ｴ螳ｹ") = "BI_2"
    m("BI_繝医う繝ｬ蜍穂ｽ・) = "BI_3"
    m("BI_蜈･豬ｴ") = "BI_4"
    m("BI_豁ｩ陦・霆翫＞縺咏ｧｻ蜍・) = "BI_5"
    m("BI_髫取ｮｵ譏・剄") = "BI_6"
    m("BI_譖ｴ陦｣") = "BI_7"
    m("BI_謗剃ｾｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ") = "BI_8"
    m("BI_謗貞ｰｿ繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ") = "BI_9"
    
    ' IADL
    m("IADL_隱ｿ逅・) = "IADL_0"
    m("IADL_豢玲ｿｯ") = "IADL_1"
    m("IADL_謗・勁") = "IADL_2"
    m("IADL_雋ｷ縺・黄") = "IADL_3"
    m("IADL_驥鷹姦邂｡逅・) = "IADL_4"
    m("IADL_譛崎脈邂｡逅・) = "IADL_5"
    m("IADL_雜｣蜻ｳ繝ｻ菴呎嚊豢ｻ蜍・) = "IADL_6"
    m("IADL_遉ｾ莨壼盾蜉・亥､門・繝ｻ蝨ｰ蝓滓ｴｻ蜍包ｼ・) = "IADL_7"
    m("IADL_繧ｳ繝溘Η繝九こ繝ｼ繧ｷ繝ｧ繝ｳ・磯崕隧ｱ繝ｻ莨夊ｩｱ・・) = "IADL_8"
    
    If m.exists(k) Then
        ADLKeyNormalize = m(k)
    Else
        ADLKeyNormalize = k ' 譛ｪ螳夂ｾｩ縺ｯ縺昴・縺ｾ縺ｾ・域ｬ｡縺ｮ謇九〒陬懷ｮ鯉ｼ・
    End If
End Function








Public Function FindADLControlByKey(ByVal key As String) As Control
    ' 萓具ｼ嗅ey="BI_0" 繧・"IADL_7" 繧・"BITotal"
    Dim p As Object, pg As Object, ctl As Control, t As String, tag As String
    On Error Resume Next
    Set p = frmEval.controls("mpADL")
    On Error GoTo 0
    If p Is Nothing Then Exit Function

    For Each pg In p.Pages
        For Each ctl In pg.controls
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
    Dim mp As MSForms.MultiPage, p As MSForms.page, ctl As Control
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

    ' mpADL 蜿門ｾ暦ｼ・wner縺九ｉ・・
    For Each ctl In owner.controls
        If TypeName(ctl) = "MultiPage" Then
            If ctl.name = "mpADL" Then Set mp = ctl: Exit For
        End If
    Next ctl
    If mp Is Nothing Then Exit Sub

    ' 辟｡蜷阪さ繝ｳ繝懶ｼ郁ｵｷ螻・ｼ夂ｫ九■荳翫′繧奇ｼ冗ｫ倶ｽ堺ｿ晄戟・峨ｒ隗｣豎ｺ
    Set p = mp.Pages(2) ' 襍ｷ螻・虚菴・
    Set cmbSU = GetRightComboByLabelCaptionIn(p, "遶九■荳翫′繧・)
    Set cmbSH = GetRightComboByLabelCaptionIn(p, "遶倶ｽ堺ｿ晄戟")

    ' 繝壹い繧帝・ｬ｡蜿肴丐
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextI
        If InStr(1, parts(i), "=") = 0 Then GoTo NextI
        k = Left$(parts(i), InStr(1, parts(i), "=") - 1)
        v = Mid$(parts(i), InStr(1, parts(i), "=") + 1)

        Select Case k
            ' --- BI (#0) ---
            Case "BITotal":                 mp.Pages(0).controls("txtBITotal").text = v
            Case "BI_0":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_0"), v
            Case "BI_1":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_1"), v
            Case "BI_2":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_2"), v
            Case "BI_3":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_3"), v
            Case "BI_4":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_4"), v
            Case "BI_5":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_5"), v
            Case "BI_6":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_6"), v
            Case "BI_7":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_7"), v
            Case "BI_8":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_8"), v
            Case "BI_9":                    SafeSetComboValue mp.Pages(0).controls("cmbBI_9"), v

            ' --- IADL (#1) ---
            Case "IADL_0":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_0"), v
            Case "IADL_1":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_1"), v
            Case "IADL_2":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_2"), v
            Case "IADL_3":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_3"), v
            Case "IADL_4":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_4"), v
            Case "IADL_5":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_5"), v
            Case "IADL_6":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_6"), v
            Case "IADL_7":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_7"), v
            Case "IADL_8":                  SafeSetComboValue mp.Pages(1).controls("cmbIADL_8"), v
            Case "IADLNote":                mp.Pages(1).controls("txtIADLNote").text = v

            ' --- 襍ｷ螻・虚菴・(#2) ---
            Case "Kyo_Roll":                SafeSetComboValue mp.Pages(2).controls("cmbKyo_Roll"), v
            Case "Kyo_SitUp":               SafeSetComboValue mp.Pages(2).controls("cmbKyo_SitUp"), v
            Case "Kyo_SitHold":             SafeSetComboValue mp.Pages(2).controls("cmbKyo_SitHold"), v
            Case "Kyo_StandUp":             If Not cmbSU Is Nothing Then SafeSetComboValue cmbSU, v
            Case "Kyo_StandHold":           If Not cmbSH Is Nothing Then SafeSetComboValue cmbSH, v
            Case "Kyo_Note":                mp.Pages(2).controls("txtKyoNote").text = v
        End Select

        n = n + 1
NextI:
    Next i

    Debug.Print "[ADL.Load] Row=" & r & " | Pairs=" & n & " | Len=" & Len(s)
End Sub

