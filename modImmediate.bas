Attribute VB_Name = "modImmediate"


'--- 陬懷勧・医％縺ｮ繝｢繧ｸ繝･繝ｼ繝ｫ蜀・□縺代〒螳檎ｵ撰ｼ・---
Private Function NzLng(v As Variant) As Long
    If IsNumeric(v) Then NzLng = CLng(v) Else NzLng = 0
End Function

Private Function NormName(ByVal s As String) As String
    s = Replace(s, vbCrLf, "")
    s = Replace(s, " ", "")
    s = Replace(s, "縲", "")
    On Error Resume Next
    s = StrConv(s, vbNarrow) ' 蜈ｨ隗停・蜊願ｧ抵ｼ亥庄閭ｽ縺ｪ繧会ｼ・
    On Error GoTo 0
    NormName = LCase$(s)
End Function

Private Sub DumpLook_Simple(ByVal look As Object, ByVal ws As Worksheet)
    Dim k As Variant, c As Variant
    For Each k In Array("Basic.Name", "Basic.ID", "Basic.Age", "Basic.EvalDate")
        If look.exists(k) Then
            c = NzLng(look(k))
            If c > 0 Then
                Debug.Print "[LOOK]", k, "Col", c, "hdr=""" & CStr(ws.Cells(1, c).value) & """"
            Else
                Debug.Print "[LOOK][MISS]", k
            End If
        Else
            Debug.Print "[LOOK][NONE]", k
        End If
    Next
End Sub

Private Sub DumpCandidate_Simple(ByVal ws As Worksheet, ByVal look As Object, ByVal rowNum As Long, ByVal idx As Long)
    Dim nm$, pid$, age$, dt$
    If NzLng(look("Basic.Name")) > 0 Then nm = CStr(ws.Cells(rowNum, look("Basic.Name")).value)
    If NzLng(look("Basic.ID")) > 0 Then pid = CStr(ws.Cells(rowNum, look("Basic.ID")).value)
    If NzLng(look("Basic.Age")) > 0 Then age = CStr(ws.Cells(rowNum, look("Basic.Age")).value)
    If NzLng(look("Basic.EvalDate")) > 0 Then dt = CStr(ws.Cells(rowNum, look("Basic.EvalDate")).value)
    Debug.Print "  [" & idx & "] row=" & rowNum & "  name=""" & nm & """  ID=" & pid & "  age=" & age & "  date=" & dt
End Sub
'=== 縺薙％縺ｾ縺ｧ ===



Public Sub RunEnsureCols()
    Call modEvalIOEntry.EnsureHeaderCol_BasicInfo(modSchema.GetEvalDataSheet())
End Sub


Public Sub Test_HeaderCheck()
    Dim ws As Worksheet, look As Object
    Set ws = modSchema.GetEvalDataSheet()
    Set look = BuildHeaderLookup(ws)

   Debug.Print "陬懷勧蜈ｷ:", modEvalIOEntry.FindColByHeaderExact(ws, "陬懷勧蜈ｷ")
Debug.Print "繝ｪ繧ｹ繧ｯ:", modEvalIOEntry.FindColByHeaderExact(ws, "繝ｪ繧ｹ繧ｯ")


End Sub


  ' ROM繝壹・繧ｸ・壽焚蛟､谺・Label+22px縲∝ｙ閠・Label+28px 繧停懃ｵｶ蟇ｾ蠎ｧ讓吮昴〒謠・∴繧具ｼ郁ｦｪ驕輔＞OK・・
Public Sub ROM_AlignFix_Set20()
    Const targetSmall As Long = 24
    Const targetNote  As Long = 28

    Dim c As Object, mp As Object, pg As Object, i As Long
    Dim bestGap As Double

    ' --- ROM繝壹・繧ｸ迚ｹ螳・---
    For Each c In frmEval.controls
        If TypeName(c) = "MultiPage" Then
            For i = 0 To c.Pages.count - 1
                If InStr(1, CStr(c.Pages(i).caption), "ROM", vbTextCompare) > 0 _
                Or InStr(1, CStr(c.Pages(i).caption), "荳ｻ隕・未遽", vbTextCompare) > 0 Then
                    Set mp = c: Set pg = c.Pages(i): Exit For
                End If
            Next
            If Not pg Is Nothing Then Exit For
        End If
    Next
    If pg Is Nothing Then Debug.Print "[ROM_AlignFix_Set20] ROM page not found": Exit Sub

    Dim ctrl As Object, txt As Object, lbl As Object, tmp As Object, tmpZ As Object
    Dim ml As Boolean, isBig As Boolean
    Dim desired As Single, curr As Single
    Dim adjSmall As Long, adjNote As Long, skip As Long

    For Each ctrl In pg.controls
        If TypeName(ctrl) = "TextBox" Then
            Set txt = ctrl

            ' --- 蛯呵・呵｣懊・蛻､螳・---
            ml = False: On Error Resume Next: ml = txt.multiline: On Error GoTo 0
            isBig = (ml Or txt.Height >= 80 Or txt.Width >= 400 _
                    Or InStr(1, UCase$(txt.name), "NOTE") > 0 _
                    Or InStr(1, UCase$(txt.tag & ""), "NOTE") > 0)

            ' --- 繝ｩ繝吶Ν迚ｹ螳・---
            Set lbl = Nothing
            If isBig Then
                ' 竭 隕ｪ蜀・・縲悟ｙ閠・阪Λ繝吶Ν
                For Each tmp In txt.parent.controls
                    If TypeName(tmp) = "Label" Then
                        If InStr(1, CStr(tmp.caption), "蛯呵・, vbTextCompare) > 0 Then Set lbl = tmp: Exit For
                    End If
                Next tmp
                ' 竭｡ 隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ繝壹・繧ｸ蜈ｨ菴薙・縲悟ｙ閠・阪Λ繝吶Ν
                If lbl Is Nothing Then
                    For Each tmpZ In pg.controls
                        If TypeName(tmpZ) = "Label" Then
                            If InStr(1, CStr(tmpZ.caption), "蛯呵・, vbTextCompare) > 0 Then Set lbl = tmpZ: Exit For
                        End If
                    Next tmpZ
                End If
            End If
            ' 竭｢-陬懶ｼ夊ｦｪ蜀・〒隕九▽縺九ｉ縺ｪ縺代ｌ縺ｰ縲√・繝ｼ繧ｸ蜈ｨ菴薙〒譛霑大ｷｦ・・=・峨Λ繝吶Ν繧呈爾縺・
If lbl Is Nothing Then
    bestGap = 1E+20
    Dim d As Double   ' 竊・竭｡縺ｧ隱ｬ譏弱☆繧・d 縺ｮ螳｣險縲よ里縺ｫ荳翫〒螳｣險縺励※縺・ｌ縺ｰ荳崎ｦ・
    For Each tmp In txt.parent.controls
        If TypeName(tmp) = "Label" Then
            If tmp.Left <= txt.Left And tmp.Width <= 120 _
               And (LenB(CStr(tmp.caption)) >= 2 Or tmp.Width >= 12) Then

                d = AbsTop(txt) - AbsTop(tmp)  ' 荳翫↓縺ゅｋ霍晞屬縺縺第治逕ｨ
                If d >= 0 And d <= 120 Then
                    If d < bestGap Then Set lbl = tmp: bestGap = d
                End If
            End If
        End If
    Next tmp
End If



            ' --- 謠・∴ ---
            If Not lbl Is Nothing Then
              desired = AbsTop(lbl) + IIf(isBig, targetNote, targetSmall)


                curr = AbsTop(txt)
                If Abs(curr - desired) > 0.5 Then
                    txt.Top = txt.Top + (desired - curr)
                    If isBig Then adjNote = adjNote + 1 Else adjSmall = adjSmall + 1
                Else
                    skip = skip + 1
                End If
            Else
                skip = skip + 1
            End If
        End If
    Next ctrl

    

    Debug.Print "[ROM_AlignFix_Set20] AdjustedSmall=" & adjSmall & "  AdjustedNote=" & adjNote & "  Skipped=" & skip
End Sub






' 蛯呵・□縺代ｒ遒ｺ螳溘↓謨ｴ蛻暦ｼ哭abel(縲悟ｙ閠・・縺ｮ24px荳九∈
Public Sub ROM_NoteFix_Once()
    Const target As Long = 24
    Dim c As Object, mp As Object, pg As Object, i As Long
    Dim ctrl As Object, lbl As Object, tb As Object
    Dim ml As Boolean, oldTop As Single

    ' ROM繝壹・繧ｸ迚ｹ螳・
    For Each c In frmEval.controls
        If TypeName(c) = "MultiPage" Then
            Set mp = c
            For i = 0 To mp.Pages.count - 1
                If InStr(1, CStr(mp.Pages(i).caption), "ROM", vbTextCompare) > 0 _
                Or InStr(1, CStr(mp.Pages(i).caption), "荳ｻ隕・未遽", vbTextCompare) > 0 Then
                    Set pg = mp.Pages(i): Exit For
                End If
            Next
            If Not pg Is Nothing Then Exit For
        End If
    Next
    If pg Is Nothing Then Debug.Print "[NoteFix] ROM page not found": Exit Sub

    ' 縲悟ｙ閠・阪Λ繝吶Ν
    For Each ctrl In pg.controls
        If TypeName(ctrl) = "Label" Then
            If InStr(1, CStr(ctrl.caption), "蛯呵・, vbTextCompare) > 0 Then Set lbl = ctrl: Exit For
        End If
    Next
    If lbl Is Nothing Then Debug.Print "[NoteFix] 蛯呵・Λ繝吶Ν縺ｪ縺・: Exit Sub

    ' 譛螟ｧ繧ｵ繧､繧ｺ縺ｮMultiLine繝・く繧ｹ繝茨ｼ亥ｙ閠・悽菴難ｼ・
    For Each ctrl In pg.controls
        If TypeName(ctrl) = "TextBox" Then
            ml = False: On Error Resume Next: ml = ctrl.multiline: On Error GoTo 0
            If ml Or ctrl.Height >= 80 Or ctrl.Width >= 400 Then
                If tb Is Nothing Then
                    Set tb = ctrl
                ElseIf ctrl.Height * ctrl.Width > tb.Height * tb.Width Then
                    Set tb = ctrl
                End If
            End If
        End If
    Next
    If tb Is Nothing Then Debug.Print "[NoteFix] 蛯呵・ユ繧ｭ繧ｹ繝医↑縺・: Exit Sub

    oldTop = tb.Top
    tb.Top = lbl.Top + target
    Debug.Print "[NoteFix] Top: " & oldTop & " -> " & tb.Top & "  (Label.Top=" & lbl.Top & ", ﾎ・" & (tb.Top - lbl.Top) & ")"
End Sub



' 隕ｪ繝√ぉ繝ｼ繝ｳ繧呈怙螟ｧ20谿ｵ縺ｾ縺ｧ縺ｧ謇薙■蛻・ｋ螳牙・迚・
Private Function AbsTop(ByVal o As Object) As Single
    On Error Resume Next
    Dim t As Single: t = NzTop(o)
    Dim p As Object: Set p = GetParentSafe(o)
    Dim guard As Long
    Do While Not p Is Nothing And guard < 20
        t = t + NzTop(p)
        Set p = GetParentSafe(p)
        guard = guard + 1
    Loop
    AbsTop = t
End Function

Private Function NzTop(ByVal o As Object) As Single
    On Error Resume Next
    NzTop = o.Top
End Function

Private Function GetParentSafe(ByVal o As Object) As Object
    On Error Resume Next
    Set GetParentSafe = o.parent
End Function



























' ROM繝壹・繧ｸ蜀・・蜈ｨCheckBox繧偵悟・菴咲ｽｮ縺九ｉ髱樒ｴｯ遨阪〒12px荳翫阪↓驟咲ｽｮ・磯ｫ倥＆/繝輔か繝ｳ繝医・隗ｦ繧峨↑縺・ｼ・
Public Sub ROM_CheckBoxes_Up12_OnROM_Recursive_Once_V2()
    Dim mp As Object, pg As Object, i As Long, found As Boolean
    Dim stk As Collection, cont As Object, c As Object
    Dim j As Long, baseTop As Double, pos As Long
    Dim tagKey As String: tagKey = "CBBase="

    ' ?? ROM繝壹・繧ｸ繧堤音螳夲ｼ・aption縺ｫ ROM / ・ｲ・ｯ・ｭ / 荳ｻ隕・未遽 / 髢｢遽蜿ｯ蜍募沺 繧貞性繧・・?
    For Each mp In frmEval.controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.count - 1
                Set pg = mp.Pages(i)
                If (InStr(1, CStr(pg.caption), "ROM", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "・ｲ・ｯ・ｭ", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "荳ｻ隕・未遽", vbTextCompare) > 0) _
                   Or (InStr(1, CStr(pg.caption), "髢｢遽蜿ｯ蜍募沺", vbTextCompare) > 0) Then
                    found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then MsgBox "ROM繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation: Exit Sub

    ' ?? 蜈･繧悟ｭ・Frame, Page, MultiPage蜀・・Page)繧ょ性繧√※襍ｰ譟ｻ・・BA Collection 繧ｹ繧ｿ繝・け・・?
    Set stk = New Collection
    stk.Add pg

    Do While stk.count > 0
        Set cont = stk(stk.count): stk.Remove stk.count

        On Error Resume Next
        For Each c In cont.controls
            ' 蟄舌さ繝ｳ繝・リ縺ｯ繧ｹ繧ｿ繝・け縺ｫ遨阪・
            Select Case TypeName(c)
                Case "Frame", "Page"
                    stk.Add c
                Case "MultiPage"
                    For j = 0 To c.Pages.count - 1
                        stk.Add c.Pages(j)
                    Next
            End Select

            ' CheckBox縺縺・Top 繧偵悟・Top?12縲阪↓・磯撼邏ｯ遨搾ｼ城ｫ倥＆縺ｯ隗ｦ繧峨↑縺・ｼ・
            If TypeName(c) = "CheckBox" Then
                With c
                    pos = InStr(1, .tag, tagKey, vbTextCompare)
                    If pos > 0 Then
                        baseTop = val(Mid$(.tag, pos + Len(tagKey)))
                    Else
                        baseTop = .Top
                        If Len(.tag) > 0 Then
                            .tag = .tag & "|" & tagKey & CStr(baseTop)
                        Else
                            .tag = tagKey & CStr(baseTop)
                        End If
                    End If
                    .Top = baseTop - 12
                End With
            End If
        Next
        On Error GoTo 0
    Loop
End Sub







' ROM繧ｿ繝悶・譛邨よ､懆ｨｼ・啜extBox鬮倥＆竕20 縺ｨ縲，heckBox縺後悟・Top?12縲阪°繧峨ぜ繝ｬ縺ｦ縺・ｋ莉ｶ謨ｰ繧定｡ｨ遉ｺ
Public Sub ROM_VerifyOnce()
    Dim mp As Object, pg As Object, c As Object, i As Long, found As Boolean
    Dim ngTB As Long, ngCB As Long, base As Double, pos As Long, cap As String

    ' ROM繝壹・繧ｸ迚ｹ螳・
    For Each mp In frmEval.controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.count - 1
                cap = CStr(mp.Pages(i).caption)
                If InStr(1, cap, "ROM", vbTextCompare) > 0 Or _
                   InStr(1, cap, "・ｲ・ｯ・ｭ", vbTextCompare) > 0 Or _
                   InStr(1, cap, "荳ｻ隕・未遽", vbTextCompare) > 0 Or _
                   InStr(1, cap, "髢｢遽蜿ｯ蜍募沺", vbTextCompare) > 0 Then
                    Set pg = mp.Pages(i): found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then Debug.Print "[VERIFY] ROM page not found": Exit Sub

    ' 讀懆ｨｼ
    For Each c In pg.controls
        If TypeName(c) = "TextBox" Then
            If c.multiline = False And c.Height <> 15 Then ngTB = ngTB + 1
        ElseIf TypeName(c) = "CheckBox" Then
            pos = InStr(1, c.tag, "CBBase=", vbTextCompare)
            If pos > 0 Then
                base = val(Mid$(c.tag, pos + 7))
                If Abs((base - 12) - c.Top) > 0.5 Then ngCB = ngCB + 1
            Else
                ngCB = ngCB + 1  ' 蝓ｺ貅匁悴險倬鹸繧ゅぜ繝ｬ謇ｱ縺・
            End If
        End If
    Next

    
End Sub




' ROM繝壹・繧ｸ蜀・・蜈ｨ縺ｦ縺ｮ蜊倩｡卦extBox繧帝ｫ倥＆20縺ｫ蝗ｺ螳夲ｼ・rame繧・・繧悟ｭ舌ｂ蜷ｫ繧√※蜀榊ｸｰ・・
Public Sub ROM_Fix_TextBoxHeight_Recursive_OnROM_Once()
    Const h As Single = 15
    Dim mp As Object, pg As Object, i As Long, found As Boolean

    ' ROM繝壹・繧ｸ繧堤音螳・
    For Each mp In frmEval.controls
        If TypeName(mp) = "MultiPage" Then
            For i = 0 To mp.Pages.count - 1
                Set pg = mp.Pages(i)
                If InStr(1, CStr(pg.caption), "ROM", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "・ｲ・ｯ・ｭ", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "荳ｻ隕・未遽", vbTextCompare) > 0 _
                   Or InStr(1, CStr(pg.caption), "髢｢遽蜿ｯ蜍募沺", vbTextCompare) > 0 Then
                    found = True: Exit For
                End If
            Next
            If found Then Exit For
        End If
    Next
    If Not found Then MsgBox "ROM繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・: Exit Sub

    ' 蜀榊ｸｰ蜃ｦ逅・幕蟋・
    Call FixTextBoxHeightRecursive(pg, h)
End Sub

Private Sub FixTextBoxHeightRecursive(cont As Object, h As Single)
    Dim c As Control, j As Long
    On Error Resume Next
    For Each c In cont.controls
        Select Case TypeName(c)
            Case "Frame", "Page"
                Call FixTextBoxHeightRecursive(c, h)
            Case "MultiPage"
                For j = 0 To c.Pages.count - 1
                    Call FixTextBoxHeightRecursive(c.Pages(j), h)
                Next
            Case "TextBox"
                If c.multiline = False Then c.Height = h
        End Select
    Next
    On Error GoTo 0
End Sub




'--- 蟇ｾ雎｡繝壹・繧ｸ・育ｭ句鴨/MMT・峨・迚ｹ螳・---
Private Function GetMMTPage() As Object
    Dim c As Object, mp As Object, i As Long
    For Each c In frmEval.controls
        If TypeName(c) = "MultiPage" Then
            Set mp = c
            For i = 0 To mp.Pages.count - 1
                Dim cap As String
                cap = mp.Pages(i).caption
                If InStr(cap, "MMT") > 0 Or InStr(cap, "遲句鴨") > 0 Then
                    Set GetMMTPage = mp.Pages(i)
                    Exit Function
                End If
            Next
        End If
    Next
    Set GetMMTPage = Nothing
End Function

'--- 閾ｪ蜍慕函謌舌ち繧ｰ縺ｮ繧ゅ・縺縺大炎髯､ ---
Private Sub MMT_ClearGen(pg As Object)
    Dim idx As Long
    For idx = pg.controls.count - 1 To 0 Step -1
        If Left$(pg.controls(idx).tag & "", 6) = "MMTGEN" Then
            pg.controls.Remove pg.controls(idx).name
        End If
    Next
End Sub

'--- 1繝壹・繧ｸ縺ｶ繧薙・繝ｬ繧､繧｢繧ｦ繝育函謌・---
Private Sub BuildPage(pg As Object, items As Variant)
    Const ROW_H As Single = 24
    Const LBL_W As Single = 130
    Const COL_W As Single = 90
    Const gap As Single = 12

    Dim x0 As Single, y0 As Single
    x0 = 20: y0 = 28

    ' 隕句・縺・
    MakeLabel pg, "lblHdr_Muscle", "遲狗ｾ､", x0, y0 - 20, 60, 18
    MakeLabel pg, "lblHdr_R", "蜿ｳ", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLabel pg, "lblHdr_L", "蟾ｦ", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, y As Single
    y = y0
    For i = LBound(items) To UBound(items)
        Dim key As String: key = CStr(items(i))

        MakeLabel pg, "lbl_" & key, key, x0, y + 3, LBL_W, 18
        MakeCombo pg, "cboR_" & key, x0 + LBL_W + gap, y, COL_W, 18
        MakeCombo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, y, COL_W, 18

        y = y + ROW_H
    Next i
End Sub

'--- Label 逕滓・ ---
Private Sub MakeLabel(pg As Object, nm As String, cap As String, L As Single, t As Single, w As Single, h As Single)
    Dim o As MSForms.label
    Set o = pg.controls.Add("Forms.Label.1", nm, True)
    With o
        .caption = cap
        .Left = L: .Top = t: .Width = w: .Height = h
        .tag = "MMTGEN"
    End With
End Sub

'--- ComboBox 逕滓・・・MT 0・・・・---
Private Sub MakeCombo(pg As Object, nm As String, L As Single, t As Single, w As Single, h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.controls.Add("Forms.ComboBox.1", nm, True)
    With o
        .Left = L: .Top = t: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .tag = "MMTGEN"
        .BoundColumn = 1
        .List = Split("0,1,2,3,4,5", ",")
    End With
End Sub



'=== MMT 蜈・ｼ溘・繝ｼ繧ｸ譁ｹ蠑擾ｼ哺MT_荳願い / MMT_荳玖い 繧定ｿｽ蜉縺励※鬆・岼逕滓・ ===
Public Sub MMT_BuildSiblingTabs()
    Dim mp As Object, iMMT As Long
    If Not FindMMT_MultiPage(mp, iMMT) Then
        MsgBox "MMT/遲句鴨 繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation: Exit Sub
    End If

    Dim pgUpper As Object, pgLower As Object
    Set pgUpper = EnsurePage(mp, "MMT_荳願い", iMMT + 1)
    Set pgLower = EnsurePage(mp, "MMT_荳玖い", iMMT + 2)

    Dim upperItems, lowerItems
    upperItems = Array("閧ｩ螻域峇", "閧ｩ莨ｸ螻・, "閧ｩ螟冶ｻ｢", "閧ｩ蜀・雷", "閧ｩ螟匁雷", "閧伜ｱ域峇", "閧倅ｼｸ螻・, "蜑崎・蝗槫・", "蜑崎・蝗槫､・, "謇矩未遽謗悟ｱ・, "謇矩未遽閭悟ｱ・, "謖・ｱ域峇", "謖・ｼｸ螻・, "豈肴欠蟇ｾ遶・)
    lowerItems = Array("閧｡螻域峇", "閧｡莨ｸ螻・, "閧｡螟冶ｻ｢", "閧｡蜀・ｻ｢", "閹晏ｱ域峇", "閹昜ｼｸ螻・, "雜ｳ髢｢遽閭悟ｱ・, "雜ｳ髢｢遽蠎募ｱ・, "豈崎ｶｾ莨ｸ螻・)

    ClearGenerated pgUpper: ClearGenerated pgLower
    BuildMMTPage pgUpper, upperItems
    BuildMMTPage pgLower, lowerItems

    MsgBox "MMT_荳願い・舟MT_荳玖い 繧剃ｽ懈・縺励∪縺励◆縲・, vbInformation
End Sub

'--- MMT繝壹・繧ｸ繧貞性繧 MultiPage 縺ｨ縺昴・繧､繝ｳ繝・ャ繧ｯ繧ｹ繧貞叙蠕・---
Private Function FindMMT_MultiPage(ByRef mp As Object, ByRef idx As Long) As Boolean
    Dim c As Object, i As Long
    For Each c In frmEval.controls
        If TypeName(c) = "MultiPage" Then
            For i = 0 To c.Pages.count - 1
                Dim cap$: cap = c.Pages(i).caption
                If InStr(cap, "MMT") > 0 Or InStr(cap, "遲句鴨") > 0 Then
                    Set mp = c: idx = i: FindMMT_MultiPage = True: Exit Function
                End If
            Next
        End If
    Next
End Function

'--- 謖・ｮ壻ｽ咲ｽｮ縺ｫ繝壹・繧ｸ繧堤畑諢擾ｼ医≠繧後・蜀榊茜逕ｨ・・---
Private Function EnsurePage(mp As Object, title As String, atIndex As Long) As Object
    Dim i As Long
    For i = 0 To mp.Pages.count - 1
        If mp.Pages(i).caption = title Then Set EnsurePage = mp.Pages(i): Exit Function
    Next
    Dim pg As Object
    Set pg = mp.Pages.Add
    pg.caption = title
    If atIndex >= 0 And atIndex < mp.Pages.count Then mp.Pages(mp.Pages.count - 1).Index = atIndex
    Set EnsurePage = pg
End Function

'--- 閾ｪ蜍慕函謌舌・謗・勁 ---
Private Sub ClearGenerated(pg As Object)
    Dim j As Long
    For j = pg.controls.count - 1 To 0 Step -1
        If Left$(pg.controls(j).tag & "", 6) = "MMTGEN" Then pg.controls.Remove pg.controls(j).name
    Next
End Sub

'--- MMT 1繝壹・繧ｸ蛻・・UI逕滓・ ---
Private Sub BuildMMTPage(pg As Object, items As Variant)
    Const ROW_H As Single = 24, LBL_W As Single = 130, COL_W As Single = 90, gap As Single = 12
    Dim x0 As Single, y0 As Single: x0 = 20: y0 = 28

    MakeLbl pg, "lblHdrMus", "遲狗ｾ､", x0, y0 - 20, 60, 18
    MakeLbl pg, "lblHdrR", "蜿ｳ", x0 + LBL_W + gap, y0 - 20, 30, 18
    MakeLbl pg, "lblHdrL", "蟾ｦ", x0 + LBL_W + gap + COL_W + gap, y0 - 20, 30, 18

    Dim i As Long, y As Single: y = y0
    For i = LBound(items) To UBound(items)
        Dim key$: key = CStr(items(i))
        MakeLbl pg, "lbl_" & key, key, x0, y + 3, LBL_W, 18
        MakeCbo pg, "cboR_" & key, x0 + LBL_W + gap, y, COL_W, 18
        MakeCbo pg, "cboL_" & key, x0 + LBL_W + gap + COL_W + gap, y, COL_W, 18
        y = y + ROW_H
    Next
End Sub

Private Sub MakeLbl(pg As Object, nm$, cap$, L!, t!, w!, h!)
    Dim o As MSForms.label
    Set o = pg.controls.Add("Forms.Label.1", nm, True)
    o.caption = cap: o.Left = L: o.Top = t: o.Width = w: o.Height = h: o.tag = "MMTGEN"
End Sub

Private Sub MakeCbo(ByVal pg As Object, ByVal nm As String, _
                    ByVal L As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim o As MSForms.ComboBox
    Set o = pg.controls.Add("Forms.ComboBox.1", nm, True)
    With o
        .Left = L: .Top = t: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .BoundColumn = 1
        .List = Split("0,1,2,3,4,5", ",")
        .tag = "MMTGEN"
    End With
End Sub





'=== MMT縺ｮ蟄舌ち繝厄ｼ井ｸ願い・丈ｸ玖い・峨ｒ縲｀MT繝壹・繧ｸ蜀・・Frame縺ｫ蜿主ｮｹ縺励※菴懈・ ===
Public Sub MMT_BuildChildTabs_Frame()
    Dim pg As Object
    Set pg = GetMMTPage()
    If pg Is Nothing Then MsgBox "MMT繝壹・繧ｸ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・, vbExclamation: Exit Sub

    Dim fra As MSForms.Frame, mp As MSForms.MultiPage
    On Error Resume Next
    Set fra = pg.controls("fraMMTWrap")
    On Error GoTo 0
    If fra Is Nothing Then
        ' Frame繧偵ヵ繧ｩ繝ｼ繝縺ｫ霑ｽ蜉縺励※縺九ｉ隕ｪ繧樽MT繝壹・繧ｸ縺ｸ
        Set fra = frmEval.controls.Add("Forms.Frame.1", "fraMMTWrap", True)
        Set fra.parent = pg
        With fra
            .caption = ""
            .Left = 12: .Top = 12
            .Width = 820: .Height = 420
            .tag = "MMTGEN"
        End With
    End If

    On Error Resume Next
    Set mp = fra.controls("mpMMTChild")
    On Error GoTo 0
    If mp Is Nothing Then
        Set mp = fra.controls.Add("Forms.MultiPage.1", "mpMMTChild", True)
        With mp
            .Left = 6: .Top = 6
            .Width = fra.Width - 12
            .Height = fra.Height - 12
            .Style = 0
            .TabsPerRow = 4
            .Pages.Clear
            .Pages.Add.caption = "荳願い"
            .Pages.Add.caption = "荳玖い"
        End With
    End If

    ' 荳ｭ霄ｫ繧呈ｧ狗ｯ・
    MMT_ClearGen mp.Pages(0)
    MMT_ClearGen mp.Pages(1)
    BuildPage mp.Pages(0), Array("閧ｩ螻域峇", "閧ｩ莨ｸ螻・, "閧ｩ螟冶ｻ｢", "閧ｩ蜀・雷", "閧ｩ螟匁雷", "閧伜ｱ域峇", "閧倅ｼｸ螻・, "蜑崎・蝗槫・", "蜑崎・蝗槫､・, "謇矩未遽謗悟ｱ・, "謇矩未遽閭悟ｱ・, "謖・ｱ域峇", "謖・ｼｸ螻・, "豈肴欠蟇ｾ遶・)
    BuildPage mp.Pages(1), Array("閧｡螻域峇", "閧｡莨ｸ螻・, "閧｡螟冶ｻ｢", "閧｡蜀・ｻ｢", "閹晏ｱ域峇", "閹昜ｼｸ螻・, "雜ｳ髢｢遽閭悟ｱ・, "雜ｳ髢｢遽蠎募ｱ・, "豈崎ｶｾ莨ｸ螻・)

    MsgBox "MMT繝壹・繧ｸ蜀・↓蟄舌ち繝厄ｼ井ｸ願い・丈ｸ玖い・峨ｒ菴懈・縺励∪縺励◆縲・, vbInformation
End Sub





Public Sub Swap_DailyLogList_ToMonthlyDraftBox()
    On Error GoTo EH

    Dim uf As frmEval
    Dim lb As MSForms.Control
    Dim host As Object          ' fraDailyLog
    Dim tb As MSForms.Control   ' TextBox

    ' 驥崎ｦ・ｼ哢ew 縺ｯ菴ｿ繧上↑縺・ｼ・nitialize縺ｧ關ｽ縺｡繧狗腸蠅・′縺ゅｋ縺溘ａ・・
    Set uf = frmEval            ' 襍ｷ蜍穂ｸｭ縺ｮ繧､繝ｳ繧ｹ繧ｿ繝ｳ繧ｹ繧貞盾辣ｧ

    Set lb = uf.controls("lstDailyLogList")
    Set host = lb.parent        ' fraDailyLog 縺ｮ縺ｯ縺・

    ' 譌｢縺ｫ菴懊▲縺ｦ縺ゅｌ縺ｰ縺昴ｌ繧剃ｽｿ縺・
    On Error Resume Next
    Set tb = host.controls("txtMonthlyMonitoringDraft")
    On Error GoTo EH

    If tb Is Nothing Then
        Set tb = host.controls.Add("Forms.TextBox.1", "txtMonthlyMonitoringDraft", True)
    End If

    ' 菴咲ｽｮ縺ｨ繧ｵ繧､繧ｺ繧・lstDailyLogList 縺ｫ蜷医ｏ縺帙ｋ
    tb.Left = lb.Left
    tb.Top = lb.Top
    tb.Width = lb.Width
    tb.Height = lb.Height

    ' 菴ｿ縺・享謇九・譛蟆城剞・・ordWrap遲峨・隗ｦ繧峨↑縺・ｼ・
    tb.multiline = True
    tb.EnterKeyBehavior = True
    tb.ScrollBars = fmScrollBarsVertical

    ' ListBox 縺ｯ髫縺呻ｼ域綾縺励◆縺・凾縺ｫ謌ｻ縺帙ｋ・・
    lb.Visible = False
    tb.Visible = True

#If APP_DEBUG Then
    Debug.Print "[Swap] lb.Visible=" & lb.Visible, _
                "tb.Name=" & tb.name, _
                "L=" & tb.Left & " T=" & tb.Top & " W=" & tb.Width & " H=" & tb.Height
#End If

    Exit Sub

EH:
#If APP_DEBUG Then
    Debug.Print "[Swap][ERR]", Err.Number, Err.Description
#End If
End Sub


Public Sub Ensure_MonthlyDraftBox_UnderFraDailyLog()
    On Error GoTo EH

    Dim uf As Object
    Dim f As Object
    Dim lb As Object
    Dim tb As Object

    Set uf = VBA.UserForms(0)
    Set f = uf.controls("fraDailyLog")
    Set lb = f.controls("lstDailyLogList")

    ' 譌｢蟄倥′縺ゅｌ縺ｰ蜿門ｾ励√↑縺代ｌ縺ｰ菴懈・・・raDailyLog 逶ｴ荳具ｼ・
    On Error Resume Next
    Set tb = f.controls("txtMonthlyMonitoringDraft")
    On Error GoTo EH

    If tb Is Nothing Then
        Set tb = f.controls.Add("Forms.TextBox.1", "txtMonthlyMonitoringDraft", True)
    End If

    ' ListBox縺ｨ蜷後§遏ｩ蠖｢縺ｫ蜷医ｏ縺帙ｋ・育｢ｺ螳壼､・・
    tb.Left = 12
    tb.Top = 294
    tb.Width = 987.3
    tb.Height = 140

    tb.multiline = True
    tb.EnterKeyBehavior = True
    tb.ScrollBars = fmScrollBarsVertical

    lb.Visible = False
    tb.Visible = True

    Debug.Print "[EnsureDraftBox] OK name=" & tb.name, _
                "L=" & tb.Left & " T=" & tb.Top & " W=" & tb.Width & " H=" & tb.Height
    Exit Sub

EH:
    Debug.Print "[EnsureDraftBox][ERR]", Err.Number, Err.Description
End Sub


Private Function JsonEsc_Local(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEsc_Local = s
End Function


Public Sub ListOpenWorkbooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        Debug.Print wb.name
    Next wb
End Sub


Public Sub Verify_MonthlyExport_OpensWorkbook()
    Debug.Print "Before: ActiveWB=" & ActiveWorkbook.name & "  Count=" & Workbooks.count

    ' 縺・∪菴ｿ縺｣縺ｦ繧句他縺ｳ蜃ｺ縺励→蜷後§
    Call ExportMonitoring_ToMonthlyWorkbook( _
        CDate(frmEval.controls("txtDailyDate").text), _
        frmEval.controls("frHeader").controls("txtHdrName").text, _
        frmEval.controls("txtMonthlyMonitoringDraft").text)

    Debug.Print "After:  ActiveWB=" & ActiveWorkbook.name & "  Count=" & Workbooks.count

    ' 髢九＞縺ｦ縺・ｋ繝悶ャ繧ｯ蜷阪ｒ蛻玲嫌
    ListOpenWorkbooks
End Sub






























